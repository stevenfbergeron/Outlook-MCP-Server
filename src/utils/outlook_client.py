"""High-performance Outlook client for mailbox access and email search.

Fork of Abhishek-Aditya-bs/Outlook-MCP-Server @ bb641f7.

Changes vs upstream:
- Replaces the single-default + single-delegated-shared architecture with a
  multi-account SMTP allowlist. The server enumerates namespace.Accounts and
  operates only on accounts whose SmtpAddress is present (case-insensitively)
  in `included_accounts` (config.properties). This is fail-closed: new
  Outlook accounts added later are NOT searchable until explicitly added.
- Each returned email is labeled with its originating account's SMTP address
  (`account` field; `mailbox_type` retained for backward compatibility).
- Cache mutations are guarded by a threading lock.
- Aux-folder search (Sent Items, Drafts) is explicitly store-scoped per
  account so results cannot leak across accounts.

References:
- Upstream base: https://github.com/Abhishek-Aditya-bs/Outlook-MCP-Server
- Security review: 40 - Knowledge/Claude-AI/2026-04-16-outlook-mcp-abhishek-security-review.md
- Install blueprint: 00 - Inbox/2026-04-16-outlook-mcp-install-blueprint.md
"""

import win32com.client
from datetime import datetime
from typing import List, Dict, Any, Optional
import logging
import pythoncom
import re
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

from ..config.config_reader import config

logger = logging.getLogger(__name__)

# OlDefaultFolders enum values (stable Outlook COM constants)
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_ITEMS = 5
OL_FOLDER_DRAFTS = 16


class OutlookClient:
    """Multi-account Outlook client gated by an SMTP allowlist.

    The allowlist is declared in config.properties as:
        included_accounts=foo@example.com, bar@example.net

    A store is searchable only if its account's SmtpAddress (case-insensitive)
    appears in the allowlist. Fail-closed: empty allowlist => no searches.
    """

    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.connected = False

        # SMTP (lowercase) -> Store COM object for allowlisted accounts
        self._allowed_stores: Dict[str, Any] = {}
        # SMTP (lowercase) -> human-friendly display name
        self._allowed_display_names: Dict[str, str] = {}

        # Caches
        self._search_cache: Dict[str, Dict[str, Any]] = {}
        self._folder_cache: Dict[str, Any] = {}
        self._cache_lock = threading.Lock()

        self._max_retries = config.get_int('max_connection_retries', 3)

    # -----------------------------------------------------------------
    # Connection + allowlist resolution
    # -----------------------------------------------------------------

    def connect(self, retry_attempt: int = 0) -> bool:
        """Connect to Outlook and resolve the SMTP allowlist."""
        try:
            logger.info("Connecting to Outlook...")
            start_time = time.time()

            pythoncom.CoInitialize()

            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                logger.info("Connected to existing Outlook instance")
            except Exception:
                logger.info("No existing Outlook instance, launching new one...")
                self.outlook = win32com.client.Dispatch("Outlook.Application")

            self.namespace = self.outlook.GetNamespace("MAPI")

            if config.get_bool('use_extended_mapi_login', True):
                try:
                    logger.info("Attempting Extended MAPI login...")
                    self.namespace.Logon(None, None, False, True)
                    logger.info("Extended MAPI login successful")
                except Exception as logon_error:
                    logger.warning(f"Extended MAPI login failed: {logon_error}")

            # Resolve the SMTP allowlist -> Store objects
            self._resolve_allowed_stores()

            self.connected = True
            logger.info(
                f"Connected to Outlook in {time.time() - start_time:.2f}s; "
                f"{len(self._allowed_stores)} account(s) on allowlist."
            )
            return True

        except Exception as e:
            logger.error(f"Failed to connect to Outlook (attempt {retry_attempt + 1}): {e}")
            self.connected = False
            if retry_attempt < self._max_retries - 1:
                wait = (2 ** retry_attempt) * 1  # 1s, 2s, 4s
                logger.info(f"Retrying in {wait}s...")
                time.sleep(wait)
                return self.connect(retry_attempt + 1)
            return False

    def _resolve_allowed_stores(self) -> None:
        """Walk namespace.Accounts and pick out allowlisted stores (fail-closed)."""
        self._allowed_stores = {}
        self._allowed_display_names = {}

        allowlist_raw = config.get_list('included_accounts', [])
        if not allowlist_raw:
            logger.warning(
                "included_accounts is empty. No mailboxes will be accessible. "
                "Add SMTP addresses (comma-separated) to config.properties to enable."
            )
            return

        allowlist = {addr.strip().lower() for addr in allowlist_raw if addr and addr.strip()}
        logger.info(f"SMTP allowlist: {sorted(allowlist)}")

        try:
            accounts = self.namespace.Accounts
        except Exception as e:
            logger.error(f"Could not enumerate namespace.Accounts: {e}")
            return

        seen_smtps = set()
        try:
            for i in range(1, accounts.Count + 1):
                try:
                    account = accounts.Item(i)
                except Exception as e:
                    logger.debug(f"Skipping account index {i}: {e}")
                    continue

                smtp = (getattr(account, 'SmtpAddress', '') or '').strip().lower()
                display = getattr(account, 'DisplayName', smtp) or smtp
                if not smtp:
                    logger.debug(f"Account '{display}' has no SmtpAddress; skipping")
                    continue

                seen_smtps.add(smtp)
                if smtp not in allowlist:
                    logger.info(f"Skipping account '{smtp}' (not on allowlist)")
                    continue

                try:
                    store = account.DeliveryStore
                except Exception as e:
                    logger.warning(f"Account '{smtp}' on allowlist but DeliveryStore failed: {e}")
                    continue

                if store is None:
                    logger.warning(f"Account '{smtp}' on allowlist but DeliveryStore is None")
                    continue

                self._allowed_stores[smtp] = store
                self._allowed_display_names[smtp] = display
                logger.info(f"Allowlisted account resolved: {smtp} -> {display}")
        except Exception as e:
            logger.error(f"Error iterating accounts: {e}")

        missing = allowlist - seen_smtps
        if missing:
            logger.warning(
                f"Allowlist entries with no matching Outlook account: {sorted(missing)}. "
                f"These will not be searchable. Check SMTP spelling or account configuration."
            )

        if not self._allowed_stores:
            logger.error(
                "No allowlisted accounts could be resolved to Outlook stores. "
                "Server will return empty results for all search calls."
            )

    # -----------------------------------------------------------------
    # Public API
    # -----------------------------------------------------------------

    def check_access(self) -> Dict[str, Any]:
        """Check connection status and per-account accessibility."""
        if not self.connected:
            if not self.connect():
                return {"error": "Could not connect to Outlook"}

        accounts_status: List[Dict[str, Any]] = []
        errors: List[str] = []

        for smtp, store in self._allowed_stores.items():
            entry = {
                "smtp": smtp,
                "display_name": self._allowed_display_names.get(smtp, smtp),
                "accessible": False,
                "inbox_name": None,
            }
            try:
                inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
                if inbox is not None:
                    entry["accessible"] = True
                    entry["inbox_name"] = getattr(inbox, 'Name', 'Inbox')
            except Exception as e:
                errors.append(f"{smtp}: {e}")
            accounts_status.append(entry)

        return {
            "outlook_connected": True,
            "allowlist_size": len(config.get_list('included_accounts', [])),
            "accounts_resolved": len(self._allowed_stores),
            "accounts": accounts_status,
            "errors": errors,
        }

    def search_emails(
        self,
        search_text: str,
        accounts: Optional[List[str]] = None,
        include_personal: bool = True,  # DEPRECATED, kept for schema compat
        include_shared: bool = True,    # DEPRECATED, kept for schema compat
    ) -> List[Dict[str, Any]]:
        """Search emails across allowlisted accounts.

        Args:
            search_text: Phrase to match (subject + body, exact).
            accounts: Optional list of SMTP addresses to restrict to (must be
                a subset of `included_accounts`). None => all allowlisted.
            include_personal / include_shared: upstream compatibility knobs.
                When both are False and `accounts` is None, nothing is searched.
        """
        if not self.connected:
            if not self.connect():
                return []

        # Resolve target accounts
        if accounts is None:
            if not (include_personal or include_shared):
                logger.info("Both include_personal and include_shared are False; nothing to search.")
                return []
            target_smtps = list(self._allowed_stores.keys())
        else:
            normalized = {a.strip().lower() for a in accounts if a}
            target_smtps = [s for s in self._allowed_stores.keys() if s in normalized]
            unknown = normalized - set(self._allowed_stores.keys())
            if unknown:
                logger.warning(
                    f"Ignoring requested accounts not on allowlist: {sorted(unknown)}"
                )

        if not target_smtps:
            logger.info("No allowlisted accounts to search; returning empty result set.")
            return []

        max_results = config.get_int('max_search_results', 50)

        cache_key = f"{search_text}__{','.join(sorted(target_smtps))}__{max_results}"
        with self._cache_lock:
            entry = self._search_cache.get(cache_key)
            if entry and (time.time() - entry['timestamp'] < 3600):
                logger.info(f"Cache hit for '{search_text}' across {len(target_smtps)} account(s)")
                return entry['data']

        all_emails: List[Dict[str, Any]] = []
        if len(target_smtps) > 1:
            with ThreadPoolExecutor(max_workers=min(len(target_smtps), 4)) as executor:
                futures = {
                    executor.submit(self._search_one_account, smtp, search_text, max_results): smtp
                    for smtp in target_smtps
                }
                for fut in as_completed(futures):
                    smtp = futures[fut]
                    try:
                        all_emails.extend(fut.result())
                    except Exception as e:
                        logger.error(f"Error searching {smtp}: {e}")
        else:
            smtp = target_smtps[0]
            try:
                all_emails.extend(self._search_one_account(smtp, search_text, max_results))
            except Exception as e:
                logger.error(f"Error searching {smtp}: {e}")

        all_emails.sort(key=lambda x: x.get('received_time', datetime.min), reverse=True)
        limited = all_emails[:max_results]

        with self._cache_lock:
            self._search_cache[cache_key] = {'data': limited, 'timestamp': time.time()}
            if len(self._search_cache) > 100:
                oldest = min(
                    self._search_cache.keys(),
                    key=lambda k: self._search_cache[k].get('timestamp', 0),
                )
                del self._search_cache[oldest]

        return limited

    def search_emails_by_subject(self, subject: str,
                                 include_personal: bool = True,
                                 include_shared: bool = True) -> List[Dict[str, Any]]:
        """Legacy method - redirects to search_emails for backward compatibility."""
        return self.search_emails(
            subject,
            accounts=None,
            include_personal=include_personal,
            include_shared=include_shared,
        )

    # -----------------------------------------------------------------
    # Internals
    # -----------------------------------------------------------------

    def _search_one_account(self, smtp: str, search_text: str,
                            max_results: int) -> List[Dict[str, Any]]:
        """Search one allowlisted account. COM-per-thread safe."""
        pythoncom.CoInitialize()
        try:
            store = self._allowed_stores.get(smtp)
            if store is None:
                logger.warning(f"No resolved store for {smtp}; skipping")
                return []
            try:
                inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
            except Exception as e:
                logger.error(f"{smtp}: GetDefaultFolder(Inbox) failed: {e}")
                return []

            return self._search_folder_comprehensive(
                inbox, store, search_text, smtp, max_results
            )
        finally:
            pythoncom.CoUninitialize()

    def _search_folder_comprehensive(self, inbox_folder, store, search_text: str,
                                     account_smtp: str,
                                     max_results: int) -> List[Dict[str, Any]]:
        """Search the inbox (and optionally Sent Items/Drafts) via AdvancedSearch."""
        emails: List[Dict[str, Any]] = []
        found_ids = set()

        try:
            scope = f"'{inbox_folder.FolderPath}'"
            search_text_escaped = search_text.replace('"', '""')
            query = (
                f'urn:schemas:httpmail:subject ci_phrasematch "{search_text_escaped}" OR '
                f'urn:schemas:httpmail:textdescription ci_phrasematch "{search_text_escaped}"'
            )

            logger.info(f"[{account_smtp}] AdvancedSearch in {scope} for '{search_text}'")
            search = self.outlook.AdvancedSearch(
                Scope=scope, Filter=query, SearchSubFolders=False,
                Tag=f"EmailBodySearch_{account_smtp}"
            )

            start_time = time.time()
            while not search.SearchComplete:
                time.sleep(0.1)
                if time.time() - start_time > 30:
                    logger.warning(f"[{account_smtp}] AdvancedSearch timed out")
                    break

            if search.SearchComplete:
                results = search.Results
                result_count = min(results.Count, max_results)
                logger.info(
                    f"[{account_smtp}] AdvancedSearch found {results.Count} "
                    f"(taking {result_count})"
                )
                for i in range(1, result_count + 1):
                    try:
                        item = results.Item(i)
                        entry_id = getattr(item, 'EntryID', '')
                        if entry_id and entry_id not in found_ids:
                            email_data = self._extract_email_data(
                                item, inbox_folder.Name, account_smtp
                            )
                            if email_data:
                                emails.append(email_data)
                                found_ids.add(entry_id)
                    except Exception as e:
                        logger.debug(f"[{account_smtp}] error processing result {i}: {e}")
                        continue
        except Exception as e:
            logger.error(f"[{account_smtp}] AdvancedSearch failed: {e}; trying fallback.")
            try:
                search_text_escaped = search_text.replace("'", "''").replace('"', '""')
                items = inbox_folder.Items
                items.Sort("[ReceivedTime]", True)
                subject_filter = (
                    f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text_escaped}%'"
                )
                filtered = items.Restrict(subject_filter)
                for item in filtered:
                    if len(emails) >= max_results:
                        break
                    entry_id = getattr(item, 'EntryID', '')
                    if entry_id and entry_id not in found_ids:
                        email_data = self._extract_email_data(item, inbox_folder.Name, account_smtp)
                        if email_data:
                            emails.append(email_data)
                            found_ids.add(entry_id)
            except Exception as fe:
                logger.debug(f"[{account_smtp}] Fallback filter failed: {fe}")

        if len(emails) < max_results and config.get_bool('search_all_folders', False):
            emails.extend(self._search_aux_folders(
                store, search_text, account_smtp,
                max_results - len(emails), found_ids
            ))

        return emails

    def _search_aux_folders(self, store, search_text: str, account_smtp: str,
                            max_results: int, found_ids: set) -> List[Dict[str, Any]]:
        """Search Sent Items and Drafts in THIS store (no cross-account leakage).

        NOTE: The upstream config name `search_all_folders` is misleading - it
        only covers Sent Items + Drafts, not every folder. Name preserved for
        backward compatibility; semantics unchanged.
        """
        emails: List[Dict[str, Any]] = []
        aux = [
            ('Sent Items', OL_FOLDER_SENT_ITEMS),
            ('Drafts', OL_FOLDER_DRAFTS),
        ]

        for folder_name, folder_id in aux:
            if len(emails) >= max_results:
                break
            try:
                folder = store.GetDefaultFolder(folder_id)
            except Exception as e:
                logger.debug(f"[{account_smtp}] no {folder_name}: {e}")
                continue
            if folder is None:
                continue

            try:
                scope = f"'{folder.FolderPath}'"
                search_text_escaped = search_text.replace('"', '""')
                query = (
                    f'urn:schemas:httpmail:subject ci_phrasematch "{search_text_escaped}" OR '
                    f'urn:schemas:httpmail:textdescription ci_phrasematch "{search_text_escaped}"'
                )
                logger.info(f"[{account_smtp}] AdvancedSearch in {folder_name} for '{search_text}'")
                search = self.outlook.AdvancedSearch(
                    Scope=scope, Filter=query, SearchSubFolders=False,
                    Tag=f"AuxSearch_{account_smtp}_{folder_name}"
                )
                start_time = time.time()
                while not search.SearchComplete:
                    time.sleep(0.1)
                    if time.time() - start_time > 10:
                        break

                if search.SearchComplete:
                    results = search.Results
                    result_count = min(results.Count, max_results - len(emails))
                    for i in range(1, result_count + 1):
                        try:
                            item = results.Item(i)
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                email_data = self._extract_email_data(
                                    item, folder_name, account_smtp
                                )
                                if email_data:
                                    emails.append(email_data)
                                    found_ids.add(entry_id)
                        except Exception as e:
                            logger.debug(f"[{account_smtp}] aux item err: {e}")
                            continue
            except Exception as e:
                logger.debug(f"[{account_smtp}] aux search {folder_name} failed: {e}")

        return emails

    def _extract_email_data(self, item, folder_name: str,
                            account_smtp: str) -> Optional[Dict[str, Any]]:
        """Extract email data; label with originating account's SMTP."""
        try:
            body = getattr(item, 'Body', '')
            max_body_chars = config.get_int('max_body_chars', 0)
            if max_body_chars > 0 and len(body) > max_body_chars:
                body = body[:max_body_chars] + " [truncated]"
            if config.get_bool('clean_html_content', True) and body:
                body = self._clean_html(body)

            recipients: List[str] = []
            max_recipients = config.get_int('max_recipients_display', 10)
            try:
                count = 0
                for recipient in item.Recipients:
                    if count >= max_recipients:
                        recipients.append(f"... and {item.Recipients.Count - count} more")
                        break
                    recipients.append(
                        getattr(recipient, 'Name', '') or getattr(recipient, 'Address', '')
                    )
                    count += 1
            except Exception:
                pass

            data = {
                'subject': getattr(item, 'Subject', 'No Subject'),
                'sender_name': getattr(item, 'SenderName', 'Unknown'),
                'sender_email': getattr(item, 'SenderEmailAddress', ''),
                'recipients': recipients,
                'received_time': getattr(item, 'ReceivedTime', datetime.now()),
                'folder_name': folder_name,
                'account': account_smtp,            # new: originating account SMTP
                'mailbox_type': account_smtp,       # retained for formatter compat
                'importance': getattr(item, 'Importance', 1),
                'body': body,
                'size': getattr(item, 'Size', 0),
                'attachments_count': (
                    getattr(item.Attachments, 'Count', 0)
                    if hasattr(item, 'Attachments') else 0
                ),
                'unread': getattr(item, 'Unread', False),
                'entry_id': getattr(item, 'EntryID', ''),
            }
            return data
        except Exception as e:
            logger.error(f"Error extracting email data: {e}")
            return None

    def _clean_html(self, text: str) -> str:
        text = re.sub(r'<[^>]+>', '', text)
        entities = {
            '&amp;': '&', '&lt;': '<', '&gt;': '>',
            '&quot;': '"', '&#39;': "'", '&nbsp;': ' ',
        }
        for e, c in entities.items():
            text = text.replace(e, c)
        text = re.sub(r'\s+', ' ', text).strip()
        return text


# Global client instance
outlook_client = OutlookClient()
