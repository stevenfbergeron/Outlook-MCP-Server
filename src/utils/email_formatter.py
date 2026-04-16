"""Simple email formatting for AI-readable responses.

Forked: multi-account aware. Emails now carry an `account` field (the SMTP
address of the originating store); `mailbox_type` is retained for backward
compatibility (it holds the same SMTP string in this fork).
"""

from typing import List, Dict, Any
from datetime import datetime
from collections import defaultdict

from ..config.config_reader import config


def format_mailbox_status(access_result: Dict[str, Any]) -> Dict[str, Any]:
    """Format mailbox access status for AI consumption (multi-account)."""
    return {
        "status": "success",
        "connection": {
            "outlook_connected": access_result.get("outlook_connected", False),
            "timestamp": datetime.now().isoformat(),
        },
        "allowlist_size": access_result.get("allowlist_size", 0),
        "accounts_resolved": access_result.get("accounts_resolved", 0),
        "accounts": access_result.get("accounts", []),
        "errors": access_result.get("errors", []),
        "notes": {
            "security_dialog": "You may need to grant permission when Outlook security dialog appears",
            "allowlist_model": "Only accounts in included_accounts are searchable (fail-closed).",
        },
    }


def format_email_chain(emails: List[Dict[str, Any]], search_subject: str) -> Dict[str, Any]:
    """Format email chain results for AI analysis."""

    if not emails:
        return {
            "status": "no_emails_found",
            "search_subject": search_subject,
            "message": f"No emails found for subject: '{search_subject}'",
        }

    conversations = group_by_conversation(emails)

    stats = {
        "total_emails": len(emails),
        "conversations": len(conversations),
        "date_range": get_date_range(emails),
        "account_distribution": get_account_distribution(emails),
        "mailbox_distribution": get_mailbox_distribution(emails),  # legacy
        "participants": get_participants(emails),
    }

    formatted_conversations = []
    for conv_id, conv_emails in conversations.items():
        conv_emails.sort(key=lambda x: x.get('received_time', datetime.min))
        formatted_conversations.append({
            "conversation_id": conv_id,
            "email_count": len(conv_emails),
            "date_range": get_date_range(conv_emails),
            "participants": get_participants(conv_emails),
            "emails": [format_single_email(email) for email in conv_emails],
        })

    formatted_conversations.sort(
        key=lambda x: max([parse_iso_time(e["received_time"]) for e in x["emails"] if e["received_time"]]),
        reverse=True,
    )

    return {
        "status": "success",
        "search_subject": search_subject,
        "summary": stats,
        "conversations": formatted_conversations,
        "all_emails_chronological": [
            format_single_email(email)
            for email in sorted(
                emails, key=lambda x: x.get('received_time', datetime.min), reverse=True
            )
        ],
    }


def format_alert_analysis(alerts: List[Dict[str, Any]], search_pattern: str) -> Dict[str, Any]:
    """Format alert analysis results for AI consumption. (Unused by current tools.)"""

    if not alerts:
        return {
            "status": "no_alerts_found",
            "search_pattern": search_pattern,
            "message": f"No alerts found for pattern: '{search_pattern}'",
        }

    urgent_alerts = []
    normal_alerts = []
    analyze_importance = config.get_bool('analyze_importance_levels', True)

    for alert in alerts:
        is_urgent = alert.get('importance', 1) > 1
        if analyze_importance:
            subject = alert.get('subject', '').lower()
            urgent_phrases = ['urgent', 'critical', 'emergency', 'asap', 'immediate']
            is_urgent = is_urgent or any(phrase in subject for phrase in urgent_phrases)
        if is_urgent:
            urgent_alerts.append(alert)
        else:
            normal_alerts.append(alert)

    alert_frequency = calculate_daily_frequency(alerts)
    recent_alerts = sorted(alerts, key=lambda x: x.get('received_time', datetime.min), reverse=True)[:10]

    stats = {
        "total_alerts": len(alerts),
        "urgent_alerts": len(urgent_alerts),
        "normal_alerts": len(normal_alerts),
        "date_range": get_date_range(alerts),
        "account_distribution": get_account_distribution(alerts),
        "mailbox_distribution": get_mailbox_distribution(alerts),
        "daily_frequency": alert_frequency,
        "response_indicators": analyze_responses(alerts),
    }

    return {
        "status": "success",
        "search_pattern": search_pattern,
        "summary": stats,
        "urgent_alerts": [format_single_email(alert) for alert in urgent_alerts[:5]],
        "recent_alerts": [format_single_email(alert) for alert in recent_alerts],
        "timeline": create_alert_timeline(alerts),
        "recommendations": generate_alert_recommendations(stats, urgent_alerts),
    }


def format_single_email(email: Dict[str, Any]) -> Dict[str, Any]:
    """Format a single email for AI consumption."""
    formatted = {
        "subject": email.get('subject', 'No Subject'),
        "sender_name": email.get('sender_name', 'Unknown'),
        "sender_email": email.get('sender_email', ''),
        "recipients": email.get('recipients', []),
        "folder": email.get('folder_name', 'Unknown'),
        "account": email.get('account', email.get('mailbox_type', 'unknown')),
        "mailbox": email.get('mailbox_type', email.get('account', 'unknown')),  # legacy
        "body_preview": email.get('body', '')[:500] if email.get('body') else '',
        "attachments": email.get('attachments_count', 0),
        "importance": get_importance_text(email.get('importance', 1)),
        "unread": email.get('unread', False),
        "size_kb": round(email.get('size', 0) / 1024, 1),
    }
    if config.get_bool('include_timestamps', True):
        received_time = email.get('received_time')
        formatted["received_time"] = received_time.isoformat() if received_time else None
    return formatted


def group_by_conversation(emails: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """Group emails by conversation based on subject similarity."""
    conversations = defaultdict(list)
    prefixes = ['re:', 'fwd:', 'fw:', 'reply:', 'forward:']
    for email in emails:
        subject = email.get('subject', '').strip()
        clean_subject = subject
        for prefix in prefixes:
            if clean_subject.lower().startswith(prefix):
                clean_subject = clean_subject[len(prefix):].strip()
        conversations[clean_subject.lower()].append(email)
    return dict(conversations)


def get_date_range(emails: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not emails:
        return {"first": None, "last": None}
    dates = [email.get('received_time') for email in emails if email.get('received_time')]
    if not dates:
        return {"first": None, "last": None}
    return {"first": min(dates).isoformat(), "last": max(dates).isoformat()}


def get_account_distribution(emails: List[Dict[str, Any]]) -> Dict[str, int]:
    """Count emails per originating account SMTP address."""
    counts: Dict[str, int] = defaultdict(int)
    for email in emails:
        key = email.get('account') or email.get('mailbox_type') or 'unknown'
        counts[key] += 1
    return dict(counts)


def get_mailbox_distribution(emails: List[Dict[str, Any]]) -> Dict[str, int]:
    """Legacy distribution counter. In this fork, each SMTP is its own bucket.

    Retained to keep downstream readers that expect this key happy.
    """
    return get_account_distribution(emails)


def get_participants(emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    participant_counts = defaultdict(int)
    participant_emails: Dict[str, str] = {}
    for email in emails:
        sender = email.get('sender_name', 'Unknown')
        sender_email = email.get('sender_email', '')
        participant_counts[sender] += 1
        if sender_email:
            participant_emails[sender] = sender_email
        for recipient in email.get('recipients', []):
            participant_counts[recipient] += 1
    participants = []
    for name, count in sorted(participant_counts.items(), key=lambda x: x[1], reverse=True):
        participants.append({
            "name": name,
            "email": participant_emails.get(name, ''),
            "participation_count": count,
        })
    return participants[:10]


def calculate_daily_frequency(alerts: List[Dict[str, Any]]) -> float:
    if not alerts:
        return 0.0
    dates = [alert.get('received_time').date() for alert in alerts if alert.get('received_time')]
    if not dates:
        return 0.0
    date_range_days = (max(dates) - min(dates)).days + 1
    return round(len(alerts) / max(date_range_days, 1), 2)


def analyze_responses(alerts: List[Dict[str, Any]]) -> Dict[str, Any]:
    replies = sum(1 for alert in alerts if alert.get('subject', '').lower().startswith(('re:', 'reply:')))
    total = len(alerts)
    return {
        "replies_found": replies,
        "total_alerts": total,
        "response_rate_percent": round((replies / total) * 100, 1) if total > 0 else 0,
    }


def create_alert_timeline(alerts: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    timeline = []
    sorted_alerts = sorted(alerts, key=lambda x: x.get('received_time', datetime.min))
    for alert in sorted_alerts:
        timeline.append({
            "timestamp": alert.get('received_time').isoformat() if alert.get('received_time') else None,
            "subject": alert.get('subject', 'No Subject')[:100],
            "sender": alert.get('sender_name', 'Unknown'),
            "account": alert.get('account', alert.get('mailbox_type', 'unknown')),
            "folder": alert.get('folder_name', 'Unknown'),
            "importance": get_importance_text(alert.get('importance', 1)),
        })
    return timeline


def generate_alert_recommendations(stats: Dict[str, Any], urgent_alerts: List[Dict[str, Any]]) -> List[str]:
    recommendations = []
    total_alerts = stats.get('total_alerts', 0)
    urgent_count = stats.get('urgent_alerts', 0)
    daily_frequency = stats.get('daily_frequency', 0)

    if daily_frequency > 5:
        recommendations.append(
            f"High alert frequency detected ({daily_frequency} alerts/day) - investigate root causes"
        )
    if urgent_count > 0 and total_alerts > 0:
        urgent_rate = (urgent_count / total_alerts) * 100
        if urgent_rate > 30:
            recommendations.append(
                f"High urgency rate ({urgent_rate:.1f}%) - review alert thresholds"
            )
    response_rate = stats.get('response_indicators', {}).get('response_rate_percent', 0)
    if response_rate < 50:
        recommendations.append(
            f"Low response rate ({response_rate}%) - review alert response procedures"
        )
    if urgent_count > 0:
        recommendations.append(f"Review {urgent_count} urgent alerts for immediate action")
    if not recommendations:
        recommendations.append("No immediate issues detected - continue monitoring")
    return recommendations


def get_importance_text(importance: int) -> str:
    importance_map = {0: "Low", 1: "Normal", 2: "High"}
    return importance_map.get(importance, "Normal")


def parse_iso_time(iso_string: str) -> datetime:
    try:
        return datetime.fromisoformat(iso_string.replace('Z', '+00:00'))
    except (ValueError, AttributeError):
        return datetime.min
