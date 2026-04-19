#!/usr/bin/env python3
"""
Outlook Email Reader — Microsoft Graph API + device_code OAuth

Reads unread Focused emails from the past 24 hours and outputs structured JSON
for the AI skill to summarize and extract insights.

Token file: ~/.outlook-mcp-tokens.json (shared with existing MCP project)
Auth flow: device_code (headless-friendly) + refresh_token
"""

import json, os, sys, time, urllib.request, urllib.parse, urllib.error
from pathlib import Path
from datetime import datetime, timedelta, timezone

# ── Config ──────────────────────────────────────────────────────────
CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "")
if not CLIENT_ID:
    print("[outlook] ERROR: OUTLOOK_CLIENT_ID environment variable is required. Set it in ~/.hermes/.env or your shell profile.", file=sys.stderr)
    sys.exit(1)
TENANT_ID = "common"  # personal outlook.com accounts
TOKEN_PATH = Path.home() / ".outlook-mcp-tokens.json"
PENDING_AUTH_PATH = Path.home() / ".outlook-mcp-pending-device-auth.json"
SCOPES = "offline_access User.Read Mail.Read"
GRAPH_BASE = "https://graph.microsoft.com/v1.0/"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
DEVICE_CODE_URL = f"{AUTHORITY}/oauth2/v2.0/devicecode"
TOKEN_URL = f"{AUTHORITY}/oauth2/v2.0/token"

# ── Token Management ────────────────────────────────────────────────
def load_tokens():
    if not TOKEN_PATH.exists():
        return None
    try:
        data = json.loads(TOKEN_PATH.read_text())
        if "expires_at" not in data and "expires_in" in data:
            data["expires_at"] = int(time.time() * 1000) + data["expires_in"] * 1000
        return data
    except Exception as e:
        print(f"[outlook] Error loading tokens: {e}", file=sys.stderr)
        return None

def save_tokens(tokens):
    TOKEN_PATH.write_text(json.dumps(tokens, indent=2))
    os.chmod(TOKEN_PATH, 0o600)

def get_access_token():
    """Get a valid access token, refreshing if needed. Returns None if auth required."""
    tokens = load_tokens()
    if not tokens or not tokens.get("access_token"):
        return None

    expires_at = tokens.get("expires_at", 0)
    # 5 minute buffer
    if time.time() * 1000 < expires_at - 300_000:
        return tokens["access_token"]

    # Try refresh
    if not tokens.get("refresh_token"):
        return None

    print("[outlook] Refreshing access token...", file=sys.stderr)
    try:
        new_tokens = refresh_token(tokens["refresh_token"])
        if new_tokens:
            save_tokens(new_tokens)
            return new_tokens["access_token"]
    except Exception as e:
        print(f"[outlook] Refresh failed: {e}", file=sys.stderr)
        return None

    return None

def refresh_token(refresh_token_str):
    payload = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token_str,
        "scope": SCOPES,
    }).encode()

    req = urllib.request.Request(TOKEN_URL, data=payload, method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read())

    if "access_token" not in body:
        raise Exception(f"Refresh failed: {body.get('error_description', body.get('error', 'unknown'))}")

    # Merge: keep old refresh_token if new one not provided
    old = load_tokens() or {}
    result = {**old, **body}
    if "refresh_token" not in body and old.get("refresh_token"):
        result["refresh_token"] = old["refresh_token"]
    result["expires_at"] = int(time.time() * 1000) + body.get("expires_in", 3600) * 1000
    return result

# ── Device Code Auth ────────────────────────────────────────────────
def start_device_code_auth():
    """Start device code flow. Returns instructions for the user."""
    payload = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "scope": SCOPES,
    }).encode()

    req = urllib.request.Request(DEVICE_CODE_URL, data=payload, method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read())

    # Save pending auth
    pending = {
        "device_code": body["device_code"],
        "user_code": body["user_code"],
        "verification_uri": body["verification_uri"],
        "expires_at": int(time.time()) + body.get("expires_in", 900),
        "interval": body.get("interval", 5),
    }
    PENDING_AUTH_PATH.write_text(json.dumps(pending, indent=2))

    return {
        "status": "pending",
        "message": f"请完成认证：\n1. 打开 {body['verification_uri']}\n2. 输入代码：{body['user_code']}\n3. 授权后再次运行本脚本",
        "user_code": body["user_code"],
        "verification_uri": body["verification_uri"],
    }

def poll_device_code():
    """Poll for device code completion. Returns tokens if authorized."""
    if not PENDING_AUTH_PATH.exists():
        return {"status": "no_pending", "message": "没有待完成的认证。请先运行 --auth"}

    pending = json.loads(PENDING_AUTH_PATH.read_text())
    if pending["expires_at"] <= time.time():
        PENDING_AUTH_PATH.unlink(missing_ok=True)
        return {"status": "expired", "message": "认证已过期，请重新运行 --auth"}

    payload = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": pending["device_code"],
    }).encode()

    req = urllib.request.Request(TOKEN_URL, data=payload, method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read())

    if "access_token" in body:
        body["expires_at"] = int(time.time() * 1000) + body.get("expires_in", 3600) * 1000
        save_tokens(body)
        PENDING_AUTH_PATH.unlink(missing_ok=True)
        return {"status": "authorized", "message": "认证成功！"}

    error = body.get("error")
    if error == "authorization_pending":
        return {"status": "pending", "message": "等待授权中，请完成浏览器操作后再次运行"}
    if error == "slow_down":
        return {"status": "pending", "message": "稍后再试"}
    if error in ("expired_token", "authorization_declined", "bad_verification_code"):
        PENDING_AUTH_PATH.unlink(missing_ok=True)
        return {"status": "failed", "message": f"认证失败：{error}"}

    return {"status": "error", "message": str(body)}

# ── Graph API Calls ─────────────────────────────────────────────────
def graph_get(access_token, path, params=None):
    """Make a GET request to Microsoft Graph API."""
    url = GRAPH_BASE + path
    if params:
        query = urllib.parse.urlencode(params)
        url += "?" + query

    req = urllib.request.Request(url, method="GET",
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())

def get_focused_folder_id(access_token):
    """Find the Focused view folder in Outlook inbox.
    
    Outlook's Focused/Other split is a UI feature, not a separate folder.
    We use the inferenceClassification to filter Focused emails.
    """
    # Focused/Other is NOT a folder — it's inferenceClassification on messages
    # We'll filter by inferenceClassification eq "Focused" in the query
    return None

# ── Email Fetching ──────────────────────────────────────────────────
def fetch_unread_focused_emails(access_token, hours=24):
    """Fetch unread Focused emails from the past N hours."""
    
    cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)
    cutoff_str = cutoff.strftime("%Y-%m-%dT%H:%M:%SZ")

    # Step 1: List unread emails in inbox from past 24h
    # We can't combine $filter (for date+unread) with $search (for focused)
    # Strategy: use $filter for unread + date, then filter focused locally
    
    params = {
        "$filter": f"isRead eq false and receivedDateTime ge {cutoff_str}",
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,isRead,importance,inferenceClassification,hasAttachments",
        "$top": 50,
    }

    try:
        response = graph_get(access_token, "me/mailFolders/inbox/messages", params)
    except urllib.error.HTTPError as e:
        if e.code == 401:
            return {"error": "auth_required", "message": "Token 无效或过期，请重新认证"}
        return {"error": "api_error", "message": f"API 错误: {e.code} {e.reason}"}

    messages = response.get("value", [])
    
    # Filter to Focused only
    focused = [m for m in messages if m.get("inferenceClassification") == "focused"]
    
    # If no focused classification data, use all (some accounts may not have Focused/Other enabled)
    if not focused and messages and all(m.get("inferenceClassification") is None for m in messages):
        focused = messages
        print(f"[outlook] No inferenceClassification data, using all {len(messages)} unread emails", file=sys.stderr)
    else:
        print(f"[outlook] Found {len(focused)} Focused unread emails (total unread: {len(messages)})", file=sys.stderr)

    # Filter out likely noise: newsletters, automated notifications
    # We'll flag them for the AI to decide, not exclude outright
    result = []
    for m in focused:
        sender = m.get("from", {}).get("emailAddress", {})
        result.append({
            "id": m.get("id"),
            "subject": m.get("subject", "(无主题)"),
            "from_name": sender.get("name", "Unknown"),
            "from_address": sender.get("address", "unknown"),
            "received_time": m.get("receivedDateTime"),
            "body_preview": m.get("bodyPreview", ""),
            "importance": m.get("importance", "normal"),
            "has_attachments": m.get("hasAttachments", False),
            "inference_classification": m.get("inferenceClassification", "N/A"),
        })

    return {"emails": result, "total_unread": len(messages), "focused_count": len(focused)}

def fetch_email_body(access_token, email_id):
    """Fetch the full body of a specific email."""
    params = {
        "$select": "id,subject,body",
    }
    try:
        response = graph_get(access_token, f"me/messages/{email_id}", params)
        body = response.get("body", {})
        content = body.get("content", "")
        content_type = body.get("contentType", "text")
        
        # If HTML, strip to text (basic)
        if content_type == "html" and content:
            content = html_to_text(content)
        
        return {"body": content, "content_type": content_type}
    except Exception as e:
        return {"error": str(e)}

def html_to_text(html):
    """HTML to text conversion — strips tags, removes hidden/invisible content, preserves structure."""
    import re

    # ── Security: remove hidden content (prompt injection prevention) ──
    # Remove elements with display:none, visibility:hidden, opacity:0
    html = re.sub(r'<[^>]+style\s*=\s*["\'][^"\']*(?:display\s*:\s*none|visibility\s*:\s*hidden|opacity\s*:\s*0\b)[^"\']*["\'][^>]*>.*?</[^>]+>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<[^>]+style\s*=\s*["\'][^"\']*(?:display\s*:\s*none|visibility\s*:\s*hidden|opacity\s*:\s*0\b)[^"\']*["\'][^>]*/?>', '', html, flags=re.IGNORECASE)
    # Remove aria-hidden elements
    html = re.sub(r'<[^>]+aria-hidden\s*=\s*["\']true["\'][^>]*>.*?</[^>]+>', '', html, flags=re.DOTALL|re.IGNORECASE)
    # Remove hidden attribute elements
    html = re.sub(r'<[^>]+\bhidden\b[^>]*>.*?</[^>]+>', '', html, flags=re.DOTALL|re.IGNORECASE)

    # ── Remove dangerous/structural elements ──
    html = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<svg[^>]*>.*?</svg>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<math[^>]*>.*?</math>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<head[^>]*>.*?</head>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<iframe[^>]*>.*?</iframe>', '', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<!--.*?-->', '', html, flags=re.DOTALL)

    # ── Convert structural elements ──
    html = re.sub(r'<br\s*/?>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'<p[^>]*>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</p>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'<div[^>]*>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</div>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'<li[^>]*>', '\n- ', html, flags=re.IGNORECASE)
    html = re.sub(r'<h[1-6][^>]*>', '\n', html, flags=re.IGNORECASE)
    html = re.sub(r'</h[1-6]>', '\n', html, flags=re.IGNORECASE)

    # ── Convert links to markdown ──
    html = re.sub(r'<a[^>]+href\s*=\s*["\']([^"\']+)["\'][^>]*>(.*?)</a>', r'[\2](\1)', html, flags=re.DOTALL|re.IGNORECASE)

    # ── Convert emphasis ──
    html = re.sub(r'<(strong|b)[^>]*>(.*?)</\1>', r'**\2**', html, flags=re.DOTALL|re.IGNORECASE)
    html = re.sub(r'<(em|i)[^>]*>(.*?)</\1>', r'*\2*', html, flags=re.DOTALL|re.IGNORECASE)

    # ── Remove all remaining tags ──
    html = re.sub(r'<[^>]+>', '', html)

    # ── Decode HTML entities ──
    html = html.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'")
    # Numeric entities
    html = re.sub(r'&#(\d+);', lambda m: chr(int(m.group(1))) if 0 < int(m.group(1)) < 65536 else '', html)
    html = re.sub(r'&#x([0-9a-f]+);', lambda m: chr(int(m.group(1),16)) if 0 < int(m.group(1),16) < 65536 else '', html, flags=re.IGNORECASE)

    # ── Remove invisible Unicode characters (critical for email newsletters) ──
    # Zero-width space, joiner, non-joiner, word joiner, invisible chars, BOM, soft hyphen, etc.
    invisible_chars = re.compile(r'[\u200B-\u200D\u2060\u2061-\u2064\u206A-\u206F\uFEFF\u00AD\u034F\u061C\u180E\u2028\u2029\u202A-\u202E\u200E\u200F\u2028\u2029]')
    html = invisible_chars.sub('', html)

    # ── Clean whitespace ──
    html = re.sub(r'[ \t]+', ' ', html)  # Collapse horizontal whitespace
    html = re.sub(r'\n\s*\n\s*\n', '\n\n', html)  # Max 2 consecutive newlines
    html = re.sub(r'\n +', '\n', html)  # Remove leading spaces on lines
    html = html.strip()

    return html

# ── CLI Entry Point ─────────────────────────────────────────────────
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Outlook Email Reader for Hermes Skill")
    parser.add_argument("--auth", action="store_true", help="Start device code authentication")
    parser.add_argument("--poll", action="store_true", help="Poll for device code completion")
    parser.add_argument("--check", action="store_true", help="Check auth status")
    parser.add_argument("--fetch", action="store_true", help="Fetch unread focused emails (past 24h)")
    parser.add_argument("--read", type=str, help="Read full body of a specific email by ID")
    parser.add_argument("--hours", type=int, default=24, help="Lookback hours for --fetch (default: 24)")
    parser.add_argument("--json", action="store_true", help="Output as JSON (default: human-readable)")
    args = parser.parse_args()

    if args.auth:
        result = start_device_code_auth()
        if args.json:
            print(json.dumps(result, ensure_ascii=False))
        else:
            print(result["message"])
        return

    if args.poll:
        result = poll_device_code()
        if args.json:
            print(json.dumps(result, ensure_ascii=False))
        else:
            print(result["message"])
        return

    if args.check:
        token = get_access_token()
        if token:
            tokens = load_tokens()
            expires = datetime.fromtimestamp(tokens["expires_at"] / 1000).isoformat()
            msg = f"✅ 已认证，token 有效至 {expires}"
        else:
            msg = "❌ 未认证或 token 过期，请运行 --auth"
        if args.json:
            print(json.dumps({"authenticated": bool(token), "expires": expires if token else None}, ensure_ascii=False))
        else:
            print(msg)
        return

    if args.read:
        token = get_access_token()
        if not token:
            print(json.dumps({"error": "auth_required"}, ensure_ascii=False))
            return
        result = fetch_email_body(token, args.read)
        print(json.dumps(result, ensure_ascii=False))
        return

    if args.fetch:
        token = get_access_token()
        if not token:
            result = {"error": "auth_required", "message": "请先认证：python outlook_mail.py --auth"}
            print(json.dumps(result, ensure_ascii=False))
            return
        result = fetch_unread_focused_emails(token, hours=args.hours)
        if args.json:
            print(json.dumps(result, ensure_ascii=False))
        else:
            # Human-readable summary
            emails = result.get("emails", [])
            if not emails:
                print(f"过去 {args.hours} 小时没有 Focused 未读邮件")
            else:
                print(f"过去 {args.hours} 小时 Focused 未读邮件 ({len(emails)} 封):")
                for i, e in enumerate(emails, 1):
                    print(f"\n{i}. {e['from_name']} ({e['from_address']})")
                    print(f"   主题: {e['subject']}")
                    print(f"   时间: {e['received_time']}")
                    print(f"   预览: {e['body_preview'][:200]}")
        return

    # Default: fetch with JSON output
    parser.print_help()

if __name__ == "__main__":
    main()