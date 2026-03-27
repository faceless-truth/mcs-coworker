from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule
import json
from datetime import datetime

class ElioEailDraftReplies(AgentPlugin):
    """
    Monitors inbox for unread emails addressed to Elio, uses Claude to draft
    intelligent contextual replies, and saves them to Drafts for review.
    Skips emails already handled by Email Triage and no-reply addresses.
    """

    def run(self, context: PluginContext) -> PluginResult:
        context.log("[ElioEmailDraftReplies] Starting plugin...")
        
        try:
            # Fetch unread emails from inbox
            unread_emails = context.graph.fetch_unread_emails('Inbox', max_count=50)
            context.log(f"[ElioEmailDraftReplies] Found {len(unread_emails)} unread emails")
            
            if not unread_emails:
                return PluginResult(success=True, message="No unread emails found.")
            
            draft_count = 0
            skipped_count = 0
            
            for email in unread_emails:
                try:
                    sender_email = email.get('from', {}).get('emailAddress', {}).get('address', '').lower()
                    sender_name = email.get('from', {}).get('emailAddress', {}).get('name', '')
                    subject = email.get('subject', 'No Subject')
                    message_id = email.get('id')
                    body = email.get('bodyPreview', '')
                    is_read = email.get('isRead', False)
                    categories = email.get('categories', [])
                    
                    # Skip if already processed by Email Triage
                    if 'EmailTriage' in categories or 'Triaged' in categories:
                        context.log(f"[ElioEmailDraftReplies] Skipping {sender_email} - already triaged")
                        skipped_count += 1
                        continue
                    
                    # Skip no-reply addresses
                    if self._is_no_reply_address(sender_email):
                        context.log(f"[ElioEmailDraftReplies] Skipping {sender_email} - no-reply address")
                        skipped_count += 1
                        continue
                    
                    # Fetch full email body for better context
                    full_email = self._fetch_full_email(context, message_id)
                    full_body = full_email.get('body', {}).get('content', body) if full_email else body
                    
                    # Check if email is to Elio (in To or CC field)
                    to_recipients = email.get('toRecipients', [])
                    cc_recipients = email.get('ccRecipients', [])
                    all_recipients = to_recipients + cc_recipients
                    
                    is_to_elio = any(
                        'elio' in recipient.get('emailAddress', {}).get('address', '').lower()
                        for recipient in all_recipients
                    )
                    
                    if not is_to_elio:
                        context.log(f"[ElioEmailDraftReplies] Skipping {sender_email} - not addressed to Elio")
                        skipped_count += 1
                        continue
                    
                    context.log(f"[ElioEmailDraftReplies] Processing email from {sender_email}: {subject}")
                    
                    # Use Claude to draft intelligent reply
                    draft_body_html = self._draft_reply_with_claude(context, sender_name, subject, full_body)
                    
                    # Create draft in Drafts folder
                    draft_result = context.graph.create_draft(
                        to=sender_email,
                        subject=f"Re: {subject}",
                        body_html=draft_body_html,
                        reply_to_id=message_id
                    )
                    
                    if draft_result:
                        context.log(f"[ElioEmailDraftReplies] Created draft for {sender_email}")
                        draft_count += 1
                    else:
                        context.log(f"[ElioEmailDraftReplies] Failed to create draft for {sender_email}")
                    
                except Exception as e:
                    context.log(f"[ElioEmailDraftReplies] Error processing email: {str(e)}")
                    continue
            
            result_message = f"Processed {len(unread_emails)} emails. Created {draft_count} drafts. Skipped {skipped_count}."
            context.log(f"[ElioEmailDraftReplies] {result_message}")
            
            return PluginResult(
                success=True,
                message=result_message,
                data={
                    'total_processed': len(unread_emails),
                    'drafts_created': draft_count,
                    'skipped': skipped_count
                }
            )
        
        except Exception as e:
            error_msg = f"Plugin error: {str(e)}"
            context.log(f"[ElioEmailDraftReplies] {error_msg}")
            return PluginResult(success=False, message=error_msg)
    
    def _is_no_reply_address(self, email_address: str) -> bool:
        """
        Check if email is from a no-reply or unsubscribe-only address.
        """
        no_reply_patterns = [
            'noreply@',
            'no-reply@',
            'donotreply@',
            'do-not-reply@',
            'unsubscribe@',
            'notification@',
            'alerts@',
            'mailer-daemon@',
            'postmaster@'
        ]
        
        email_lower = email_address.lower()
        return any(pattern in email_lower for pattern in no_reply_patterns)
    
    def _fetch_full_email(self, context: PluginContext, message_id: str) -> dict:
        """
        Fetch the full email content via Graph API.
        """
        try:
            # Use graph's internal request method to fetch full email
            endpoint = f"/me/messages/{message_id}"
            response = context.graph._make_request('GET', endpoint)
            return response
        except Exception as e:
            context.log(f"[ElioEmailDraftReplies] Could not fetch full email: {str(e)}")
            return None
    
    def _draft_reply_with_claude(self, context: PluginContext, sender_name: str, subject: str, email_body: str) -> str:
        """
        Use Claude Haiku to draft an intelligent contextual reply.
        """
        try:
            prompt = f"""You are drafting a professional reply to an email on behalf of Elio from an accounting practice.

Original email from: {sender_name}
Subject: {subject}

Email content:
{email_body[:2000]}  # Limit to 2000 chars for token efficiency

Draft a professional, concise reply that:
1. Thanks them for their email
2. Addresses the main point or request
3. Provides next steps if applicable
4. Keeps professional tone appropriate for accounting practice
5. Is 2-4 sentences maximum

Respond with ONLY the reply body text, no subject line, no salutation, no signature."""
            
            response = context.claude.generate(
                model=self.get_claude_model(),
                messages=[
                    {"role": "user", "content": prompt}
                ],
                max_tokens=300
            )
            
            reply_text = response.get('content', [{}])[0].get('text', '')
            
            # Format as HTML
            html_body = f"""<html><body>
<p>{reply_text.replace(chr(10), '</p><p>')}</p>
<br/>
<p>Best regards,<br/>Elio</p>
</body></html>"""
            
            return html_body
        
        except Exception as e:
            context.log(f"[ElioEmailDraftReplies] Claude error: {str(e)}")
            # Fallback generic reply
            return f"""<html><body>
<p>Thank you for your email regarding {subject}.</p>
<p>I will review this and get back to you shortly.</p>
<br/>
<p>Best regards,<br/>Elio</p>
</body></html>"""
    
    def get_schedule(self) -> Schedule:
        """
        Run every 15 minutes to monitor for new unread emails.
        """
        return Schedule(interval_seconds=900)  # 15 minutes