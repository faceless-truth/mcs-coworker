from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule

class AutoReplyRoss(AgentPlugin):
    name = "Auto Reply to Ross"
    description = "Automatically sends a reply to every email from ross@mcands.com.au"
    detail = "Monitors the Inbox for unread emails from Ross and sends an automatic reply to each one."
    version = "1.0.0"
    icon = "📧"
    author = "CoWorker AI"
    requires_graph = True
    requires_claude = False
    default_schedule = Schedule.every_minutes(5)

    def load(self, context: PluginContext) -> bool:
        return bool(context.graph)

    def run(self, context: PluginContext) -> PluginResult:
        try:
            # Fetch unread emails
            emails = context.graph.fetch_unread_emails("Inbox", 25)
            
            # Filter for emails from ross@mcands.com.au
            ross_emails = [
                e for e in emails
                if e.get('from', {}).get('emailAddress', {}).get('address', '').lower()
                == 'ross@mcands.com.au'
            ]
            
            actions_taken = 0
            
            # Reply to each email from Ross
            for email in ross_emails:
                subject = email.get('subject', 'No Subject')
                message_id = email.get('id')
                sender_address = email.get('from', {}).get('emailAddress', {}).get('address', '')
                
                # Send automatic reply
                context.graph.send_email(
                    sender_address,
                    f"Re: {subject}",
                    "<p>Thank you for your email. I will get back to you shortly.</p>",
                    message_id
                )
                
                # Mark as read
                context.graph.mark_as_read(message_id)
                
                actions_taken += 1
                context.log(f"Auto-reply sent to Ross for: {subject}")
            
            return PluginResult(
                success=True,
                summary=f"{actions_taken} auto-reply email(s) sent to Ross.",
                actions_taken=actions_taken
            )
        
        except Exception as e:
            return PluginResult(
                success=False,
                error=str(e),
                summary="Failed to process emails from Ross."
            )
