from plugin_base import AgentPlugin, PluginContext, PluginResult, Schedule

class AutoResponseElioClause(AgentPlugin):
    name = "Auto Response to Elio (Claude-Generated)"
    description = "Automatically generates professional draft responses to emails from elioscarton@gmail.com using Claude AI."
    detail = "Monitors inbox every minute for emails from elioscarton@gmail.com and creates contextual draft responses using Claude Haiku based on the email content."
    version = "1.0.0"
    icon = "✉️"
    author = "CoWorker AI"
    requires_graph = True
    requires_claude = True
    default_schedule = Schedule.every_minutes(1)

    def load(self, context: PluginContext) -> bool:
        return bool(context.graph and context.claude)

    def email_templates_schema(self):
        return {
            'draft_prompt': {
                'type': 'text',
                'label': 'Claude Prompt',
                'default': 'You are a professional accountant. Read the following email and draft a thoughtful, professional response. Keep it concise (2-3 sentences max). Do not include a signature line.'
            }
        }

    def run(self, context: PluginContext) -> PluginResult:
        try:
            # Fetch unread emails from inbox
            emails = context.graph.fetch_unread_emails("Inbox", 25)
            
            # Filter for emails from elioscarton@gmail.com
            target_emails = [
                e for e in emails
                if e.get('from', {}).get('emailAddress', {}).get('address', '').lower() == 'elioscarton@gmail.com'
            ]
            
            drafts_created = 0
            
            for email in target_emails:
                try:
                    message_id = email.get('id')
                    subject = email.get('subject', '(No Subject)')
                    body = email.get('bodyPreview', '') or email.get('body', {}).get('content', '')
                    sender_address = email.get('from', {}).get('emailAddress', {}).get('address', '')
                    
                    # Get the custom prompt or use default
                    prompt_template = self.get_email_template(
                        'draft_prompt',
                        'You are a professional accountant. Read the following email and draft a thoughtful, professional response. Keep it concise (2-3 sentences max). Do not include a signature line.'
                    )
                    
                    # Build Claude prompt
                    claude_prompt = f"""{prompt_template}

Email from: {sender_address}
Subject: {subject}
Body: {body}

Draft response:"""
                    
                    # Call Claude to generate response
                    response = context.claude.messages.create(
                        model=self.get_claude_model(),
                        max_tokens=300,
                        messages=[{"role": "user", "content": claude_prompt}]
                    )
                    
                    generated_response = response.content[0].text.strip()
                    
                    # Create draft with signature image support
                    sig_image_path = context.graph.get_signature_image_path()
                    
                    reply_subject = f"Re: {subject}"
                    reply_body = f"<p>{generated_response}</p>"
                    
                    if sig_image_path:
                        context.graph.create_draft_with_inline_image(
                            sender_address,
                            reply_subject,
                            reply_body,
                            sig_image_path,
                            "signature_image",
                            message_id
                        )
                    else:
                        context.graph.create_draft(
                            sender_address,
                            reply_subject,
                            reply_body,
                            message_id
                        )
                    
                    context.log(f"Draft created for email from {sender_address} with subject '{subject}'")
                    drafts_created += 1
                    
                except Exception as e:
                    context.log(f"Error processing email: {str(e)}")
                    continue
            
            return PluginResult(
                success=True,
                summary=f"Processed {len(target_emails)} emails from elioscarton@gmail.com.",
                drafts_created=drafts_created,
                actions_taken=drafts_created
            )
            
        except Exception as e:
            return PluginResult(
                success=False,
                error=f"Plugin error: {str(e)}",
                summary="Failed to process emails."
            )
