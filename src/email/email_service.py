# src/email/email_service.py
"""Email service with hardcoded email resolver"""
import win32com.client
import json
import os
from typing import Dict, List, Tuple, Optional
from src.utils.logger import setup_logger
from src.utils.email_resolver import EmailResolver
from config.settings import Config

class OutlookEmailService:
    """Handles email operations via Outlook with hardcoded email mapping"""
    
    def __init__(self):
        self.logger = setup_logger(self.__class__.__name__)
        self.outlook = None
        self.is_connected = False
        self.address_book = None
        
        # Initialize email resolver (uses hardcoded path)
        self.email_resolver = EmailResolver()
        
        # Statistics
        self.resolution_stats = {
            'mapped': 0,
            'outlook_resolved': 0,
            'fallback_used': 0,
            'failed': 0
        }
    
    def connect(self) -> bool:
        """Connect to Outlook application"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            try:
                self.address_book = self.outlook.Session.AddressLists.Item("Global Address List")
                self.logger.info("Successfully connected to Outlook and Global Address List")
            except:
                self.logger.warning("Connected to Outlook but couldn't access Global Address List")
            
            self.is_connected = True
            return True
        except Exception as e:
            self.logger.error(f"Error connecting to Outlook: {str(e)}")
            self.is_connected = False
            return False
    
    def search_contact_email(self, name_or_email: str) -> Optional[str]:
        """Search for email address using mapping file, then Outlook, then fallback"""
        if not name_or_email:
            return None
        
        # If it's already an email, return it
        if '@' in name_or_email and '.' in name_or_email:
            self.logger.info(f"'{name_or_email}' is already an email address")
            return name_or_email
        
        # Step 1: Try mapping file resolution
        mapped_email = self.email_resolver.resolve_email(name_or_email)
        if mapped_email:
            self.resolution_stats['mapped'] += 1
            self.logger.info(f"✅ Mapped resolution: '{name_or_email}' -> '{mapped_email}'")
            return mapped_email
        
        # Step 2: Try Outlook Global Address List search
        if self.is_connected and Config.SEARCH_OUTLOOK_CONTACTS and self.address_book:
            try:
                entries = self.address_book.AddressEntries
                for entry in entries:
                    if name_or_email.lower() in entry.Name.lower():
                        email = entry.GetExchangeUser().PrimarySmtpAddress
                        self.resolution_stats['outlook_resolved'] += 1
                        self.logger.info(f"📧 Outlook resolution: '{name_or_email}' -> '{email}'")
                        return email
            except Exception as e:
                self.logger.warning(f"Error searching Outlook GAL for '{name_or_email}': {e}")
        
        # Step 3: No resolution found - record as failed
        self.resolution_stats['failed'] += 1
        self.logger.warning(f"❌ No email resolution found for: '{name_or_email}'")
        return None
    
    def send_email(self, to_address: str, subject: str, body: str, 
                   cc_addresses: List[str] = None, attachments: List[str] = None, 
                   is_html: bool = True) -> bool:
        """Send email via Outlook with mapping resolution"""
        if not self.is_connected:
            self.logger.error("Not connected to Outlook. Call connect() first.")
            return False
        
        try:
            # Resolve email address using mapping
            resolved_email = self.search_contact_email(to_address)
            if not resolved_email:
                self.logger.error(f"Could not resolve email for: {to_address}")
                return False
            
            # Create mail item
            mail = self.outlook.CreateItem(0)  # 0 = Mail item
            
            # Set email properties
            mail.To = resolved_email
            mail.Subject = subject
            
            # Set body format (HTML or plain text)
            if is_html:
                mail.HTMLBody = self._convert_to_html(body)
            else:
                mail.Body = body
            
            # Add CC if provided
            if cc_addresses:
                cc_emails = []
                for cc_addr in cc_addresses:
                    cc_email = self.search_contact_email(cc_addr)
                    if cc_email:
                        cc_emails.append(cc_email)
                
                if cc_emails:
                    mail.CC = "; ".join(cc_emails)
            
            # Add attachments if provided
            if attachments:
                for attachment_path in attachments:
                    if os.path.exists(attachment_path):
                        try:
                            mail.Attachments.Add(attachment_path)
                            self.logger.info(f"Added attachment: {os.path.basename(attachment_path)}")
                        except Exception as e:
                            self.logger.warning(f"Could not attach file {attachment_path}: {e}")
                    else:
                        self.logger.warning(f"Attachment file not found: {attachment_path}")
            
            # Set sender info
            try:
                mail.SenderName = Config.SENDER_DISPLAY_NAME
            except Exception as e:
                self.logger.warning(f"Could not set sender name: {e}")
            
            # Send the email
            mail.Send()
            
            attachment_count = len(attachments) if attachments else 0
            self.logger.info(f"📧 Email sent successfully: {to_address} -> {resolved_email} (attachments: {attachment_count})")
            return True
            
        except Exception as e:
            self.logger.error(f"Error sending email to {to_address}: {str(e)}")
            return False
    
    def _convert_to_html(self, body: str) -> str:
        """Convert plain text with HTML elements to proper HTML email"""
        
        html_email = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            line-height: 1.6; 
            color: #333; 
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }
        .header { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
        }
        .summary { 
            background-color: #f8f9fa; 
            padding: 20px; 
            border-radius: 8px; 
            margin: 20px 0;
            border-left: 5px solid #4CAF50;
        }
        .poc-notice {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
            border-left: 5px solid #f39c12;
        }
        .footer { 
            margin-top: 40px; 
            font-size: 12px; 
            color: #666;
            border-top: 1px solid #ddd; 
            padding-top: 15px;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin: 20px 0; 
        }
        th, td { 
            border: 1px solid #ddd; 
            padding: 8px; 
            text-align: left; 
        }
        th { 
            background-color: #4CAF50; 
            color: white; 
            font-weight: bold; 
        }
        tr:nth-child(even) { 
            background-color: #f9f9f9; 
        }
    </style>
</head>
<body>
"""
        
        lines = body.split('\n')
        in_table = False
        
        for line in lines:
            line = line.strip()
            
            # Handle HTML table
            if line.startswith('<table'):
                in_table = True
                html_email += line + '\n'
            elif line.endswith('</table>') or line.startswith('</table>'):
                html_email += line + '\n'
                in_table = False
            elif in_table:
                html_email += line + '\n'
            # Handle special sections
            elif line.startswith('SUMMARY:'):
                html_email += '<div class="summary"><h3>Summary</h3>\n'
            elif line.startswith('IMPORTANT NOTICE:'):
                html_email += '<div class="poc-notice"><h3>Important Notice</h3>\n'
            elif line.startswith('• '):
                html_email += f'<p>• {line[2:]}</p>\n'
            elif line.startswith('If you have any questions'):
                html_email += '</div><div class="section"><p>' + line + '</p>\n'
            # Handle regular text
            elif line and not line.startswith('<'):
                if line.startswith('Hello '):
                    html_email += f'<div class="header"><h2>{line}</h2></div>\n'
                elif 'Proto4Lab Team' in line:
                    html_email += f'</div><div class="footer"><p><strong>{line}</strong></p>\n'
                elif line.startswith('---'):
                    continue
                elif 'automated email generated' in line:
                    html_email += f'<p><em>{line}</em></p></div>\n'
                else:
                    html_email += f'<p>{line}</p>\n'
        
        html_email += """
</body>
</html>
"""
        
        return html_email
    
    def send_bulk_emails(self, email_data: List[Dict]) -> Tuple[int, int]:
        """Send multiple emails with resolution statistics"""
        successful = 0
        failed = 0
        
        # Reset stats for this batch
        self.resolution_stats = {
            'mapped': 0,
            'outlook_resolved': 0,
            'fallback_used': 0,
            'failed': 0
        }
        
        total_emails = len(email_data)
        self.logger.info(f"Starting bulk email send: {total_emails} emails")
        
        for i, email_info in enumerate(email_data, 1):
            
            success = self.send_email(
                email_info['to'],
                email_info['subject'],
                email_info['body'],
                email_info.get('cc', []),
                email_info.get('attachments', []),
                is_html=True
            )
            
            if success:
                successful += 1
            else:
                failed += 1
        
        # Log resolution statistics
        self.logger.info(f"📊 Email Resolution Statistics:")
        self.logger.info(f"   ✅ Mapped via file: {self.resolution_stats['mapped']}")
        self.logger.info(f"   📧 Outlook resolved: {self.resolution_stats['outlook_resolved']}")
        self.logger.info(f"   ❌ Failed: {self.resolution_stats['failed']}")
        self.logger.info(f"📧 Bulk email results: {successful} successful, {failed} failed")
        
        return successful, failed
    
    def test_email_resolution(self, names: List[str]) -> Dict[str, str]:
        """Test email resolution for a list of names"""
        results = {}
        for name in names:
            email = self.search_contact_email(name)
            results[name] = email
            self.logger.info(f"Test resolution: '{name}' -> '{email if email else 'NOT FOUND'}'")
        
        return results
    
    def get_resolution_stats(self) -> Dict:
        """Get email resolution statistics"""
        return self.resolution_stats.copy()
    
    def get_mapping_info(self) -> Dict:
        """Get information about the email mapping"""
        return self.email_resolver.get_mapping_stats()