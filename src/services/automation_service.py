# src/services/automation_service.py
"""Main automation service with unmapped user export"""
import pandas as pd
import os
from datetime import datetime
from typing import Dict, List, Any, Tuple
from src.data.excel_processor import ExcelProcessor
from src.email.email_service import OutlookEmailService
from src.email.email_templates import EmailTemplate
from src.utils.logger import setup_logger
from config.settings import Config

class ERFAutomationService:
    """Main service that orchestrates the ERF email automation process with unmapped export"""
    
    def __init__(self):
        self.logger = setup_logger(self.__class__.__name__)
        self.excel_processor = ExcelProcessor()
        self.email_service = OutlookEmailService()
        self.email_template = EmailTemplate()
        
    def initialize(self) -> bool:
        """Initialize all services"""
        self.logger.info("Initializing ERF Automation Service with Email Mapping")
        
        # Show mapping file info
        mapping_info = self.email_service.get_mapping_info()
        self.logger.info(f"Loaded email mappings: {mapping_info}")
        
        # Connect to Outlook
        if not self.email_service.connect():
            self.logger.error("Failed to connect to Outlook")
            return False
        
        self.logger.info("All services initialized successfully")
        return True
    
    def process_excel_file(self, file_path: str) -> bool:
        """Process Excel file and prepare data"""
        self.logger.info(f"Processing Excel file: {file_path}")
        
        # Load file
        if not self.excel_processor.load_file(file_path):
            return False
        
        # Filter data
        if not self.excel_processor.filter_data():
            return False
        
        # Group by requester
        if not self.excel_processor.group_by_requester():
            return False
        
        self.logger.info("Excel file processed successfully")
        return True
    
    def generate_email_data_with_resolution(self) -> List[Dict[str, Any]]:
        """Generate email data with actual resolved emails"""
        email_data = []
        grouped_data = self.excel_processor.get_grouped_data()
        
        for requester, items_df in grouped_data.items():
            # Resolve the actual email
            resolved_email = self.email_service.search_contact_email(requester)
            
            email_content = self.email_template.generate_status_email(requester, items_df)
            
            email_data.append({
                'to': requester,
                'resolved_email': resolved_email,
                'subject': email_content['subject'],
                'body': email_content['body'],
                'requester_name': requester,
                'item_count': len(items_df),
                'email_found': resolved_email is not None
            })
        
        return email_data
    
    def preview_emails(self) -> Dict[str, Any]:
        """Generate preview of emails to be sent with resolution info"""
        email_data = self.generate_email_data_with_resolution()
        
        # Test email resolution for preview
        mapped_count = sum(1 for email in email_data if email['email_found'])
        unmapped_count = len(email_data) - mapped_count
        
        preview = {
            'total_emails': len(email_data),
            'mapped_count': mapped_count,
            'unmapped_count': unmapped_count,
            'emails': []
        }
        
        for email in email_data:
            preview['emails'].append({
                'to': email['to'],
                'resolved_email': email['resolved_email'],
                'subject': email['subject'],
                'item_count': email['item_count'],
                'email_found': email['email_found'],
                'body_preview': email['body'][:300] + "..." if len(email['body']) > 300 else email['body']
            })
        
        return preview
    
    def export_unmapped_users(self, mode: str = "demo") -> str:
        """Export unmapped users to Excel file"""
        unmapped_users = self.email_service.email_resolver.get_unmapped_users()
        
        if not unmapped_users:
            self.logger.info("No unmapped users to export")
            return None
        
        # Create filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"unmapped_users_{mode}_{timestamp}.xlsx"
        
        try:
            # Create detailed unmapped users report
            unmapped_data = []
            for user in unmapped_users:
                unmapped_data.append({
                    'Username': user,
                    'Status': 'Email Not Found',
                    'Mode': mode,
                    'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'Recommended_Action': 'Add to email mapping file or verify username'
                })
            
            unmapped_df = pd.DataFrame(unmapped_data)
            unmapped_df.to_excel(filename, index=False)
            
            self.logger.info(f"Exported {len(unmapped_users)} unmapped users to: {filename}")
            return filename
            
        except Exception as e:
            self.logger.error(f"Error exporting unmapped users: {e}")
            return None
    
    def manager_demo_mode(self) -> Tuple[int, int, Dict[str, Any]]:
        """Demo mode with actual email resolution display"""
        email_data = self.generate_email_data_with_resolution()
        
        if not email_data:
            return 0, 0, {'error': 'No email data available'}
        
        # Get first 5 unique requesters
        demo_data = email_data[:5]
        
        print(f"\nMANAGER DEMO MODE WITH EMAIL MAPPING")
        print("=" * 60)
        print(f"Found {len(email_data)} total requesters. Showing first 5 for demo:")
        print()
        
        for i, email in enumerate(demo_data, 1):
            status = "✅ FOUND" if email['email_found'] else "❌ NOT FOUND"
            resolved = email['resolved_email'] if email['resolved_email'] else "Email not found"
            print(f"{i}. {email['requester_name']} - {email['item_count']} items")
            print(f"   Email Resolution: {status} -> {resolved}")
        
        # Ask for test email addresses
        test_emails = []
        print(f"\nEnter email addresses to receive the demo emails:")
        
        while len(test_emails) < 3:
            email = input(f"Test email #{len(test_emails) + 1} (or press Enter if done): ").strip()
            if not email:
                if len(test_emails) > 0:
                    break
                else:
                    print("Please provide at least one email address.")
                    continue
            
            if '@' in email:
                test_emails.append(email)
                print(f"Added: {email}")
            else:
                print("Please enter a valid email address.")
        
        if not test_emails:
            print("No test emails provided. Demo cancelled.")
            return 0, 0, {'error': 'No test emails'}
        
        print(f"\nDemo will send {len(demo_data)} ERF reports to {len(test_emails)} test email(s)")
        confirm = input("Proceed with demo? (y/n): ").lower().startswith('y')
        
        if not confirm:
            print("Demo cancelled.")
            return 0, 0, {'cancelled': True}
        
        # Prepare demo emails with actual resolution info
        demo_email_list = []
        
        for email_info in demo_data:
            original_recipient = email_info['requester_name']
            resolved_email = email_info['resolved_email'] if email_info['resolved_email'] else "Email not found"
            
            for test_email in test_emails:
                demo_subject = f"[DEMO] ERF Status for {original_recipient} - {email_info['item_count']} Items"
                
                demo_body = f"""THIS IS A DEMO EMAIL FOR MANAGER PRESENTATION
============================================================

Email Resolution Demo:
Original Recipient: {original_recipient}
Resolved Email: {resolved_email}
Items: {email_info['item_count']}
Demo sent to: {test_email}

{email_info['body']}

============================================================
END OF DEMO EMAIL - Original would go to: {resolved_email}
"""
                
                demo_email_list.append({
                    'to': test_email,
                    'subject': demo_subject,
                    'body': demo_body
                })
        
        print(f"\nSending {len(demo_email_list)} demo emails...")
        
        # Send demo emails
        successful, failed = self.email_service.send_bulk_emails(demo_email_list)
        
        # Export unmapped users for demo
        unmapped_file = self.export_unmapped_users("demo")
        
        # Get resolution stats
        resolution_stats = self.email_service.get_resolution_stats()
        unmapped_users = self.email_service.email_resolver.get_unmapped_users()
        
        summary = {
            'demo_mode': True,
            'original_requesters': len(email_data),
            'demo_emails_sent': successful,
            'demo_emails_failed': failed,
            'test_addresses': test_emails,
            'demo_requesters': [email['requester_name'] for email in demo_data],
            'resolution_stats': resolution_stats,
            'unmapped_users': unmapped_users,
            'unmapped_users_file': unmapped_file
        }
        
        # Show results
        if unmapped_users:
            print(f"\n⚠️  Found {len(unmapped_users)} users with unmapped emails:")
            for user in unmapped_users[:10]:  # Show first 10
                print(f"   - {user}")
            if len(unmapped_users) > 10:
                print(f"   ... and {len(unmapped_users) - 10} more")
            
            if unmapped_file:
                print(f"\n📁 Exported unmapped users to: {unmapped_file}")
        
        return successful, failed, summary
    
    def send_emails(self, test_mode: bool = True) -> Tuple[int, int, Dict[str, Any]]:
        """Send emails with unmapped user tracking"""
        email_data = self.generate_email_data_with_resolution()
        
        if test_mode:
            self.logger.info("Running in PREVIEW MODE - emails will not be sent")
            
            # Show resolution summary
            mapped_count = sum(1 for email in email_data if email['email_found'])
            unmapped_count = len(email_data) - mapped_count
            
            print(f"\nEmail Resolution Preview:")
            print(f"   ✅ Successfully resolved: {mapped_count}")
            print(f"   ❌ Not resolved: {unmapped_count}")
            
            if unmapped_count > 0:
                unmapped_users = [email['requester_name'] for email in email_data if not email['email_found']]
                print(f"\nUsers with unresolved emails:")
                for user in unmapped_users[:10]:
                    print(f"   - {user}")
                if len(unmapped_users) > 10:
                    print(f"   ... and {len(unmapped_users) - 10} more")
            
            return len(email_data), 0, {
                'test_mode': True, 
                'emails': email_data,
                'mapped_count': mapped_count,
                'unmapped_count': unmapped_count
            }
        
        # Live mode
        print(f"\n⚠️  FINAL CONFIRMATION")
        print(f"About to send {len(email_data)} emails to actual recipients")
        
        # Show resolution summary
        mapped_count = sum(1 for email in email_data if email['email_found'])
        unmapped_count = len(email_data) - mapped_count
        
        print(f"\nEmail Resolution Summary:")
        print(f"   ✅ Successfully resolved: {mapped_count}")
        print(f"   ❌ Not resolved: {unmapped_count}")
        
        if unmapped_count > 0:
            unmapped_users = [email['requester_name'] for email in email_data if not email['email_found']]
            print(f"\nUsers that will NOT receive emails (no email found):")
            for user in unmapped_users[:10]:
                print(f"   - {user}")
            if len(unmapped_users) > 10:
                print(f"   ... and {len(unmapped_users) - 10} more")
        
        final_confirm = input(f"\nProceed to send {mapped_count} emails to resolved addresses? Type 'SEND LIVE' to confirm: ")
        if final_confirm != 'SEND LIVE':
            print("Live email sending cancelled.")
            return 0, 0, {'cancelled': True, 'reason': 'User cancelled'}
        
        # Filter out emails without resolved addresses
        sendable_emails = [email for email in email_data if email['email_found']]
        
        print(f"\nSending {len(sendable_emails)} emails to resolved addresses...")
        successful, failed = self.email_service.send_bulk_emails(sendable_emails)
        
        # Export unmapped users for live mode
        unmapped_file = self.export_unmapped_users("live")
        
        # Get final stats
        summary = self.excel_processor.get_summary()
        resolution_stats = self.email_service.get_resolution_stats()
        unmapped_users = self.email_service.email_resolver.get_unmapped_users()
        
        summary.update({
            'emails_sent': successful,
            'emails_failed': failed,
            'test_mode': False,
            'total_requesters': len(email_data),
            'mapped_count': mapped_count,
            'unmapped_count': unmapped_count,
            'resolution_stats': resolution_stats,
            'unmapped_users': unmapped_users,
            'unmapped_users_file': unmapped_file
        })
        
        # Show final results
        if unmapped_users:
            print(f"\n📁 Exported {len(unmapped_users)} unmapped users to: {unmapped_file}")
        
        return successful, failed, summary
    
    def get_processing_summary(self) -> Dict[str, Any]:
        """Get complete processing summary with email mapping info"""
        summary = self.excel_processor.get_summary()
        summary['mapping_info'] = self.email_service.get_mapping_info()
        return summary