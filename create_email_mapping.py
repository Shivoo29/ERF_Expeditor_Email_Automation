# create_email_mapping.py
"""Extract all users from ERF file and create complete email mapping using Outlook auto-complete"""

import pandas as pd
import win32com.client
import time
from datetime import datetime
import os

class EmailMappingGenerator:
    """Generate email mapping from ERF data using Outlook auto-complete"""
    
    def __init__(self):
        self.outlook = None
        self.resolved_emails = {}
        self.failed_resolutions = []
    
    def connect_outlook(self):
        """Connect to Outlook"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("✅ Connected to Outlook")
            return True
        except Exception as e:
            print(f"❌ Failed to connect to Outlook: {e}")
            return False
    
    def extract_users_from_erf(self, erf_file_path):
        """Extract all unique users from the 'Entered by' column"""
        print(f"📊 Reading ERF file: {erf_file_path}")
        
        try:
            # Read the Excel file and find the Main data sheet
            excel_file = pd.ExcelFile(erf_file_path)
            sheet_names = excel_file.sheet_names
            
            # Find the sheet with ERF data (likely 'Main data')
            target_sheet = None
            for sheet in sheet_names:
                df_test = pd.read_excel(erf_file_path, sheet_name=sheet, nrows=1)
                if 'Entered by' in df_test.columns:
                    target_sheet = sheet
                    break
            
            if not target_sheet:
                print("❌ Could not find 'Entered by' column in any sheet")
                return []
            
            print(f"📋 Using sheet: '{target_sheet}'")
            
            # Read the full sheet
            df = pd.read_excel(erf_file_path, sheet_name=target_sheet)
            
            # Extract unique users from 'Entered by' column
            if 'Entered by' not in df.columns:
                print("❌ 'Entered by' column not found")
                return []
            
            # Get unique, non-null users
            users = df['Entered by'].dropna().unique()
            users = [str(user).strip() for user in users if str(user).strip()]
            users = sorted(list(set(users)))  # Remove duplicates and sort
            
            print(f"📋 Found {len(users)} unique users:")
            for i, user in enumerate(users, 1):
                print(f"   {i:2d}. {user}")
            
            return users
            
        except Exception as e:
            print(f"❌ Error reading ERF file: {e}")
            return []
    
    def resolve_email_autocomplete(self, username):
        """Resolve single email using Outlook auto-complete"""
        if not self.outlook:
            return None
        
        try:
            # Create a mail item
            mail = self.outlook.CreateItem(0)
            mail.To = username
            
            # Try to resolve the recipient
            recipients = mail.Recipients
            if recipients.Count > 0:
                recipient = recipients.Item(1)
                
                if recipient.Resolve():
                    resolved_email = recipient.Address
                    
                    # If it's an Exchange address, get SMTP address
                    if resolved_email.startswith('/'):
                        try:
                            exchange_user = recipient.AddressEntry.GetExchangeUser()
                            if exchange_user:
                                resolved_email = exchange_user.PrimarySmtpAddress
                        except:
                            pass
                    
                    # Clean up
                    mail = None
                    
                    if resolved_email and '@' in resolved_email:
                        return resolved_email
            
            # Clean up if failed
            mail = None
            return None
            
        except Exception as e:
            print(f"      Error resolving {username}: {e}")
            return None
    
    def bulk_resolve_all_users(self, users):
        """Resolve emails for all users"""
        if not self.connect_outlook():
            return
        
        total = len(users)
        print(f"\n🔍 Resolving emails for {total} users using Outlook auto-complete...")
        print("=" * 70)
        
        for i, user in enumerate(users, 1):
            print(f"{i:2d}/{total}: {user:<15} ", end="")
            
            email = self.resolve_email_autocomplete(user)
            
            if email:
                self.resolved_emails[user] = email
                print(f"-> {email}")
            else:
                self.failed_resolutions.append(user)
                print("-> NOT FOUND")
            
            # Small delay to avoid overwhelming Outlook
            time.sleep(0.2)
        
        print("=" * 70)
        print(f"✅ Resolved: {len(self.resolved_emails)}")
        print(f"❌ Failed: {len(self.failed_resolutions)}")
        print(f"📊 Success rate: {len(self.resolved_emails)/total*100:.1f}%")
    
    def create_mapping_excel(self, output_file):
        """Create clean email mapping Excel file"""
        mapping_data = []
        
        # Add resolved mappings
        for username, email in self.resolved_emails.items():
            mapping_data.append({
                'Username': username,
                'Email': email,
                'Status': 'RESOLVED',
                'Method': 'Outlook Auto-Complete',
                'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
        
        # Add failed resolutions for manual completion
        for username in self.failed_resolutions:
            mapping_data.append({
                'Username': username,
                'Email': '',  # Empty for manual filling
                'Status': 'NEEDS_MANUAL_INPUT',
                'Method': 'Auto-Complete Failed',
                'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
        
        # Create DataFrame and save
        df = pd.DataFrame(mapping_data)
        df.to_excel(output_file, index=False)
        
        print(f"\n📁 Email mapping saved to: {output_file}")
        print(f"   Resolved emails: {len(self.resolved_emails)}")
        print(f"   Need manual input: {len(self.failed_resolutions)}")
        
        return output_file
    
    def create_simple_mapping_for_resolver(self, output_file):
        """Create simple 2-column mapping file for email resolver"""
        if not self.resolved_emails:
            print("❌ No resolved emails to export")
            return
        
        # Create simple format: Username | Email
        simple_data = []
        for username, email in self.resolved_emails.items():
            simple_data.append({
                'Eng': username,
                'Email': email
            })
        
        df = pd.DataFrame(simple_data)
        df.to_excel(output_file, index=False)
        
        print(f"📁 Simple mapping for resolver saved to: {output_file}")
        return output_file

def main():
    """Main execution"""
    print("🚀 ERF Email Mapping Generator")
    print("=" * 50)
    
    # Get ERF file path
    erf_file = input("Enter path to ERF Excel file: ").strip()
    if not os.path.exists(erf_file):
        print(f"❌ File not found: {erf_file}")
        return
    
    # Initialize generator
    generator = EmailMappingGenerator()
    
    # Extract users from ERF file
    users = generator.extract_users_from_erf(erf_file)
    if not users:
        print("❌ No users found in ERF file")
        return
    
    # Confirm before proceeding
    print(f"\n⚠️  About to resolve emails for {len(users)} users")
    confirm = input("This may take a few minutes. Proceed? (y/n): ").lower()
    if not confirm.startswith('y'):
        print("❌ Cancelled by user")
        return
    
    # Resolve all emails
    generator.bulk_resolve_all_users(users)
    
    # Create output files
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Detailed mapping file
    detailed_file = f"email_mapping_detailed_{timestamp}.xlsx"
    generator.create_mapping_excel(detailed_file)
    
    # Simple mapping file for email resolver
    simple_file = f"email_mapping_simple_{timestamp}.xlsx"
    generator.create_simple_mapping_for_resolver(simple_file)
    
    print(f"\n🎉 Email mapping generation complete!")
    print(f"📁 Files created:")
    print(f"   1. {detailed_file} (detailed report)")
    print(f"   2. {simple_file} (for email resolver)")
    
    if generator.failed_resolutions:
        print(f"\n📋 Users needing manual email lookup:")
        for user in generator.failed_resolutions:
            print(f"   - {user}")
        print(f"\nYou can manually add these to the simple mapping file")

if __name__ == "__main__":
    main()