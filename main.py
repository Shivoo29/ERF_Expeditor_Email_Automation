# main.py - SIMPLIFIED WITH HARDCODED EMAIL MAPPING
"""Main entry point for ERF Email Automation with Hardcoded Email Mapping"""
import sys
import os
from src.services.automation_service import ERFAutomationService
from src.utils.logger import setup_logger
from config.settings import Config

def main():
    """Main execution function"""
    logger = setup_logger("main")
    
    print("ERF Email Automation System with Email Mapping")
    print("=" * 60)
    
    # Check for help
    if len(sys.argv) > 1 and sys.argv[1] == "--help":
        print("""
Usage:
  python main.py [Excel file path]
  python main.py --help                         # Show this help

Examples:
  python main.py "ERF report.xlsx"
  python main.py                                # Will prompt for file path
        """)
        return
    
    # Ensure directories exist
    Config.ensure_directories()
    
    # Get Excel file path
    if len(sys.argv) > 1 and not sys.argv[1].startswith('--'):
        excel_file_path = sys.argv[1]
    else:
        excel_file_path = input("Enter the path to your ERF Excel file: ").strip()
    
    if not os.path.exists(excel_file_path):
        logger.error(f"ERF file not found: {excel_file_path}")
        return
    
    print(f"\nConfiguration:")
    print(f"  ERF Data File: {excel_file_path}")
    print(f"  Email Mapping: Hardcoded in email_resolver.py")
    
    # Initialize automation service
    automation = ERFAutomationService()
    
    try:
        # Step 1: Initialize services
        print("\n1️⃣ Initializing services...")
        if not automation.initialize():
            logger.error("Failed to initialize services")
            return
        
        # Step 2: Process Excel file
        print("\n2️⃣ Processing Excel file...")
        if not automation.process_excel_file(excel_file_path):
            logger.error("Failed to process Excel file")
            return
        
        # Step 3: Show processing summary
        print("\n3️⃣ Processing Summary:")
        summary = automation.get_processing_summary()
        print(f"   Total rows: {summary['total_rows']}")
        print(f"   Filtered rows: {summary['filtered_rows']}")
        print(f"   Unique requesters: {summary['unique_requesters']}")
        print(f"   Status breakdown: {summary['status_breakdown']}")
        
        if 'mapping_info' in summary:
            mapping_info = summary['mapping_info']
            print(f"   Email mappings loaded: {mapping_info['total_mappings']}")
        
        # Step 4: Preview emails
        print("\n4️⃣ Email Preview:")
        preview = automation.preview_emails()
        print(f"   Total emails to send: {preview['total_emails']}")
        
        for i, email in enumerate(preview['emails'][:3], 1):
            print(f"   {i}. To: {email['to']} | Items: {email['item_count']}")
        
        if len(preview['emails']) > 3:
            print(f"   ... and {len(preview['emails']) - 3} more")
        
        # Step 5: Mode selection
        print("\n5️⃣ Send Mode Selection:")
        print("Available modes:")
        print("  1. Preview mode (no emails sent, show resolution test)")
        print("  2. Manager Demo mode (send to test emails with mapping info)")
        print("  3. Live mode (send to actual recipients using email mapping)")
        
        mode = input("Select mode (1/2/3): ").strip()
        
        if mode == "1":
            # Pure preview mode
            print("\n6️⃣ Running in PREVIEW MODE...")
            successful, failed, results = automation.send_emails(test_mode=True)
            
            # Show results
            print(f"\nPreview Results:")
            print(f"   Total requesters: {len(results['emails'])}")
            print(f"   Would resolve emails: {results.get('mapped_count', 0)}")
            print(f"   Would fail resolution: {results.get('unmapped_count', 0)}")
            
        elif mode == "2":
            # Manager demo mode
            print("\n6️⃣ Running MANAGER DEMO MODE...")
            successful, failed, results = automation.manager_demo_mode()
            
        elif mode == "3":
            # Live mode
            print(f"\n6️⃣ Preparing for LIVE MODE...")
            print(f"⚠️  This will send emails to actual recipients using email mapping!")
            
            successful, failed, results = automation.send_emails(test_mode=False)
            
        else:
            print("Invalid mode selection")
            return
        
        # Show results
        print(f"\n🎉 Process completed!")
        print(f"   ✅ Successful: {successful}")
        print(f"   ❌ Failed: {failed}")
        print(f"   📧 Total: {successful + failed}")
        
        # Show unmapped users info
        if 'unmapped_users' in results and results['unmapped_users']:
            unmapped_count = len(results['unmapped_users'])
            print(f"\n⚠️  Users with unmapped emails: {unmapped_count}")
            
            # Show first few unmapped users
            for user in results['unmapped_users'][:5]:
                print(f"   - {user}")
            if unmapped_count > 5:
                print(f"   ... and {unmapped_count - 5} more")
            
            if 'unmapped_users_file' in results and results['unmapped_users_file']:
                print(f"\n📁 Exported unmapped users to: {results['unmapped_users_file']}")
        
        # Show resolution stats for live/demo modes
        if mode in ["2", "3"] and 'resolution_stats' in results:
            stats = results['resolution_stats']
            print(f"\nEmail Resolution Stats:")
            print(f"   📋 Mapped via file: {stats.get('mapped', 0)}")
            print(f"   📧 Outlook resolved: {stats.get('outlook_resolved', 0)}")
            print(f"   ❌ Resolution failed: {stats.get('failed', 0)}")
        
        if mode == "1":
            print("\n💡 Run Mode 2 for manager demo or Mode 3 for live emails")
        
    except KeyboardInterrupt:
        print("\n❌ Process interrupted by user")
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        print(f"❌ An error occurred: {str(e)}")

if __name__ == "__main__":
    main()