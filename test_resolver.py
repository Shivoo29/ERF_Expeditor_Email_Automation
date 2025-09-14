# test_resolver.py
"""Test script to verify Outlook auto-complete resolution"""

from src.utils.email_resolver import EmailResolver

def test_specific_users():
    """Test resolution for your specific problematic users"""
    
    # Initialize resolver
    resolver = EmailResolver()
    
    # Test users that were missing from your mapping
    test_users = ['CHAEBY', 'ALIHA', 'BANGADI', 'BALAKYA', 'BANSASO']
    
    print("Testing Outlook Auto-Complete Resolution")
    print("=" * 50)
    
    for user in test_users:
        print(f"\nTesting: {user}")
        email = resolver.test_single_resolution(user)
    
    print("\n" + "=" * 50)
    print("Test complete!")

def test_bulk_resolution():
    """Test bulk resolution for all your users"""
    
    # Your 41 users that need email resolution
    all_users = [
        'ADALAPR', 'ALIHA', 'BALAKYA', 'BANGADI', 'BANSASO', 'BHALADA', 
        'BILLAPR', 'BTE', 'CHAEBY', 'CHANGPA3'
        # Add the rest of your 41 users here
    ]
    
    resolver = EmailResolver()
    
    print("Bulk Testing Outlook Auto-Complete Resolution")
    print("=" * 60)
    
    results = resolver.bulk_resolve_with_autocomplete(all_users)
    
    # Export results
    report_file = "email_resolution_test_results.xlsx"
    resolver.export_resolution_report(results, report_file)
    
    return results

if __name__ == "__main__":
    print("Choose test mode:")
    print("1. Test specific users (CHAEBY, ALIHA, etc.)")
    print("2. Test bulk resolution")
    
    choice = input("Enter choice (1 or 2): ").strip()
    
    if choice == "1":
        test_specific_users()
    elif choice == "2":
        test_bulk_resolution()
    else:
        print("Invalid choice")
        