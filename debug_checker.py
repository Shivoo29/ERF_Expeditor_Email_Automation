import pandas as pd
import sys
import os

def debug_sheet(file_path, sheet_name):
    """Debug a specific sheet to see what's going wrong"""
    try:
        print(f"🔍 DEBUGGING SHEET: '{sheet_name}'")
        print("=" * 60)
        
        # Read the sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        print(f"📊 Rows: {len(df)}")
        print(f"📋 Columns: {len(df.columns)}")
        
        print(f"\n📝 ALL COLUMN NAMES:")
        print("-" * 50)
        for i, col in enumerate(df.columns, 1):
            print(f"{i:2d}. '{col}'")
        
        print(f"\n📋 FIRST 5 ROWS:")
        print("-" * 50)
        print(df.head().to_string())
        
        print(f"\n🔍 LOOKING FOR REQUIRED COLUMNS:")
        print("-" * 50)
        target_cols = ['ERF Sched Line Status', 'Entered by']
        for col in target_cols:
            if col in df.columns:
                print(f"✅ Found: '{col}'")
                # Show some sample values
                sample_values = df[col].dropna().unique()[:5]
                print(f"   Sample values: {list(sample_values)}")
            else:
                print(f"❌ Missing: '{col}'")
        
        # Check for variations of the column names
        print(f"\n🔍 CHECKING FOR SIMILAR COLUMN NAMES:")
        print("-" * 50)
        
        for target in target_cols:
            print(f"\nLooking for variations of '{target}':")
            variations = []
            
            for col in df.columns:
                col_str = str(col).lower()
                target_words = target.lower().split()
                
                # Check if any words match
                matches = sum(1 for word in target_words if word in col_str)
                if matches > 0:
                    variations.append((col, matches))
            
            if variations:
                variations.sort(key=lambda x: x[1], reverse=True)
                for col, match_count in variations[:3]:
                    print(f"   📋 '{col}' (matches: {match_count})")
            else:
                print(f"   ❌ No similar columns found")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def quick_fix_check(file_path):
    """Quick check to see what we're dealing with"""
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        print("🚀 QUICK ANALYSIS OF ALL SHEETS")
        print("=" * 60)
        
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # Basic info
                print(f"\n📋 {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
                
                # Show first few column names
                cols_preview = list(df.columns)[:5]
                print(f"   First 5 columns: {cols_preview}")
                
                # Check for key terms in column names
                col_text = ' '.join(str(col).lower() for col in df.columns)
                has_erf = 'erf' in col_text or 'status' in col_text
                has_entered = 'entered' in col_text or 'by' in col_text
                
                print(f"   Has ERF/Status terms: {'✅' if has_erf else '❌'}")
                print(f"   Has Entered/By terms: {'✅' if has_entered else '❌'}")
                
                if has_erf and has_entered:
                    print(f"   🎯 POTENTIAL CANDIDATE!")
                
            except Exception as e:
                print(f"\n📋 {sheet_name}: ❌ Error reading - {e}")
        
        # Ask user which sheet to debug
        print(f"\n💡 WHICH SHEET SHOULD WE DEBUG?")
        print("Based on the screenshot, 'Main data' looks promising with 2262 rows")
        
        target_sheet = input(f"\nEnter sheet name to debug (or press Enter for 'Main data'): ").strip()
        if not target_sheet:
            target_sheet = "Main data"
        
        if target_sheet in sheet_names:
            debug_sheet(file_path, target_sheet)
        else:
            print(f"❌ Sheet '{target_sheet}' not found")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("Enter Excel file path: ")
    
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
    else:
        quick_fix_check(file_path)