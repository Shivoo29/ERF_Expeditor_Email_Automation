import pandas as pd
import sys
import os

def analyze_sheet(df, sheet_name):
    """Analyze a specific sheet"""
    print(f"\n📋 SHEET: '{sheet_name}'")
    print("-" * 60)
    
    if df.empty:
        print("   ❌ Sheet is empty")
        return 0, False
    
    # Check if it looks like a pivot table
    is_pivot = is_pivot_table(df)
    if is_pivot:
        print("   ❌ Appears to be a pivot table or summary")
        return 0, False
    
    print(f"   📊 Rows: {len(df)}")
    print(f"   📋 Columns: {len(df.columns)}")
    
    # Check for critical columns
    has_status = 'ERF Sched Line Status' in df.columns
    has_entered_by = 'Entered by' in df.columns
    
    print(f"   🎯 Has 'ERF Sched Line Status': {'✅' if has_status else '❌'}")
    print(f"   🎯 Has 'Entered by': {'✅' if has_entered_by else '❌'}")
    
    # Score the sheet
    required_columns = [
        'Plnt', 'Ship-To-Plant', 'ERF Nr', 'Item', 'Entered by',
        'Material', 'Material Description', 'Unit', 'ERF Itm Qty',
        'Date Req.', 'ERF Sched Line Status', 'Due Date', 'Expeditor',
        'Expeditor Status', 'ETA', 'Expeditor Remarks', 'END'
    ]
    
    score = 0
    found_columns = []
    missing_columns = []
    
    for col in required_columns:
        if col in df.columns:
            score += 1
            found_columns.append(col)
        else:
            missing_columns.append(col)
    
    print(f"   📈 Score: {score}/{len(required_columns)} required columns found")
    
    # Show column details
    print(f"\n   📝 ALL COLUMNS IN THIS SHEET:")
    for i, col in enumerate(df.columns, 1):
        status = "✅" if col in required_columns else "📋"
        print(f"      {i:2d}. {status} '{col}'")
    
    if found_columns:
        print(f"\n   ✅ FOUND REQUIRED COLUMNS ({len(found_columns)}):")
        for col in found_columns[:10]:  # Show first 10
            print(f"      • '{col}'")
        if len(found_columns) > 10:
            print(f"      ... and {len(found_columns) - 10} more")
    
    if missing_columns:
        print(f"\n   ❌ MISSING REQUIRED COLUMNS ({len(missing_columns)}):")
        for col in missing_columns[:10]:  # Show first 10
            print(f"      • '{col}'")
        if len(missing_columns) > 10:
            print(f"      ... and {len(missing_columns) - 10} more")
    
    # Show sample data if it looks good
    if has_status and has_entered_by and score > 5:
        print(f"\n   📋 SAMPLE DATA (first 3 rows):")
        print("   " + "-" * 50)
        sample_cols = ['ERF Sched Line Status', 'Entered by']
        if 'ERF Nr' in df.columns:
            sample_cols.append('ERF Nr')
        if 'Material' in df.columns:
            sample_cols.append('Material')
        
        sample_data = df[sample_cols].head(3)
        for idx, row in sample_data.iterrows():
            print(f"   Row {idx + 1}: {dict(row)}")
    
    return score, (has_status and has_entered_by)

def is_pivot_table(df):
    """Check if the dataframe looks like a pivot table"""
    if len(df) < 3:
        return False
    
    # Convert to string and check for common pivot table patterns
    first_rows = df.head(5).astype(str)
    
    # Look for pivot table indicators
    pivot_indicators = [
        'Column Labels', 'Row Labels', 'Count of', 'Sum of', 
        'Grand Total', 'Unnamed:', 'nan'
    ]
    
    for indicator in pivot_indicators:
        if any(indicator in str(cell) for row in first_rows.values for cell in row):
            return True
    
    # Check if first row is all NaN
    if df.iloc[0].isna().all():
        return True
    
    # Check if there are many unnamed columns
    unnamed_cols = [col for col in df.columns if 'Unnamed:' in str(col)]
    if len(unnamed_cols) > len(df.columns) * 0.5:
        return True
    
    return False

def check_excel_file(file_path):
    """Check all sheets in Excel file and find the best one for ERF data"""
    try:
        print("🔍 MULTI-SHEET EXCEL FILE ANALYSIS")
        print("=" * 60)
        print(f"📁 File: {file_path}")
        
        # Get all sheet names
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        print(f"📊 Total sheets found: {len(sheet_names)}")
        print(f"📋 Sheet names: {sheet_names}")
        
        best_sheet = None
        best_score = 0
        sheet_results = []
        
        # Analyze each sheet
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                score, is_suitable = analyze_sheet(df, sheet_name)
                
                sheet_results.append({
                    'name': sheet_name,
                    'score': score,
                    'suitable': is_suitable,
                    'rows': len(df) if not df.empty else 0
                })
                
                if is_suitable and score > best_score:
                    best_sheet = sheet_name
                    best_score = score
                    
            except Exception as e:
                print(f"\n📋 SHEET: '{sheet_name}'")
                print("-" * 60)
                print(f"   ❌ Error reading sheet: {e}")
        
        # Summary
        print(f"\n🎯 ANALYSIS SUMMARY")
        print("=" * 60)
        
        print(f"📊 SHEET SCORES:")
        for result in sorted(sheet_results, key=lambda x: x['score'], reverse=True):
            status = "🎯 BEST" if result['name'] == best_sheet else "✅ GOOD" if result['suitable'] else "❌ SKIP"
            print(f"   {status} '{result['name']}' - Score: {result['score']}/16, Rows: {result['rows']}")
        
        if best_sheet:
            print(f"\n🎉 RECOMMENDED SHEET: '{best_sheet}'")
            print(f"   📈 Score: {best_score}/16 required columns")
            print(f"   ✅ This sheet should work with the ERF automation system")
            
            # Show what the system will do
            print(f"\n🚀 WHAT WILL HAPPEN:")
            print(f"   1. System will automatically select sheet '{best_sheet}'")
            print(f"   2. Load {sheet_results[[r['name'] for r in sheet_results].index(best_sheet)]['rows']} rows of data")
            print(f"   3. Filter for 'On order' and 'Received' status items")
            print(f"   4. Group by 'Entered by' field")
            print(f"   5. Generate emails for each requester")
            
        else:
            print(f"\n❌ NO SUITABLE SHEET FOUND")
            print(f"   📋 None of the sheets contain the required ERF data structure")
            print(f"   🔧 Required: 'ERF Sched Line Status' and 'Entered by' columns")
            
        print(f"\n💡 NEXT STEPS:")
        if best_sheet:
            print(f"   ✅ Run the automation system - it will work with your file!")
            print(f"   📝 Command: python main.py \"{file_path}\"")
        else:
            print(f"   📥 Get the raw ERF data export (not pivot tables)")
            print(f"   📋 Ensure data has individual ERF line items")
            print(f"   🔧 Required columns: 'ERF Sched Line Status', 'Entered by'")
        
        return best_sheet is not None
        
    except Exception as e:
        print(f"❌ Error analyzing Excel file: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("Enter Excel file path: ")
    
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
    else:
        check_excel_file(file_path)