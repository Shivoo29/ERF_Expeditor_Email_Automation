# src/data/excel_processor.py - FIXED VERSION
"""Excel data processing module with improved multi-sheet support"""
import pandas as pd
from typing import Dict, Any, Tuple, Optional
from src.utils.logger import setup_logger
from src.utils.validators import validate_excel_file
from config.settings import Config

class ExcelProcessor:
    """Handles Excel file processing and data filtering with multi-sheet support"""
    
    def __init__(self):
        self.logger = setup_logger(self.__class__.__name__)
        self.raw_data = None
        self.filtered_data = None
        self.grouped_data = None
        self.selected_sheet = None
    
    def find_data_sheet(self, file_path: str) -> Tuple[bool, Optional[str], Optional[pd.DataFrame]]:
        """Find the sheet that contains actual ERF data"""
        try:
            # Get all sheet names
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            self.logger.info(f"Found {len(sheet_names)} sheets: {sheet_names}")
            
            best_sheet = None
            best_score = 0
            best_data = None
            
            for sheet_name in sheet_names:
                self.logger.info(f"Analyzing sheet: '{sheet_name}'")
                
                try:
                    # Read the sheet
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    
                    # Skip empty sheets
                    if df.empty:
                        self.logger.info(f"  ❌ Sheet '{sheet_name}' is empty")
                        continue
                    
                    # Check for critical columns first (this is the key fix!)
                    has_status = 'ERF Sched Line Status' in df.columns
                    has_entered_by = 'Entered by' in df.columns
                    
                    if not (has_status and has_entered_by):
                        self.logger.info(f"  ❌ Sheet '{sheet_name}' missing critical columns")
                        missing = []
                        if not has_status:
                            missing.append("'ERF Sched Line Status'")
                        if not has_entered_by:
                            missing.append("'Entered by'")
                        self.logger.info(f"      Missing: {', '.join(missing)}")
                        continue
                    
                    # If we have critical columns, check if it's a real pivot table
                    if self._is_real_pivot_table(df):
                        self.logger.info(f"  ❌ Sheet '{sheet_name}' appears to be a pivot table")
                        continue
                    
                    # Score the sheet based on how many required columns it has
                    score = self._score_sheet(df)
                    self.logger.info(f"  📊 Sheet '{sheet_name}' score: {score}/{len(Config.REQUIRED_COLUMNS)}")
                    self.logger.info(f"  ✅ Sheet '{sheet_name}' has critical columns and real data")
                    
                    if score > best_score:
                        best_sheet = sheet_name
                        best_score = score
                        best_data = df
                
                except Exception as e:
                    self.logger.warning(f"  ❌ Error reading sheet '{sheet_name}': {str(e)}")
                    continue
            
            if best_sheet:
                self.logger.info(f"🎯 Selected sheet: '{best_sheet}' with score {best_score}")
                return True, best_sheet, best_data
            else:
                self.logger.error("❌ No suitable data sheet found")
                return False, None, None
                
        except Exception as e:
            self.logger.error(f"Error analyzing Excel file: {str(e)}")
            return False, None, None
    
    def _is_real_pivot_table(self, df: pd.DataFrame) -> bool:
        """More accurate check if the dataframe is actually a pivot table"""
        if len(df) < 3:
            return False
        
        # Check if more than 70% of columns are unnamed (strong indicator)
        unnamed_cols = [col for col in df.columns if 'Unnamed:' in str(col)]
        if len(unnamed_cols) > len(df.columns) * 0.7:
            return True
        
        # Check if first row is mostly NaN (common in pivot tables)
        if df.iloc[0].isna().sum() > len(df.columns) * 0.8:
            return True
        
        # Check for specific pivot table patterns in the data
        first_few_cells = []
        for i in range(min(5, len(df))):
            for j in range(min(5, len(df.columns))):
                cell_value = str(df.iloc[i, j]).lower()
                first_few_cells.append(cell_value)
        
        cell_text = ' '.join(first_few_cells)
        pivot_indicators = ['column labels', 'row labels', 'count of', 'sum of', 'grand total']
        
        for indicator in pivot_indicators:
            if indicator in cell_text:
                return True
        
        # If we have good columns and actual data, it's probably not a pivot table
        return False
    
    def _score_sheet(self, df: pd.DataFrame) -> int:
        """Score a sheet based on how many required columns it has"""
        score = 0
        df_columns = [str(col).strip() for col in df.columns]
        
        for required_col in Config.REQUIRED_COLUMNS:
            if required_col in df_columns:
                score += 1
        
        return score
    
    def load_file(self, file_path: str) -> bool:
        """Load Excel file and find the correct sheet with data"""
        self.logger.info(f"Loading Excel file: {file_path}")
        
        # Validate file
        is_valid, message = validate_excel_file(file_path)
        if not is_valid:
            self.logger.error(message)
            return False
        
        try:
            # Find the best data sheet
            found_sheet, sheet_name, sheet_data = self.find_data_sheet(file_path)
            
            if not found_sheet:
                self.logger.error("Could not find a suitable data sheet")
                self.logger.error("Please ensure your Excel file contains raw ERF data with the required columns")
                return False
            
            # Use the found sheet data
            self.raw_data = sheet_data
            self.selected_sheet = sheet_name
            
            self.logger.info(f"Successfully loaded {len(self.raw_data)} rows from sheet '{sheet_name}'")
            
            # Show available columns
            self.logger.info(f"Found {len(self.raw_data.columns)} columns in selected sheet")
            
            # Check which required columns are missing
            missing_columns = []
            available_columns = list(self.raw_data.columns)
            
            for required_col in Config.REQUIRED_COLUMNS:
                if required_col not in available_columns:
                    missing_columns.append(required_col)
            
            if missing_columns:
                self.logger.warning(f"Missing optional columns: {missing_columns}")
            else:
                self.logger.info("All required columns found!")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error loading Excel file: {str(e)}")
            return False
    
    def filter_data(self) -> bool:
        """Filter data for target statuses"""
        if self.raw_data is None:
            self.logger.error("No data loaded. Call load_file() first.")
            return False
        
        try:
            # Check if ERF Sched Line Status column exists
            if 'ERF Sched Line Status' not in self.raw_data.columns:
                self.logger.error("Cannot filter: 'ERF Sched Line Status' column not found")
                return False
            
            # Show available statuses first
            available_statuses = self.raw_data['ERF Sched Line Status'].dropna().unique()
            self.logger.info(f"Available statuses in data: {list(available_statuses)}")
            self.logger.info(f"Looking for statuses: {Config.TARGET_STATUSES}")
            
            # Filter for required statuses
            status_filter = self.raw_data['ERF Sched Line Status'].isin(Config.TARGET_STATUSES)
            self.filtered_data = self.raw_data[status_filter].copy()
            
            self.logger.info(f"Filtered to {len(self.filtered_data)} items with target statuses")
            
            if len(self.filtered_data) == 0:
                self.logger.warning("No items found with target statuses")
                self.logger.info(f"Available statuses: {list(available_statuses)}")
                self.logger.info("Consider updating target statuses in config if needed")
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error filtering data: {str(e)}")
            return False
    
    def group_by_requester(self) -> bool:
        """Group filtered data by 'Entered by' field"""
        if self.filtered_data is None:
            self.logger.error("No filtered data available. Call filter_data() first.")
            return False
        
        try:
            # Check if Entered by column exists
            if 'Entered by' not in self.filtered_data.columns:
                self.logger.error("Cannot group: 'Entered by' column not found")
                return False
            
            # Remove rows where 'Entered by' is empty/null
            clean_data = self.filtered_data.dropna(subset=['Entered by'])
            clean_data = clean_data[clean_data['Entered by'].str.strip() != '']
            
            if len(clean_data) == 0:
                self.logger.error("No valid 'Entered by' entries found")
                return False
            
            self.grouped_data = clean_data.groupby('Entered by')
            
            requesters = list(self.grouped_data.groups.keys())
            self.logger.info(f"Found {len(requesters)} unique requesters")
            
            for requester in requesters[:10]:  # Show first 10
                count = len(self.grouped_data.get_group(requester))
                self.logger.info(f"  - {requester}: {count} items")
            
            if len(requesters) > 10:
                self.logger.info(f"  ... and {len(requesters) - 10} more requesters")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error grouping data: {str(e)}")
            return False
    
    def get_grouped_data(self) -> Dict[str, pd.DataFrame]:
        """Return grouped data as dictionary"""
        if self.grouped_data is None:
            return {}
        
        return {name: group for name, group in self.grouped_data}
    
    def get_summary(self) -> Dict[str, Any]:
        """Get processing summary"""
        summary = {
            'selected_sheet': self.selected_sheet,
            'total_rows': len(self.raw_data) if self.raw_data is not None else 0,
            'filtered_rows': len(self.filtered_data) if self.filtered_data is not None else 0,
            'unique_requesters': len(self.grouped_data.groups) if self.grouped_data is not None else 0,
            'status_breakdown': {}
        }
        
        if self.filtered_data is not None and 'ERF Sched Line Status' in self.filtered_data.columns:
            summary['status_breakdown'] = self.filtered_data['ERF Sched Line Status'].value_counts().to_dict()
        
        return summary