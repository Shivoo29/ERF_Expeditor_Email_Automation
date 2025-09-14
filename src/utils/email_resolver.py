# src/utils/email_resolver.py
"""Fast email resolution utility using simple dictionary lookup"""
import pandas as pd
import os
from typing import Optional, Dict, List
from src.utils.logger import setup_logger

class EmailResolver:
    """Fast email resolver using preloaded dictionary"""
    
    def __init__(self):
        self.logger = setup_logger(self.__class__.__name__)
        self.email_mapping = {}
        self.unmapped_users = set()  # Use set for faster lookup
        
        # HARDCODED PATH - Update this to your actual email mapping file path
        self.mapping_file_path = r"email_mapping_detailed_20250911_121356.xlsx"
        
        # Load once at startup
        self._load_mapping_fast()
    
    def _load_mapping_fast(self):
        """Load email mapping once using fast pandas operations"""
        if not os.path.exists(self.mapping_file_path):
            self.logger.warning(f"Email mapping file not found at: {self.mapping_file_path}")
            # Use fallback approach for missing mapping file
            self._setup_fallback_mode()
            return
        
        try:
            # Read Excel file once
            df = pd.read_excel(self.mapping_file_path)
            self.logger.info(f"Read Excel file with {len(df)} rows and {len(df.columns)} columns")
            
            # Find email column efficiently - look for '@' in column names or data
            email_col = None
            eng_col = df.columns[0]  # First column is usually Eng ID
            
            # Check column names first (faster)
            for col in df.columns:
                if '@' in str(col):
                    email_col = col
                    break
            
            # If not found in column names, check data (sample first few rows only)
            if not email_col:
                for col in df.columns[1:]:  # Skip first column (Eng ID)
                    sample = df[col].dropna().head(3).astype(str)
                    if any('@' in val for val in sample):
                        email_col = col
                        break
            
            if not email_col:
                self.logger.error("No email column found")
                self._setup_fallback_mode()
                return
            
            # Create mapping using vectorized operations (much faster)
            valid_rows = (
                df[eng_col].notna() & 
                df[email_col].notna() & 
                df[email_col].astype(str).str.contains('@', na=False) &
                df[eng_col].astype(str).str.strip().ne('') &
                df[eng_col].astype(str).str.upper().ne('ENG')
            )
            
            # Extract valid mappings
            valid_data = df[valid_rows]
            eng_ids = valid_data[eng_col].astype(str).str.strip().str.upper()
            emails = valid_data[email_col].astype(str).str.strip()
            
            # Create dictionary mapping
            self.email_mapping = dict(zip(eng_ids, emails))
            
            self.logger.info(f"Loaded {len(self.email_mapping)} email mappings from '{email_col}' column")
            
            # Show samples
            for i, (eng_id, email) in enumerate(list(self.email_mapping.items())[:3]):
                self.logger.info(f"  Sample {i+1}: {eng_id} -> {email}")
                
        except Exception as e:
            self.logger.error(f"Error loading email mapping: {e}")
            self._setup_fallback_mode()
    
    def _setup_fallback_mode(self):
        """Setup fallback mode when mapping file is incomplete or missing"""
        self.logger.info("Setting up fallback mode for email resolution")
        
        # Add some common known mappings if available
        # You can manually add known mappings here
        known_mappings = {
            # Add any known mappings manually
            # 'USERID': 'user.name@lamresearch.com',
        }
        
        self.email_mapping.update(known_mappings)
        self.logger.info(f"Fallback mode active with {len(self.email_mapping)} known mappings")
    
    def resolve_email(self, username: str) -> Optional[str]:
        """Fast email resolution using dictionary lookup"""
        if not username:
            return None
        
        # Clean and uppercase for consistent lookup
        clean_username = str(username).strip().upper()
        
        # Direct dictionary lookup (O(1) operation - very fast)
        if clean_username in self.email_mapping:
            email = self.email_mapping[clean_username]
            self.logger.info(f"Resolved {username} -> {email}")
            return email
        
        # Fast partial matching using dictionary iteration (only if needed)
        for eng_id, email in self.email_mapping.items():
            if clean_username in eng_id or eng_id in clean_username:
                self.logger.info(f"Partial match: {username} -> {email} (via {eng_id})")
                return email
        
        # Add to unmapped set for tracking
        self.unmapped_users.add(username)
        self.logger.warning(f"No email found for: {username}")
        return None
    
    def bulk_resolve_emails(self, usernames: List[str]) -> Dict[str, Optional[str]]:
        """Resolve multiple emails at once (more efficient)"""
        results = {}
        clean_usernames = [str(u).strip().upper() for u in usernames]
        
        # Bulk lookup using dictionary comprehension
        for original, clean in zip(usernames, clean_usernames):
            results[original] = self.email_mapping.get(clean)
            if not results[original]:
                self.unmapped_users.add(original)
        
        return results
    
    def get_unmapped_users(self) -> List[str]:
        """Get list of unmapped users"""
        return list(self.unmapped_users)
    
    def add_manual_mapping(self, username: str, email: str):
        """Add manual mapping for missing users"""
        if username and email and '@' in email:
            clean_username = str(username).strip().upper()
            self.email_mapping[clean_username] = email.strip()
            # Remove from unmapped if it was there
            self.unmapped_users.discard(username)
            self.logger.info(f"Added manual mapping: {clean_username} -> {email}")
    
    def export_unmapped_users(self, output_file: str):
        """Export unmapped users for manual completion"""
        if not self.unmapped_users:
            self.logger.info("No unmapped users to export")
            return
        
        try:
            # Create template for manual completion
            unmapped_df = pd.DataFrame({
                'Username': list(self.unmapped_users),
                'Email': '',  # Empty for manual filling
                'Full_Name': '',  # Empty for manual filling
                'Status': 'NEEDS_EMAIL',
                'Instructions': 'Please fill in Email column manually'
            })
            
            unmapped_df.to_excel(output_file, index=False)
            self.logger.info(f"Exported {len(self.unmapped_users)} unmapped users to: {output_file}")
            
            print(f"\n📋 INCOMPLETE EMAIL MAPPING DETECTED")
            print(f"   Missing emails for {len(self.unmapped_users)} users")
            print(f"   Exported template to: {output_file}")
            print(f"   Please manually fill in the Email column and reload")
            
        except Exception as e:
            self.logger.error(f"Error exporting unmapped users: {e}")
    
    def load_manual_mappings(self, manual_file: str):
        """Load manually completed mappings"""
        if not os.path.exists(manual_file):
            return
        
        try:
            df = pd.read_excel(manual_file)
            if 'Username' in df.columns and 'Email' in df.columns:
                for _, row in df.iterrows():
                    username = str(row['Username']).strip()
                    email = str(row['Email']).strip()
                    if username and email and '@' in email:
                        self.add_manual_mapping(username, email)
                
                self.logger.info(f"Loaded additional manual mappings from {manual_file}")
        except Exception as e:
            self.logger.error(f"Error loading manual mappings: {e}")
    
    def get_mapping_stats(self) -> Dict:
        """Get mapping statistics"""
        return {
            'total_mappings': len(self.email_mapping),
            'unmapped_users_count': len(self.unmapped_users),
            'unmapped_users': list(self.unmapped_users),
            'coverage_percentage': 0 if not self.unmapped_users else 
                round((len(self.email_mapping) / (len(self.email_mapping) + len(self.unmapped_users))) * 100, 1),
            'sample_mappings': dict(list(self.email_mapping.items())[:5])
        }