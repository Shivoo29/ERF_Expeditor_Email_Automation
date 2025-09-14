# src/email/email_templates.py
"""Email template generation with improved formatting and Expeditor Remarks handling"""
import pandas as pd
from datetime import datetime
from typing import Dict, Any

class EmailTemplate:
    """Handles email template generation"""
    
    @staticmethod
    def generate_status_email(requester_name: str, items_df: pd.DataFrame) -> Dict[str, str]:
        """Generate email subject and body for status update"""
        
        # Count items by status
        on_order_count = len(items_df[items_df['ERF Sched Line Status'] == 'On order'])
        received_count = len(items_df[items_df['ERF Sched Line Status'] == 'Received'])
        total_items = len(items_df)
        
        # Generate subject
        subject = f"ERF Status Update - {total_items} Items"
        
        # Generate HTML table with all required columns including Expeditor Remarks
        html_table = EmailTemplate._generate_html_table(items_df)
        
        # Generate body
        body = f"""Hello {requester_name},

I hope this email finds you well. This is an automated status update for your ERF items.

SUMMARY:
• Items On Order: {on_order_count}
• Items Received: {received_count}
• Total Items: {total_items}

Please find the detailed information in the table below:

{html_table}

If you have any questions or concerns regarding these items, please don't hesitate to reach out.

IMPORTANT NOTICE:
This is a Proof of Concept (POC) system active until Monday. For any queries related to these automated reports, please contact:
• Hemanth.Mathad@lamresearch.com (Mathad, Hemanth)
• Kishor.Ghantani@lamresearch.com (Ghantani, Kishor)

Best regards,
Proto4Lab Team
Lam Research

---
This is an automated email generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""
        
        return {
            'subject': subject,
            'body': body
        }
    
    @staticmethod
    def _generate_html_table(items_df: pd.DataFrame) -> str:
        """Generate an HTML table for the ERF items with better Expeditor Remarks handling"""
        
        # Select key columns including Expeditor Remarks
        display_columns = [
            'ERF Nr', 'Material', 'Material Description', 'ERF Itm Qty', 'Unit',
            'ERF Sched Line Status', 'END', 'PO Due Date', 'Expeditor', 
            'Expeditor Status', 'Expeditor Remarks'
        ]
        
        # Column name mapping for display purposes only
        column_display_names = {
            'ERF Nr': 'ERF Nr',
            'Material': 'Material',
            'Material Description': 'Material Description',
            'ERF Itm Qty': 'ERF Itm Qty',
            'Unit': 'Unit',
            'END': 'END',
            'ERF Sched Line Status': 'ERF Sched Line Status', 
            'Expeditor': 'Expeditor',
            'Expeditor Status': 'Expeditor Status',
            'PO Due Date': 'Commit Date',  
            'Expeditor Remarks': 'Expeditor Remarks',
        }
        
        # Filter to only include columns that exist in the dataframe
        available_columns = [col for col in display_columns if col in items_df.columns]
        
        # Create the table data
        table_data = items_df[available_columns].fillna('N/A')
        
        # Start building HTML with improved styling
        html = """
<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; width: 100%; table-layout: fixed;">
    <thead>
        <tr style="background-color: #4CAF50; color: white; font-weight: bold;">
"""
        
        # Add headers with specific column widths using display names
        for col in available_columns:
            display_name = column_display_names.get(col, col)
            
            # Give Expeditor Remarks more width
            if col == 'Expeditor Remarks':
                width_style = 'width: 200px;'
            elif col in ['ERF Nr', 'Unit', 'ERF Itm Qty']:
                width_style = 'width: 80px;'
            elif col in ['Material', 'Material Description']:
                width_style = 'width: 150px;'
            else:
                width_style = 'width: 120px;'
                
            html += f'            <th style="text-align: left; padding: 8px; border: 1px solid #ddd; {width_style}">{display_name}</th>\n'
        
        html += """        </tr>
    </thead>
    <tbody>
"""
        
        # Add data rows
        for idx, (_, row) in enumerate(table_data.iterrows()):
            # Alternate row colors
            bg_color = "#f9f9f9" if idx % 2 == 0 else "#ffffff"
            
            # Color code based on status
            status = row.get('ERF Sched Line Status', '')
            if status == 'On order':
                status_color = "#FFF3CD"  # Light yellow
            elif status == 'Received':
                status_color = "#D4EDDA"  # Light green
            else:
                status_color = bg_color
            
            html += f'        <tr style="background-color: {status_color};">\n'
            
            for col in available_columns:
                value = str(row.get(col, 'N/A'))
                
                # Special handling for different columns
                if col == 'Expeditor Remarks':
                    # Keep full expeditor remarks with better wrapping
                    if len(value) > 150:
                        value = value[:147] + "..."
                    cell_style = 'padding: 6px; border: 1px solid #ddd; text-align: left; word-wrap: break-word; white-space: normal; max-width: 200px;'
                elif col in ['ERF Nr', 'ERF Itm Qty', 'Unit']:
                    # Keep these short columns compact
                    if len(value) > 20:
                        value = value[:17] + "..."
                    cell_style = 'padding: 6px; border: 1px solid #ddd; text-align: center;'
                else:
                    # Regular columns
                    if len(value) > 50:
                        value = value[:47] + "..."
                    cell_style = 'padding: 6px; border: 1px solid #ddd; text-align: left;'
                
                html += f'            <td style="{cell_style}">{value}</td>\n'
            
            html += '        </tr>\n'
        
        html += """    </tbody>
</table>

<p style="font-size: 11px; color: #666; margin-top: 10px;">
<strong>Legend:</strong> 
<span style="background-color: #FFF3CD; padding: 2px 4px;">On Order</span>
<span style="background-color: #D4EDDA; padding: 2px 4px;">Received</span>
</p>
"""
        
        return html