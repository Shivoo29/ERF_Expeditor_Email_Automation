# src/utils/chart_generator.py
"""Chart generation utilities for ERF data visualization"""
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
from datetime import datetime
import os
from typing import Dict, List, Tuple
import numpy as np

class ERFChartGenerator:
    """Generates charts for ERF data visualization"""
    
    def __init__(self, output_dir: str = "charts"):
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        
        # Set matplotlib style
        plt.style.use('default')
        sns.set_palette("husl")
    
    def generate_requester_summary_chart(self, requester_name: str, items_df: pd.DataFrame) -> str:
        """Generate a comprehensive summary chart for a requester"""
        
        # Create figure with subplots
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))
        fig.suptitle(f'ERF Status Summary for {requester_name}', fontsize=16, fontweight='bold')
        
        # 1. Status Distribution (Pie Chart)
        status_counts = items_df['ERF Sched Line Status'].value_counts()
        colors = ['#FF9999', '#66B2FF', '#99FF99', '#FFD700', '#FF6B6B']
        
        ax1.pie(status_counts.values, labels=status_counts.index, autopct='%1.1f%%', 
                startangle=90, colors=colors[:len(status_counts)])
        ax1.set_title(f'Status Distribution\n({len(items_df)} Total Items)', fontweight='bold')
        
        # 2. Items by Due Date (Timeline)
        if 'Due Date' in items_df.columns:
            due_dates = pd.to_datetime(items_df['Due Date'], errors='coerce').dropna()
            if not due_dates.empty:
                due_dates_grouped = due_dates.dt.to_period('M').value_counts().sort_index()
                
                ax2.bar(range(len(due_dates_grouped)), due_dates_grouped.values, 
                       color='skyblue', alpha=0.7)
                ax2.set_title('Items by Due Date (Monthly)', fontweight='bold')
                ax2.set_xlabel('Month')
                ax2.set_ylabel('Number of Items')
                
                # Format x-axis labels
                if len(due_dates_grouped) > 0:
                    labels = [str(period) for period in due_dates_grouped.index]
                    ax2.set_xticks(range(len(labels)))
                    ax2.set_xticklabels(labels, rotation=45, ha='right')
            else:
                ax2.text(0.5, 0.5, 'No valid due dates', ha='center', va='center', transform=ax2.transAxes)
                ax2.set_title('Items by Due Date', fontweight='bold')
        
        # 3. Top Materials (Horizontal Bar Chart)
        if 'Material' in items_df.columns:
            material_counts = items_df['Material'].value_counts().head(10)
            
            if not material_counts.empty:
                y_pos = np.arange(len(material_counts))
                ax3.barh(y_pos, material_counts.values, color='lightcoral', alpha=0.7)
                ax3.set_yticks(y_pos)
                ax3.set_yticklabels([str(mat)[:20] + '...' if len(str(mat)) > 20 else str(mat) 
                                   for mat in material_counts.index])
                ax3.set_xlabel('Count')
                ax3.set_title('Top 10 Materials', fontweight='bold')
                ax3.invert_yaxis()
            else:
                ax3.text(0.5, 0.5, 'No material data', ha='center', va='center', transform=ax3.transAxes)
                ax3.set_title('Top Materials', fontweight='bold')
        
        # 4. Quantity Distribution
        if 'ERF Itm Qty' in items_df.columns:
            quantities = pd.to_numeric(items_df['ERF Itm Qty'], errors='coerce').dropna()
            
            if not quantities.empty:
                ax4.hist(quantities, bins=min(20, len(quantities.unique())), 
                        color='lightgreen', alpha=0.7, edgecolor='black')
                ax4.set_title('Quantity Distribution', fontweight='bold')
                ax4.set_xlabel('Quantity')
                ax4.set_ylabel('Frequency')
                
                # Add summary stats
                stats_text = f'Total Qty: {quantities.sum():.0f}\nAvg: {quantities.mean():.1f}\nMax: {quantities.max():.0f}'
                ax4.text(0.7, 0.8, stats_text, transform=ax4.transAxes, 
                        bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.7))
            else:
                ax4.text(0.5, 0.5, 'No quantity data', ha='center', va='center', transform=ax4.transAxes)
                ax4.set_title('Quantity Distribution', fontweight='bold')
        
        # Adjust layout and save
        plt.tight_layout()
        
        # Save chart
        chart_filename = f"{requester_name}_ERF_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        chart_path = os.path.join(self.output_dir, chart_filename)
        plt.savefig(chart_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        return chart_path
    
    def generate_status_timeline_chart(self, requester_name: str, items_df: pd.DataFrame) -> str:
        """Generate a timeline chart showing status progression"""
        
        fig, ax = plt.subplots(figsize=(12, 6))
        
        # Create timeline data
        if 'Due Date' in items_df.columns and 'ERF Sched Line Status' in items_df.columns:
            df_copy = items_df.copy()
            df_copy['Due Date'] = pd.to_datetime(df_copy['Due Date'], errors='coerce')
            df_copy = df_copy.dropna(subset=['Due Date'])
            
            if not df_copy.empty:
                # Group by month and status
                df_copy['Month'] = df_copy['Due Date'].dt.to_period('M')
                timeline_data = df_copy.groupby(['Month', 'ERF Sched Line Status']).size().unstack(fill_value=0)
                
                # Create stacked bar chart
                timeline_data.plot(kind='bar', stacked=True, ax=ax, 
                                 color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7'])
                
                ax.set_title(f'ERF Timeline for {requester_name}', fontsize=14, fontweight='bold')
                ax.set_xlabel('Month', fontweight='bold')
                ax.set_ylabel('Number of Items', fontweight='bold')
                ax.legend(title='Status', bbox_to_anchor=(1.05, 1), loc='upper left')
                
                # Rotate x-axis labels
                ax.tick_params(axis='x', rotation=45)
            else:
                ax.text(0.5, 0.5, 'No timeline data available', ha='center', va='center', transform=ax.transAxes)
                ax.set_title(f'ERF Timeline for {requester_name}', fontsize=14, fontweight='bold')
        
        plt.tight_layout()
        
        # Save chart
        chart_filename = f"{requester_name}_timeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        chart_path = os.path.join(self.output_dir, chart_filename)
        plt.savefig(chart_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        return chart_path
    
    def generate_summary_table_image(self, requester_name: str, items_df: pd.DataFrame) -> str:
        """Generate a summary table as an image (for key metrics only)"""
        
        # Create summary data
        summary_data = {
            'Metric': ['Total Items', 'On Order', 'Received', 'Total Quantity', 'Unique Materials', 'Avg Quantity'],
            'Value': []
        }
        
        # Calculate metrics
        total_items = len(items_df)
        on_order = len(items_df[items_df['ERF Sched Line Status'] == 'On order'])
        received = len(items_df[items_df['ERF Sched Line Status'] == 'Received'])
        
        total_qty = 0
        avg_qty = 0
        if 'ERF Itm Qty' in items_df.columns:
            quantities = pd.to_numeric(items_df['ERF Itm Qty'], errors='coerce').dropna()
            total_qty = quantities.sum() if not quantities.empty else 0
            avg_qty = quantities.mean() if not quantities.empty else 0
        
        unique_materials = items_df['Material'].nunique() if 'Material' in items_df.columns else 0
        
        summary_data['Value'] = [
            total_items, on_order, received, f"{total_qty:.0f}", 
            unique_materials, f"{avg_qty:.1f}"
        ]
        
        # Create figure
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.axis('tight')
        ax.axis('off')
        
        # Create table
        table_data = list(zip(summary_data['Metric'], summary_data['Value']))
        table = ax.table(cellText=table_data, 
                        colLabels=['Metric', 'Value'],
                        cellLoc='center',
                        loc='center',
                        colWidths=[0.7, 0.3])
        
        # Style the table
        table.auto_set_font_size(False)
        table.set_fontsize(12)
        table.scale(1, 2)
        
        # Header styling
        for i in range(2):
            table[(0, i)].set_facecolor('#4CAF50')
            table[(0, i)].set_text_props(weight='bold', color='white')
        
        # Row styling
        for i in range(1, len(table_data) + 1):
            for j in range(2):
                if i % 2 == 0:
                    table[(i, j)].set_facecolor('#f0f0f0')
        
        plt.title(f'ERF Summary for {requester_name}', fontsize=14, fontweight='bold', pad=20)
        
        # Save chart
        chart_filename = f"{requester_name}_summary_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        chart_path = os.path.join(self.output_dir, chart_filename)
        plt.savefig(chart_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        return chart_path
    
    def cleanup_old_charts(self, days_old: int = 7):
        """Clean up chart files older than specified days"""
        import time
        
        current_time = time.time()
        for filename in os.listdir(self.output_dir):
            file_path = os.path.join(self.output_dir, filename)
            if os.path.isfile(file_path):
                file_age = current_time - os.path.getctime(file_path)
                if file_age > (days_old * 24 * 60 * 60):  # Convert days to seconds
                    try:
                        os.remove(file_path)
                    except Exception:
                        pass  # Ignore errors during cleanup