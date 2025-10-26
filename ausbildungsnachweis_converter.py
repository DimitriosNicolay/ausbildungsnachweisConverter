#!/usr/bin/env python3
"""
Ausbildungsnachweis Excel to Markdown Converter

Converts German vocational training reports (Ausbildungsnachweise) 
from Excel format to Markdown format suitable for Obsidian.

Author: Dimitrios Nicolay
Date: 2025-10-26
"""

import re
import zipfile
from pathlib import Path
from datetime import datetime
from typing import Tuple, Optional
import pandas as pd


class AusbildungsnachweisConverter:
    """Handles conversion of Ausbildungsnachweis Excel files to Markdown format."""
    
    def __init__(self, default_year: str = "2025"):
        self.default_year = default_year
        
    def parse_dates_from_filename(self, filename: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """Extract start and end dates from Excel filename."""
        base = Path(filename).stem
        
        # Match various date formats in filename
        m = re.search(r'_(\d{2}\.\d{2}\.?)(?:\.\d{2})?\s*[-]\s*(\d{2}[.\-]\d{2}\.?)(?:\.\d{2})?', base)
        
        if not m:
            return None, None, None
            
        start_raw, end_raw = m.group(1), m.group(2)
        start_raw = start_raw.rstrip('.')
        end_raw = end_raw.replace('-', '.').rstrip('.')
        
        return start_raw, end_raw, base
    
    def parse_date(self, date_str: str, default_year: str) -> datetime:
        """Parse date string with optional year component."""
        parts = date_str.split('.')
        
        if len(parts) == 3:
            return datetime.strptime(date_str, "%d.%m.%y")
        elif len(parts) == 2:
            return datetime.strptime(f"{date_str}.{default_year}", "%d.%m.%Y")
        else:
            raise ValueError(f"Invalid date format: {date_str}")
    
    def clean_text(self, text) -> str:
        """Remove extra whitespace from text."""
        if pd.isna(text):
            return ""
        text = str(text).strip()
        return " ".join(text.split())
    
    def convert_excel_to_markdown(self, excel_path: Path) -> Tuple[str, str]:
        """
        Convert single Excel file to Markdown format.
        
        Returns:
            Tuple of (output_filename, markdown_content)
        """
        start_raw, end_raw, base = self.parse_dates_from_filename(excel_path.name)
        if not start_raw:
            raise ValueError(f"Cannot parse dates from filename: {excel_path.name}")
        
        # Read Excel file to extract metadata
        df = pd.read_excel(excel_path, header=None)
        
        # Extract year from Excel metadata or filename
        actual_year = self.default_year
        if df.shape[1] > 11 and not pd.isna(df.iloc[2, 11]):
            try:
                actual_year = str(int(df.iloc[2, 11]))
            except:
                actual_year = "2024" if ".24" in excel_path.name else self.default_year
        elif ".24" in excel_path.name:
            actual_year = "2024"
        
        start_dt = self.parse_date(start_raw, actual_year)
        end_dt = self.parse_date(end_raw, actual_year)
        
        iso_year, iso_week, _ = start_dt.isocalendar()
        
        # Extract header metadata
        name = self.clean_text(df.iloc[0, 7]) if df.shape[1] > 7 else "Nicolay, Dimitrios"
        if not name:
            name = "Nicolay, Dimitrios"
            
        course = self.clean_text(df.iloc[1, 7]) if df.shape[1] > 7 else "Fachinformatiker SI - U27B (IHK)"
        if not course:
            course = "Fachinformatiker SI - U27B (IHK)"
        
        year = actual_year
        
        # Locate data start row
        data_start_row = None
        for i in range(len(df)):
            cell = df.iloc[i, 1] if df.shape[1] > 1 else None
            if isinstance(cell, str) and "Tag" in cell:
                data_start_row = i + 1
                break
        
        if data_start_row is None:
            data_start_row = 0
        
        # Build Markdown document
        md_lines = []
        md_lines.append(f"# KW{iso_week:02d} - Ausbildungsnachweis ({start_dt.strftime('%d.%m.%Y')} - {end_dt.strftime('%d.%m.%Y')})")
        md_lines.append("")
        md_lines.append(f"**Name:** {name}  ")
        md_lines.append(f"**Ausbildung:** {course}  ")
        md_lines.append(f"**Jahr:** {year}  ")
        md_lines.append("")
        md_lines.append("---")
        md_lines.append("")
        
        current_date = None
        daily_activities = []
        pending_hours = None
        
        def flush_day():
            """Write accumulated daily activities to output."""
            nonlocal current_date, daily_activities, pending_hours
            
            if not current_date or not daily_activities:
                return
                
            md_lines.append(f"## {current_date.strftime('%d.%m.%Y')}")
            md_lines.append("")
            
            for activity in daily_activities:
                md_lines.append(f"- {activity}")
            
            if pending_hours is not None:
                md_lines.append("")
                md_lines.append(f"**Stunden:** {int(pending_hours)}")
            
            md_lines.append("")
        
        # Process data rows
        for i in range(data_start_row, len(df)):
            row = df.iloc[i, :]
            
            date_cell = row[1] if len(row) > 1 else None
            activity_cell = row[2] if len(row) > 2 else None
            hours_cell = row[11] if len(row) > 11 else None
            
            if pd.notna(date_cell) and isinstance(date_cell, (pd.Timestamp, datetime)):
                flush_day()
                current_date = pd.to_datetime(date_cell).to_pydatetime()
                daily_activities = []
                pending_hours = None
                continue
            
            if pd.notna(activity_cell):
                activity_text = self.clean_text(activity_cell)
                if activity_text:
                    daily_activities.append(activity_text)
            
            if pd.notna(hours_cell) and current_date:
                try:
                    pending_hours = float(hours_cell)
                except:
                    pending_hours = None
                flush_day()
                current_date = None
                daily_activities = []
                pending_hours = None
        
        flush_day()
        
        out_filename = f"{start_dt.year}-KW{iso_week:02d}-Ausbildungsnachweis-{start_dt.strftime('%d.%m')}-{end_dt.strftime('%d.%m')}.md"
        markdown_content = "\n".join(md_lines).strip()
        
        return out_filename, markdown_content


def main():
    """Execute conversion workflow."""
    converter = AusbildungsnachweisConverter()
    
    input_dir = Path('input')
    output_dir = Path('output')
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)
    
    excel_files = sorted(input_dir.glob('AusbildungsnachweisU27_*.xlsx'))
    
    if not excel_files:
        print("No Excel files found in 'input' directory.")
        print("Please place your Excel files there and run again.")
        return
    
    print(f"Processing {len(excel_files)} file(s)\n")
    print("=" * 70)
    
    converted_files = []
    failed_files = []
    
    for excel_path in excel_files:
        try:
            md_filename, md_content = converter.convert_excel_to_markdown(excel_path)
            output_path = output_dir / md_filename
            output_path.write_text(md_content, encoding='utf-8')
            converted_files.append(md_filename)
            print(f"[OK] {excel_path.name} -> {md_filename}")
            
        except Exception as e:
            failed_files.append((excel_path.name, str(e)))
            print(f"[ERROR] {excel_path.name}: {e}")
    
    print("=" * 70)
    
    if converted_files:
        zip_path = output_dir / "Ausbildungsnachweise-Markdown.zip"
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for md_file in sorted(converted_files):
                zipf.write(output_dir / md_file, md_file)
        
        print(f"\nCreated: {zip_path}")
        print(f"Total converted: {len(converted_files)} file(s)")
    
    if failed_files:
        print(f"\nFailed: {len(failed_files)} file(s)")
        for filename, error in failed_files:
            print(f"  - {filename}: {error}")
    
    print(f"\nSummary: {len(converted_files)} successful, {len(failed_files)} failed")


if __name__ == "__main__":
    main()
