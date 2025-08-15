#!/usr/bin/env python3
"""
Triangle Agency Field Agent Dossier Generator
Creates HTML dossiers for field agents from spreadsheet data
"""

import pandas as pd
import os
from datetime import datetime
import argparse
import base64
from PIL import Image
import io

def load_agent_data(file_path):
    """Load agent data from Excel, ODS, or CSV file"""
    try:
        # Check file extension
        if file_path.lower().endswith('.ods'):
            df = pd.read_excel(file_path, engine='odf')
            print(f"Loaded ODS file: {file_path}")
        elif file_path.lower().endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path)
            print(f"Loaded Excel file: {file_path}")
            if file_path.lower().endswith('.xls'):
                print("WARNING: .xls format has a 256 character limit per cell. Consider using .ods or .xlsx format.")
        else:
            # Try different encodings for CSV files with unlimited field size
            encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
            df = None
            for encoding in encodings:
                try:
                    # Set csv.field_size_limit to maximum to handle long fields
                    import csv
                    csv.field_size_limit(1000000)  # Set to 1 million characters
                    df = pd.read_csv(file_path, encoding=encoding)
                    print(f"Loaded CSV file with {encoding} encoding: {file_path}")
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                raise Exception("Could not decode CSV file with any standard encoding")
        
        # Debug: Check for truncated fields
        for col in df.columns:
            for idx, value in enumerate(df[col]):
                if isinstance(value, str) and len(value) == 256:
                    print(f"WARNING: Field '{col}' in row {idx} is exactly 256 characters - may be truncated!")
                elif isinstance(value, str) and len(value) > 500:
                    print(f"INFO: Field '{col}' in row {idx} has {len(value)} characters (looks complete)")
        
        return df
    except Exception as e:
        print(f"Error loading file: {e}")
        print("If using .ods format, make sure you have 'odfpy' installed: pip install odfpy")
        return None

def sanitize_filename(name):
    """Convert agent name to safe filename format"""
    safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in name).strip()
    safe_name = safe_name.replace(' ', '_')
    # Remove any double underscores that might have been created
    while '__' in safe_name:
        safe_name = safe_name.replace('__', '_')
    return safe_name

def load_and_process_photo(agent_name, photos_dir="photos"):
    """Load agent photo, crop to fit, and convert to base64"""
    try:
        # Generate filename from agent name
        safe_name = sanitize_filename(agent_name)
        photo_path = os.path.join(photos_dir, f"{safe_name}.jpg")
        
        if not os.path.exists(photo_path):
            print(f"No photo found for {agent_name} at {photo_path}")
            return None
            
        # Load image
        with Image.open(photo_path) as img:
            # Convert to RGB if necessary (handles RGBA, etc.)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Target dimensions (matching the CSS: 150px width, 200px height)
            target_width = 150
            target_height = 200
            target_ratio = target_width / target_height  # 0.75
            
            # Get current dimensions
            current_width, current_height = img.size
            current_ratio = current_width / current_height
            
            # Crop to center if aspect ratios don't match
            if abs(current_ratio - target_ratio) > 0.01:  # Small tolerance for floating point
                if current_ratio > target_ratio:
                    # Image is too wide, crop sides
                    new_width = int(current_height * target_ratio)
                    left = (current_width - new_width) // 2
                    img = img.crop((left, 0, left + new_width, current_height))
                else:
                    # Image is too tall, crop top and bottom
                    new_height = int(current_width / target_ratio)
                    top = (current_height - new_height) // 2
                    img = img.crop((0, top, current_width, top + new_height))
            
            # Resize to exact target dimensions
            img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
            
            # Convert to base64
            buffer = io.BytesIO()
            img.save(buffer, format='JPEG', quality=85)
            img_data = buffer.getvalue()
            base64_string = base64.b64encode(img_data).decode('utf-8')
            
            print(f"Processed photo for {agent_name}")
            return f"data:image/jpeg;base64,{base64_string}"
            
    except Exception as e:
        print(f"Error processing photo for {agent_name}: {e}")
        return None
    """Load HTML template from file"""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error loading template file '{template_path}': {e}")
        return None

def load_template(template_path):
    """Load HTML template from file"""
    try:
        # Try UTF-8 first (most common for HTML files)
        with open(template_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Check if the triangle symbol is properly loaded
            if '▲' in content or '△' in content:
                return content
            
        # If UTF-8 didn't work or triangle not found, try other encodings
        encodings = ['utf-8-sig', 'latin-1', 'cp1252']
        for encoding in encodings:
            try:
                with open(template_path, 'r', encoding=encoding) as f:
                    content = f.read()
                    # Replace any corrupted triangle symbols with the correct one
                    content = content.replace('â–³', '▲')
                    content = content.replace('â–²', '▲')
                    print(f"Loaded template with {encoding} encoding")
                    return content
            except UnicodeDecodeError:
                continue
                
        # Last resort - read as binary and decode manually
        with open(template_path, 'rb') as f:
            raw_content = f.read()
            content = raw_content.decode('utf-8', errors='replace')
            content = content.replace('â–³', '▲')
            content = content.replace('â–²', '▲')
            print("Loaded template with fallback decoding")
            return content
            
    except Exception as e:
        print(f"Error loading template file '{template_path}': {e}")
        return None

def generate_dossier_html(agent_data, template, photos_dir="photos"):
    """Generate HTML content for a single agent dossier using external template"""
    
    try:
        # Helper function to safely get data or return placeholder
        def safe_get(field_name, default="[DATA NOT PROVIDED]"):
            return str(agent_data.get(field_name, default))
        
        # Load and process photo
        agent_name = safe_get('Name')
        photo_base64 = load_and_process_photo(agent_name, photos_dir)
        
        # Create photo HTML
        if photo_base64:
            photo_html = f'<img src="{photo_base64}" alt="Agent Photo" style="width: 150px; height: 200px; object-fit: cover;">'
        else:
            photo_html = 'PHOTO<br>[PENDING]'
        
        # Create a mapping of template variables to agent data
        template_vars = {
            '{name}': safe_get('Name'),
            '{looks}': safe_get('Looks'),
            '{anomaly}': safe_get('Anomaly', '[ANOMALY TYPE]'),
            '{reality}': safe_get('Reality', '[REALITY LEVEL]'),
            '{competency}': safe_get('Competency', '[COMPETENCY LEVEL]'),
            '{anomaly_contact}': safe_get('Anomaly_Contact'),
            '{agency_contact}': safe_get('Agency_Contact'),
            '{power_visual}': safe_get('Power_Visual'),
            '{annual_salary}': safe_get('Annual_Salary'),
            '{coffee}': safe_get('Coffee'),
            '{collaboration}': safe_get('Collaboration'),
            '{work_experience}': safe_get('Work_Experience'),
            '{primary_contact}': safe_get('Primary_Contact'),
            '{first_connection}': safe_get('First_Connection'),
            '{second_connection}': safe_get('Second_Connection'),
            '{third_connection}': safe_get('Third_Connection'),
            '{timestamp}': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            '{photo}': photo_html
        }
        
        # Replace template variables using simple string replacement
        html_content = template
        for placeholder, value in template_vars.items():
            html_content = html_content.replace(placeholder, value)
            
        return html_content
        
    except Exception as e:
        print(f"Error generating dossier for {agent_data.get('Name', 'Unknown')}: {e}")
        return None

def create_dossiers(input_file, template_file="dossier_template.html", output_dir="dossiers", photos_dir="photos"):
    """Main function to create dossiers from Excel or CSV data"""
    
    # Load the template
    template = load_template(template_file)
    if template is None:
        print(f"Failed to load template file. Make sure '{template_file}' exists in the current directory.")
        return
    
    # Load the data
    agents_df = load_agent_data(input_file)
    if agents_df is None:
        return
    
    # Create output directory
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Check if photos directory exists
    if not os.path.exists(photos_dir):
        print(f"Photos directory '{photos_dir}' not found. Dossiers will be created without photos.")
    
    # Generate dossier for each agent
    for index, agent in agents_df.iterrows():
        html_content = generate_dossier_html(agent, template, photos_dir)
        
        if html_content is None:
            print(f"Skipping agent {agent.get('Name', 'Unknown')} due to template error.")
            continue
        
        # Create filename (sanitize agent name)
        safe_name = sanitize_filename(agent['Name'])
        filename = f"Agent_{safe_name}_Dossier.html"
        filepath = os.path.join(output_dir, filename)
        
        # Write HTML file
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"Created dossier: {filepath}")
        except Exception as e:
            print(f"Error creating dossier for {agent['Name']}: {e}")
    
    print(f"\nDossier generation complete! Files saved to: {output_dir}/")
    print("You can now open these HTML files in a browser or convert them to PDF.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generate Triangle Agency field agent dossiers')
    parser.add_argument('input_file', help='Path to the ODS (.ods), Excel (.xlsx/.xls), or CSV file containing agent data')
    parser.add_argument('-t', '--template', default='dossier_template.html', help='Path to HTML template file (default: dossier_template.html)')
    parser.add_argument('-o', '--output', default='dossiers', help='Output directory (default: dossiers)')
    parser.add_argument('-p', '--photos', default='photos', help='Photos directory (default: photos)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found!")
    elif not os.path.exists(args.template):
        print(f"Error: Template file '{args.template}' not found!")
    else:
        create_dossiers(args.input_file, args.template, args.output, args.photos)