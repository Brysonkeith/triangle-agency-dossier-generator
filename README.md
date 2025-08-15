# Triangle Agency Dossier Generator

A Python script for creating HTML dossiers for field agents in the Triangle Agency TTRPG.

Create some dossiers to share with players about their fellow field agents based on their Onboarding Questionnaire

I used Google forms to ask my players the questionaire then copied the answers from that into my OpenOffice Calc spreadsheet to load into this script.

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Brysonkeith/triangle-agency-dossier-generator.git
   cd triangle-agency-dossier-generator
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

   Or install individually:
   ```bash
   pip install pandas Pillow openpyxl xlrd odfpy xlsxwriter
   ```

## Quick Start

1. **Prepare your data file** with agent information (see [Data Format](#data-format) below)
2. **Add agent photos** to a `photos/` directory (optional)
3. **Run the generator:**
   ```bash
   python Automatic_Dossier_Creator.py your_agent_data.csv
   ```

Your dossiers will be created in the `dossiers/` directory!

## Usage

### Basic Usage
```bash
python Automatic_Dossier_Creator.py Player_Answers_Example.ods
```

### Advanced Options
```bash
python Automatic_Dossier_Creator.py Player_Answers_Example.ods -t custom_template.html -p photos -o output_folder
```

### Command Line Arguments
- `input_file`: Path to your CSV/Excel/ODS file with agent data
- `-t, --template`: Custom HTML template file (default: `dossier_template.html`)
- `-p, --photos`: Photos directory (default: `photos`)
- `-o, --output`: Output directory for generated dossiers (default: `dossiers`)

## Data Format

Your CSV/Excel file should include these columns:

### Required Fields
- `Name`: Agent's full name
- `Looks`: Physical description
- `Anomaly_Contact`: How they first encountered anomalies
- `Agency_Contact`: How Triangle Agency recruited them
- `Power_Visual`: Description of their anomalous abilities
- `Annual_Salary`: Salary classification
- `Coffee`: Coffee preferences (important for morale!)
- `Collaboration`: Team integration assessment
- `Work_Experience`: Technical skills and background
- `Primary_Contact`: Emergency contact information
- `First_Connection`: Known associate #1
- `Second_Connection`: Known associate #2
- `Third_Connection`: Known associate #3

### Optional Fields (will show placeholders if missing)
- `Anomaly`: Anomaly type classification
- `Reality`: Reality stability level
- `Competency`: Competency classification

## Photo Management

### Photo Requirements
- **Format**: JPG files only
- **Naming**: Use agent name with spaces/special characters replaced by underscores
- **Location**: Place in `photos/` directory (or specify with `-p`)

### Photo Processing
The script automatically:
- Crops photos to 3:4 aspect ratio (centers the image)
- Resizes to exactly 150x200 pixels
- Converts to base64 and embeds in HTML

### Example Photo Names
- "Dojima Tanaka" -> `photos/Dojima_Tanaka.jpg`
- "Agent Smith" -> `photos/Agent_Smith.jpg`
- "Dr. Sarah O'Connor" -> `photos/Dr__Sarah_O_Connor.jpg`

## Template Customization

The default template provides a basic confidential government document look, you can customize it however you like by editing the html template.
**Use template variables** like `{name}`, `{looks}`, `{photo}`, etc. in your template to get embedding to work.
**Run with custom template**: `python script.py data.csv -t my_custom_template.html`

### Available Template Variables
- `{name}`, `{looks}`, `{photo}`
- `{anomaly}`, `{reality}`, `{competency}`
- `{anomaly_contact}`, `{agency_contact}`, `{power_visual}`
- `{annual_salary}`, `{coffee}`, `{collaboration}`, `{work_experience}`
- `{primary_contact}`, `{first_connection}`, `{second_connection}`, `{third_connection}`
- `{timestamp}`

## About Triangle Agency

Check it out at https://www.hauntedtable.games/
I'm absolutely enamored with this game since I found it at GenCon 2025

## Contributing and License

I'm no frontend engineer, if you want to contribute feel free!

This project is open source. Feel free to use, modify, and distribute for your TTRPG campaigns.

## Troubleshooting

**"Missing required field" errors:**
- Check that your CSV has all the required column names
- Ensure column names match exactly (case-sensitive)

**Photos not loading:**
- Check photo filenames match agent names (with underscores for spaces or other special characters)
- Ensure photos are in JPG format
- Verify the photos directory path

**Large file size:**
- Photos are embedded as base64, making files larger
- Consider reducing photo file sizes before processing
- Use JPEG compression for source photos

### Getting Help

If you encounter issues:
1. Check the console output for specific error messages
2. Verify your data format matches the requirements
3. Test with a single agent first
4. Open an issue on GitHub with error details

---

Example Photo by Sindre Fs: https://www.pexels.com/photo/grayscale-photo-of-man-wearing-denim-jacket-1040880/
