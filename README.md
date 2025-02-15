# Lazy Transcript Generator
A Python application that automates the generation of personalized transcripts based on employee data and accolades. The program provides both a GUI interface and command-line functionality for generating transcripts in CSV or DOCX formats.

## Features
- GUI interface for easy file selection and output configuration
- Support for both CSV and DOCX output formats
- Customizable templates for transcript generation
- Validation of input files and data
- Batch processing of multiple records
- Variable substitution in templates

## Quick Start - Executable Version

### Note: Currently, the executable is only available for Windows.

1. Go to the Releases page
2. Download the latest version: LazyTranscript.exe
3. Double-click the .exe file to run the application

For instructions on how to use the GUI Interface, refer to the Usage section.

For other operating systems, please follow the Python installation instructions below.

## Installation

1. Clone the repository:
```bash
git clone https://github.com/joysonchewjin/LazyTranscript.git
cd LazyTranscript
```

2. Ensure you have Python 3.8 or higher installed:
```bash
python --version
```

3. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Interface

1. Run the application via the terminal
   
```bash
python gui.py
```

2. Select input files:
  - Data CSV File
  - Writeups CSV File
  - DOCX Template (if generating DOCX output)

3. Choose output format (CSV or DOCX)
4. Select output location
5. Click "Generate Transcripts"

### Command Line Interface
The application can also be run from the command line with the following arguments:

```bash
python transcript_generator.py --data DATA_FILE --writeups WRITEUPS_FILE [--output_type {csv,docx}] [--output_dir OUTPUT_DIR] [--template TEMPLATE]
```

Required arguments:
  - --data: Path to CSV file containing personnel data, accolades, and other variables to be included in the transcripts
  - --writeups: Path to CSV file containing writeup templates

Optional arguments:
  - --output_type: Output format, either 'csv' to compile the transcripts into a single CSV file, or 'docx' to compile the transcripts into multiple word documents based on a template
  - --output_dir: Path to directory to save output files
  - --template: Path to DOCX tempalte file (required if output_type is 'docx'

## Input File Requirements

### Data CSV File

Contains each personnel's information with the accolades he or she has achieved, and other variables to be included in the transcript. Must follow the following requirements:
- Must contain a 'name' column for personnel names
- Must have at least one accolade column (prefixed with 'accolade_')
- Accolade columns must contain 'yes' or 'no' values
- Variable and accolade names must not include spaces

Refer to the example_data.csv file for an example data file.

### Writeups CSV File

Contains each accolade's writeup information to be included in the transcript. Must follow the following requirements:
- Must contain an 'accolade' and a 'writeup' column
- Writeups must utilise the placeholder format '${variable_name}' to reference variables in the data file
- Accolade names must be the exact same as in the data file (although prefixing with 'accolade_' is not necessary)
- Writeups cannot be empty
- Accolades cannot be duplicated

Refer to the example_writeups.csv file for an example writeups file.

### Template DOCX File

Contains the template word document required for DOCX output format, for generation of printable transcripts from a single template. Must follow the following requirements:
- Must contain {{transcript}} placeholder for the transcript content
- Must utilise the placeholder format {{variable_name}} to reference variables from the data file

Refer to the example_template.docx file for an example template file.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## Support

For bugs, questions, and discussions, please use the GitHub Issues page.
