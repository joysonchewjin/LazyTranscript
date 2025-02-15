import pandas as pd
from string import Template
from docxtpl import DocxTemplate
from pathlib import Path
from typing import Set, Dict, Optional
import os
from datetime import datetime
import logging
from document_validator import DocumentValidator

class ValidationError(Exception):
    """Custom exception for validatiohn errors that contain all the validation messages"""
    def __init__(self, validation_errors: dict):
        self.validation_errors = validation_errors
        message = "\n".join(f"{k}: {v}" for k, v in validation_errors.items())
        super().__init__(message)

class TranscriptGenerator:
    """A class to generate transcripts from accolades and templates.
    
    Attributes:
        ACCOLADE_PREFIX (str): Prefix used to identify accolade columns
        data_df (pd.DataFrame): DataFrame containing person data and accolades
        writeups_df (pd.DataFrame): DataFrame containing writeup templates
        writeups (Dict[str, Template]): Dictionary of prepared writeup templates
        accolade_columns (Set[str]): Set of identified accolade column names
        variable_columns (Set[str]): Set of identified variable column names
    """
    
    ACCOLADE_PREFIX = "accolade_"

    def __init__(self, data_path: str, writeups_path: str):
        """Initialize the TranscriptGenerator.
        
        Args:
            data_path (str): Path to CSV file containing person data and accolades
            writeups_path (str): Path to CSV file containing writeup templates
            
        Raises:
            FileNotFoundError: If either input file doesn't exist
            pd.errors.EmptyDataError: If either input file is empty
        """

        data_errors = DocumentValidator.validate_data_file(data_path)
        if data_errors:
            raise ValidationError(data_errors)
        
        writeups_errors = DocumentValidator.validate_writeups_file(writeups_path)
        if writeups_errors:
            raise ValidationError(writeups_errors)

        self.data_df = pd.read_csv(data_path)
        self.writeups_df = pd.read_csv(writeups_path)
            
        self.writeups = self._prepare_writeups()
        self.accolade_columns = self._identify_accolade_columns()
        self.variable_columns = self._identify_variable_columns()

    def _identify_accolade_columns(self) -> Set[str]:
        """Identify columns that represent accolades.
        
        Returns:
            Set[str]: Set of column names that start with ACCOLADE_PREFIX
        """
        return {
            col
            for col in self.data_df.columns
            if col.startswith(self.ACCOLADE_PREFIX)
        }
    
    def _identify_variable_columns(self) -> Set[str]:
        """Identify columns that represent variables.
        
        Returns:
            Set[str]: Set of column names that don't start with ACCOLADE_PREFIX
        """
        return {
            col for col in self.data_df.columns
            if not col.startswith(self.ACCOLADE_PREFIX)
        }

    def _get_person_variables(self, person: pd.Series) -> Dict[str, str]:
        """Extract variables from a person's data.
        
        Args:
            person (pd.Series): Row from data_df containing person's information
            
        Returns:
            Dict[str, str]: Dictionary of variable names and values
        """
        return {
            col: str(person[col]) if pd.notna(person[col]) else "None"
            for col in self.variable_columns
        }

    def _prepare_writeups(self) -> Dict[str, Template]:
        """Prepare writeup templates from the templates file.
        
        Returns:
            Dict[str, Template]: Dictionary of event names to their templates
            
        Raises:
            KeyError: If required columns are missing from writeups_df
        """
        if not all(col in self.writeups_df.columns for col in ['accolade', 'writeup']):
            raise KeyError("Templates file must contain 'accolade' and 'writeup' columns")
            
        templates = {}
        for _, row in self.writeups_df.iterrows():
            accolade_name = self.ACCOLADE_PREFIX + row['accolade']
            template_text = row['writeup']
            templates[accolade_name] = Template(template_text)
        return templates

    def generate_transcript(self, person: pd.Series) -> str:
        """Generate a transcript for a person based on their accolades.
        
        Args:
            person (pd.Series): Row from data_df containing person's information
            
        Returns:
            str: Generated transcript text
            
        Raises:
            KeyError: If template variables are missing
        """
        transcript_parts = []
        template_vars = self._get_person_variables(person)

        for a in self.accolade_columns:
            if pd.notna(person[a]) and person[a].lower() == 'yes' and a in self.writeups:
                try:
                    section = self.writeups[a].substitute(template_vars)
                    transcript_parts.append(section)
                except KeyError as e:
                    logging.warning(f"Template for {a} is missing required variable: {e}")
                    raise
    
        return ' '.join(transcript_parts)
    
    def export_transcripts_csv(self, export_dir: Optional[str] = None) -> str:
        """Export all transcripts to a CSV file.

        Args:
            export_dir(Optional[str]): Path to the export directory. If None, directory will be created in the current working directory.
        
        Returns:
            str: Path to the exported CSV file
            
        Raises:
            OSError: If there's an error writing the file
        """
        if export_dir:
            dir_errors = DocumentValidator.validate_output_directory(export_dir)
            if dir_errors:
                raise ValidationError(dir_errors)

        results = []
        for _, person in self.data_df.iterrows():
            try:
                transcript = self.generate_transcript(person)
                results.append({'transcript': transcript})
            except KeyError as e:
                logging.error(f"Error generating transcript: {e}")
                results.append({'transcript': f"Error: {str(e)}"})

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'transcripts_{timestamp}.csv'

        if export_dir is None:
            output_dir = Path.cwd()
        else:
            output_dir = Path(export_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        
        output_path = output_dir / filename

        try:
            results_df = pd.DataFrame(results)
            results_df.to_csv(output_path, index=False)
        except OSError as e:
            logging.error(f"Error saving CSV file: {e}")
            raise

        return str(output_path)
    
    def _replace_template_variables(self, text: str, variables: Dict[str, str]) -> str:
        """Replace template variables in text with their values.
        
        Args:
            text (str): Text containing template variables in [VARIABLE] format
            variables (Dict[str, str]): Dictionary of variable names and their values
            
        Returns:
            str: Text with variables replaced
        """
        result = text
        for var_name, value in variables.items():
            placeholder = f'[{var_name.upper()}]'
            result = result.replace(placeholder, value)
        return result
    

    def export_from_docx_template(self, docx_template_path: str, export_dir: Optional[str] = None) -> str:
        """Export transcripts to individual DOCX files using a template.
        
        The template can include both {{transcript}} for the full transcript and
        any person variable in square brackets (e.g., {{name}}, {{rank}}, etc.)
        
        Args:
            docx_template_path (str): Path to the DOCX template file
            export_dir(Optional[str]): Path to the export directory where DOCX files are to be saved. If None, directory will be created in the current working directory
            
        Returns:
            str: Path to the directory containing exported files
            
        Raises:
            FileNotFoundError: If template file doesn't exist
            OSError: If there's an error creating output directory
        """
        template_errors = DocumentValidator.validate_docx_template(docx_template_path)
        if template_errors:
            raise ValidationError(template_errors)
        
        if export_dir:
            dir_errors = DocumentValidator.validate_output_directory(export_dir)
            if dir_errors:
                raise ValidationError(dir_errors)
        if not os.path.exists(docx_template_path):
            raise FileNotFoundError(f"Template file not found: {docx_template_path}")


        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if export_dir is None:
            output_dir = Path.cwd() / f'transcripts_{timestamp}'
        else: 
            output_dir = Path(export_dir)
        
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            logging.error(f"Error creating output directory: {e}")
            raise


        for _, person in self.data_df.iterrows():
            name = self._get_person_variables(person).get('name', 'unknown')
            try:
                # Create a new instance of DocxTemplate for each person
                doc = DocxTemplate(docx_template_path)

                #Generate transcript
                transcript = self.generate_transcript(person)

                # Prep context dictionary
                context = self._get_person_variables(person)

                context['transcript'] = transcript

                # Normalize context keys for template matching
                normalized_context = {}
                for k, v in context.items():
                    # Convert to lowercase and replace spaces with underscores
                    key = k.lower()
                    key = key.replace(' ', '_')
                    normalized_context[key] = v
                
                context = normalized_context

                # Render the content
                doc.render(context)
                
                # Generate filename using person's name if available
                safe_name = ''.join(c for c in name if c.isalnum() or c in ('-', '_')).strip()
                output_filename = f'transcript_{safe_name}_{timestamp}.docx'
                output_path = output_dir / output_filename

                doc.save(output_path)
                
            except Exception as e:
                logging.error(f'Error processing document for {name}: {e}')
                continue

        return str(output_dir)

def main():
    """Main function to run the transcript generator."""
    logging.basicConfig(level=logging.INFO)
    
    try:
        generator = TranscriptGenerator(
            data_path='../examples/data/example_data.csv', 
            writeups_path='../examples/writeups/example_writeups.csv',
        )


        csv_path = generator.export_transcripts_csv()
        logging.info(f"CSV exported to: {csv_path}")

        docx_dir = generator.export_from_docx_template('../examples/template/example_template.docx')
        logging.info(f"DOCX files exported to: {docx_dir}")
        
    except Exception as e:
        logging.error(f"Error running transcript generator: {e}")
        raise

if __name__ == '__main__':
    main()