from typing import Dict, List, Optional, Set
import pandas as pd
from pathlib import Path
import logging
from docxtpl import DocxTemplate
import os

class DocumentValidator:
    REQUIRED_WRITEUP_COLUMNS = {'accolade', 'writeup'}
    REQUIRED_TEMPLATE_VARIABLES = {'transcript'}

    @staticmethod
    def validate_data_file(file_path: str) -> Optional[Dict[str, str]]:
        """Validate data file for the following:
        1. File exists
        2. File is a CSV
        3. File is not empty
        4. File contains a 'name' column
        5. File has at least one accolade column (prefixed by 'accolade_')
        6. Accolade columns have yes or no values

        Args:
            file_path(str): Path to the data CSV file
        
        Returns:
            errors(str): All errors detected
        """
        errors = {}

        path = Path(file_path)

        if not path.exists():
            errors['file_existence'] = f"Data file does not exist: {file_path}" 
            return errors
        if path.suffix.lower() != '.csv':
            errors['file_extension'] = f"Data file must be a CSV file, got: {path.suffix}"
            return errors

        try:
            df = pd.read_csv(file_path)

            if df.empty:
                errors['data_empty'] = "Data file is empty"
                return errors

            if 'name' not in df.columns:
                errors['missing_name'] = "Data file must contain a 'name' column"

            accolade_columns = [col for col in df.columns if col.startswith('accolade_')]
            
            if not accolade_columns:
                errors['no_accolades'] = "Data file must contain at least one accolade column (prefix: 'accolade_')"

            for col in accolade_columns:
                invalid_values = df[col].dropna().apply(lambda x: str(x).lower() not in ['yes', 'no'])
                if invalid_values.any():
                    invalid_rows = df.index[invalid_values].tolist()
                    errors[f'invalid_{col}'] = f"Invalid values in {col} at rows {invalid_rows}. Must be 'yes' or 'no'"

        except pd.errors.EmptyDataError:
            errors['file_empty'] = "Data file is empty" 
        except pd.errors.ParserError:
            errors['file_format'] = "Invalid CSV file format"
        except Exception as e:
            errors['unexpected'] = f"Unexpected error validating data file: {str(e)}"

        return errors if errors else None
    
    @staticmethod
    def validate_writeups_file(file_path: str) -> Optional[Dict[str, str]]:
        """Validate data file for the following:
        1. File exists
        2. File is a CSV
        3. File is not empty
        4. File contains 'accolade' and 'writeup' columns
        5. 'writeup' values are not empty
        6. No duplicate accolades

        Args:
            file_path(str): Path to the data CSV file
        
        Returns:
            errors(str): All errors detected
        """
        errors = {}

        path = Path(file_path)

        if not path.exists():
            errors['file_existence'] = f"Writeups file does not exist: {file_path}"
            return errors
        
        if path.suffix.lower() != '.csv':
            errors['file_extension'] = f"Writeups file must be a CSV file, got: {path.suffix}"
            return errors

        try:
            df = pd.read_csv(file_path)

            if df.empty:
                errors['data_empty'] = "Writeups file is empty"
                return errors

            # Check for required columns
            missing_columns = DocumentValidator.REQUIRED_WRITEUP_COLUMNS - set(df.columns)
            if missing_columns:
                errors['missing_columns'] = f"Writeups file missing required columns: {missing_columns}"

            # Check for empty writeups
            empty_writeups = df['writeup'].isna() | (df['writeup'].str.strip() == '')
            if empty_writeups.any():
                empty_rows = df.index[empty_writeups].tolist()
                errors['empty_writeups'] = f"Empty writeup templates found in rows: {empty_rows}"

            # Check for empty accolades
            empty_accolades = df['accolade'].isna() | (df['accolade'].str.strip() == '')
            if empty_accolades.any():
                empty_rows = df.index[empty_accolades].tolist()
                errors['empty_accolades'] = f"Empty accolade names found in rows: {empty_rows}"

            # Check for duplicate accolades
            duplicates = df['accolade'].duplicated()
            if duplicates.any():
                duplicate_rows = df.index[duplicates].tolist()
                errors['duplicate_accolades'] = f"Duplicate accolade names found in rows: {duplicate_rows}"

        except pd.errors.EmptyDataError:
            errors['file_empty'] = "Writeups file is empty"
        except pd.errors.ParserError:
            errors['file_format'] = "Invalid CSV file format"
        except Exception as e:
            errors['unexpected'] = f"Unexpected error validating writeups file: {str(e)}"

        return errors if errors else None

    @staticmethod 
    def validate_docx_template(template_path: str) -> Optional[Dict[str, str]]:
        path = Path(template_path)
        errors = {}

        if not path.exists():
            errors['file_existence'] = f"Template file does not exist: {template_path}"
            return errors
        if path.suffix.lower() != '.docx':
            errors['file_extension'] = f"Template file must be a DOCX file, got: {path.suffix}"

        
        try:
            doc = DocxTemplate(template_path)

            template_vars = set()

            for var in doc.get_undeclared_template_variables():
                template_vars.add(var.lower())

            missing_vars = DocumentValidator.REQUIRED_TEMPLATE_VARIABLES - template_vars

            if missing_vars:
                errors['missing_transcript'] = "Template missing transcript insertion"

        except Exception as e:
            errors['template_error'] = f"Error processing template file: {str(e)}"

        return errors if errors else None
    

    @staticmethod
    def validate_output_directory(directory_path: str) -> Optional[Dict[str, str]]:
        errors = {}

        path = Path(directory_path)

        if path.exists():
            if not path.is_dir():
                errors[['not_directory']] = f"Output path exists but is not a directory: {directory_path}"

            if not os.access(path, os.W_OK):
                errors['parent_not_writeable'] = f"Output directory is not writeable: {directory_path}"

        else:
            parent = path.parent
            if not parent.exists():
                errors['parent_missing'] = f"Parent directory does not exist: {parent}"
            
            elif not os.access(parent, os.W_OK):
                errors['parent_not_writeable'] = f"Parent directory is not writeable: {parent}"

        return errors if errors else None