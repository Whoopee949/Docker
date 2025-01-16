import os
import pandas as pd
from typing import Optional
import openai
from pandas import DataFrame
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

class EmployeeDataProcessor:
    def __init__(self, file_path: str, api_key: str, sheet_name: str = 'Table 1'):
        if not file_path.endswith('.xlsx'):
            raise ValueError("File must be an Excel (.xlsx) file")
        self.file_path = file_path
        self.api_key = api_key
        self.sheet_name = sheet_name
        self.data: Optional[DataFrame] = None
        openai.api_key = api_key

    def read_excel(self) -> None:
        try:
            if not os.path.exists(self.file_path):
                raise FileNotFoundError(f"Excel file not found: {self.file_path}")
            
            # Get available sheet names
            xl = pd.ExcelFile(self.file_path)
            logger.debug(f"Available sheets: {xl.sheet_names}")
            
            if self.sheet_name not in xl.sheet_names:
                raise ValueError(f"Sheet '{self.sheet_name}' not found. Available sheets: {xl.sheet_names}")
            
            # Read Excel with explicit parameters
            self.data = pd.read_excel(
                self.file_path,
                sheet_name=self.sheet_name,
                engine='openpyxl',
                na_filter=True,
                dtype=object  # Read all columns as objects initially
            )
            
            logger.debug(f"Data shape: {self.data.shape}")
            logger.debug(f"Data types: {self.data.dtypes}")
            
            # Convert data types appropriately
            for column in self.data.columns:
                try:
                    # Try numeric conversion
                    pd.to_numeric(self.data[column], errors='raise')
                    self.data[column] = pd.to_numeric(self.data[column], errors='coerce')
                except:
                    # Keep as object if numeric conversion fails
                    pass
            
            print(f"Successfully loaded {len(self.data)} records from Excel")
            print(f"Columns: {', '.join(self.data.columns)}")
            
        except pd.errors.EmptyDataError:
            logger.error("The Excel file is empty")
            self.data = None
        except Exception as e:
            logger.error(f"Error reading the Excel file: {str(e)}")
            self.data = None
            raise

if __name__ == "__main__":
    file_path = "/home/hadi/Desktop/RAG/Employee_Data.xlsx"
    api_key = "your-api-key-here"
    processor = EmployeeDataProcessor(file_path, api_key)
    processor.read_excel()