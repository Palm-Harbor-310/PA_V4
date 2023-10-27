# core_functions.py

# Import required libraries
import yaml
import os
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential


# Load the YAML configuration file
with open("config.yaml", 'r') as stream:
    try:
        config = yaml.safe_load(stream)
    except yaml.YAMLError as exc:
        print(exc)

azure_settings = config["azure_settings"]

def on_user_selection(selection):
    if selection == "Invoice":
        service_type = "phh-invoices"
    elif selection == "PO":
        service_type = "pos"
    else:
        raise ValueError("Invalid selection")

    # Initialize Azure client
    client = get_azure_client(service_type)

# Initialize Azure client

def get_azure_client(service_type):
    if service_type not in azure_settings:
        raise ValueError(f"Invalid service_type: {service_type}")
    
    key_name = f"AZURE_API_KEY_{service_type.upper()}"  
    key = os.getenv(key_name)
    if not key:
        raise EnvironmentError(f"Azure API key for {service_type} not set.")

    endpoint = azure_settings[service_type]["endpoint"]
    client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key))
    
    return client


# Function to split multi-page PDFs into single-page PDFs
def split_pdfs(input_folder, output_folder):
    """
    Splits multi-page PDFs into single-page PDFs.
    """
    pass

# Function to send a PDF to Azure for data extraction
def send_pdf_to_azure(file_path, client):
    """
    Sends a PDF to Azure for data extraction.
    """
    pass

# Function to extract invoice data from Azure's response
def extract_invoice_data(file_path, client):
    """
    Extracts invoice data from Azure's response.
    """
    pass

# Function to extract purchase order data from Azure's response
def extract_po_data(file_path, client):
    """
    Extracts purchase order data from Azure's response.
    """
    pass

# Function to create Pandas DataFrames from the extracted data
def create_dataframes(data_list):
    """
    Creates Pandas DataFrames from the extracted data.
    """
    pass

# Function to rename DataFrame headers to match Deacom import requirements
def rename_headers(df):
    """
    Renames DataFrame headers to match Deacom import requirements.
    """
    pass

# Function to save a DataFrame to an Excel file
def save_df_to_excel(df, file_path):
    """
    Saves a DataFrame to an Excel file.
    """
    pass

# Function to wait for user approval before proceeding to the next step
def user_approval():
    """
    Waits for user approval before proceeding to the next step.
    """
    pass

# Function to compare prices and quantities between PO and Invoice DataFrames
def compare_prices(po_df, invoice_df):
    """
    Compares prices and quantities between PO and Invoice DataFrames.
    """
    pass

# Function to generate a new file for comparing prices
def generate_comparison_file(po_df, invoice_df):
    """
    Generates a new file for comparing prices.
    """
    pass

# Function to update the PO DataFrame based on the Invoice DataFrame
def update_po(po_df, invoice_df):
    """
    Updates the PO DataFrame based on the Invoice DataFrame.
    """
    pass

# Function to save the updated PO to a temp folder on the server
def save_to_server(df, server_path):
    """
    Saves the updated PO to a temp folder on the server.
    """
    pass

# Function to archive or delete files based on user settings
def archive_files(archive_path, files_to_archive):
    """
    Archives or deletes files based on user settings.
    """
    pass

# Optional function to clean up any temporary or intermediate files
def cleanup():
    """
    Cleans up any temporary or intermediate files.
    """
    pass

# Function to log any errors that occur during the process
def log_errors(errors):
    """
    Logs any errors that occur during the process.
    """
    pass
