# pa_v4
 Purchasing Automation


---

# PDF Processing App with PyQt5

## Description

This application is designed to automate the process of handling PDFs and Excel files for purchase orders and invoices. It uses Azure Form Recognizer for data extraction and PyQt5 for the GUI.

## Installation

1. Clone the repository.
2. Run `pip install -r requirements.txt` to install the required packages.

## Folder Structure

```
MyProject/
|-- src/
|-- data/
|-- logs/
|-- archive/
```

## Core Functions

- `split_pdfs(input_folder, output_folder)`: Splits multi-page PDFs into single-page PDFs.
- `send_pdf_to_azure(file_path, client)`: Sends a PDF to Azure for data extraction.
- `extract_invoice_data(file_path, client)`: Extracts invoice data from Azure's response.
- `extract_po_data(file_path, client)`: Extracts purchase order data from Azure's response.
- `create_dataframes(data_list)`: Creates Pandas DataFrames from the extracted data.
- `rename_headers(df)`: Renames DataFrame headers to match Deacom import requirements.
- `save_df_to_excel(df, file_path)`: Saves a DataFrame to an Excel file.
- `user_approval()`: Waits for user approval before proceeding to the next step.
- `compare_prices(po_df, invoice_df)`: Compares prices and quantities between PO and Invoice DataFrames.
- `generate_comparison_file(po_df, invoice_df)`: Generates a new file for comparing prices.
- `update_po(po_df, invoice_df)`: Updates the PO DataFrame based on the Invoice DataFrame.
- `save_to_server(df, server_path)`: Saves the updated PO to a temp folder on the server.
- `archive_files(archive_path, files_to_archive)`: Archives or deletes files.
- `cleanup()`: Optional function to clean up any temporary or intermediate files.
- `log_errors(errors)`: Logs any errors that occur during the process.

## Usage

Run `python src/main_app.py` to start the application.

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct, and the process for submitting pull requests.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

---

