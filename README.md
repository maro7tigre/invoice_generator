# Invoice Generator

A Python application for generating professional invoices from Excel data. This application uses a template-based approach to create customized invoices with proper formatting and automatic calculations.

## Features

- **Template-Based Design**: Use your own Excel template for a consistent, professional look
- **Batch Processing**: Automatically generates multiple invoices for large datasets
- **Client Information**: Include client details (name, address, ICE number)
- **Automatic Calculations**: Handles item totals, subtotals, taxes, and final amounts
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **User-Friendly Interface**: Simple GUI for easy invoice generation

## Installation

### Prerequisites

- Python 3.6+
- Required Python packages:
  - pandas
  - openpyxl
  - tkinter (usually included with Python)

### Setup

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/invoice-generator.git
   cd invoice-generator
   ```

2. Install the required packages:
   ```bash
   pip install pandas openpyxl
   ```

3. Place your invoice template Excel file (named "FACTURE COMPT.xlsx") in a folder called "facture" on your desktop
   - The template should have a table structure with 23 rows for invoice items
   - Table should start at row 12
   - Description in column A (possibly merged with B-D)
   - Quantity in column H
   - Unit Price in column I
   - Total cells should be after row 35 with labels like "Total HT", "TVA", etc.

## Usage

### Running the Application

```bash
python invoice_generator.py
```

### Using the GUI

1. **Select Data File**: Choose the Excel file containing your invoice data
   - Data file should have columns for description, quantity, and unit price
   - Each row represents one invoice item

2. **Invoice Number**: Optionally specify an invoice number or leave blank for automatic generation

3. **Client Information**: Enter client details:
   - Name/Company
   - Address
   - ICE Number

4. **Output Options**: Choose where to save the generated invoice(s)

5. **Generate**: Click "Générer la facture" to create the invoice(s)

### Data File Format

The data file should be an Excel file with at least 3 columns:
- Column 1: Description of the item
- Column 2: Quantity
- Column 3: Unit price

Example:
```
Computer Repair Service | 1 | 500
RAM Upgrade 8GB         | 2 | 300
SSD Installation 500GB  | 1 | 600
```

## Multiple Invoices

If your data file contains more than 23 items:
- The application will automatically generate multiple invoice files
- Each invoice will contain up to 23 items
- Invoices will be numbered sequentially (e.g., FA 001/2025_1, FA 001/2025_2)
- Continuation invoices include references to the original invoice

## Customization

### Template Structure

The application assumes the following Excel template structure:
- Invoice number in cell E3
- Date in cell I3
- Client information in cells H5 (name), H7 (address), H9 (ICE)
- Item table starting at row 12 with 23 rows
- Column A for description (may be merged A-D)
- Column H for quantity
- Column I for unit price
- Column J for item total (calculated)
- Total section below the item table with labels for Total HT, TVA, and Total TTC

### Modifying the Code

If your template has a different structure, modify these variables in the `create_invoice` method:
- `start_row`: The first row of the items table (default: 12)
- `max_rows_per_invoice`: Number of rows in your template (default: 23)

## Troubleshooting

### Common Issues

- **Template Not Found**: Ensure "FACTURE COMPT.xlsx" exists in the "facture" folder on your desktop
- **Data File Format**: Check that your data file has at least 3 columns with proper values
- **Merged Cell Errors**: The application handles merged cells, but complex templates may cause issues

### Getting Help

If you encounter problems:
1. Check the log area in the application for error messages
2. Verify your template structure matches the expected format
3. Ensure your data file contains valid information

## License

[MIT License](LICENSE)

## Acknowledgments

- This application uses openpyxl for Excel manipulation
- Built with tkinter for the graphical user interface
