# Export Excel Data to DataGridView in C#

This sample demonstrates how to export data from an Excel worksheet into a Windows Forms DataGridView control using C# and Syncfusion XlsIO. It provides a simple yet powerful example of integrating Excel data with .NET applications, allowing developers to load spreadsheet content directly into a grid for display, editing, or analysis. By leveraging Syncfusion’s Excel engine, the project avoids the need for Microsoft Excel to be installed, making it lightweight and efficient for desktop applications.

## What this sample shows
* Initializes the Syncfusion Excel engine and sets the default version.
* Opens an embedded Excel workbook (Sample.xlsx) from application resources.
* Accesses the first worksheet in the workbook.
* Exports the used range of the worksheet into a DataTable using ExportDataTable.
* Binds the resulting DataTable to a Windows Forms DataGridView.
* Provides a button (Excel to DataGridView) to trigger the import process.
* Displays a label with instructions for the user.
* Ensures no exceptions are thrown if unsaved workbooks exist by setting ThrowNotSavedOnDestroy = false.

## Prerequisites
* Visual Studio
* .NET Framework
* Syncfusion XlsIO NuGet package

## Use cases
* Loading Excel data into a grid for quick visualization.
* Providing end users with editable tabular views of spreadsheet content.
* Automating the import of structured data from Excel into .NET applications.
* Building reporting or dashboard tools that combine Excel input with DataGridView features.

## Repository Structure
* **Excel-to-DataGridView:** Contains the Windows Forms project with a form (Form1) that demonstrates importing Excel data into a grid. The form includes a button, label, and DataGridView control, along with event handlers and helper methods for file access.

## Key Highlights
* **Embedded Resource Access:** Uses Assembly.GetManifestResourceStream to load the Excel file packaged with the application.
* **Data Export:** Converts worksheet ranges into a DataTable for easy binding.
* **UI Integration:** Displays the imported data in a familiar grid interface.
* **Automation:** Eliminates manual copy-paste operations by programmatically importing Excel content.

This project serves as a practical reference for developers who want to integrate Excel data with Windows Forms applications. By following the example, you can adapt the approach to your own business scenarios, ensuring that spreadsheet data is imported seamlessly into DataGridView controls with minimal effort.