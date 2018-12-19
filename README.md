# powershell-pdftables-api
Example of using the https://pdftables.com/ API from Windows PowerShell. 

The scripts require PowerShell 3.0 or later. This is installed by default from Windows 8. In Windows 7 you can install it from:

https://docs.microsoft.com/en-us/powershell/scripting/setup/installing-windows-powershell?view=powershell-5.1

# powershell-convert-dir.ps1 
This enables you to batch convert PDFs from the command line in Windows.

The script will convert all PDFs in a directory (folder) using PDFTables. It will also process PDFs in sub directories and will mirror the input sub-directory structure in the output sub-directory.

This script takes three mandatory parameters:

* Your PDFTables API Key (displayed on https://pdftables.com/pdf-to-excel-api after you login)
* The input directory containing PDF files
* The output directory containing PDF files (this can be the same as the input directory)

It also has one optional parameter: the format in which you want the output files. This is specified with `-format VALUE` where `VALUE` can be `xlsx-single` (the default), `xlsx-multiple`, `csv` or `xml`.

Example usage from Windows command line:
```
powershell -ExecutionPolicy Unrestricted -File pdftables-convert-dir.ps1 YOUR-API-KEY INPUT_DIR OUTPUT_DIR
```
Example usage output in CSV format:
```
powershell -ExecutionPolicy Unrestricted -File pdftables-convert-dir.ps1 YOUR-API-KEY INPUT_DIR OUTPUT_DIR -format csv
```

# powershell-convert-single.ps1 
This enables you to convert a PDF from one directory into another (or same) directory from the command line in Windows.

The script will convert a single PDF in one directory to another directory.

This script takes three mandatory parameters:

* Your PDFTables API Key (displayed on https://pdftables.com/pdf-to-excel-api after you login)
* The input directory including your PDF file name
* The output directory where you'd like to save your converted PDF (this can be the same as the input directory)

It also has one optional parameter: the format in which you want the output files. This is specified with `-format VALUE` where `VALUE` can be `xlsx-single` (the default), `xlsx-multiple`, `csv` or `xml`.

Example usage from Windows command line:
```
powershell -ExecutionPolicy Unrestricted -File pdftables-convert-single.ps1 YOUR-API-KEY INPUT_DIR_FILE OUTPUT_DIR
```
Example usage output in CSV format:
```
powershell -ExecutionPolicy Unrestricted -File pdftables-convert-single.ps1 YOUR-API-KEY INPUT_DIR_FILE OUTPUT_DIR -format csv
```
