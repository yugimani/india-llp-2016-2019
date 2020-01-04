# How To Configure VB Script

## Input and Intermediate Files
I renamed all input files to "yyyy-mm-Month.xlsx" and created blank intermediate files as "yyyy-Month.xlsx". *SrcFile* points to input file location and *DestFile* points to intermediate file location.  

*SrcSheet* - Script expects information under "Domestic LLP" worksheet inside the *SrcFile*. Typically MCA released file uses "Indian LLP's Registered in <Month>" or "LLP Registered", among few other uncommon worksheet names. 

## Input Fields Location 
###### Values are column numbers (A = 1, B = 2, C = 3, etc...) of respective field in *SrcFile*. For example, if LLPIN is in column B, then *SrcNoCol* = 2. 

1. *SrcNoCol* - ID aka LLPIN
1. *SrcNameCol* - Company Name
1. *SrcDateCol* - Date of Incorporation
1. *SrcStateCol* - State or Union Territory
1. *SrcROCCol* - Registrar of Companies
1. *SrcPartnersCol* - Number of Partners
1. *SrcDPartnersCol* - Number of Designated Partners
1. *SrcObligationCol* - Total Obligation of Contribution
1. *SrcDivisionCol* - Industrial Activity Code aka NIC code
1. *SrcActivityCol* - Activity Description

## Date Format
Google Data Studio works well with "yyyy-mm-dd" date format. MCA uses "dd.mm.yyyy", "dd/mm/yyyy", "mm/dd/yyyy", "mm.dd.yyyy" and "yyyy-mm-dd" formats interchangeably. To ensure consistency in date format fed to Data Studio platform, Date of Incorporation is formatted to "yyy-mm-dd". These below variables are used to configure the position of Year, Month and Day in the input field. For example, if the input file has date of incorporation as "mm/dd/yyyy", then *YearPos* = 2, *MonthPos* = 0, *DayPos* = 1.

These values are 0-indexed. 

1. *YearPos* - Position of Year
1. *MonthPos* - Position of Month
1. *DayPos* - Position of Day

## Data Range

If valid data starts from *n*th row, set the cell values of range appropriately in these places. 
```VBScript
Set SrcRange = SrcSheet.Range("A2", SrcSheet.Range("M2").End(xlDown))
```
and
```VBScript
NumRows = SrcSheet.Range("E2", SrcSheet.Range("E2").End(xlDown)).Rows.Count
```

For example, if data starts from 5th row, then these two places need to be hardcoded as
```VBScript
Set SrcRange = SrcSheet.Range("A5", SrcSheet.Range("M5").End(xlDown))
```
and
```VBScript
NumRows = SrcSheet.Range("E5", SrcSheet.Range("E5").End(xlDown)).Rows.Count
```

Arbitray column E is chosen to find the number of data rows in input worksheet. Any other column may also be used given that there are no null/empty cells in that column. 

## Output File

Running this macro would save the contents to intermediate file, *DestFile*. 
Open the *DestFile* in xlsx format and save as csv format. Data Studio platform only allows importing UTF-8 encoded CSV files. 
