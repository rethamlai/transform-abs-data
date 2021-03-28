# Transform ABS TableBuilder data
VBA code to consolidate data downloaded from the Australian Bureau of Statistics (ABS) TableBuilder into a more conveninent format for pivot table use.

## What is this for?
Data downloaded from the ABS with row fields, column fields and wafers populated is arranged into an excel table format with multiple tabs for each wafer. This can sometimes be hard to deal with especially if you want to use pivot table functions. The data being split into multiple tabs for each wafer can also be time consuming to deal with if the user does not want the data broken down the way it has been.

## Example
1. Download and open the file "abs_transform_data.xlsm"
2. Press the button to see how the data will be transformed

## How to use this for your own ABS data
1. Open your ABS workbook's developers tab and then press "View Code"
  - If the Developers tab is not showing in your Excel Ribbon, please see this link on how to open the tab: https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45

2. Open any module and then paste the below into the module:

``` Ruby
Function check_if_sheet_exists(sh As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sh)
    On Error GoTo 0

    If Not ws Is Nothing Then check_if_sheet_exists = True
End Function


Sub abs_table_builder_data_flat_to_long()

    Dim RowFields, colFields, sheetCount, Index As Long
    Dim xCount As Integer
    
    Dim sh As String: sh = "Transformed data"
    
    ' Check if worksheet exists, if not then create
    If check_if_sheet_exists(sh) Then
    Else
        ThisWorkbook.Sheets.Add.Name = sh
    End If
    
    ' Count number of row fields
    RowFields = Worksheets("Data Sheet 0").Range("B:B").Cells.SpecialCells(xlCellTypeConstants).Count - 4
    
    ' Count number of col fields
    colFields = Worksheets("Data Sheet 0").Rows(10).Cells.SpecialCells(xlCellTypeConstants).Count - 3
    
    ' Count number of worksheets with the string "Data Sheet"
    For k = 1 To ThisWorkbook.Sheets.Count
        If Mid(Sheets(k).Name, 1, 10) = "Data Sheet" Then xCount = xCount + 1
    Next
    
    sheetCount = CStr(xCount)
    
    ' Number iterator
    Index = 0
    
    ' Populate headings
    Worksheets(sh).Range("A1").Value = Worksheets("Data Sheet 0").Range("B11").Value
    Worksheets(sh).Range("B1").Value = Worksheets("Data Sheet 0").Range("A10").Value
    Worksheets(sh).Range("C1").Value = "Wafer field"
    Worksheets(sh).Range("D1").Value = "Value"
    
    For s = 0 To sheetCount - 1
    
        For I = 1 To colFields
        
            For j = 2 To RowFields + 1
                
                ' Populate row fields
                Worksheets(sh).Range("A" & j + Index).Value = Worksheets("Data Sheet " & s).Range("B" & 10 + j).Value
                
                ' Populate col fields
                Worksheets(sh).Range("B" & j + Index).Value = Worksheets("Data Sheet " & s).Cells(10, I + 2).Value
                
                ' Populate worksheet wafer
                Worksheets(sh).Range("C" & j + Index).Value = Worksheets("Data Sheet " & s).Range("A9").Value
                
                ' Populate value
                Worksheets(sh).Range("D" & j + Index).Value = Worksheets("Data Sheet " & s).Cells(10 + j, 2 + I).Value
                
            Next j
            
            Index = Index + RowFields
            
        Next I

    Next s
    
End Sub
```
3. Run the sub function "abs_table_builder_data_flat_to_long"

