Attribute VB_Name = "Module1"
Sub MainStocksAnalysis() ' Main sub Loop through Worksheet

Dim WS_Count As Integer
Dim SheetNUM As Integer
Dim answer As Integer

answer = MsgBox("Please note: Tickers data cannot be empty and it shall be ordered in alphabetical order. For Tickers with 0 opening or closing" _
        & "values  the % calculation will return an empty cell. Do you want to Continue?", vbQuestion + vbYesNo)
 
  If answer = vbNo Then Exit Sub
  
  ' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count


' Begin the loop for all the sheet in the worksheet
    For SheetNUM = 1 To WS_Count

       
       'Run GetUniqueItems and StockAnalisys subs for each sheet in the workbook
        ActiveWorkbook.Worksheets(SheetNUM).Select
       
       'Run AdvFilter for each sheet in the workbook
        AdvFilter
       
       'Run StockAnalisys subs for each sheet in the workbook
        StockAnalisys
    
    Next SheetNUM
    
End Sub


Sub GetUniqueItems() 'Excel VBA to extract the unique tickers code. Code from https://www.thesmallman.com/list-unique-items-with-vba
Dim UItem As Collection
Dim UV As New Collection
Dim Rng As Range
Dim I As Long
Dim datalenght As Long

Set UItem = New Collection

On Error Resume Next
For Each Rng In Range("A2", Range("A" & Rows.Count).End(xlUp))
UItem.Add CStr(Rng), CStr(Rng)
datalenght = datalenght + 1
Next
On Error GoTo 0
For I = 1 To UItem.Count
Range("I" & I + 1) = UItem(I)
Next I

'Sort the Range
Range("I2", Range("i" & Rows.Count).End(xlUp)).Sort Range("i2"), 1
End Sub

Sub AdvFilter() ' ' AdvFilter Macro https://www.get-digital-help.com/create-a-unique-distinct-list-of-a-long-list-without-sacrificing-performance-using-vba-in-excel/

' Select first cell in column (Sheet1!A2)
ActiveSheet.Range("A2").Select
ActiveSheet.Range("A2", Range("A" & Rows.Count).End(xlUp)).Select

' Extending the selection down to the cell just above the first blank cell in this column
Range(Selection, Selection.End(xlDown)).Select

' Run Advanced Filter on selection and copy to Sheet1!C2
Selection.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("i2", Range("i" & Rows.Count).End(xlUp)), Unique:=True
Range("i1") = Empty

End Sub



Sub StockAnalisys()


'Assign the last row with data to a variable
Dim LastRow As Long
'LastRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Dim rng as As Range of the uniques Tickers in column I
Dim Rng As Range

'Conditional Formatting
Dim condition1 As FormatCondition, condition2 As FormatCondition

'Counter definition
Dim I, j, k As Long
I = 2
j = 2
k = 2

'Reference of the Tikers row relative position in the list
Dim Target As String
Dim start As Long
Dim finish As Long


Dim Greatest_increase As Long
Dim Greatest_decrease As Long
Dim Greatest_totalvolume As Long

Range("A:A").NumberFormat = "@"
Range("I:I").NumberFormat = "@"

' Header set up as "The ticker symbol" / "Yearly change"/"The percent change"/"The total stock volume"
Cells(1, 9).Value = "The ticker symbol"
Cells(1, 10).Value = "Yearly change"
Cells(1, 11).Value = "The percent change"
Cells(1, 12).Value = "The total stock volume"

' Reporting matrix set up as "Greatest % increase" / "Greatest % decrease"/"Greatest total volume" by "Tickers" and "Value"
Cells(1, 16).Value = "Tickers"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"

'loop for all the Tickers in column I until find empty cell
While IsEmpty(Cells(j, 9).Value) = False
 
    
    'Loop for all the Tickers in column A and to identify the range applicable to specific ticker listed in column I
    'check if the content is empty
    If IsEmpty(Cells(j, 9).Value) = False Then
            'select the tickers column
            Range("A:A").Select
            
            'Find the start of the range using the row number where found the ticker using find in a range
            start = Selection.Find(Cells(j, 9).Value, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=True).Row
            
            'Find the end of the range using the row number where found the next ticker using find in a range
            finish = Selection.Find(Cells(j + 1, 9).Value, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=True, SearchFormat:=True).Row - 1
    
            'Calculate the difference between close price and open price and paste the value in column J at the relevant Row
            Cells(j, 10).Value = Cells(finish, 6).Value - Cells(start, 3).Value
            
            'Calculate the % change from opening price at the beginning of a given year to the closing price at the end _
            of that year. If either the start (column 3) or finish price (column 6)is 0 it return empty %
            
            If Cells(finish, 6).Value And Cells(start, 3).Value <> 0 Then
                Cells(j, 11).Value = Cells(finish, 6).Value / Cells(start, 3).Value - 1
            Else
                Cells(j, 11).Value = ""
            End If
            
            'Calculate the total stock volume of the stock
            Cells(j, 12).Value = Application.Sum(Range(Cells(start, 7), Cells(finish, 7)))
            
            'Next Ticker
            j = j + 1
'    End If
    Else
    End If
Wend
 
 'Calculation of the Max "Greatest % increase" / "Greatest % decrease"/"Greatest total volume"
Cells(2, 17).Value = Application.Max(Range(Cells(2, 11), Cells(j - 1, 11)))
Cells(3, 17).Value = Application.Min(Range(Cells(2, 11), Cells(j - 1, 11)))
Cells(4, 17).Value = Application.Max(Range(Cells(2, 12), Cells(j - 1, 12)))

'Search for the tickers associated with "Greatest % increase" / "Greatest % decrease"/"Greatest total volume" _
by creating a formula in Excel
Cells(2, 16).Formula = "=INDIRECT(""I""&MATCH(MAX(K:K),K:K,0))"
Cells(3, 16).Formula = "=INDIRECT(""I""&MATCH(MIN(K:K),K:K,0))"
Cells(4, 16).Formula = "=INDIRECT(""I""&MATCH(MAX(l:l),l:l,0))"

'Format the columns as %
Range(Cells(2, 11), Cells(j - 1, 11)).NumberFormat = "0.00%"
Range(Cells(2, 17), Cells(3, 17)).NumberFormat = "0.00%"

'Defined range based on the number of uniques tickers
Set Rng = Range(Cells(2, 10), Cells(j - 1, 10))

'Delete/clear any existing conditional formatting from the range
Rng.FormatConditions.Delete

'Defining and setting the criteria for each conditional format (e.g. positive or negative) for the cells defined in rng
Set condition1 = Rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set condition2 = Rng.FormatConditions.Add(xlCellValue, xlLess, "=0")

'Defining and setting the format to be applied for each condition - green RGB(0, 255, 0) - red RGB(255, 0, 0)
With condition1
    .Interior.Color = RGB(0, 255, 0)
End With

With condition2
    .Interior.Color = RGB(255, 0, 0)
End With


End Sub



