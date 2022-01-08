Attribute VB_Name = "Module1"
Sub Mainsub() ' Main sub

Dim WS_Count As Integer
Dim SheetNUM As Integer
Dim answer As Integer

answer = MsgBox("Please note: Tickers data cannot be empty and it shall be ordered in alphabetical order. For Tickers with 0 opening or closing" _
        & "value, the % calculation will return an empty cell. Do you want to Continue?", vbQuestion + vbYesNo)
 
  If answer = vbNo Then Exit Sub
  
  ' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count


' Begin the loop for all the sheet in the worksheet
    For SheetNUM = 1 To WS_Count

       
       'Run GetUniqueItems and StockAnalisys subs for each sheet in the workbook
       ActiveWorkbook.Worksheets(SheetNUM).Select
       
       'Run GetUniqueItems for each sheet in the workbook
       GetUniqueItems
       
       'Run StockAnalisys subs for each sheet in the workbook
       StockAnalisys
    
    Next SheetNUM
    
End Sub


Sub GetUniqueItems() 'Excel VBA to extract the unique tickers code. https://www.thesmallman.com/list-unique-items-with-vba
Dim UItem As Collection
Dim UV As New Collection
Dim rng As Range
Dim I As Long
Dim datalenght As Long

Set UItem = New Collection

On Error Resume Next
For Each rng In Range("A2", Range("A" & Rows.Count).End(xlUp))
UItem.Add CStr(rng), CStr(rng)
datalenght = datalenght + 1
Next
On Error GoTo 0
For I = 1 To UItem.Count
Range("I" & I + 1) = UItem(I)
Next I
'Sort the Range
Range("I2", Range("i" & Rows.Count).End(xlUp)).Sort Range("i2"), 1
End Sub

Sub StockAnalisys()


'Assign the last row with data to a variable
Dim LastRow As Long
'LastRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Dim rng as As Range of the uniques Tickers in column I
Dim rng As Range

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

'Data reporting variables
Dim Greatest_increase As Long
Dim Greatest_decrease As Long
Dim Greatest_totalvolume As Long

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

'loop for all the Tickers in column I
While IsEmpty(Cells(j, 9).Value) = False
 'MsgBox (Cells(j, 9).Value)
    
    'Loop for all the Tickers in column A and to identify the range applicable to specific ticker listed in column I
    If IsEmpty(Cells(j, 9).Value) = False Then
        For k = 2 To LastRow
        
            If Cells(k, 1).Value = Cells(j, 9).Value And Cells(k - 1, 1).Value <> Cells(j, 9).Value Then
                start = k
            ElseIf Cells(k, 1).Value = Cells(j, 9).Value And Cells(k + 1, 1).Value <> Cells(j, 9).Value Then
                finish = k
            End If
        Next
        
        'Calulate the difference between close price and open price and paste the value in column J at the relevant Row
        Cells(j, 10).Value = Cells(finish, 6).Value - Cells(start, 3).Value
        
        'Calulate the % change from opening price at the beginning of a given year to the closing price at the end _
        of that year. If either the start of finish price is 0 it return empty %
        If Cells(finish, 6).Value And Cells(start, 3).Value <> 0 Then
            Cells(j, 11).Value = Cells(finish, 6).Value / Cells(start, 3).Value - 1
        Else
            Cells(j, 11).Value = ""
        End If
        
        'Calculate the total stock volume of the stock
        Cells(j, 12).Value = Application.Sum(Range(Cells(start, 7), Cells(finish, 7)))
        
        'Next Ticker
        j = j + 1
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
Set rng = Range(Cells(2, 11), Cells(j - 1, 11))

'Delete/clear any existing conditional formatting from the range
rng.FormatConditions.Delete

'Defining and setting the criteria for each conditional format (e.g. positive or negative) for the cells defined in rng
Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")

'Defining and setting the format to be applied for each condition green RGB(0, 255, 0) / red RGB(255, 0, 0)
With condition1
    .Interior.Color = RGB(0, 255, 0)
End With

With condition2
    .Interior.Color = RGB(255, 0, 0)
End With


End Sub


