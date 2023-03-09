Sub Notes_Sub()
'WIP Mindmap Notes:
'colour row | if value (i, n)> value(i-1, n) then green else red,where 1=columns, j = rows and n= a fixed value
' ticker counter extract | create a counter of unique values | if (i, 1) <> (i-1, 1), print in (counter +1, 10), else next i
' Opening and closing values | Double variable in an if statement, based on value from ticker counter extract when true

Dim Ticker As String



End Sub

Sub Test_n_Learn()
Dim Nmbr_of_Data_Rows As Long
Dim Nmbr_of_Data_Columns As Long
Dim DistinctV_Count As Integer
Dim x As Integer
Dim y As Integer
Dim RowWindow As Long
Dim UTickerWindow As Long
Dim Distinct_Values As String
DistinctV_Count = 2

'My job is to find you the last populated Cell (final row) in Column A
                        Nmbr_of_Data_Rows = Cells(Rows.Count, 1).End(xlUp).Row

'And my job is to find you the last populated cell (final Column) in Row 1
                        Nmbr_of_Data_Columns = Cells(1, Columns.Count).End(xlToLeft).Column

'I provide the range of rows of data we want to check
                    For RowWindow = 2 To 700
                
'I will be creating the headings for our new table
                    Cells(1, 10).Value = "Ticker Name Temp"
                    Cells(1, 11).Value = "Min Date"
                    Cells(1, 12).Value = "Opening Stock on Min Date"
                    Cells(1, 13).Value = "Max Date"
                    Cells(1, 14).Value = "Close Stock on Max Date"
                     Cells(1, 15).Value = "Ticker Name"
                    Cells(1, 16).Value = "Yearly Change"
                    Cells(1, 17).Value = "Percent Change"
                    Cells(1, 18).Value = "Total Stock Value"
                    Columns("A:I").ColumnWidth = 10
                    Columns("J:W").ColumnWidth = 18

'I am the if statement for checking if the cell value matches the value in the row above.
'If they do not match then we will capture
'The Ticker Name | a min date value based on the rows we have checked | the opening Stock value for this date | a max date based on the data we have checked | the Closing stock value for this date
'We will also be keeping a tally of how many unique values we get.
'That way we can assign each unique value it's own row in the new data table
                    If Cells(RowWindow, 1).Value <> Cells(RowWindow - 1, 1).Value Then
                    Cells(DistinctV_Count, 10).Value = Cells(RowWindow, 1).Value
                   Cells(DistinctV_Count, 11).Value = Cells(RowWindow, 2).Value
                    Cells(DistinctV_Count, 12).Value = Cells(RowWindow, 3).Value
                    Cells(DistinctV_Count, 13).Value = Cells(RowWindow, 2).Value
                    Cells(DistinctV_Count, 14).Value = Cells(RowWindow, 6).Value
                    Cells(DistinctV_Count, 15).Value = Cells(RowWindow, 1).Value
                    Cells(DistinctV_Count, 18).Value = Cells(RowWindow, 7).Value
                    Cells(DistinctV_Count, 16).Value = (Cells(DistinctV_Count, 14).Value - Cells(DistinctV_Count, 12).Value)
                    DistinctV_Count = DistinctV_Count + 1

ElseIf Cells(RowWindow, 1).Value = Cells(RowWindow - 1, 1).Value And Cells(DistinctV_Count - 1, 11).Value > Cells(RowWindow, 2).Value Then
Cells(DistinctV_Count - 1, 11).Value = Cells(RowWindow, 2).Value
Cells(DistinctV_Count - 1, 12).Value = Cells(RowWindow, 3).Value
Cells(DistinctV_Count - 1, 18).Value = (Cells(DistinctV_Count - 1, 18).Value + Cells(RowWindow, 7).Value)
Cells(DistinctV_Count - 1, 16).Value = (Cells(DistinctV_Count - 1, 14).Value - Cells(DistinctV_Count - 1, 12).Value)



ElseIf Cells(RowWindow, 1).Value = Cells(RowWindow - 1, 1).Value And Cells(DistinctV_Count - 1, 13).Value < Cells(RowWindow, 2).Value Then
Cells(DistinctV_Count - 1, 13).Value = Cells(RowWindow, 2).Value
Cells(DistinctV_Count - 1, 14).Value = Cells(RowWindow, 6).Value
Cells(DistinctV_Count - 1, 18).Value = (Cells(DistinctV_Count - 1, 18).Value + Cells(RowWindow, 7).Value)
Cells(DistinctV_Count - 1, 16).Value = (Cells(DistinctV_Count - 1, 14).Value - Cells(DistinctV_Count - 1, 12).Value)

End If
Next RowWindow

MsgBox ((DistinctV_Count - 2) & " Tickers Found")
End Sub
