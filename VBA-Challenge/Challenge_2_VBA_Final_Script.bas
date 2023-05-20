Attribute VB_Name = "Module2"
Sub Test_n_Rebuild()
Dim Nmbr_of_Data_Rows As Long
Dim Nmbr_of_Data_Columns As Long
Dim DistinctV_Count As Integer
Dim x As Integer
Dim y As Integer
Dim RowWindow As Long
Dim UTickerWindow As Long
Dim Distinct_Values As String
DistinctV_Count = 2
Application.ScreenUpdating = False

For Each WS_Tab In Worksheets
WS_Tab.Select
Application.ScreenUpdating = True

'My job is to find you the last populated Cell (final row) in Column A
                        Nmbr_of_Data_Rows = Cells(Rows.Count, 1).End(xlUp).Row

'And my job is to find you the last populated cell (final Column) in Row 1
                        Nmbr_of_Data_Columns = Cells(1, Columns.Count).End(xlToLeft).Column

'I do the sorting
    Range("A1:G" & Nmbr_of_Data_Rows).Sort Key1:=Range("A1:A" & Nmbr_of_Data_Rows), Order1:=xlAscending, Key1:=Range("B1:B" & Nmbr_of_Data_Rows), Order2:=xlAscending, Header:=xlYes
                    
'I provide the range of rows of data we want to check
                    For RowWindow = 2 To Nmbr_of_Data_Rows
                
'I will be creating the headings for our new table
                    Cells(1, 10).Value = "Ticker Name"
                    Cells(1, 11).Value = "Min Date"
                    Cells(1, 12).Value = "Opening Stock on Min Date"
                    Cells(1, 13).Value = "Max Date"
                    Cells(1, 14).Value = "Close Stock on Max Date"
                    Cells(1, 15).Value = "Total Stock Value"
                    Cells(1, 16).Value = "Yearly Change"
                    Cells(1, 17).Value = "Percent Change"
                    
                    Cells(1, 21).Value = "Ticker Name"
                    Cells(1, 22).Value = "Value"
                    Cells(2, 20).Value = "Greatest % Increase"
                    Cells(3, 20).Value = "Greatest % Decrease"
                    Cells(4, 20).Value = "Greatest Total Volume"
                    Cells(2, 21).Value = "Load Ticker"
                    Cells(3, 21).Value = "Load Ticker"
                    Cells(4, 21).Value = "Load Ticker"
                    Cells(2, 22).Value = -1
                    Cells(3, 22).Value = 1
                    Cells(4, 22).Value = 0.2
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
                    Cells(DistinctV_Count, 15).Value = Cells(RowWindow, 7).Value
                    DistinctV_Count = DistinctV_Count + 1

ElseIf Cells(RowWindow, 1).Value = Cells(RowWindow - 1, 1).Value And Cells(RowWindow, 1).Value <> Cells(RowWindow + 1, 1).Value Then
Cells(DistinctV_Count - 1, 13).Value = Cells(RowWindow, 2).Value
Cells(DistinctV_Count - 1, 14).Value = Cells(RowWindow, 6).Value
Cells(DistinctV_Count - 1, 15).Value = (Cells(DistinctV_Count - 1, 15).Value + Cells(RowWindow, 7).Value)
Cells(DistinctV_Count - 1, 16).Value = (Cells(DistinctV_Count - 1, 14).Value - Cells(DistinctV_Count - 1, 12).Value)
Cells(DistinctV_Count - 1, 17).Value = FormatPercent(((Cells(DistinctV_Count - 1, 14).Value / Cells(DistinctV_Count - 1, 12).Value)) - 1)

ElseIf Cells(RowWindow, 1).Value = Cells(RowWindow - 1, 1).Value Then
Cells(DistinctV_Count - 1, 15).Value = (Cells(DistinctV_Count - 1, 15).Value + Cells(RowWindow, 7).Value)

End If

'Yearly Change and % Change Calculation of Harvested Data
Cells(DistinctV_Count - 1, 16).Value = (Cells(DistinctV_Count - 1, 14).Value - Cells(DistinctV_Count - 1, 12).Value)
Cells(DistinctV_Count - 1, 17).Value = FormatPercent(((Cells(DistinctV_Count - 1, 14).Value / Cells(DistinctV_Count - 1, 12).Value)) - 1)

'Colur Yearly Change
If Cells(DistinctV_Count - 1, 16).Value > 0 Then
Cells(DistinctV_Count - 1, 16).Interior.ColorIndex = 4
Else
Cells(DistinctV_Count - 1, 16).Interior.ColorIndex = 3
End If

'Colour %Change
If Cells(DistinctV_Count - 1, 17).Value > 0 Then
Cells(DistinctV_Count - 1, 17).Interior.ColorIndex = 4
Else
Cells(DistinctV_Count - 1, 17).Interior.ColorIndex = 3
End If
Next RowWindow

Distinct_Responses = Range("P" & Rows.Count).End(xlUp).Row

For Bonus_Checker = 2 To Distinct_Responses

    If Range("Q" & Bonus_Checker).Value > Range("V2").Value Then
    Range("V2").Value = FormatPercent(Range("Q" & Bonus_Checker).Value)
    Range("U2").Value = Range("J" & Bonus_Checker).Value
End If

If Range("Q" & Bonus_Checker).Value < Range("V3").Value Then
Range("V3").Value = FormatPercent(Range("Q" & Bonus_Checker).Value)
Range("U3").Value = Range("J" & Bonus_Checker).Value
End If

If Range("O" & Bonus_Checker).Value > Range("V4").Value Then
Range("V4").Value = Range("O" & Bonus_Checker).Value
Range("U4").Value = Range("J" & Bonus_Checker).Value
End If


If Distinctive_Response > 1 Then
Destinctive_Response = 0
End If
Next Bonus_Checker

If DistinctV_Count > 2 Then
DistinctV_Count = 2

Else: End If
Next
End Sub

