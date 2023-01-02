Attribute VB_Name = "BenfordsLawMacro"
Option Base 1

Function NewBlankWorksheet()
'
' NewBlankWorksheet Macro
'
'
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim count As Integer

    For Each ws In Worksheets
        If Left(ws.Name, 16) = "Benford's Report" Then
            count = count + 1
        End If
    Next ws
    
    If count > 0 Then
        Sheets.Add(Before:=Sheets(1)).Name = "Benford's Report " & count
    Else
        Sheets.Add(Before:=Sheets(1)).Name = "Benford's Report"
    End If
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
    
End Function

Function HSLToRGB(h As Single, s As Single, l As Single) As Long
    ' Convert HSL values to RGB values
    Dim r As Single, g As Single, b As Single
    
    If s = 0 Then
        ' If saturation is 0, all colors are the same.
        ' The RGB values are equal to the lightness value.
        r = l
        g = l
        b = l
    Else
        ' Calculate temporary values based on lightness and saturation
        Dim temp1 As Single, temp2 As Single
        If l < 0.5 Then
            temp1 = l * (1 + s)
        Else
            temp1 = l + s - (l * s)
        End If
        temp2 = 2 * l - temp1
        
        ' Calculate the red, green, and blue values
        r = HueToRGB(temp1, temp2, h + (1 / 3))
        g = HueToRGB(temp1, temp2, h)
        b = HueToRGB(temp1, temp2, h - (1 / 3))
    End If
    
    ' Return the RGB values as a single long value
    HSLToRGB = RGB(r * 255, g * 255, b * 255)
End Function

Function HueToRGB(temp1 As Single, temp2 As Single, tempH As Single) As Single
    ' Calculate the red, green, or blue value based on the hue
    If tempH < 0 Then
        tempH = tempH + 1
    ElseIf tempH > 1 Then
        tempH = tempH - 1
    End If
    
    If 6 * tempH < 1 Then
        HueToRGB = temp2 + (temp1 - temp2) * 6 * tempH
    ElseIf 2 * tempH < 1 Then
        HueToRGB = temp1
    ElseIf 3 * tempH < 2 Then
        HueToRGB = temp2 + (temp1 - temp2) * ((2 / 3) - tempH) * 6
    Else
        HueToRGB = temp2
    End If
End Function

Public Function max(ByVal val1 As Variant, ByVal val2 As Variant) As Variant
    If val1 > val2 Then
        max = val1
    Else
        max = val2
    End If
End Function

Public Function min(ByVal val1 As Variant, ByVal val2 As Variant) As Variant
    If val1 < val2 Then
        min = val1
    Else
        min = val2
    End If
End Function

Sub CountSelection()
    Dim data As Range
    Dim myChart As ChartObject
    Dim first_table_row As Integer
    Dim first_table_col As Integer
    
    Dim selection_n_cols As Long, selection_n_rows As Long, i As Long, j As Long, aux As Long
    
    Dim LastSheetRow As Long, LastSheetColumn As Long, FirstSheetRow As Long, FirstSheetColumn As Long
    Dim selection_first_col As Long, selection_first_row As Long, selection_last_col As Long, selection_last_row As Long
    
    Dim digits() As Long
    Dim total_digits() As Long
    Dim max_values() As Long
    Dim min_values() As Long
    
    Dim dark_purple As Long
    Dim medium_purple As Long
    Dim light_purple As Long
    Dim good As Long
    Dim meh As Long
    Dim bad As Long
    dark_purple = RGB(217, 217, 217)
    medium_purple = RGB(217, 217, 217)
    light_purple = RGB(247, 247, 247)
    good = RGB(176, 216, 164)
    meh = RGB(254, 225, 145)
    bad = RGB(253, 128, 96)
    
    Dim hasHeaders As Boolean

    
    On Error GoTo ErrorHandler
    
    Set data = Application.InputBox( _
      Title:="Number Format Rule From Cell", _
      Prompt:="Select a cell to pull in your number format rule", _
      Type:=8)
      
    ' Get the dimensions of the used range in data's parent sheet
    With data.Parent
        FirstSheetRow = .UsedRange.Row
        FirstSheetColumn = .UsedRange.Column
        LastSheetRow = FirstSheetRow + .UsedRange.Rows.count - 1
        LastSheetColumn = FirstSheetColumn + .UsedRange.Columns.count - 1
    End With
    
    ' Get the dimensions of the data selection
    selection_n_cols = data.Columns.count()
    selection_n_rows = data.Rows.count()
    selection_first_row = data.Cells(1, 1).Row
    selection_first_col = data.Cells(1, 1).Column
    selection_last_row = selection_first_row + selection_n_rows - 1
    selection_last_col = selection_first_col + selection_n_cols - 1


    
    ' Check if the data selection is within the used range of data's parent sheet
    If selection_first_row > LastSheetRow Or _
       selection_first_col > LastSheetColumn Or _
       selection_last_row < FirstSheetRow Or _
       selection_last_col < FirstSheetColumn Then
        ' If not, show an error message and exit the subroutine
        MsgBox "Error: Selection Empty"
        Exit Sub
    End If

    ' Calculate the row and column offsets of the data selection within the used range
    i = max(1, FirstSheetRow - selection_first_row)
    j = max(1, FirstSheetColumn - selection_first_col)
    
    ' Trim the size of the data selection to fit within the used range
    selection_n_rows = min(selection_n_rows, LastSheetRow - selection_first_row + 1)
    selection_n_cols = min(selection_n_cols, LastSheetColumn - selection_first_col + 1)
    
    ' Initialize arrays for storing data about the data selection
    ReDim digits(1 To 9, 1 To selection_n_cols) As Long
    ReDim total_digits(1 To selection_n_cols) As Long
    ReDim max_values(1 To selection_n_cols) As Long
    ReDim min_values(1 To selection_n_cols) As Long


    
    ' Loop through each column in the data range
    For j = 1 To selection_n_cols
        ' Initialize the minimum value for the current column to the maximum possible value
        min_values(j) = 2147483647
        
        ' Loop through each row in the data range
        For i = 1 To selection_n_rows
            ' Check if the current cell is not empty and contains a numeric value
            If Not IsEmpty(data.Cells(i, j)) And IsNumeric(data.Cells(i, j)) Then
                ' Extract the first character from the cell value
                aux = Left(data.Cells(i, j).Value, 1)
                
                ' If the character is greater than 0, count it as a digit and increment the total digit count
                If aux > 0 Then
                    digits(aux, j) = digits(aux, j) + 1
                    total_digits(j) = total_digits(j) + 1
                End If
                
                ' Update the maximum value for the current column using the max function
                max_values(j) = max(max_values(j), data.Cells(i, j).Value)
                
                ' If the cell value is greater than 0, update the minimum value for the current column using the min function
                If CLng(data.Cells(i, j).Value) > 0 Then
                    min_values(j) = min(min_values(j), data.Cells(i, j).Value)
                End If
            End If
        Next i
    Next j
    
    '-----------------------CHECK HEADERS-----------------------
    
    ' Check if the first row contains headers
    hasHeaders = False
    
    ' Iterate through the columns in the first row
    For i = FirstSheetColumn To LastSheetColumn
      ' If the current cell is not numeric but the cell below it is numeric, then it is likely a header
      If Not IsNumeric(data.Parent.Cells(FirstSheetRow, i)) And IsNumeric(data.Parent.Cells(FirstSheetRow + 1, i)) Then
        hasHeaders = True
        Exit For
      End If
    Next i

    
    ' -----------------------CREATE TABLE-----------------------
    
    Call NewBlankWorksheet
    
    ' Set the starting row and column for the table
    first_table_col = 2
    first_table_row = 4
    
    ' Populate the first column with the digits from 1 to 9
    ' Populate the second column with the values from Benford's Law
    For i = 1 To 9
      Sheets(1).Cells(first_table_row + i, first_table_col).Value = i
      Sheets(1).Cells(first_table_row + i, first_table_col + 1).Value = Log(1 + 1 / i) / Log(10)
    Next i


    Sheets(1).Cells(first_table_row, first_table_col).Value = "digits"
    Sheets(1).Cells(first_table_row, first_table_col + 1).Value = "Benford's Law Frequency"
    Sheets(1).Cells(first_table_row + 10, first_table_col).Value = "Total Entries > 0:"
    Sheets(1).Cells(first_table_row + 11, first_table_col).Value = "Chi Square (X^2):"
    Sheets(1).Cells(first_table_row + 12, first_table_col).Value = "p-value:"
    Sheets(1).Cells(first_table_row + 13, first_table_col).Value = "Follows Benford's Law?"
    Sheets(1).Cells(first_table_row + 14, first_table_col).Value = "Max value:"
    Sheets(1).Cells(first_table_row + 15, first_table_col).Value = "Min value:"
    
    Dim ws As Worksheet  ' Declare a variable to hold a reference to the worksheet
    Set ws = Sheets(1)  ' Set the variable to reference the first worksheet in the workbook
    
    For j = 1 To selection_n_cols  ' Loop through the number of columns in the selection
        If hasHeaders Then  ' Check if the selection has headers
            ' If the selection has headers, copy the header value from the selection to the worksheet
            ws.Cells(first_table_row, first_table_col + j + 1).Value = data.Parent.Cells(FirstSheetRow, j + selection_first_col - 1).Value
        Else
            ' If the selection does not have headers, assign a default header value to the worksheet
            ws.Cells(first_table_row, first_table_col + j + 1).Value = "Column " & Split(Cells(1, j + selection_first_col - 1).Address(True, False), "$")(0)
        End If
        
        ' Loop through the number of digits (1-9)
        For i = 1 To 9
            If total_digits(j) = 0 Then  ' Check if there are no digits in the current column
                ' If there are no digits, set the cell value to 0
                ws.Cells(first_table_row + i, first_table_col + j + 1).Value = 0
            Else
                ' If there are digits, calculate the proportion of each digit and assign it to the cell
                ws.Cells(first_table_row + i, first_table_col + j + 1).Value = digits(i, j) / total_digits(j)
                
                ' Calculate the chi-squared statistic for the current column and assign it to the cell
                ws.Cells(first_table_row + 11, first_table_col + j + 1).Value = ws.Cells(first_table_row + 11, first_table_col + j + 1).Value + (((digits(i, j) - (total_digits(j) * ws.Cells(first_table_row + i, first_table_col + 1).Value)) ^ 2) / (total_digits(j) * ws.Cells(first_table_row + i, first_table_col + 1).Value))
            End If
        Next i
        
        ' Assign the total number of digits in the current column to the cell
        ws.Cells(first_table_row + 10, first_table_col + j + 1).Value = total_digits(j)
        
        ' Calculate the p-value for the chi-squared statistic and assign it to the cell
        ws.Cells(first_table_row + 12, first_table_col + j + 1).Value = Application.WorksheetFunction.ChiSq_Dist_RT(ws.Cells(first_table_row + 11, first_table_col + j + 1).Value, 8)
        
    
        ' Check if the p-value is greater than the significance level (0.05)
        If ws.Cells(first_table_row + 12, first_table_col + j + 1).Value > 0.05 Then
            ' If the p-value is greater than the significance level, set the cell value to "Yes?" and color it "good"
            ws.Cells(first_table_row + 13, first_table_col + j + 1).Value = "Yes?"
            ws.Cells(first_table_row + 13, first_table_col + j + 1).Interior.Color = good
        Else
            ' If the p-value is not greater than the significance level, set the cell value to "No?" and color it "bad"
            ws.Cells(first_table_row + 13, first_table_col + j + 1).Value = "No?"
            ws.Cells(first_table_row + 13, first_table_col + j + 1).Interior.Color = bad
        End If
        
        ' Assign the maximum value in the current column to the cell
        ws.Cells(first_table_row + 14, first_table_col + j + 1).Value = max_values(j)
        
        ' Assign the minimum value in the current column to the cell
        ws.Cells(first_table_row + 15, first_table_col + j + 1).Value = min_values(j)
    Next j


    '-----------------------FORMAT TABLE-----------------------
    
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 10, first_table_col), Sheets(1).Cells(first_table_row + 10, first_table_col + 1)).MergeCells = True
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 11, first_table_col), Sheets(1).Cells(first_table_row + 11, first_table_col + 1)).MergeCells = True
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 12, first_table_col), Sheets(1).Cells(first_table_row + 12, first_table_col + 1)).MergeCells = True
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 13, first_table_col), Sheets(1).Cells(first_table_row + 13, first_table_col + 1)).MergeCells = True
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 14, first_table_col), Sheets(1).Cells(first_table_row + 14, first_table_col + 1)).MergeCells = True
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 15, first_table_col), Sheets(1).Cells(first_table_row + 15, first_table_col + 1)).MergeCells = True
    
    With Sheets(1).Range(Sheets(1).Cells(first_table_row, first_table_col), Sheets(1).Cells(first_table_row + 15, first_table_col + selection_n_cols + 1))
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlNone
        .EntireColumn.AutoFit
    End With
    With Sheets(1).Range(Sheets(1).Cells(first_table_row, first_table_col), Sheets(1).Cells(first_table_row + 9, first_table_col + selection_n_cols + 1))
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .NumberFormat = "0.000"
        .Interior.Color = light_purple
        
    End With
    With Sheets(1).Range(Sheets(1).Cells(first_table_row + 10, first_table_col), Sheets(1).Cells(first_table_row + 15, first_table_col + selection_n_cols + 1))
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    
    With Sheets(1).Range(Sheets(1).Cells(first_table_row, first_table_col), Sheets(1).Cells(first_table_row, first_table_col + selection_n_cols + 1))
        .Font.Bold = True
        .Interior.Color = dark_purple
    End With
    With Sheets(1).Range(Sheets(1).Cells(first_table_row + 1, first_table_col), Sheets(1).Cells(first_table_row + 15, first_table_col))
        .Font.Bold = True
        .NumberFormat = "0"
        .Interior.Color = medium_purple
    End With
    With Sheets(1).Range(Sheets(1).Cells(first_table_row - 1, first_table_col), Sheets(1).Cells(first_table_row - 1, first_table_col + selection_n_cols + 1))
        .MergeCells = True
        .Value = data.Parent.Name
        .Font.Size = 24
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .EntireColumn.AutoFit
    End With
    
    Sheets(1).Range(Sheets(1).Cells(first_table_row + 11, first_table_col + 2), Sheets(1).Cells(first_table_row + 11, first_table_col + selection_n_cols + 1)).NumberFormat = "0.000"
    
    For j = 1 To selection_n_cols
        With Sheets(1).Cells(first_table_row + 10, first_table_col + j + 1)
            If .Value >= 1000 Then
                .Interior.Color = good
            ElseIf .Value >= 500 Then
                .Interior.Color = meh
            Else
                .Interior.Color = bad
            End If
        End With
        
        With Sheets(1).Cells(first_table_row + 12, first_table_col + j + 1)
            If .Value <= 0.05 Then
                .Interior.Color = bad
                .NumberFormat = "0.00E+00"
            ElseIf .Value <= 0.1 Then
                .Interior.Color = meh
                .NumberFormat = "0.000"
            Else
                .Interior.Color = good
                .NumberFormat = "0.000"
            End If
        End With
        
        If Len(Sheets(1).Cells(first_table_row + 14, first_table_col + j + 1).Value) - Len(Sheets(1).Cells(first_table_row + 15, first_table_col + j + 1).Value) <= 1 Then
            Sheets(1).Cells(first_table_row + 14, first_table_col + j + 1).Interior.Color = bad
            Sheets(1).Cells(first_table_row + 15, first_table_col + j + 1).Interior.Color = bad
        ElseIf Len(Sheets(1).Cells(first_table_row + 14, first_table_col + j + 1).Value) - Len(Sheets(1).Cells(first_table_row + 15, first_table_col + j + 1).Value) <= 3 Then
            Sheets(1).Cells(first_table_row + 14, first_table_col + j + 1).Interior.Color = meh
            Sheets(1).Cells(first_table_row + 15, first_table_col + j + 1).Interior.Color = meh
        Else
            Sheets(1).Cells(first_table_row + 14, first_table_col + j + 1).Interior.Color = good
            Sheets(1).Cells(first_table_row + 15, first_table_col + j + 1).Interior.Color = good
        End If
    Next j
        
        
    '-----------------------CREATE CHART-----------------------
    
    Set myChart = Sheets(1).ChartObjects.Add( _
        Left:=Sheets(1).Cells(first_table_row, first_table_col + selection_n_cols + 4).Left, _
        Width:=500, _
        Top:=Sheets(1).Cells(first_table_row - 1, first_table_col).Top, _
        Height:=300)
    
    With myChart.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Frequency analysis by column"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    With myChart.Chart.SeriesCollection.NewSeries
        .XValues = Sheets(1).Range(Sheets(1).Cells(first_table_row + 1, first_table_col), Sheets(1).Cells(first_table_row + 9, first_table_col))
        .Values = Sheets(1).Range(Sheets(1).Cells(first_table_row + 1, first_table_col + 1), Sheets(1).Cells(first_table_row + 9, first_table_col + 1))
        .Name = Sheets(1).Cells(first_table_row, first_table_col + 1)
        .Format.Fill.ForeColor.RGB = RGB(255, 158, 0)
    End With
    
    For j = 1 To selection_n_cols
        With myChart.Chart.SeriesCollection.NewSeries
            .XValues = Sheets(1).Range(Sheets(1).Cells(first_table_row + 1, first_table_col), Sheets(1).Cells(first_table_row + 9, first_table_col))
            .Values = Sheets(1).Range(Sheets(1).Cells(first_table_row + 1, first_table_col + j + 1), Sheets(1).Cells(first_table_row + 9, first_table_col + j + 1))
            .Name = Sheets(1).Cells(first_table_row, first_table_col + j + 1)
            .Format.Fill.ForeColor.RGB = HSLToRGB(0.6, 0.9, 0.9 * j / (selection_n_cols + 1) + 0.1)
        End With
    Next j

    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 424 Then
        ' Input box was canceled
        Exit Sub
    Else
        MsgBox "An error has occurred: " & Err.Description
        Exit Sub
    End If

End Sub


