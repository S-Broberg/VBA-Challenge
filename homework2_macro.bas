Attribute VB_Name = "Module1"

Sub macro2()
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim percent1 As Double
Dim percent2 As Double
Dim v1 As Long
Dim v2 As Long
Dim v3 As Long
Dim volume As Long


Dim lastrow As Long
'1

Cells(1, 9).Value = "ticker"
Range("A:A").Copy
Range("I:I").Insert
Cells(1, 10).Value = "Yearly change"
Cells(1, 11).Value = "Percent change"
Cells(1, 12).Value = "Total Stock Volume"
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'2
For i = 2 To lastrow
percent1 = Sheets("2014").Cells(i, 3)
percent2 = Sheets("2016").Cells(i, 3)

yearly_change = percent2 - percent1


Cells(i, 10).Value = yearly_change
'3
If percent1 = 0 Then
percent_change = 0
ElseIf percent2 > percent1 Then
percent_change = (yearly_change / percent1) * -100
Cells(i, 10).Interior.ColorIndex = 4
ElseIf percent2 < percent1 Then

percent_change = (yearly_change / percent1) * 100
Cells(i, 10).Interior.ColorIndex = 3
End If
Cells(i, 11).Value = percent_change

'4
v1 = Sheets("2014").Cells(i, 7)
v2 = Sheets("2015").Cells(i, 7)
v3 = Sheets("2016").Cells(i, 7)
volume = v1 + v2 + v3
Cells(i, 12).Value = volume



Next i
End Sub

