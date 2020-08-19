Attribute VB_Name = "Module1"
Option Explicit

Sub RunForm()

Application.ScreenUpdating = False
Worksheets("Sheet1").Visible = True

PopulateStates
UserForm1.Show

End Sub

Sub PopulateStates()

Dim tWB As Workbook
Dim ncategories As Integer, i As Integer

Set tWB = Application.Workbooks("Geography Toolbox")

Sheets("Sheet1").Select: Sheets("Sheet1").Range("E1").Select

ncategories = WorksheetFunction.CountA(Columns("E:E"))

For i = 1 To ncategories
    UserForm1.state1select.AddItem Range("E1:E" & ncategories).Cells(i, 1)
    UserForm1.state2select.AddItem Range("E1:E" & ncategories).Cells(i, 1)
Next i

UserForm1.state1select.Text = Range("E1:E" & ncategories).Cells(1, 1)
UserForm1.state2select.Text = Range("E1:E" & ncategories).Cells(1, 1)

Application.ScreenUpdating = True

End Sub

