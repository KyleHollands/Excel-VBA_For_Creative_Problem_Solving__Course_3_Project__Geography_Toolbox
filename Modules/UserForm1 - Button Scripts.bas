VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Geography Toolbox"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14835
   OleObjectBlob   =   "UserForm1 - Button Scripts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub city1input_Change()

End Sub

Private Sub GoButton1_Click()

Application.ScreenUpdating = False

Dim tWB As Workbook
Dim i As Integer
Dim lat1 As Double, lat2 As Double, lon1 As Double, lon2 As Double, d As Double
Dim ngroups As Variant

Set tWB = ThisWorkbook

tWB.Sheets("Sheet1").Range("A1").Select

ngroups = WorksheetFunction.CountA(Columns("A:A"))

If UserForm1.City1Label = "" Or UserForm1.City2Label = "" Then
    MsgBox ("One of the city inputs is blank.")
    Exit Sub
End If

For i = 1 To ngroups
    If Range("A1:A" & ngroups).Cells(i, 1) = UserForm1.City1Label Then
        Range("A" & i).Select
        lat1 = ActiveCell.Offset(0, 1)
        lon1 = ActiveCell.Offset(0, 2)
    End If
Next i

For i = 1 To ngroups
    If Range("A1:A" & ngroups).Cells(i, 1) = UserForm1.City2Label Then
        Range("A" & i).Select
        lat2 = ActiveCell.Offset(0, 1)
        lon2 = ActiveCell.Offset(0, 2)
    End If
Next i

d = Distance(lat1, lat2, lon1, lon2)

MsgBox (UserForm1.City1Label & " and " & UserForm1.City2Label & " are " & FormatNumber(d, 0) & " miles apart as the crow flies.")

Application.ScreenUpdating = True

End Sub

Private Sub GoButton2_Click()

Application.ScreenUpdating = False

Dim tWB As Workbook
Dim i As Integer
Dim lat1 As Double, lat2 As Double, lon1 As Double, lon2 As Double, d As Double
Dim ngroups As Variant

Set tWB = ThisWorkbook

tWB.Sheets("Sheet1").Range("A1").Select

ngroups = WorksheetFunction.CountA(Columns("A:A"))

For i = 1 To ngroups
    If Range("A1:A" & ngroups).Cells(i, 1) = UserForm1.city1select Then
        Range("A" & i).Select
        lat1 = ActiveCell.Offset(0, 1)
        lon1 = ActiveCell.Offset(0, 2)
    End If
Next i

For i = 1 To ngroups
    If Range("A1:A" & ngroups).Cells(i, 1) = UserForm1.city2select Then
        Range("A" & i).Select
        lat2 = ActiveCell.Offset(0, 1)
        lon2 = ActiveCell.Offset(0, 2)
    End If
Next i

d = Distance(lat1, lat2, lon1, lon2)

MsgBox (UserForm1.city1select.Text & " and " & UserForm1.city2select.Text & " are " & FormatNumber(d, 0) & " miles apart as the crow flies.")



End Sub

Private Sub state1select_Change()

Application.ScreenUpdating = False


Dim i As Integer, ngroups As Integer, j As Integer
Dim tWB As Workbook

UserForm1.city1select.Clear

Set tWB = ThisWorkbook

tWB.Sheets("Sheet1").Range("A1").Select

ngroups = WorksheetFunction.CountA(Columns("A:A"))

For i = 1 To ngroups
    j = 1
    If Range("A1:A" & ngroups).Cells(i, 1) = UserForm1.state1select.Text Then
        Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            UserForm1.city1select.AddItem ActiveCell.Offset(j, 0)
            j = j + 1
        Loop
        UserForm1.city1select = ActiveCell.Offset(1, 0)
        Exit For
    End If
Next i

Application.ScreenUpdating = True


End Sub

Private Sub state2select_Change()

Application.ScreenUpdating = False

Dim i As Integer, ngroups As Integer, j As Integer
Dim tWB As Workbook

UserForm1.city2select.Clear

Set tWB = ThisWorkbook

tWB.Sheets("Sheet1").Range("A1").Select

ngroups = WorksheetFunction.CountA(Columns("A:A"))

For i = 1 To ngroups
    j = 1
    If Range("A1:A" & ngroups).Cells(i, 1) = UserForm1.state2select.Text Then
        Range("A" & i).Select
        Do While Not IsEmpty(ActiveCell.Offset(j, 0))
            UserForm1.city2select.AddItem ActiveCell.Offset(j, 0)
            j = j + 1
        Loop
        UserForm1.city2select.Text = ActiveCell.Offset(1, 0)
        Exit For
    End If
Next i

Application.ScreenUpdating = True

End Sub

Private Sub SearchButton1_Click()

Dim i As Integer, j As Integer, k As Integer
Dim Cities() As String, States() As String
Dim Ans As Integer, commaposition

If UserForm1.city1input.Text = "" Then
    MsgBox ("You cannot leave this blank.")
    Exit Sub
End If

On Error GoTo Reset

For i = 1 To 855
    commaposition = InStr(UserForm1.city1input.Text, ",")
    If commaposition = 0 Then
        If UCase(Left(Range("A1:A855").Cells(i, 1), Len(UserForm1.city1input.Text))) = UCase(UserForm1.city1input.Text) _
        And Not UCase(Range("A1:A855").Cells(i, 1)) = Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
            ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not UCase(ActiveCell.Offset(-j, 0)) = ActiveCell.Offset(-j, 0) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j, 0)
                    Exit Do
                End If
            Loop
        End If
    Else
        If UCase(Left(Range("A1:A855").Cells(i, 1), commaposition - 1)) = UCase(Left(UserForm1.city1input.Text, commaposition - 1)) _
        And Not UCase(Range("A1:A855").Cells(i, 1)) = Range("A1:A855").Cells(i, 1) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
        ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not IsEmpty(ActiveCell.Offset(-j, 0)) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j + 1, 0)
                    Exit Do
                End If
            Loop
        End If
    End If
Next i

If k > 1 Then
    UserForm2.ComboBox1.Clear
    For j = 1 To k
        UserForm2.ComboBox1.AddItem Cities(j) & ", " & States(j)
    Next j
    UserForm2.ComboBox1.Text = Cities(1) & ", " & States(1)
    UserForm2.Show
    City1Label = Cities(UserForm2.ComboBox1.ListIndex + 1)
    State1Label = States(UserForm2.ComboBox1.ListIndex + 1)
Else
    Ans = MsgBox("Did you mean " & Cities(1) & ", " & States(1) & "?", vbYesNo)
    If Ans = 7 Then
        MsgBox "Sorry, that's the only location we could find meeting your search criterion."
        Exit Sub
    End If
    City1Label = Cities(1)
    State1Label = States(1)
End If

Exit Sub

Reset:
    MsgBox ("You must enter a valid city.")

End Sub

Private Sub SearchButton2_Click()

Dim i As Integer, j As Integer, k As Integer
Dim Cities() As String, States() As String
Dim Ans As Integer, commaposition

If UserForm1.city2input.Text = "" Then
    MsgBox ("You cannot leave this blank.")
    Exit Sub
End If

On Error GoTo Reset

For i = 1 To 855
    commaposition = InStr(UserForm1.city2input.Text, ",")
    If commaposition = 0 Then
        If UCase(Left(Range("A1:A855").Cells(i, 1), Len(UserForm1.city2input.Text))) = UCase(UserForm1.city2input.Text) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
            ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not UCase(ActiveCell.Offset(-j, 0)) = ActiveCell.Offset(-j, 0) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j, 0)
                    Exit Do
                End If
            Loop
        End If
    Else
        If UCase(Left(Range("A1:A855").Cells(i, 1), commaposition - 1)) = UCase(Left(UserForm1.city2input.Text, commaposition - 1)) Then
            Range("A1:A855").Cells(i, 1).Select
                k = k + 1
        ReDim Preserve Cities(k) As String
            ReDim Preserve States(k) As String
            Cities(k) = ActiveCell
            j = 1
            Do
                If Not IsEmpty(ActiveCell.Offset(-j, 0)) Then
                    j = j + 1
                Else
                    States(k) = ActiveCell.Offset(-j + 1, 0)
                    Exit Do
                End If
            Loop
        End If
    End If
Next i

If k > 1 Then
    UserForm2.ComboBox1.Clear
    For j = 1 To k
        UserForm2.ComboBox1.AddItem Cities(j) & ", " & States(j)
    Next j
    UserForm2.ComboBox1.Text = Cities(1) & ", " & States(1)
    UserForm2.Show
    City2Label = Cities(UserForm2.ComboBox1.ListIndex + 1)
    State2Label = States(UserForm2.ComboBox1.ListIndex + 1)
Else
    Ans = MsgBox("Did you mean " & Cities(1) & ", " & States(1) & "?", vbYesNo)
    If Ans = 7 Then
        MsgBox "Sorry, that's the only location we could find meeting your search criterion."
        Exit Sub
    End If
    City2Label = Cities(1)
    State2Label = States(1)
End If

Exit Sub

Reset:
    MsgBox ("You must enter a valid city.")

End Sub

Function Distance(lat1 As Double, lat2 As Double, lon1 As Double, lon2 As Double) As Double
'YOUR CODE GOES HERE - CALCULATION AND OUTPUT OF DISTANCE BETWEEN TWO POINTS
Dim pi As Double, Rad As Double, a As Double, b As Double, c As Double

pi = WorksheetFunction.pi()
Rad = 3960

a = Cos(lat1 * pi / 180) * Cos(lat2 * pi / 180) * Cos(lon1 * pi / 180) * Cos(lon2 * pi / 180)
b = Cos(lat1 * pi / 180) * Sin(lon1 * pi / 180) * Cos(lat2 * pi / 180) * Sin(lon2 * pi / 180)
c = Sin(lat1 * pi / 180) * Sin(lat2 * pi / 180)

Distance = WorksheetFunction.Acos(a + b + c) * Rad

End Function

Private Sub QuitButton_Click()

Worksheets("Sheet1").Visible = False
Unload UserForm1

End Sub

Private Sub ResetButton_Click()

Unload UserForm1: Call RunForm

End Sub


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ' Your codes
        ' Tip: If you want to prevent closing UserForm by Close (×) button in the right-top corner of the UserForm, just uncomment the following line:
         Cancel = True
    End If
End Sub
