VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} departman 
   Caption         =   "Tracker Türünü Deðiþtir"
   ClientHeight    =   2928
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3636
   OleObjectBlob   =   "departman.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "departman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
If ComboBox1.Value = "Mold Quality" Or ComboBox1.Value = "Engineering" Then
OptionButton1.Enabled = True
OptionButton2.Enabled = True
Else
OptionButton1.Enabled = False
OptionButton2.Enabled = False
End If
End Sub

Private Sub CommandButton1_Click()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
Application.Calculation = xlManual
Application.ScreenUpdating = False
If TextBox1.Value = "1709" Then
    wstracker.Unprotect Password:="4135911"
    Application.ScreenUpdating = False
    If OptionButton1.Value = True Then
    wstracker.Cells(1, 21) = "Vestas"
    Else
    wstracker.Cells(1, 21) = "Nordex"
    End If
    If ComboBox1.Value = "Finish Quality" Then
    wstracker.Cells(4, 1) = "NCR MANAGEMENT - FINISH QUALITY"
    ElseIf ComboBox1.Value = "Mold Quality" Then
    wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY"
    ElseIf ComboBox1.Value = "Production" Then
    wstracker.Cells(4, 1) = "NCR MANAGEMENT - PRODUCTION"
    ElseIf ComboBox1.Value = "Engineering" Then
    wstracker.Cells(4, 1) = "NCR MANAGEMENT - ENGINEERING"
    ElseIf ComboBox1.Value = "UT" Then
    wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT"
    ElseIf ComboBox1.Value = "Read-Only" Then
    wstracker.Cells(4, 1) = "NCR MANAGEMENT - Read Only"
    End If
    Call refresh
Else
    MsgBox "Þifre hatalý"
    Exit Sub
End If
Unload Me
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Read-Only", "Production", "Mold Quality", "Finish Quality", "Engineering", "UT")
End Sub

Private Sub UserForm_Terminate()
Call refresh
End Sub
