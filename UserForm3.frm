VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Çýktý Al"
   ClientHeight    =   5316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6840
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox2.ListCount - 1 To 0 Step -1
        If Me.ListBox2.Selected(iCtr) = True Then
            Me.ListBox2.RemoveItem iCtr
        End If
    Next iCtr

End Sub

Private Sub CommandButton2_Click()
    Dim iCtr As Long
    For iCtr = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox2.AddItem Me.ListBox1.List(iCtr)
        End If
    Next iCtr

    For iCtr = Me.ListBox1.ListCount - 1 To 0 Step -1
        If Me.ListBox1.Selected(iCtr) = True Then
            Me.ListBox1.RemoveItem iCtr
        End If
    Next iCtr
End Sub

Private Sub CommandButton3_Click()
    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox1.ListCount - 1
        Me.ListBox2.AddItem Me.ListBox1.List(iCtr)
    Next iCtr

    Me.ListBox1.Clear
End Sub

Private Sub CommandButton4_Click()
    Dim iCtr As Long

    For iCtr = 0 To Me.ListBox2.ListCount - 1
        Me.ListBox1.AddItem Me.ListBox2.List(iCtr)
    Next iCtr

    Me.ListBox2.Clear
End Sub

Private Sub CommandButton5_Click()
        Dim departman As String
        If ListBox2.ListCount < 1 Then
        MsgBox "Kanat seçilmedi"
        End If
        Dim wstracker As Worksheet
            Dim Rng, rng2, rng10, rng11, rng12 As Range
            Dim c As Range
            Dim Item As Variant
'            Dim list2 As New Collection
            Dim list3(500) As Variant
'            For i = 0 To ListBox2.ListCount - 1
'                list2.Add ListBox2.List(i), CStr(ListBox2.List(i))
'            Next i
            Application.ScreenUpdating = False
        '   *** Change wstrackereet name to suit ***
        '   wstracker.ListObjects("Tablo1"
            Set wstracker = Worksheets("NCR TRACKER")
            departman = wstracker.Cells(4, 1)
            wstracker.Unprotect Password:="4135911"
            wstracker.Cells(5, 1) = "Tamir Yeri"
            LastRow = wstracker.Range("D65536").End(xlUp).Row
            Set Rng = wstracker.Range("D6:D" & LastRow)
            On Error GoTo 0
            Set Rng = wstracker.Range("A4:R" & LastRow)
            Set rng11 = wstracker.Range("AA4:AA" & LastRow)
            'Set Rng = Union(rng10, rng11)
            Set rng2 = wstracker.Range("A6:R" & LastRow)
            With wstracker.PageSetup
            .Orientation = xlLandscape
            .FitToPagesWide = 1
            .FitToPagesTall = False
            For i = 0 To ListBox2.ListCount - 1
                list3(i) = ListBox2.List(i)
            Next i
                wstracker.ListObjects("Tablo1").Range.AutoFilter 4, Criteria1:=list3, Operator:=xlFilterValues
                 wstracker.Cells(4, 1) = Item & " Rapor Tarihi: " & Now()
                    For Each cell In rng2.Columns(1).Cells.SpecialCells(xlCellTypeVisible)
                        ' BURAYA ÇALIÞILACAK
                        If wstracker.Cells(cell.Row, 26) = "UT" Then
                        wstracker.Cells(cell.Row, 1) = "UT"
                        ElseIf (wstracker.Cells(cell.Row, 8) = "SS" And wstracker.Cells(cell.Row, 10) = "A") _
                        Or (wstracker.Cells(cell.Row, 8) = "PS" And wstracker.Cells(cell.Row, 10) = "B") Then
                        wstracker.Cells(cell.Row, 1) = "Ýç Tamir"
                        Else
                        wstracker.Cells(cell.Row, 1) = "Trim"
                        End If
                    Next cell
                 wstracker.ListObjects("Tablo1").Range.Sort Key1:=Range("D5"), _
                                          DataOption1:=xlSortNormal, _
                                          Header:=xlYes
                 Rng.PrintOut , copies:=ComboBox1.Value
            End With
            wstracker.Cells(4, 1) = departman
            wstracker.Cells(5, 1) = "ID"
            Call refresh
            Unload Me
            Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.List = Array("1", "2", "3", "4", "5")
    ComboBox1.Value = "1"
    Me.ListBox1.MultiSelect = fmMultiSelectMulti
    Me.ListBox2.MultiSelect = fmMultiSelectMulti
        For Each Item In Listkanat
        ListBox1.AddItem Item
        Next Item
End Sub
Private Sub UserForm_Terminate()
Call refresh
End Sub
