Attribute VB_Name = "Module1"
Option Explicit
Public id As Integer
Public idrow As Integer
Public oraclenovar, hop2 As Boolean
Public yenino, ttgrupcombo As String
Public Listkanat As New Collection
Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export "C:\Users\ftek\OneDrive - TPI Composites Inc\Masaüstü\Vba\" & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub

Sub goster()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
On Error GoTo err
        Dim sonuc As Integer
        On Error Resume Next
        sonuc = InputBox("Enter the password", "Database path change")
        If sonuc = "1474" Then
            wstracker.Unprotect Password:="4135911"
            Dim fname As Variant
            fname = Application.GetOpenFilename(filefilter:="Access Files,*.acc*")
            wstracker.Cells(1, 9) = fname
            With wstracker
                .Protect Password:="4135911", AllowFiltering:=True
                .EnableSelection = xlNoRestrictions
            End With
        Else
            MsgBox "Incorrect Password"
            Exit Sub
        End If
        On Error GoTo 0
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub
Sub goster2()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
On Error GoTo err
        Dim sonuc As Integer
        On Error Resume Next
        sonuc = InputBox("Enter the password", "Photograph path change")
        If sonuc = "1474" Then
            wstracker.Unprotect Password:="4135911"
            
                Dim sFolder As String
                
                With Application.FileDialog(msoFileDialogFolderPicker)
                .InitialFileName = "C:\"
                If .Show = -1 Then
                    sFolder = .SelectedItems(1)
                End If
                End With
                If sFolder <> "" Then
                wstracker.Cells(2, 9) = sFolder
                End If
        On Error GoTo 0
With wstracker
    .Protect Password:="4135911", AllowFiltering:=True
    .EnableSelection = xlNoRestrictions
    
End With
Else
            MsgBox "Incorrect Password"
            Exit Sub
        End If
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub
Sub refresh()
'On Error GoTo err
Application.DisplayStatusBar = True
Application.Calculation = xlManual
Application.ScreenUpdating = False
Dim wsozet1, wsozet2, wstracker As Worksheet
Set wsozet2 = ThisWorkbook.Worksheets("KANATBAÞI ÖZET")
Set wsozet1 = ThisWorkbook.Worksheets("KANAT ÖZETÝ")
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")

wstracker.Unprotect Password:="4135911"
wsozet2.Unprotect Password:="4135911"

Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsvers As ADODB.Recordset
Dim rspivot As ADODB.Recordset
Dim dbPath, keyword2 As String
Dim SQL, sqlvers, versiyon, SQLpivot, kanatno As String
Dim i, j, k, l, q, w, r, t As Integer
Dim prsay, qusay, utsay, engsay As Integer
Dim var
Dim Sht  As Worksheet
Dim MyLbl As OLEObject
Dim LastRow, lastrow2 As Long
prsay = 0
qusay = 0
utsay = 0
engsay = 0
j = 0
k = 0
l = 0



wsozet1.Range("A2:Z10000").ClearContents
wsozet2.Range("A2:Z500").ClearContents
wsozet2.Range("A2:Z500").Borders.LineStyle = xlLineStyleNone


dbPath = wstracker.Cells(1, 9)
Set cnn = New ADODB.Connection
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"

    sqlvers = "SELECT VERS FROM VERS"
    Set rsvers = New ADODB.Recordset
    rsvers.Open sqlvers, cnn
    If rsvers.EOF And rsvers.BOF Then
            rsvers.Close
            cnn.Close
            Set rsvers = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
    End If
        versiyon = rsvers.Fields(0).Value
        If wstracker.Cells(3, 9) <> versiyon Then
        MsgBox "Eski bir sürüm kullanýyorsunuz baðlantý yapýlmadý!"
        Exit Sub
        End If
Dim mainstr As String
Dim sqlproj As String
Dim proj As String
If wstracker.Cells(1, 21) = "Vestas" Then
proj = "V"
Else
proj = "N"
End If
sqlproj = "AND [Project] LIKE '%" & proj & "%' ORDER BY [OPEN DATE]"
mainstr = "SELECT ID, [ORACLE NUMBER], [PROJECT], [BLADE NUMBER], [MOLD NUMBER], [OPEN DATE], [DEFECT CODE / TYPE], [PS / SS], EDGE, SURFACE, [Z (mm)], [X/C (mm)], [Length], [Width], [Depth], EXPLANATION, SOURCE, STATUS, [NCR PICTURE], RR, SCOP, SCMIN, LAMOP, LAMMIN, TOTALMIN, QRESP, DISP, TAMIRTURU, KATSAYISI, [TAHMINI SURE] FROM NCRDB"
If wstracker.Cells(4, 1) = "NCR MANAGEMENT - FINISH QUALITY" Then
SQL = mainstr & " WHERE QRESP = 'Fin Q' AND (STATUS = 'Açýk' OR STATUS = '1. Aþama Onayý' OR STATUS = 'After Kontrol' OR STATUS = '2. Aþama Onayý' OR STATUS = '1. Aþama NOK' OR STATUS = '2. Aþama NOK' OR STATUS = '2. Aþama' OR STATUS = 'Müh. Onayý') ORDER BY [OPEN DATE]"
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY" Then
SQL = mainstr & " WHERE QRESP = 'Mold Q' AND (STATUS = 'Açýk' OR STATUS = '1. Aþama Onayý' OR STATUS = 'After Kontrol' OR STATUS = '2. Aþama Onayý' OR STATUS = '1. Aþama NOK' OR STATUS = '2. Aþama NOK' OR STATUS = '2. Aþama' OR STATUS = 'Müh. Onayý') " & sqlproj
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - ENGINEERING" Then
SQL = mainstr & " WHERE STATUS = 'Müh. Onayý' " & sqlproj
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" Then
SQL = mainstr & " WHERE QRESP = 'UT' AND (STATUS LIKE 'UT%' OR STATUS = 'AÇIK' OR STATUS = '1. Aþama Onayý' OR STATUS = 'After Kontrol' OR STATUS = '2. Aþama Onayý') ORDER BY [OPEN DATE]"
Else
SQL = mainstr & " WHERE [BLADE NUMBER] LIKE 'KANAT%' AND (STATUS = 'Açýk' OR STATUS = '1. Aþama Onayý' OR STATUS = '2. Aþama Onayý' OR STATUS = '1. Aþama NOK' OR STATUS = '2. Aþama NOK' OR STATUS = '2. Aþama' OR STATUS = 'Müh. Onayý' OR STATUS = 'After Kontrol' OR STATUS LIKE 'UT%') ORDER BY [PROJECT], [BLADE NUMBER], [OPEN DATE]"
End If
SQLpivot = "SELECT [PROJECT], [BLADE NUMBER], STATUS FROM NCRDB WHERE [BLADE NUMBER] LIKE 'KANAT%' AND ( STATUS = 'Açýk' OR STATUS = '1. Aþama Onayý' OR STATUS = 'After Kontrol' OR STATUS = '2. Aþama Onayý' OR STATUS = '1. Aþama NOK' OR STATUS = '2. Aþama NOK' OR STATUS = '2. Aþama' OR STATUS = 'Müh. Onayý' OR STATUS LIKE 'UT%') ORDER BY [PROJECT], [BLADE NUMBER]"

Set rs = New ADODB.Recordset
rs.Open SQL, cnn
If rs.EOF And rs.BOF Then
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
    MsgBox "Herhangi bir kayýt bulunamadý", vbCritical, "Kayýt Bulunamadý"
    With wstracker.ListObjects("Tablo1")
    .AutoFilter.ShowAllData
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Rows.Delete
    End If
End With
    Exit Sub
End If
MsgBox "a"
'ÖZET TABLO ÇIKAR
Set rspivot = New ADODB.Recordset
rspivot.Open SQLpivot, cnn


With wstracker.ListObjects("Tablo1")
    .AutoFilter.ShowAllData
    If Not .DataBodyRange Is Nothing Then
        .DataBodyRange.Rows.Delete
    End If
    Call .Range(2, 1).CopyFromRecordset(rs)
End With
wsozet1.Cells(2, 1).CopyFromRecordset rspivot
wsozet1.Range("C:C").Insert

LastRow = wsozet1.Range("A" & Rows.Count).End(xlUp).Row
For w = 2 To LastRow
    wsozet1.Cells(w, 3).Value = wsozet1.Cells(w, 1).Value & "-" & wsozet1.Cells(w, 2).Value
    wsozet1.Cells(w, 5).Value = 0
Next w
wsozet1.Columns("A:B").EntireColumn.Delete
wsozet1.Range("A2", wsozet1.Range("A2").End(xlDown)).Copy Destination:=wsozet2.Range("A2")
wsozet2.Range("A2", wsozet2.Range("A2").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
lastrow2 = wsozet2.Range("A" & Rows.Count).End(xlUp).Row
For q = 2 To lastrow2
    t = 0
    kanatno = wsozet2.Cells(q, 1).Value
    For r = 2 To LastRow
    If wsozet1.Cells(r, 1).Value = kanatno Then
    If wsozet1.Cells(r, 2).Value = "Açýk" Or wsozet1.Cells(r, 2).Value = "1. Aþama NOK" Or wsozet1.Cells(r, 2).Value = "2. Aþama NOK" Or wsozet1.Cells(r, 2).Value Like "UT M*" Or wsozet1.Cells(r, 2).Value Like "2. Aþama" Then
    wsozet2.Cells(q, 2).Value = wsozet2.Cells(q, 2).Value + 1
    t = t + 1
    prsay = prsay + 1
    ElseIf wsozet1.Cells(r, 2).Value = "1. Aþama Onayý" Or wsozet1.Cells(r, 2).Value = "2. Aþama Onayý" Or wsozet1.Cells(r, 2).Value = "After Kontrol" Then
    wsozet2.Cells(q, 3).Value = wsozet2.Cells(q, 3).Value + 1
    t = t + 1
    qusay = qusay + 1
    ElseIf wsozet1.Cells(r, 2).Value Like "UT O*" Then
    wsozet2.Cells(q, 4).Value = wsozet2.Cells(q, 4).Value + 1
    t = t + 1
    utsay = utsay + 1
    ElseIf wsozet1.Cells(r, 2).Value Like "Müh. Onayý" Then
    wsozet2.Cells(q, 5).Value = wsozet2.Cells(q, 5).Value + 1
    t = t + 1
    engsay = engsay + 1
    End If
    End If
    Next r
    wsozet2.Cells(q, 6).Value = t
Next q

'Formatlamaya baþla
wsozet2.Range("B2", wsozet2.Range("F2").End(xlDown)).NumberFormat = "0"
wsozet2.Range("B2", wsozet2.Range("F2").End(xlDown)).HorizontalAlignment = xlCenter
wsozet2.Range("F2", wsozet2.Range("F2").End(xlDown)).Font.FontStyle = "Bold"
wsozet2.Range("A2", wsozet2.Range("F2").End(xlDown)).Borders.LineStyle = xlContinuous


rspivot.Close
rs.Close
cnn.Close
Set rs = Nothing
Set cnn = Nothing
Set rsvers = Nothing
LastRow = wstracker.Range("A" & Rows.Count).End(xlUp).Row
MsgBox "a2"
For i = 6 To LastRow
wstracker.Hyperlinks.Add Anchor:=wstracker.Cells(i, 19), Address:=wstracker.Cells(i, 19).Value, TextToDisplay:="Link"
            If wstracker.Cells(4, 1) = "NCR MANAGEMENT - FINISH QUALITY" Then
                    If wstracker.Cells(i, 18).Value = "1. Aþama Onayý" Or wstracker.Cells(i, 18).Value = "2. Aþama Onayý" Or wstracker.Cells(i, 18).Value = "After Kontrol" Then
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 18).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 18).Font.ColorIndex = 1
                    End If
                    If wstracker.Cells(i, 26).Value = "Mold Q" Then
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 26).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 26).Font.ColorIndex = 1
                    End If
                    If wstracker.Cells(i, 2).Value = "000000" Then
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 2).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 2).Font.ColorIndex = 1
                    End If
            ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY" Then
                    If wstracker.Cells(i, 18).Value = "Açýk" Or _
                        wstracker.Cells(i, 18).Value = "2. Aþama" Or _
                        wstracker.Cells(i, 18).Value = "1. Aþama NOK" Or _
                        wstracker.Cells(i, 18).Value = "2. Aþama NOK" Or _
                        wstracker.Cells(i, 18).Value = "1. Aþama Onayý" Or _
                        wstracker.Cells(i, 18).Value = "2. Aþama Onayý" Or _
                        wstracker.Cells(i, 18).Value = "After Kontrol" Or _
                        wstracker.Cells(i, 18).Value Like "UT M*" _
                        Then
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 18).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 18).Font.ColorIndex = 1
                    End If
                    
                    
                    If wstracker.Cells(i, 26).Value = "Mold Q" Then
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 26).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 26).Font.ColorIndex = 1
                    End If
                    If wstracker.Cells(i, 2).Value = "000000" Then
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 2).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 2).Font.ColorIndex = 1
                    End If
                    
            ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - PRODUCTION" Then
                    If wstracker.Cells(i, 18).Value = "Açýk" Or _
                        wstracker.Cells(i, 18).Value = "2. Aþama" Or _
                        wstracker.Cells(i, 18).Value = "1. Aþama NOK" Or _
                        wstracker.Cells(i, 18).Value = "2. Aþama NOK" Or _
                        wstracker.Cells(i, 18).Value Like "UT M*" _
                        Then
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 18).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 18).Font.ColorIndex = 1
                    End If
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 26).Font.ColorIndex = 1
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 2).Font.ColorIndex = 1
            ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - ENGINEERING" Then
                    If wstracker.Cells(i, 18).Value = "Müh. Onayý" Then
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 18).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 18).Font.ColorIndex = 1
                    End If
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 26).Font.ColorIndex = 1
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 2).Font.ColorIndex = 1
            ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" Then
                    If wstracker.Cells(i, 18).Value Like "UT O*" Then
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 18).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 18).Font.ColorIndex = 1
                    End If
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 26).Font.ColorIndex = 1
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 2).Font.ColorIndex = 1
                    If wstracker.Cells(i, 2).Value = "000000" Then
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 2).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 2).Font.ColorIndex = 1
                    End If
                    If wstracker.Cells(i, 26).Value = "UT" Then
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleSingle
                        wstracker.Cells(i, 26).Font.ColorIndex = 23
                    Else
                        wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                        wstracker.Cells(i, 26).Font.ColorIndex = 1
                    End If
            ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - Read Only" Then
                    wstracker.Cells(i, 18).Font.Underline = xlUnderlineStyleNone
                    wstracker.Cells(i, 18).Font.ColorIndex = 1
                    wstracker.Cells(i, 26).Font.Underline = xlUnderlineStyleNone
                    wstracker.Cells(i, 26).Font.ColorIndex = 1
                    wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                    wstracker.Cells(i, 2).Font.ColorIndex = 1
                    wstracker.Cells(i, 2).Font.Underline = xlUnderlineStyleNone
                    wstracker.Cells(i, 2).Font.ColorIndex = 1

            End If
Next i
MsgBox "b"
wstracker.Columns("G:G").AutoFit
'wstracker.Columns("P:P").AutoFit
'wstracker.Columns("U:U").AutoFit
'wstracker.Columns("W:W").AutoFit
wstracker.Cells(1, 1).Select
'Call aUTorefresh
wsozet2.Protect Password:="4135911"
With wstracker
    .Protect Password:="4135911", AllowFiltering:=True
    .EnableSelection = xlNoRestrictions
End With
wstracker.OLEObjects("Label1").Object.Caption = "Tamir bekleyen: " & prsay & "      " & "Onay bekleyen: " & qusay _
                                                                            & vbNewLine & "Tarama bekleyen: " & utsay & "      " & "Müh. Onayý: " & engsay
Application.Calculation = xlAutomatic
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    On Error Resume Next
    Set rs = Nothing
    Exit Sub
End Sub


Sub aUTorefresh()
     Application.OnTime Now + TimeValue("00:10:00"), "refresh"
End Sub
Function IsFileOpen(fileName As String)

Dim fileNum As Integer
Dim errNum As Integer

'Allow all errors to happen
On Error Resume Next
fileNum = FreeFile()

'Try to Açýk and close the file for inpUT.
'Errors mean the file is already Açýk
Open fileName For Input Lock Read As #fileNum
Close fileNum

'Get the error number
errNum = err

'Do not allow errors to happen
On Error GoTo 0

'Check the Error Number
Select Case errNum

    'errNum = 0 means no errors, therefore file Kapatýldý
    Case 0
    IsFileOpen = False
 
    'errNum = 70 means the file is already Açýk
    Case 70
    IsFileOpen = True

    'Something else went wrong
    Case Else
    IsFileOpen = errNum

End Select

End Function



