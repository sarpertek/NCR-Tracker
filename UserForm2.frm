VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Detaylý Rapor"
   ClientHeight    =   4776
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3732
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RemoveChar As Integer
Dim DefaultValue As Long
Private Sub CommandButton1_Click()
On Error GoTo err
ThisWorkbook.Worksheets("NCR REPORT").Unprotect Password:="4135911"
Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rs As ADODB.Recordset
Dim rsvers As ADODB.Recordset 'dim the ADO recordset class
Dim dbPath As String
Dim SQL As String
Dim i As Integer
Dim var As Boolean
Dim results(0 To 29) As Variant
Application.ScreenUpdating = False
dbPath = ThisWorkbook.Worksheets("NCR TRACKER").Range("I1").Value
Set cnn = New ADODB.Connection ' Initialise the collection class variable
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
        If ThisWorkbook.Worksheets("NCR TRACKER").Cells(3, 9) <> versiyon Then
        MsgBox "Eski bir sürüm kullanýyorsunuz baðlantý yapýlmadý!"
        Exit Sub
        End If
If OptionButton2 = True Then
            If TextBox3.Value <> "" Then
            keyword = TextBox3.Value
            Key = " LIKE *" & keyword & "*"
            'sqlkeyword = "WHERE [BLADE NUMBER] LIKE '% " & keyword & " %' "
            SQL = "SELECT DBLOG.LOGDATE, DBLOG.WHOM, NCRDB.[NCR PICTURE], DBLOG.[ORACLE NUMBER], NCRDB.PROJECT, NCRDB.[BLADE NUMBER], " & _
                  "NCRDB.[MOLD NUMBER] , NCRDB.[DEFECT CODE / TYPE], NCRDB.[Z (mm)], NCRDB.[X/C (mm)], NCRDB.[Length], NCRDB.[Width], NCRDB.[Depth], " & _
                  "NCRDB.[PS / SS], NCRDB.EDGE, NCRDB.SURFACE, " & _
                  "NCRDB.EXPLANATION, NCRDB.Source, DBLOG.Status, DBLOG.RR, " & _
                  "DBLOG.SCOP, DBLOG.SCMIN, DBLOG.LAMOP, DBLOG.LAMMIN, NCRDB.TOTALMIN, DBLOG.QRESP, NCRDB.[DISP], NCRDB.[TAMIRTURU], NCRDB.[KATSAYISI] " & _
                  "FROM NCRDB INNER JOIN DBLOG ON NCRDB.[ID] = DBLOG.[ID] ORDER BY DBLOG.LOGDATE"
            'ORDER BY DBLOG.LOGDATE;
            Set rs = New ADODB.Recordset 'assign memory to the recordset
            rs.Open SQL, cnn
            filterstring = "([BLADE NUMBER]" & _
                            Key & " OR [Status]" & _
                            Key & " OR [ORACLE NUMBER]" & _
                            Key & " OR [DEFECT CODE / TYPE]" & _
                            Key & " OR [PS / SS]" & _
                            Key & " OR [EDGE]" & _
                            Key & " OR [SURFACE]" & _
                            Key & " OR [EXPLANATION]" & _
                            Key & " OR [Source]" & _
                            Key & " OR [RR]" & _
                            Key & " OR [SCOP]" & _
                            Key & " OR [LAMOP]" & _
                            Key & " OR [QRESP]" & _
                            Key & " OR [WHOM]" & _
                            Key & ")"
                           
            rs.Filter = filterstring

                If rs.EOF And rs.BOF Then
                rs.Close
                cnn.Close
                Set rs = Nothing
                Set cnn = Nothing
                MsgBox "Herhangi bir kayýt bulunamadý!", vbCritical, "Kayýt Bulunamadý"
                Exit Sub
                End If
                                      
            Else
            MsgBox "Aramak için bir kelime giriniz"
            Exit Sub
            End If
ElseIf OptionButton3 = True Then
        If Date1.Value <> "" Then
        split1 = Split(Date1.Text, ".")
        bastarih = split1(2) & "-" & split1(1) & "-" & split1(0)
        Else
        bastarih = "01-01-2020"
        End If

        If Date2.Value <> "" Then
        split2 = Split(Date2.Text, ".")
        sontarih = split2(2) & "-" & split2(1) & "-" & split2(0) & " 23:59:00"
        Else
        sontarih = "01-01-2040"
        End If
        SQL = "SELECT DBLOG.LOGDATE, DBLOG.WHOM, NCRDB.[NCR PICTURE], DBLOG.[ORACLE NUMBER], NCRDB.PROJECT, NCRDB.[BLADE NUMBER], " & _
                  "NCRDB.[MOLD NUMBER] , NCRDB.[DEFECT CODE / TYPE], NCRDB.[Z (mm)], NCRDB.[X/C (mm)], NCRDB.[Length], NCRDB.[Width], NCRDB.[Depth], " & _
                  "NCRDB.[PS / SS], NCRDB.EDGE, NCRDB.SURFACE, " & _
                  "NCRDB.EXPLANATION, NCRDB.Source, DBLOG.Status, DBLOG.RR, " & _
                  "DBLOG.SCOP, DBLOG.SCMIN, DBLOG.LAMOP, DBLOG.LAMMIN, NCRDB.TOTALMIN, DBLOG.QRESP, NCRDB.[DISP], NCRDB.[TAMIRTURU], NCRDB.[KATSAYISI] " & _
                  "FROM NCRDB INNER JOIN DBLOG ON NCRDB.[ID] = DBLOG.[ID] WHERE " & _
                  "DBLOG.LOGDATE >= #" & bastarih & "# AND DBLOG.LOGDATE <= #" & sontarih & "#"
                  
            If TextBox1.Value <> "" Then
            var = True
            SQL1 = " AND [BLADE NUMBER] LIKE '%" & TextBox1.Value & "%' "
            Else
            SQL1 = ""
            End If
                        
            If TextBox2.Value <> "" Then
            var = True
            SQL2 = " AND NCRDB.PROJECT LIKE '%" & TextBox2.Value & "%' "
            Else
            SQL2 = ""
            End If
            
            If TextBox4.Value <> "" Then
            var = True
            SQL3 = " AND DBLOG.[ORACLE NUMBER] LIKE '%" & TextBox4.Value & "%' "
            Else
            SQL3 = ""
            End If
            
            If ComboBox1.Value <> "" Then
            var = True
            SQL4 = " AND NCRDB.[DEFECT CODE / TYPE] LIKE '%" & ComboBox1.Value & "%' "
            Else
            SQL4 = ""
            End If
            
            If TextBox5.Value <> "" Then
            var = True
            SQL5 = " AND DBLOG.STATUS LIKE '%" & TextBox5.Value & "%' "
            Else
            SQL5 = ""
            End If
            sql6 = " ORDER BY DBLOG.LOGDATE "
            If var = True Then
            SQL = SQL & SQL1 & SQL2 & SQL3 & SQL4 & SQL5 & sql6
            Else
            SQL = SQL & sql6
            End If
            'ORDER BY DBLOG.LOGDATE;
            Set rs = New ADODB.Recordset 'assign memory to the recordset
            rs.Open SQL, cnn
                
            'Check if the recordset is empty.
            If rs.EOF And rs.BOF Then
                rs.Close
                cnn.Close
                Set rs = Nothing
                Set cnn = Nothing
                MsgBox "Herhangi bir kayýt bulunamadý!", vbCritical, "Kayýt Bulunamadý"
                Exit Sub
            End If
             
Else
        If Date1.Value <> "" Then
        split1 = Split(Date1.Text, ".")
        bastarih = split1(2) & "-" & split1(1) & "-" & split1(0)
        Else
        bastarih = "01-01-2020"
        End If

        If Date2.Value <> "" Then
        split2 = Split(Date2.Text, ".")
        sontarih = split2(2) & "-" & split2(1) & "-" & split2(0) & " 23:59:00"
        Else
        sontarih = "01-01-2040"
        End If
        SQL = "SELECT NCRDB.LOGDATE, NCRDB.WHOM, NCRDB.[NCR PICTURE], NCRDB.[ORACLE NUMBER], NCRDB.PROJECT, NCRDB.[BLADE NUMBER], " & _
                  "NCRDB.[MOLD NUMBER] , NCRDB.[DEFECT CODE / TYPE], NCRDB.[Z (mm)], NCRDB.[X/C (mm)], NCRDB.[Length], NCRDB.[Width], NCRDB.[Depth], " & _
                  "NCRDB.[PS / SS], NCRDB.EDGE, NCRDB.SURFACE, " & _
                  "NCRDB.EXPLANATION, NCRDB.Source, NCRDB.Status, NCRDB.RR, " & _
                  "NCRDB.SCOP, NCRDB.SCMIN, NCRDB.LAMOP, NCRDB.LAMMIN, NCRDB.TOTALMIN, NCRDB.QRESP, NCRDB.[DISP], NCRDB.[TAMIRTURU], NCRDB.[KATSAYISI] " & _
                  "FROM NCRDB WHERE " & _
                  "NCRDB.LOGDATE >= #" & bastarih & "# AND NCRDB.LOGDATE <= #" & sontarih & "#"
            If TextBox1.Value <> "" Then
            var = True
            SQL1 = " AND [BLADE NUMBER] LIKE '%" & TextBox1.Value & "%' "
            Else
            SQL1 = ""
            End If
                        
            If TextBox2.Value <> "" Then
            var = True
            SQL2 = " AND NCRDB.PROJECT LIKE '%" & TextBox2.Value & "%' "
            Else
            SQL2 = ""
            End If
            
            If TextBox4.Value <> "" Then
            var = True
            SQL3 = " AND NCRDB.[ORACLE NUMBER] LIKE '%" & TextBox4.Value & "%' "
            Else
            SQL3 = ""
            End If
            
            If ComboBox1.Value <> "" Then
            var = True
            SQL4 = " AND NCRDB.[DEFECT CODE / TYPE] LIKE '%" & ComboBox1.Value & "%' "
            Else
            SQL4 = ""
            End If
            
            If TextBox5.Value <> "" Then
            var = True
            SQL5 = " AND NCRDB.STATUS LIKE '%" & TextBox5.Value & "%' "
            Else
            SQL5 = ""
            End If
            sql6 = " ORDER BY NCRDB.LOGDATE "
            If var = True Then
            SQL = SQL & SQL1 & SQL2 & SQL3 & SQL4 & SQL5 & sql6
            Else
            SQL = SQL & sql6
            End If
            'ORDER BY NCRDB.LOGDATE;
            Set rs = New ADODB.Recordset 'assign memory to the recordset
            rs.Open SQL, cnn

            If rs.EOF And rs.BOF Then
                rs.Close
                cnn.Close
                Set rs = Nothing
                Set cnn = Nothing
                MsgBox "Herhangi bir kayýt bulunamadý!", vbCritical, "Kayýt Bulunamadý"
                Exit Sub
            End If
             
End If

'DÝREK DOLDUR
Application.ScreenUpdating = False
ThisWorkbook.Worksheets("NCR REPORT").Unprotect Password:="4135911"
ThisWorkbook.Worksheets("NCR REPORT").Range("A6:AG50000").ClearContents
        
            With ThisWorkbook.Worksheets("NCR REPORT").ListObjects("Tablo2")
                .AutoFilter.ShowAllData
                If Not .DataBodyRange Is Nothing Then
                    .DataBodyRange.Rows.Delete
                End If
                Call .Range(2, 1).CopyFromRecordset(rs)
            End With
            LastRow = ThisWorkbook.Worksheets("NCR REPORT").Range("C" & Rows.Count).End(xlUp).Row
            For i = 6 To LastRow
            ThisWorkbook.Worksheets("NCR REPORT").Hyperlinks.Add Anchor:=ThisWorkbook.Worksheets("NCR REPORT").Cells(i, 3), Address:=ThisWorkbook.Worksheets("NCR REPORT").Cells(i, 3).Value, TextToDisplay:="Link"
            Next i
        

ThisWorkbook.Worksheets("NCR REPORT").Columns("Q:AA").AutoFit
With ThisWorkbook.Worksheets("NCR REPORT")
    .Protect Password:="4135911", AllowFiltering:=True, AllowSorting:=True
    .EnableSelection = xlNoRestrictions
End With
ActiveWindow.ScrollRow = 1
    rs.Close
    cnn.Close
    'clear memory
    Set rs = Nothing
    Set cnn = Nothing
'Call AUToSizeColumnLv(ListView1)
'Label10.Caption = "Bulunan kayýt sayýsý: " & ListView1.ListItems.Count
'With ThisWorkbook.Worksheets("NCR REPORT")
'    .Protect Password:="4135911", AllowFiltering:=True
'    .EnableSelection = xlNoRestrictions
'
'End With
If LastRow - 5 <> "" Then
ThisWorkbook.Worksheets("NCR REPORT").OLEObjects("Label1").Object.Caption = "Rapor Baþlangýcý: " & ThisWorkbook.Worksheets("NCR REPORT").Cells(6, 1).Value & "" _
                                                                            & vbNewLine & "Bulunan Kayýt: " & LastRow - 5 & "     "
Else
ThisWorkbook.Worksheets("NCR REPORT").OLEObjects("Label1").Object.Caption = " "
End If
Unload Me
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub

End Sub
Private Sub Date1_Change()
On Error GoTo err:
RemoveChar = 1
If Len(Date1.Value) = "11" Then
            Date1.Value = Left(Date1.Value, Len(Date1.Value) - RemoveChar)
End If
If Len(Date1.Value) = "1" Or Len(Date1.Value) = "4" Or Len(Date1.Value) = "7" Or Len(Date1.Value) = "8" Or Len(Date1.Value) = "9" Or Len(Date1.Value) = "10" Then
    If Not IsNumeric(Right(Date1.Value, 1)) Then
            If Len(Date1.Value) < 1 Then RemoveChar = 0
            Date1.Value = Left(Date1.Value, Len(Date1.Value) - RemoveChar)
    End If
End If
If Len(Date1.Value) = "2" Then
    If Right(Date1.Value, 2) < 32 Then
    Else
    MsgBox "Days can not be higher than 31!"
    End If
End If
If Len(Date1.Value) = "5" Then
    If Right(Date1.Value, 2) < 13 Then
    Else
    MsgBox "Months can not be higher than 12!"
    End If
End If
If Len(Date1.Value) = "3" Or Len(Date1.Value) = "6" Then
        If Not Right(Date1.Value, 1) = "." Then
            If Len(Date1.Value) < 1 Then RemoveChar = 0
            Date1.Value = Left(Date1.Value, Len(Date1.Value) - RemoveChar)
        End If
End If
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub

Private Sub Date2_Change()
On Error GoTo err
RemoveChar = 1
If Len(Date2.Value) = "11" Then
            Date2.Value = Left(Date2.Value, Len(Date2.Value) - RemoveChar)
End If
If Len(Date2.Value) = "1" Or Len(Date2.Value) = "4" Or Len(Date2.Value) = "7" Or Len(Date2.Value) = "8" Or Len(Date2.Value) = "9" Or Len(Date2.Value) = "10" Then
    If Not IsNumeric(Right(Date2.Value, 1)) Then
            If Len(Date2.Value) < 1 Then RemoveChar = 0
            Date2.Value = Left(Date2.Value, Len(Date2.Value) - RemoveChar)
    End If
End If
If Len(Date2.Value) = "2" Then
    If Right(Date2.Value, 2) < 32 Then
    Else
    MsgBox "Days can not be higher than 31!"
    End If
End If
If Len(Date2.Value) = "5" Then
    If Right(Date2.Value, 2) < 13 Then
    Else
    MsgBox "Months can not be higher than 12!"
    End If
End If
If Len(Date2.Value) = "3" Or Len(Date2.Value) = "6" Then
        If Not Right(Date2.Value, 1) = "." Then
            If Len(Date2.Value) < 1 Then RemoveChar = 0
            Date2.Value = Left(Date2.Value, Len(Date2.Value) - RemoveChar)
        End If
End If
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub

Private Sub Optionbutton1_Change()
    If OptionButton1 = True Or OptionButton3 = True Then
    Date1.Enabled = True
    Date2.Enabled = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
    ComboBox1.Enabled = True
    TextBox3.Enabled = False
    TextBox4.Enabled = True
    TextBox5.Enabled = True
    ElseIf OptionButton2 = True Then
    Date1.Enabled = False
    Date2.Enabled = False
    TextBox1.Enabled = False
    TextBox2.Enabled = False
    ComboBox1.Enabled = False
    TextBox3.Enabled = True
    TextBox4.Enabled = False
    TextBox5.Enabled = False
    End If
End Sub

Private Sub OptionButton3_Change()
    If OptionButton1 = True Or OptionButton3 = True Then
    Date1.Enabled = True
    Date2.Enabled = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
    ComboBox1.Enabled = True
    TextBox3.Enabled = False
    TextBox4.Enabled = True
    TextBox5.Enabled = True
    ElseIf OptionButton2 = True Then
    Date1.Enabled = False
    Date2.Enabled = False
    TextBox1.Enabled = False
    TextBox2.Enabled = False
    ComboBox1.Enabled = False
    TextBox3.Enabled = True
    TextBox4.Enabled = False
    TextBox5.Enabled = False
    End If
End Sub
Private Sub UserForm_Initialize()
    OptionButton1 = True
    ComboBox1.List = Array( _
            "D010 - Dry Glass", "D020 - Semi-Dry", "D050 - Entrained Air", "D055 - Damaged Glass", "D060 - FOD", "D070 - Crack", "D072 - Delamination", _
            "D074 - Rich Resin", "D080 - Paste Void - Exposed", "D082 - Paste Void - Non-Exposed", "D090 - Resin Voids", "D100 - Paint Rougness and Thickness", _
            "D101 - LPS Damage", "D102 - LPS Continuity", "D103 - Shiny Spots", "D106 - Folded Laminate", "D107 - LE Bond Cap Width Deviation", _
            "D109 - Glass Location", "D111 - Core Gap", "D112 - Missing Core", "D114 - Core Location", "D150 - Wave", "D160 - Overbite-Underbite", _
            "D170 - TE Thickness", "D180 - Root Face Flatness", "D220 - Bond Line Thickness", "D310 - Incompletely Cured ", _
            "D320 - Blade-Subcomponent Geometry", "D330 - Subcomponent Position", "D350 - Mass or Moment", "D370 - Coatings", _
            "D401 - Disbonding", "D405 - LPS-SPL Position", "D500 - Other", "D505 - Glue Spill", "D510 - Blocked Drainage Hole", _
            "D625 - Burned Area", "D635 - Core Step", "D640 - Edge Step", "D640 - Edge Step", _
            "D655 - Hole in Laminate", "D675 - Short Flange", "D680 - Bond Width", "D700 - Too much tackifier ", _
            "D705 - Bond Cap Deviation", "D710 - SMT Shift", "D715 - Carbon Fleece Shift ", "D720 - SWR Protrusion ", _
            "D725 - Glass Overlap", "D730 - Receptor Bolt Length", "D735 - Pultrusion Length", "D740 - Pultrusion Chamfer", "D750 - Pultrusion Other", _
            "D755 - Grooves ", "D775 - Red Zone-Wrinkle", "D776 - Green Zone - Wrinkle", "D777 - Yellow Zone - Wrinkle", "D778 - White Zone - Wrinkle ")
    TextBox2.List = Array("V162", "NX74.5", "V136")
    Date1.Enabled = True
    Date2.Enabled = True
    TextBox1.Enabled = True
    TextBox2.Enabled = True
    ComboBox1.Enabled = True
    TextBox3.Enabled = False
    TextBox4.Enabled = True
    TextBox5.Enabled = True
End Sub

Private Sub AUToSizeColumnLv(lv As ListView)

Dim i      As Integer
Dim szKey   As String
Dim objItem As ListItem

    With lv
        szKey = CStr(Now)

        Set objItem = .ListItems.Add(1, szKey, .ColumnHeaders(1).Text & "  ")
        SendMessage .hWnd, LVM_SETCOLUMNWIDTH, 0, LVSCW_AUTOSIZE

        For i = 1 To .ColumnHeaders.Count - 1
            objItem.SubItems(i) = .ColumnHeaders(i + 1).Text & "  "
            SendMessage .hWnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE
        Next i

        .ListItems.Remove szKey
    End With

End Sub
