VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "NCR Olu�tur"
   ClientHeight    =   7356
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8136
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim patharray(50) As String
Dim fname() As String
Dim patharray2(50) As String
Dim fname2() As String
Dim asamasay As Integer
Dim UTsayi, durum As String

Private Sub ComboBox8_Change()
ComboBox5.Enabled = True
If ComboBox8.Value = "V136" Then
ComboBox7.List = Array("KANAT", "SHELL", "MSW", "TSWROOT", "TSWTIP", "TEINSEND", "TEINSMID", "RF")
ComboBox2.List = Array("LE", "TE", "LE-TE")
ComboBox7.Value = ""
ComboBox2.Value = ""
    If hop2 = False Then
    ComboBox2.Value = ""
    ComboBox5.List = Array("D010-Dry Glass-Kuru Kuma�", "D020-Semi Dry Glass-Yar� Kuru Kuma�", _
    "D050-Entrained Air-Hava Alm�� B�lge", "D055-Damaged Glass-Laminasyonda Hasar", _
    "D060-Foreign Objects-Yabanc� Madde", "D070-Cracks-�atlak", "D072 Delamination-Delaminasyon", _
    "D074 Resin Rich-Zengin Re�ine", "D080-Paste Voids - Exposed-Yetersiz Krema", _
    "D082 Paste Voids - Non-Exposed-Kremada Hava", "D090-Resin Voids-Re�inede Hava", _
    "D101 Lps Damage-Lps Hasar�", "D102 Lps Continuity-Lps Ba�lant� - Diren� Hatas�", _
    "D103 Shiny Spots-Parlak Noktalar", "D104 Loose Laminates-Kalk�k Kuma� Lifleri", _
    "D106 Folded Laminates-Kuma� Katlanmas�", "D109 Glass Or Fabric Location-Kuma� Konum Hatas�", _
    "D110-Core Discoloration-N�ve Yan���", "D111 Core Gap-N�ve Bo�lu�u", _
    "D112 Missing Core-Eksik N�ve", "D114 Core Location-N�ve Konum Hatas�-Kayma", _
    "D150-Waves & Wrinkles-Dalga", "D160-Overbite-Underbite-Ps-Ss Kabu Ka��kl�k", _
    "D170-Trailing Edge Thickness-Te Kal�nl�k", "D220-Bond Line Thickness-Yap��ma Hatt�-Krema Kal�nl��� ", _
    "D310-insufficient Cure Or Failed Tg-Yetersiz K�rlenme (Tg)", "D320-Blade-Subcomponent Geometry-Kanat-Par�a Geometrisi", _
    "D330-Subcomponent Position-Par�a Pozisyon Hatas�-Hizas�zl���", _
    "D350-Mass Or Moment-A��rl�k Veya Moment", "D370-Coatings-Y�zey Kaplamalar�", _
    "D401 Disbond-Yap��mama", "D405-Lps Position-Lps Pozisyonu", _
    "D505-Excessive Adhesive-Krema Ta�mas�", "D510-Blocked Drainage Hole-T�kal� Drenaj Deli�i", _
    "D635-Core Step-N�ve Y�kseklik Fark�", "D640-Edge Step-N�ve Kenar Y�kseklik Fark�", _
    "D655-Holes in Laminate-Laminasyonda Delik", "D670-Flange Step-Flan� Ge�i� Y�kseklik Fark�", _
    "D675-Short Flange-K�sa Flan�", "D680-Bond Width-Yap��ma Geni�li�i", _
    "D700-Excessive Tackifier-Fazla Sprey Kullan�m�", "D705-Bond Cap Deviation-Yap��ma Kebi Hatas�", _
    "D710-Lps Tip Position-Smt Kaymas�", "D720-Steel insert Protrusion-Y�zeyde Insert ��k�nt�s�", _
    "D725-Glass Overlap-Kuma� Bindirmesi", "D740-Pultrusion Chamfer-Pultruzyon Pah Hatas�", _
    "D755-Groove-Oluk", "D760-Handling Damage-Ta��ma Hasar�", "D770-Pultrusion Void-Pultruzyonda Hava", "D924 Laminate Discoloration-Laminasyonda Renk Hatas�", _
    "D970-Exceeding Surface Preparation Time-Y�zey Aktivasyon S�re A��m�")
    ComboBox5.Value = ""
    End If
ElseIf ComboBox8.Value = "V162" Then
ComboBox7.List = Array("KANAT", "SHELL", "MSW", "TSW", "TEINS", "RF")
ComboBox2.List = Array("LE", "TE", "LE-TE")
ComboBox7.Value = ""
ComboBox2.Value = ""
    If hop2 = False Then
    ComboBox2.Value = ""
    ComboBox5.List = Array("D010-Dry Glass-Kuru Kuma�", "D020-Semi Dry Glass-Yar� Kuru Kuma�", _
    "D050-Entrained Air-Hava Alm�� B�lge", "D055-Damaged Glass-Laminasyonda Hasar", _
    "D060-Foreign Objects-Yabanc� Madde", "D070-Cracks-�atlak", "D072 Delamination-Delaminasyon", _
    "D074 Resin Rich-Zengin Re�ine", "D080-Paste Voids - Exposed-Yetersiz Krema", _
    "D082 Paste Voids - Non-Exposed-Kremada Hava", "D090-Resin Voids-Re�inede Hava", _
    "D101 Lps Damage-Lps Hasar�", "D102 Lps Continuity-Lps Ba�lant� - Diren� Hatas�", _
    "D103 Shiny Spots-Parlak Noktalar", "D104 Loose Laminates-Kalk�k Kuma� Lifleri", _
    "D106 Folded Laminates-Kuma� Katlanmas�", "D109 Glass Or Fabric Location-Kuma� Konum Hatas�", _
    "D110-Core Discoloration-N�ve Yan���", "D111 Core Gap-N�ve Bo�lu�u", _
    "D112 Missing Core-Eksik N�ve", "D114 Core Location-N�ve Konum Hatas�-Kayma", _
    "D150-Waves & Wrinkles-Dalga", "D160-Overbite-Underbite-Ps-Ss Kabu Ka��kl�k", _
    "D170-Trailing Edge Thickness-Te Kal�nl�k", "D220-Bond Line Thickness-Yap��ma Hatt�-Krema Kal�nl��� ", _
    "D310-insufficient Cure Or Failed Tg-Yetersiz K�rlenme (Tg)", "D320-Blade-Subcomponent Geometry-Kanat-Par�a Geometrisi", _
    "D330-Subcomponent Position-Par�a Pozisyon Hatas�-Hizas�zl���", _
    "D350-Mass Or Moment-A��rl�k Veya Moment", "D370-Coatings-Y�zey Kaplamalar�", _
    "D401 Disbond-Yap��mama", "D405-Lps Position-Lps Pozisyonu", _
    "D505-Excessive Adhesive-Krema Ta�mas�", "D510-Blocked Drainage Hole-T�kal� Drenaj Deli�i", _
    "D635-Core Step-N�ve Y�kseklik Fark�", "D640-Edge Step-N�ve Kenar Y�kseklik Fark�", _
    "D655-Holes in Laminate-Laminasyonda Delik", "D670-Flange Step-Flan� Ge�i� Y�kseklik Fark�", _
    "D675-Short Flange-K�sa Flan�", "D680-Bond Width-Yap��ma Geni�li�i", _
    "D700-Excessive Tackifier-Fazla Sprey Kullan�m�", "D705-Bond Cap Deviation-Yap��ma Kebi Hatas�", _
    "D710-Lps Tip Position-Smt Kaymas�", "D720-Steel insert Protrusion-Y�zeyde Insert ��k�nt�s�", _
    "D725-Glass Overlap-Kuma� Bindirmesi", "D740-Pultrusion Chamfer-Pultruzyon Pah Hatas�", _
    "D755-Groove-Oluk", "D760-Handling Damage-Ta��ma Hasar�", "D770-Pultrusion Void-Pultruzyonda Hava", "D924 Laminate Discoloration-Laminasyonda Renk Hatas�", _
    "D970-Exceeding Surface Preparation Time-Y�zey Aktivasyon S�re A��m�")
    ComboBox5.Value = ""
    End If
Else
ComboBox7.List = Array("KANAT", "SHELL", "MG", "MSW", "TSW")
ComboBox2.List = Array("LE", "TE", "LE-TE", "P1", "P2", "P3", "P4", "P5", "P6")
ComboBox7.Value = ""
ComboBox2.Value = ""
    If hop2 = False Then
    ComboBox5.List = Array( _
                "D010-Dry Glass", "D020-Semi-Dry", "D050-Entrained Air", "D055-Damaged Glass", "D060-FOD", "D070-Crack", "D072-Delamination", _
                "D074-Rich Resin", "D080-Paste Void-Exposed", "D082-Paste Void-Non-Exposed", "D090-Resin Voids", "D100-Paint Rougness and Thickness", _
                "D101-LPS Damage", "D102-LPS Continuity", "D103-Shiny Spots", "D106-Folded Laminate", "D107-LE Bond Cap Width Deviation", _
                "D109-Glass Location", "D111-Core Gap", "D112-Missing Core", "D114-Core Location", "D150-Wave", "D160-Overbite-Underbite", _
                "D170-TE Thickness", "D180-Root Face Flatness", "D220-Bond Line Thickness", "D310-Incompletely Cured ", _
                "D320-Blade-Subcomponent Geometry", "D330-Subcomponent Position", "D350-Mass or Moment", "D370-Coatings", _
                "D401-Disbonding", "D405-LPS-SPL Position", "D500-Other", "D505-Glue Spill", "D510-Blocked Drainage Hole", _
                "D625-Burned Area", "D635-Core Step", "D640-Edge Step", "D640-Edge Step", _
                "D655-Holes in Laminate/Laminasyonda Delik", "D675-Short Flange", "D680-Bond Width", "D700-Too much tackifier ", _
                "D705-Bond Cap Deviation", "D710-SMT Shift", "D715-Carbon Fleece Shift ", "D720-SWR Protrusion ", _
                "D725-Glass Overlap", "D730-Receptor Bolt Length", "D735-Pultrusion Length", "D740-Pultrusion Chamfer", "D750-Pultrusion Other", _
                "D755-Grooves ", "D775-Red Zone-Wrinkle", "D776-Green Zone-Wrinkle", "D777-Yellow Zone-Wrinkle", "D778-White Zone-Wrinkle ")
    ComboBox5.Value = ""
    End If
End If
End Sub

Private Sub CommandButton1_Click()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
'On Error GoTo err
Application.ScreenUpdating = False
Application.Calculation = xlManual

wstracker.Unprotect Password:="4135911"

UserName = Application.UserName

Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset
Dim rsvers As ADODB.Recordset 'dim the ADO recordset class
Dim rsorano As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim dbPath
Dim x As Long, i As Long
Dim fso6 As New FileSystemObject
Dim fso5 As New FileSystemObject
Dim fso4 As New FileSystemObject
Dim fso3 As New FileSystemObject
Dim fso2 As New FileSystemObject
Dim fso As New FileSystemObject
Dim link, opisim, opmin, sqlvers As String
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
            MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
            Exit Sub
    End If
        versiyon = rsvers.Fields(0).Value
        If wstracker.Cells(3, 9) <> versiyon Then
        MsgBox "Eski bir s�r�m kullan�yorsunuz ba�lant� yap�lmad�!"
        Exit Sub
        End If

'Erro handler
'On Error GoTo errHandler:
'Foto kopyala
folderpath = wstracker.Cells(2, 9)

If TextBox9.Enabled = True Then
    If ComboBox5.Value = "D500 - Other" Or ComboBox5.Value = "D060 - FOD" Or ComboBox5.Value = "D655 - Hole in Laminate" Then
        If TextBox9.Value = "" Then
        MsgBox "A��klama k�sm� doldurulmal�d�r.", vbCritical
        Exit Sub
        End If
    End If
End If

'If Me.TreeView1.Nodes.Count = 0 And UserForm1.Caption <> "�ki a�amal� tamir" And UserForm1.Caption <> "M�hendislik Onay�" Then
If Me.TreeView1.Nodes.Count = 0 And Me.TreeView1.Visible = True Then
    MsgBox "Devam etmek i�in foto�raflar� ekleyiniz"
    Exit Sub
End If
If Me.TreeView2.Nodes.Count = 0 And Me.TreeView2.Visible = True Then
    MsgBox "Devam etmek i�in foto�raflar� ekleyiniz"
    Exit Sub
End If

 If Len(TextBox2.Value) < 4 Then
    var = TextBox2.Value
    For i = 1 To 4 - Len(var)
    var = "0" & var
    Next i
 Else
    var = TextBox2.Value
 End If
 
var = ComboBox7.Value & "-" & var
var2 = TextBox10.Value

'klas�r i�in yeni id al

If UserForm1.Caption = "Son Kontroller ve After Onay�" Then
    If TextBox1.Value = "000000" Then
    MsgBox "Oracle numaras� al�nmadan NCR kapat�lamaz"
    Exit Sub
    Else
    SQL = "UPDATE NCRDB SET STATUS = 'M�h. Onay�', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
    SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
    Set cnn = New ADODB.Connection
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
    SQL0 = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & ""
    Set rs = New ADODB.Recordset
    rs.Open SQL0, cnn
    If rs.EOF And rs.BOF Then
            rs.Close
            cnn.Close
            Set rs = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
            Exit Sub
    End If
        link = rs.Fields(0).Value
        rs.Close
'link okay
        finalpath = link & "5 - Tamir Sonras�\"
        If Not fso6.FolderExists(finalpath) Then
               fso6.CreateFolder finalpath
        End If
        'fotoyu da y�kle
        For Each picturepath In patharray()
            If Not picturepath = "" Then
            fname = Split(picturepath, "\")
            Result = fname(UBound(fname))
                NewFile = finalpath & Result
                FileCopy picturepath, NewFile
            End If
        Next
        
        For Each picturepath2 In patharray2()
            If Not picturepath2 = "" Then
            fname2 = Split(picturepath2, "\")
            Result2 = fname2(UBound(fname2))
        finalpath = link & "1 - NCR �n ve arka y�z�\"
                            If Not fso6.FolderExists(finalpath) Then
                                  fso6.CreateFolder finalpath
                            End If
            NewFile2 = finalpath & Result2
            FileCopy picturepath2, NewFile2
            End If
        Next

    cnn.Execute SQL
    cnn.Execute SQL2
    End If
ElseIf UserForm1.Caption = "UT Onay�" Then
    If TextBox1.Value = "000000" Then
    MsgBox "Oracle numaras� al�nmadan NCR kapat�lamaz"
    Exit Sub
    Else
    SQL = "UPDATE NCRDB SET STATUS = 'M�h. Onay�', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
    SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
    Set cnn = New ADODB.Connection
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
    SQL0 = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & ""
    Set rs = New ADODB.Recordset
    rs.Open SQL0, cnn
    If rs.EOF And rs.BOF Then
            rs.Close
            cnn.Close
            Set rs = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
            Exit Sub
    End If
        link = rs.Fields(0).Value
        rs.Close
'link okay
        finalpath = link & "1 - NCR �n ve arka y�z�\"
        If Not fso6.FolderExists(finalpath) Then
               fso6.CreateFolder finalpath
        End If
        'fotoyu da y�kle
        For Each picturepath In patharray()
            If Not picturepath = "" Then
            fname = Split(picturepath, "\")
            Result = fname(UBound(fname))
                NewFile = finalpath & Result
                FileCopy picturepath, NewFile
            End If
        Next
    cnn.Execute SQL
    cnn.Execute SQL2
    End If
ElseIf UserForm1.Caption = "NCR Olu�tur" Or UserForm1.Caption = "UT NCR'� Olu�tur" Then
        If TextBox2.Value = "" Or ComboBox1.Value = "" Or ComboBox2.Value = "" _
        Or ComboBox3.Value = "" Or ComboBox4.Value = "" _
        Or ComboBox6.Value = "" Or ComboBox7.Value = "" Or ComboBox7.Value = "" Then
        MsgBox "T�m alanlar doldurulmal�d�r.", vbCritical, "ZORUNLU ALANLAR"
        Exit Sub
        End If
                If Len(TextBox1.Value) = 6 Or TextBox1.Value = "" Or TextBox1.Value = "000000" Then
                Else
                MsgBox "NCR No hatal� veya eksik girildi"
                Exit Sub
                End If
        
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        SQL = "SELECT max(ID) FROM NCRDB"
        Set rs = New ADODB.Recordset
        rs.Open SQL, cnn
        If rs.EOF And rs.BOF Then
            rs.Close
            cnn.Close
            Set rs = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
            Exit Sub
        End If
        
        '        'yapay zeka kasal�m :)
        Dim alan, tahmin As Integer
        alan = length.Value * En.Value
        Dim sqlhesap As String
        SQL = "SELECT TOP 5 * " & _
                   "FROM NCRDB " & _
                   "WHERE NCRDB.[DEFECT CODE / TYPE] = '" & ComboBox5.Value & "' And NCRDB.SURFACE = '" & ComboBox3.Value & "' And NCRDB.TOTALMIN > 0" & _
                   " ORDER BY Abs(" & alan & "-([NCRDB].[LENGTH]*[NCRDB].[Width]))"
        sqlhesap = "SELECT AVG(TOTALMIN) AS TheAverage FROM (" & SQL & ")"
        
'        'BURDAN SONRASI S�L�NECEK
'        Set rsvers2 = New ADODB.Recordset
'        rsvers2.Open SQL, cnn
'        ThisWorkbook.Worksheets("Sayfa1").Range("A2").CopyFromRecordset rsvers2
'        rsvers2.Close
'
'
        Set rsvers = New ADODB.Recordset
        rsvers.Open sqlhesap, cnn
            If rsvers.EOF And rsvers.BOF Then
            tahmin = 0
            Else
                If IsNull(rsvers(0).Value) Then tahmin = 0 Else tahmin = rsvers(0).Value
            End If
        rsvers.Close
       'yapay zeka son
        
        Dim newid As Integer
        newid = rs.Fields(0).Value + 1
            rs.Close
            cnn.Close
            Set rs = Nothing
            Set cnn = Nothing
        
        For Each picturepath In patharray()
                    If Not picturepath = "" Then
                    fname = Split(picturepath, "\")
                    Result = fname(UBound(fname))
                    'KLAS�RLEME
                    'ana klas�r
                    newfolder1 = folderpath & "\" & ComboBox8.Value & "\"
                    'kanatsa veya shellse 2. klas�r
                    If ComboBox7.Value = "KANAT" Or ComboBox7.Value = "SHELL" Then
                    newpath = newfolder1 & "KANAT\"
                    newpath2 = newpath & var & "\"
                        'kanat m� shell mi
                        If ComboBox7.Value = "KANAT" Then
                        newpath3 = newpath2 & "TPT\"
                        Else
                        'shellse
                        newpath3 = newpath2 & "SHELL\"
                        End If
                        'oracle no yoksa
                                If TextBox1.Value = "" Then
                                        newpath4 = newpath3 & newid & " - " & ComboBox5.Value & "\"
                        'oracle no varsa
                                Else
                                        newpath4 = newpath3 & yenino & " - " & ComboBox5.Value & "\"
                                End If
                    Else
                    'k���k par�aysa
                    newpath = newfolder1 & "SMALL PARTS\"
                    newpath2 = newpath & ComboBox7.Value & "\"
                    newpath3 = newpath2 & var & "\"
                        'oracle no yoksa BURDA KALDIK
                                If TextBox1.Value = "" Then
                                        newpath4 = newpath3 & newid & " - " & ComboBox5.Value & "\"
                        'oracle no varsa
                                Else
                                        newpath4 = newpath3 & yenino & " - " & ComboBox5.Value & "\"
                                End If
                    End If
                    If UserForm1.Caption = "UT NCR'� Olu�tur" Then
                    finalpath = newpath4 & "2 - Tamir �ncesi\"
                    'else kalite a��yorsa ncr'�
                    Else
                    finalpath = newpath4 & "1 - NCR �n ve arka y�z�\"
                    End If
                                    If Not fso.FolderExists(newfolder1) Then
                                          fso.CreateFolder newfolder1
                                    End If
                                    If Not fso2.FolderExists(newpath) Then
                                          fso2.CreateFolder newpath
                                    End If
                                    If Not fso3.FolderExists(newpath2) Then
                                          fso3.CreateFolder newpath2
                                    End If
                                    If Not fso4.FolderExists(newpath3) Then
                                          fso4.CreateFolder newpath3
                                    End If
                                    If Not fso5.FolderExists(newpath4) Then
                                          fso5.CreateFolder newpath4
                                    End If
                                    If Not fso6.FolderExists(finalpath) Then
                                          fso6.CreateFolder finalpath
                                    End If
                    NewFile = finalpath & Result
                    FileCopy picturepath, NewFile
                    End If
        Next
        
        For Each picturepath2 In patharray2()
                    If Not picturepath2 = "" Then
                    fname2 = Split(picturepath2, "\")
                    Result2 = fname2(UBound(fname2))
                    finalpath = newpath4 & "2 - Tamir �ncesi\"
                                    If Not fso6.FolderExists(finalpath) Then
                                          fso6.CreateFolder finalpath
                                    End If
                    NewFile2 = finalpath & Result2
                    FileCopy picturepath2, NewFile2
                    End If
        Next

        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        Set rst = New ADODB.Recordset
        rst.Open Source:="NCRDB", ActiveConnection:=cnn, _
        CursorType:=adOpenDynamic, LockType:=adLockOptimistic, _
        Options:=adCmdTable
        With rst
        .AddNew
        If TextBox1.Value = "" Then
        .Fields("ORACLE NUMBER").Value = "000000"
        Else
        .Fields("ORACLE NUMBER").Value = yenino
        End If
        .Fields("PROJECT").Value = ComboBox8.Value
        .Fields("BLADE NUMBER").Value = var
        .Fields("MOLD NUMBER").Value = ComboBox6.Value
        .Fields("OPEN DATE").Value = Now
        .Fields("DEFECT CODE / TYPE").Value = ComboBox5.Value
        .Fields("PS / SS").Value = ComboBox1.Value
        .Fields("EDGE").Value = ComboBox2.Value
        .Fields("SURFACE").Value = ComboBox3.Value
        If ztext.Value = "" Then ztext.Value = 0
        If xtext.Value = "" Then xtext.Value = 0
        If length.Value = "" Then length.Value = 0
        If En.Value = "" Then En.Value = 0
        If Depth.Value = "" Then Depth.Value = 0
        .Fields("Z (mm)").Value = ztext.Value
        .Fields("X/C (mm)").Value = xtext.Value
        .Fields("Length").Value = length.Value
        .Fields("Width").Value = En.Value
        .Fields("Depth").Value = Depth.Value
        .Fields("EXPLANATION").Value = TextBox9.Value
        .Fields("SOURCE").Value = ComboBox4.Value
        .Fields("STATUS").Value = "A��k"
        .Fields("NCR PICTURE").Value = newpath4
        .Fields("WHOM").Value = UserName
        .Fields("LOGDATE").Value = Now
        If wstracker.Cells(4, 1) = "NCR MANAGEMENT - FINISH QUALITY" Then
        .Fields("QRESP").Value = "Fin Q"
        ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" And ComboBox5.Value = "D655-Holes in Laminate/Laminasyonda Delik" Then
        .Fields("QRESP").Value = "Fin Q"
        ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" Then
        .Fields("QRESP").Value = "UT"
        Else
        .Fields("QRESP").Value = "Mold Q"
        End If
        .Fields("TAHMINI SURE").Value = tahmin
        .Update
        End With
        SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = (select * from (SELECT max(ID) from NCRDB) as t) "
        cnn.Execute SQL2
        MsgBox " NCR Ba�ar�yla olu�turuldu ve foto�raflar ar�ivlendi."
        If wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" And ComboBox5.Value = "D655 - Hole in Laminate" Then
        MsgBox " Delikler aras� mesafe NCR'� Finish Kaliteye Aktar�ld�, ka��d� onlara teslim edin"
        End If
ElseIf UserForm1.Caption = "Tek a�amal� tamir" Then
        If opminbox.Value = "" Or opadi.Value = "" Then
        MsgBox "Onaya g�ndermek i�in gerekli alanlar� doldurun."
        Exit Sub
        End If
'production tek a�amal� tamir upload
        'Link yakala
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        SQL = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & ""
        Set rs = New ADODB.Recordset
        rs.Open SQL, cnn
        If rs.EOF And rs.BOF Then
            rs.Close
            cnn.Close
            Set rs = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
            Exit Sub
        End If
        link = rs.Fields(0).Value
        rs.Close
'link okay
        opisim = opadi.Value
        opmin = opminbox.Value
            If wstracker.Cells(idrow, 26).Value <> "UT" Then
            SQL = "UPDATE NCRDB SET STATUS = '1. A�ama Onay�', RR = ' ', SCOP = '" & opisim & "', SCMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
            Else
                If UTsayi = "A��k" Then
                durum = "UT Onay� 0"
                durum = Right(durum, 1) + 1
                durum = "UT Onay� " & durum
                ElseIf UTsayi Like "UT M*" Then
                durum = Right(UTsayi, 1)
                durum = "UT Onay� " & durum
                ElseIf durum = "" Then durum = "UT Onay� 1"
                End If
            SQL = "UPDATE NCRDB SET STATUS = '" & durum & "', RR = ' ', SCOP = '" & opisim & "', SCMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
            End If
            finalpath = link & "3 - 1. A�ama Tamir\"
            If Not fso6.FolderExists(finalpath) Then
                  fso6.CreateFolder finalpath
            End If
        'fotoyu da y�kle
        For Each picturepath In patharray()
            If Not picturepath = "" Then
            fname = Split(picturepath, "\")
            Result = fname(UBound(fname))
                NewFile = finalpath & Result
                FileCopy picturepath, NewFile
            End If
        Next
            cnn.Execute SQL
            SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
            cnn.Execute SQL2

ElseIf UserForm1.Caption = "1. A�ama Onay�" Or UserForm1.Caption = "2. A�ama Onay�" Then
        'link yakala
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        SQL = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & ""
        Set rs = New ADODB.Recordset
        rs.Open SQL, cnn
        If rs.EOF And rs.BOF Then
            rs.Close
            cnn.Close
            Set rs = Nothing
            Set cnn = Nothing
            MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
            Exit Sub
        End If
        link = rs.Fields(0).Value
        rs.Close
        'link okay
    If UserForm1.Caption = "1. A�ama Onay�" Then
        finalpath = link & "3 - 1. A�ama Tamir\"
            'tamir e�er tek a�amal�ysa
            If wstracker.Cells(idrow, 27).Value = "1 - Laminasyon UYGULANMAYACAK" Or wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then
            SQL = "UPDATE NCRDB SET STATUS = 'After Kontrol', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
            'tamir e�er �ok a�amal�ysa
            Else
            SQL = "UPDATE NCRDB SET STATUS = '2. A�ama', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
            End If
    ElseIf UserForm1.Caption = "2. A�ama Onay�" Then
        SQL = "UPDATE NCRDB SET STATUS = 'After Kontrol', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        finalpath = link & "4 - 2. A�ama Tamir\"
    Else
    End If

    If Not fso6.FolderExists(finalpath) Then
          fso6.CreateFolder finalpath
    End If
    
'fotoyu da y�kle
        For Each picturepath In patharray()
            If Not picturepath = "" Then
            fname = Split(picturepath, "\")
            Result = fname(UBound(fname))
                NewFile = finalpath & Result
                FileCopy picturepath, NewFile
          End If
        Next
        cnn.Execute SQL
        SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
        cnn.Execute SQL2
        

ElseIf UserForm1.Caption = "�ki a�amal� tamir" Then
        If opminbox.Value = "" Or opadi.Value = "" Then
        MsgBox "Onaya g�ndermek i�in gerekli alanlar� doldurun."
        Exit Sub
        End If
'production iki a�amal� tamir upload
        opisim = opadi.Value
        opmin = opminbox.Value
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        'daha 1. a�amadaysa
        If wstracker.Cells(idrow, 18).Value = "A��k" Or wstracker.Cells(idrow, 18).Value = "1. A�ama NOK" Then
            SQL = "UPDATE NCRDB SET STATUS = '1. A�ama Onay�', RR = ' ', SCOP = '" & opisim & "', SCMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
        '2. a�amadaysa
        ElseIf wstracker.Cells(idrow, 18).Value = "2. A�ama NOK" Or wstracker.Cells(idrow, 18).Value = "2. A�ama" Then
            SQL = "UPDATE NCRDB SET STATUS = '2. A�ama Onay�', RR = ' ', LAMOP = '" & opisim & "', LAMMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
            '2. a�amadaysa foto�raf y�kle
                If wstracker.Cells(idrow, 18).Value = "2. A�ama" Then
                'foto�raf y�klenecek linki bul
                    SQL0 = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & ""
                    Set rs = New ADODB.Recordset
                    rs.Open SQL0, cnn
                        If rs.EOF And rs.BOF Then
                                rs.Close
                                cnn.Close
                                Set rs = Nothing
                                Set cnn = Nothing
                                MsgBox "Herhangi bir kay�t bulunamad� server ba�lant�n�z� kontrol edin", vbCritical, "Kay�t Bulunamad�"
                                Exit Sub
                        End If
                        link = rs.Fields(0).Value
                        rs.Close
                            'link okay �imdi fotoyu y�kle
                                finalpath = link & "1 - NCR �n ve arka y�z�\"
                        If Not fso6.FolderExists(finalpath) Then
                               fso6.CreateFolder finalpath
                        End If
                        For Each picturepath In patharray()
                            If Not picturepath = "" Then
                            fname = Split(picturepath, "\")
                            Result = fname(UBound(fname))
                                NewFile = finalpath & Result
                                FileCopy picturepath, NewFile
                            End If
                        Next
                End If
        End If
            cnn.Execute SQL
            SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
            cnn.Execute SQL2

ElseIf UserForm1.Caption = "M�hendislik Onay�" Then
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        'onayla
        SQL = "UPDATE NCRDB SET STATUS = 'Kapat�ld�', RR = ' ', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
        cnn.Execute SQL
        SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
        cnn.Execute SQL2

't�m se�imlerin sonu end ifi end if gibi end if
End If

On Error Resume Next
rs.Close
rst.Close
rsorano.Close
cnn.Close
Set rsorano = Nothing
Set rs = Nothing
Set rst = Nothing
Set cnn = Nothing
On Error GoTo 0

Unload Me
Call refresh
Exit Sub
err:
    MsgBox "�nternete ba�l� oldu�unu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran g�r�nt�s� at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "S�k�nt� var :("
    Exit Sub
End Sub

Private Sub CommandButton2_Click()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
'On Error GoTo err
Application.ScreenUpdating = False
Application.Calculation = xlManual
Dim resp As String
Dim fso6 As New FileSystemObject

    If rrcombo.Value = "" Then
    MsgBox "Reddetmek i�in sebep girilmedi!"
    Exit Sub
    ElseIf rrtext.Visible = True And rrtext.Value = "" Then
    MsgBox "A��klama girilmedi!"
    Exit Sub
    ElseIf rrcombo.Value = "Tamir tekrar�" Then
        If TextBox1.Value = "000000" Then
        MsgBox "Tamir tekrar� vermek i�in bu NCR i�in Oracle NO girilmelidir!"
        Exit Sub
        ElseIf rrtext.Value = "" Then
        MsgBox "Tamir tekrar� i�in a��klama giriniz"
        Exit Sub
        End If
    
    rreason = "Tamir tekrar� - " & rrtext.Text
    ElseIf rrcombo.Value = "Di�er" Then
    rreason = "Di�er - " & rrtext.Text
    ElseIf rrcombo.Value = "NCR �ptali" Then
    rreason = "�ptal - " & rrtext.Text
    ElseIf rrcombo.Value = "Hata ge�medi" Then
    rreason = "Hata ge�medi"
    Else
    rreason = rrcombo.Value & " - " & rrtext.Value
    End If
    UserName = Application.UserName
    
Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst, rs As ADODB.Recordset 'dim the ADO recordset class
Dim dbPath
Dim x As Long, i As Long
dbPath = wstracker.Cells(1, 9)
Set cnn = New ADODB.Connection
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
Set rst = New ADODB.Recordset
Set rs = New ADODB.Recordset

    resp = wstracker.Cells(idrow, 26).Value
    If rrcombo.Value = "Tamir tekrar�" Then
            TTGrup.Show
            tekrardurum = "Tamir tekrar� - " & rrtext.Value
            yenincrexplan = "Tamir tekrar� - " & wstracker.Cells(idrow, 2).Value
            Dim newid As Integer
            sqlnewid = "SELECT max(ID) FROM NCRDB"
            'COMBOBOX5.VALUE --- DEFECT KODU
            SQL4 = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & " "
            rst.Open SQL4, cnn
            rs.Open sqlnewid, cnn
            photolink = rst.Fields(0).Value
            newid = rs.Fields(0).Value + 1
            Dim test As Long
            Dim oUTpUT As String
            photolink = Left(photolink, Len(photolink) - 1)
            test = InStrRev(photolink, "\")
            oUTpUT = Left(photolink, test)
           'YEN� KLAS�R OLU�TURMAMIZ LAZIM
            finalpath = oUTpUT & newid & " - " & ComboBox5.Value & "\"
            If Not fso6.FolderExists(finalpath) Then
            fso6.CreateFolder finalpath
            End If
            SQL = "UPDATE NCRDB SET STATUS = '" & tekrardurum & "', [SOURCE] = '" & ttgrupcombo & "', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now(), QRESP = '" & resp & "'  WHERE [ID] = " & id & "  "
            SQL0 = "INSERT INTO NCRDB SELECT [PROJECT], [BLADE NUMBER], [MOLD NUMBER], [OPEN DATE], [PS / SS], [EDGE], [SURFACE], [Z (mm)], [X/C (mm)], [Length], [Width], [Depth], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
            'YEN� NCRIN GEREKL� ALANLARINI DOLDUR
            SQL3 = "UPDATE NCRDB SET [ORACLE NUMBER] = '000000', [DEFECT CODE / TYPE] = '" & ComboBox5.Value & "', [EXPLANATION] = '" & rreason & "', [SOURCE] = '" & yenincrexplan & "', [STATUS] = 'A��k', [NCR PICTURE] = '" & finalpath & "' WHERE [ID] = " & newid & " "
    ElseIf rrcombo.Value = "NCR �ptali" Then
            SQL = "UPDATE NCRDB SET STATUS = 'Kapat�ld�', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
    Else
        If UserForm1.Caption = "M�hendislik Onay�" Then
                    If resp = "UT" Then
                    SQL = "UPDATE NCRDB SET STATUS = 'UT Onay� 1', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now(), QRESP = '" & resp & "' WHERE [ID] = " & id & "  "
                    Else
                    SQL = "UPDATE NCRDB SET STATUS = 'After Kontrol', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now(), QRESP = '" & resp & "' WHERE [ID] = " & id & "  "
                    End If
        ElseIf UserForm1.Caption = "2. A�ama Onay�" Or UserForm1.Caption = "Son Kontroller ve After Onay�" Then
                            SQL = "UPDATE NCRDB SET STATUS = '2. A�ama NOK', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        ElseIf UserForm1.Caption = "1. A�ama Onay�" Then
                            SQL = "UPDATE NCRDB SET STATUS = '1. A�ama NOK', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        ElseIf UserForm1.Caption = "UT Onay�" And rrcombo.Value = "Hata ge�medi" Then
                            If length.Value = "" Or En.Value = "" Then
                            MsgBox "Tarama sonras� yeni hata boyutunu giriniz!"
                            Exit Sub
                            End If
                            durum = Right(UTsayi, 1) + 1
                            durum = "UT M�dahalesi " & durum
                            SQL = "UPDATE NCRDB SET STATUS = '" & durum & "', Length = " & length.Value & ", Width = " & En.Value & ", RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        End If
    End If
wstracker.Unprotect Password:="4135911"
cnn.Execute SQL
If SQL0 <> "" Then cnn.Execute SQL0
If SQL3 <> "" Then cnn.Execute SQL3
SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
SQL4 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & newid & " "
cnn.Execute SQL2
cnn.Execute SQL4
cnn.Close
Set rs = Nothing
Set rst = Nothing
Set cnn = Nothing

MsgBox "NCR Red/�ptal edildi."
Unload Me
Call refresh
Exit Sub
err:
    MsgBox "�nternete ba�l� oldu�unu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran g�r�nt�s� at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "S�k�nt� var :("
    Exit Sub
End Sub





Private Sub opminbox_Change()
Dim RemoveChar As Integer
Dim DefaultValue As Long
RemoveChar = 1

'Allow only numbers (no blanks)
  If Not IsNumeric(opminbox.Value) And Len(opminbox) <> 0 Then
    If Len(opminbox) < 1 Then RemoveChar = 0
    opminbox = Left(opminbox, Len(opminbox) - RemoveChar)
  End If
'Change to default value if blank
End Sub

Private Sub rrcombo_Change()
If rrcombo.Value = "Di�er" Or rrcombo.Value = "Tamir tekrar�" Or rrcombo.Value = "NCR �ptali" _
        Or rrcombo.Value = "NCR �n-Arka Foto" Or rrcombo.Value = "Ta�lama Foto" Or rrcombo.Value = "Laminasyon Foto" Or _
        rrcombo.Value = "Dispozisyon uygun de�il" Then
    rrtext.Visible = True
    Label21.Visible = True
ElseIf rrcombo.Value = "Hata ge�medi" Then
    length.Enabled = True
    length.Value = ""
    En.Enabled = True
    En.Value = ""
    Label21.Visible = False
    rrtext.Visible = False
Else
    Label21.Visible = False
    rrtext.Visible = False
End If
End Sub

Private Sub TextBox1_Change()
                Dim wstracker As Worksheet
                Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
If wstracker.Cells(4, 1).Value = "NCR MANAGEMENT - ENGINEERING" Then
Else

Dim RemoveChar As Integer
Dim DefaultValue As Long
RemoveChar = 1

'Allow only numbers (no blanks)
  If Not IsNumeric(TextBox1.Value) And Len(TextBox1) <> 0 Then
    If Len(TextBox1) < 1 Then RemoveChar = 0
    TextBox1 = Left(TextBox1, Len(TextBox1) - RemoveChar)

  End If
  If Len(TextBox1) > 6 Then TextBox1 = Left(TextBox1, Len(TextBox1) - 1)
  If Len(TextBox1) = 6 Then
                dbPath = wstracker.Cells(1, 9)
                Set cnn = New ADODB.Connection
                cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
                sqlorano = "SELECT TOP 1 * FROM NCRDB WHERE [ORACLE NUMBER] LIKE '" & TextBox1.Value & "%' ORDER BY [OPEN DATE] DESC"
                Set rsorano = New ADODB.Recordset
                rsorano.Open sqlorano, cnn
                        If rsorano.EOF And rsorano.BOF Then
                        'MsgBox "BU NO BULUNAMADI"
                        yenino = TextBox1.Value & "-1"
                        ComboBox8.Value = ""
                        ComboBox7.Value = ""
                        TextBox2.Value = ""
                        ComboBox6.Value = ""
                        ComboBox5.Value = ""
                        
                        ComboBox8.Enabled = True
                        ComboBox7.Enabled = True
                        TextBox2.Enabled = True
                        ComboBox6.Enabled = True
                        ComboBox5.Enabled = True
                        Else
                        hop2 = True
                        'buraya ncr'a numara al�nacak
                        'MsgBox "Bu NCR no al�nm�� ya da ayn� NCR daha �nce girilmi�", vbCritical, "ORACLE NO HATASI"
                        yenino = TextBox1.Value & "-" & Split(rsorano(1).Value, "-")(UBound(Split(rsorano(1).Value, "-"))) + 1
                        ComboBox8.Value = rsorano(2).Value
                        ComboBox7.Value = Split(rsorano(3).Value, "-")(0)
                        TextBox2.Value = Split(rsorano(3).Value, "-")(UBound(Split(rsorano(3).Value, "-")))
                        ComboBox6.Value = rsorano(4).Value
                        ComboBox5.Value = rsorano(6).Value
                        ComboBox8.Enabled = False
                        ComboBox7.Enabled = False
                        TextBox2.Enabled = False
                        ComboBox6.Enabled = False
                        ComboBox5.Enabled = False
                        End If
                On Error Resume Next
                hop2 = False
                rsorano.Close
                cnn.Close
                Set rsorano = Nothing
                Set cnn = Nothing
     Else
                        ComboBox8.Value = ""
                        ComboBox7.Value = ""
                        TextBox2.Value = ""
                        ComboBox6.Value = ""
                        ComboBox5.Value = ""
                        ComboBox8.Enabled = True
                        ComboBox7.Enabled = True
                        TextBox2.Enabled = True
                        ComboBox6.Enabled = True
                        ComboBox5.Enabled = True
  End If
End If
'Change to default value if blank
End Sub

Private Sub TextBox2_Change()

Dim RemoveChar As Integer
Dim DefaultValue As Long
RemoveChar = 0

'Allow only numbers (no blanks)
    If Len(TextBox2) > 4 Then
    RemoveChar = 1
    MsgBox "Par�a no 4 haneden b�y�k olamaz", vbExclamation, "Yanl�� Kanat No"
    End If
    TextBox2 = Left(TextBox2, Len(TextBox2) - RemoveChar)
'Change to default value if blank

End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, BUTton As Integer, Shift As Integer, x As Single, y As Single)
    i = 0
    For Each picturepath In Data.Files()
    If Not picturepath = "" Then
        fname = Split(picturepath, "\")
        Result = fname(UBound(fname))
      patharray(i) = picturepath
      On Error Resume Next
      TreeView1.Nodes.Add Key:=Result, Text:=Result
      i = i + 1
    End If
    Next picturepath

End Sub
Private Sub TreeView2_OLEDragDrop(Data2 As MSComctlLib.DataObject, Effect As Long, BUTton As Integer, Shift As Integer, x As Single, y As Single)
    j = 0
    For Each picturepath2 In Data2.Files()
    If Not picturepath2 = "" Then
        fname2 = Split(picturepath2, "\")
        Result2 = fname2(UBound(fname2))
      patharray2(j) = picturepath2
      On Error Resume Next
      TreeView2.Nodes.Add Key:=Result2, Text:=Result2
      j = j + 1
    End If
    Next picturepath2

End Sub

Private Sub UserForm_Initialize()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
'On Error GoTo err
TreeView1.OLEDropMode = ccOLEDropManual
TreeView2.OLEDropMode = ccOLEDropManual
oraclenovar = False
If idrow = Key Then
ComboBox1.List = Array("PS", "SS", "PS-SS")
ComboBox2.List = Array("LE", "TE", "LE-TE")
ComboBox3.List = Array("A", "B", "A-B")
ComboBox4.List = Array("Debag Inspection", "Demold Inspection", "Trim Inspection", "Firewall", "Other")
TreeView1.OLEDropMode = ccOLEDropManual
TreeView2.OLEDropMode = ccOLEDropManual
ComboBox6.List = Array("1", "2", "3", "4")
ComboBox7.List = Array("KANAT", "SHELL", "MSW", "TSWROOT", "TSWTIP", "TEINSEND", "TEINSMID", "RF")
ComboBox8.List = Array("V162", "NX74.5", "V136")
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - PRODUCTION" Or wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY" And ((wstracker.Cells(idrow, 18).Value = "A��k" Or wstracker.Cells(idrow, 18).Value = "1. A�ama NOK" Or wstracker.Cells(idrow, 18).Value = "2. A�ama NOK" Or wstracker.Cells(idrow, 18).Value = "2. A�ama" Or wstracker.Cells(idrow, 18).Value Like "UT M*")) Then
               If wstracker.Cells(idrow, 27).Value = "1 - Laminasyon UYGULANMAYACAK" Or wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then
            '4 - Delbas ve kuma� at
            '2 - Kat kald�r ve kuma� at
            '3 - Delbas
            '1 - Laminasyon UYGULANMAYACAK
                TreeView1.OLEDropMode = ccOLEDropManual
                TreeView2.Visible = False
                If wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then TreeView1.Visible = True Else TreeView1.Visible = False
                If wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then Label11.Visible = True Else Label11.Visible = False

                Label2.Visible = False
                ComboBox7.Visible = False
                    TextBox2.Visible = False
                    TextBox10.Visible = True
                    TextBox1.Value = wstracker.Cells(idrow, 2).Value
                    TextBox1.Enabled = False
                    ComboBox8.Value = wstracker.Cells(idrow, 3).Value
                    ComboBox8.Enabled = False
                    TextBox10.Value = wstracker.Cells(idrow, 4).Value
                    TextBox10.Enabled = False
                    ComboBox6.Value = wstracker.Cells(idrow, 5).Value
                    ComboBox6.Enabled = False
                    ComboBox5.Value = wstracker.Cells(idrow, 7).Value
                    ComboBox5.Enabled = False
                    ComboBox1.Value = wstracker.Cells(idrow, 8).Value
                    ComboBox1.Enabled = False
                    ComboBox2.Value = wstracker.Cells(idrow, 9).Value
                    ComboBox2.Enabled = False
                    ComboBox3.Value = wstracker.Cells(idrow, 10).Value
                    ComboBox3.Enabled = False
                    ztext.Value = wstracker.Cells(idrow, 11).Value
                    ztext.Enabled = False
                    xtext.Value = wstracker.Cells(idrow, 12).Value
                    xtext.Enabled = False
                    length.Value = wstracker.Cells(idrow, 13).Value
                    length.Enabled = False
                    En.Value = wstracker.Cells(idrow, 14).Value
                    En.Enabled = False
                    Depth.Value = wstracker.Cells(idrow, 15).Value
                    Depth.Enabled = False
                    TextBox9.Value = wstracker.Cells(idrow, 16).Value
                    TextBox9.Enabled = False
                    ComboBox4.Value = wstracker.Cells(idrow, 17).Value
                    ComboBox4.Enabled = False
                    Label20.Caption = "Tamiri yapanlar"
                    Label11.Caption = "Del-bas Foto�raflar�"
                    opminbox.Enabled = True
                    opadi.Enabled = True
                    UTsayi = wstracker.Cells(idrow, 18).Value
                UserForm1.Caption = "Tek a�amal� tamir"
                CommandButton1.Caption = "Onaya g�nder"
                Else
                                If wstracker.Cells(idrow, 18).Value = "2. A�ama" Then
                                TreeView1.Visible = False
                                TreeView1.OLEDropMode = ccOLEDropManual
                                Label11.Visible = False
                                Label11.Caption = "NCR Arka y�z� ve dispozisyon"
                                Else
                                TreeView1.Visible = False
                                Label11.Visible = False
                                End If
                    TreeView2.Visible = False
                    Label2.Visible = False
                    ComboBox7.Visible = False
                    TextBox2.Visible = False
                    TextBox10.Visible = True
                    TextBox1.Value = wstracker.Cells(idrow, 2).Value
                    TextBox1.Enabled = False
                    ComboBox8.Value = wstracker.Cells(idrow, 3).Value
                    ComboBox8.Enabled = False
                    TextBox10.Value = wstracker.Cells(idrow, 4).Value
                    TextBox10.Enabled = False
                    ComboBox6.Value = wstracker.Cells(idrow, 5).Value
                    ComboBox6.Enabled = False
                    ComboBox5.Value = wstracker.Cells(idrow, 7).Value
                    ComboBox5.Enabled = False
                    ComboBox1.Value = wstracker.Cells(idrow, 8).Value
                    ComboBox1.Enabled = False
                    ComboBox2.Value = wstracker.Cells(idrow, 9).Value
                    ComboBox2.Enabled = False
                    ComboBox3.Value = wstracker.Cells(idrow, 10).Value
                    ComboBox3.Enabled = False
                    ztext.Value = wstracker.Cells(idrow, 11).Value
                    ztext.Enabled = False
                    xtext.Value = wstracker.Cells(idrow, 12).Value
                    xtext.Enabled = False
                    length.Value = wstracker.Cells(idrow, 13).Value
                    length.Enabled = False
                    En.Value = wstracker.Cells(idrow, 14).Value
                    En.Enabled = False
                    Depth.Value = wstracker.Cells(idrow, 15).Value
                    Depth.Enabled = False
                    TextBox9.Value = wstracker.Cells(idrow, 16).Value
                    TextBox9.Enabled = False
                    ComboBox4.Value = wstracker.Cells(idrow, 17).Value
                    ComboBox4.Enabled = False
                    Label20.Caption = "Tamiri yapanlar"
                    opminbox.Enabled = True
                    opadi.Enabled = True

                UserForm1.Caption = "�ki a�amal� tamir"
                CommandButton1.Caption = "Onaya g�nder"
                End If
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - FINISH QUALITY" Or wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY" Or wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" And (wstracker.Cells(idrow, 18).Value <> "A��k" Or wstracker.Cells(idrow, 18).Value <> "1. A�ama NOK" Or wstracker.Cells(idrow, 18).Value <> "2. A�ama NOK" Or wstracker.Cells(idrow, 18).Value <> "2. A�ama") Then
                rrcombo.List = Array( _
                    "Yetersiz ta�lama", "Tamir tekrar�", "Hatal� kat a��lmas�", _
                    "Di�er", "NCR �ptali")
                TreeView2.Visible = False
                Label21.Visible = False
                Label21.Caption = "Detayl� a��klama"
                If wstracker.Cells(idrow, 18).Value = "1. A�ama Onay�" Then
                UserForm1.Caption = "1. A�ama Onay�"
                Label11.Caption = "1. A�ama tamir foto�raflar�"
                ElseIf wstracker.Cells(idrow, 18).Value = "2. A�ama Onay�" Then
                UserForm1.Caption = "2. A�ama Onay�"
                Label11.Caption = "2. A�ama tamir foto�raflar�"
                ElseIf wstracker.Cells(idrow, 18).Value = "After Kontrol" Then
                UserForm1.Caption = "Son Kontroller ve After Onay�"
                Label11.Caption = "AFTER FOTO�RAFLARI"
                TreeView2.Visible = True
                TreeView2.Height = 48
                TreeView2.Top = 276
                Label21.Visible = True
                Label21.Caption = "NCR �n ve Arka Y�z�"
                                On Error GoTo atla2
                                If wstracker.Cells(idrow, 27).Value = "1 - Laminasyon UYGULANMAYACAK" Or wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then
                                Label11.Visible = False
                                TreeView1.Visible = False
                                TreeView2.Top = TreeView1.Top
                                TreeView2.Left = TreeView1.Left
                                TreeView2.Height = TreeView1.Height
                                Label21.Top = Label11.Top
                                Label21.Left = Label11.Left
                                Label21.Height = Label11.Height
                                End If
atla2:
                ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" Then
                UTsayi = wstracker.Cells(idrow, 18).Value
                UserForm1.Caption = "UT Onay�"
                Label11.Caption = "AFTER FOTO�RAFLARI"
                rrcombo.List = Array("Hata ge�medi", _
                    "Tamir tekrar�", _
                    "Di�er", "NCR �ptali")
                ComboBox4.Value = "UT"
                ComboBox4.Enabled = False
                End If
                TreeView1.OLEDropMode = ccOLEDropManual
                rrcombo.Visible = True
                rrtext.Visible = True
                    ComboBox7.Visible = False
                    TextBox2.Visible = False
                    TextBox10.Visible = True
                    TextBox1.Value = wstracker.Cells(idrow, 2).Value
                    TextBox1.Enabled = False
                    ComboBox8.Value = wstracker.Cells(idrow, 3).Value
                    ComboBox8.Enabled = False
                    TextBox10.Value = wstracker.Cells(idrow, 4).Value
                    TextBox10.Enabled = False
                    ComboBox6.Value = wstracker.Cells(idrow, 5).Value
                    ComboBox6.Enabled = False
                    ComboBox5.Value = wstracker.Cells(idrow, 7).Value
                    ComboBox5.Enabled = False
                    ComboBox1.Value = wstracker.Cells(idrow, 8).Value
                    ComboBox1.Enabled = False
                    ComboBox2.Value = wstracker.Cells(idrow, 9).Value
                    ComboBox2.Enabled = False
                    ComboBox3.Value = wstracker.Cells(idrow, 10).Value
                    ComboBox3.Enabled = False
                    ztext.Value = wstracker.Cells(idrow, 11).Value
                    ztext.Enabled = False
                    xtext.Value = wstracker.Cells(idrow, 12).Value
                    xtext.Enabled = False
                    length.Value = wstracker.Cells(idrow, 13).Value
                    length.Enabled = False
                    En.Value = wstracker.Cells(idrow, 14).Value
                    En.Enabled = False
                    Depth.Value = wstracker.Cells(idrow, 15).Value
                    Depth.Enabled = False
                    TextBox9.Value = wstracker.Cells(idrow, 16).Value
                    TextBox9.Enabled = False
                    ComboBox4.Value = wstracker.Cells(idrow, 17).Value
                    ComboBox4.Enabled = False
                    Label20.Visible = True
                    Label20.Caption = "Tamiri reddetme nedeni"
                    opminbox.Enabled = False
                    opadi.Enabled = False
                    opminbox.Visible = False
                    opadi.Visible = False
                CommandButton1.Caption = "Onayla"
                CommandButton2.Visible = True
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - ENGINEERING" Then
                UserForm1.Caption = "M�hendislik Onay�"
                Label11.Visible = False
                TreeView1.Visible = False
                rrcombo.Visible = True
                rrtext.Visible = False
                rrcombo.List = Array("NCR �n-Arka Foto", "Ta�lama Foto", "Laminasyon Foto", _
                    "Tamir tekrar�", _
                    "Dispozisyon uygun de�il", "Di�er")
                TreeView2.Visible = False
                    Label2.Visible = False
                    ComboBox7.Visible = False
                    TextBox2.Visible = False
                    TextBox10.Visible = True
                    TextBox1.Value = wstracker.Cells(idrow, 2).Value
                    TextBox1.Enabled = False
                    ComboBox8.Value = wstracker.Cells(idrow, 3).Value
                    ComboBox8.Enabled = False
                    TextBox10.Value = wstracker.Cells(idrow, 4).Value
                    TextBox10.Enabled = False
                    ComboBox6.Value = wstracker.Cells(idrow, 5).Value
                    ComboBox6.Enabled = False
                    ComboBox5.Value = wstracker.Cells(idrow, 7).Value
                    ComboBox5.Enabled = False
                    ComboBox1.Value = wstracker.Cells(idrow, 8).Value
                    ComboBox1.Enabled = False
                    ComboBox2.Value = wstracker.Cells(idrow, 9).Value
                    ComboBox2.Enabled = False
                    ComboBox3.Value = wstracker.Cells(idrow, 10).Value
                    ComboBox3.Enabled = False
                    ztext.Value = wstracker.Cells(idrow, 11).Value
                    ztext.Enabled = False
                    xtext.Value = wstracker.Cells(idrow, 12).Value
                    xtext.Enabled = False
                    length.Value = wstracker.Cells(idrow, 13).Value
                    length.Enabled = False
                    En.Value = wstracker.Cells(idrow, 14).Value
                    En.Enabled = False
                    Depth.Value = wstracker.Cells(idrow, 15).Value
                    Depth.Enabled = False
                    TextBox9.Value = wstracker.Cells(idrow, 16).Value
                    TextBox9.Enabled = False
                    ComboBox4.Value = wstracker.Cells(idrow, 17).Value
                    ComboBox4.Enabled = False
                    Label20.Visible = True
                    Label21.Visible = False
                    Label20.Caption = "Tamiri reddetme nedeni"
                    Label21.Caption = "Detayl� a��klama"
                    opminbox.Enabled = False
                    opadi.Enabled = False
                    opminbox.Visible = False
                    opadi.Visible = False
                CommandButton1.Caption = "Onayla"
                CommandButton2.Visible = True
Else
Exit Sub
End If
If wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" Then
                'ut ncr� olu�turma k�sm�
                UserForm1.Caption = "UT NCR'� Olu�tur"
                ComboBox4.List = Array("Termal", "TE Insert", "MSW", "TESW")
                ComboBox4.Enabled = True
                ComboBox4.Value = "Termal"
                TreeView2.Visible = False
                Label20.Visible = False
                rrcombo.Visible = False
                rrtext.Visible = False
                Label21.Visible = False
                opadi.Visible = False
                opminbox.Visible = False
                Label11.Caption = "Hata Foto�raf� ya da Analiz Exceli"
                On Error GoTo atla
                If wstracker.Cells(idrow, 18).Value Like "UT O*" Then
                UserForm1.Caption = "UT Onay�"
                Label11.Caption = "NCR �n ve Arka Y�z�"
                rrcombo.Visible = True
                rrtext.Visible = False
                Label20.Visible = True
                End If
atla:
End If

Exit Sub
err:
    MsgBox "�nternete ba�l� oldu�unu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran g�r�nt�s� at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "S�k�nt� var :("
    Exit Sub
End Sub

 
Private Sub UserForm_Terminate()
Call refresh
End Sub
