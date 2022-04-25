VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "NCR Oluþtur"
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
    ComboBox5.List = Array("D010-Dry Glass-Kuru Kumaþ", "D020-Semi Dry Glass-Yarý Kuru Kumaþ", _
    "D050-Entrained Air-Hava Almýþ Bölge", "D055-Damaged Glass-Laminasyonda Hasar", _
    "D060-Foreign Objects-Yabancý Madde", "D070-Cracks-Çatlak", "D072 Delamination-Delaminasyon", _
    "D074 Resin Rich-Zengin Reçine", "D080-Paste Voids - Exposed-Yetersiz Krema", _
    "D082 Paste Voids - Non-Exposed-Kremada Hava", "D090-Resin Voids-Reçinede Hava", _
    "D101 Lps Damage-Lps Hasarý", "D102 Lps Continuity-Lps Baðlantý - Direnç Hatasý", _
    "D103 Shiny Spots-Parlak Noktalar", "D104 Loose Laminates-Kalkýk Kumaþ Lifleri", _
    "D106 Folded Laminates-Kumaþ Katlanmasý", "D109 Glass Or Fabric Location-Kumaþ Konum Hatasý", _
    "D110-Core Discoloration-Nüve Yanýðý", "D111 Core Gap-Nüve Boþluðu", _
    "D112 Missing Core-Eksik Nüve", "D114 Core Location-Nüve Konum Hatasý-Kayma", _
    "D150-Waves & Wrinkles-Dalga", "D160-Overbite-Underbite-Ps-Ss Kabu Kaçýklýk", _
    "D170-Trailing Edge Thickness-Te Kalýnlýk", "D220-Bond Line Thickness-Yapýþma Hattý-Krema Kalýnlýðý ", _
    "D310-insufficient Cure Or Failed Tg-Yetersiz Kürlenme (Tg)", "D320-Blade-Subcomponent Geometry-Kanat-Parça Geometrisi", _
    "D330-Subcomponent Position-Parça Pozisyon Hatasý-Hizasýzlýðý", _
    "D350-Mass Or Moment-Aðýrlýk Veya Moment", "D370-Coatings-Yüzey Kaplamalarý", _
    "D401 Disbond-Yapýþmama", "D405-Lps Position-Lps Pozisyonu", _
    "D505-Excessive Adhesive-Krema Taþmasý", "D510-Blocked Drainage Hole-Týkalý Drenaj Deliði", _
    "D635-Core Step-Nüve Yükseklik Farký", "D640-Edge Step-Nüve Kenar Yükseklik Farký", _
    "D655-Holes in Laminate-Laminasyonda Delik", "D670-Flange Step-Flanþ Geçiþ Yükseklik Farký", _
    "D675-Short Flange-Kýsa Flanþ", "D680-Bond Width-Yapýþma Geniþliði", _
    "D700-Excessive Tackifier-Fazla Sprey Kullanýmý", "D705-Bond Cap Deviation-Yapýþma Kebi Hatasý", _
    "D710-Lps Tip Position-Smt Kaymasý", "D720-Steel insert Protrusion-Yüzeyde Insert Çýkýntýsý", _
    "D725-Glass Overlap-Kumaþ Bindirmesi", "D740-Pultrusion Chamfer-Pultruzyon Pah Hatasý", _
    "D755-Groove-Oluk", "D760-Handling Damage-Taþýma Hasarý", "D770-Pultrusion Void-Pultruzyonda Hava", "D924 Laminate Discoloration-Laminasyonda Renk Hatasý", _
    "D970-Exceeding Surface Preparation Time-Yüzey Aktivasyon Süre Aþýmý")
    ComboBox5.Value = ""
    End If
ElseIf ComboBox8.Value = "V162" Then
ComboBox7.List = Array("KANAT", "SHELL", "MSW", "TSW", "TEINS", "RF")
ComboBox2.List = Array("LE", "TE", "LE-TE")
ComboBox7.Value = ""
ComboBox2.Value = ""
    If hop2 = False Then
    ComboBox2.Value = ""
    ComboBox5.List = Array("D010-Dry Glass-Kuru Kumaþ", "D020-Semi Dry Glass-Yarý Kuru Kumaþ", _
    "D050-Entrained Air-Hava Almýþ Bölge", "D055-Damaged Glass-Laminasyonda Hasar", _
    "D060-Foreign Objects-Yabancý Madde", "D070-Cracks-Çatlak", "D072 Delamination-Delaminasyon", _
    "D074 Resin Rich-Zengin Reçine", "D080-Paste Voids - Exposed-Yetersiz Krema", _
    "D082 Paste Voids - Non-Exposed-Kremada Hava", "D090-Resin Voids-Reçinede Hava", _
    "D101 Lps Damage-Lps Hasarý", "D102 Lps Continuity-Lps Baðlantý - Direnç Hatasý", _
    "D103 Shiny Spots-Parlak Noktalar", "D104 Loose Laminates-Kalkýk Kumaþ Lifleri", _
    "D106 Folded Laminates-Kumaþ Katlanmasý", "D109 Glass Or Fabric Location-Kumaþ Konum Hatasý", _
    "D110-Core Discoloration-Nüve Yanýðý", "D111 Core Gap-Nüve Boþluðu", _
    "D112 Missing Core-Eksik Nüve", "D114 Core Location-Nüve Konum Hatasý-Kayma", _
    "D150-Waves & Wrinkles-Dalga", "D160-Overbite-Underbite-Ps-Ss Kabu Kaçýklýk", _
    "D170-Trailing Edge Thickness-Te Kalýnlýk", "D220-Bond Line Thickness-Yapýþma Hattý-Krema Kalýnlýðý ", _
    "D310-insufficient Cure Or Failed Tg-Yetersiz Kürlenme (Tg)", "D320-Blade-Subcomponent Geometry-Kanat-Parça Geometrisi", _
    "D330-Subcomponent Position-Parça Pozisyon Hatasý-Hizasýzlýðý", _
    "D350-Mass Or Moment-Aðýrlýk Veya Moment", "D370-Coatings-Yüzey Kaplamalarý", _
    "D401 Disbond-Yapýþmama", "D405-Lps Position-Lps Pozisyonu", _
    "D505-Excessive Adhesive-Krema Taþmasý", "D510-Blocked Drainage Hole-Týkalý Drenaj Deliði", _
    "D635-Core Step-Nüve Yükseklik Farký", "D640-Edge Step-Nüve Kenar Yükseklik Farký", _
    "D655-Holes in Laminate-Laminasyonda Delik", "D670-Flange Step-Flanþ Geçiþ Yükseklik Farký", _
    "D675-Short Flange-Kýsa Flanþ", "D680-Bond Width-Yapýþma Geniþliði", _
    "D700-Excessive Tackifier-Fazla Sprey Kullanýmý", "D705-Bond Cap Deviation-Yapýþma Kebi Hatasý", _
    "D710-Lps Tip Position-Smt Kaymasý", "D720-Steel insert Protrusion-Yüzeyde Insert Çýkýntýsý", _
    "D725-Glass Overlap-Kumaþ Bindirmesi", "D740-Pultrusion Chamfer-Pultruzyon Pah Hatasý", _
    "D755-Groove-Oluk", "D760-Handling Damage-Taþýma Hasarý", "D770-Pultrusion Void-Pultruzyonda Hava", "D924 Laminate Discoloration-Laminasyonda Renk Hatasý", _
    "D970-Exceeding Surface Preparation Time-Yüzey Aktivasyon Süre Aþýmý")
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
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
    End If
        versiyon = rsvers.Fields(0).Value
        If wstracker.Cells(3, 9) <> versiyon Then
        MsgBox "Eski bir sürüm kullanýyorsunuz baðlantý yapýlmadý!"
        Exit Sub
        End If

'Erro handler
'On Error GoTo errHandler:
'Foto kopyala
folderpath = wstracker.Cells(2, 9)

If TextBox9.Enabled = True Then
    If ComboBox5.Value = "D500 - Other" Or ComboBox5.Value = "D060 - FOD" Or ComboBox5.Value = "D655 - Hole in Laminate" Then
        If TextBox9.Value = "" Then
        MsgBox "Açýklama kýsmý doldurulmalýdýr.", vbCritical
        Exit Sub
        End If
    End If
End If

'If Me.TreeView1.Nodes.Count = 0 And UserForm1.Caption <> "Ýki aþamalý tamir" And UserForm1.Caption <> "Mühendislik Onayý" Then
If Me.TreeView1.Nodes.Count = 0 And Me.TreeView1.Visible = True Then
    MsgBox "Devam etmek için fotoðraflarý ekleyiniz"
    Exit Sub
End If
If Me.TreeView2.Nodes.Count = 0 And Me.TreeView2.Visible = True Then
    MsgBox "Devam etmek için fotoðraflarý ekleyiniz"
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

'klasör için yeni id al

If UserForm1.Caption = "Son Kontroller ve After Onayý" Then
    If TextBox1.Value = "000000" Then
    MsgBox "Oracle numarasý alýnmadan NCR kapatýlamaz"
    Exit Sub
    Else
    SQL = "UPDATE NCRDB SET STATUS = 'Müh. Onayý', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
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
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
    End If
        link = rs.Fields(0).Value
        rs.Close
'link okay
        finalpath = link & "5 - Tamir Sonrasý\"
        If Not fso6.FolderExists(finalpath) Then
               fso6.CreateFolder finalpath
        End If
        'fotoyu da yükle
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
        finalpath = link & "1 - NCR ön ve arka yüzü\"
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
ElseIf UserForm1.Caption = "UT Onayý" Then
    If TextBox1.Value = "000000" Then
    MsgBox "Oracle numarasý alýnmadan NCR kapatýlamaz"
    Exit Sub
    Else
    SQL = "UPDATE NCRDB SET STATUS = 'Müh. Onayý', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
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
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
    End If
        link = rs.Fields(0).Value
        rs.Close
'link okay
        finalpath = link & "1 - NCR ön ve arka yüzü\"
        If Not fso6.FolderExists(finalpath) Then
               fso6.CreateFolder finalpath
        End If
        'fotoyu da yükle
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
ElseIf UserForm1.Caption = "NCR Oluþtur" Or UserForm1.Caption = "UT NCR'ý Oluþtur" Then
        If TextBox2.Value = "" Or ComboBox1.Value = "" Or ComboBox2.Value = "" _
        Or ComboBox3.Value = "" Or ComboBox4.Value = "" _
        Or ComboBox6.Value = "" Or ComboBox7.Value = "" Or ComboBox7.Value = "" Then
        MsgBox "Tüm alanlar doldurulmalýdýr.", vbCritical, "ZORUNLU ALANLAR"
        Exit Sub
        End If
                If Len(TextBox1.Value) = 6 Or TextBox1.Value = "" Or TextBox1.Value = "000000" Then
                Else
                MsgBox "NCR No hatalý veya eksik girildi"
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
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
        End If
        
        '        'yapay zeka kasalým :)
        Dim alan, tahmin As Integer
        alan = length.Value * En.Value
        Dim sqlhesap As String
        SQL = "SELECT TOP 5 * " & _
                   "FROM NCRDB " & _
                   "WHERE NCRDB.[DEFECT CODE / TYPE] = '" & ComboBox5.Value & "' And NCRDB.SURFACE = '" & ComboBox3.Value & "' And NCRDB.TOTALMIN > 0" & _
                   " ORDER BY Abs(" & alan & "-([NCRDB].[LENGTH]*[NCRDB].[Width]))"
        sqlhesap = "SELECT AVG(TOTALMIN) AS TheAverage FROM (" & SQL & ")"
        
'        'BURDAN SONRASI SÝLÝNECEK
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
                    'KLASÖRLEME
                    'ana klasör
                    newfolder1 = folderpath & "\" & ComboBox8.Value & "\"
                    'kanatsa veya shellse 2. klasör
                    If ComboBox7.Value = "KANAT" Or ComboBox7.Value = "SHELL" Then
                    newpath = newfolder1 & "KANAT\"
                    newpath2 = newpath & var & "\"
                        'kanat mý shell mi
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
                    'küçük parçaysa
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
                    If UserForm1.Caption = "UT NCR'ý Oluþtur" Then
                    finalpath = newpath4 & "2 - Tamir öncesi\"
                    'else kalite açýyorsa ncr'ý
                    Else
                    finalpath = newpath4 & "1 - NCR ön ve arka yüzü\"
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
                    finalpath = newpath4 & "2 - Tamir öncesi\"
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
        .Fields("STATUS").Value = "Açýk"
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
        MsgBox " NCR Baþarýyla oluþturuldu ve fotoðraflar arþivlendi."
        If wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" And ComboBox5.Value = "D655 - Hole in Laminate" Then
        MsgBox " Delikler arasý mesafe NCR'ý Finish Kaliteye Aktarýldý, kaðýdý onlara teslim edin"
        End If
ElseIf UserForm1.Caption = "Tek aþamalý tamir" Then
        If opminbox.Value = "" Or opadi.Value = "" Then
        MsgBox "Onaya göndermek için gerekli alanlarý doldurun."
        Exit Sub
        End If
'production tek aþamalý tamir upload
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
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
        End If
        link = rs.Fields(0).Value
        rs.Close
'link okay
        opisim = opadi.Value
        opmin = opminbox.Value
            If wstracker.Cells(idrow, 26).Value <> "UT" Then
            SQL = "UPDATE NCRDB SET STATUS = '1. Aþama Onayý', RR = ' ', SCOP = '" & opisim & "', SCMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
            Else
                If UTsayi = "Açýk" Then
                durum = "UT Onayý 0"
                durum = Right(durum, 1) + 1
                durum = "UT Onayý " & durum
                ElseIf UTsayi Like "UT M*" Then
                durum = Right(UTsayi, 1)
                durum = "UT Onayý " & durum
                ElseIf durum = "" Then durum = "UT Onayý 1"
                End If
            SQL = "UPDATE NCRDB SET STATUS = '" & durum & "', RR = ' ', SCOP = '" & opisim & "', SCMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
            End If
            finalpath = link & "3 - 1. Aþama Tamir\"
            If Not fso6.FolderExists(finalpath) Then
                  fso6.CreateFolder finalpath
            End If
        'fotoyu da yükle
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

ElseIf UserForm1.Caption = "1. Aþama Onayý" Or UserForm1.Caption = "2. Aþama Onayý" Then
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
            MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
            Exit Sub
        End If
        link = rs.Fields(0).Value
        rs.Close
        'link okay
    If UserForm1.Caption = "1. Aþama Onayý" Then
        finalpath = link & "3 - 1. Aþama Tamir\"
            'tamir eðer tek aþamalýysa
            If wstracker.Cells(idrow, 27).Value = "1 - Laminasyon UYGULANMAYACAK" Or wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then
            SQL = "UPDATE NCRDB SET STATUS = 'After Kontrol', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
            'tamir eðer çok aþamalýysa
            Else
            SQL = "UPDATE NCRDB SET STATUS = '2. Aþama', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
            End If
    ElseIf UserForm1.Caption = "2. Aþama Onayý" Then
        SQL = "UPDATE NCRDB SET STATUS = 'After Kontrol', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        finalpath = link & "4 - 2. Aþama Tamir\"
    Else
    End If

    If Not fso6.FolderExists(finalpath) Then
          fso6.CreateFolder finalpath
    End If
    
'fotoyu da yükle
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
        

ElseIf UserForm1.Caption = "Ýki aþamalý tamir" Then
        If opminbox.Value = "" Or opadi.Value = "" Then
        MsgBox "Onaya göndermek için gerekli alanlarý doldurun."
        Exit Sub
        End If
'production iki aþamalý tamir upload
        opisim = opadi.Value
        opmin = opminbox.Value
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        'daha 1. aþamadaysa
        If wstracker.Cells(idrow, 18).Value = "Açýk" Or wstracker.Cells(idrow, 18).Value = "1. Aþama NOK" Then
            SQL = "UPDATE NCRDB SET STATUS = '1. Aþama Onayý', RR = ' ', SCOP = '" & opisim & "', SCMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
        '2. aþamadaysa
        ElseIf wstracker.Cells(idrow, 18).Value = "2. Aþama NOK" Or wstracker.Cells(idrow, 18).Value = "2. Aþama" Then
            SQL = "UPDATE NCRDB SET STATUS = '2. Aþama Onayý', RR = ' ', LAMOP = '" & opisim & "', LAMMIN = " & opmin & ", WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
            '2. aþamadaysa fotoðraf yükle
                If wstracker.Cells(idrow, 18).Value = "2. Aþama" Then
                'fotoðraf yüklenecek linki bul
                    SQL0 = "SELECT [NCR PICTURE] FROM NCRDB WHERE [ID] = " & id & ""
                    Set rs = New ADODB.Recordset
                    rs.Open SQL0, cnn
                        If rs.EOF And rs.BOF Then
                                rs.Close
                                cnn.Close
                                Set rs = Nothing
                                Set cnn = Nothing
                                MsgBox "Herhangi bir kayýt bulunamadý server baðlantýnýzý kontrol edin", vbCritical, "Kayýt Bulunamadý"
                                Exit Sub
                        End If
                        link = rs.Fields(0).Value
                        rs.Close
                            'link okay þimdi fotoyu yükle
                                finalpath = link & "1 - NCR ön ve arka yüzü\"
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

ElseIf UserForm1.Caption = "Mühendislik Onayý" Then
        Set cnn = New ADODB.Connection
        cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
        'onayla
        SQL = "UPDATE NCRDB SET STATUS = 'Kapatýldý', RR = ' ', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
        cnn.Execute SQL
        SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
        cnn.Execute SQL2

'tüm seçimlerin sonu end ifi end if gibi end if
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
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
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
    MsgBox "Reddetmek için sebep girilmedi!"
    Exit Sub
    ElseIf rrtext.Visible = True And rrtext.Value = "" Then
    MsgBox "Açýklama girilmedi!"
    Exit Sub
    ElseIf rrcombo.Value = "Tamir tekrarý" Then
        If TextBox1.Value = "000000" Then
        MsgBox "Tamir tekrarý vermek için bu NCR için Oracle NO girilmelidir!"
        Exit Sub
        ElseIf rrtext.Value = "" Then
        MsgBox "Tamir tekrarý için açýklama giriniz"
        Exit Sub
        End If
    
    rreason = "Tamir tekrarý - " & rrtext.Text
    ElseIf rrcombo.Value = "Diðer" Then
    rreason = "Diðer - " & rrtext.Text
    ElseIf rrcombo.Value = "NCR Ýptali" Then
    rreason = "Ýptal - " & rrtext.Text
    ElseIf rrcombo.Value = "Hata geçmedi" Then
    rreason = "Hata geçmedi"
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
    If rrcombo.Value = "Tamir tekrarý" Then
            TTGrup.Show
            tekrardurum = "Tamir tekrarý - " & rrtext.Value
            yenincrexplan = "Tamir tekrarý - " & wstracker.Cells(idrow, 2).Value
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
           'YENÝ KLASÖR OLUÞTURMAMIZ LAZIM
            finalpath = oUTpUT & newid & " - " & ComboBox5.Value & "\"
            If Not fso6.FolderExists(finalpath) Then
            fso6.CreateFolder finalpath
            End If
            SQL = "UPDATE NCRDB SET STATUS = '" & tekrardurum & "', [SOURCE] = '" & ttgrupcombo & "', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now(), QRESP = '" & resp & "'  WHERE [ID] = " & id & "  "
            SQL0 = "INSERT INTO NCRDB SELECT [PROJECT], [BLADE NUMBER], [MOLD NUMBER], [OPEN DATE], [PS / SS], [EDGE], [SURFACE], [Z (mm)], [X/C (mm)], [Length], [Width], [Depth], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
            'YENÝ NCRIN GEREKLÝ ALANLARINI DOLDUR
            SQL3 = "UPDATE NCRDB SET [ORACLE NUMBER] = '000000', [DEFECT CODE / TYPE] = '" & ComboBox5.Value & "', [EXPLANATION] = '" & rreason & "', [SOURCE] = '" & yenincrexplan & "', [STATUS] = 'Açýk', [NCR PICTURE] = '" & finalpath & "' WHERE [ID] = " & newid & " "
    ElseIf rrcombo.Value = "NCR Ýptali" Then
            SQL = "UPDATE NCRDB SET STATUS = 'Kapatýldý', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
    Else
        If UserForm1.Caption = "Mühendislik Onayý" Then
                    If resp = "UT" Then
                    SQL = "UPDATE NCRDB SET STATUS = 'UT Onayý 1', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now(), QRESP = '" & resp & "' WHERE [ID] = " & id & "  "
                    Else
                    SQL = "UPDATE NCRDB SET STATUS = 'After Kontrol', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now(), QRESP = '" & resp & "' WHERE [ID] = " & id & "  "
                    End If
        ElseIf UserForm1.Caption = "2. Aþama Onayý" Or UserForm1.Caption = "Son Kontroller ve After Onayý" Then
                            SQL = "UPDATE NCRDB SET STATUS = '2. Aþama NOK', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        ElseIf UserForm1.Caption = "1. Aþama Onayý" Then
                            SQL = "UPDATE NCRDB SET STATUS = '1. Aþama NOK', RR = '" & rreason & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & "  "
        ElseIf UserForm1.Caption = "UT Onayý" And rrcombo.Value = "Hata geçmedi" Then
                            If length.Value = "" Or En.Value = "" Then
                            MsgBox "Tarama sonrasý yeni hata boyutunu giriniz!"
                            Exit Sub
                            End If
                            durum = Right(UTsayi, 1) + 1
                            durum = "UT Müdahalesi " & durum
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

MsgBox "NCR Red/Ýptal edildi."
Unload Me
Call refresh
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
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
If rrcombo.Value = "Diðer" Or rrcombo.Value = "Tamir tekrarý" Or rrcombo.Value = "NCR Ýptali" _
        Or rrcombo.Value = "NCR Ön-Arka Foto" Or rrcombo.Value = "Taþlama Foto" Or rrcombo.Value = "Laminasyon Foto" Or _
        rrcombo.Value = "Dispozisyon uygun deðil" Then
    rrtext.Visible = True
    Label21.Visible = True
ElseIf rrcombo.Value = "Hata geçmedi" Then
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
                        'buraya ncr'a numara alýnacak
                        'MsgBox "Bu NCR no alýnmýþ ya da ayný NCR daha önce girilmiþ", vbCritical, "ORACLE NO HATASI"
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
    MsgBox "Parça no 4 haneden büyük olamaz", vbExclamation, "Yanlýþ Kanat No"
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
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - PRODUCTION" Or wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY" And ((wstracker.Cells(idrow, 18).Value = "Açýk" Or wstracker.Cells(idrow, 18).Value = "1. Aþama NOK" Or wstracker.Cells(idrow, 18).Value = "2. Aþama NOK" Or wstracker.Cells(idrow, 18).Value = "2. Aþama" Or wstracker.Cells(idrow, 18).Value Like "UT M*")) Then
               If wstracker.Cells(idrow, 27).Value = "1 - Laminasyon UYGULANMAYACAK" Or wstracker.Cells(idrow, 27).Value = "3 - Delbas" Then
            '4 - Delbas ve kumaþ at
            '2 - Kat kaldýr ve kumaþ at
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
                    Label11.Caption = "Del-bas Fotoðraflarý"
                    opminbox.Enabled = True
                    opadi.Enabled = True
                    UTsayi = wstracker.Cells(idrow, 18).Value
                UserForm1.Caption = "Tek aþamalý tamir"
                CommandButton1.Caption = "Onaya gönder"
                Else
                                If wstracker.Cells(idrow, 18).Value = "2. Aþama" Then
                                TreeView1.Visible = False
                                TreeView1.OLEDropMode = ccOLEDropManual
                                Label11.Visible = False
                                Label11.Caption = "NCR Arka yüzü ve dispozisyon"
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

                UserForm1.Caption = "Ýki aþamalý tamir"
                CommandButton1.Caption = "Onaya gönder"
                End If
ElseIf wstracker.Cells(4, 1) = "NCR MANAGEMENT - FINISH QUALITY" Or wstracker.Cells(4, 1) = "NCR MANAGEMENT - MOLD QUALITY" Or wstracker.Cells(4, 1) = "NCR MANAGEMENT - UT" And (wstracker.Cells(idrow, 18).Value <> "Açýk" Or wstracker.Cells(idrow, 18).Value <> "1. Aþama NOK" Or wstracker.Cells(idrow, 18).Value <> "2. Aþama NOK" Or wstracker.Cells(idrow, 18).Value <> "2. Aþama") Then
                rrcombo.List = Array( _
                    "Yetersiz taþlama", "Tamir tekrarý", "Hatalý kat açýlmasý", _
                    "Diðer", "NCR Ýptali")
                TreeView2.Visible = False
                Label21.Visible = False
                Label21.Caption = "Detaylý açýklama"
                If wstracker.Cells(idrow, 18).Value = "1. Aþama Onayý" Then
                UserForm1.Caption = "1. Aþama Onayý"
                Label11.Caption = "1. Aþama tamir fotoðraflarý"
                ElseIf wstracker.Cells(idrow, 18).Value = "2. Aþama Onayý" Then
                UserForm1.Caption = "2. Aþama Onayý"
                Label11.Caption = "2. Aþama tamir fotoðraflarý"
                ElseIf wstracker.Cells(idrow, 18).Value = "After Kontrol" Then
                UserForm1.Caption = "Son Kontroller ve After Onayý"
                Label11.Caption = "AFTER FOTOÐRAFLARI"
                TreeView2.Visible = True
                TreeView2.Height = 48
                TreeView2.Top = 276
                Label21.Visible = True
                Label21.Caption = "NCR Ön ve Arka Yüzü"
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
                UserForm1.Caption = "UT Onayý"
                Label11.Caption = "AFTER FOTOÐRAFLARI"
                rrcombo.List = Array("Hata geçmedi", _
                    "Tamir tekrarý", _
                    "Diðer", "NCR Ýptali")
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
                UserForm1.Caption = "Mühendislik Onayý"
                Label11.Visible = False
                TreeView1.Visible = False
                rrcombo.Visible = True
                rrtext.Visible = False
                rrcombo.List = Array("NCR Ön-Arka Foto", "Taþlama Foto", "Laminasyon Foto", _
                    "Tamir tekrarý", _
                    "Dispozisyon uygun deðil", "Diðer")
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
                    Label21.Caption = "Detaylý açýklama"
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
                'ut ncrý oluþturma kýsmý
                UserForm1.Caption = "UT NCR'ý Oluþtur"
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
                Label11.Caption = "Hata Fotoðrafý ya da Analiz Exceli"
                On Error GoTo atla
                If wstracker.Cells(idrow, 18).Value Like "UT O*" Then
                UserForm1.Caption = "UT Onayý"
                Label11.Caption = "NCR Ön ve Arka Yüzü"
                rrcombo.Visible = True
                rrtext.Visible = False
                Label20.Visible = True
                End If
atla:
End If

Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub

 
Private Sub UserForm_Terminate()
Call refresh
End Sub
