VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dispselect 
   Caption         =   "Dispozisyon Tipini Seç"
   ClientHeight    =   2484
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4212
   OleObjectBlob   =   "Dispselect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dispselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
If ComboBox1.Value = "1 - Laminasyon UYGULANMAYACAK" Then
ElseIf ComboBox1.Value = "2 - Kat kaldýr ve kumaþ at" Then
ComboBox2.List = Array("El Yatýrmasý", "Vacuum Bag", "Infusion", "Güçlendirme", "Pencere", "Te ins Güçlendirme", "Aðýr Tamir")
ElseIf ComboBox1.Value = "3 - Delbas" Then
End If
End Sub

Private Sub CommandButton1_Click()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
On Error GoTo err
Application.ScreenUpdating = False
Application.Calculation = xlManual

    UserName = Application.UserName
    
    If ComboBox1.Value <> "" And ComboBox2.Value <> "" And ComboBox3.Value <> "" Then
            Dim opmin2, opmin3 As String
            Dim opmin4 As Integer
            opmin2 = ComboBox1.Value
            opmin3 = ComboBox2.Value
            opmin4 = ComboBox3.Value
            SQL = "UPDATE NCRDB SET DISP = '" & opmin2 & "', TAMIRTURU = '" & opmin3 & "', KATSAYISI = '" & opmin4 & "', WHOM = '" & UserName & "', LOGDATE = now() WHERE [ID] = " & id & " "
    Else
    MsgBox "Bir tamir tipi seçiniz ve tüm alanlarý doldurunuz"
    Exit Sub
    End If

wstracker.Unprotect Password:="4135911"

Dim cnn As ADODB.Connection 'dim the ADO collection class
Dim rst As ADODB.Recordset 'dim the ADO recordset class
Dim dbPath
Dim x As Long, i As Long
dbPath = wstracker.Cells(1, 9)
Set cnn = New ADODB.Connection
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=18811938;"
Set rst = New ADODB.Recordset
rst.Open Source:="NCRDB", ActiveConnection:=cnn, _
CursorType:=adOpenDynamic, LockType:=adLockOptimistic, _
Options:=adCmdTable
cnn.Execute SQL
SQL2 = "INSERT INTO DBLOG SELECT [ID], [ORACLE NUMBER], [STATUS], [RR], [SCOP], [SCMIN], [LAMOP], [LAMMIN], [QRESP], [LOGDATE], [WHOM] FROM NCRDB WHERE [ID] = " & id & " "
cnn.Execute SQL2
cnn.Close
Set rs = Nothing
Set cnn = Nothing

MsgBox " Tamir türü baþarýyla güncellendi"
Unload Me
Call refresh
Exit Sub
err:
    MsgBox "Ýnternete baðlý olduðunu kontrol et" & vbNewLine & "Hala olmuyorsa Sarper'e ekran görüntüsü at :) " & vbNewLine _
    & vbNewLine & "Hata kodu: " & err.Number & vbNewLine & _
    "Hata: " & err.Description, vbExclamation, "Sýkýntý var :("
    Exit Sub
End Sub

Private Sub UserForm_Initialize()
Dim wstracker As Worksheet
Set wstracker = ThisWorkbook.Worksheets("NCR TRACKER")
ComboBox1.List = Array("1 - Laminasyon UYGULANMAYACAK", "2 - Kat kaldýr ve kumaþ at", "3 - Delbas")
ComboBox3.List = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "20+")
If wstracker.Cells(idrow, 3).Value = "V136" Then
ComboBox2.List = Array("Kumaþ Atýlmayacak", "NC", "El yatýrmasý", "Vacuum Bag", "Ýnfüzyon", "Güçlendirme", "Nüve Yerleþtirme", "Delbas", "Dolgu", "Pencere Tamiri", "MSW Tamiri")
Else
ComboBox2.List = Array("Kumaþ Atýlmayacak", "SDR", "El yatýrmasý", "Vacuum Bag", "Ýnfüzyon", "Güçlendirme", "Nüve Yerleþtirme", "Delbas", "Dolgu", "Pencere Tamiri", "Splitline yarýlacak")
End If
End Sub

Private Sub UserForm_Terminate()
Call refresh
End Sub
