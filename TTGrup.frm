VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TTGrup 
   Caption         =   "Tamir Tekrar� Grupland�rma"
   ClientHeight    =   1560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4020
   OleObjectBlob   =   "TTGrup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TTGrup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
If Not ComboBox1.Value = "" Then
ttgrupcombo = ComboBox1.Value
Unload Me
Else
MsgBox "Tamir tekrar� nedenini se�iniz"
End If

End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Laminasyonda Hava", "Kuruluk", "Delikler aras� mesafe", "Kuma� hasar�", _
"Dispozisyon", "FOD", "Hata giderilmedi", "Delaminasyon", "Egzoterm", "Yap��mama", _
"Bindirme hatas�", "Yetersiz ta�lama", "Kuma� kaymas�", "D�k�mantasyon")
End Sub

Private Sub UserForm_Terminate()
Call refresh
End Sub
