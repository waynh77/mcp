VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} INVOICE 
   Caption         =   "INVOICE"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "INVOICE.dsx":0000
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "INVOICE.dsx":1CCA
End
Attribute VB_Name = "INVOICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim no As Byte
Dim tot As Double

Private Sub ActiveReport_ReportStart()
If cetak_ulang = True Then
    Field9.DataField = "simbol"
    Field10.DataField = "jumlah"
    Field11.DataField = "satuan_rp"
    Field12.DataField = "total_rp"
Else
    Field9.DataField = "curr"
    Field10.DataField = "jml"
    Field11.DataField = "satuan_rp"
    Field12.DataField = "total"
End If
    no = 0
    tot = 0
End Sub

Private Sub ActiveReport_Terminate()
If ctk_inv = True Then
    Unload Beli_frm
End If
End Sub

Private Sub Detail_Format()
no = no + 1
Field8 = no
tot = tot + Val(Format(Field12, "###.##"))
End Sub

Private Sub GroupFooter1_Format()
Field7.Text = Format(tot, "###,###.00")
Label64.Caption = Mid(MoneyChanger.Label1.Caption, 13)
End Sub
