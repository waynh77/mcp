VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LABA_RUGI 
   Caption         =   "LAPORAN LABA RUGI"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "LABA_RUGI.dsx":0000
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   18177
   _ExtentY        =   13573
   SectionData     =   "LABA_RUGI.dsx":3482
End
Attribute VB_Name = "LABA_RUGI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim penjualan As Double
Dim persediaan1 As Double
Dim pembelian As Double
Dim persediaan2 As Double
Dim gaji As Double
Dim sewa As Double
Dim iklan As Double
Dim listrik As Double
Dim transport As Double
Dim penyusutan As Double
Dim lain As Double
Dim bunga As Double

penjualan = Val(Format(Field2, "###.00"))
persediaan1 = Val(Format(Field4, "###.00"))
pembelian = Val(Format(Field5, "###.00"))
persediaan2 = Val(Format(Field6, "###.00"))
gaji = Val(Format(Field7, "###.00"))
sewa = Val(Format(Field8, "###.00"))
iklan = Val(Format(Field9, "###.00"))
listrik = Val(Format(Field10, "###.00"))
transport = Val(Format(Field11, "###.00"))
penyusutan = Val(Format(Field12, "###.00"))
lain = Val(Format(Field13, "###.00"))
bunga = Val(Format(Field14, "###.00"))

Field3 = 0
Field20 = Field2
Field21 = Format(persediaan1 + pembelian - persediaan2, "###,###.00")
Field22 = Format(penjualan - (persediaan1 + pembelian - persediaan2), "###,###.00")
Field23 = Format(gaji + sewa + iklan + listrik + transport + penyusutan + lain, "###,###.00")
Field24 = Format(Val(Format(Field22, "###.00")) - Val(Format(Field23, "###.00")), "###,###.00")
Field25 = Field14
Field26 = Format(Val(Format(Field24, "###.00")) - Val(Format(Field25, "###.00")), "###,###.00")
Field28 = Format(Val(Format(Field26, "###.00")) - Val(Format(Field27, "###.00")), "###,###.00")

'update bb
With BB_frm
    .Data1.RecordSource = "select * from BB_BULANAN where no_akun='3-200' and bulan ='" & Lap_frm.Combo1 & "' and tahun='" & Lap_frm.Combo2 & "'"
    .Data1.Refresh
    If Not .Data1.Recordset.BOF Then
    .Data1.Recordset.Edit
    .Data1.Recordset!saldo = Val(Format(Field28, "###.00"))
    .Data1.Recordset.Update
    End If
End With
Unload BB_frm
End Sub

Private Sub PageHeader_Format()
'Field1 = Format(Date, "d mmmm yyyy")
End Sub
