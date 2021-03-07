VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LProfit 
   Caption         =   "LAPORAN PROFIT HARIAN"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   Icon            =   "LProfit.dsx":0000
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   18785
   _ExtentY        =   15161
   SectionData     =   "LProfit.dsx":3482
End
Attribute VB_Name = "LProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tot_beli As Double
Dim tot_jual As Double
Dim TOT_profit As Double

Private Sub ActiveReport_ReportStart()
Field2.DataField = "simbol"
Field3.DataField = "jml_jual"
Field4.DataField = "rate_beli"
Field5.DataField = "tot_beli"
Field6.DataField = "rate_jual"
Field7.DataField = "tot_jual"
Field8.DataField = "profit"
tot_beli = 0
tot_jual = 0
TOT_profit = 0
End Sub

Private Sub Detail_Format()
tot_beli = tot_beli + Field5
tot_jual = tot_jual + Field7
TOT_profit = TOT_profit + Field8
End Sub

Private Sub GroupFooter1_Format()
Field9 = Format(tot_beli, "###,###.00")
Field10 = Format(tot_jual, "###,###.00")
Field11 = Format(TOT_profit, "###,###.00")
End Sub
