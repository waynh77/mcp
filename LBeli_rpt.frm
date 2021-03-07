VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form LBeli_rpt 
   Caption         =   "LAPORAN TRANSAKSI PENJUALAN DAN PEMBELIAN"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "LBeli_rpt.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      CausesValidation=   0   'False
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "LBeli_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New LBeli_CR85

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
With LBeli_frm
If .Combo1.ListIndex = 0 And .Combo2.ListIndex = 0 Then
    If .Combo3.ListIndex = 0 Then
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ")"
    Else
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.simbol}='" & .Combo3 & "'"
    End If
ElseIf .Combo1.ListIndex = 0 And .Combo2.ListIndex = 1 Then
    If .Combo3.ListIndex = 0 Then
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,m,d") & ")"
    Else
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,m,d") & ") and {trans_jualbeli.simbol}='" & .Combo3 & "'"
    End If
ElseIf .Combo1.ListIndex = 1 And .Combo2.ListIndex = 0 Then
    If .Combo3.ListIndex = 0 Then
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.status}='BELI'"
    Else
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.simbol}='" & .Combo3 & "' and {trans_jualbeli.status}='BELI'"
    End If
ElseIf .Combo1.ListIndex = 1 And .Combo2.ListIndex = 1 Then
    If .Combo3.ListIndex = 0 Then
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,m,d") & ") and {trans_jualbeli.status}='BELI'"
    Else
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,m,d") & ") and {trans_jualbeli.simbol}='" & .Combo3 & "' and {trans_jualbeli.status}='BELI'"
    End If
ElseIf .Combo1.ListIndex = 2 And .Combo2.ListIndex = 0 Then
    If .Combo3.ListIndex = 0 Then
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.status}='JUAL'"
    Else
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.simbol}='" & .Combo3 & "' and {trans_jualbeli.status}='JUAL'"
    End If
ElseIf .Combo1.ListIndex = 2 And .Combo2.ListIndex = 1 Then
    If .Combo3.ListIndex = 0 Then
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,m,d") & ") and {trans_jualbeli.status}='JUAL'"
    Else
        Report.RecordSelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,m,d") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,m,d") & ") and {trans_jualbeli.simbol}='" & .Combo3 & "' and {trans_jualbeli.status}='JUAL'"
    End If
End If
If .Combo1.ListIndex = 0 Then
    Report.ReportTitle = "LAPORAN PENJUALAN DAN PEMBELIAN"
ElseIf .Combo1.ListIndex = 1 Then
    Report.ReportTitle = "LAPORAN PEMBELIAN"
Else
    Report.ReportTitle = "LAPORAN PENJUALAN"
End If
End With
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
