VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LBeli_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "LAPORAN PEMBELIAN VALAS"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "LBeli_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11400
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   7320
      TabIndex        =   12
      Text            =   "Combo3"
      Top             =   2520
      Width           =   3015
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "LBeli_frm.frx":3482
      Height          =   4695
      Left            =   2520
      TabIndex        =   10
      Top             =   3720
      Width           =   10575
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   18653
      _ExtentY        =   8281
      _StockProps     =   79
      Caption         =   "DATA LAPORAN PEMBELIAN"
      ForeColor       =   -2147483634
      BackColor       =   8421504
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES CARI DATA"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   3000
      Width           =   5535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   1560
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Index           =   0
      Left            =   7320
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   60293121
      CurrentDate     =   39864
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14520
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":3496
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":4E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":67BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":814C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":9ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":B470
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":C14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LBeli_frm.frx":CE24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   11010
      Left            =   14430
      TabIndex        =   0
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   19420
      ButtonWidth     =   1191
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Tambah"
            Object.ToolTipText     =   "Tambah Data Baru"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Hapus"
            Object.ToolTipText     =   "Hapus Data"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cari"
            Object.ToolTipText     =   "Cari Data"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cetak"
            Object.ToolTipText     =   "Cetak Data Ke Printer"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Tutup"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "LBeli_frm.frx":E7B6
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Index           =   1
      Left            =   8880
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   60293121
      CurrentDate     =   39864
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PERIODE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   3
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CURRENCY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   2
      Left            =   4920
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS LAPORAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LAPORAN PEMBELIAN DAN PENJUALAN VALAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "LBeli_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Change()
If Combo2.ListIndex = 0 Then
    DTPicker1(1).Visible = False
Else
    DTPicker1(1).Visible = True
End If
End Sub

Private Sub Combo2_Click()
If Combo2.ListIndex = 0 Then
    DTPicker1(1).Visible = False
Else
    DTPicker1(1).Visible = True
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 Then
    If Combo3.ListIndex = 0 Then
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)='" & DTPicker1(0) & "'"
    Else
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)='" & DTPicker1(0) & "' and simbol='" & Combo3 & "'"
    End If
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 0 Then
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)>=#" & DTPicker1(0) & "# and cdate(tgl)<=#" & DTPicker1(1) & "#"
    Else
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)>=#" & DTPicker1(0) & "# and cdate(tgl)<=#" & DTPicker1(1) & "# and simbol='" & Combo3 & "'"
    End If
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 Then
    If Combo3.ListIndex = 0 Then
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)='" & DTPicker1(0) & "' and status='BELI'"
    Else
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)='" & DTPicker1(0) & "' and simbol='" & Combo3 & "' and status='BELI'"
    End If
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 0 Then
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)>=#" & DTPicker1(0) & "# and cdate(tgl)<=#" & DTPicker1(1) & "#  and status='BELI'"
    Else
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)>=#" & DTPicker1(0) & "# and cdate(tgl)<=#" & DTPicker1(1) & "# and simbol='" & Combo3 & "' and status='BELI'"
    End If
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 Then
    If Combo3.ListIndex = 0 Then
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)='" & DTPicker1(0) & "' and status='JUAL'"
    Else
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)='" & DTPicker1(0) & "' and simbol='" & Combo3 & "' and status='JUAL'"
    End If
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 0 Then
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)>=#" & DTPicker1(0) & "# and cdate(tgl)<=#" & DTPicker1(1) & "#  and status='JUAL'"
    Else
        Data1.RecordSource = "select * from trans_jualbeli where cdate(tgl)>=#" & DTPicker1(0) & "# and cdate(tgl)<=#" & DTPicker1(1) & "# and simbol='" & Combo3 & "' and status='JUAL'"
    End If
End If
Data1.Refresh
End Sub

Private Sub Data1_Reposition()
cek_data
End Sub

Sub cek_data()
If Data1.Recordset.BOF Then
    Toolbar1.Buttons(5).Enabled = False
Else
    Toolbar1.Buttons(5).Enabled = True
End If
End Sub

Private Sub Form_Activate()
Call db_lbeli
isi_cmb1
ISI_cmb2
isi_cmb3
'Command1_Click
End Sub

Private Sub Form_Load()
DTPicker1(0) = Date
DTPicker1(1) = Date
Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 5
'    LPembelian.Show
CrystalReport1.ReportFileName = App.Path & "\laporan penjualan dan pembelian.rpt"
If Combo1.ListIndex = 0 And Combo2.ListIndex = 0 Then
    If Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.void}=false"
    Else
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.simbol}='" & Combo3 & "' and {trans_jualbeli.void}=false"
    End If
ElseIf Combo1.ListIndex = 0 And Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.void}=false"
'        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Year(DTPicker1(0)) & "," & Month(DTPicker1(0)) & Day(DTPicker1(0)) & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.void}=false"
    Else
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.simbol}='" & Combo3 & "' and {trans_jualbeli.void}=false"
    End If
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 0 Then
    If Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.status}='BELI' and {trans_jualbeli.void}=false"
    Else
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.simbol}='" & Combo3 & "' and {trans_jualbeli.status}='BELI' and {trans_jualbeli.void}=false"
    End If
ElseIf Combo1.ListIndex = 1 And Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.status}='BELI' and {trans_jualbeli.void}=false"
    Else
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.simbol}='" & Combo3 & "' and {trans_jualbeli.status}='BELI' and {trans_jualbeli.void}=false"
    End If
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 0 Then
    If Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.status}='JUAL' and {trans_jualbeli.void}=false"
    Else
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.simbol}='" & Combo3 & "' and {trans_jualbeli.status}='JUAL' and {trans_jualbeli.void}=false"
    End If
ElseIf Combo1.ListIndex = 2 And Combo2.ListIndex = 1 Then
    If Combo3.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.status}='JUAL' and {trans_jualbeli.void}=false"
    Else
        CrystalReport1.SelectionFormula = "{trans_jualbeli.tgl}>= date(" & Format(LBeli_frm.DTPicker1(0), "yyyy,mm,dd") & ") and {trans_jualbeli.tgl}<= date(" & Format(LBeli_frm.DTPicker1(1), "yyyy,mm,dd") & ") and {trans_jualbeli.simbol}='" & Combo3 & "' and {trans_jualbeli.status}='JUAL' and {trans_jualbeli.void}=false"
    End If
End If
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
'    LBeli_rpt.Show
Case 7
    Unload Me
End Select
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "All..."
Combo1.AddItem "Pembelian"
Combo1.AddItem "Penjualan"
Combo1.ListIndex = 0
End Sub

Sub ISI_cmb2()
Combo2.Clear
Combo2.AddItem "Harian"
Combo2.AddItem "Periodik"
Combo2.ListIndex = 0
End Sub

Sub isi_cmb3()
Combo3.Clear
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Combo3.AddItem "All..."
    Do While Not .EOF
        Combo3.AddItem !simbol
        .MoveNext
    Loop
    Combo3.ListIndex = 0
End If
End With
End Sub
