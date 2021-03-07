VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Data_Admin 
   BackColor       =   &H00FF8080&
   Caption         =   "DATABASE MANAGEMENT"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "Data_Admin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   1440
      Top             =   720
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Data_Admin.frx":3482
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   14175
      _Version        =   196614
      BevelColorFace  =   192
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      RowSelectionStyle=   2
      SelectTypeCol   =   0
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   25003
      _ExtentY        =   12091
      _StockProps     =   79
      Caption         =   "SSDBGrid1"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   3375
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
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
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
            Picture         =   "Data_Admin.frx":3496
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":4E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":67BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":814C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":9ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":B470
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":C14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Data_Admin.frx":CE24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   11010
      Left            =   14430
      TabIndex        =   4
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
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DEL ALL"
            Object.ToolTipText     =   "Hapus Semua Data"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CLOSE"
            Object.ToolTipText     =   "Tutup"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Data_Admin.frx":E7B6
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   8280
      Width           =   11160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECT DATABASE"
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
      Left            =   11520
      TabIndex        =   5
      Top             =   8280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECT DATABASE"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MONEY CHANGER DATABASE MANAGEMENT"
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
      TabIndex        =   0
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "Data_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Boolean

Private Sub Combo1_Change()
isi_grid
Label1(1).Caption = ""
Label1(1).Caption = "TOTAL DATA : " & Data1.Recordset.RecordCount
End Sub

Private Sub Combo1_Click()
isi_grid
Label1(1).Caption = ""
Label1(1).Caption = "TOTAL DATA : " & Data1.Recordset.RecordCount
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
Call dB_Admin
ISI_cmb
End Sub

Sub ISI_cmb()
Combo1.Clear
Combo1.AddItem "REKAP JURNAL"
Combo1.AddItem "BUKU BESAR"
Combo1.AddItem "TRANSAKSI JUAL BELI"
Combo1.AddItem "TRANSAKSI HUTANG PIUTANG"
Combo1.AddItem "STOK BULANAN"
Combo1.AddItem "PERSEDIAAN"
Combo1.AddItem "STOK HARIAN"
Combo1.AddItem "KAS HARIAN"
Combo1.AddItem "SALDO PINJAMAN"
Combo1.AddItem "SALDO BONUS"
Combo1.AddItem "TRANSAKSI PINJAMAN"
Combo1.AddItem "TRANSAKSI BONUS"
Combo1.AddItem "TRANSAKSI GAJI"
Combo1.ListIndex = 0
End Sub

Private Sub Timer1_Timer()
Dim lbl As String
lbl = "PERHATIAN : Hati-hati dalam mengubah/menghapus data...!!!"
If t = True Then
    Label2.Caption = lbl
    t = False
Else
    Label2.Caption = ""
    t = True
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    With Data1.Recordset
        If Not .BOF Then
            X = MsgBox("Apakah anda yakin menghapus semua data...???", vbYesNo, "Hapus Semua Data")
            If X = vbYes Then
                .MoveFirst
                Do While Not .EOF
                    .Delete
                    .MoveNext
                Loop
                MsgBox "Data Telah Dihapus...", vbInformation, "Validasi Data"
            End If
        Else
            MsgBox "Maaf Data Masih Kosong...", vbInformation, "Validasi Data"
        End If
    End With
Case 2
    Unload Me
End Select
End Sub

Sub isi_grid()
SSDBGrid1.Caption = "DATA " & Combo1
Select Case Combo1.ListIndex
Case 0
    Data1.RecordSource = "select * from rekap_jurnal order by tgl desc,jam desc,no_akun asc"
Case 1
    Data1.RecordSource = "select * from bb_bulanan order by tahun desc,bulan desc,no_akun asc"
Case 2
    Data1.RecordSource = "select * from trans_jualbeli order by tgl desc,jam desc,no_inv asc"
Case 3
    Data1.RecordSource = "select * from trans_hutangpiutang order by tgl desc,jam desc,no_inv asc"
Case 4
    Data1.RecordSource = "select * from STOK_BULANAN order by tahun desc,bulan desc,currency asc"
Case 5
    Data1.RecordSource = "select * from PERSEDIAAN order by simbol asc"
Case 6
    Data1.RecordSource = "select * from STOK_harian order by tgl desc,currency asc"
Case 7
    Data1.RecordSource = "select * from kas_harian order by tgl desc"
Case 8
    Data1.RecordSource = "select * from saldo_pinjaman"
Case 9
    Data1.RecordSource = "select * from saldo_bonus"
Case 10
    Data1.RecordSource = "select * from trans_pinjaman order by tgl desc"
Case 11
    Data1.RecordSource = "select * from trans_bonus order by tgl desc"
Case 12
    Data1.RecordSource = "select * from trans_gaji order by tgl desc"
End Select
Data1.Refresh
End Sub
