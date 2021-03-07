VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Gaji_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Slip Gaji Pegawai"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Gaji_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data9 
      Caption         =   "kas harian"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data8 
      Caption         =   "bb_bulanan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data7 
      Caption         =   "Rekap Jurnal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data6 
      Caption         =   "Trans bonus"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data5 
      Caption         =   "saldo bonus"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data4 
      Caption         =   "trans pinjaman"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data3 
      Caption         =   "saldo pinjaman"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   11280
      TabIndex        =   5
      Text            =   "Text8"
      Top             =   7320
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11280
      TabIndex        =   42
      Text            =   "Text7"
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   11280
      TabIndex        =   4
      Text            =   "Text6"
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   300
      Left            =   11280
      TabIndex        =   41
      Text            =   "Text2"
      Top             =   7680
      Width           =   2775
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   11280
      TabIndex        =   2
      Text            =   "Combo3"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Data data1 
      Caption         =   "tabel gaji"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   11280
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11280
      TabIndex        =   35
      Text            =   "Text5"
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   11280
      TabIndex        =   33
      Text            =   "Text4"
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   11280
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   11280
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   11280
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   11280
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   11280
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   11280
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   5
      Left            =   11280
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   6
      Left            =   11280
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   7
      Left            =   11280
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   10
      Left            =   11280
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   8
      Left            =   11280
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   9
      Left            =   11280
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Data data2 
      BackColor       =   &H00000000&
      Caption         =   "DATA SLIP GAJI PEGAWAI"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Width           =   5175
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
            Picture         =   "Gaji_frm.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":4E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":67A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":8138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":9ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":B45C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":C136
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Gaji_frm.frx":CE10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   11010
      Left            =   14430
      TabIndex        =   6
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
            Caption         =   "Tambah"
            Object.ToolTipText     =   "Tambah Data Baru"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "Edit Data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.Visible         =   0   'False
            Caption         =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      MouseIcon       =   "Gaji_frm.frx":E7A2
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid2 
      Bindings        =   "Gaji_frm.frx":EABC
      Height          =   7575
      Left            =   120
      TabIndex        =   36
      Top             =   840
      Width           =   8535
      _Version        =   196614
      BevelColorFace  =   32768
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   15055
      _ExtentY        =   13361
      _StockProps     =   79
      Caption         =   "Transaksi Gaji Pegawai"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BAYAR BONUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Index           =   19
      Left            =   8880
      TabIndex        =   40
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BAYAR PINJAMAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Index           =   18
      Left            =   8880
      TabIndex        =   39
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HUTANG BONUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Index           =   17
      Left            =   8880
      TabIndex        =   38
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL PINJAMAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Index           =   16
      Left            =   8880
      TabIndex        =   37
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL GAJI DITERIMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Index           =   15
      Left            =   8880
      TabIndex        =   34
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL LEMBUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   14
      Left            =   8880
      TabIndex        =   32
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH LEMBUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   13
      Left            =   8880
      TabIndex        =   31
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TAHUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   12
      Left            =   8880
      TabIndex        =   30
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BULAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Index           =   11
      Left            =   8880
      TabIndex        =   29
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR PEGAWAI"
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
      Left            =   8880
      TabIndex        =   28
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PEGAWAI"
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
      Left            =   8880
      TabIndex        =   27
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GAJI POKOK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   2
      Left            =   8880
      TabIndex        =   26
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TUNJANGAN MAKAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   3
      Left            =   8880
      TabIndex        =   25
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TUNJ. TRANSPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   4
      Left            =   8880
      TabIndex        =   24
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TUNJANGAN TELP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   5
      Left            =   8880
      TabIndex        =   23
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UPAH LEMBUR/JAM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   6
      Left            =   8880
      TabIndex        =   22
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIVISI"
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
      Index           =   8
      Left            =   8880
      TabIndex        =   21
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JABATAN"
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
      Index           =   9
      Left            =   8880
      TabIndex        =   20
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TUNJ. KESEHATAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   7
      Left            =   8880
      TabIndex        =   19
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TUNJ. PERUMAHAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   10
      Left            =   8880
      TabIndex        =   18
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SLIP GAJI PEGAWAI"
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
      TabIndex        =   7
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "Gaji_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cek As Boolean
Dim tambah As Boolean
Dim cek_dat As Boolean
Dim lembur_awal As Double
Dim pinjaman_awal As Double
Dim bonus_awal As Double
Dim gaji_awal As Double

Private Sub Combo1_Click()
isi_grid
End Sub

Private Sub Combo2_Click()
isi_grid
End Sub

Private Sub Combo3_Click()
isi
isi_grid
isi_pinjaman
isi_bonus
End Sub

Private Sub Data2_Reposition()
isi_data
End Sub

Private Sub Form_Activate()
data1.Refresh
isi_cmb3
isi
isi_data
End Sub

Private Sub Form_Load()
Call db_Gaji
Tutup
Kosong
kosong_Trans
cmd_awal
isi_cmb1
ISI_cmb2
End Sub

Private Sub SSDBGrid1_Click()
isi
End Sub

Private Sub Text3_Change()
isi_ttlLembur
ttl_gaji
End Sub

Private Sub Text6_Change()
ttl_gaji
End Sub

Private Sub Text8_Change()
ttl_gaji
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Tambah" Then
        Buka
        tambah = True
        cmd_Simpan
        kosong_Trans
        Text3.SetFocus
        lembur_awal = 0
        pinjam_awal = 0
        bonus_awal = 0
        gaji_awal = 0
    Else
        simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        If Combo3 <> "" Then
            Buka
            cmd_Simpan
            tambah = False
            Text3.SetFocus
            lembur_awal = Val(Format(Text4, "###.00"))
            pinjam_awal = Val(Format(Text6, "###.00"))
            bonus_awal = Val(Format(Text8, "###.00"))
            gaji_awal = Val(Format(Text2, "###.00"))
        Else
            MsgBox "Data Kosong", vbInformation, "Validasi Data"
        End If
    Else
        cmd_awal
        Tutup
        isi
    End If
Case 3
Case 4
Case 5
    cetak_bukti
Case 6
Case 7
    Unload Me
End Select
End Sub

Sub Kosong()
'Combo3 = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Text1(4) = ""
Text1(5) = ""
Text1(6) = ""
Text1(7) = ""
Text1(8) = ""
Text1(9) = ""
Text1(10) = ""
End Sub

Sub kosong_Trans()
'Text2 = ""
Text3 = ""
'Text4 = ""
'Text5 = ""
Text6 = ""
'Text7 = ""
Text8 = ""
End Sub

Sub isi()
If data1.Enabled = True Then
Kosong
With data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until Combo3 = !no_peg Or .EOF
            .MoveNext
        Loop
'        Combo3 = !no_peg
        Text1(1) = !nama
        Text1(2) = !Div
        Text1(3) = !Jab
        Text1(4) = Format(!pokok, "###,###.00")
        Text1(5) = Format(!makan, "###,###.00")
        Text1(6) = Format(!transport, "###,###.00")
        Text1(7) = Format(!telp, "###,###.00")
        Text1(8) = Format(!kesehatan, "###,###.00")
        Text1(9) = Format(!perumahan, "###,###.00")
        Text1(10) = Format(!lembur, "###,###.00")
    End If
End With
End If
End Sub

Sub isi_cmb1()
Combo1.Clear
Combo1.AddItem "Januari"
Combo1.AddItem "Februari"
Combo1.AddItem "Maret"
Combo1.AddItem "April"
Combo1.AddItem "Mei"
Combo1.AddItem "Juni"
Combo1.AddItem "Juli"
Combo1.AddItem "Agustus"
Combo1.AddItem "September"
Combo1.AddItem "Oktober"
Combo1.AddItem "November"
Combo1.AddItem "Desember"
Combo1.ListIndex = 0
End Sub

Sub ISI_cmb2()
Data8.RecordSource = "select tahun from bb_bulanan group by tahun order by tahun desc"
Data8.Refresh
Combo2.Clear
With Data8.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo2.AddItem !tahun
        .MoveNext
    Loop
    Combo2.ListIndex = 0
Else
    Combo2 = ""
End If
End With
End Sub

Sub isi_cmb3()
data1.Enabled = False
data1.Refresh
Combo3.Clear
With data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo3.AddItem !no_peg
        .MoveNext
    Loop
    Combo3.ListIndex = 0
Else
    Combo3 = ""
End If
End With
data1.Refresh
data1.Enabled = True
End Sub

Sub isi_grid()
If Combo1 <> "Combo1" And Combo2 <> "Combo2" Then
data2.RecordSource = "select * from trans_gaji where bulan=" & Combo1.ListIndex + 1 & " and tahun=" & Combo2 '& "'"
data2.Refresh
End If
End Sub

Sub isi_pinjaman()
Data3.RecordSource = "select * from saldo_pinjaman where no_peg='" & Combo3 & "'"
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    Text5 = Format(!saldo, "###,###.00")
Else
    Text5 = 0
End If
End With
End Sub

Sub isi_bonus()
Data5.RecordSource = "select * from saldo_bonus where no_peg='" & Combo3 & "'"
Data5.Refresh
With Data5.Recordset
If Not .BOF Then
    Text7 = Format(!saldo, "###,###.00")
Else
    Text7 = 0
End If
End With
End Sub

Sub isi_ttlLembur()
Text4 = Format(Val(Format(Text1(10), "###.00")) * Val(Format(Text3, "###")), "###,###.00")
End Sub

Sub isi_data()
With data2.Recordset
If data2.Enabled = True And Not .BOF Then
    Combo3 = !no_peg
    Text1(1) = !nama
    Text1(2) = !Div
    Text1(3) = !Jab
    Text3 = !jml_lembur
    Text6 = Format(!pinjaman, "###,###.00")
    Text8 = Format(!bonus, "###,###.00")
    Text1(4) = Format(!pokok, "###,###.00")
    Text1(5) = Format(!makan, "###,###.00")
    Text1(6) = Format(!transport, "###,###.00")
    Text1(7) = Format(!telp, "###,###.00")
    Text1(8) = Format(!kesehatan, "###,###.00")
    Text1(9) = Format(!perumahan, "###,###.00")
    Text1(10) = Format(!lembur, "###,###.00")
End If
End With
isi_ttlLembur
ttl_gaji
End Sub

Sub Tutup()
'    Combo3.Enabled = False
    Text1(1).Enabled = False
    Text1(2).Enabled = False
    Text1(3).Enabled = False
    Text1(4).Enabled = False
    Text1(5).Enabled = False
    Text1(6).Enabled = False
    Text1(7).Enabled = False
    Text1(8).Enabled = False
    Text1(9).Enabled = False
    Text1(10).Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
End Sub

Sub Buka()
Text3.Enabled = True
Text6.Enabled = True
Text8.Enabled = True
Text3 = Format(Text3, "###")
Text6 = Format(Text6, "###")
Text8 = Format(Text8, "###")
End Sub

Sub simpan()
Cek_Input
If cek = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With data2.Recordset
        If tambah = True Then
            cek_data
            If cek_dat = False Then
                cetak_bukti
                update_pinjaman
                update_bonus
                UPDATE_JURNAL
                edit_BB
                Update_KAs
                .AddNew
                !tgl = Date
                !BULAN = Combo1.ListIndex + 1
                !tahun = Combo2
                !no_peg = Combo3
                !nama = Text1(1)
                !Div = Text1(2)
                !Jab = Text1(3)
                !pokok = Val(Format(Text1(4), "###.00"))
                !makan = Val(Format(Text1(5), "###.00"))
                !transport = Val(Format(Text1(6), "###.00"))
                !telp = Val(Format(Text1(7), "###.00"))
                !kesehatan = Val(Format(Text1(8), "###.00"))
                !perumahan = Val(Format(Text1(9), "###.00"))
                !lembur = Val(Format(Text1(10), "###.00"))
                !jml_lembur = Val(Format(Text3, "###.00"))
                !ttl_lembur = Val(Format(Text4, "###.00"))
                !pinjaman = Val(Format(Text6, "###.00"))
                !bonus = Val(Format(Text8, "###.00"))
                .Update
                Tutup
                cmd_awal
            Else
                MsgBox "Maaf data sudah ada... silahkan masukan yang lain", vbInformation, "Validasi Data"
            End If
        Else
            edit_pinjamBonus
            UPDATE_JURNAL
            edit_BB
            Update_KAs
            .Edit
            !pokok = Val(Format(Text1(4), "###.00"))
            !makan = Val(Format(Text1(5), "###.00"))
            !transport = Val(Format(Text1(6), "###.00"))
            !telp = Val(Format(Text1(7), "###.00"))
            !kesehatan = Val(Format(Text1(8), "###.00"))
            !perumahan = Val(Format(Text1(9), "###.00"))
            !lembur = Val(Format(Text1(10), "###.00"))
            !jml_lembur = Val(Format(Text3, "###.00"))
            !ttl_lembur = Val(Format(Text4, "###.00"))
            !pinjaman = Val(Format(Text6, "###.00"))
            !bonus = Val(Format(Text8, "###.00"))
            .Update
            Tutup
            cmd_awal
        End If
    End With
    data2.Refresh
End If
data2.Refresh
End Sub

Sub Cek_Input()
cek = False
If Text1(4) = "" Or Text1(5) = "" Or Text1(6) = "" Or Text1(7) = "" Or Text1(8) = "" Or Text1(9) = "" Or Text1(10) = "" Then
    cek = False
Else
    cek = True
End If
End Sub

Sub cek_data()
cek_dat = False
With data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo3 = !no_peg Then
            cek_dat = True
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub cmd_awal()
With Toolbar1
    .Buttons(1).Image = 1
    .Buttons(2).Image = 2
    .Buttons(1).Caption = "Tambah"
    .Buttons(2).Caption = "Edit"
    .Buttons(1).ToolTipText = "Tambah Data"
    .Buttons(2).ToolTipText = "Edit Data"
    .Buttons(3).Visible = True
    .Buttons(4).Visible = False
    .Buttons(5).Visible = True
    .Buttons(6).Visible = False
    .Buttons(7).Visible = True
End With
data2.Enabled = True
SSDBGrid2.Enabled = True
End Sub

Sub cmd_Simpan()
With Toolbar1
    .Buttons(1).Image = 8
    .Buttons(2).Image = 3
    .Buttons(1).Caption = "Simpan"
    .Buttons(2).Caption = "Batal"
    .Buttons(1).ToolTipText = "Simpan Data"
    .Buttons(2).ToolTipText = "Batal Data"
    .Buttons(3).Visible = False
    .Buttons(4).Visible = False
    .Buttons(5).Visible = False
    .Buttons(6).Visible = False
    .Buttons(7).Visible = False
End With
data2.Enabled = False
SSDBGrid2.Enabled = False
End Sub

Sub ttl_gaji()
Dim ttl As Double
ttl = Val(Format(Text1(4), "###.00")) + Val(Format(Text1(5), "###.00")) + Val(Format(Text1(6), "###.00")) + Val(Format(Text1(7), "###.00")) + Val(Format(Text1(8), "###.00")) + Val(Format(Text1(9), "###.00")) + Val(Format(Text4, "###.00")) - Val(Format(Text6, "###.00")) + Val(Format(Text8, "###.00"))
Text2 = Format(ttl, "###,###.00")
End Sub

Sub update_pinjaman()
With Data3.Recordset
    .Edit
    !saldo = !saldo - Val(Format(Text6, "###.00"))
    .Update
End With

With Data4.Recordset
    .AddNew
    !no_peg = Combo3
    !tgl = Date
    !ket = "BAYAR PINJAMAN"
    !jml = Val(Format(Text6, "###.00"))
    .Update
End With
End Sub

Sub update_bonus()
With Data5.Recordset
If Not .BOF Then
    .Edit
    !saldo = !saldo - Val(Format(Text8, "###.00"))
    .Update
End If
End With

With Data6.Recordset
    .AddNew
    !no_peg = Combo3
    !tgl = Date
    !ket = "BAYAR BONUS"
    !jml = Val(Format(Text8, "###.00"))
    .Update
End With
End Sub

Sub edit_pinjamBonus()
Dim sel_pinjam As Double
Dim sel_bonus As Double

sel_pinjam = pinjam_awal - Val(Format(Text6, "###.00"))
With Data3.Recordset
    .Edit
    !saldo = !saldo - sel_pinjam
    .Update
End With

With Data4.Recordset
    .AddNew
    !no_peg = Combo3
    !tgl = Date
    !ket = "BAYAR PINJAMAN"
    !jml = sel_pinjam
    .Update
End With

sel_bonus = bonus_awal - Val(Format(Text8, "###.00"))
Data5.Refresh
With Data5.Recordset
    .Edit
    !saldo = !saldo + sel_bonus
    .Update
End With

With Data6.Recordset
    .AddNew
    !no_peg = Combo3
    !tgl = Date
    !ket = "BAYAR BONUS"
    !jml = sel_bonus
    .Update
End With

End Sub

Sub UPDATE_JURNAL()
'jurnal gaji
With Data7.Recordset
    .AddNew
    !tgl = Date
    !jam = Time
    !no_akun = "1-110"
    !dk = "KREDIT"
    If tambah = True Then
        !ket = "BAYAR GAJI " & Text1(1)
        !jml = Val(Format(Text2, "###.00"))
    Else
        !ket = "REVISI BAYAR GAJI " & Text1(1)
        !jml = gaji_awal - Val(Format(Text2, "###.00"))
    End If
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = "6-011"
    .Update

    .AddNew
    !tgl = Date
    !jam = Time
    !no_akun = "6-011"
    !dk = "DEBET"
    If tambah = True Then
        !ket = "BAYAR GAJI " & Text1(1)
        !jml = Val(Format(Text2, "###.00"))
    Else
        !ket = "REVISI BAYAR GAJI " & Text1(1)
        !jml = gaji_awal - Val(Format(Text2, "###.00"))
    End If
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = "1-110"
    .Update
End With

End Sub

Sub edit_BB()
Dim sal As Double
Dim sel As Double
Dim dk As String
sel = 0
sel = Val(Format(Text2, "###")) - gaji_awal

'bb kas
Data8.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
Data8.Refresh
With Data8.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal - sel + Val(Format(Text6, "###"))
    .Update
End With

'bb gaji
Data8.RecordSource = "select * from bb_bulanan where no_akun='6-011' order by tahun desc,bulan desc"
Data8.Refresh
With Data8.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal + sel
    .Update
End With
Data8.Refresh

'bb pinjaman
Data8.RecordSource = "select * from bb_bulanan where no_akun='1-131' order by tahun desc,bulan desc"
Data8.Refresh
With Data8.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal - Val(Format(Text6, "###.00"))
    .Update
End With
Data8.Refresh

'bb bonus
Data8.RecordSource = "select * from bb_bulanan where no_akun='2-120' order by tahun desc,bulan desc"
Data8.Refresh
With Data8.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal - Val(Format(Text8, "###.00"))
    .Update
End With
Data8.Refresh

'update laba rugi
Data8.RecordSource = "select * from bb_bulanan where no_akun='3-200' order by tahun desc,bulan desc"
Data8.Refresh
With Data8.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal - Val(Format(Text2, "###.00"))
    .Update
End With

End Sub

Sub Update_KAs()
'update kas harian
Data9.RecordSource = "select * from kas_harian where cdate(tgl)>='" & DTPicker1 & "'"
Data9.Refresh
With Data9.Recordset
If Not .BOF Then
    .MoveFirst
    If tambah = True Then
        .Edit
        !saldo = !saldo - Val(Format(Text2, "###.00"))
        .Update
    Else
        Do While Not .EOF
            .Edit
            !saldo = !saldo - (Val(Format(Text2, "###.00")) - gaji_awal)
            .Update
            .MoveNext
        Loop
    End If
End If
End With

End Sub

Sub cetak_bukti()
With Bukti_Gaji
    .Field1 = Format(Date, "d mmm yyyy")
    .Field2 = Combo3
    .Field3 = Text1(1)
    .Field4 = Text1(2)
    .Field5 = Text1(3)
    .Field6 = Combo1 & " " & Combo2
    .Field7 = Format(Text1(4), "###,###.00")
    .Field8 = Format(Text1(5), "###,###.00")
    .Field9 = Format(Text1(6), "###,###.00")
    .Field10 = Format(Text1(7), "###,###.00")
    .Field11 = Format(Text1(8), "###,###.00")
    .Field12 = Format(Text1(9), "###,###.00")
    .Field13 = Format(Text1(10), "###,###.00")
    .Field14 = Format(Text3, "###")
    .Field15 = Format(Text4, "###,###.00")
    .Field16 = Format(Text5, "###,###.00")
    .Field17 = Format(Text6, "###,###.00")
    .Field18 = Format(Text7, "###,###.00")
    .Field19 = Format(Text8, "###,###.00")
    .Field20 = Format(Text2, "###,###.00")
    .Label80 = Text1(1)
    .Show
    .WindowState = 2
End With
End Sub

