VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Beli_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Transaksi Pembelian Valas"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15240
   Icon            =   "Beli_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data12 
      Caption         =   "Data12"
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
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data11 
      Caption         =   "Stok Bulanan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   10440
      TabIndex        =   35
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Data Data10 
      Caption         =   "Piutang"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data9 
      Caption         =   "Hutang"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data8 
      Caption         =   "Rekap Jurnal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data7 
      Caption         =   "Rekening Bank"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      Left            =   11160
      TabIndex        =   33
      Text            =   "Combo2"
      Top             =   8880
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      Left            =   11160
      TabIndex        =   31
      Text            =   "Combo2"
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Data Data6 
      Caption         =   "Buku Besar"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "HAPUS TRANSAKSI"
      Height          =   495
      Left            =   12360
      TabIndex        =   29
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Data Data5 
      Caption         =   "beli/jual"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA TRANSAKSI"
      Connect         =   "Access"
      DatabaseName    =   " dbmcp"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   " "
      Top             =   5880
      Width           =   4575
   End
   Begin VB.Data Data4 
      Caption         =   "Nasabah"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Stok"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Currency"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   1  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   10440
      TabIndex        =   27
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   0
      Left            =   10440
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   10440
      TabIndex        =   21
      Top             =   3360
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   10440
      TabIndex        =   20
      Top             =   1920
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   10440
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "EDIT  TRANSAKSI"
      Height          =   495
      Left            =   10800
      TabIndex        =   18
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TAMBAH TRANSAKSI"
      Height          =   495
      Left            =   9240
      TabIndex        =   17
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   5
      Left            =   11160
      TabIndex        =   14
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   4
      Left            =   11160
      TabIndex        =   13
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   3
      Left            =   11160
      TabIndex        =   12
      Top             =   4680
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   11160
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Recalculate"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Selected Row"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid2 
      Bindings        =   "Beli_frm.frx":3482
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   7695
      _Version        =   196614
      BevelColorFace  =   8388608
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2566
      Columns(0).Caption=   "CURRENCY"
      Columns(0).Name =   "CURRENCY"
      Columns(0).CaptionAlignment=   2
      Columns(0).DataField=   "Simbol"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3387
      Columns(1).Caption=   "STOCK"
      Columns(1).Name =   "STOCK"
      Columns(1).Alignment=   1
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Jumlah"
      Columns(1).DataType=   5
      Columns(1).NumberFormat=   "###"
      Columns(1).FieldLen=   256
      Columns(2).Width=   3466
      Columns(2).Caption=   "RATE (RP)"
      Columns(2).Name =   "RATE"
      Columns(2).Alignment=   1
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Rate_Rp"
      Columns(2).DataType=   5
      Columns(2).NumberFormat=   "###,###.00"
      Columns(2).FieldLen=   256
      Columns(3).Width=   3572
      Columns(3).Caption=   "TOTAL (RP)"
      Columns(3).Name =   "TOTAL (RP)"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Total_Rp"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "###,###.00"
      Columns(3).FieldLen=   256
      _ExtentX        =   13573
      _ExtentY        =   4895
      _StockProps     =   79
      Caption         =   "CURRENCY STOCK"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   11160
      TabIndex        =   3
      Top             =   7680
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14490
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
            Picture         =   "Beli_frm.frx":3496
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":4E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":67BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":814C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":9ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":B470
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":C14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Beli_frm.frx":CE24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   10710
      Left            =   14430
      TabIndex        =   1
      Top             =   0
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   18891
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Baru"
            Object.ToolTipText     =   "Transaksi Baru"
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
            Caption         =   "Cari"
            Object.ToolTipText     =   "Cari Data Transaksi"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Preview"
            Object.ToolTipText     =   "Simpan dan Cetak Transaksi"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Tutup"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Beli_frm.frx":E7B6
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Beli_frm.frx":EAD0
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   8535
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3175
      Columns(0).Caption=   "VALUTA ASING"
      Columns(0).Name =   "VALUTA ASING"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Curr"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   3
      Columns(0).Nullable=   0
      Columns(1).Width=   2223
      Columns(1).Caption=   "JUMLAH"
      Columns(1).Name =   "JUMLAH"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Jml"
      Columns(1).DataType=   4
      Columns(1).FieldLen=   256
      Columns(2).Width=   4128
      Columns(2).Caption=   "HARGA SATUAN"
      Columns(2).Name =   "HARGA SATUAN"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Satuan_Rp"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   4736
      Columns(3).Caption=   "EQUIVALENT RUPIAH"
      Columns(3).Name =   "EQUIVALENT RUPIAH"
      Columns(3).Alignment=   1
      Columns(3).DataField=   "Total"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).Locked=   -1  'True
      _ExtentX        =   15055
      _ExtentY        =   7435
      _StockProps     =   79
      Caption         =   "TRANSAKSI"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   8400
      X2              =   14160
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL KAS (RP)"
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
      Index           =   13
      Left            =   8520
      TabIndex        =   34
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REKENING"
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
      Index           =   12
      Left            =   9240
      TabIndex        =   32
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARA BAYAR"
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
      Index           =   11
      Left            =   9240
      TabIndex        =   30
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL STOK (RP)"
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
      Index           =   2
      Left            =   8520
      TabIndex        =   28
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   13560
      MouseIcon       =   "Beli_frm.frx":EAE4
      MousePointer    =   99  'Custom
      Picture         =   "Beli_frm.frx":EDEE
      Stretch         =   -1  'True
      ToolTipText     =   "Tambah Nasabah Baru"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA NASABAH"
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
      Left            =   8520
      TabIndex        =   26
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALAMAT"
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
      Index           =   0
      Left            =   8520
      TabIndex        =   25
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO. TELEPON"
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
      Index           =   1
      Left            =   8520
      TabIndex        =   24
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS NASABAH"
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
      Left            =   8520
      TabIndex        =   23
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   9240
      X2              =   13800
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL RP"
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
      Index           =   10
      Left            =   9240
      TabIndex        =   16
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   9240
      X2              =   13800
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSAKSI"
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
      Left            =   9240
      TabIndex        =   15
      Top             =   3840
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EQUIVALENT RP"
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
      Left            =   9240
      TabIndex        =   10
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SATUAN RP"
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
      Index           =   6
      Left            =   9240
      TabIndex        =   9
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH"
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
      Index           =   5
      Left            =   9240
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VALUTA ASING"
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
      Index           =   4
      Left            =   9240
      TabIndex        =   7
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSAKSI PEMBELIAN"
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
   Begin VB.Menu ctkulang 
      Caption         =   "Cetak Ulang Invoice"
   End
End
Attribute VB_Name = "Beli_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean
Dim cek2 As Boolean
Dim cek3 As Boolean
Dim cekstok As Boolean
Dim nomor_inv As String
Dim ulang As Boolean

Private Sub Combo1_Click(Index As Integer)
Select Case Index
Case 0
If Combo1(0).ListIndex = 0 Then
    Data4.RecordSource = "msnasabah_perseorangan"
Else
    Data4.RecordSource = "msnasabah_perusahaan"
End If
Data4.Refresh
ISI_cmb2
Case 1
isi_nasabah
End Select
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
'KeyAscii = 0
End Sub

Private Sub Combo2_Change(Index As Integer)
Select Case Index
Case 0
    If Combo2(0).ListIndex = 1 Then
        isi_Rek
    Else
        Combo2(1).Visible = False
        Label1(12).Visible = False
    End If
End Select
End Sub

Private Sub Combo2_Click(Index As Integer)
Select Case Index
Case 0
    If Combo2(0).ListIndex = 1 Then
        isi_Rek
    Else
        Combo2(1).Visible = False
        Label1(12).Visible = False
    End If
End Select
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command2_Click()
    SSDBGrid1.DeleteSelected
    SSDBGrid1.Refresh
    hitung
End Sub

Private Sub Command3_Click()
hitung
End Sub

Private Sub Command4_Click()
If Command4.Caption = "TAMBAH TRANSAKSI" Then
    If Data1.Recordset.RecordCount < 5 Then
        cmd_simpanTrans
        Data1.Enabled = False
        tambah = True
        kosong_Trans
        buka_trans
        isi_cmb3
        Combo1(2).SetFocus
    Else
        MsgBox "Maaf maksimal 5 transaksi", vbInformation, "Validasi Transaksi"
    End If
Else
    SIMPAN_TRANS
End If
End Sub

Private Sub Command5_Click()
If Command4.Caption = "TAMBAH TRANSAKSI" Then
    If Not Data1.Recordset.BOF Then
        buka_trans
        tambah = False
        isi_cmb3
        isi_trans
        cmd_simpanTrans
        Combo1(2).SetFocus
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
Else
    kosong_Trans
    cmd_awalTrans
    Data1.Refresh
    tutup_Trans
End If
End Sub

Private Sub Command6_Click()
Hapus
End Sub

Private Sub ctkulang_Click()
cetak_ulang = True
ctk_ulang.Show
ctk_ulang.WindowState = 2
End Sub

Private Sub Data1_Reposition()
isi_trans
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
Data3.Refresh
Data4.Refresh
Data5.Refresh
Data6.Refresh
Data7.Refresh
isi_stok
Isi_KAs
If Mid(MoneyChanger.Label1.Caption, 13) <> "admin" Then
    ctkulang.Visible = False
Else
    ctkulang.Visible = True
End If
End Sub

Private Sub Form_Load()
Call db_beli
Kosong
Tutup
cmd_awal
If MoneyChanger.user_mnu.Visible = False Then
    ctkulang.Visible = False
Else
    ctkulang.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
hapus_temp
update_kasharian
End Sub

Sub update_kasharian()
Data1.Enabled = False
Data7.RecordSource = "select * from kas_harian where cdate(tgl)='" & Date & "'"
Data7.Refresh
With Data7.Recordset
If Not .BOF Then
    .Edit
    !saldo = Val(Format(Text3, "###"))
    .Update
Else
    .AddNew
    !tgl = Date
    !saldo = Val(Format(Text3, "###"))
    .Update
End If
End With
End Sub

Sub hapus_temp()
Dim C As Single
Data1.Enabled = False
Data1.RecordSource = "select * from temp_trans where user ='" & Mid(MoneyChanger.Label1.Caption, 13) & "'"
Data1.Refresh
C = Data1.Recordset.RecordCount
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until C = 0
            .Delete
            C = C - 1
            .MoveNext
        Loop
        Kosong
    End If
End With
Data1.Enabled = True
Data1.Refresh
End Sub

Private Sub Image1_Click()
Nasabah_frm.Show
End Sub

Sub hitung()
Dim baris As Single
Dim counter As Single
Data1.Enabled = False
SSDBGrid1.Refresh
SSDBGrid1.MoveFirst
If SSDBGrid1.Columns(1).Value <> "" And SSDBGrid1.Columns(2).Value <> "" Then
With SSDBGrid1
.Update
.Refresh
baris = .Rows
counter = 0
If baris > 0 Then
    .MoveFirst
    Do Until counter = baris
        If .Columns(3).Value <> "" Then
        X = Val(X) + Format(.Columns(3).Value, "###.##")
        End If
        .MoveNext
        counter = counter + 1
    Loop
    Text2 = Format(X, "###,###.00")
    .MoveLast
End If
End With
End If
Data1.Enabled = True
End Sub

Private Sub SSDBGrid1_AfterUpdate(RtnDispErrMsg As Integer)
'hitung
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 3, 4
    If Text1(3) <> "" And Text1(4) <> "" Then
        Text1(5) = Format(Val(Text1(3)) * Val(Format(Text1(4), "###.00")), "###,###.00")
    End If
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 3, 4, 5
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Baru" Then
        If Label3.Caption <> "TRANSAKSI PENJUALAN" Then
            Buka
            Kosong
            isi_cmb1
            isi_Byr
            cmd_Simpan
            Command4_Click
        Else
            If Data3.Recordset.BOF Then
                MsgBox "Maaf transaksi tidak dapat dilanjutkan karena Persediaan = 0", vbCritical, "Validasi Stok"
            Else
                Buka
                Kosong
                isi_cmb1
                isi_Byr
                cmd_Simpan
                Command4_Click
            End If
        End If
    Else
        Data1.RecordSource = "select * from temp_trans where user ='" & Mid(MoneyChanger.Label1.Caption, 13) & "'"
        Data1.Refresh

        If Not Data1.Recordset.BOF Then
            X = MsgBox("Apakah anda yakin data sudah benar...???", vbYesNo, "Cek Data")
            If X = vbYes Then
                inv_auto
                INVOICE.ToolbarVisible = True
                cetak
                edit_BB
                add_jurnal
                Isi_ARAP
                MsgBox "Transaksi disimpan dengan No.Invoice :" & nomor_inv, vbInformation, "Simpan Transaksi"
                simpan
                update_stok
                cmd_awal
                Kosong
                Tutup
                ctk_inv = True
            End If
        Else
            MsgBox "Maaf anda belum melakukan transaksi...", vbInformation, "Validasi Transaksi"
        End If
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        Buka
        cmd_Simpan
    Else
        hapus_temp
        cmd_awal
        cmd_awalTrans
        Kosong
        Tutup
    End If
Case 3
Case 4
Case 5
Case 6
    nomor_inv = "-"
    INVOICE.ToolbarVisible = False
    cetak
    ctk_inv = False
Case 7
    Unload Me
End Select
End Sub

Sub cetak()
With INVOICE
    .Field1 = Format(Date, "dd-mmm-yyyy")
    .Field2 = Combo1(1)
    .Field3 = Text1(0)
    .Field4 = Text1(1)
    .Label11.Caption = Label3.Caption
    .Field7 = Text2
    .Field5 = nomor_inv
    .DAODataControl1.DatabaseName = Data1.DatabaseName
    .DAODataControl1.RecordSource = Data1.RecordSource
    .Refresh
    .Show
End With
End Sub

Sub Kosong()
Combo1(0) = ""
Combo1(1) = ""
Text1(0) = ""
Text1(1) = ""
'Text1(2) = ""
Text1(3) = ""
Text1(4) = ""
Text1(5) = ""
Text2 = ""
'Text3 = ""
Combo2(0) = ""
Combo2(1) = ""
kosong_Trans
End Sub

Sub kosong_Trans()
Combo1(2) = ""
Text1(3) = ""
Text1(4) = ""
Text1(5) = ""
End Sub

Sub isi()
With Data1.Recordset
    If Not .BOF And Data1.Enabled = True Then
        Combo1(0) = ""
        Combo1(1) = ""
        Text1(0) = ""
        Text1(1) = ""
        Text1(2) = ""
        Text2 = ""
    End If
End With
End Sub

Sub isi_trans()
With Data1.Recordset
    If Not .BOF And Data1.Enabled = True Then
        Combo1(2) = !curr
        Text1(3) = !jml
        Text1(4) = !satuan_rp
        Text1(5) = !total
    End If
End With
End Sub

Sub isi_cmb1()
Combo1(0).Clear
Combo1(0).AddItem "PERSEORANGAN"
Combo1(0).AddItem "PERUSAHAAN"
Combo1(0).ListIndex = 0
End Sub

Sub ISI_cmb2()
Combo1(1).Clear
With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo1(0) = "PERSEORANGAN" Then
            Combo1(1).AddItem !nama_nasabah
        Else
            Combo1(1).AddItem !nama_Perusahaan
        End If
        .MoveNext
    Loop
    Combo1(1).ListIndex = 0
End If
End With
End Sub

Sub isi_cmb3()
Combo1(2).Clear
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1(2).AddItem !simbol
        .MoveNext
    Loop
    Combo1(2).ListIndex = 0
End If
End With
End Sub

Sub isi_nasabah()
Dim nama As String
Data4.Refresh
With Data4.Recordset
If Not .BOF Then
    .MoveFirst
    If Combo1(0).ListIndex = 0 Then
    Do Until Combo1(1) = !nama_nasabah
        .MoveNext
    Loop
    Else
    Do Until Combo1(1) = !nama_Perusahaan
        .MoveNext
    Loop
    End If
    Text1(0) = !alamat
    Text1(1) = !telp
End If
End With
End Sub

Sub isi_Byr()
Combo2(0).Clear
Combo2(0).AddItem "KAS"
Combo2(0).AddItem "BANK"
If Label3.Caption = UCase("TRANSAKSI PEMBELIAN") Then
    Combo2(0).AddItem "HUTANG"
Else
    Combo2(0).AddItem "PIUTANG"
End If
Combo2(0).ListIndex = 0
End Sub

Sub isi_Rek()
Combo2(1).Clear
With Data7.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            Combo2(1).AddItem (!kode_bank & "-" & !no_rekening)
            .MoveNext
        Loop
        Combo2(1).ListIndex = 0
    End If
End With
End Sub

Sub isi_stok()
Dim ttl As Double
ttl = 0
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        ttl = ttl + !total_rp
        .MoveNext
    Loop
End If
Text1(2) = Format(ttl, "###,###.00")
End With
Data3.Refresh
End Sub

Sub Tutup()
Text3.Enabled = False
Combo1(0).Enabled = False
Combo1(1).Enabled = False
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text2.Enabled = False
Combo2(0).Enabled = False
Combo2(1).Enabled = False
tutup_Trans
End Sub

Sub tutup_Trans()
Combo1(2).Enabled = False
Text1(3).Enabled = False
Text1(4).Enabled = False
Text1(5).Enabled = False
End Sub

Sub Buka()
Combo1(0).Enabled = True
Combo1(1).Enabled = True
Combo2(0).Enabled = True
Combo2(1).Enabled = True
End Sub

Sub buka_trans()
Combo1(2).Enabled = True
Text1(3).Enabled = True
Text1(4).Enabled = True
Text1(5).Enabled = True
End Sub

Sub SIMPAN_TRANS()
Cek_Input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            cek_stok
            Data1.Refresh
            If cekstok = True Then
                .AddNew
                !curr = Combo1(2)
                !jml = Text1(3)
                !satuan_rp = Text1(4)
                !total = Text1(5)
                !user = Mid(MoneyChanger.Label1.Caption, 13)
                .Update
                tutup_Trans
                cmd_awalTrans
                Data1.Refresh
                hitung
            Else
                MsgBox "Stok Rupiah tidak cukup, silahkan isi yang lain...", vbInformation, "Validasi Data"
                Text1(3).SetFocus
            End If
        Else
            .Edit
            !curr = Combo1(2)
            !jml = Text1(3)
            !satuan_rp = Text1(4)
            !total = Text1(5)
            .Update
            tutup_Trans
            cmd_awalTrans
            Data1.Refresh
            hitung
        End If
    End With
End If
Data1.Refresh
End Sub

Sub simpan()
Cek_Input2
If cek3 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    transfer
    Tutup
    Kosong
    cmd_awal
    Data1.Refresh
    Data3.Refresh
    Data5.Refresh
End If
End Sub

Sub transfer()
Dim smbl As String
Dim rate As Double
Dim tot_beli As Double
tot_beli = 0
Data1.Enabled = False
Data5.RecordSource = "trans_jualbeli"
Data5.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        smbl = !curr
        With Data5.Recordset
            .AddNew
            If Label3.Caption = UCase("TRANSAKSI PEMBELIAN") Then
                !Status = "BELI"
            Else
                !Status = "JUAL"
            End If
            !jenis_nasabah = Combo1(0)
            !nama_nasabah = Combo1(1)
            !tgl = Date
            !jam = Time
            !user = Mid(MoneyChanger.Label1.Caption, 13)
            !no_inv = nomor_inv
            !simbol = Data1.Recordset!curr
            !jumlah = Data1.Recordset!jml
            !satuan_rp = Data1.Recordset!satuan_rp
            !total_rp = Data1.Recordset!total
            !cara_bayar = Combo2(0)
            With Data3.Recordset
                Data3.Refresh
                If Not .BOF Then
                    .MoveFirst
                    Do Until !simbol = smbl
                        .MoveNext
                    Loop
                    rate = !rate_rp
                End If
            End With
            !rate_dasar = rate
            !hrg_dasar = rate * Data1.Recordset!jml
            !net = !total_rp - !hrg_dasar
            tot_beli = tot_beli + !hrg_dasar
            .Update
        End With
        .MoveNext
    Loop
End If
End With
Data1.Enabled = True
End Sub

Sub Cek_Input2()
If Combo1(1) = "" Or Data1.Recordset.BOF Then
    cek3 = False
Else
    cek3 = True
End If
End Sub

Sub transfer_Trans()
With Data1.Recordset
    !curr = Combo1(2)
    !jml = Text1(3)
    !satuan_rp = Text1(4)
    !total = Text1(5)
End With
End Sub

Sub Cek_Input()
If Combo1(2) = "" Or Text1(3) = "" Or Text1(4) = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub cek_tambah()
cek = False
End Sub

Sub Hapus()
With Data1.Recordset
    If Not .BOF Then
        X = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If X = vbYes Then
            .Delete
            kosong_Trans
            hitung
            Data1.Refresh
        End If
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
End With
End Sub

Sub cmd_awal()
With Toolbar1
    .Buttons(1).Image = 1
    .Buttons(2).Image = 2
    .Buttons(1).Caption = "Baru"
    .Buttons(2).Caption = "Edit"
    .Buttons(1).ToolTipText = "Transaksi Baru"
    .Buttons(2).ToolTipText = "Edit Data"
    .Buttons(1).Visible = True 'add
    .Buttons(2).Visible = False 'edit
    .Buttons(3).Visible = False 'hapus
    .Buttons(4).Visible = False 'cari
    .Buttons(5).Visible = False 'preview
    .Buttons(6).Visible = False 'cetak
    .Buttons(7).Visible = True 'close
End With
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
End Sub

Sub cmd_Simpan()
With Toolbar1
    .Buttons(1).Image = 8
    .Buttons(2).Image = 3
    .Buttons(1).Caption = "Simpan"
    .Buttons(2).Caption = "Batal"
    .Buttons(1).ToolTipText = "Simpan Data"
    .Buttons(2).ToolTipText = "Batal Data"
    .Buttons(1).Visible = True 'Simpan
    .Buttons(2).Visible = True 'Batal
    .Buttons(3).Visible = False 'hapus
    .Buttons(4).Visible = False 'cari
    .Buttons(5).Visible = False 'preview
    .Buttons(6).Visible = True 'cetak
    .Buttons(7).Visible = False
End With
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
End Sub

Sub cmd_awalTrans()
Command4.Caption = "TAMBAH TRANSAKSI"
Command5.Caption = "EDIT TRANSAKSI"
Command6.Visible = True
Data1.Enabled = True
End Sub

Sub cmd_simpanTrans()
Command4.Caption = "SIMPAN TRANSAKSI"
Command5.Caption = "BATAL TRANSAKSI"
Command6.Visible = False
Data1.Enabled = False
End Sub

Sub update_stok()
Dim tot_rp As Double
Dim ada As Boolean
Dim valas As String
Dim valas2 As String
Dim tot As Double
Dim jml As Single
Dim sat As Double
Dim CUR As String
Dim tgl_awal As Date
Dim tot_stok As Double
Dim rate_akhir As Double
Dim jml_akhir As Single
Dim tot_akhir As Double

tot_stok = 0
Data1.Enabled = False
If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
    With Data1.Recordset
        If Not .BOF Then
            Data1.Refresh
            tot_rp = 0
            .MoveFirst
            Do While Not .EOF
                valas = !curr
                tot = !total
                jml = !jml
                sat = !satuan_rp
                tot_rp = tot_rp + !total
                Data5.RecordSource = "select status,sum(jumlah) as qty,sum(total_rp) as tot from trans_jualbeli where simbol='" & valas & "'and cdate(tgl)='" & Date & "' and status='BELI' group by status"
                Data5.Refresh
                Data7.RecordSource = "select * from stok_harian where cdate(tgl)<'" & Date & "' and currency='" & valas & "' order by tgl desc"
                Data7.Refresh
'                If Not Data7.Recordset.EOF Then
                'Data7.Recordset.MoveFirst
                With Data3.Recordset
                    Data3.Refresh
                    If Not .BOF Then
                        .MoveFirst
                        ada = False
                        Do While Not .EOF
                            If valas = !simbol Then
                                .Edit
                                !tgl = Date
                                !jam = Time
                                !jumlah = !jumlah + jml
                                If Not Data7.Recordset.BOF Then
                                    !rate_rp = (Data5.Recordset!tot + Data7.Recordset!total_rp) / (Data5.Recordset!qty + Data7.Recordset!jml)
                                Else
                                    !rate_rp = (Data5.Recordset!tot) / (Data5.Recordset!qty)
                                End If
                                !total_rp = !jumlah * !rate_rp
                                .Update
                                .MoveLast
                                ada = True
                            End If
                            .MoveNext
                        Loop
                        If ada = False Then
                            Data3.Refresh
                            .AddNew
                            !simbol = valas
                            !tgl = Date
                            !jam = Time
                            !rate_rp = sat
                            !jumlah = jml
                            !total_rp = !jumlah * !rate_rp
                            .Update
                            Data3.Refresh
                        End If
                    Else
                        Data3.Refresh
                        .AddNew
                        !simbol = valas
                        !tgl = Date
                        !jam = Time
                        !rate_rp = sat
                        !jumlah = jml
                        !total_rp = !jumlah * !rate_rp
                        .Update
                        Data3.Refresh
                    End If
                End With
 '               End If
                
                'update bulanan
                Data11.RecordSource = "select * from stok_bulanan where tahun='" & Year(Date) & "' and bulan='" & Month(Date) & "'"
                Data11.Refresh
                With Data11.Recordset
                    If Not .BOF Then
                        .MoveFirst
                        ada = False
                        Do While Not .EOF
                            If valas = !Currency Then
                                .Edit
                                !jml = !jml + jml
                                If Not Data7.Recordset.BOF Then
                                    !rate = (Data5.Recordset!tot + Data7.Recordset!total_rp) / (Data5.Recordset!qty + Data7.Recordset!jml)
                                Else
                                    !rate = (Data5.Recordset!tot) / (Data5.Recordset!qty)
                                End If
                                !total_rp = !jml * !rate
                                .Update
                                .MoveLast
                                ada = True
                            End If
                            .MoveNext
                        Loop
                        If ada = False Then
                            Data11.Refresh
                            .AddNew
                            !Currency = valas
                            !tahun = Year(Date)
                            !BULAN = Month(Date)
                            !rate = sat
                            !jml = jml
                            !total_rp = !jml * !rate
                            .Update
                            Data11.Refresh
                        End If
                    Else
                        If Month(Date) = 1 Then
                            Data11.RecordSource = "select * from stok_bulanan where tahun='" & Year(Date) - 1 & "' and bulan='12'"
                        Else
                            Data11.RecordSource = "select * from stok_bulanan where tahun='" & Year(Date) & "' and bulan='" & Month(Date) - 1 & "'"
                        End If
                        Data11.Refresh
                        .AddNew
                        !Currency = valas
                        !tahun = Year(Date)
                        !BULAN = Month(Date)
                        !rate = sat
                        !jml = jml
                        !total_rp = !jml * !rate
                        .Update
                        Data11.Refresh
                    End If
                End With
                
                'update harian
                Data11.RecordSource = "select * from stok_harian where cdate(tgl)='" & Date & "'"
                Data11.Refresh
                With Data11.Recordset
                    If Not Data11.Recordset.BOF Then
                            Data11.Recordset.MoveFirst
                            ada = False
                            Do While Not Data11.Recordset.EOF
                                If valas = Data11.Recordset!Currency Then
                                    .Edit
                                    !jml = !jml + jml
                                    If Not Data7.Recordset.BOF Then
                                        !rate = (Data5.Recordset!tot + Data7.Recordset!total_rp) / (Data5.Recordset!qty + Data7.Recordset!jml)
                                    Else
                                        !rate = (Data5.Recordset!tot) / (Data5.Recordset!qty)
                                    End If
                                    !total_rp = !jml * !rate
                                    .Update
                                    .MoveLast
                                End If
                                .MoveNext
                            Loop
                    Else
                        .Edit
                        !jml = !jml + jml
                        !rate = (Data5.Recordset!tot) / (Data5.Recordset!qty)
                        !total_rp = !jml * !rate
                        .Update
                    End If
                End With
                .MoveNext
                Data3.Refresh
            Loop
        End If
    End With
    

Else
    With Data1.Recordset
    tot_rp = 0
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            valas = !curr
            tot = !total
            jml = !jml
            sat = !satuan_rp
            tot_rp = tot_rp + !total
            With Data3.Recordset
                Data3.Refresh
                If Not .BOF Then
                    .MoveFirst
                    ada = False
                    Do While Not .EOF
                        If Data1.Recordset!curr = !simbol Then
                            .Edit
                            !tgl = Date
                            !jam = Time
                            !jumlah = !jumlah - jml
                            If !jumlah = 0 Then
'                                !rate_rp = 0
                                !total_rp = 0
                            Else
                                !total_rp = !jumlah * !rate_rp
                            End If
                            .Update
                            .MoveLast
                            ada = True
                        End If
                        .MoveNext
                    Loop
                    Data3.Refresh
                    If ada = False Then
                        .AddNew
                        !simbol = Data1.Recordset!curr
                        !tgl = Date
                        !jam = Time
'                        !rate_rp = Data1.Recordset!satuan_rp
                        !jumlah = Data1.Recordset!jml
                        !total_rp = !jumlah * !rate_rp
                        .Update
                        Data3.Refresh
                    End If
                Else
                    Data3.Refresh
                    .AddNew
                    !simbol = Data1.Recordset!curr
                    !tgl = Date
                    !jam = Time
'                    !rate_rp = Data1.Recordset!satuan_rp
                    !jumlah = Data1.Recordset!jml
                    !total_rp = !jumlah * !rate_rp
                    .Update
                    Data3.Refresh
                End If
            End With
            tot_rp = tot_rp + !total
            .MoveNext
        Loop
    End If
    End With
    
    'update stok bulanan
    With Data1.Recordset
    tot_rp = 0
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            valas = !curr
            tot = !total
            jml = !jml
            sat = !satuan_rp
            tot_rp = tot_rp + !total
            Data11.RecordSource = "select * from stok_bulanan where tahun='" & Year(Date) & "' and bulan='" & Month(Date) & "'"
            Data11.Refresh
            With Data11.Recordset
                If Not .BOF Then
                    .MoveFirst
                    ada = False
                    Do While Not .EOF
                        If Data1.Recordset!curr = !Currency Then
                            .Edit
                            !jml = !jml - jml
                            If !jml = 0 Then
                                !rate = 0
                                !total_rp = 0
                            Else
                                !total_rp = !jml * !rate
                            End If
                            .Update
                            .MoveLast
                            ada = True
                        End If
                        .MoveNext
                    Loop
                    Data11.Refresh
                    If ada = False Then
                        .AddNew
                        !Currency = Data1.Recordset!curr
                        !tahun = Year(Date)
                        !BULAN = Month(Date)
'                        !rate = Data1.Recordset!satuan_rp
                        !jml = Data1.Recordset!jml
                        !total_rp = !jml * !rate
                        .Update
                        Data11.Refresh
                    End If
                Else
                    If Month(Date) = 1 Then
                        Data11.RecordSource = "select * from stok_bulanan where tahun='" & Year(Date) - 1 & "' and bulan='12'"
                    Else
                        Data11.RecordSource = "select * from stok_bulanan where tahun='" & Year(Date) & "' and bulan='" & Month(Date) - 1 & "'"
                    End If
                    Data11.Refresh
                    If Not .BOF Then
                        .MoveFirst
                        rate_akhir = 0
                        jml_akhir = 0
                        tot_akhir = 0
                        Do While Not .EOF
                            If !Currency = valas Then
                                rate_akhir = !rate
                                jml_akhir = !jml
                                tot_akhir = !total_rp
                                .MoveLast
                            End If
                            .MoveNext
                        Loop
                    Else
                        rate_akhir = 0
                        jml_akhir = 0
                        tot_akhir = 0
                    End If
                    .AddNew
                    !Currency = Data1.Recordset!curr
                    !tahun = Year(Date)
                    !BULAN = Month(Date)
'                    !rate = Data1.Recordset!satuan_rp
                    !jml = jml_akhir - jml
                    !total_rp = !jml * !rate
                    .Update
                    Data11.Refresh
                End If
            End With
            tot_rp = tot_rp + !total
            .MoveNext
        Loop
    End If
    End With

    'update stok harian
    With Data1.Recordset
    If Not .BOF Then
        Data1.Refresh
        tot_rp = 0
        .MoveFirst
        Do While Not .EOF
            valas = !curr
            tot = !total
            jml = !jml
            sat = !satuan_rp
            tot_rp = tot_rp + !total
            Data11.RecordSource = "select * from stok_harian where cdate(tgl)='" & Date & "'"
            Data11.Refresh
                If Not Data11.Recordset.BOF Then
                    With Data11.Recordset
                        Data11.Recordset.MoveFirst
                        ada = False
                        Do While Not Data11.Recordset.EOF
                            If valas = Data11.Recordset!Currency Then
                                .Edit
                                !jml = !jml - jml
                                If !jml = 0 Then
'                                    !rate = 0
                                    !total_rp = 0
                                Else
                                    !total_rp = !jml * !rate
                                End If
                                .Update
                                .MoveLast
                            End If
                            .MoveNext
                        Loop
                    End With
                End If
            .MoveNext
            Data11.Refresh
        Loop
    End If
    End With
End If
Data1.Refresh
Data3.Refresh
Load Stok_frm
Stok_frm.update_BB
Unload Stok_frm
Data1.Enabled = True
End Sub

Sub inv_auto()
Dim urutan As String * 10
Dim hitung As Single
Data5.RecordSource = "select no_inv from trans_jualbeli order by no_inv asc"
Data5.Refresh
With Data5.Recordset
    If .RecordCount = 0 Then
        urutan = "000" & "0000001"
    Else
        .MoveLast
'        If Val(Left(.Fields("No_inv"), 7)) <> "0000000" Then
'            urutan = "0000000" & "0000001"
'        Else
            hitung = Val(Right(.Fields("no_inv"), 7)) + 1
            urutan = "000" & Right("0000000" & hitung, 7)
'        End If
    End If
    nomor_inv = urutan
End With
End Sub

Sub cek_stok()
With Data3.Recordset
Data3.Refresh
cekstok = True
If Not .BOF And Label3.Caption = "TRANSAKSI PENJUALAN" Then
    .MoveFirst
    Do Until !simbol = Combo1(2)
        .MoveNext
    Loop
    If Val(Format(Text1(3), "###")) > !jumlah Then
        cekstok = False
    Else
        cekstok = True
    End If
    Data3.Refresh
End If
End With
End Sub

Sub edit_BB()
Dim sal As Double
Dim no As String
Dim nil As Double
Dim nil2 As Double
Dim rate As Double
Data1.Enabled = False
If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
    If Combo2(0).ListIndex = 0 Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
    ElseIf Combo2(0).ListIndex = 1 Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-120' order by tahun desc,bulan desc"
    Else
        Data6.RecordSource = "select * from bb_bulanan where no_akun='2-110' order by tahun desc,bulan desc"
    End If
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = !saldo
        .Edit
        If Combo2(0).ListIndex = 2 Then
            !saldo = sal + Val(Format(Text2, "###.##"))
        Else
            !saldo = sal - Val(Format(Text2, "###.##"))
        End If
        .Update
    End With
    Data6.RecordSource = "select * from bb_bulanan where no_akun='1-160' order by tahun desc,bulan desc"
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal + Val(Format(Text2, "###.##"))
        .Update
    End With
Else
    If Combo2(0).ListIndex = 0 Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
    ElseIf Combo2(0).ListIndex = 1 Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-120' order by tahun desc,bulan desc"
    Else
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-130' order by tahun desc,bulan desc"
    End If
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal + Val(Format(Text2, "###.##"))
        .Update
    End With
    Data6.RecordSource = "select * from bb_bulanan where no_akun='4-100' order by tahun desc,bulan desc"
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal + Val(Format(Text2, "###.##"))
        .Update
    End With
End If

If Label3.Caption = "TRANSAKSI PENJUALAN" Then
nil = 0
nil2 = 0
Data1.Refresh
With Data1.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            With Data3.Recordset
                If Not .BOF Then
                    .MoveFirst
                    Do Until Data1.Recordset!curr = !simbol
                        .MoveNext
                    Loop
                    rate = !rate_rp
                End If
            End With
            nil = nil + (!jml * rate)
            nil2 = nil2 + !total
            .MoveNext
        Loop
    End If
End With

'hit hpp
    Data6.RecordSource = "select * from BB_BULANAN where no_akun='" & "5-100" & "' order by tahun desc,bulan desc"
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = Val(Format(Text1(2), "###.##"))
        .Edit
        !saldo = sal - (sal - nil)
        .Update
    End With
    
    'laba ditahan
    Data6.RecordSource = "select * from BB_BULANAN where no_akun='" & "3-200" & "' order by tahun desc,bulan desc"
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal + (nil2 - nil)
        .Update
    End With
    
    'update persediaan
    Data6.RecordSource = "select * from BB_BULANAN where no_akun='" & "1-160" & "' order by tahun desc,bulan desc"
    Data6.Refresh
    With Data6.Recordset
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal - nil
        .Update
    End With
    Data6.RecordSource = "select * from BB_BULANAN order by tahun desc,bulan desc"
    Data6.Refresh
End If

End Sub


Sub add_jurnal()
    Data8.Refresh
    With Data8.Recordset
        .AddNew
        !tgl = Date
        !jam = Time
        !user = Mid(MoneyChanger.Label1.Caption, 13)
        !jml = Val(Format(Text2, "###.##"))
        If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
            !dk = "DEBET"
            !ket = "TRANSAKSI PEMBELIAN VALAS NO.INV : " & nomor_inv
            !no_akun = "1-160"
        Else
            !dk = "KREDIT"
            !ket = "TRANSAKSI PENJUALAN VALAS NO.INV : " & nomor_inv
            !no_akun = "4-100"
        End If
        If Combo2(0).ListIndex = 0 Then
            !sumber_akun = "1-110"
        ElseIf Combo2(0).ListIndex = 1 Then
            !sumber_akun = "1-120"
        Else
            If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
                !sumber_akun = "2-110"
            Else
                !sumber_akun = "1-130"
            End If
        End If
        .Update
    Data8.Refresh
        .AddNew
        !tgl = Date
        !jam = Time
        !user = Mid(MoneyChanger.Label1.Caption, 13)
        !jml = Val(Format(Text2, "###.##"))
        If Combo2(0).ListIndex = 0 Then
            !no_akun = "1-110"
        ElseIf Combo2(0).ListIndex = 1 Then
            !no_akun = "1-120"
        Else
            If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
                !no_akun = "2-110"
            Else
                !no_akun = "1-130"
            End If
        End If
        If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
            !dk = "KREDIT"
            !ket = "TRANSAKSI PEMBELIAN VALAS"
        Else
            !dk = "DEBET"
            !ket = "TRANSAKSI PENJUALAN VALAS"
        End If
        !sumber_akun = "1-160"
        .Update
    End With
End Sub


Sub Isi_ARAP()
If Combo2(0).ListIndex = 2 Then
    If Label3.Caption = "TRANSAKSI PEMBELIAN" Then
        With Data9.Recordset
            Data9.Refresh
            .AddNew
            !no_inv = nomor_inv
            !tgl = Date
            !jam = Time
            !saldo = Val(Format(Text2, "###.##"))
            !jenis_nasabah = Combo1(0)
            !nama_nasabah = Combo1(1)
            .Update
        End With
    Else
        With Data10.Recordset
            Data10.Refresh
            .AddNew
            !no_inv = nomor_inv
            !tgl = Date
            !jam = Time
            !saldo = Val(Format(Text2, "###.##"))
            !jenis_nasabah = Combo1(0)
            !nama_nasabah = Combo1(1)
            .Update
        End With
    End If
End If
End Sub

Sub Isi_KAs()
Data6.RecordSource = "select * from BB_BULANAN where no_akun='1-110' order by tahun desc,bulan desc"
Data6.Refresh
With Data6.Recordset
If Not .BOF Then
    .MoveFirst
    Text3 = Format(!saldo, "###,###.00")
End If
End With
Data6.RecordSource = "bb_bulanan"
Data6.Refresh
End Sub
