VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Kas_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Transaksi Kas"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Kas_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   300
      Index           =   5
      Left            =   9600
      TabIndex        =   33
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "BB"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "AKUN"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Index           =   2
      Left            =   9600
      TabIndex        =   31
      Top             =   4560
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Index           =   1
      Left            =   9600
      TabIndex        =   30
      Top             =   4080
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9600
      TabIndex        =   28
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10320
      Top             =   1560
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "TRANSAKSI KAS"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   12000
      TabIndex        =   26
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   5
      Left            =   4920
      TabIndex        =   24
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   4
      Left            =   9600
      TabIndex        =   22
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   1140
      Index           =   3
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5520
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Index           =   0
      Left            =   9600
      TabIndex        =   20
      Top             =   3600
      Width           =   2055
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Kas_frm.frx":3482
      Height          =   3615
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   6735
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   11880
      _ExtentY        =   6376
      _StockProps     =   79
      Caption         =   "DATA TRANSAKSI"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   12000
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   7320
      TabIndex        =   11
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Format          =   60882945
      CurrentDate     =   39848
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   12000
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   7320
      TabIndex        =   6
      Text            =   "KAS"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Text            =   "1-110"
      Top             =   840
      Width           =   2055
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
            Picture         =   "Kas_frm.frx":3496
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":4E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":67BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":814C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":9ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":B470
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":C14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Kas_frm.frx":CE24
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
      MouseIcon       =   "Kas_frm.frx":E7B6
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JAM TRANSAKSI"
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
      Index           =   15
      Left            =   7200
      TabIndex        =   32
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUB AKUN"
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
      Left            =   7200
      TabIndex        =   29
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDO AKHIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   14
      Left            =   9600
      TabIndex        =   27
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   240
      X2              =   14040
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JENIS AKUN"
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
      Index           =   13
      Left            =   7200
      TabIndex        =   25
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL TRANSAKSI"
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
      Index           =   12
      Left            =   3000
      TabIndex        =   23
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH RP"
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
      Left            =   7200
      TabIndex        =   18
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "URAIAN/KETERANGAN"
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
      Left            =   7200
      TabIndex        =   17
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE AKUN"
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
      Left            =   7200
      TabIndex        =   16
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA PERKIRAAN"
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
      Index           =   7
      Left            =   7200
      TabIndex        =   15
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KAS = DEBET/KREDIT ..... TRANSAKSI = DEBET/KREDIT"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   13815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JAM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   5
      Left            =   9600
      TabIndex        =   12
      Top             =   1560
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
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   4
      Left            =   4920
      TabIndex        =   10
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KAS DEBET/KREDIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   14040
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDO SAAT INI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   2
      Left            =   9600
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA AKUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KODE AKUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSAKSI KAS MASUK"
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
Attribute VB_Name = "Kas_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek2 As Boolean
Dim keter As String
Dim sld As Double

Private Sub Combo1_Change()
isi_akun
End Sub

Private Sub Combo1_Click()
isi_akun
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Data1_Reposition()
Isi
End Sub

Private Sub DTPicker1_Change()
isi_grid
End Sub

Private Sub DTPicker1_Click()
isi_grid
End Sub

Private Sub Form_Activate()
isi_grid
isi_saldoAwal
Data2.Refresh
Data3.Refresh
hit_trans
End Sub

Sub isi_grid()
If Label3 = "TRANSAKSI KAS MASUK" Then
    Data1.RecordSource = "select * from rekap_jurnal where cdate(tgl)='" & DTPicker1 & "' and no_akun <> '1-110' and sumber_akun='1-110' and dk='KREDIT'"
Else
    Data1.RecordSource = "select * from rekap_jurnal where cdate(tgl)='" & DTPicker1 & "' and no_akun <> '1-110' and sumber_akun='1-110' and dk='DEBET'"
End If
Data1.Refresh
hit_trans
Isi
End Sub

Private Sub Form_Load()
Call DB_Kas
Kosong
Tutup
DTPicker1 = Date
cmd_awal
End Sub

Private Sub Form_Unload(Cancel As Integer)
update_kasharian
End Sub

Private Sub Label3_Change()
Kosong
Tutup
cmd_awal
isi_grid
isi_saldoAwal
End Sub

Private Sub SSDBGrid1_Click()
Isi
End Sub

Private Sub Text2_Change(Index As Integer)
isi_akun2
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 4
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub Timer1_Timer()
Text1(4) = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Tambah" Then
        Buka
        Kosong
        ISI_cmb
        tambah = True
        cmd_Simpan
        sld = 0
    Else
        Simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        Buka
        tambah = False
        keter = "1-110" & DTPicker1 & Text2(5) & Text2(3)
        sld = Val(Format(Text2(4), "###"))
        ISI_cmb
        Isi
        Text2(4) = Format(Text2(4), "###")
        cmd_Simpan
    Else
        cmd_awal
        Tutup
        Data1.Refresh
    End If
Case 3
Case 4
Case 5
Case 6
Case 7
    Unload Me
End Select
End Sub

Sub Kosong()
Combo1 = ""
Text2(0) = ""
Text2(1) = ""
Text2(2) = ""
Text2(3) = ""
Text2(4) = ""
Text2(5) = ""
End Sub

Sub Isi()
With Data1.Recordset
If Not .BOF And Data1.Enabled = True Then
    Text2(0) = !no_akun
    Text2(3) = !ket
    Text2(4) = Format(!jml, "###,###.00")
    Text2(5) = !jam
End If
End With
End Sub

Sub ISI_cmb()
If Label3.Caption = "TRANSAKSI KAS MASUK" Then
    Data2.RecordSource = "select * from tbl_akun where no_akun <> '1-110' and left(no_akun,1) <> '6' order by no_akun"
Else
    Data2.RecordSource = "select * from tbl_akun where no_akun <> '1-110'  and left(no_akun,1) <> '5' order by no_akun"
End If
Data2.Refresh
Combo1.Clear
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !nama_akun
        .MoveNext
    Loop
    Combo1.ListIndex = 0
End If
End With
End Sub

Sub isi_akun()
Data2.Refresh
With Data2.Recordset
    If Not .BOF And Combo1 <> "" Then
        .MoveFirst
        Do Until !nama_akun = Combo1
            .MoveNext
        Loop
        Text2(0) = !no_akun
        Text2(1) = !nama_jenisakun
        Text2(2) = !nama_subakun
    End If
End With
End Sub

Sub isi_akun2()
Data2.Refresh
With Data2.Recordset
    If Not .BOF And Text2(0) <> "" And Not Data1.Recordset.BOF Then
        .MoveFirst
        Do Until !no_akun = Text2(0)
            .MoveNext
        Loop
        Combo1 = !nama_akun
        Text2(1) = !nama_jenisakun
        Text2(2) = !nama_subakun
    End If
End With
End Sub

Sub isi_saldoAwal()
Data3.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
Data3.Refresh
With Data3.Recordset
    If Not .BOF Then
        .MoveFirst
        Text1(2) = Format(!saldo, "###,###.00")
    End If
End With
Data3.Refresh
End Sub

Sub Tutup()
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(3).Enabled = False
Text1(4).Enabled = False
Text1(5).Enabled = False
Text1(6).Enabled = False
Text2(0).Enabled = False
Text2(1).Enabled = False
Text2(2).Enabled = False
Text2(3).Enabled = False
Text2(4).Enabled = False
Text2(5).Enabled = False
Combo1.Enabled = False
End Sub

Sub Buka()
Combo1.Enabled = True
Text2(3).Enabled = True
Text2(4).Enabled = True
Text2(4) = Format(Text2(4), "###")
End Sub

Sub Simpan()
Cek_Input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            edit_BB
            Data1.Refresh
            .AddNew
            !tgl = DTPicker1
            !jam = Text1(4)
            !no_akun = Text2(0)
            !ket = Text2(3)
            !jml = Text2(4)
            If Label3.Caption = "TRANSAKSI KAS MASUK" Then
                !dk = "KREDIT"
            Else
                !dk = "DEBET"
            End If
            !user = Mid(MoneyChanger.Label1.Caption, 13)
            !sumber_akun = "1-110"
            .Update
            input_kas
        Else
            edit_BB
            .Edit
            !tgl = DTPicker1
            !jam = Text1(4)
            !no_akun = Text2(0)
            !ket = Text2(3)
            !jml = Text2(4)
            If Label3.Caption = "TRANSAKSI KAS MASUK" Then
                !dk = "KREDIT"
            Else
                !dk = "DEBET"
            End If
            !user = Mid(MoneyChanger.Label1.Caption, 13)
            !sumber_akun = "1-110"
            .Update
            Edit_kas
        End If
    End With
    Tutup
    cmd_awal
    Data1.Refresh
    hit_trans
    isi_saldoAwal
End If
End Sub

Sub transfer()
With Data1.Recordset
    !tgl = DTPicker1
    !jam = Text1(4)
    !no_akun = Text2(0)
    !ket = Text2(3)
    !jml = Text2(4)
    If Label3.Caption = "TRANSAKSI KAS MASUK" Then
        !dk = "KREDIT"
    Else
        !dk = "DEBET"
    End If
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = "1-110"
End With
End Sub

Sub input_kas()
With Data1.Recordset
    Data1.Refresh
    .AddNew
    !tgl = DTPicker1
    !jam = Text1(4)
    !no_akun = Text1(0)
    !ket = Text2(3)
    !jml = Text2(4)
    If Label3.Caption = "TRANSAKSI KAS KELUAR" Then
        !dk = "KREDIT"
    Else
        !dk = "DEBET"
    End If
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = Text2(0)
    .Update
End With
End Sub

Sub Edit_kas()
Data1.RecordSource = "select * from REKAP_JURNAL where cdate(tgl)='" & DTPicker1 & "'"
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until !no_akun & !tgl & !jam & !ket = keter
        .MoveNext
    Loop
    .Edit
    !jam = Text1(4)
    !ket = Text2(3)
    !jml = Text2(4)
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = Text2(0)
    .Update
    
    .MoveFirst
End If
End With
isi_grid
End Sub

Sub Cek_Input()
If Text2(0) = "" Or Text2(4) = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub Hapus()
With Data1.Recordset
    If Not .Recordset.BOF And Not .Recordset.EOF Then
        X = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If X = vbYes Then
            .Recordset.Delete
            .Refresh
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
    .Buttons(1).Caption = "Tambah"
    .Buttons(2).Caption = "Edit"
    .Buttons(1).ToolTipText = "Tambah Data"
    .Buttons(2).ToolTipText = "Edit Data"
    .Buttons(3).Visible = False
    .Buttons(4).Visible = False
    .Buttons(5).Visible = False
    .Buttons(6).Visible = False
    .Buttons(7).Visible = True
End With
Data1.Enabled = True
SSDBGrid1.Enabled = True
DTPicker1.Enabled = True
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
Data1.Enabled = False
SSDBGrid1.Enabled = False
DTPicker1.Enabled = False
End Sub

Sub hit_trans()
Dim t As Double
Data1.Enabled = False
Data1.Refresh
t = 0
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        t = t + !jml
        .MoveNext
    Loop
End If
Text1(5) = Format(t, "###,###.00")
Data1.Enabled = True
End With
Data1.Refresh
End Sub

Sub Add_Bb()
Dim sal As Double
Dim dk As String
Data3.RecordSource = "select * from bukubesar where no_akun='1-110' order by tgl desc,jam desc"
Data3.Refresh
With Data3.Recordset
    .MoveFirst
    sal = !saldo
    .AddNew
    !tgl = DTPicker1
    !jam = Text1(4)
    !no_akun = "1-110"
    If Label3.Caption = "TRANSAKSI KAS MASUK" Then
        !saldo = sal + Val(Format(Text2(4), "###"))
    Else
        !saldo = sal - Val(Format(Text2(4), "###"))
    End If
    .Update
End With
Data3.RecordSource = "select * from bukubesar where no_akun='" & Text2(0) & "' order by tgl desc,jam desc"
Data3.Refresh
With Data3.Recordset
    .MoveFirst
    sal = !saldo
    .AddNew
    !tgl = DTPicker1
    !jam = Text1(4)
    !no_akun = Text2(0)
    With Data2.Recordset
        If Not .BOF Then
            .MoveFirst
            Do Until !no_akun = Data3.Recordset!no_akun
                .MoveNext
            Loop
            dk = !dk
        End If
    End With
    If Label3.Caption = "TRANSAKSI KAS MASUK" And dk = "DEBET" Then
        !saldo = sal - Val(Format(Text2(4), "###"))
    Else
        !saldo = sal + Val(Format(Text2(4), "###"))
    End If
    .Update
End With
Data3.RecordSource = "select * from bukubesar where no_akun='1-110' order by tgl desc,jam desc"
Data3.Refresh
End Sub

Sub edit_BB()
Dim sal As Double
Dim sel As Double
Dim dk As String
sel = 0
sel = Val(Format(Text2(4), "###")) - sld
Data3.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
Data3.Refresh
With Data3.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    If Label3.Caption = "TRANSAKSI KAS MASUK" Then
        !saldo = sal + sel
    Else
        !saldo = sal - sel
    End If
    .Update
End With
Data3.RecordSource = "select * from bb_bulanan where no_akun='" & Text2(0) & "' order by tahun desc,bulan desc"
Data3.Refresh
With Data3.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    With Data2.Recordset
        If Not .BOF Then
            .MoveFirst
            Do Until !no_akun = Data3.Recordset!no_akun
                .MoveNext
            Loop
            dk = !dk
        End If
    End With
    If (Label3.Caption = "TRANSAKSI KAS MASUK" And dk = "DEBET") Or (Label3.Caption = "TRANSAKSI KAS KELUAR" And dk = "KREDIT") Then
        !saldo = sal - sel
    Else
        !saldo = sal + sel
    End If
    .Update
End With
Data3.Refresh

'update laba rugi
If Mid(Text2(0), 1, 1) = "6" And Label3.Caption = "TRANSAKSI KAS KELUAR" Then
Data3.RecordSource = "select * from bb_bulanan where no_akun='3-200' order by tahun desc,bulan desc"
Data3.Refresh
With Data3.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal - sel
    .Update
End With
End If
End Sub

Sub update_kasharian()
Data1.Enabled = False
Data1.RecordSource = "select * from kas_harian where cdate(tgl)='" & Date & "'"
Data1.Refresh
With Data1.Recordset
If Not .BOF Then
    .Edit
    !saldo = Val(Format(Text1(2), "###"))
    .Update
Else
    .AddNew
    !tgl = Date
    !saldo = Val(Format(Text1(2), "###"))
    .Update
End If
End With
End Sub

