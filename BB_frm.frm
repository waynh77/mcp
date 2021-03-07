VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form BB_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Buku Besar"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "BB_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   9960
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   7
      Left            =   9960
      TabIndex        =   19
      Top             =   6240
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   9960
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   9960
      TabIndex        =   14
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   9960
      TabIndex        =   13
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   9960
      TabIndex        =   12
      Top             =   3600
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   9960
      TabIndex        =   11
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   9960
      TabIndex        =   4
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   9960
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "BUKU BESAR"
      Connect         =   "Access"
      DatabaseName    =   "C:\WaynhSoft\Proyek\Money Changer\Source Code MCP\DBMCP-recover.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BB_Bulanan"
      Top             =   6960
      Width           =   3375
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
            Picture         =   "BB_frm.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":4E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":67A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":8138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":9ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":B45C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":C136
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BB_frm.frx":CE10
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
      ButtonWidth     =   1032
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
            Object.Visible         =   0   'False
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
      MouseIcon       =   "BB_frm.frx":E7A2
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "BB_frm.frx":EABC
      Height          =   7455
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   7095
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2725
      Columns(0).Caption=   "TAHUN"
      Columns(0).Name =   "TANGGAL"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Tahun"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2646
      Columns(1).Caption=   "BULAN"
      Columns(1).Name =   "JAM"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Bulan"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2858
      Columns(2).Caption=   "NO AKUN"
      Columns(2).Name =   "NO AKUN"
      Columns(2).Alignment=   2
      Columns(2).DataField=   "No_Akun"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "SALDO"
      Columns(3).Name =   "SALDO"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Saldo"
      Columns(3).DataType=   5
      Columns(3).FieldLen=   256
      _ExtentX        =   12515
      _ExtentY        =   13150
      _StockProps     =   79
      Caption         =   "DATABASE BUKU BESAR"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   10320
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   57212929
      CurrentDate     =   39856
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
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   8
      Left            =   7560
      TabIndex        =   18
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Line Line4 
      X1              =   7560
      X2              =   14280
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDO RP"
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
      Left            =   7560
      TabIndex        =   17
      Top             =   6240
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
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   9
      Left            =   7560
      TabIndex        =   15
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   13680
      MouseIcon       =   "BB_frm.frx":EAD0
      MousePointer    =   99  'Custom
      Picture         =   "BB_frm.frx":EDDA
      Stretch         =   -1  'True
      ToolTipText     =   "Tambah Nasabah Baru"
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   7560
      X2              =   14280
      Y1              =   5040
      Y2              =   5040
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
      Index           =   0
      Left            =   7560
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR AKUN"
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
      Left            =   7560
      TabIndex        =   9
      Top             =   3120
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
      Index           =   2
      Left            =   7560
      TabIndex        =   8
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEBET/KREDIT"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NERACA/LABA-RUGI"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   4560
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
      Index           =   5
      Left            =   7560
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   7560
      X2              =   14280
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   14280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSAKSI BUKU BESAR"
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
Attribute VB_Name = "BB_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cek As Boolean
Dim cek2 As Boolean

Private Sub Data1_Reposition()
isi
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
End Sub

Private Sub Form_Load()
Call Db_BB
cmd_awal
Kosong
tutup
End Sub

Private Sub SSDBGrid1_Click()
'Isi
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Tambah" Then
    Else
        simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
    If Not Data1.Recordset.BOF Then
        buka
        cmd_simpan
        Text1(7) = Format(Text1(7), "###")
        Text1(7).SetFocus
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
    Else
        tutup
        Data1.Refresh
        cmd_awal
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
Dim x As Byte
x = 0
Do Until x = 8
    Text1(x) = ""
    x = x + 1
Loop
DTPicker1 = Date
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF And Data1.Enabled = True Then
    Data2.Refresh
    With Data2.Recordset
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If !no_akun = Data1.Recordset!no_akun Then
                    Text1(0) = !nama_jenisakun
                    Text1(1) = !nama_subakun
                    Text1(2) = !no_akun
                    Text1(3) = !nama_akun
                    Text1(4) = !dk
                    Text1(5) = !nrlr
                    .MoveLast
                End If
                .MoveNext
            Loop
        End If
    End With
'    DTPicker1 = !tgl
    Text2 = !BULAN
    Text1(6) = !tahun
    Text1(7) = Format(!saldo, "###,###.00")
End If
End With
End Sub

Sub tutup()
Dim x As Byte
x = 0
Do Until x = 8
    Text1(x).Enabled = False
    x = x + 1
Loop
Text2.Enabled = False
'DTPicker1.Enabled = False
End Sub

Sub buka()
Text1(7).Enabled = True
End Sub

Sub simpan()
Cek_Input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        .Edit
        !saldo = Text1(7)
        .Update
    End With
    cmd_awal
    Data1.Refresh
End If
End Sub

Sub transfer()
With Data1.Recordset
    !saldo = Text1(7)
End With
End Sub

Sub Cek_Input()
If Text1(7) = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub hapus()
With Data1.Recordset
    If Not .Recordset.BOF And Not .Recordset.EOF Then
        x = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If x = vbYes Then
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
'    .Buttons(1).Caption = "Baru"
    .Buttons(2).Caption = "Edit"
'    .Buttons(1).ToolTipText = "Transaksi Baru"
    .Buttons(2).ToolTipText = "Edit Data"
    .Buttons(1).Visible = False 'add
    .Buttons(2).Visible = True 'edit
    .Buttons(3).Visible = False 'hapus
    .Buttons(4).Visible = False 'cari
    .Buttons(5).Visible = False 'preview
    .Buttons(6).Visible = False 'cetak
    .Buttons(7).Visible = True 'close
End With
Data1.Enabled = True
SSDBGrid1.Enabled = True
tutup
Text1(7) = Format(Text1(7), "###,###.00")
End Sub

Sub cmd_simpan()
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
    .Buttons(6).Visible = False 'cetak
    .Buttons(7).Visible = False
End With
Data1.Enabled = False
SSDBGrid1.Enabled = False
buka
Text1(7) = Format(Text1(7), "###")
End Sub


