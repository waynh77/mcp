VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Currency_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Tabel Currency"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "Currency_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA CURRENCY"
      Connect         =   "Access"
      DatabaseName    =   " "
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Ms_Currency"
      Top             =   2160
      Width           =   3375
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Currency_frm.frx":3482
      Height          =   5775
      Left            =   4560
      TabIndex        =   6
      Top             =   2640
      Width           =   5895
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   6509
      Columns(0).Caption=   "NAMA CURRENCY"
      Columns(0).Name =   "Nama_Currency"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Nama_Currency"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "SIMBOL CURRENCY"
      Columns(1).Name =   "Simbol"
      Columns(1).Alignment=   2
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Simbol"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   10398
      _ExtentY        =   10186
      _StockProps     =   79
      Caption         =   "TABEL DATA CURRENCY"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   7320
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   7320
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
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
            Picture         =   "Currency_frm.frx":3496
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":4E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":67BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":814C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":9ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":B470
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":C14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Currency_frm.frx":CE24
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
      MouseIcon       =   "Currency_frm.frx":E7B6
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SIMBOL CURRENCY"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA CURRENCY"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TABEL CURRENCY"
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
Attribute VB_Name = "Currency_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek As Boolean
Dim cek2 As Boolean
Dim smbl As String

Private Sub Data1_Reposition()
isi
End Sub

Private Sub Form_Activate()
Data1.Refresh
isi
End Sub

Private Sub Form_Load()
Call db_Currency
Tutup
Kosong
End Sub

Private Sub SSDBGrid1_Click()
isi
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Tambah" Then
        Buka
        Kosong
        tambah = True
        cmd_Simpan
        Text1(0).SetFocus
    Else
        simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        If Text1(0) <> "" Then
            Buka
            tambah = False
            smbl = Text1(1)
            cmd_Simpan
        Else
            MsgBox "Data Kosong", vbInformation, "Validasi Data"
        End If
    Else
        cmd_awal
        Tutup
        isi
    End If
Case 3
    Hapus
Case 4
Case 5
Case 6
Case 7
    Unload Me
End Select
End Sub

Sub Kosong()
Text1(0) = ""
Text1(1) = ""
End Sub

Sub isi()
With Data1.Recordset
    If Not .BOF And Data1.Enabled = True Then
        Text1(0) = !nama_currency
        Text1(1) = !simbol
    End If
End With
End Sub

Sub Tutup()
    Text1(0).Enabled = False
    Text1(1).Enabled = False
End Sub

Sub Buka()
    Text1(0).Enabled = True
    Text1(1).Enabled = True
End Sub

Sub simpan()
Cek_Input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            cek_tambah
            If cek = False Then
            Data1.Refresh
                .AddNew
                !nama_currency = Text1(0)
                !simbol = Text1(1)
                .Update
                Tutup
                cmd_awal
                Tambah_Persediaan
            Else
                MsgBox "Data sudah ada,silahkan isi yang lain...", vbInformation, "Validasi Data"
            End If
        Else
            .Edit
            !nama_currency = Text1(0)
            !simbol = Text1(1)
            .Update
            Tutup
            cmd_awal
'            Tambah_Persediaan
            Update_Persediaan
        End If
    End With
    Data1.Refresh
End If
End Sub

Sub transfer()
With Data1.Recordset
    !nama_currency = Text1(0)
    !simbol = Text1(1)
End With
End Sub

Sub Cek_Input()
cek2 = False
If Text1(0) = "" Or Text1(1) = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub cek_tambah()
cek = False
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Text1(0) = !nama_currency Then
            cek = True
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
End Sub

Sub Hapus()
With Data1.Recordset
    If Not .BOF Then
        X = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If X = vbYes Then
            Hapus_Persediaan
            .Delete
            Kosong
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
    .Buttons(1).Caption = "Tambah"
    .Buttons(2).Caption = "Edit"
    .Buttons(1).ToolTipText = "Tambah Data"
    .Buttons(2).ToolTipText = "Edit Data"
    .Buttons(3).Visible = True
    .Buttons(4).Visible = False
    .Buttons(5).Visible = False
    .Buttons(6).Visible = False
    .Buttons(7).Visible = True
End With
Data1.Enabled = True
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
End Sub

Sub Tambah_Persediaan()
With Data2.Recordset
Data2.Refresh
    .AddNew
    !tgl = Date
    !jam = Time
    !simbol = Text1(1)
    !rate_rp = 0
    !jumlah = 0
    !total_rp = 0
    .Update
End With
Data3.RecordSource = "stok_bulanan"
Data3.Refresh
With Data3.Recordset
    .AddNew
    !tahun = Year(Date)
    !BULAN = Month(Date)
    !Currency = Text1(1)
    !rate = 0
    !jml = 0
    !total_rp = 0
    .Update
End With
Data3.RecordSource = "stok_harian"
Data3.Refresh
With Data3.Recordset
    .AddNew
    !tgl = Date
    !Currency = Text1(1)
    !rate = 0
    !jml = 0
    !total_rp = 0
    .Update
    .AddNew
    !tgl = Date - 1
    !Currency = Text1(1)
    !rate = 0
    !jml = 0
    !total_rp = 0
    .Update
End With
End Sub

Sub Update_Persediaan()
Data2.Refresh
Dim ketemu As Boolean
ketemu = False
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do While Not .EOF
            If !simbol = smbl Then
                .Edit
                !simbol = Text1(1)
                .Update
                ketemu = True
                .MoveLast
            End If
            .MoveNext
        Loop
    End If
End With
If ketemu = False Then
    Tambah_Persediaan
End If
End Sub

Sub Hapus_Persediaan()
Data2.Refresh
With Data2.Recordset
    If Not .BOF Then
        .MoveFirst
        Do Until !simbol = Text1(1)
            .MoveNext
        Loop
        X = MsgBox("Apakah data persediaan akan dihapus juga?", vbYesNo, "Hapus Persediaan")
        If X = vbYes Then
            .Delete
        End If
        Data2.Refresh
    End If
End With
End Sub
