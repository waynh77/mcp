VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form ctk_ulang 
   BackColor       =   &H00FF8080&
   Caption         =   "CETAK ULANG INVOICE"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "ctk_ulang.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data7 
      Caption         =   "Kas Harian"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   11
      Left            =   11880
      TabIndex        =   26
      Text            =   " "
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Data Data6 
      Caption         =   "Buku_Besar"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data5 
      Caption         =   "Stok_Bulanan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data4 
      Caption         =   "persediaan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9480
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctk_ulang.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctk_ulang.frx":365C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctk_ulang.frx":4FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctk_ulang.frx":6980
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctk_ulang.frx":765A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   10
      Left            =   11880
      TabIndex        =   25
      Text            =   " "
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   9
      Left            =   11880
      TabIndex        =   1
      Text            =   " "
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   11880
      TabIndex        =   0
      Text            =   " "
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   11880
      TabIndex        =   24
      Text            =   " "
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   11880
      TabIndex        =   18
      Text            =   " "
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Data Data3 
      Caption         =   "Nasabah"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   11880
      TabIndex        =   16
      Text            =   " "
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   11880
      TabIndex        =   14
      Text            =   " "
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   11880
      TabIndex        =   12
      Text            =   " "
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   11880
      TabIndex        =   11
      Text            =   " "
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   11880
      TabIndex        =   10
      Text            =   " "
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   11880
      TabIndex        =   9
      Text            =   " "
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Data Data2 
      Caption         =   "Trans_JualBeli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "ctk_ulang.frx":8334
      Height          =   6015
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   7455
      _Version        =   196614
      BevelColorFace  =   255
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   13150
      _ExtentY        =   10610
      _StockProps     =   79
      Caption         =   "TRANSAKSI"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Data Data1 
      BackColor       =   &H00404040&
      Caption         =   "DATA TRANSAKSI"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   11880
      TabIndex        =   4
      Text            =   " "
      Top             =   1080
      Width           =   2655
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
      Index           =   12
      Left            =   9480
      TabIndex        =   27
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   14040
      MouseIcon       =   "ctk_ulang.frx":8348
      MousePointer    =   99  'Custom
      Picture         =   "ctk_ulang.frx":8652
      ToolTipText     =   "Cetak/Print"
      Top             =   7200
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL (A)*(B)"
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
      Left            =   9480
      TabIndex        =   23
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RATE (B)"
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
      Left            =   9480
      TabIndex        =   22
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUMLAH (A)"
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
      Left            =   9480
      TabIndex        =   21
      Top             =   5760
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
      Index           =   8
      Left            =   9480
      TabIndex        =   20
      Top             =   5280
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
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   6
      Left            =   9480
      TabIndex        =   19
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TELEPON"
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
      Left            =   9480
      TabIndex        =   17
      Top             =   3240
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   4
      Left            =   9480
      TabIndex        =   15
      Top             =   2880
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
      Index           =   3
      Left            =   9480
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   13320
      MouseIcon       =   "ctk_ulang.frx":9FD4
      MousePointer    =   99  'Custom
      Picture         =   "ctk_ulang.frx":A2DE
      ToolTipText     =   "Tutup/Close"
      Top             =   7200
      Width           =   480
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
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   2
      Left            =   9480
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   1
      Left            =   9480
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS TRANSAKSI"
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
      Left            =   9480
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   12600
      MouseIcon       =   "ctk_ulang.frx":AFA8
      MousePointer    =   99  'Custom
      Picture         =   "ctk_ulang.frx":B2B2
      ToolTipText     =   "Cetak/Print"
      Top             =   7200
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOMOR INVOICE"
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
      Left            =   9480
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CETAK ULANG INVOICE"
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
      TabIndex        =   2
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "ctk_ulang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rate_awal As Double
Dim jml_awal As Single
Dim tot_awal As Double
Dim net_awal As Double
Dim net_akhir As Double
Dim sedia As Double


Private Sub Combo1_Click()
isi_data
End Sub

Private Sub Data1_Reposition()
isi_trans
End Sub

Private Sub Form_Activate()
ISI_cmb
End Sub

Private Sub Form_Load()
Call db_ctkUlang
cetak_ulang = True
Data2.RecordSource = "select tgl,no_inv,status,jenis_nasabah,nama_nasabah,CARA_BAYAR from trans_jualbeli group by no_inv,status,jenis_nasabah,nama_nasabah,tgl,CARA_BAYAR order by no_inv desc"
cmd_awal
End Sub

Sub ISI_cmb()
Combo1.Clear
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !no_inv
        .MoveNext
    Loop
    Combo1.ListIndex = 0
End If
End With
End Sub

Sub isi_data()
Dim nil As Double
Data2.Refresh
With Data2.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        If Combo1 = !no_inv Then
            Text1(0) = !Status
            Text1(1) = !jenis_nasabah
            Text1(2) = !nama_nasabah
            Text1(3) = !tgl
            Text1(11) = !cara_bayar
            .MoveLast
        End If
        .MoveNext
    Loop
End If
End With
Data1.RecordSource = "select tgl,jam,simbol,jumlah,satuan_rp,total_rp,rate_dasar,hrg_dasar,net from trans_jualbeli where no_inv='" & Combo1 & "'"
Data1.Refresh
nil = 0
Data1.Enabled = False
With Data1.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        nil = nil + !total_rp
        .MoveNext
    Loop
End If
Text1(6) = Format(nil, "###,###.00")
End With
Data1.Enabled = True
Data1.Refresh
If Text1(1).Text = "PERSEORANGAN" Then
    Data3.RecordSource = "select * from msnasabah_perseorangan where nama_nasabah='" & Text1(2).Text & "'"
Else
    Data3.RecordSource = "select * from msnasabah_perusahaan where nama_perusahaan='" & Text1(2).Text & "'"
End If
Data3.Refresh
With Data3.Recordset
If Not .BOF Then
    Text1(4) = !alamat
    Text1(5) = !telp
End If
End With
End Sub

Sub isi_trans()
With Data1.Recordset
If Not .BOF And Data1.Enabled = True Then
    Text1(7) = !simbol
    Text1(8) = Format(!jumlah, "###,###")
    Text1(9) = Format(!satuan_rp, "###,###.00")
    Text1(10) = Format(!total_rp, "###,###.00")
End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
cetak_ulang = False
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
If Image1(0).ToolTipText = "Edit Transaksi" Then
    If Combo1 <> "" Then
        cmd_Simpan
        buka_trans
        Text1(8).SetFocus
        jml_awal = Val(Format(Text1(8), "###"))
        rate_awal = Val(Format(Text1(9), "###"))
        tot_awal = Val(Format(Text1(10), "###"))
        net_awal = Data1.Recordset!net
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
Else
    simpan
End If
Case 1
If Image1(1).ToolTipText = "Cetak Ulang" And Combo1 <> "" Then
    With INVOICE
        .Field1 = ": " & Format(Text1(3), "dd-mmm-yyyy")
        .Field2 = ": " & Text1(2) 'nama
        .Field3 = ": " & Text1(4) 'alamat
        .Field4 = ": " & Text1(5) 'telepon
        If Text1(0) = "BELI" Then
            .Label11.Caption = "TRANSAKSI PEMBELIAN"
        Else
            .Label11.Caption = "TRANSAKSI PENJUALAN"
        End If
        .Field7 = Text1(6) 'total transaksi
        .Field5 = Combo1 'no_inv
        .DAODataControl1.DatabaseName = Data1.DatabaseName
        .DAODataControl1.RecordSource = Data1.RecordSource
        .Refresh
        .Show
    End With
Else
    cmd_awal
    tutup_Trans
End If
Case 2
    Unload Me
End Select
End Sub


Sub cmd_awal()
Image1(0).Picture = ImageList1.ListImages(3).Picture
Image1(1).Picture = ImageList1.ListImages(4).Picture
Image1(2).Picture = ImageList1.ListImages(5).Picture
Image1(0).ToolTipText = "Edit Transaksi"
Image1(1).ToolTipText = "Cetak Ulang"
Image1(2).ToolTipText = "Keluar"
Image1(2).Visible = True
Data1.Enabled = True
SSDBGrid1.Enabled = True
End Sub

Sub cmd_Simpan()
Image1(0).Picture = ImageList1.ListImages(1).Picture
Image1(1).Picture = ImageList1.ListImages(2).Picture
Image1(2).Picture = ImageList1.ListImages(5).Picture
Image1(0).ToolTipText = "Simpan Transaksi"
Image1(1).ToolTipText = "Batal"
Image1(2).ToolTipText = "Keluar"
Image1(2).Visible = False
Data1.Enabled = False
SSDBGrid1.Enabled = False
End Sub

Sub buka_trans()
Text1(8).Enabled = True
Text1(9).Enabled = True
Text1(8) = Format(Text1(8), "###")
Text1(9) = Format(Text1(9), "###")
End Sub

Sub tutup_Trans()
Text1(8).Enabled = False
Text1(9).Enabled = False
End Sub


Sub hit()
Text1(10) = Format(Val(Format(Text1(8), "###")) * Val(Format(Text1(9), "###")), "###,###.00")
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 8, 9
    If Text1(8) <> "" And Text1(9) <> "" Then
        hit
    End If
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 8, 9
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Select
End Sub

Sub simpan()
Dim sel_tot As Double
sel_tot = 0
If Text1(8) = "" Or Text1(9) = "" Then
    MsgBox "Maaf data belum lengkap...", vbInformation, "Validasi Input"
Else
    update_stok
    update_stokBulanan
    With Data1.Recordset
        .Edit
        !jumlah = Text1(8)
        !satuan_rp = Text1(9)
        !total_rp = Val(Format(Text1(10), "###"))
        !hrg_dasar = !jumlah * !rate_dasar
        !net = !total_rp - !hrg_dasar
        net_akhir = !net
        .Update
        Data1.Refresh
        tutup_Trans
        cmd_awal
    End With
    update_bukubesar
End If
End Sub

Sub update_stok()
Dim sel_jml As Single
Dim sel_tot As Double
sel_jml = 0
sel_tot = 0
Data4.RecordSource = "select * from persediaan where simbol='" & Text1(7) & "'"
Data4.Refresh
With Data4.Recordset
    sel_jml = jml_awal - Val(Format(Text1(8), "###"))
'    sel_tot = sel_jml * Data1.Recordset!rate_dasar
    sel_tot = tot_awal - (Val(Format(Text1(8), "###")) * Val(Format(Text1(9), "###")))
    .Edit
    If Text1(0) = "BELI" Then
        !jumlah = !jumlah - sel_jml
        !total_rp = !total_rp - sel_tot
        If Text1(10) = 0 Then
            !rate_rp = 0
        Else
            !rate_rp = !total_rp / !jumlah
        End If
    Else
        !jumlah = !jumlah + sel_jml
        !total_rp = !jumlah * !rate_rp
    End If
    .Update
End With
Data4.Refresh
End Sub

Sub update_stokBulanan()
Dim sel_jml As Single
Dim sel_tot As Double
Dim bln_awal As String
Dim thn_awal As String
bln_awal = Month(Text1(3))
thn_awal = Year(Text1(3))
sel_jml = 0
sel_tot = 0
sel_jml = jml_awal - Val(Format(Text1(8), "###"))
'sel_tot = sel_jml * Data1.Recordset!rate_dasar
sel_tot = tot_awal - Val(Format(Text1(8), "###")) * Val(Format(Text1(9), "###"))
Data5.RecordSource = "select * from stok_bulanan where currency='" & Text1(7) & "' and bulan >= '" & bln_awal & "' and tahun >= '" & thn_awal & "'"
Data5.Refresh
With Data5.Recordset
    .MoveFirst
    Do While Not .EOF
        .Edit
        If Text1(0) = "BELI" Then
            !jml = !jml - sel_jml
            !total_rp = !total_rp - sel_tot
            If Text1(10) = 0 Then
                !rate = 0
            Else
                !rate = !total_rp / !jml
            End If
        Else
            !jml = !jml + sel_jml
            !total_rp = !jml * !rate
        End If
        .Update
        .MoveNext
    Loop
End With
Data5.Refresh

'update persediaan bb_bulanan
Data6.RecordSource = "select * from bb_bulanan where no_akun='1-160' and bulan >= '" & bln_awal & "' and tahun >= '" & thn_awal & "'"
Data6.Refresh
With Data6.Recordset
    .MoveFirst
    Do While Not .EOF
        .Edit
        If Text1(0) = "BELI" Then
            !saldo = !saldo - sel_tot
        Else
            !saldo = !saldo + sel_jml * Data5.Recordset!rate
        End If
        .Update
        .MoveNext
    Loop
End With
Data6.Refresh

'update stok harian
Data5.RecordSource = "select * from stok_harian where currency='" & Text1(7) & "' and cdate(tgl) >= '" & Text1(3) & "'"
Data5.Refresh
With Data5.Recordset
    .MoveFirst
    Do While Not .EOF
        .Edit
        If Text1(0) = "BELI" Then
            !jml = !jml - sel_jml
            !total_rp = !total_rp - sel_tot
            If Text1(10) = 0 Then
                !rate = 0
            Else
                !rate = !total_rp / !jml
            End If
        Else
            !jml = !jml + sel_jml
            !total_rp = !jml * !rate
        End If
        .Update
        .MoveNext
    Loop
End With
Data5.Refresh
End Sub

Sub update_bukubesar()
Dim sel_tot As Double
sel_tot = 0
sel_tot = tot_awal - Val(Format(Text1(10), "###"))
If Text1(11) = "KAS" Then
    If Text1(0) = "BELI" Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-110'   and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    Else
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-110'  or no_akun='4-100' or no_akun='3-200' or no_akun='5-100' and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    End If
ElseIf Text1(11) = "BANK" Then
    If Text1(0) = "BELI" Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-120'  and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    Else
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-120' or no_akun='4-100' or no_akun='3-200' or no_akun='5-100' and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    End If
ElseIf Text1(11) = "HUTANG" Then
    If Text1(0) = "BELI" Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='2-110' and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    Else
        Data6.RecordSource = "select * from bb_bulanan where no_akun='2-110' or no_akun='4-100' or no_akun='3-200' or no_akun='5-100' and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    End If
ElseIf Text1(11) = "PIUTANG" Then
    If Text1(0) = "BELI" Then
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-130'  and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    Else
        Data6.RecordSource = "select * from bb_bulanan where no_akun='1-130'  or no_akun='4-100' or no_akun='3-200' or no_akun='5-100' and bulan >= '" & Month(Text1(3)) & "' and tahun>='" & Year(Text1(3)) & "'"
    End If
End If
Data6.Refresh
With Data6.Recordset
    .MoveFirst
    Do While Not .EOF
        .Edit
        If !no_akun = "3-200" Or !no_akun = "5-100" Then
            !saldo = !saldo - (net_awal - net_akhir)
        Else
            If Text1(0) = "BELI" Then
                !saldo = !saldo + sel_tot
            Else
                !saldo = !saldo - sel_tot
            End If
        End If
        .Update
        .MoveNext
    Loop
End With
Data5.Refresh

'update kas harian
Data7.RecordSource = "select * from kas_harian where cdate(tgl)='" & Format(Text1(3), "mm/dd/yyyy") & "'"
Data7.Refresh
With Data7.Recordset
If Not .BOF Then
    .Edit
    If Text1(0) = "BELI" Then
        !saldo = !saldo + sel_tot
    Else
        !saldo = !saldo - sel_tot
    End If
    .Update
End If
End With

End Sub

