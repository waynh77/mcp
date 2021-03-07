VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Bonus_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Bonus"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "Bonus_frm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Caption         =   "DATA TRANSAKSI BONUS PEGAWAI"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   10320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2880
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10320
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   1
      Left            =   10320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4320
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Index           =   2
      Left            =   10320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5040
      Width           =   3735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   10320
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Data Data2 
      Caption         =   "SALDO BONUS"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "data peg"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data4 
      Caption         =   "rekap jurnal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data5 
      Caption         =   "update kas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   10320
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39901
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
            Picture         =   "Bonus_frm.frx":1CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":365C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":4FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":6980
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":8312
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":9CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":A97E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Bonus_frm.frx":B658
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
      MouseIcon       =   "Bonus_frm.frx":CFEA
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Bonus_frm.frx":D304
      Height          =   6375
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   7335
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   3200
      Columns(0).Caption=   "NO PEGAWAI"
      Columns(0).Name =   "NO PEGAWAI"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "no_peg"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1799
      Columns(1).Caption=   "TANGGAL"
      Columns(1).Name =   "TANGGAL"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "tgl"
      Columns(1).DataType=   7
      Columns(1).FieldLen=   256
      Columns(2).Width=   4101
      Columns(2).Caption=   "KETERANGAN"
      Columns(2).Name =   "KETERANGAN"
      Columns(2).CaptionAlignment=   2
      Columns(2).DataField=   "Ket"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "JUMLAH"
      Columns(3).Name =   "JUMLAH"
      Columns(3).Alignment=   1
      Columns(3).CaptionAlignment=   2
      Columns(3).DataField=   "Jml"
      Columns(3).DataType=   5
      Columns(3).NumberFormat=   "###,###.00"
      Columns(3).FieldLen=   256
      _ExtentX        =   12938
      _ExtentY        =   11245
      _StockProps     =   79
      Caption         =   "TRANSAKSI BONUS PEGAWAI"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSAKSI HUTANG BONUS"
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
      TabIndex        =   14
      Top             =   120
      Width           =   15255
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
      Left            =   7920
      TabIndex        =   13
      Top             =   2400
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
      Left            =   7920
      TabIndex        =   12
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KETERANGAN"
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
      Left            =   7920
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
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
      Left            =   7920
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KURANG BAYAR"
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
      Left            =   7920
      TabIndex        =   9
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   7920
      X2              =   14040
      Y1              =   4800
      Y2              =   4800
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
      Index           =   2
      Left            =   7920
      TabIndex        =   8
      Top             =   3360
      Width           =   2295
   End
End
Attribute VB_Name = "Bonus_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tambah As Boolean
Dim cek2 As Boolean
Dim nil_awal As Double

Private Sub Combo1_Click()
isi_peg
isi_grid
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub Data1_Reposition()
isi
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
Data3.Refresh
isi_cmb1
ISI_cmb2
End Sub

Private Sub Form_Load()
Call Db_bONUS
Kosong
Tutup
cmd_awal
End Sub

Private Sub SSDBGrid1_Click()
isi
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Toolbar1.Buttons(1).Caption = "Tambah" Then
        Buka
        Kosong
        cmd_Simpan
        tambah = True
        Combo1.Enabled = False
        Combo2.Enabled = True
        nil_awal = 0
    Else
        simpan
    End If
Case 2
    If Toolbar1.Buttons(2).Caption = "Edit" Then
        If Not Data1.Recordset.BOF Then
            Buka
            cmd_Simpan
            tambah = False
            nil_awal = 0
            nil_awal = Val(Text1(1))
            Combo1.Enabled = False
            Combo2.Enabled = False
        Else
            MsgBox "Maaf data kosong/belum dipilih...", vbInformation, "Validasi Data"
        End If
    Else
        cmd_awal
        Tutup
        Data1.Refresh
        isi_tot
        Combo1.Enabled = True
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
DTPicker1 = Date
Text1(1) = ""
'Text1(2) = ""
End Sub

Sub isi()
With Data1.Recordset
If Not .BOF And Data1.Enabled = True Then
    Combo1 = !no_peg
    DTPicker1 = !tgl
    Combo2 = !ket
    Text1(1) = Format(!jml, "###,###.00")
End If
If .BOF And Data1.Enabled = True Then
    Kosong
End If
End With
End Sub

Sub isi_cmb1()
Combo1.Clear
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        Combo1.AddItem !no_peg
        .MoveNext
    Loop
    Combo1.ListIndex = 0
Else
    Combo1 = ""
End If
End With
End Sub

Sub ISI_cmb2()
Combo2.Clear
Combo2.AddItem "BONUS"
Combo2.AddItem "BAYAR BONUS"
Combo2.ListIndex = 0
End Sub

Sub isi_peg()
With Data3.Recordset
If Not .BOF Then
    .MoveFirst
    Do Until !no_peg = Combo1 Or .EOF
        .MoveNext
    Loop
    Text1(0) = !nama_peg
Else
    Text1(0) = ""
End If
End With
End Sub

Sub isi_grid()
Data1.RecordSource = "select * from trans_bonus where no_peg='" & Combo1 & "' order by tgl desc"
Data1.Refresh
isi_tot
End Sub

Sub isi_tot()
Data2.RecordSource = "select * from saldo_bonus where no_peg='" & Combo1 & "'"
Data2.Refresh
If Not Data2.Recordset.BOF Then
    Text1(2) = Format(Data2.Recordset!saldo, "###,###.00")
Else
    Text1(2) = 0
End If
End Sub

Sub Tutup()
Text1(0).Enabled = False
DTPicker1.Enabled = False
Combo2.Enabled = False
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
End Sub

Sub Buka()
'DTPicker1.Enabled = True
Combo2.Enabled = True
Text1(1).Enabled = True
Text1(1) = Format(Text1(1), "###")
End Sub

Sub simpan()
Cek_Input
If cek2 = False Then
    MsgBox "Input tidak valid, mohon diperiksa kembali", vbInformation, "Validasi Input"
Else
    With Data1.Recordset
        If tambah = True Then
            .AddNew
            !no_peg = Combo1
            !tgl = DTPicker1
            !ket = Combo2
            !jml = Text1(1)
            .Update
            cetak_bukti
        Else
            .Edit
            !no_peg = Combo1
            !tgl = DTPicker1
            !ket = Combo2
            !jml = Text1(1)
            .Update
        End If
    End With
    UPDATE_JURNAL
    update_saldo
    Update_KAs
    edit_BB
    Tutup
    cmd_awal
    Data1.Refresh
    isi_tot
    Combo1.Enabled = True
End If
End Sub

Sub update_saldo()
Dim sal As Double
With Data2.Recordset
If Not .BOF Then
    sal = !saldo
    .Edit
    If Combo2 = "BONUS" Then
        !saldo = sal - (nil_awal - Val(Text1(1)))
    Else
        !saldo = sal + (nil_awal - Val(Text1(1)))
    End If
    .Update
Else
    .AddNew
    !no_peg = Combo1
    If Combo2 = "BONUS" Then
        !saldo = Val(Text1(1))
    Else
        !saldo = -Val(Text1(1))
    End If
    .Update
End If
End With
End Sub

Sub UPDATE_JURNAL()
With Data4.Recordset
    .AddNew
    !tgl = DTPicker1
    !jam = Time
    !no_akun = "6-018"
    !jml = Text1(1)
    If Combo2 = "BONUS" Then
        !dk = "KREDIT"
        If tambah = True Then
            !ket = "BONUS " & Text1(0)
            !jml = Text1(1)
        Else
            !ket = "REVISI BONUS " & Text1(0)
            !jml = nil_awal - Text1(1)
        End If
    Else
        !dk = "DEBET"
        If tambah = True Then
            !ket = "BAYAR BONUS " & Text1(0)
            !jml = Text1(1)
        Else
            !ket = "REVISI BAYAR BONUS " & Text1(0)
            !jml = nil_awal - Text1(1)
        End If
    End If
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = "1-131"
    .Update

    .AddNew
    !tgl = DTPicker1
    !jam = Time
    !no_akun = "2-120"
    If Combo2 = "BONUS" Then
        !dk = "DEBET"
        If tambah = True Then
            !ket = "BONUS " & Text1(0)
            !jml = Text1(1)
        Else
            !ket = "REVISI BONUS " & Text1(0)
            !jml = nil_awal - Text1(1)
        End If
    Else
        !dk = "KREDIT"
        If tambah = True Then
            !ket = "BAYAR BONUS " & Text1(0)
            !jml = Text1(1)
        Else
            !ket = "REVISI BAYAR BONUS " & Text1(0)
            !jml = nil_awal - Text1(1)
        End If
    End If
    !user = Mid(MoneyChanger.Label1.Caption, 13)
    !sumber_akun = "1-110"
    .Update
End With
End Sub

Sub Update_KAs()
'update kas harian
Data5.RecordSource = "select * from kas_harian where cdate(tgl)>='" & DTPicker1 & "'"
Data5.Refresh
With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    If tambah = True Then
        .Edit
        If Combo2 = "BONUS" Then
            '!saldo = !saldo + Text1(1)
        Else
            !saldo = !saldo - Text1(1)
        End If
        .Update
    Else
        Do While Not .EOF
            .Edit
            If Combo2 = "BONUS" Then
            '    !saldo = !saldo - (nil_awal - Val(Text1(1)))
            Else
                !saldo = !saldo + (nil_awal - Val(Text1(1)))
            End If
            .Update
            .MoveNext
        Loop
    End If
End If
End With

End Sub

Sub edit_BB()
Dim sal As Double
Dim sel As Double
Dim dk As String
sel = 0
sel = Val(Format(Text1(1), "###")) - nil_awal
Data5.RecordSource = "select * from bb_bulanan where no_akun='6-018' order by tahun desc,bulan desc"
Data5.Refresh
With Data5.Recordset
    If Combo2 = "BONUS" Then
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal + sel
        .Update
    End If
End With

Data5.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
Data5.Refresh
With Data5.Recordset
    If Combo2 = "BAYAR BONUS" Then
        .MoveFirst
        sal = !saldo
        .Edit
        !saldo = sal - sel
        .Update
    End If
End With

Data5.RecordSource = "select * from bb_bulanan where no_akun='2-120' order by tahun desc,bulan desc"
Data5.Refresh
With Data5.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    If Combo2 = "BONUS" Then
        !saldo = sal + sel
    Else
        !saldo = sal - sel
    End If
    .Update
End With
Data5.Refresh

'update laba rugi
If Combo2 = "BONUS" Then
Data5.RecordSource = "select * from bb_bulanan where no_akun='3-200' order by tahun desc,bulan desc"
Data5.Refresh
With Data5.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    !saldo = sal - sel
    .Update
End With
End If
End Sub

Sub Cek_Input()
If Combo1 = "" Or Combo2 = "" Or Text1(0) = "" Then
    cek2 = False
Else
    cek2 = True
End If
End Sub

Sub Hapus()
With Data1.Recordset
    If Not .BOF Then
        X = MsgBox("Apakah anda yakin menghapus data?", vbYesNo, "Hapus Data")
        If X = vbYes Then
            Hapus_Kas
            .Delete
            Kosong
            Data1.Refresh
        End If
    Else
        MsgBox "Data masih kosong/belum dipilih...", vbInformation, "Validasi Data"
    End If
End With
End Sub

Sub Hapus_Kas()
Dim sal As Double
Dim sel As Double
Dim dk As String
sel = 0
sel = Val(Format(Text1(1), "###")) - nil_awal
'update bb
Data5.RecordSource = "select * from bb_bulanan where no_akun='1-110' order by tahun desc,bulan desc"
Data5.Refresh
With Data5.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    If Combo2 = "BAYAR BONUS" Then
        !saldo = sal - sel
    Else
        !saldo = sal + sel
    End If
    .Update
End With
Data5.RecordSource = "select * from bb_bulanan where no_akun='1-131' order by tahun desc,bulan desc"
Data5.Refresh
With Data5.Recordset
    .MoveFirst
    sal = !saldo
    .Edit
    If Combo2 = "BONUS" Then
        !saldo = sal - sel
    Else
        !saldo = sal + sel
    End If
    .Update
End With
Data5.Refresh

'update kas harian
Data5.RecordSource = "select * from kas_harian where cdate(tgl)>='" & DTPicker1 & "'"
Data5.Refresh
With Data5.Recordset
If Not .BOF Then
    .MoveFirst
    Do While Not .EOF
        .Edit
        If Combo2 = "BONUS" Then
            !saldo = !saldo + Val(Format(Text1(1), "###.00"))
        Else
            !saldo = !saldo - Val(Format(Text1(1), "###.00"))
        End If
        .Update
        .MoveNext
    Loop
End If
End With

With Data2.Recordset
If Not .BOF Then
    sal = !saldo
    .Edit
    If Combo2 = "BONUS" Then
        !saldo = sal - Val(Format(Text1(1), "###.00"))
    Else
        !saldo = sal + Val(Format(Text1(1), "###.00"))
    End If
    .Update
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
    .Buttons(1).Visible = True
    .Buttons(2).Visible = True
    .Buttons(3).Visible = True
    .Buttons(4).Visible = False
    .Buttons(5).Visible = True
    .Buttons(6).Visible = False
    .Buttons(7).Visible = True
End With
Data1.Enabled = True
SSDBGrid1.Enabled = True
End Sub

Sub cmd_Simpan()
With Toolbar1
    .Buttons(1).Image = 8
    .Buttons(2).Image = 3
    .Buttons(1).Caption = "Simpan"
    .Buttons(2).Caption = "Batal"
    .Buttons(1).ToolTipText = "Simpan Data"
    .Buttons(2).ToolTipText = "Batal Data"
    .Buttons(1).Visible = True
    .Buttons(2).Visible = True
    .Buttons(3).Visible = False
    .Buttons(4).Visible = False
    .Buttons(5).Visible = False
    .Buttons(6).Visible = False
    .Buttons(7).Visible = False
End With
Data1.Enabled = False
SSDBGrid1.Enabled = False
End Sub

Sub cetak_bukti()
With Bukti_Bonus
    .Field1 = Format(Date, "d mmm yyyy")
    .Field2 = Combo1
    .Field3 = Text1(0)
    .Field4 = Data3.Recordset!Div
    .Field5 = Data3.Recordset!Jab
    .Field7 = Format(Text1(1), "###,###.00")
    .Field6 = Format(Text1(2), "###,###.00")
    If Combo2 = "BONUS" Then
        .Label70.Caption = "Bonus Baru"
        .Field8 = Format(Val(Format(.Field6, "###.00")) + Val(Format(.Field7, "###.00")), "###,###.00")
    Else
        .Label70.Caption = "Bayar Bonus"
        .Field8 = Format(Val(Format(.Field6, "###.00")) - Val(Format(.Field7, "###.00")), "###,###.00")
    End If
    .Label80 = Text1(0)
    .Label81.Caption = "Keterangan : " & Combo2
    .Show
    .WindowState = 2
End With
End Sub
