VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form Lap_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "Nama Laporan"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6480
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   840
      Width           =   735
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
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES CARI DATA"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1320
      Width           =   5535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9120
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   840
      Width           =   1215
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
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "Lap_frm.frx":0000
      Height          =   6495
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   10575
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   18653
      _ExtentY        =   11456
      _StockProps     =   79
      Caption         =   "DATA LAPORAN PEMBELIAN"
      ForeColor       =   16777215
      BackColor       =   8421504
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
            Picture         =   "Lap_frm.frx":0014
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":19A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":3338
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":4CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":665C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":7FEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":8CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lap_frm.frx":99A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   11010
      Left            =   14430
      TabIndex        =   3
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
            Caption         =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Tutup"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Lap_frm.frx":B334
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
      Index           =   3
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   1575
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
      Index           =   1
      Left            =   7440
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA LAPORAN"
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
      TabIndex        =   4
      Top             =   120
      Width           =   15255
   End
End
Attribute VB_Name = "Lap_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim thn_awal As String
Dim bln_awal As String
Dim jml_awal As Double
Dim saldo_awal As Double
Dim jml_beli As Double
Dim rate_beli As Double
Dim jml_jual As Double
Dim rate_jual As Double
Dim rate_stok As Double
Dim curr As String
Dim tot_beli As Double
Dim tot_jual As Double
Dim tot_stok As Double


If Me.Caption = "NERACA" Then
    Data1.RecordSource = "select * from bb_bulanan where  tahun ='" & Combo2 & "' and bulan='" & Combo1 & "'"
    Data1.Refresh

ElseIf Me.Caption = "LABA-RUGI" Then
    If Combo1 <> "1" Then
        bln_awal = Combo1 - 1
        thn_awal = Combo2
    Else
        bln_awal = "12"
        thn_awal = Combo2 - 1
    End If
    Data1.RecordSource = "select * from bb_bulanan where tahun ='" & Combo2 & "' and bulan='" & Combo1 & "'"
    Data2.RecordSource = "select * from trans_jualbeli  where status='BELI' AND month(tgl)='" & Combo1 & "' and year(tgl)='" & Combo2 & "'"
    Data3.RecordSource = "persediaan"
    Data4.RecordSource = "select * from bb_bulanan where tahun='" & thn_awal & "' and bulan='" & bln_awal & "'"
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh

ElseIf Me.Caption = "LAPORAN UANG KERTAS ASING" Then
    If Combo1 <> "1" Then
        bln_awal = Combo1 - 1
        thn_awal = Combo2
    Else
        bln_awal = "12"
        thn_awal = Combo2 - 1
    End If
    Data1.RecordSource = "temp_Uka"
    Data2.RecordSource = "SELEct simbol,status,sum(total_rp)as tot_beli,sum(jumlah) as jml,(sum(total_rp)/sum(jumlah)) as rj from trans_jualbeli where Status='BELI' and month(tgl)='" & Combo1 & "' and year(tgl)='" & Combo2 & "' group by simbol,status order by simbol asc"
    Data3.RecordSource = "SELEct simbol,status,sum(total_rp)as tot_jual,sum(jumlah) as jml,(sum(total_rp)/sum(jumlah)) as rj from trans_jualbeli where Status='JUAL' and month(tgl)='" & Combo1 & "' and year(tgl)='" & Combo2 & "' group by simbol,status order by simbol asc"
    Data4.RecordSource = "select * from stok_bulanan where bulan='" & bln_awal & "' and tahun='" & thn_awal & "' order by currency asc"
    Data5.RecordSource = "select * from stok_bulanan where bulan='" & Combo1 & "' and tahun='" & Combo2 & "' order by currency asc"
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    Data5.Refresh
    hapus_temp
    With Data4.Recordset
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                curr = !Currency
                jml_awal = !jml
                saldo_awal = !total_rp
                Data2.Refresh
                With Data2.Recordset
                    .MoveFirst
                    jml_beli = 0
                    tot_beli = 0
                    rate_beli = 0
                    Do While Not .EOF
                        If !simbol = curr Then
                            jml_beli = !jml
                            rate_beli = !rj
                            tot_beli = !tot_beli
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                End With
                Data3.Refresh
                With Data3.Recordset
                    .MoveFirst
                    jml_jual = 0
                    rate_jual = 0
                    tot_jual = 0
                    Do While Not .EOF
                        If !simbol = curr Then
                            jml_jual = !jml
                            rate_jual = !rj
                            tot_jual = !tot_jual
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                End With
                With Data1.Recordset
                    .AddNew
                    !curr = curr
                    !jml_awal = jml_awal
                    !saldo_awal = saldo_awal
                    !jml_beli = jml_beli
                    !rate_beli = rate_beli
                    !jml_jual = jml_jual
                    !rate_jual = rate_jual
                    !ttl_beli = tot_beli
                    !ttl_jual = tot_jual
                    With Data5.Recordset
                    If Not .BOF Then
                        .MoveFirst
                        tot_stok = 0
                        Do While Not .EOF
                            If curr = !Currency Then
                                rate_stok = !rate
                                tot_stok = !total_rp
                                .MoveLast
                            End If
                            .MoveNext
                        Loop
                    End If
                    End With
                    !rate_stok = rate_stok
                    !ttl_stok = tot_stok
                    .Update
                End With
                .MoveNext
            Loop
        Else
            With Data2.Recordset
                If Not .BOF Then
                    .MoveFirst
                    Do While Not .EOF
                        curr = !simbol
                        jml_awal = 0
                        saldo_awal = o
                        tot_beli = !tot_beli
                        jml_beli = !jml
                        rate_beli = !rj
                        Data3.Refresh
                        With Data3.Recordset
                            .MoveFirst
                            jml_jual = 0
                            rate_jual = 0
                            tot_jual = 0
                            Do While Not .EOF
                                If !simbol = curr Then
                                    jml_jual = !jml
                                    tot_jual = !tot_jual
                                    rate_jual = !rj
                                    .MoveLast
                                End If
                                .MoveNext
                            Loop
                        End With
                        With Data1.Recordset
                            .AddNew
                            !curr = curr
                            !jml_awal = jml_awal
                            !saldo_awal = saldo_awal
                            !jml_beli = jml_beli
                            !rate_beli = rate_beli
                            !jml_jual = jml_jual
                            !rate_jual = rate_jual
                            !ttl_beli = tot_beli
                            !ttl_jual = tot_jual
                            With Data5.Recordset
                            If Not .BOF Then
                                .MoveFirst
                                tot_stok = 0
                                Do While Not .EOF
                                    If curr = !Currency Then
                                        rate_stok = !rate
                                        tot_stok = !total_rp
                                        .MoveLast
                                    End If
                                    .MoveNext
                                Loop
                            End If
                            End With
                            !rate_stok = rate_stok
                            !ttl_stok = tot_stok
                            .Update
                        End With
                        .MoveNext
                    Loop
                End If
            End With
        End If
    End With
    Data1.Refresh
'    Data4.RecordSource = "select * from bb_bulanan where no_akun='1-160' and bulan='" & Combo1 & "' and tahun='" & Combo2 & "'"
'    Data4.Refresh
End If
End Sub

Private Sub Form_Load()
Call dB_Lapor
isi_cmb1
ISI_cmb2
End Sub

Sub isi_cmb1()
Dim X As Byte
X = 0
Combo1.Clear
Do Until X = 12
    X = X + 1
    Combo1.AddItem X
Loop
Combo1 = Month(Date)
End Sub

Sub ISI_cmb2()
Dim X As Integer
X = 2008
Combo2.Clear
Do Until X = 2108
    X = X + 1
    Combo2.AddItem X
Loop
Combo2 = Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Caption = "LAPORAN UANG KERTAS ASING" Then
    hapus_temp
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim aktiva As Double
Dim penyusutan As Double
Dim beli As Double
Select Case Button.Index
Case 1
    If Me.Caption <> "LAPORAN UANG KERTAS ASING" Then
        Command1_Click
    End If
    If Data1.Recordset.BOF Then
        MsgBox "Maaf data tidak diketemukan...", vbInformation, "Data Kosong"
    Else
        If Me.Caption = "NERACA" Then
            With Data1.Recordset
                If Not .BOF Then
                    'ambil saldo kas
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-110" Then
                            NERACA.Field2 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil saldo bank
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-120" Then
                            NERACA.Field3 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil saldo persediaan
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-160" Then
                            NERACA.Field4 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil saldo piutang
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-130" Then
                            NERACA.Field10 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil saldo pinjaman
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-131" Then
                            NERACA.Field24 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil saldo aktiva tetap
                    aktiva = 0
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-210" Then
                            aktiva = aktiva + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-220" Then
                            aktiva = aktiva + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-230" Then
                            aktiva = aktiva + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-240" Then
                            aktiva = aktiva + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    NERACA.Field6 = Format(aktiva, "###,###.00")
                    'ambil saldo penyusutan aktiva tetap
                    aktiva = 0
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-270" Then
                            penyusutan = penyusutan + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-280" Then
                            penyusutan = penyusutan + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-290" Then
                            penyusutan = penyusutan + !saldo
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    NERACA.Field7 = Format(penyusutan, "###,###.00")
                    'ambil saldo hutang
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "2-110" Then
                            NERACA.Field14 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "2-120" Then
                            NERACA.Field16 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil MODAL
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "3-100" Then
                            NERACA.Field18 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil laba
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "3-200" Then
                            'NERACA.Field19 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil PRIVE/EKUITAS LAIN2
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "3-110" Then
                            NERACA.Field22 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    
                End If
            End With
            NERACA.Field1 = MonthName(Combo1) & " " & Combo2
            NERACA.Show
            
        ElseIf Me.Caption = "LABA-RUGI" Then
            With Data1.Recordset
                    'ambil penjualan
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "4-100" Then
                            LABA_RUGI.Field2 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil persediaan awal
                    With Data4.Recordset
                    If Not .BOF Then
                        .MoveFirst
                        Do While Not .EOF
                            If !no_akun = "1-160" Then
                                LABA_RUGI.Field4 = Format(Val(!saldo), "###,###.00")
                                .MoveLast
                            End If
                            .MoveNext
                        Loop
                    Else
                        LABA_RUGI.Field4 = 0
                    End If
                    End With
                    'ambil pembelian
                    With Data2.Recordset
                        If Not .BOF Then
                            .MoveFirst
                            beli = 0
                            Do While Not .EOF
                                beli = beli + !total_rp
                                .MoveNext
                            Loop
                            LABA_RUGI.Field5 = Format(Val(beli), "###,###.00")
                        Else
                            LABA_RUGI.Field5 = 0
                        End If
                    End With
                    'ambil saldo akhir persediaan
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "1-160" Then
                            LABA_RUGI.Field6 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil gaji
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-011" Then
                            LABA_RUGI.Field7 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil sewa
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-015" Then
                            LABA_RUGI.Field8 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil promosi
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-018" Then
                            LABA_RUGI.Field9 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil listrik
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-012" Then
                            LABA_RUGI.Field10 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil transport
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-020" Then
                            LABA_RUGI.Field11 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil PENYUSUTAN
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-014" Then
                            LABA_RUGI.Field12 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil lain-lain
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "6-019" Then
                            LABA_RUGI.Field13 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil PEndapatan bunga
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "7-110" Then
                            LABA_RUGI.Field14 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
                    'ambil pajak
                    .MoveFirst
                    Do While Not .EOF
                        If !no_akun = "2-112" Then
                            LABA_RUGI.Field27 = Format(Val(!saldo), "###,###.00")
                            .MoveLast
                        End If
                        .MoveNext
                    Loop
            End With
            LABA_RUGI.Field1 = MonthName(Combo1) & " " & Combo2
            LABA_RUGI.Show
            
        ElseIf Me.Caption = "LAPORAN UANG KERTAS ASING" Then
            With UKA
                .DAODataControl1.DatabaseName = Data1.DatabaseName
                .DAODataControl1.RecordSource = Data1.RecordSource
                .Field1 = MonthName(Combo1) & " " & Combo2
                .Show
            End With
        End If
    End If

Case 2
    Unload Me
End Select
End Sub

Sub hapus_temp()
Dim C As Single
Data1.Enabled = False
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
    End If
End With
Data1.Enabled = True
Data1.Refresh
End Sub
