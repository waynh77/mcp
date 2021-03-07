VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form LAporan_frm 
   BackColor       =   &H00FF8080&
   Caption         =   "NAMA LAPORAN"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "LAporan_frm.frx":0000
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
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   11010
      Left            =   14550
      TabIndex        =   8
      Top             =   0
      Width           =   690
      _ExtentX        =   1217
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
      MouseIcon       =   "LAporan_frm.frx":3482
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
      Top             =   2400
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
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7200
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES CARI DATA"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   1800
      Width           =   5535
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
   Begin SSDataWidgets_B.SSDBGrid SSDBGrid1 
      Bindings        =   "LAporan_frm.frx":379C
      Height          =   6015
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   10575
      _Version        =   196614
      BevelColorFace  =   192
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   18653
      _ExtentY        =   10610
      _StockProps     =   79
      Caption         =   "DATA LAPORAN PEMBELIAN"
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Index           =   0
      Left            =   7200
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   59113473
      CurrentDate     =   39864
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
            Picture         =   "LAporan_frm.frx":37B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":5142
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":6AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":8466
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":9DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":B78A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":C464
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LAporan_frm.frx":D13E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Index           =   1
      Left            =   8760
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   59113473
      CurrentDate     =   39864
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
      TabIndex        =   7
      Top             =   120
      Width           =   15255
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
      Left            =   4800
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
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
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "LAporan_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sal_awal As Double
Dim sal_akhir As Double
Dim tot_pers As Double
Dim profit As Double

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

Private Sub Command1_Click()
Dim curr As String
Dim jml_awal As Double
Dim rate_stok As Double
Dim rate_beli As Double
Dim tot_beli As Double
Dim tot_jual As Double
Dim tot_stok As Double

If Me.Caption = "REKAP JURNAL" Then
    Select Case Combo2.ListIndex
        Case 0
            Data1.RecordSource = "select TGL,JAM,REKAP_JURNAL.NO_AKUN,NAMA_AKUN,rekap_jurnal.DK,JML,KET,sumber_akun,USER from rekap_jurnal,tbl_akun where rekap_jurnal.no_akun=tbl_akun.no_akun and cdate(tgl)='" & DTPicker1(0) & "'"
        Case 1
            Data1.RecordSource = "select TGL,JAM,REKAP_JURNAL.NO_AKUN,NAMA_AKUN,rekap_jurnal.DK,JML,KET,sumber_akun,USER from rekap_jurnal,tbl_akun where rekap_jurnal.no_akun=tbl_akun.no_akun and cdate(tgl)>='" & DTPicker1(0) & "' and cdate(tgl)<='" & DTPicker1(1) & "'"
    End Select
    Data1.Refresh
ElseIf Me.Caption = "PROFIT PENJUALAN" Then
    Select Case Combo2.ListIndex
        Case 0
            Data2.RecordSource = "select simbol,sum(jumlah)as jml, sum(total_rp)as total,total/jml as rata from trans_jualbeli where status='JUAL' and cdate(tgl)='" & DTPicker1(0) & "' group by simbol"
        Case 1
            Data2.RecordSource = "select simbol,sum(jumlah)as jml, sum(total_rp)as total,total/jml as rata from trans_jualbeli  where status='JUAL' and cdate(tgl)>='" & DTPicker1(0) & "' and cdate(tgl)<='" & DTPicker1(1) & "' group by simbol"
    End Select
    Data2.Refresh
    Data1.RecordSource = "temp_profit"
    Data1.Refresh
    hapus_temp
    With Data2.Recordset
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                With Data1.Recordset
                    .AddNew
                    !simbol = Data2.Recordset!simbol
                    !jml_jual = Data2.Recordset!jml
                    !tot_jual = Data2.Recordset!total
                    If !jml_jual = 0 Then
                        !rate_jual = 0
                    Else
                        !rate_jual = Data2.Recordset!rata
                    End If
                    Data3.RecordSource = "select * from stok_harian where currency='" & Data2.Recordset!simbol & "' and cdate(tgl)='" & DTPicker1(0) & "'"
                    Data3.Refresh
                    !rate_beli = Data3.Recordset!rate
                    !tot_beli = !jml_jual * !rate_beli
                    !profit = !tot_jual - !tot_beli
                    .Update
                End With
                .MoveNext
            Loop
        End If
    End With
    Data1.Refresh
ElseIf Me.Caption = "LAPORAN REKAP STOK TRANSAKSI HARIAN" Then
    Data1.RecordSource = "temp_Uka"
    Data2.RecordSource = "SELEct simbol,status,sum(total_rp) as total,sum(jumlah) as jml,(sum(total_rp)/sum(jumlah)) as rj,void from trans_jualbeli where Status='BELI' and CDATE(tgl)='" & DTPicker1(0) & "' and void=0 group by simbol,status,void order by simbol asc"
    Data3.RecordSource = "SELEct simbol,status,sum(total_rp) as total,sum(jumlah) as jml,(sum(total_rp)/sum(jumlah)) as rj,void from trans_jualbeli where Status='JUAL' and cdate(tgl)='" & DTPicker1(0) & "' and void=0 group by simbol,status,void order by simbol asc"
    Data4.RecordSource = "select * from stok_harian where cdate(tgl)<'" & DTPicker1(0) & "' order by currency asc,tgl desc"
    Data5.RecordSource = "select * from stok_harian where cdate(tgl)='" & DTPicker1(0) & "' order by currency asc"
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    Data5.Refresh
    hapus_temp
    tot_pers = 0
    With Data4.Recordset
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If curr <> !Currency Then
                    curr = !Currency
                    jml_awal = !jml
                    saldo_awal = !total_rp
                    Data2.Refresh
                    With Data2.Recordset
                        jml_beli = 0
                        rate_beli = 0
                        If Not .BOF Then
                            .MoveFirst
                            tot_beli = 0
                            Do While Not .EOF
                                If !simbol = curr Then
                                    jml_beli = !jml
                                    If !total = 0 Then
                                        rate_beli = 0
                                    Else
                                        rate_beli = !total / !jml
                                    End If
                                    tot_beli = !total
                                    .MoveLast
                                End If
                                .MoveNext
                            Loop
                        End If
                    End With
                    If rate_beli = 0 Then
                        rate_beli = Data4.Recordset!rate
                    End If
                    Data3.Refresh
                    With Data3.Recordset
                        jml_jual = 0
                        rate_jual = 0
                        If Not .BOF Then
                            .MoveFirst
                            tot_jual = 0
                            Do While Not .EOF
                                If !simbol = curr Then
                                    jml_jual = !jml
                                    If !total = 0 Then
                                        rate_jual = 0
                                    Else
                                        rate_jual = !total / !jml
                                    End If
                                    tot_jual = !total
                                    .MoveLast
                                End If
                                .MoveNext
                            Loop
                        End If
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
                        !ttl_beli = tot_beli
                        !ttl_jual = tot_jual
                        !ttl_stok = tot_stok
                        .Update
                    End With
                End If
                .MoveNext
            Loop
        Else
            With Data2.Recordset
                If Not .BOF Then
                    .MoveFirst
                    Do While Not .EOF
                        curr = !simbol
                        jml_awal = 0
                        saldo_awal = 0
                        jml_beli = !jml
                        If !total = 0 Then
                            rate_beli = 0
                        Else
                            rate_beli = !total / !jml
                        End If
                        tot_beli = !total
                        If rate_beli = 0 Then
                            rate_beli = Data4.Recordset!rate
                        End If
                        Data3.Refresh
                        With Data3.Recordset
                            jml_jual = 0
                            rate_jual = 0
                            If Not .BOF Then
                                .MoveFirst
                                tot_jual = 0
                                Do While Not .EOF
                                    If !simbol = curr Then
                                        jml_jual = !jml
                                        If !total = 0 Then
                                            rate_jual = 0
                                        Else
                                            rate_jual = !total / !jml
                                        End If
                                        tot_jual = !total
                                        .MoveLast
                                    End If
                                    .MoveNext
                                Loop
                            End If
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
                            !ttl_beli = tot_beli
                            !ttl_jual = tot_jual
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
    Data2.RecordSource = "select * from kas_harian where cdate(tgl)<'" & DTPicker1(0) & "' order by tgl desc"
    Data2.Refresh
    If Not Data2.Recordset.BOF Then
        sal_awal = Data2.Recordset!saldo
    Else
        sal_awal = 0
    End If
    Data3.RecordSource = "select * from kas_harian where cdate(tgl)='" & DTPicker1(0) & "' order by tgl desc"
    Data3.Refresh
    If Not Data3.Recordset.BOF Then
        sal_akhir = Data3.Recordset!saldo
    Else
        sal_akhir = 0
    End If
    
    'hit profit
    Data4.RecordSource = "SELEct simbol,sum(jumlah) as jml, sum(hrg_dasar) as tot_dsr,sum(total_rp) as tot_jual,(sum(hrg_dasar)/sum(jumlah)) as rd ,(sum(total_rp)/sum(jumlah)) as rj, sum(net) as profit from trans_jualbeli  where status='JUAL' and cdate(tgl)='" & DTPicker1(0) & "'group by simbol"
    Data4.Refresh
    With Data4.Recordset
    If Not .BOF Then
        .MoveFirst
        profit = 0
        Do While Not .EOF
            Data5.RecordSource = "select * from stok_harian where currency='" & !simbol & "' and cdate(tgl)='" & DTPicker1(0) & "'"
            Data5.Refresh
            profit = profit + (!tot_jual - (!jml * Data5.Recordset!rate))
            .MoveNext
        Loop
    End If
    End With
End If
If Not Data1.Recordset.BOF Then
    Toolbar1.Buttons(1).Enabled = True
End If
End Sub

Private Sub Form_Load()
ISI_cmb
DTPicker1(0) = Date
DTPicker1(1) = Date
Call dB_lAP
If Me.Caption = "REKAP JURNAL" Then
    Data1.RecordSource = "rekap_jurnal"
End If
Toolbar1.Buttons(1).Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim aktiva As Double
Dim penyusutan As Double
Dim beli As Double
Select Case Button.Index
Case 1
    If Me.Caption <> "LAPORAN REKAP STOK TRANSAKSI HARIAN" And Me.Caption <> "PROFIT PENJUALAN" Then
        Command1_Click
    End If
    If Data1.Recordset.BOF Then
        MsgBox "Maaf data tidak diketemukan...", vbInformation, "Data Kosong"
    Else
        
        If Me.Caption = "REKAP JURNAL" Then
            With Rekap_Jurnal
                .DAODataControl1.DatabaseName = Data1.DatabaseName
                .DAODataControl1.RecordSource = Data1.RecordSource
                .Show
            End With
        ElseIf Me.Caption = "PROFIT PENJUALAN" Then
            With LProfit
                If Combo2.ListIndex = 0 Then
                    .Field1 = Format(DTPicker1(0), "d mmm yyyy")
                Else
                    .Field1 = Format(DTPicker1(0), "d mmm yyyy") & " - " & Format(DTPicker1(1), "d mmm yyyy")
                End If
                .DAODataControl1.DatabaseName = Data1.DatabaseName
                .DAODataControl1.RecordSource = Data1.RecordSource
                .Show
            End With
        ElseIf Me.Caption = "LAPORAN REKAP STOK TRANSAKSI HARIAN" Then
            With UKA_bln
                Data1.Refresh
                .DAODataControl1.DatabaseName = Data1.DatabaseName
                .DAODataControl1.RecordSource = Data1.RecordSource
                .Field1 = Format(DTPicker1(0), "d mmmm yyyy")
                .Field21 = Format(sal_awal, "###,###.00")
                .Field28 = Format(profit, "###,###.00")
                .Show
            End With
        End If
    End If
Case 2
    Unload Me
End Select
End Sub

Sub ISI_cmb()
Combo2.Clear
Combo2.AddItem "Harian"
Combo2.AddItem "Periodik"
Combo2.ListIndex = 0
End Sub

Sub hapus_temp()
Dim C As Single
Data1.Enabled = False
'Data1.RecordSource = "temp_profit"
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
    End If
End With
Data1.Enabled = True
Data1.Refresh
End Sub

