Attribute VB_Name = "MCP_mod"

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim db As String

Public cetak_ulang As Boolean
Public ctk_inv As Boolean

Private Declare Function CreateRoundRectRgn _
Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, _
ByVal X2 As Long, _
ByVal Y2 As Long, _
ByVal X3 As Long, _
ByVal Y3 As Long) As Long

Sub buka_Db()
db = App.Path + "\DBMCP-recover.mdb"
'db = "\\Namora-31c4a3e9\Program Money Changer\dbmcp-recover.mdb"
'db = "DBMCP"
End Sub

Public Sub DB_PEg()
buka_Db
With Peg_frm
.Data1.DatabaseName = db
.Data1.RecordSource = "ms_pegawai"
.Div.DatabaseName = db
.Jab.DatabaseName = db
.gaji.DatabaseName = db
.Div.RecordSource = "tbl_divisi"
.Jab.RecordSource = "tbl_jabatan"
.gaji.RecordSource = "tbl_gaji"
.Data1.DefaultType = 2
.Div.DefaultType = 2
.Jab.DefaultType = 2
End With
End Sub

Public Sub db_DivJab()
buka_Db
DivJab_frm.Data1.DatabaseName = db
DivJab_frm.Data1.DefaultType = 2
End Sub

Public Sub db_Nasabah()
buka_Db
With Nasabah_frm
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data1.RecordSource = "msNasabah_Perseorangan"
.Data2.RecordSource = "msnasabah_perusahaan"
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data1.Refresh
.Data2.Refresh
End With
End Sub

Public Sub db_JenisAkun()
buka_Db
JenisAkun_frm.Data1.DatabaseName = db
JenisAkun_frm.Data1.RecordSource = "tblJenis_Akun"
JenisAkun_frm.Data1.DefaultType = 2
JenisAkun_frm.Data1.Refresh
End Sub

Public Sub db_SubAkun()
buka_Db
With SubAkun_frm
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data1.RecordSource = "tblsub_akun"
.Data2.RecordSource = "tbljenis_akun"
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data1.Refresh
.Data2.Refresh
End With
End Sub

Public Sub db_TabelAkun()
buka_Db
With TabelAkun_frm
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data3.DatabaseName = db
.Data4.DatabaseName = db
.Data1.RecordSource = "select * from Tbl_akun order by no_akun asc "
.Data2.RecordSource = "tbljenis_akun"
.Data3.RecordSource = "tblsub_akun"
.Data4.RecordSource = "BB_BULANAN"
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data4.DefaultType = 2
.Data1.Refresh
.Data2.Refresh
.Data3.Refresh
.Data4.Refresh
End With
End Sub

Public Sub db_Currency()
buka_Db
With Currency_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data1.DatabaseName = db
.Data1.RecordSource = "ms_currency"
.Data1.Refresh
.Data2.DatabaseName = db
.Data2.RecordSource = "persediaan"
.Data2.Refresh
.Data3.DatabaseName = db
.Data3.RecordSource = "STOK_BULANAN"
.Data3.Refresh
End With
End Sub

Public Sub db_TblBank()
buka_Db
With TblBank_frm
.Data1.DefaultType = 2
.Data1.DatabaseName = db
.Data1.RecordSource = "tbl_bank"
.Data1.Refresh
End With
End Sub

Public Sub db_Rekening()
buka_Db
With Rek_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data1.RecordSource = "MS_rekening"
.Data2.RecordSource = "tbl_bank"
.Data1.Refresh
.Data2.Refresh
End With
End Sub

Sub db_beli()
buka_Db
With Beli_frm
.Data1.DefaultType = 2
.Data10.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data4.DefaultType = 2
.Data5.DefaultType = 2
.Data6.DefaultType = 2
.Data7.DefaultType = 2
.Data8.DefaultType = 2
.Data9.DefaultType = 2
.Data10.DefaultType = 2
.Data11.DefaultType = 2
.Data1.DatabaseName = db
.Data1.RecordSource = "temp_trans"
.Data2.DatabaseName = db
.Data2.RecordSource = "ms_currency"
.Data3.DatabaseName = db
.Data3.RecordSource = "PERSEDIAAN"
.Data4.DatabaseName = db
.Data5.DatabaseName = db
.Data5.RecordSource = "trans_jualbeli"
.Data6.DatabaseName = db
.Data6.RecordSource = "BB_BULANAN"
.Data7.DatabaseName = db
.Data7.RecordSource = "ms_rekening"
.Data8.DatabaseName = db
.Data8.RecordSource = "rekap_jurnal"
.Data9.DatabaseName = db
.Data9.RecordSource = "Hutang"
.Data10.DatabaseName = db
.Data10.RecordSource = "piutang"
.Data11.DatabaseName = db
.Data11.RecordSource = "Stok_bulanan"
.Data12.DatabaseName = db
.Data12.RecordSource = "Stok_harian"
End With
End Sub

Sub db_stok()
buka_Db
With Stok_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data1.RecordSource = "persediaan"
End With
End Sub

Public Sub db_lbeli()
buka_Db
With LBeli_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data2.RecordSource = "ms_currency"
End With
End Sub

Public Sub Db_BB()
buka_Db
With BB_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data1.DatabaseName = db
.Data1.RecordSource = "select * from bb_bulanan order by tahun desc,bulan desc,no_akun asc"
.Data2.DatabaseName = db
.Data2.RecordSource = "tbl_akun"
End With
End Sub

Public Sub DB_Kas()
buka_Db
With Kas_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data3.DatabaseName = db
.Data2.RecordSource = "select * from tbl_akun where no_akun <> '1-110' order by no_akun"
.Data3.RecordSource = "select * from BB_BULANAN where no_akun='1-110' order by tahun desc,bulan desc"
End With
End Sub

Public Sub DB_Bank()
buka_Db
With Bank_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data3.DatabaseName = db
.Data2.RecordSource = "select * from tbl_akun where no_akun <> '1-120' order by no_akun"
.Data3.RecordSource = "select * from BB_BULANAN where no_akun='1-120' order by tahun desc,bulan desc"
End With
End Sub

Sub DB_Utang()
buka_Db
With Utang_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data4.DefaultType = 2
.Data5.DefaultType = 2
.Data6.DefaultType = 2
.Data7.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data3.DatabaseName = db
.Data4.DatabaseName = db
.Data5.DatabaseName = db
.Data6.DatabaseName = db
.Data7.DatabaseName = db
.Data5.RecordSource = "trans_hutangpiutang"
.Data6.RecordSource = "BB_BULANAN"
.Data7.RecordSource = "rekap_jurnal"
End With
End Sub

Public Sub Db_Memorial()
buka_Db
With Memorial_frm
.Data1.DefaultType = 2
.Data2.DefaultType = 2
.Data3.DefaultType = 2
.Data4.DefaultType = 2
.Data5.DefaultType = 2
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data3.DatabaseName = db
.Data4.DatabaseName = db
.Data5.DatabaseName = db
.Data2.RecordSource = "select * from tbl_akun where no_akun <>'1-110'  and no_akun<>'1-160' and no_akun<>'1-120' and no_akun<>'1-130' and no_akun<>'2-110' order by no_akun asc"
.Data3.RecordSource = "BB_BULANAN"
.Data5.RecordSource = "tEMP_JURNAL"
End With
End Sub

Public Sub dB_Admin()
buka_Db
Data_Admin.Data1.DatabaseName = db
End Sub

Public Sub dB_lAP()
buka_Db
LAporan_frm.Data1.DatabaseName = db
LAporan_frm.Data2.DatabaseName = db
LAporan_frm.Data3.DatabaseName = db
LAporan_frm.Data4.DatabaseName = db
LAporan_frm.Data5.DatabaseName = db
End Sub

Public Sub dB_Lapor()
buka_Db
Lap_frm.Data1.DatabaseName = db
Lap_frm.Data2.DatabaseName = db
Lap_frm.Data3.DatabaseName = db
Lap_frm.Data4.DatabaseName = db
Lap_frm.Data5.DatabaseName = db
End Sub

Public Sub db_ctkUlang()
buka_Db
With ctk_ulang
.Data2.DatabaseName = db
.Data1.DatabaseName = db
.Data3.DatabaseName = db
.Data4.DatabaseName = db
.Data5.DatabaseName = db
.Data6.DatabaseName = db
.Data7.DatabaseName = db
End With
End Sub

Public Sub DB_User()
buka_Db
With User_frm
.Data1.DatabaseName = db
.Data2.DatabaseName = db
.Data1.RecordSource = "select * from user where user_name<>'admin'"
.Data2.RecordSource = "ms_pegawai"
End With
End Sub

Public Sub DB_Login()
buka_Db
frmLogin.Data1.DatabaseName = db
frmLogin.Data1.RecordSource = "user"
End Sub

Public Sub Kontrol_User()
buka_Db
With MoneyChanger
.Data4.DatabaseName = db
.Data5.DatabaseName = db
.Data6.DatabaseName = db
.Data7.DatabaseName = db
End With
End Sub

Sub db_tblgaji()
buka_Db
With TblGaji_frm
    .Data1.DatabaseName = db
    .Data2.DatabaseName = db
    .Data1.RecordSource = "tbl_gaji"
    .Data2.RecordSource = "ms_pegawai"
End With
End Sub

Sub db_Gaji()
buka_Db
With Gaji_frm
    .Data1.DatabaseName = db
    .Data2.DatabaseName = db
    .Data3.DatabaseName = db
    .Data4.DatabaseName = db
    .Data5.DatabaseName = db
    .Data6.DatabaseName = db
    .Data7.DatabaseName = db
    .Data8.DatabaseName = db
    .Data9.DatabaseName = db
    .Data1.RecordSource = "tbl_gaji"
    .Data2.RecordSource = "trans_gaji"
    .Data3.RecordSource = "saldo_pinjaman"
    .Data4.RecordSource = "Trans_pinjaman"
    .Data5.RecordSource = "saldo_bonus"
    .Data6.RecordSource = "trans_bonus"
    .Data7.RecordSource = "rekap_jurnal"
    .Data8.RecordSource = "bb_bulanan"
    .Data9.RecordSource = "kas_harian"
End With
End Sub

Sub Db_Pinjaman()
buka_Db
With Pinjam_frm
    .Data1.DatabaseName = db
    .Data2.DatabaseName = db
    .Data1.RecordSource = "trans_pinjaman"
    .Data2.RecordSource = "saldo_pinjaman"
    .Data3.DatabaseName = db
    .Data3.RecordSource = "ms_pegawai"
    .Data4.DatabaseName = db
    .Data4.RecordSource = "rekap_jurnal"
    .Data5.DatabaseName = db
End With
End Sub

Sub Db_bONUS()
buka_Db
With Bonus_frm
    .Data1.DatabaseName = db
    .Data2.DatabaseName = db
    .Data1.RecordSource = "trans_bonus"
    .Data2.RecordSource = "saldo_bonus"
    .Data3.DatabaseName = db
    .Data3.RecordSource = "ms_pegawai"
    .Data4.DatabaseName = db
    .Data4.RecordSource = "rekap_jurnal"
    .Data5.DatabaseName = db
End With
End Sub


Sub Main()
frmLogin.Show
'MoneyChanger.Show
End Sub
