Attribute VB_Name = "mod_koneksi"
Public koneksi As New ADODB.Connection
Public dt_user As ADODB.Recordset
Public dt_setting As ADODB.Recordset
Public dt_buku As ADODB.Recordset
Public dt_anggota As ADODB.Recordset
Public dt_staff As ADODB.Recordset
Public dt_pinjam As ADODB.Recordset
Public dt_detilpinjam As ADODB.Recordset
Public dt_temp As ADODB.Recordset
Public dt_custom As ADODB.Recordset

Public Sub koneksi_db()
Set koneksi = New ADODB.Connection
Set dt_user = New ADODB.Recordset
Set dt_setting = New ADODB.Recordset
Set dt_buku = New ADODB.Recordset
Set dt_anggota = New ADODB.Recordset
Set dt_staff = New ADODB.Recordset
Set dt_pinjam = New ADODB.Recordset
Set dt_detilpinjam = New ADODB.Recordset
Set dt_temp = New ADODB.Recordset
Set dt_custom = New ADODB.Recordset

koneksi.ConnectionString = "driver=mysql odbc 3.51 driver;server=localhost;uid=root;db=db_perpus;"
koneksi.Open
End Sub
