VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLaporanAnggota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Laporan Anggota --"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr_anggota 
      Left            =   7800
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   13920
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak Semua Data"
      Height          =   495
      Left            =   11760
      TabIndex        =   4
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   3
      Top             =   5400
      Width           =   5175
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Aktif"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   2
      Top             =   5400
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Tidak Aktif"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Daftar Hari ini"
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   0
      Top             =   5400
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_anggota 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9128
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nama Anggota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1260
   End
End
Attribute VB_Name = "frmLaporanAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Laporan Anggota - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private querynya As String

Private Sub cetak()
Call koneksi_db
cr_anggota.ReportFileName = App.Path & "\report\Laporan_Anggota.rpt"
cr_anggota.WindowState = crptNormal
cr_anggota.RetrieveDataFiles
cr_anggota.Action = 1
End Sub

Private Sub tampil_data_custom()
Call koneksi_db
    dt_custom.Open "select * from T_ANGGOTA where " + querynya, koneksi
    
    Set grid_anggota.DataSource = dt_custom
    
    grid_anggota.ColWidth(0) = 100
    grid_anggota.ColWidth(1) = 1500
    grid_anggota.ColWidth(2) = 2500
    grid_anggota.ColWidth(3) = 1500
    grid_anggota.ColWidth(4) = 1500
    grid_anggota.ColWidth(5) = 2500
    grid_anggota.ColWidth(6) = 1500
    grid_anggota.ColWidth(7) = 2500
    grid_anggota.ColWidth(8) = 1500
    grid_anggota.TextMatrix(0, 1) = "No Anggota"
    grid_anggota.TextMatrix(0, 2) = "Nama Anggota"
    grid_anggota.TextMatrix(0, 3) = "Tanggal Daftar"
    grid_anggota.TextMatrix(0, 4) = "Tanggal Lahir"
    grid_anggota.TextMatrix(0, 5) = "Alamat"
    grid_anggota.TextMatrix(0, 6) = "NIM"
    grid_anggota.TextMatrix(0, 7) = "Keterangan"
    grid_anggota.TextMatrix(0, 8) = "Status"
    dt_custom.Close
End Sub

Private Sub cmdCetak_Click()
Call cetak
End Sub

Private Sub cmdKeluar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call koneksi_db
dt_anggota.Open "select * from T_ANGGOTA", koneksi
grid_anggota.Clear
Set grid_anggota.DataSource = dt_anggota
    grid_anggota.ColWidth(0) = 100
    grid_anggota.ColWidth(1) = 1500
    grid_anggota.ColWidth(2) = 2500
    grid_anggota.ColWidth(3) = 1500
    grid_anggota.ColWidth(4) = 1500
    grid_anggota.ColWidth(5) = 2500
    grid_anggota.ColWidth(6) = 1500
    grid_anggota.ColWidth(7) = 2500
    grid_anggota.ColWidth(8) = 1500
    grid_anggota.TextMatrix(0, 1) = "No Anggota"
    grid_anggota.TextMatrix(0, 2) = "Nama Anggota"
    grid_anggota.TextMatrix(0, 3) = "Tanggal Daftar"
    grid_anggota.TextMatrix(0, 4) = "Tanggal Lahir"
    grid_anggota.TextMatrix(0, 5) = "Alamat"
    grid_anggota.TextMatrix(0, 6) = "NIM"
    grid_anggota.TextMatrix(0, 7) = "Keterangan"
    grid_anggota.TextMatrix(0, 8) = "Status"
dt_anggota.Close
End Sub

Private Sub optPilihan_Click(Index As Integer)
If Index = 0 Then
    querynya = " status='AKTIF' "
    Call tampil_data_custom
    ElseIf Index = 1 Then
            querynya = " status='TIDAK AKTIF' "
            Call tampil_data_custom
        ElseIf Index = 2 Then
                querynya = " tgl_daftar = curdate()  "
                Call tampil_data_custom
End If
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Call koneksi_db
    dt_anggota.Open "select * from T_ANGGOTA where NM_ANGGOTA like'%" & txtCari.Text & "%'", koneksi
    
    Set grid_anggota.DataSource = dt_anggota
    grid_anggota.ColWidth(0) = 100
    grid_anggota.ColWidth(1) = 1500
    grid_anggota.ColWidth(2) = 2500
    grid_anggota.ColWidth(3) = 1500
    grid_anggota.ColWidth(4) = 1500
    grid_anggota.ColWidth(5) = 2500
    grid_anggota.ColWidth(6) = 1500
    grid_anggota.ColWidth(7) = 2500
    grid_anggota.ColWidth(8) = 1500
    grid_anggota.TextMatrix(0, 1) = "No Anggota"
    grid_anggota.TextMatrix(0, 2) = "Nama Anggota"
    grid_anggota.TextMatrix(0, 3) = "Tanggal Daftar"
    grid_anggota.TextMatrix(0, 4) = "Tanggal Lahir"
    grid_anggota.TextMatrix(0, 5) = "Alamat"
    grid_anggota.TextMatrix(0, 6) = "NIM"
    grid_anggota.TextMatrix(0, 7) = "Keterangan"
    grid_anggota.TextMatrix(0, 8) = "Status"
End Sub
