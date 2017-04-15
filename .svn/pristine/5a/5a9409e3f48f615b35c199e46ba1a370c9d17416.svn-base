VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLaporanPeminjaman 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Laporan Peminjaman --"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr_pinjam 
      Left            =   10920
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Bulan"
      Height          =   375
      Index           =   2
      Left            =   9360
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Minggu"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Hari ini"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   4
      Top             =   5400
      Width           =   5175
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak Semua Data"
      Height          =   495
      Left            =   11520
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   13920
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_pinjam 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
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
      Caption         =   "No Peminjaman"
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
      TabIndex        =   3
      Top             =   5400
      Width           =   1320
   End
End
Attribute VB_Name = "frmLaporanPeminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Laporan Peminjaman - Aplikasi Perpustakaan
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
cr_pinjam.ReportFileName = App.Path & "\report\Laporan_Peminjaman.rpt"
cr_pinjam.WindowState = crptNormal
cr_pinjam.RetrieveDataFiles
cr_pinjam.Action = 1
End Sub

Private Sub tampil_data_custom()
Call koneksi_db
    dt_custom.Open "select * from V_PEMINJAMAN where " + querynya, koneksi
    
    Set grid_pinjam.DataSource = dt_custom
    
    grid_pinjam.ColWidth(0) = 100
    grid_pinjam.ColWidth(1) = 1500
    grid_pinjam.ColWidth(2) = 2500
    grid_pinjam.ColWidth(3) = 1500
    grid_pinjam.ColWidth(4) = 2500
    grid_pinjam.ColWidth(5) = 1500
    grid_pinjam.ColWidth(6) = 1500
    grid_pinjam.ColWidth(7) = 1500
    grid_pinjam.ColWidth(8) = 1500
    grid_pinjam.TextMatrix(0, 1) = "No Pinjam"
    grid_pinjam.TextMatrix(0, 2) = "Nama Staff"
    grid_pinjam.TextMatrix(0, 3) = "No Anggota"
    grid_pinjam.TextMatrix(0, 4) = "Nama Anggota"
    grid_pinjam.TextMatrix(0, 5) = "Tanggal Pinjam"
    grid_pinjam.TextMatrix(0, 6) = "Tanggal Kembali"
    grid_pinjam.TextMatrix(0, 7) = "Total Denda"
    grid_pinjam.TextMatrix(0, 8) = "Update By"
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
dt_pinjam.Open "select * from V_PEMINJAMAN", koneksi
grid_pinjam.Clear
Set grid_pinjam.DataSource = dt_pinjam
    grid_pinjam.ColWidth(0) = 100
    grid_pinjam.ColWidth(1) = 1500
    grid_pinjam.ColWidth(2) = 2500
    grid_pinjam.ColWidth(3) = 1500
    grid_pinjam.ColWidth(4) = 2500
    grid_pinjam.ColWidth(5) = 1500
    grid_pinjam.ColWidth(6) = 1500
    grid_pinjam.ColWidth(7) = 1500
    grid_pinjam.ColWidth(8) = 1500
    grid_pinjam.TextMatrix(0, 1) = "No Pinjam"
    grid_pinjam.TextMatrix(0, 2) = "Nama Staff"
    grid_pinjam.TextMatrix(0, 3) = "No Anggota"
    grid_pinjam.TextMatrix(0, 4) = "Nama Anggota"
    grid_pinjam.TextMatrix(0, 5) = "Tanggal Pinjam"
    grid_pinjam.TextMatrix(0, 6) = "Tanggal Kembali"
    grid_pinjam.TextMatrix(0, 7) = "Total Denda"
    grid_pinjam.TextMatrix(0, 8) = "Update By"
dt_pinjam.Close
End Sub

Private Sub optPilihan_Click(Index As Integer)
If Index = 0 Then
    querynya = " tgl_pinjam = curdate() "
    Call tampil_data_custom
    ElseIf Index = 1 Then
            querynya = " week(tgl_pinjam) = week(current_date()) "
            Call tampil_data_custom
        ElseIf Index = 2 Then
                querynya = " month(tgl_pinjam) = month(current_date()) "
                Call tampil_data_custom
End If
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
    Call koneksi_db
    dt_pinjam.Open "select * from V_PEMINJAMAN where NO_PINJAM like'%" & txtCari.Text & "%'", koneksi
    
    Set grid_pinjam.DataSource = dt_pinjam
    
    grid_pinjam.ColWidth(0) = 100
    grid_pinjam.ColWidth(1) = 1500
    grid_pinjam.ColWidth(2) = 2500
    grid_pinjam.ColWidth(3) = 1500
    grid_pinjam.ColWidth(4) = 2500
    grid_pinjam.ColWidth(5) = 1500
    grid_pinjam.ColWidth(6) = 1500
    grid_pinjam.ColWidth(7) = 1500
    grid_pinjam.ColWidth(8) = 1500
    grid_pinjam.TextMatrix(0, 1) = "No Pinjam"
    grid_pinjam.TextMatrix(0, 2) = "Nama Staff"
    grid_pinjam.TextMatrix(0, 3) = "No Anggota"
    grid_pinjam.TextMatrix(0, 4) = "Nama Anggota"
    grid_pinjam.TextMatrix(0, 5) = "Tanggal Pinjam"
    grid_pinjam.TextMatrix(0, 6) = "Tanggal Kembali"
    grid_pinjam.TextMatrix(0, 7) = "Total Denda"
    grid_pinjam.TextMatrix(0, 8) = "Update By"
    
End Sub
