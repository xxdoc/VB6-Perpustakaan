VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLaporanBuku 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Laporan Buku --"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optPilihan 
      Caption         =   "10 Buku terfavorit"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   10
      Top             =   5880
      Width           =   2655
   End
   Begin Crystal.CrystalReport cr_buku 
      Left            =   6000
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCari 
      Caption         =   "&Cari.."
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtCariRak 
      Height          =   375
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "Out Of Stok"
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.OptionButton optPilihan 
      Caption         =   "In Stok"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   2
      Top             =   5520
      Width           =   5415
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak.."
      Height          =   495
      Left            =   12000
      TabIndex        =   1
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   13920
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_buku 
      Height          =   5175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9128
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Kode Rak"
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
      TabIndex        =   8
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nama Buku"
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
      TabIndex        =   6
      Top             =   5520
      Width           =   990
   End
End
Attribute VB_Name = "frmLaporanBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Laporan Buku - Aplikasi Perpustakaan
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
cr_buku.ReportFileName = App.Path & "\report\Laporan_Buku.rpt"
cr_buku.WindowState = crptNormal
cr_buku.RetrieveDataFiles
cr_buku.Action = 1
End Sub

Private Sub cetak_favorit()
Call koneksi_db
cr_buku.ReportFileName = App.Path & "\report\Laporan_Buku_Favorit.rpt"
cr_buku.WindowState = crptNormal
cr_buku.RetrieveDataFiles
cr_buku.Action = 1
End Sub

Private Sub tampil_data()
    Call koneksi_db
    dt_buku.Open "select * from T_BUKU where NM_BUKU like '%" & txtCari.Text & "%' And KD_RAK like '%" & txtCariRak.Text & "%'", koneksi
    
    Set grid_buku.DataSource = dt_buku
    grid_buku.ColWidth(0) = 100
    grid_buku.ColWidth(1) = 1000
    grid_buku.ColWidth(2) = 2500
    grid_buku.ColWidth(3) = 1500
    grid_buku.ColWidth(4) = 2500
    grid_buku.ColWidth(5) = 2500
    grid_buku.ColWidth(6) = 1500
    grid_buku.ColWidth(7) = 1500
    grid_buku.ColWidth(8) = 1500
    grid_buku.ColWidth(9) = 1000
    grid_buku.ColWidth(10) = 3500
    grid_buku.TextMatrix(0, 1) = "Kode Buku"
    grid_buku.TextMatrix(0, 2) = "Nama Buku"
    grid_buku.TextMatrix(0, 3) = "Jenis Buku"
    grid_buku.TextMatrix(0, 4) = "Pengarang"
    grid_buku.TextMatrix(0, 5) = "Penerbit"
    grid_buku.TextMatrix(0, 6) = "Tahun Terbit"
    grid_buku.TextMatrix(0, 7) = "Harga"
    grid_buku.TextMatrix(0, 8) = "Kode Rak"
    grid_buku.TextMatrix(0, 9) = "Stok"
    grid_buku.TextMatrix(0, 10) = "Keterangan"
End Sub

Private Sub tampil_data_favorit()
Call koneksi_db
    dt_custom.Open "select * from V_BUKU_FAVORIT", koneksi
    
    Set grid_buku.DataSource = dt_custom
    
    grid_buku.ColWidth(0) = 100
    grid_buku.ColWidth(1) = 1000
    grid_buku.ColWidth(2) = 2500
    grid_buku.ColWidth(3) = 2500
    grid_buku.ColWidth(4) = 2500
    grid_buku.ColWidth(5) = 1000
    grid_buku.ColWidth(6) = 2000
    grid_buku.TextMatrix(0, 1) = "Kode Buku"
    grid_buku.TextMatrix(0, 2) = "Nama Buku"
    grid_buku.TextMatrix(0, 3) = "Pengarang"
    grid_buku.TextMatrix(0, 4) = "Penerbit"
    grid_buku.TextMatrix(0, 5) = "Stok"
    grid_buku.TextMatrix(0, 6) = "Dipinjam sebanyak"
    dt_custom.Close
End Sub

Private Sub tampil_data_custom()
Call koneksi_db
    dt_custom.Open "select * from T_BUKU where " + querynya, koneksi
    
    Set grid_buku.DataSource = dt_custom
    
    grid_buku.ColWidth(0) = 100
    grid_buku.ColWidth(1) = 1000
    grid_buku.ColWidth(2) = 2500
    grid_buku.ColWidth(3) = 1500
    grid_buku.ColWidth(4) = 2500
    grid_buku.ColWidth(5) = 2500
    grid_buku.ColWidth(6) = 1500
    grid_buku.ColWidth(7) = 1500
    grid_buku.ColWidth(8) = 1500
    grid_buku.ColWidth(9) = 1000
    grid_buku.ColWidth(10) = 3500
    grid_buku.TextMatrix(0, 1) = "Kode Buku"
    grid_buku.TextMatrix(0, 2) = "Nama Buku"
    grid_buku.TextMatrix(0, 3) = "Jenis Buku"
    grid_buku.TextMatrix(0, 4) = "Pengarang"
    grid_buku.TextMatrix(0, 5) = "Penerbit"
    grid_buku.TextMatrix(0, 6) = "Tahun Terbit"
    grid_buku.TextMatrix(0, 7) = "Harga"
    grid_buku.TextMatrix(0, 8) = "Kode Rak"
    grid_buku.TextMatrix(0, 9) = "Stok"
    grid_buku.TextMatrix(0, 10) = "Keterangan"
    dt_custom.Close
End Sub

Private Sub cmdCari_Click()
If txtCari.Text = "" And txtCariRak.Text = "" Then
    MsgBox "Nama Buku atau Kode Rak masih kosong", vbExclamation, "Aplikasi Perpustakaan"
Else
    Call tampil_data
End If
End Sub

Private Sub cmdCetak_Click()
If optPilihan(2).Value = True Then
    Call cetak_favorit
Else
    Call cetak
End If
End Sub

Private Sub cmdKeluar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call koneksi_db
dt_buku.Open "select * from T_BUKU", koneksi
grid_buku.Clear
Set grid_buku.DataSource = dt_buku
    grid_buku.ColWidth(0) = 100
    grid_buku.ColWidth(1) = 1000
    grid_buku.ColWidth(2) = 2500
    grid_buku.ColWidth(3) = 1500
    grid_buku.ColWidth(4) = 2500
    grid_buku.ColWidth(5) = 2500
    grid_buku.ColWidth(6) = 1500
    grid_buku.ColWidth(7) = 1500
    grid_buku.ColWidth(8) = 1500
    grid_buku.ColWidth(9) = 1000
    grid_buku.ColWidth(10) = 3500
    grid_buku.TextMatrix(0, 1) = "Kode Buku"
    grid_buku.TextMatrix(0, 2) = "Nama Buku"
    grid_buku.TextMatrix(0, 3) = "Jenis Buku"
    grid_buku.TextMatrix(0, 4) = "Pengarang"
    grid_buku.TextMatrix(0, 5) = "Penerbit"
    grid_buku.TextMatrix(0, 6) = "Tahun Terbit"
    grid_buku.TextMatrix(0, 7) = "Harga"
    grid_buku.TextMatrix(0, 8) = "Kode Rak"
    grid_buku.TextMatrix(0, 9) = "Stok"
    grid_buku.TextMatrix(0, 10) = "Keterangan"
dt_buku.Close
End Sub

Private Sub optPilihan_Click(Index As Integer)
If Index = 0 Then
    querynya = " STOK > 0 "
    Call tampil_data_custom
    ElseIf Index = 1 Then
            querynya = " STOK=0 "
            Call tampil_data_custom
            ElseIf Index = 2 Then
                Call tampil_data_favorit
End If
End Sub
