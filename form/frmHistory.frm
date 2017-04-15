VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form History --"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtnopinjam 
      Height          =   375
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_h_detil 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   5530
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   12120
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_h_pinjam 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   3836
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtNo 
      Height          =   375
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data Detil Peminjaman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2385
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Peminjaman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No Anggota"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form History - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private Sub awal()
txtNo.Text = ""
txtnopinjam.Enabled = False
txtnopinjam.Text = ""
End Sub

Private Sub tampil_data()
    Call koneksi_db
    dt_pinjam.Open "select * from V_HISTORY where NO_ANGGOTA='" & txtNo.Text & "'", koneksi
    
    grid_h_pinjam.Clear
    Set grid_h_pinjam.DataSource = dt_pinjam
    
    grid_h_pinjam.ColWidth(0) = 100
    grid_h_pinjam.ColWidth(1) = 1500
    grid_h_pinjam.ColWidth(2) = 1500
    grid_h_pinjam.ColWidth(3) = 2500
    grid_h_pinjam.ColWidth(4) = 2000
    grid_h_pinjam.ColWidth(5) = 2000
    grid_h_pinjam.ColWidth(6) = 1500
    grid_h_pinjam.TextMatrix(0, 1) = "No Pinjam"
    grid_h_pinjam.TextMatrix(0, 2) = "No Anggota"
    grid_h_pinjam.TextMatrix(0, 3) = "Nama Anggota"
    grid_h_pinjam.TextMatrix(0, 4) = "Tanggal Pinjam"
    grid_h_pinjam.TextMatrix(0, 5) = "Tanggal Kembali"
    grid_h_pinjam.TextMatrix(0, 6) = "Total Denda"
    dt_pinjam.Close
End Sub

Private Sub tampil_data_recent()
    Call koneksi_db
    dt_pinjam.Open "select * from V_HISTORY where NO_PINJAM='" & txtnopinjam.Text & "'", koneksi
    
    grid_h_pinjam.Clear
    Set grid_h_pinjam.DataSource = dt_pinjam
    
    grid_h_pinjam.ColWidth(0) = 100
    grid_h_pinjam.ColWidth(1) = 1500
    grid_h_pinjam.ColWidth(2) = 1500
    grid_h_pinjam.ColWidth(3) = 2500
    grid_h_pinjam.ColWidth(4) = 2000
    grid_h_pinjam.ColWidth(5) = 2000
    grid_h_pinjam.ColWidth(6) = 1500
    grid_h_pinjam.TextMatrix(0, 1) = "No Pinjam"
    grid_h_pinjam.TextMatrix(0, 2) = "No Anggota"
    grid_h_pinjam.TextMatrix(0, 3) = "Nama Anggota"
    grid_h_pinjam.TextMatrix(0, 4) = "Tanggal Pinjam"
    grid_h_pinjam.TextMatrix(0, 5) = "Tanggal Kembali"
    grid_h_pinjam.TextMatrix(0, 6) = "Total Denda"
    dt_pinjam.Close
End Sub

Private Sub cmdKeluar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call awal
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call koneksi_db
    dt_anggota.Open "select * from V_HISTORY where NO_ANGGOTA='" & txtNo.Text & "'", koneksi
    
    With dt_anggota
        If .BOF And .EOF Then
            MsgBox "Nomor Anggota '" + txtNo.Text + "' Tidak Ditemukan.", vbInformation, "Aplikasi Perpustakaan"
            Call awal
        Else
            Call tampil_data
            txtnopinjam.Enabled = True
            txtnopinjam.SetFocus
        End If
    End With
End If
End Sub

Private Sub txtnopinjam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi_db
Call tampil_data_recent
dt_detilpinjam.Open "select KD_BUKU,NM_BUKU,PENGARANG,HARGA,STATUS,KONDISI,DENDA,QTY  from V_DETIL_PINJAM where NO_PINJAM='" & txtnopinjam.Text & "'", koneksi
grid_h_detil.Clear
Set grid_h_detil.DataSource = dt_detilpinjam

grid_h_detil.ColWidth(0) = 100
grid_h_detil.ColWidth(1) = 1200
grid_h_detil.ColWidth(2) = 3000
grid_h_detil.ColWidth(3) = 2000
grid_h_detil.ColWidth(4) = 1200
grid_h_detil.ColWidth(5) = 1200
grid_h_detil.ColWidth(6) = 1200
grid_h_detil.ColWidth(7) = 1000
grid_h_detil.ColWidth(8) = 800
grid_h_detil.TextMatrix(0, 1) = "Kode Buku"
grid_h_detil.TextMatrix(0, 2) = "Nama Buku"
grid_h_detil.TextMatrix(0, 3) = "Pengarang"
grid_h_detil.TextMatrix(0, 4) = "Harga"
grid_h_detil.TextMatrix(0, 5) = "Status"
grid_h_detil.TextMatrix(0, 6) = "Kondisi"
grid_h_detil.TextMatrix(0, 7) = "Denda"
grid_h_detil.TextMatrix(0, 8) = "QTY"

dt_detilpinjam.Close
End If
End Sub
