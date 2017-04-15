VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAnggota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Data Anggota --"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNIM 
      Height          =   375
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox comStatus 
      Height          =   315
      Left            =   1320
      TabIndex        =   23
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtAlamat 
      Height          =   735
      Left            =   1320
      MaxLength       =   64
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   2040
      Width           =   6255
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   18
      Top             =   6840
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   9255
      Begin VB.CommandButton cmdBersih 
         Caption         =   "Bersih"
         Height          =   495
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
         Height          =   495
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   7920
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtKeterangan 
      Height          =   735
      Left            =   3960
      MaxLength       =   64
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox txtTglLahir 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtTglDaftar 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   2
      Top             =   600
      Width           =   6255
   End
   Begin VB.TextBox txtNo 
      Height          =   375
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAnggota 
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3413
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "NIM"
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
      Left            =   2880
      TabIndex        =   25
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Alamat"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label12 
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
      TabIndex        =   19
      Top             =   6840
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Keterangan"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   2880
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Lahir"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Daftar"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Anggota"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1065
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
Attribute VB_Name = "frmAnggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Anggota - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private status_form As String

Private Function cek_nim(nim As String) As String
Call koneksi_db
    dt_anggota.Open "select * from T_ANGGOTA where NIM='" & nim & "'", koneksi
    With dt_anggota
        If .BOF And .EOF Then
            cek_nim = "TIDAK ADA"
        Else
            cek_nim = "ADA"
        End If
    End With
    dt_anggota.Close
End Function

Private Sub after_simpan()
cmdTambah.Enabled = True
cmdEdit.Enabled = True
cmdHapus.Enabled = False
cmdSimpan.Enabled = False
cmdBatal.Enabled = False
cmdBersih.Enabled = False

Call tampil_data
Call bersih
Call awal
Call nonaktif
txtCari.Enabled = True
End Sub

Private Sub generate_no()
TANGGAL = Format(Date, "DD/MM/YYYY")

Call koneksi_db
dt_anggota.Open "select * from T_ANGGOTA order by NO_ANGGOTA desc", koneksi

With dt_anggota
    If .EOF Then
        txtNo.Text = Format(Date, "yymm") + "AG100"
    Else
        no = Right(dt_anggota!NO_ANGGOTA, 3) + 1
        txtNo.Text = Format(Date, "yymm") + "AG" + CStr(no)
    End If
End With
dt_anggota.Close
End Sub

Private Sub aktif()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Enabled = True
Next
comStatus.Enabled = True
End Sub

Private Sub nonaktif()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Enabled = False
Next
comStatus.Enabled = False
End Sub

Private Sub bersih()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Text = ""
Next
    comStatus.Text = ""
End Sub

Private Sub bersih_edit()
    txtNIM.Text = ""
    txtNama.Text = ""
    txtAlamat.Text = ""
    txtKeterangan.Text = ""
    txtTglLahir.Text = ""
    comStatus.Text = ""
End Sub

Private Sub awal()
    Call bersih
    cmdTambah.Enabled = True
    cmdSimpan.Enabled = False
    cmdEdit.Enabled = True
    cmdBatal.Enabled = False
    cmdHapus.Enabled = False
    cmdBersih.Enabled = False
    comStatus.Enabled = False
End Sub

Private Sub tampil_data()
Call koneksi_db
dt_anggota.Open "select * from T_ANGGOTA order by NM_ANGGOTA asc", koneksi
GridAnggota.Clear
Set GridAnggota.DataSource = dt_anggota
GridAnggota.ColWidth(0) = 100
GridAnggota.ColWidth(1) = 1200
GridAnggota.ColWidth(2) = 3000
GridAnggota.ColWidth(3) = 1500
GridAnggota.ColWidth(4) = 1500
GridAnggota.ColWidth(5) = 3000
GridAnggota.ColWidth(6) = 1200
GridAnggota.ColWidth(7) = 2000
GridAnggota.ColWidth(8) = 1200
GridAnggota.TextMatrix(0, 1) = "Nomor Anggota"
GridAnggota.TextMatrix(0, 2) = "Nama Anggota"
GridAnggota.TextMatrix(0, 3) = "Tanggal Daftar"
GridAnggota.TextMatrix(0, 4) = "Tanggal Lahir"
GridAnggota.TextMatrix(0, 5) = "Alamat"
GridAnggota.TextMatrix(0, 6) = "NIM"
GridAnggota.TextMatrix(0, 7) = "Keterangan"
GridAnggota.TextMatrix(0, 8) = "Status"
End Sub

Private Sub simpan_data()
simpan = "insert into T_ANGGOTA values('" & txtNo.Text & "','" & txtNama.Text & "','" & txtTglDaftar.Text & _
         "','" & txtTglLahir.Text & "','" & txtAlamat.Text & "','" & txtNIM.Text & "','" _
         & txtKeterangan & "','" & comStatus.Text & "')"
         
koneksi.Execute simpan
MsgBox "Data Anggota Berhasil disimpan", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub update_data()
update = "update T_ANGGOTA set NM_ANGGOTA='" & txtNama.Text & "', TGL_DAFTAR='" & txtTglDaftar.Text & _
         "', TGL_LAHIR='" & txtTglLahir.Text & "', ALAMAT='" & txtAlamat.Text & "', NIM='" & txtNIM.Text & _
         "',KETERANGAN='" & txtKeterangan.Text & "',STATUS ='" & comStatus.Text & _
         "' where NO_ANGGOTA='" & txtNo.Text & "'"
         
koneksi.Execute update
MsgBox "Data Anggota Berhasil diupdate", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub hapus_data()
hapus = "delete from T_ANGGOTA where NO_ANGGOTA='" & txtNo.Text & "'"
         
koneksi.Execute hapus
MsgBox "Data Anggota Berhasil dihapus", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub cmdBatal_Click()
Call bersih
Call nonaktif
Call awal
txtCari.Enabled = True
End Sub

Private Sub cmdBersih_Click()
Call bersih_edit
txtNIM.SetFocus
End Sub

Private Sub cmdEdit_Click()
status_form = "EDIT"
Call aktif

txtNo.Enabled = True
txtNo.SetFocus
cmdTambah.Enabled = False
cmdEdit.Enabled = False
cmdHapus.Enabled = False
cmdSimpan.Enabled = False
cmdBatal.Enabled = True
cmdBersih.Enabled = True
End Sub

Private Sub cmdHapus_Click()
Dim tanya
tanya = MsgBox("Anda yakin ingin menghapus data?", vbQuestion + vbYesNo, "Aplikasi Perpustakaan")
        If tanya = vbYes Then
            Call hapus_data
            Call tampil_data
            Call bersih
            Call awal
            Call nonaktif
            txtCari.Enabled = True
        End If
End Sub

Private Sub cmdKeluar_Click()
If cmdSimpan.Enabled = True Then
    MsgBox "Simpan dahulu data anda", vbExclamation, "Aplikasi Perpustakaan"
Else
    Unload Me
End If
End Sub

Private Sub cmdSimpan_Click()
Dim cek As String

If status_form = "BARU" Then
    cek = cek_nim(txtNIM.Text)
    If cek = "ADA" Then
        MsgBox "Data NIM sudah ada", vbExclamation, "Aplikasi Perpustakaan"
        txtNIM.SetFocus
    Else
        Call simpan_data
        Call after_simpan
    End If
ElseIf status_form = "EDIT" Then
    Call update_data
    Call after_simpan
End If
End Sub

Private Sub cmdTambah_Click()
status_form = "BARU"
Call aktif
Call generate_no
txtTglDaftar.Text = Format(Date, "YYYY-MM-DD")

txtNIM.SetFocus
cmdTambah.Enabled = False
cmdEdit.Enabled = False
cmdHapus.Enabled = False
cmdSimpan.Enabled = True
cmdBatal.Enabled = True
cmdBersih.Enabled = True
End Sub

Private Sub Form_Load()
    Call awal
    Call nonaktif
    Call tampil_data
    comStatus.AddItem "AKTIF"
    comStatus.AddItem "TIDAK AKTIF"
    txtCari.Enabled = True
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
Call koneksi_db

dt_anggota.Open "select * from T_ANGGOTA where NM_ANGGOTA like '%" & txtCari.Text & "%'", koneksi
GridAnggota.Clear
Set GridAnggota.DataSource = dt_anggota
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi_db
    dt_anggota.Open "select * from T_ANGGOTA where NO_ANGGOTA='" & txtNo.Text & "'", koneksi
    With dt_anggota
        If .BOF And .EOF Then
            MsgBox "Nomor Anggota '" + txtNo.Text + "' Tidak Ditemukan.", vbInformation, "Aplikasi Perpustakaan"
            Call bersih
            txtNo.SetFocus
        Else
            txtNo.Locked = True
            txtTglDaftar.Locked = True
            txtNama.Text = .Fields("NM_ANGGOTA")
            txtTglDaftar.Text = Format(.Fields("TGL_DAFTAR"), "YYYY-MM-DD")
            txtTglLahir.Text = Format(.Fields("TGL_LAHIR"), "YYYY-MM-DD")
            txtAlamat.Text = .Fields("ALAMAT")
            txtNIM.Text = .Fields("NIM")
            txtKeterangan.Text = .Fields("KETERANGAN")
            comStatus.Text = .Fields("STATUS")
            cmdSimpan.Enabled = True
            cmdHapus.Enabled = True
        End If
    End With
    dt_anggota.Close
End If
End Sub
