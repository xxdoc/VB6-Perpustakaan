VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuku 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Data Buku --"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   28
      Top             =   6840
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   9375
      Begin VB.CommandButton cmdBersih 
         Caption         =   "Bersih"
         Height          =   495
         Left            =   4080
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   8040
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   6720
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
         Height          =   495
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridBuku 
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3201
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtKeterangan 
      Height          =   735
      Left            =   4800
      MaxLength       =   64
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox txtStok 
      Height          =   375
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtKodeRak 
      Height          =   375
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   14
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtHarga 
      Height          =   375
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtTahun 
      Height          =   375
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtPenerbit 
      Height          =   375
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2160
      Width           =   5295
   End
   Begin VB.ComboBox comJenis 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtPengarang 
      Height          =   375
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1680
      Width           =   5295
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   1200
      MaxLength       =   35
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.TextBox txtKode 
      Height          =   375
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblterbilang 
      AutoSize        =   -1  'True
      Caption         =   "Nol Rupiah"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3720
      TabIndex        =   30
      Top             =   2640
      Width           =   795
   End
   Begin VB.Label Label12 
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
      TabIndex        =   29
      Top             =   6840
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Keterangan"
      Height          =   195
      Left            =   3720
      TabIndex        =   19
      Top             =   3120
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Stok"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   330
   End
   Begin VB.Label Label9 
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
      Left            =   3720
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Harga"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tahun Terbit"
      Height          =   195
      Left            =   3720
      TabIndex        =   11
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Penerbit"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Pengarang"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Buku"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Buku"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kode Buku"
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
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmBuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Buku - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private status_form, kodejenis As String

Private Sub generate_no()
Dim no As Double
Call koneksi_db
dt_buku.Open "select RIGHT(KD_BUKU,4) AS KODE FROM T_bUKU ORDER BY KODE DESC", koneksi

With dt_buku
    If .EOF Then
        txtKode.Text = "1000"
    Else
        no = Val(dt_buku!KODE) + 1
        txtKode.Text = kodejenis + CStr(no)
    End If
End With
dt_buku.Close
End Sub

Private Sub aktif()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Enabled = True
Next
comJenis.Enabled = True
End Sub

Private Sub nonaktif()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Enabled = False
Next
comJenis.Enabled = False
txtCari.Enabled = True
End Sub

Private Sub bersih()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Text = ""
Next
    comJenis.Text = ""
End Sub

Private Sub awal()
    Call bersih
    cmdTambah.Enabled = True
    cmdSimpan.Enabled = False
    cmdEdit.Enabled = True
    cmdBatal.Enabled = False
    cmdHapus.Enabled = False
    cmdBersih.Enabled = False
    comJenis.Enabled = False
End Sub

Private Sub tampil_data()
Call koneksi_db
dt_buku.Open "select * from T_BUKU order by KD_BUKU asc", koneksi
GridBuku.Clear
Set GridBuku.DataSource = dt_buku
GridBuku.ColWidth(0) = 100
GridBuku.ColWidth(1) = 1000
GridBuku.ColWidth(2) = 3000
GridBuku.ColWidth(3) = 900
GridBuku.ColWidth(4) = 1200
GridBuku.ColWidth(5) = 1200
GridBuku.ColWidth(6) = 700
GridBuku.ColWidth(7) = 700
GridBuku.ColWidth(8) = 900
GridBuku.ColWidth(9) = 700
GridBuku.ColWidth(10) = 2000
GridBuku.TextMatrix(0, 1) = "Kode Buku"
GridBuku.TextMatrix(0, 2) = "Nama Buku"
GridBuku.TextMatrix(0, 3) = "Jenis Buku"
GridBuku.TextMatrix(0, 4) = "Pengarang"
GridBuku.TextMatrix(0, 5) = "Penerbit"
GridBuku.TextMatrix(0, 6) = "Tahun"
GridBuku.TextMatrix(0, 7) = "Harga"
GridBuku.TextMatrix(0, 8) = "Kode Rak"
GridBuku.TextMatrix(0, 9) = "Stok"
GridBuku.TextMatrix(0, 10) = "Keterangan"
End Sub

Private Sub simpan_data()
simpan = "insert into T_BUKU values('" & txtKode.Text & "','" & txtNama.Text & "','" & comJenis.Text & _
         "','" & txtPengarang.Text & "','" & txtPenerbit.Text & "','" & txtTahun.Text & "','" _
         & txtHarga.Text & "','" & txtKodeRak.Text & "','" & txtStok.Text & "','" & txtKeterangan.Text & "')"
         
koneksi.Execute simpan
MsgBox "Data Buku Berhasil disimpan", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub update_data()
update = "update T_BUKU set NM_BUKU='" & txtNama.Text & "', JNS_BUKU='" & comJenis.Text & _
         "', PENGARANG='" & txtPengarang.Text & "', PENERBIT='" & txtPenerbit.Text & "', THN_TERBIT='" & txtTahun.Text & _
         "',HARGA=" & txtHarga.Text & ", KD_RAK='" & txtKodeRak.Text & "', STOK=" & txtStok.Text & _
         ",KETERANGAN ='" & txtKeterangan.Text & _
         "' where KD_BUKU='" & txtKode.Text & "'"
         
koneksi.Execute update
MsgBox "Data Buku Berhasil diupdate", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub hapus_data()
hapus = "delete from T_BUKU where KD_BUKU='" & txtKode.Text & "'"
         
koneksi.Execute hapus
MsgBox "Data Buku Berhasil dihapus", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub cmdBatal_Click()
Call bersih
Call nonaktif
Call awal
txtCari.Enabled = True
End Sub

Private Sub cmdBersih_Click()
If status_form = "BARU" Then
    Call bersih
    generate_no
    txtNama.SetFocus
ElseIf status_form = "EDIT" Then
    Call bersih
    txtKode.SetFocus
End If
End Sub

Private Sub cmdEdit_Click()
status_form = "EDIT"
Call aktif

txtKode.Enabled = True
txtKode.SetFocus
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
If status_form = "BARU" Then
    Call simpan_data
ElseIf status_form = "EDIT" Then
    Call update_data
End If

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

Private Sub cmdTambah_Click()
status_form = "BARU"
kodejenis = "KM"
Call aktif
Call generate_no

comJenis.Text = "Komputer"
txtKode.Enabled = False
txtNama.SetFocus
cmdTambah.Enabled = False
cmdEdit.Enabled = False
cmdHapus.Enabled = False
cmdSimpan.Enabled = True
cmdBatal.Enabled = True
cmdBersih.Enabled = True
End Sub

Private Sub comJenis_Click()
If comJenis.Text = "Komputer" Then
    kodejenis = "KM"
    generate_no
ElseIf comJenis.Text = "Umum" Then
    kodejenis = "UM"
    generate_no
ElseIf comJenis.Text = "Novel" Then
    kodejenis = "NV"
    generate_no
ElseIf comJenis.Text = "Komik" Then
    kodejenis = "KK"
    generate_no
End If
End Sub

Private Sub Form_Load()
    Call awal
    Call nonaktif
    Call tampil_data
    comJenis.AddItem "Komputer"
    comJenis.AddItem "Umum"
    comJenis.AddItem "Novel"
    comJenis.AddItem "Komik"
    txtCari.Enabled = True
    kodejenis = "KM"
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
Call koneksi_db

dt_buku.Open "select * from T_BUKU where NM_BUKU like '%" & txtCari.Text & "%'", koneksi
GridBuku.Clear
Set GridBuku.DataSource = dt_buku
End Sub


Private Sub txtHarga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lblterbilang.Caption = Terbilang(Val(txtHarga.Text))
    txtStok.SetFocus
End If
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi_db
    dt_buku.Open "select * from T_BUKU where KD_BUKU='" & txtKode.Text & "'", koneksi
    With dt_buku
        If .BOF And .EOF Then
            MsgBox "Kode Buku '" + txtKode.Text + "' Tidak Ditemukan.", vbInformation, "Aplikasi Perpustakaan"
            Call bersih
            txtKode.SetFocus
        Else
            txtKode.Locked = True
            txtNama.Text = .Fields("NM_BUKU")
            comJenis.Text = .Fields("JNS_BUKU")
            txtPengarang.Text = .Fields("PENGARANG")
            txtPenerbit.Text = .Fields("PENERBIT")
            txtTahun.Text = .Fields("THN_TERBIT")
            txtHarga.Text = .Fields("HARGA")
            txtStok.Text = .Fields("STOK")
            txtKodeRak.Text = .Fields("KD_RAK")
            txtKeterangan.Text = .Fields("KETERANGAN")
            cmdSimpan.Enabled = True
            cmdHapus.Enabled = True
        End If
    End With
    dt_buku.Close
End If
End Sub
