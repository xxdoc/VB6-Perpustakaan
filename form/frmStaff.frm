VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStaff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Data Staff --"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox comLocked 
      Height          =   315
      Left            =   6120
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtKode 
      Height          =   375
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   1320
      MaxLength       =   35
      TabIndex        =   11
      Top             =   720
      Width           =   6255
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   9375
      Begin VB.CommandButton cmdBersih 
         Caption         =   "Bersih"
         Height          =   495
         Left            =   4080
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   8040
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   6720
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
         Height          =   495
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCari 
      Height          =   375
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   2
      Top             =   6720
      Width           =   6135
   End
   Begin VB.TextBox txtAlamat 
      Height          =   735
      Left            =   1320
      MaxLength       =   64
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   6255
   End
   Begin VB.ComboBox comStatus 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridStaff 
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Locked"
      Height          =   195
      Left            =   4920
      TabIndex        =   22
      Top             =   2520
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kode Staff"
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
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Staff"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Nama Staff"
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
      TabIndex        =   16
      Top             =   6720
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Alamat"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   450
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Staff - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private status_form As String

Private Sub aktif()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Enabled = True
Next
comStatus.Enabled = True
comLocked.Enabled = True
End Sub

Private Sub nonaktif()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Enabled = False
Next
comStatus.Enabled = False
comLocked.Enabled = False
End Sub

Private Sub bersih()
Dim objek As Control
For Each objek In Me.Controls
If TypeOf objek Is TextBox Then objek.Text = ""
Next
    comStatus.Text = ""
    comLocked.Text = ""
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
dt_staff.Open "select * from T_STAFF order by NM_STAFF asc", koneksi
GridStaff.Clear
Set GridStaff.DataSource = dt_staff
GridStaff.ColWidth(0) = 100
GridStaff.ColWidth(1) = 900
GridStaff.ColWidth(2) = 2000
GridStaff.ColWidth(3) = 1200
GridStaff.ColWidth(4) = 3000
GridStaff.ColWidth(5) = 1200
GridStaff.ColWidth(6) = 1000
GridStaff.ColWidth(7) = 1200
GridStaff.TextMatrix(0, 1) = "Kode Staff"
GridStaff.TextMatrix(0, 2) = "Nama Staff"
GridStaff.TextMatrix(0, 3) = "Password"
GridStaff.TextMatrix(0, 4) = "Alamat"
GridStaff.TextMatrix(0, 5) = "Status"
GridStaff.TextMatrix(0, 6) = "Locked"
GridStaff.TextMatrix(0, 7) = "Login Terakhir"
End Sub

Private Sub simpan_data()
simpan = "insert into T_STAFF (KD_STAFF,PASS,NM_STAFF,ALAMAT,STATUS,LOCKED) values('" & txtKode.Text & "','" & txtPass.Text & "','" & txtNama.Text & _
         "','" & txtAlamat.Text & "','" & comStatus.Text & "','" & comLocked.Text & "')"
         
koneksi.Execute simpan
MsgBox "Data Staff Berhasil disimpan", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub update_data()
update = "update T_STAFF set NM_STAFF='" & txtNama.Text & "', PASS='" & txtPass.Text & _
         "', ALAMAT='" & txtAlamat.Text & "',STATUS ='" & comStatus.Text & "',LOCKED ='" & comLocked.Text & _
         "' where KD_STAFF='" & txtKode.Text & "'"
         
koneksi.Execute update
MsgBox "Data Staff Berhasil diupdate", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub hapus_data()
hapus = "delete from T_STAFF where KD_STAFF='" & txtKode.Text & "'"
         
koneksi.Execute hapus
MsgBox "Data Staff Berhasil dihapus", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub cmdBatal_Click()
Call bersih
Call nonaktif
Call awal
txtCari.Enabled = True
End Sub

Private Sub cmdBersih_Click()
Call bersih
txtKode.SetFocus
End Sub

Private Sub cmdEdit_Click()
status_form = "EDIT"
Call aktif

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
Call aktif

txtKode.SetFocus
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
    comStatus.AddItem "ADMIN"
    comStatus.AddItem "USER"
    comLocked.AddItem "TRUE"
    comLocked.AddItem "FALSE"
    txtCari.Enabled = True
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
Call koneksi_db

dt_staff.Open "select * from T_STAFF where NM_STAFF like '%" & txtCari.Text & "%'", koneksi
GridStaff.Clear
Set GridStaff.DataSource = dt_staff
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi_db
    dt_staff.Open "select * from T_STAFF where KD_STAFF='" & txtKode.Text & "'", koneksi
    With dt_staff
        If .BOF And .EOF Then
            MsgBox "Kode Staff '" + txtKode.Text + "' Tidak Ditemukan.", vbInformation, "Aplikasi Perpustakaan"
            Call bersih
            txtKode.SetFocus
        Else
            txtKode.Locked = True
            txtNama.Text = .Fields("NM_STAFF")
            txtPass.Text = .Fields("PASS")
            txtAlamat.Text = .Fields("ALAMAT")
            comStatus.Text = .Fields("STATUS")
            comLocked.Text = .Fields("LOCKED")
            cmdSimpan.Enabled = True
            cmdHapus.Enabled = True
        End If
    End With
    dt_staff.Close
End If
End Sub