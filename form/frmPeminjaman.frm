VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPeminjaman 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Peminjaman Buku --"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr_pinjam 
      Left            =   4680
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   2880
      TabIndex        =   28
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Bata&l"
      Height          =   495
      Left            =   8400
      TabIndex        =   27
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   9720
      TabIndex        =   26
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simp&an"
      Height          =   495
      Left            =   1560
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "&Input"
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   6360
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_pinjam 
      Height          =   3375
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5953
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtQTY 
      Height          =   375
      Left            =   10200
      MaxLength       =   2
      TabIndex        =   21
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtStok 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtTahunTerbit 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtPengarang 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtNamaBuku 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txtKodeBuku 
      Height          =   375
      Left            =   240
      MaxLength       =   6
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtNamaPeminjam 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   9
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtNoAnggota 
      Height          =   375
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.TextBox txtTgl 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtStaff 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtNoPinjam 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal"
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
         Left            =   8520
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Staff"
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
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Width           =   420
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
         TabIndex        =   2
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Label lbljumbuku 
      Caption         =   "N/A"
      Height          =   255
      Left            =   9480
      TabIndex        =   32
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Buku"
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
      Left            =   7080
      TabIndex        =   31
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lbltglkembali 
      Caption         =   "N/A"
      Height          =   255
      Left            =   1920
      TabIndex        =   30
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Kembali"
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
      Left            =   240
      TabIndex        =   29
      Top             =   6000
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      Height          =   195
      Left            =   10440
      TabIndex        =   22
      Top             =   1800
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Stok"
      Height          =   195
      Left            =   9000
      TabIndex        =   20
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Tahun Terbit"
      Height          =   195
      Left            =   7680
      TabIndex        =   18
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Pengarang"
      Height          =   195
      Left            =   5280
      TabIndex        =   16
      Top             =   1800
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nama Buku"
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Kode Buku"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      X1              =   240
      X2              =   11160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nama Peminjam"
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "No Anggota"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmPeminjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Peminjaman - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private jum_buku, max_buku As Integer

Private Sub cetak()
Call koneksi_db
cr_pinjam.SelectionFormula = "{v_peminjaman.NO_PINJAM}='" & txtNoPinjam.Text & "'"
cr_pinjam.ReportFileName = App.Path & "\report\Peminjaman.rpt"
cr_pinjam.WindowState = crptNormal
cr_pinjam.RetrieveDataFiles
cr_pinjam.Action = 1
End Sub

Private Sub generate_no()
TANGGAL = Format(Date, "DD/MM/YYYY")

Call koneksi_db
dt_pinjam.Open "select * from T_PEMINJAMAN order by NO_PINJAM desc", koneksi

With dt_pinjam
    If .EOF Then
        txtNoPinjam.Text = Format(Date, "yymm") + "PJ1000"
    Else
        no = Right(dt_pinjam!NO_PINJAM, 4) + 1
        txtNoPinjam.Text = Format(Date, "yymm") + "PJ" + CStr(no)
    End If
End With
End Sub

Private Sub aktif()
    txtNoAnggota.Enabled = True
    txtKodeBuku.Enabled = True
    txtQTY.Enabled = True
    cmdBatal.Enabled = True
    
    cmdCetak.Enabled = False
    cmdInput.Enabled = False
End Sub

Private Sub nonaktif()
    txtNoAnggota.Enabled = False
    txtKodeBuku.Enabled = False
    txtQTY.Enabled = False
    cmdSimpan.Enabled = False
    cmdBatal.Enabled = False
    cmdCetak.Enabled = False
    
    cmdInput.Enabled = True
End Sub

Private Sub bersih()
    txtNoAnggota.Text = ""
    txtNamaPeminjam.Text = ""
    txtKodeBuku.Text = ""
    txtNamaBuku.Text = ""
    txtPengarang.Text = ""
    txtTahunTerbit.Text = ""
    txtStok.Text = ""
    txtQTY.Text = ""
    txtNoAnggota.SetFocus
    grid_pinjam.Clear
End Sub

Private Sub batal()
    txtNoAnggota.Text = ""
    txtNamaPeminjam.Text = ""
    txtKodeBuku.Text = ""
    txtNamaBuku.Text = ""
    txtPengarang.Text = ""
    txtTahunTerbit.Text = ""
    txtStok.Text = ""
    txtQTY.Text = ""
    grid_pinjam.Clear
End Sub

Private Sub input_grid()
Call koneksi_db
dt_temp.Open " select * from T_TEMP", koneksi
Set grid_pinjam.DataSource = dt_temp
grid_pinjam.ColWidth(0) = 100
grid_pinjam.ColWidth(1) = 1200
grid_pinjam.ColWidth(2) = 3000
grid_pinjam.ColWidth(3) = 2000
grid_pinjam.ColWidth(4) = 1200
grid_pinjam.ColWidth(5) = 1200
grid_pinjam.TextMatrix(0, 1) = "Kode Buku"
grid_pinjam.TextMatrix(0, 2) = "Nama Buku"
grid_pinjam.TextMatrix(0, 3) = "Pengarang"
grid_pinjam.TextMatrix(0, 4) = "Tahun Terbit"
grid_pinjam.TextMatrix(0, 5) = "Jumlah Pinjam"

dt_temp.Close
End Sub

Private Sub buat_table_temp()
Dim buat As String

buat = "create table T_TEMP(KD_BUKU varchar(6), NM_BUKU varchar(35), PENGARANG varchar(30), THN varchar(4), QTY int)"
koneksi.Execute buat
End Sub

Private Sub hapus_table_temp()
hapus = "drop table if exists T_TEMP"
koneksi.Execute hapus
End Sub

Private Sub bersih_table_temp()
Dim bersih As String

bersih = "truncate table T_TEMP"
koneksi.Execute bersih
End Sub

Private Sub simpan_table_temp()
simpan = "insert into T_TEMP() values('" & txtKodeBuku.Text & "','" & txtNamaBuku.Text & "','" & txtPengarang.Text & _
         "','" & txtTahunTerbit.Text & "'," & txtQTY.Text & ")"
koneksi.Execute (simpan)
End Sub

Private Sub kurangin_stok()
Dim stok As Double

stok = Val(txtStok.Text) - Val(txtQTY.Text)
update = "update T_BUKU set STOK=" & stok & " where KD_BUKU='" & txtKodeBuku.Text & "'"
koneksi.Execute (update)
End Sub

Private Sub simpan_peminjaman()
TANGGAL = Format(Date, "YYYY-MM-DD")
simpan = "insert into T_PEMINJAMAN (NO_PINJAM, KD_STAFF, NO_ANGGOTA, TGL_PINJAM) " & _
         " values('" & txtNoPinjam.Text & "','" & kode_staff & "','" & txtNoAnggota.Text & "','" & txtTgl.Text & "')"
koneksi.Execute (simpan)
End Sub

Private Sub simpan_detil_peminjaman()
Dim simpan, nopinjam, kdbuku As String

For a = 1 To (grid_pinjam.Rows - 1)
kdbuku = grid_pinjam.TextMatrix(a, 1)
simpan = "insert into T_DETIL_PINJAM (NO_PINJAM,KD_BUKU,QTY,STATUS,KONDISI,DENDA) values( " & _
        "'" & txtNoPinjam.Text & "','" & kdbuku & "'," & txtQTY.Text & ",'-','-',0)"

Set dt_detilpinjam = koneksi.Execute(simpan)
Next a
End Sub

Private Sub cmdBatal_Click()
Call nonaktif
Call batal
Call bersih_table_temp
txtNoPinjam.Text = ""
lbltglkembali.Caption = ""
lbljumbuku.Caption = ""
End Sub

Private Sub cmdCetak_Click()
Call cetak
End Sub

Private Sub cmdInput_Click()
Call generate_no
Call aktif
Call hapus_table_temp
Call buat_table_temp

Call koneksi_db
Dim max_pinjam As Integer

dt_setting.Open "select NILAI from T_SETTINGS where NAMA_SETTING ='LAMA_PINJAM'", koneksi
max_pinjam = dt_setting!NILAI
dt_setting.Close
    
Dim tgl_temp
tgl_temp = DateAdd("d", max_pinjam, Date)

lbltglkembali.Caption = Format(tgl_temp, "YYYY-MM-DD")
txtNoAnggota.SetFocus
End Sub

Private Sub cmdKeluar_Click()
If cmdSimpan.Enabled = True Then
    MsgBox "Simpan dahulu data anda", vbExclamation, "Aplikasi Perpustakaan"
Else
    Unload Me
End If
End Sub

Private Sub cmdSimpan_Click()
Call simpan_peminjaman
Call simpan_detil_peminjaman
Call nonaktif
cmdCetak.Enabled = True
MsgBox "Data Peminjaman Berhasil disimpan", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub Form_Load()
    txtTgl.Text = Format(Date, "YYYY-MM-DD")
    txtStaff.Text = nama_staff
    Call nonaktif
    jum_buku = 0
    
    Call koneksi_db
    dt_setting.Open "select NILAI from T_SETTINGS where NAMA_SETTING ='MAX_BUKU'", koneksi
    max_buku = dt_setting!NILAI
    dt_setting.Close
End Sub

Private Sub txtKodeBuku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call koneksi_db
    
    dt_buku.Open "select * from T_BUKU where KD_BUKU='" & txtKodeBuku.Text & "'", koneksi
    With dt_buku
        If .BOF And .EOF Then
            MsgBox "Kode Buku '" + txtKodeBuku.Text + "' Tidak Ditemukan.", vbInformation, "Aplikasi Perpustakaan"
            txtKodeBuku.Text = ""
            txtNamaBuku.Text = ""
            txtPengarang.Text = ""
            txtTahunTerbit.Text = ""
            txtStok.Text = ""
            txtQTY.Text = ""
            txtKodeBuku.SetFocus
        Else
            txtNamaBuku.Text = .Fields("NM_BUKU")
            txtPengarang.Text = .Fields("PENGARANG")
            txtTahunTerbit.Text = .Fields("THN_TERBIT")
            txtStok.Text = .Fields("STOK")
            txtQTY.SetFocus
        End If
    End With
    dt_buku.Close
End If
End Sub

Private Sub txtNoAnggota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call koneksi_db
    
    dt_anggota.Open "select * from T_ANGGOTA where NO_ANGGOTA='" & txtNoAnggota.Text & "'", koneksi
    With dt_anggota
        If .BOF And .EOF Then
            MsgBox "Kode Anggota '" + txtNoAnggota.Text + "' Tidak Ditemukan.", vbInformation, "Aplikasi Perpustakaan"
            txtNoAnggota.Text = ""
            txtNamaPeminjam.Text = ""
            txtNoAnggota.SetFocus
        Else
            txtNamaPeminjam.Text = .Fields("NM_ANGGOTA")
            txtKodeBuku.SetFocus
        End If
    End With
End If
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
Dim jum_awal As Integer
If KeyAscii = 13 Then
jum_awal = Val(lbljumbuku.Caption)
    If Val(txtStok.Text) < 1 Then
        MsgBox "Stok Buku ini tidak ada di sistem", vbExclamation, "Aplikasi Perpustakaan"
        txtKodeBuku.Text = ""
        txtNamaBuku.Text = ""
        txtPengarang.Text = ""
        txtTahunTerbit.Text = ""
        txtStok.Text = ""
        txtQTY.Text = ""
        txtKodeBuku.SetFocus
    End If
    
    If Val(txtQTY) > Val(txtStok.Text) Then
            MsgBox "Stok Buku ini kurang", vbExclamation, "Aplikasi Perpustakaan"
            txtQTY.Text = ""
    Else
        jum_buku = jum_buku + Val(txtQTY.Text)
        If jum_buku > max_buku Then
            MsgBox "Jumlah buku sudah mencapai batas maksimal", vbExclamation, "Aplikasi Perpustakaan"
            lbljumbuku.Caption = jum_awal
            cmdSimpan.Enabled = True
        Else
            Dim tanya
            Call simpan_table_temp
            Call input_grid
            Call kurangin_stok
            tanya = MsgBox("Anda ingin input buku lagi?", vbQuestion + vbYesNo, "Aplikasi Perpustakaan")
                If tanya = vbYes Then
                    txtKodeBuku.Text = ""
                    txtNamaBuku.Text = ""
                    txtPengarang.Text = ""
                    txtTahunTerbit.Text = ""
                    txtStok.Text = ""
                    txtQTY.Text = ""
                    txtKodeBuku.SetFocus
                    lbljumbuku.Caption = jum_buku
                Else
                    Call nonaktif
                    cmdSimpan.Enabled = True
                    cmdBatal.Enabled = True
                    lbljumbuku.Caption = jum_buku
                End If
        End If
    End If
End If
End Sub
