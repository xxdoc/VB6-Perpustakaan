VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPengembalian 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Pengembalian Buku --"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr_pinjam 
      Left            =   4440
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtNamaBuku 
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox txtDenda 
      Height          =   375
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox comKondisi 
      Height          =   315
      Left            =   7200
      TabIndex        =   30
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox comStatus 
      Height          =   315
      Left            =   5280
      TabIndex        =   28
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   12375
      Begin VB.TextBox txtTglKembali 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtPeminjam 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   23
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtNoAnggota 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   22
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtNoPinjam 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtStaff 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtTglPinjam 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label14 
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
         Height          =   435
         Left            =   7560
         TabIndex        =   27
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Peminjam"
         Height          =   195
         Left            =   3720
         TabIndex        =   25
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No Anggota"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   855
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
         TabIndex        =   14
         Top             =   240
         Width           =   1320
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
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label10 
         Caption         =   "Tanggal  Pinjam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7560
         TabIndex        =   12
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.TextBox txtKodeBuku 
      Height          =   375
      Left            =   240
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtQTY 
      Height          =   375
      Left            =   11760
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "&Input"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simp&an"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   11280
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Bata&l"
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid_pinjam 
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   5953
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbltotdenda 
      Caption         =   "N/A"
      Height          =   255
      Left            =   1680
      TabIndex        =   38
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total Denda"
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
      TabIndex        =   37
      Top             =   5760
      Width           =   1065
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Nama Buku"
      Height          =   195
      Left            =   1680
      TabIndex        =   36
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "%"
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
      Left            =   11520
      TabIndex        =   34
      Top             =   1800
      Width           =   150
   End
   Begin VB.Label lblpersendenda 
      AutoSize        =   -1  'True
      Caption         =   "N/A"
      Height          =   195
      Left            =   11160
      TabIndex        =   33
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Denda"
      Height          =   195
      Left            =   9120
      TabIndex        =   32
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Kondisi"
      Height          =   195
      Left            =   7200
      TabIndex        =   29
      Top             =   1560
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Kode Buku"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   5280
      TabIndex        =   20
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      Height          =   195
      Left            =   11760
      TabIndex        =   19
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Buku yang dipinjam "
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
      Left            =   8160
      TabIndex        =   18
      Top             =   5760
      Width           =   2370
   End
   Begin VB.Label lblbukupinjam 
      Caption         =   "N/A"
      Height          =   255
      Left            =   11040
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Buku yang dikembalikan"
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
      Left            =   8160
      TabIndex        =   16
      Top             =   6120
      Width           =   2730
   End
   Begin VB.Label lblbukukembali 
      Caption         =   "N/A"
      Height          =   255
      Left            =   11040
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPengembalian"
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

Private harga_buku, stoknya, denda_harian, max_pinjam As Double

Private Sub cetak()
Call koneksi_db
cr_pinjam.SelectionFormula = "{v_peminjaman.NO_PINJAM}='" & txtNoPinjam.Text & "'"
cr_pinjam.ReportFileName = App.Path & "\report\Pengembalian.rpt"
cr_pinjam.WindowState = crptNormal
cr_pinjam.RetrieveDataFiles
cr_pinjam.Action = 1
End Sub

Private Sub aktif()
cmdInput.Enabled = False
cmdSimpan.Enabled = True
cmdCetak.Enabled = False
cmdBatal.Enabled = True

txtNoPinjam.Enabled = True
txtKodeBuku.Enabled = True
txtQTY.Enabled = True
End Sub

Private Sub nonaktif()
cmdInput.Enabled = True
cmdSimpan.Enabled = False
cmdCetak.Enabled = False
cmdBatal.Enabled = False

txtNoPinjam.Enabled = False
txtKodeBuku.Enabled = False
txtQTY.Enabled = False
End Sub

Private Sub jum_buku()
Call koneksi_db

dt_pinjam.Open "select sum(QTY) as JUMLAH from V_DETIL_PINJAM where NO_PINJAM like '%" & txtNoPinjam.Text & "%'", koneksi
lblbukupinjam.Caption = dt_pinjam!JUMLAH
dt_pinjam.Close
End Sub

Private Sub jum_denda()
Call koneksi_db

dt_setting.Open "select NILAI from T_SETTINGS where NAMA_SETTING='PERSEN_DENDA'", koneksi
lblpersendenda.Caption = dt_setting!NILAI
dt_setting.Close

dt_setting.Open "select NILAI from T_SETTINGS where NAMA_SETTING='DENDA_PERHARI'", koneksi
denda_harian = dt_setting!NILAI
dt_setting.Close

dt_setting.Open "select NILAI from T_SETTINGS where NAMA_SETTING='LAMA_PINJAM'", koneksi
max_pinjam = dt_setting!NILAI
dt_setting.Close
End Sub

Private Sub update_detilpinjam()
update = "update T_DETIL_PINJAM set KONDISI='" & comKondisi.Text & _
         "', STATUS='" & comStatus.Text & "', DENDA=" & txtDenda.Text & ", QTY=" & txtQTY.Text & _
         " where NO_PINJAM='" & txtNoPinjam.Text & "' and KD_BUKU='" & txtKodeBuku.Text & "'"
koneksi.Execute (update)
End Sub

Private Sub tampil_detilpinjam()
Call koneksi_db
dt_detilpinjam.Open "select KD_BUKU,NM_BUKU,PENGARANG,HARGA,STATUS,KONDISI,DENDA,QTY from V_DETIL_PINJAM where NO_PINJAM='" & txtNoPinjam.Text & "'", koneksi
grid_pinjam.Clear
Set grid_pinjam.DataSource = dt_detilpinjam
grid_pinjam.ColWidth(0) = 100
grid_pinjam.ColWidth(1) = 1200
grid_pinjam.ColWidth(2) = 3000
grid_pinjam.ColWidth(3) = 2000
grid_pinjam.ColWidth(4) = 1200
grid_pinjam.ColWidth(5) = 1200
grid_pinjam.ColWidth(6) = 1200
grid_pinjam.ColWidth(7) = 1000
grid_pinjam.ColWidth(8) = 800
grid_pinjam.TextMatrix(0, 1) = "Kode Buku"
grid_pinjam.TextMatrix(0, 2) = "Nama Buku"
grid_pinjam.TextMatrix(0, 3) = "Pengarang"
grid_pinjam.TextMatrix(0, 4) = "Harga"
grid_pinjam.TextMatrix(0, 5) = "Status"
grid_pinjam.TextMatrix(0, 6) = "Kondisi"
grid_pinjam.TextMatrix(0, 7) = "Denda"
grid_pinjam.TextMatrix(0, 8) = "QTY"

dt_detilpinjam.Close
End Sub

Private Sub update_peminjaman()
update = "update T_PEMINJAMAN set TGL_KEMBALI='" & txtTglKembali.Text & _
         "', TOTAL_DENDA=" & Val(lbltotdenda.Caption) & ", UPDATE_BY='" & kode_staff & _
         "' where NO_PINJAM='" & txtNoPinjam.Text & "'"
koneksi.Execute (update)

MsgBox "Data Peminjaman " + txtNoPinjam.Text & ", berhasil di Update", vbInformation, "Aplikasi Perpustakaan"
cmdCetak.Enabled = True
cmdSimpan.Enabled = False
End Sub

Private Sub tambahin_stok()
Dim stok As Double

stok = stoknya + Val(txtQTY.Text)
update = "update T_BUKU set STOK=" & stok & " where KD_BUKU='" & txtKodeBuku.Text & "'"
koneksi.Execute (update)
End Sub

Private Sub awal()
txtNoPinjam.Text = ""
txtPeminjam.Text = ""
txtNoAnggota.Text = ""
txtTglPinjam.Text = ""
txtKodeBuku.Text = ""
txtNamaBuku.Text = ""
comStatus.Text = ""
comKondisi.Text = ""
txtDenda.Text = ""
txtQTY.Text = ""
grid_pinjam.Clear
End Sub

Private Sub bersih()
txtKodeBuku.Text = ""
txtNamaBuku.Text = ""
comStatus.Text = ""
comKondisi.Text = ""
txtDenda.Text = ""
txtQTY.Text = ""
grid_pinjam.Clear
End Sub

Private Sub cmdBatal_Click()
Call bersih
Call nonaktif
lblbukupinjam.Caption = 0
lblbukukembali.Caption = 0
lbltotdenda.Caption = 0
Call awal
End Sub

Private Sub cmdCetak_Click()
Call cetak
End Sub

Private Sub cmdInput_Click()
Call bersih
Call aktif
Call jum_denda
txtNoPinjam.SetFocus
End Sub

Private Sub cmdKeluar_Click()
If cmdSimpan.Enabled = True Then
    MsgBox "Simpan dahulu data anda", vbExclamation, "Aplikasi Perpustakaan"
Else
    Unload Me
End If
End Sub

Private Sub cmdSimpan_Click()
Call update_peminjaman
cmdSimpan.Enabled = False
cmdCetak.Enabled = True
cmdBatal.Enabled = False
cmdInput.Enabled = True
End Sub

Private Sub comKondisi_Click()
If comKondisi.Text = "RUSAK" Then
    txtDenda.Text = harga_buku * Val(lblpersendenda.Caption) / 100
Else
    txtDenda.Text = 0
End If
End Sub

Private Sub comStatus_Click()
If comStatus.Text = "TIDAK KEMBALI" Then
    comKondisi.Enabled = False
    comKondisi.Text = "HILANG"
    txtDenda.Text = harga_buku
Else
    txtDenda.Text = 0
    comKondisi.Enabled = True
    comKondisi.Text = ""
End If
End Sub

Private Sub Form_Load()
comStatus.AddItem "KEMBALI"
comStatus.AddItem "TIDAK KEMBALI"

comKondisi.AddItem "BAIK"
comKondisi.AddItem "RUSAK"

txtStaff.Text = nama_staff
txtTglKembali.Text = Format(Date, "YYYY-MM-DD")
lbltotdenda.Caption = 0
lblbukukembali.Caption = 0
lblbukupinjam.Caption = 0
Call nonaktif
End Sub

Private Sub txtKodeBuku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi_db
    
    dt_buku.Open "select * from V_DETIL_PINJAM where NO_PINJAM='" & txtNoPinjam.Text & "' and KD_BUKU='" & txtKodeBuku.Text & "'", koneksi
    With dt_buku
        If .BOF And .EOF Then
            MsgBox "Kode Buku '" + txtKodeBuku.Text + "' Tidak Ditemukan di Peminjaman nomor :" + txtNoPinjam.Text + ".", vbInformation, "Aplikasi Perpustakaan"
            txtKodeBuku.Text = ""
            txtNamaBuku.Text = ""
            txtQTY.Text = ""
            txtKodeBuku.SetFocus
        Else
            txtNamaBuku.Text = .Fields("NM_BUKU")
            comStatus.Text = .Fields("STATUS")
            comKondisi.Text = .Fields("KONDISI")
            txtDenda.Text = .Fields("DENDA")
            txtQTY.Text = .Fields("QTY")
            harga_buku = Val(.Fields("HARGA"))
            
            dt_custom.Open "select STOK from T_BUKU where KD_BUKU='" & txtKodeBuku.Text & "'", koneksi
            stoknya = dt_custom!stok
            dt_custom.Close
            
            txtQTY.SetFocus
        End If
    End With
    dt_buku.Close
End If
End Sub

Private Sub txtNoPinjam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi_db

dt_pinjam.Open "select * from V_PEMINJAMAN where NO_PINJAM like '%" & txtNoPinjam.Text & "%'", koneksi

With dt_pinjam
    If .BOF And .EOF Then
        MsgBox "Data Peminjaman tidak ditemukan", vbExclamation, "Aplikasi Perpustakaan"
    Else
        txtNoPinjam.Text = dt_pinjam!NO_PINJAM
        txtPeminjam.Text = dt_pinjam!NM_ANGGOTA
        txtNoAnggota.Text = dt_pinjam!NO_ANGGOTA
        txtTglPinjam.Text = Format(dt_pinjam!TGL_PINJAM, "YYYY-MM-DD")
        dt_pinjam.Close
        
        Call tampil_detilpinjam
        Call jum_buku
        txtKodeBuku.SetFocus
        
            Dim selisihnya As Double
            lama_hari = CDate(txtTglKembali.Text) - CDate(txtTglPinjam.Text)
                            
            MsgBox lama_hari
            If Val(lama_hari) > max_pinjam Then
                selisihnya = Val(lama_hari) - max_pinjam
                lbltotdenda.Caption = denda_harian * selisihnya
            End If
    End If
End With
End If
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
Dim jum_awal As Integer

If KeyAscii = 13 Then
jum_awal = Val(lblbukukembali.Caption)
    
    If comStatus.Text = "-" Then
        MsgBox "Status harus dipilih terlebih dahulu", vbExclamation, "Aplikasi Perpustakaan"
    Else
        If Val(lblbukukembali.Caption) > Val(lblbukupinjam.Caption) Then
            MsgBox "Ada kesalahan, buku yang dikembalikan lebih banyak dari yang dipinjam.", vbExclamation, "Aplikasi Perpustakaan"
            lblbukukembali.Caption = jum_awal
            txtQTY.SetFocus
        Else
            'Count qty buku yg dikembalikan
            lblbukukembali.Caption = Val(lblbukukembali.Caption) + Val(txtQTY.Text)
            update_detilpinjam
            Call tambahin_stok
            lbltotdenda.Caption = Val(lbltotdenda.Caption) + Val(txtDenda.Text)
        
            bersih
            txtKodeBuku.SetFocus
            tampil_detilpinjam
        End If
    End If
End If
End Sub
