VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Login --"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Batal"
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtPass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtUser 
         Height          =   375
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Silahkan isi username dan password anda dengan benar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   795
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Copyright(c)2016 - BSI Margonda Depok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   240
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4215
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   -720
      Width           =   3135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Login - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private coba As Integer

Private Sub update_login()
TANGGAL = Format(Now, "YYYY-MM-DD:HH-MM")

update = "update T_STAFF set LAST_LOGIN='" & TANGGAL & "' where KD_STAFF='" & txtUser.Text & "'"
koneksi.Execute (update)
End Sub
Private Sub kunci_akun()
update = "update T_STAFF set LOCKED='TRUE' where KD_STAFF='" & txtUser.Text & "'"
koneksi.Execute (update)
End Sub

Private Sub buka_master()
If hak_akses = "ADMIN" Then
    frmMaster.mn_data_anggota.Enabled = True
    frmMaster.mn_data_buku.Enabled = True
    frmMaster.mn_data_staff.Enabled = True
    frmMaster.mn_gpass.Enabled = True
    frmMaster.mn_peminjaman.Enabled = True
    frmMaster.mn_pengembalian.Enabled = True
    frmMaster.mn_set_apl.Enabled = True
    frmMaster.mn_lap_data_anggota.Enabled = True
    frmMaster.mn_lap_data_buku.Enabled = True
    frmMaster.mn_lap_pinjam.Enabled = True
    frmMaster.mn_history.Enabled = True
    frmMaster.mn_logout.Enabled = True
ElseIf hak_akses = "USER" Then
    frmMaster.mn_data_anggota.Enabled = True
    frmMaster.mn_data_buku.Enabled = True
    frmMaster.mn_data_staff.Enabled = False
    frmMaster.mn_gpass.Enabled = True
    frmMaster.mn_peminjaman.Enabled = True
    frmMaster.mn_pengembalian.Enabled = True
    frmMaster.mn_set_apl.Enabled = False
    frmMaster.mn_lap_data_anggota.Enabled = True
    frmMaster.mn_lap_data_buku.Enabled = True
    frmMaster.mn_lap_pinjam.Enabled = True
    frmMaster.mn_history.Enabled = True
    frmMaster.mn_logout.Enabled = True
Else
    MsgBox "Hak akses anda tidak dikenal atau tidak diketahui,hubungi admin", vbExclamation, "Aplikasi Perpustakaan"
End If
End Sub

Private Sub bersih()
    txtUser.Text = ""
    txtPass.Text = ""
    txtUser.SetFocus
End Sub
Private Sub cmdBatal_Click()
    Call bersih
End Sub

Private Sub cmdLogin_Click()
Dim num_coba As Integer
Dim is_kunci As String
Call koneksi_db

dt_user.Open "select * from T_STAFF where KD_STAFF ='" & txtUser.Text & "' and PASS='" & txtPass.Text & "'", koneksi
dt_setting.Open "select NILAI from T_SETTINGS where NAMA_SETTING ='KESEMPATAN_LOGIN'", koneksi
num_coba = dt_setting!NILAI

    If dt_user.EOF Then
            If coba > num_coba Then
                MsgBox "Akun anda terkunci karena salah login sebanyak " & num_coba & " x," & Chr(13) & _
                       " hubungi admin untuk info lebih lanjut.", vbExclamation, "Aplikasi Perpustakaan"
                Call kunci_akun
            Else
                MsgBox "kesempatan ke " & coba & " salah" & Chr(13) & _
                " Username atau Password '" & txtPass.Text & "' salah", vbExclamation, "Aplikasi Perpustakaan"
                
                txtPass.Text = ""
                txtPass.SetFocus
                coba = coba + 1
            End If
    Else
        nama_staff = dt_user!NM_STAFF
        kode_staff = dt_user!KD_STAFF
        hak_akses = dt_user!Status
        is_kunci = dt_user!Locked
        dt_user.Close
        dt_setting.Close
        
        If is_kunci = "TRUE" Then
            MsgBox "Akun anda terkunci, hubungi administrator", vbExclamation, "Aplikasi Perpustakaan"
            Call bersih
        Else
            Call update_login
            MsgBox "Selamat datang " & nama_staff & " , Anda berhasil Login", vbInformation, "Aplikasi Perpustakaan"
            Call buka_master
            frmMaster.sbar.Panels(1).Text = nama_staff
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Activate()
coba = 1
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdLogin.SetFocus
End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPass.SetFocus
End If
End Sub
