VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form Setting --"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fdenda 
      Height          =   4095
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   8055
      Begin VB.TextBox txtPersenDenda 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "0"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtDendaHarian 
         Height          =   285
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   13
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblketpersen 
         Caption         =   "N/A"
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   7575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "% dari harga buku"
         Height          =   195
         Left            =   3120
         TabIndex        =   18
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Penggantian Buku"
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
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblketharian 
         Caption         =   "N/A"
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "/ Hari"
         Height          =   195
         Left            =   3120
         TabIndex        =   14
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Denda Harian"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   7200
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame fpeminjaman 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   8055
      Begin VB.TextBox txtMaxBuku 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtLamaPinjam 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblketmaxbuku 
         Caption         =   "N/A"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   7575
      End
      Begin VB.Label lblketlamapinjam 
         Caption         =   "N/A"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Buku"
         Height          =   195
         Left            =   2640
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Maks Buku "
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
         TabIndex        =   5
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hari"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lama Peminjaman"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1530
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Peminjaman"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Denda"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Setting - Aplikasi Perpustakaan
'copyright(c)2016 Yudha Tri Putra
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private Sub tampil_data()
Call koneksi_db

'load lama pinjam
dt_setting.Open "select NILAI,KETERANGAN from T_SETTINGS where NAMA_SETTING='LAMA_PINJAM'", koneksi
txtLamaPinjam.Text = dt_setting!NILAI
lblketlamapinjam.Caption = dt_setting!KETERANGAN
dt_setting.Close

'load maksimal jumlah buku pinjam
dt_setting.Open "select NILAI,KETERANGAN from T_SETTINGS where NAMA_SETTING='MAX_BUKU'", koneksi
txtMaxBuku.Text = dt_setting!NILAI
lblketmaxbuku.Caption = dt_setting!KETERANGAN
dt_setting.Close

'load denda perhari
dt_setting.Open "select NILAI,KETERANGAN from T_SETTINGS where NAMA_SETTING='DENDA_PERHARI'", koneksi
txtDendaHarian.Text = dt_setting!NILAI
lblketharian.Caption = dt_setting!KETERANGAN
dt_setting.Close

'load persen denda perharga buku
dt_setting.Open "select NILAI,KETERANGAN from T_SETTINGS where NAMA_SETTING='PERSEN_DENDA'", koneksi
txtPersenDenda.Text = dt_setting!NILAI
lblketpersen.Caption = dt_setting!KETERANGAN
dt_setting.Close
End Sub

Private Sub cmdSimpan_Click()
update = "update T_SETTINGS set NILAI ='" & txtLamaPinjam.Text & "' where NAMA_SETTING='LAMA_PINJAM'"
koneksi.Execute update

update = "update T_SETTINGS set NILAI ='" & txtMaxBuku.Text & "' where NAMA_SETTING='MAX_BUKU'"
koneksi.Execute update

update = "update T_SETTINGS set NILAI ='" & txtDendaHarian.Text & "' where NAMA_SETTING='DENDA_PERHARI'"
koneksi.Execute update

update = "update T_SETTINGS set NILAI ='" & txtPersenDenda.Text & "' where NAMA_SETTING='PERSEN_DENDA'"
koneksi.Execute update

MsgBox "Data Setting Berhasil diupdate", vbInformation, " Aplikasi Perpustakaan"
End Sub

Private Sub Form_Load()
fdenda.Visible = False
Call tampil_data
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(1).Selected = True Then
    fpeminjaman.Visible = True
    fdenda.Visible = False
ElseIf TabStrip1.Tabs(2).Selected = True Then
    fpeminjaman.Visible = False
    fdenda.Visible = True
End If
End Sub
