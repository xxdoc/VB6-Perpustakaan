VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-- Form About --"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "frmAbout.frx":0000
      Left            =   120
      List            =   "frmAbout.frx":0016
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Copyright(c)2016 . Bina Sarana Informatika"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   3030
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Dibuat Oleh :"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0"
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
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Aplikasi Perpustakaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form About - Aplikasi Perpustakaan
'copyright(c)2016
' - Yudha Tri Putra
' - Asti Aprilliyanti
' - Bangun Subkhi Ismawanto
' - Manan Sabili
' - Dwi Hardianto Putra
' - Fera Waningsih

Private Sub cmdOK_Click()
Unload Me
End Sub
