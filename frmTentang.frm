VERSION 5.00
Begin VB.Form frmTentang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   4320
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTentang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2981.741
   ScaleMode       =   0  'User
   ScaleWidth      =   7141.487
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Label lblProfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   1020
         TabIndex        =   7
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label lblProfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1020
         TabIndex        =   6
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label lblProfile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   1020
         TabIndex        =   5
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program ini dilisensikan kepada :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   600
         TabIndex        =   4
         Top             =   1860
         Width           =   2820
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single Store Edition"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4380
         TabIndex        =   3
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.winsor.co.id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5820
         TabIndex        =   2
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   300
         Picture         =   "frmTentang.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   5835
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1020
         TabIndex        =   1
         Top             =   3240
         Width           =   810
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         Picture         =   "frmTentang.frx":2FD5
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmTentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sPesan As String

Private Sub Form_Load()
  lblCopyright.Caption = "Hak Cipta (c) 2001 oleh PT Winsor Satria Persada"
  '
  lblProfile(0).Caption = ICNama
  lblProfile(1).Caption = ICAlamat
  lblProfile(2).Caption = ICKota
  '
End Sub

Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 27) Or (KeyAscii = 13) Then
    Unload Me
  End If
End Sub

Private Sub Frame1_Click()
  Unload Me
End Sub
