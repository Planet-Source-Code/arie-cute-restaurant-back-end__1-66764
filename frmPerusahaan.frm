VERSION 5.00
Begin VB.Form frmPerusahaan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identitas Perusahaan"
   ClientHeight    =   3165
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmPerusahaan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   180
      TabIndex        =   7
      Top             =   720
      Width           =   4755
      Begin VB.TextBox txtFields 
         DataField       =   "KodePos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1140
         TabIndex        =   3
         Top             =   1440
         Width           =   1155
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Negara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1140
         TabIndex        =   4
         Top             =   1800
         Width           =   1155
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telepon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   3180
         TabIndex        =   5
         Top             =   1260
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Fax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   3180
         TabIndex        =   6
         Top             =   1620
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Perusahaan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   0
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   1140
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1140
         TabIndex        =   2
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1500
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negara"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1860
         Width           =   525
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   2460
         TabIndex        =   13
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   2460
         TabIndex        =   12
         Top             =   1680
         Width           =   270
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   180
         Width           =   405
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1140
         Width           =   330
      End
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   4635
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPerusahaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPerusahaan As ADODB.Recordset
Dim i As Integer

Private Sub Form_Load()
  '
  sPesan = "Masukkan profil perusahaan anda. "
  sPesan = sPesan & "Informasi ini akan disimpan dalam database aplikasi."
  lblLabels(10).Caption = sPesan
  '
  Set rsPerusahaan = New ADODB.Recordset
  rsPerusahaan.Open "Select * from [Perusahaan]", db, adOpenStatic, adLockOptimistic
  '
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = rsPerusahaan
  Next
  '
  If (rsPerusahaan.RecordCount = 0) And (bGrant = True) Then
    rsPerusahaan.AddNew
  Else
    Frame1.Enabled = False
  End If
  '
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
  Dim i As Byte
  '
  If bGrant Then
    For i = 0 To 4
      If txtFields(i).Text = "" Then
        MsgBox "Field ini dibutuhkan", vbCritical
        txtFields(i).SetFocus
        Cancel = -1
        Exit Sub
      End If
    Next
    '
    rsPerusahaan.Update
  End If
  '
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub
