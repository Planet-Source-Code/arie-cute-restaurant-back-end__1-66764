VERSION 5.00
Begin VB.Form frmSetDatabase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Lokasi Database"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&elesai"
      Height          =   795
      Left            =   2820
      Picture         =   "frmSetDatabase.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Simpan"
      Height          =   795
      Left            =   1980
      Picture         =   "frmSetDatabase.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   795
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   5415
      Begin VB.TextBox txtPathData 
         Height          =   315
         Left            =   240
         MaxLength       =   128
         TabIndex        =   0
         Top             =   1140
         Width           =   4935
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   300
         Picture         =   "frmSetDatabase.frx":0614
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOKASI DATABASE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   1020
         TabIndex        =   6
         Top             =   300
         Width           =   2640
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path Database"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   900
         Width           =   1245
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmSetDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fNum As Integer
Dim bolInfo As Boolean

Private Sub Form_Load()
  bolInfo = False
  txtPathData.Text = App.Path
End Sub

Private Sub cmdBrowse_Click()
  '
  Dim lpIDList As Long
  Dim sBuffer As String
  Dim szTitle As String
  Dim tBrowseInfo As BrowseInfo

  szTitle = "Pilih sebuah folder yang akan menjadi tempat penyimpanan database "
  
  With tBrowseInfo
   .hWndOwner = Me.hWnd
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With

  lpIDList = SHBrowseForFolder(tBrowseInfo)

  If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   txtPathData.Text = sBuffer
  End If
  '
End Sub

Private Sub cmdUpdate_Click()
  '
  If txtPathData = "" Then
    MsgBox "Lokasi Database tidak boleh kosong", vbCritical
    cmdBrowse.SetFocus
    Exit Sub
  End If
  '
  If Len(txtPathData.Text) = 3 Then
    txtPathData.Text = Left$(txtPathData.Text, 2)
  End If
  fNum = FreeFile()
  Open App.Path & "\Hospitality.fnb" For Output As #fNum
  Print #fNum, txtPathData.Text
  Close #fNum
  '
  sPesan = "Lokasi Database sekarang ialah " & txtPathData.Text
  MsgBox sPesan, vbInformation
  '
  bolInfo = True
  '
End Sub

Private Sub cmdClose_Click()
  '
  If bolInfo Then
    sPesan = "Untuk mengaktifkan pengaturan lokasi database yang baru, anda harus" & vbCrLf
    sPesan = sPesan & "keluar dan menjalankan Winsor F & B Control kembali"
    MsgBox sPesan, vbExclamation
  End If
  Unload Me
  '
End Sub
