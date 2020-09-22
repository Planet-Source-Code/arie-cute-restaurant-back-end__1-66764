VERSION 5.00
Begin VB.Form frmCPMarkup 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faktor Mark Up"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
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
   ScaleHeight     =   2085
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Selesai"
      Height          =   795
      Left            =   1920
      Picture         =   "frmCPMarkup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   795
   End
   Begin VB.CommandButton cmdHitung 
      Caption         =   "&Hitung"
      Height          =   795
      Left            =   1080
      Picture         =   "frmCPMarkup.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
      Width           =   795
   End
   Begin VB.TextBox txtMarkUp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   2100
      MaxLength       =   3
      TabIndex        =   0
      Top             =   300
      Width           =   915
   End
   Begin VB.Label lblInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faktor Mark Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   2
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label lblInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   7
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   195
   End
End
Attribute VB_Name = "frmCPMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtMarkUp_KeyPress(KeyAscii As Integer)
  Dim strValid As String
  '
  strValid = "0123456789,."
  
  If KeyAscii > 26 Then
    If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
    End If
  End If
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
  '
End Sub

Private Sub cmdHitung_Click()
  '
  With frmCProdukMenu
    .txtFields(5).Text = txtMarkUp.Text
    .txtFields(6).Text = CCur(.txtFields(4).Text) + ((txtMarkUp.Text / 100) * CCur(.txtFields(4).Text))
    .txtFields(7).Text = CCur(.txtFields(6).Text) - CCur(.txtFields(4).Text)
  End With
  cmdClose_Click
  frmCProdukMenu.txtFields(6).SetFocus
  '
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

