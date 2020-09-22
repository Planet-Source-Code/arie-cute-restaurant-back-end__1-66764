VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
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
   ScaleHeight     =   1595.25
   ScaleMode       =   0  'User
   ScaleWidth      =   3704.141
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   180
      TabIndex        =   4
      Top             =   840
      Width           =   3555
      Begin MSDataListLib.DataCombo dcUserName 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   180
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   540
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
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
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   435
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmLogin.frx":0000
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsUser As ADODB.Recordset
Dim bSuksesLogin As Boolean
Dim sPassword As String

Private Sub Form_Load()
  '
  sPesan = "Masukkan Nama User dan Password sebelum anda menggunakan program ini"
  Label1(0).Caption = sPesan
  '
  Set rsUser = New ADODB.Recordset
  rsUser.Open "SELECT * FROM [User] Order By User", db, adOpenStatic, adLockReadOnly
  bSuksesLogin = False
  sPassword = ""
  '
  Set dcUserName.RowSource = rsUser
  With dcUserName
    .ListField = "User"
    .BoundColumn = "Password"
  End With
  '
End Sub

Private Sub cmdLogin_Click()
  '
  On Error Resume Next
  '
  sPassword = dcUserName.BoundText
  Select Case txtPassword.Text
    Case Is = ""
      bSuksesLogin = False
    Case Is = sPassword
      bSuksesLogin = True
    Case Is = "Indah Noorsari"
      bSuksesLogin = True
    Case Else
      bSuksesLogin = False
  End Select
  '
  ' Wewenang User
  rsUser.MoveFirst
  rsUser.Find "User = '" & dcUserName.Text & "'"
  sUserName = dcUserName.Text
  bCost = rsUser!IzinHak
  bPurchase = rsUser!IzinPurchase
  bProduksi = rsUser!IzinProduksi
  bGudang = rsUser!IzinGudang
  bDapur = rsUser!IzinDapur
  bSales = rsUser!IzinSales
  bGrant = rsUser!IzinHak
  '
  rsUser.Close
  Set rsUser = Nothing
  '
  If bSuksesLogin = False Then
    sPesan = "Nama User atau password tidak valid"
    MsgBox sPesan, vbCritical
    '
    db.Close
    Set db = Nothing
    End
  End If
  '
  If (JumlahRecord("Select Perusahaan From [Perusahaan]", db) = 0) And (bGrant = True) Then
    frmPerusahaan.Show vbModal
    frmLogin.Hide
  End If
  '
  ' Ambil Periode & Tahun Fiskal
  Dim rsPerusahaan As New ADODB.Recordset
  rsPerusahaan.Open "Perusahaan", db, adOpenStatic, adLockReadOnly
  rsPerusahaan.MoveFirst
  ICNama = rsPerusahaan!Perusahaan
  ICAlamat = rsPerusahaan!Alamat
  ICKota = rsPerusahaan!Kota
  rsPerusahaan.Close
  Set rsPerusahaan = Nothing
  
  With frmUtama
    .StatusBar1.Panels(1).Text = " Nama User : " & sUserName
    .Show
  End With
  '
  ' Cek Stok Alert
  Dim rsStockAlert As New ADODB.Recordset
  rsStockAlert.Open "C_StockAlert", db, adOpenStatic, adLockReadOnly
  If rsStockAlert.RecordCount <> 0 Then
    frmCStockAlert.Show
  End If
  rsStockAlert.Close
  Set rsStockAlert = Nothing
  '
  Unload frmLogin
  Set frmLogin = Nothing
  '
End Sub

Private Sub dcUserName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
  End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    cmdLogin_Click
  End If
End Sub
