VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCPilihResep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resep yang akan dibuat"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   4815
      Begin MSDataListLib.DataCombo dcResep 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   300
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Resep"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Resep"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   345
      Left            =   780
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   180
      Picture         =   "frmCPilihResep.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   525
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   495
      Left            =   60
      Top             =   180
      Width           =   4815
   End
End
Attribute VB_Name = "frmCPilihResep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPilihResep As ADODB.Recordset

Private Sub Form_Load()
  '
  Set rsPilihResep = New ADODB.Recordset
  rsPilihResep.Open "Select Inventory, NamaInventory From [C_Inventory] Where IsResep=True", db, adOpenStatic, adLockReadOnly
  Set dcResep.RowSource = rsPilihResep
  dcResep.ListField = "NamaInventory"
  dcResep.BoundColumn = "Inventory"
  '
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsPilihResep.Close
  Set rsPilihResep = Nothing
End Sub

Private Sub dcResep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If dcResep.Text = "" Then Exit Sub
    SendKeys "{TAB}"
  End If
End Sub

Private Sub cmdCetak_Click()
  If dcResep.Text = "" Then
    MsgBox "Pilih dulu resep yang akan dibuat", vbInformation
  Else
    '
    Dim sSQL As String
    '
    sSQL = "SELECT C_Resep.Resep, C_Resep.NamaResep, C_Resep.NamaSatuan, C_ResepBahan.Inventory, C_ResepBahan.NamaSatuan, C_ResepBahan.NamaInventory FROM C_Resep INNER JOIN C_ResepBahan ON C_Resep.Resep = C_ResepBahan.Resep WHERE (((C_Resep.Resep)='" & dcResep.BoundText & "'))"

    With frmUtama.datReport
      .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
      .RecordSource = sSQL
      .Refresh
    End With
    With frmUtama.crReport
      .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
      .WindowTitle = "Formulir Permintaan Pembuatan Resep"
      .ReportFileName = App.Path & "\FormBuatResepModifier.Rpt"
      .Formulas(0) = "IDCompany= '" & ICNama & "'"
      .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
      .Formulas(2) = "IDKota= '" & ICKota & "'"
      .Action = 1
    End With
    '
  End If
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

