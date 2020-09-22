VERSION 5.00
Begin VB.Form frmCAddSupplier 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menambah Supplier"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLookUp 
      Caption         =   "..."
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Width           =   315
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "Supplier"
      Height          =   315
      Index           =   2
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   1
      Top             =   960
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&elesai"
      Height          =   795
      Left            =   3120
      Picture         =   "frmCAddSupplier.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1500
      Width           =   795
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Simpan"
      Height          =   795
      Left            =   2280
      Picture         =   "frmCAddSupplier.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1500
      Width           =   795
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Supplier"
      Height          =   315
      Index           =   0
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1035
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00E0E0E0&
      DataField       =   "NamaSupplier"
      Height          =   315
      Index           =   1
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Satuan"
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
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   1020
      Width           =   600
   End
   Begin VB.Label lblInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga per"
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
      Left            =   300
      TabIndex        =   6
      Top             =   1020
      Width           =   840
   End
   Begin VB.Label lblInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Kode Supplier"
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
      Index           =   0
      Left            =   300
      TabIndex        =   5
      Top             =   300
      Width           =   2070
   End
End
Attribute VB_Name = "frmCAddSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAddSupp As New ADODB.Recordset

Private Sub Form_Load()
  '
  sFormAktif = Me.Name
  Set rsAddSupp = New ADODB.Recordset
  rsAddSupp.Open "C_Supplier_Inventory", db, adOpenStatic, adLockOptimistic
  '
  lblInv(3).Caption = frmCInventory.txtFields(7).Text
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsAddSupp.Close
  Set rsAddSupp = Nothing
  sFormAktif = "FRMCINVENTORY"
End Sub

Private Sub cmdUpdate_Click()
  '
  ' Cek Supplier
  For iCounter = 0 To 2 Step 2
    If txtFields(iCounter).Text = "" Then
      MsgBox "Field ini dibutuhkan", vbCritical
      txtFields(iCounter).SetFocus
      Exit Sub
    End If
  Next
  If IsNumeric(txtFields(2).Text) = False Then
    MsgBox "Field ini bertipe numerik", vbCritical
    txtFields(2).SetFocus
    SendKeys "{END}+{HOME}"
    Exit Sub
  End If
  '
  iCounter = JumlahRecord("Select * From [C_Supplier_Inventory] Where Supplier = '" & txtFields(0).Text & "' And Inventory = '" & frmCInventory.txtFields(0).Text & "'", db)
  If iCounter > 0 Then
    MsgBox "Supplier telah ada", vbCritical
    txtFields(0).Text = ""
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    txtFields(0).SetFocus
    Exit Sub
  End If
  '
  rsAddSupp.AddNew
  rsAddSupp!Inventory = frmCInventory.txtFields(0).Text
  rsAddSupp!NamaInventory = frmCInventory.txtFields(1).Text
  rsAddSupp!Supplier = txtFields(0).Text
  rsAddSupp!NamaSupplier = txtFields(1).Text
  rsAddSupp!LastPrice = CCur(txtFields(2).Text)
  rsAddSupp.Update
  '
  MsgBox "Supplier baru telah disimpan dalam database", vbInformation
  txtFields(0).Text = ""
  txtFields(1).Text = ""
  txtFields(2).Text = ""
  txtFields(0).SetFocus
  '
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    Select Case Index
      Case 0
        sKontrolAktif = "AIA_1"
        Dim rsSupplier As New ADODB.Recordset
        rsSupplier.Open "Select Supplier, NamaSupplier from [C_Supplier] Where Supplier='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsSupplier.RecordCount <> 0 Then
          txtFields(1).Text = rsSupplier!NamaSupplier
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(1).Text = ""
          frmCari.Show vbModal
          txtFields(0).SetFocus
        End If
        rsSupplier.Close
        sKontrolAktif = ""
      Case 2
        If Not IsNumeric(txtFields(2).Text) Then
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub cmdLookUp_Click()
  sKontrolAktif = "AIA_1"
  frmCari.Show vbModal
  txtFields(0).SetFocus
  If txtFields(0).Text <> "" Then SendKeys "{ENTER}"
End Sub
