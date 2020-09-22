VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCIPOBarang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PO - Pesan Barang"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
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
   ScaleHeight     =   3675
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   5355
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   1
         Top             =   660
         Width           =   1035
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   2
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   2
         Left            =   2220
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1755
      End
      Begin MSDataListLib.DataCombo dcInventory 
         Height          =   315
         Left            =   2220
         TabIndex        =   0
         Top             =   300
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga per"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   12
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Barang dipesan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label lblSatuan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   3420
         TabIndex        =   9
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblSatuan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Sub Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Simpan"
      Height          =   795
      Left            =   1920
      Picture         =   "frmCIPOBarang.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&elesai"
      Height          =   795
      Left            =   2760
      Picture         =   "frmCIPOBarang.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2820
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   120
      Picture         =   "frmCIPOBarang.frx":0614
      Stretch         =   -1  'True
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Purchase Order"
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
      Left            =   720
      TabIndex        =   13
      Top             =   300
      Width           =   2910
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   435
      Left            =   60
      Top             =   240
      Width           =   5355
   End
End
Attribute VB_Name = "frmCIPOBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAddBarang As ADODB.Recordset
Dim rsInventory As ADODB.Recordset

Dim sSatuanBeli As String
Dim iFSatuanKecil As Single

Private Sub ClearForm()
  dcInventory.Text = ""
  For iCounter = 0 To 2
    txtFields(iCounter).Text = 0
  Next
  For iCounter = 0 To 1
    lblSatuan(iCounter).Caption = "Satuan"
  Next
End Sub

Private Sub Form_Load()
  Set rsAddBarang = New ADODB.Recordset
  rsAddBarang.Open "C_IStockCard", db, adOpenStatic, adLockOptimistic
  '
  Set rsInventory = New ADODB.Recordset
  rsInventory.Open "SELECT NamaSupplier, Inventory, NamaInventory FROM [C_Supplier_Inventory] Where Supplier='" & frmCIPOrder.txtFields(2).Text & "'", db, adOpenStatic, adLockReadOnly
  Set dcInventory.RowSource = rsInventory
  dcInventory.ListField = "NamaInventory"
  dcInventory.BoundColumn = "Inventory"
  '
  ClearForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsInventory.Close
  Set rsInventory = Nothing
  rsAddBarang.Close
  Set rsAddBarang = Nothing
  sFormAktif = "FRMCIPORDER"
End Sub

Private Sub cmdUpdate_Click()
  '
  ' Cek Barang
  If dcInventory.Text = "" Then
    MsgBox "Data Inventory yang dipesan tidak boleh kosong", vbCritical
    dcInventory.SetFocus
    Exit Sub
  End If
  iCounter = JumlahRecord("Select * From [C_IStockCard] Where Inventory = '" & dcInventory.BoundText & "' And PO = '" & frmCIPOrder.txtFields(0).Text & "'", db)
  If iCounter > 0 Then
    MsgBox "Barang telah ada dalam daftar PO", vbCritical
    ClearForm
    Exit Sub
  End If
  '
  rsAddBarang.AddNew
  rsAddBarang!Tanggal = frmCIPOrder.txtFields(1).Text
  rsAddBarang!Inventory = dcInventory.BoundText
  rsAddBarang!NamaInventory = dcInventory.Text
  rsAddBarang!JumlahPO = txtFields(0).Text
  rsAddBarang!SisaPO = txtFields(0).Text
  rsAddBarang!SatuanBesar = lblSatuan(0).Caption
  rsAddBarang!PO = frmCIPOrder.txtFields(0).Text
  rsAddBarang!Keterangan = frmCIPOrder.txtFields(6).Text
  rsAddBarang!Harga = txtFields(1).Text
  rsAddBarang!UnitPesan = txtFields(0).Text * iFSatuanKecil
  rsAddBarang!HargaSubTotal = txtFields(2).Text
  rsAddBarang.Update
  '
  ' Isi File History
  Dim rsHistoryDetail As New ADODB.Recordset
  rsHistoryDetail.Open "HistoryDetail", db, adOpenStatic, adLockOptimistic
  rsHistoryDetail.AddNew
  rsHistoryDetail!KodeRef = frmCIPOrder.txtFields(0).Text
  rsHistoryDetail!Inventory = dcInventory.BoundText
  rsHistoryDetail!NamaInventory = dcInventory.Text
  rsHistoryDetail!Jumlah = txtFields(0).Text
  rsHistoryDetail!Satuan = lblSatuan(0).Caption
  rsHistoryDetail!HargaSatuan = txtFields(1).Text
  rsHistoryDetail!SubTotal = txtFields(2).Text
  rsHistoryDetail.Update
  rsHistoryDetail.Close
  Set rsHistoryDetail = Nothing
  '
  ' Update File Inventory
  Dim rsItemInventory As New ADODB.Recordset
  rsItemInventory.Open "C_Inventory", db, adOpenStatic, adLockOptimistic
  rsItemInventory.Find "Inventory='" & dcInventory.BoundText & "'"
  rsItemInventory!QtyOnOrder = rsItemInventory!QtyOnOrder + txtFields(0).Text
  rsItemInventory.Update
  
  frmCIPOrder.txtFields(7).Text = CCur(frmCIPOrder.txtFields(7).Text) + CCur(txtFields(2).Text)
  MsgBox "Barang telah masuk dalam daftar PO", vbInformation
  ClearForm
  '
  dcInventory.SetFocus
  '
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub dcInventory_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    If dcInventory.Text = "" Then Exit Sub
    '
    Dim rsCariInventory As New ADODB.Recordset
    rsCariInventory.Open "SELECT NamaSupplier, Inventory, NamaInventory, LastPrice FROM [C_Supplier_Inventory] Where Supplier='" & frmCIPOrder.txtFields(2).Text & "' and Inventory='" & dcInventory.BoundText & "'", db, adOpenStatic, adLockReadOnly
    txtFields(1).Text = rsCariInventory!LastPrice
    rsCariInventory.Close
    Set rsCariInventory = Nothing
    '
    Dim rsDetailInventory As New ADODB.Recordset
    rsDetailInventory.Open "Select Inventory, SatuanBesar, FSatuanKecil From [C_Inventory] Where Inventory='" & dcInventory.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If rsDetailInventory.RecordCount <> 0 Then
      sSatuanBeli = rsDetailInventory!SatuanBesar
      iFSatuanKecil = rsDetailInventory!FSatuanKecil
      lblSatuan(0).Caption = sSatuanBeli
      lblSatuan(1).Caption = sSatuanBeli
    End If
    rsDetailInventory.Close
    Set rsDetailInventory = Nothing
    '
    SendKeys "{TAB}"
  End If
  '
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim strValid As String
  '
  strValid = "0123456789."
  '
  If KeyAscii > 26 Then
    If InStr(strValid, Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
    End If
  End If
  '
  If KeyAscii = 13 Then
    Select Case Index
      Case 0
        If txtFields(0).Text = "" Then txtFields(0).Text = 0
        SendKeys "{TAB}"
      Case 1
        txtFields(2).Text = txtFields(0).Text * txtFields(1).Text
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub
