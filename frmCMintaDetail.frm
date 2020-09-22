VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCMintaDetail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory yang akan diminta"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
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
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&elesai"
      Height          =   795
      Left            =   2940
      Picture         =   "frmCMintaDetail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Simpan"
      Height          =   795
      Left            =   2100
      Picture         =   "frmCMintaDetail.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2820
      Width           =   795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   60
      TabIndex        =   5
      Top             =   720
      Width           =   5655
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   3960
         TabIndex        =   2
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1755
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   2
         Left            =   2220
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1755
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   1
         Top             =   660
         Width           =   675
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
      Begin VB.Label lblSatuan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   16
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "per"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   4080
         TabIndex        =   14
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Sub Total"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblSatuan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   12
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label lblSatuan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   11
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah diminta"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Inventory"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   300
         TabIndex        =   8
         Top             =   1080
         Width           =   435
      End
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   120
      Picture         =   "frmCMintaDetail.frx":0614
      Stretch         =   -1  'True
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Inventory"
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
      TabIndex        =   15
      Top             =   240
      Width           =   1845
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   495
      Left            =   60
      Top             =   180
      Width           =   5655
   End
End
Attribute VB_Name = "frmCMintaDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsMintaInventory As ADODB.Recordset
Dim rsInventory As ADODB.Recordset

Dim sSatuanBesar, sSatuanKecil As String
Dim iFSatuanKecil As Single
Dim bolResep As Boolean

Dim BatasBesar, BatasKecil, CompStockAwal, CompStockAmbil As Single

Private Sub ClearForm()
  dcInventory.Text = ""
  For iCounter = 0 To 3
    txtFields(iCounter).Text = 0
  Next
  For iCounter = 0 To 2
    lblSatuan(iCounter).Caption = "Satuan"
  Next
End Sub

Private Sub Form_Load()
  Set rsMintaInventory = New ADODB.Recordset
  rsMintaInventory.Open "C_AmbilKembaliDetail", db, adOpenStatic, adLockOptimistic
  '
  Set rsInventory = New ADODB.Recordset
  rsInventory.Open "Select Inventory, NamaInventory From [C_Inventory]", db, adOpenStatic, adLockReadOnly
  Set dcInventory.RowSource = rsInventory
  dcInventory.ListField = "NamaInventory"
  dcInventory.BoundColumn = "Inventory"
  '
  ClearForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsInventory.Close
  Set rsInventory = Nothing
  rsMintaInventory.Close
  Set rsMintaInventory = Nothing
  sFormAktif = "FRMCMINTA"
End Sub

Private Sub cmdUpdate_Click()
  '
  ' Cek Barang
  If dcInventory.Text = "" Then
    MsgBox "Data Inventory yang diterima tidak boleh kosong", vbCritical
    dcInventory.SetFocus
    Exit Sub
  End If
  iCounter = JumlahRecord("Select * From [C_AmbilKembaliDetail] Where Inventory = '" & dcInventory.BoundText & "' And PI = '" & frmCMinta.txtFields(0).Text & "'", db)
  If iCounter > 0 Then
    MsgBox "Item inventory telah diminta", vbCritical
    ClearForm
    Exit Sub
  End If
  '
  rsMintaInventory.AddNew
  rsMintaInventory!Tanggal = frmCMinta.txtFields(1).Text
  rsMintaInventory!Inventory = dcInventory.BoundText
  rsMintaInventory!NamaInventory = dcInventory.Text
  rsMintaInventory!Satuan = lblSatuan(1).Caption
  rsMintaInventory!PI = frmCMinta.txtFields(0).Text
  rsMintaInventory!UnitKeluar = Format((txtFields(0).Text * iFSatuanKecil) + txtFields(3).Text, "###,###,###.##")
  rsMintaInventory!Harga = txtFields(1).Text
  rsMintaInventory!HargaSubTotal = Format(txtFields(2).Text, "###,###,###.##")
  rsMintaInventory.Update
  '
  frmCMinta.txtFields(7).Text = CCur(frmCMinta.txtFields(7).Text) + CCur(txtFields(2).Text)
  MsgBox "Item Inventory telah masuk dalam daftar permintaan", vbInformation
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
    Dim rsDetailInventory As New ADODB.Recordset
    rsDetailInventory.Open "Select Inventory, QtyOnHand, QtyOnHandKecil, IsResep, HargaPerUnit, LastInvoicePrice, SatuanBesar, FSatuanKecil, SatuanKecil From [C_Inventory] Where Inventory='" & dcInventory.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If rsDetailInventory.RecordCount <> 0 Then
      bolResep = rsDetailInventory!IsResep
      txtFields(1).Text = Format(rsDetailInventory!HargaPerUnit, "###,###,###.##")
      If Not bolResep Then
        BatasBesar = rsDetailInventory!QtyOnHand
        BatasKecil = rsDetailInventory!QtyOnHandKecil
        With txtFields(3)
          .TabStop = True
          .Locked = False
          .BackColor = &HFFFFFF
        End With
      Else
        BatasKecil = rsDetailInventory!QtyOnHandKecil
        With txtFields(3)
          .Text = 0
          .TabStop = False
          .Locked = True
          .BackColor = &HE0E0E0
        End With
      End If
      sSatuanBesar = rsDetailInventory!SatuanBesar
      sSatuanKecil = rsDetailInventory!SatuanKecil
      iFSatuanKecil = rsDetailInventory!FSatuanKecil
    
      lblSatuan(0).Caption = sSatuanBesar
      lblSatuan(1).Caption = sSatuanKecil
      lblSatuan(2).Caption = sSatuanKecil
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
  Dim curSubTotal As Currency
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
        If Not bolResep Then
          If CSng(txtFields(0).Text) > BatasBesar Then
            MsgBox "Jumlah stok yang ada hanya sebanyak " & BatasBesar & " " & sSatuanBesar & " " & BatasKecil & " " & sSatuanKecil, vbCritical
            txtFields(0).Text = 0
            dcInventory.SetFocus
            Exit Sub
          End If
        Else
          If CSng(txtFields(0).Text) > BatasKecil Then
            MsgBox "Jumlah stok yang ada hanya sebanyak " & BatasKecil & " " & sSatuanKecil, vbCritical
            txtFields(0).Text = 0
            dcInventory.SetFocus
            Exit Sub
          End If
        End If
        txtFields(3).Text = 0
        curSubTotal = CCur(((txtFields(0).Text * iFSatuanKecil) + txtFields(3).Text) * txtFields(1).Text)
        txtFields(2).Text = Format(curSubTotal, "###,###,###.##")
        SendKeys "{TAB}"
      Case 3
        If txtFields(3).Text = "" Then txtFields(3).Text = 0
        If CSng(txtFields(3).Text) >= iFSatuanKecil Then
          txtFields(0).Text = Int(txtFields(3).Text / iFSatuanKecil)
          txtFields(3).Text = txtFields(3).Text - (txtFields(0).Text * iFSatuanKecil)
        End If
        If (CSng((txtFields(0).Text * iFSatuanKecil) + txtFields(3).Text) > (BatasBesar * iFSatuanKecil) + BatasKecil) Then
          MsgBox "Jumlah stok yang ada hanya sebanyak " & BatasBesar & " " & sSatuanBesar & " " & BatasKecil & " " & sSatuanKecil, vbCritical
          txtFields(0).Text = 0
          txtFields(3).Text = 0
          dcInventory.SetFocus
          Exit Sub
        End If
        curSubTotal = CCur(((txtFields(0).Text * iFSatuanKecil) + txtFields(3).Text) * txtFields(1).Text)
        txtFields(2).Text = Format(curSubTotal, "###,###,###.##")
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub
