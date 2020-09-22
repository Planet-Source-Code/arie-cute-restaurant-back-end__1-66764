VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCIPIBarang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PI - Terima Barang"
   ClientHeight    =   3660
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
   ScaleHeight     =   3660
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&elesai"
      Height          =   795
      Left            =   2760
      Picture         =   "frmCIPIBarang.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Simpan"
      Height          =   795
      Left            =   1920
      Picture         =   "frmCIPIBarang.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
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
      TabIndex        =   0
      Top             =   720
      Width           =   5355
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   13
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   660
         Width           =   1035
      End
      Begin MSDataListLib.DataCombo dcInventory 
         Height          =   315
         Left            =   2220
         TabIndex        =   1
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
         TabIndex        =   12
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         Left            =   3420
         TabIndex        =   7
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Barang diterima"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   720
         Width           =   1665
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   360
         Width           =   960
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
         TabIndex        =   4
         Top             =   1080
         Width           =   435
      End
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   180
      Picture         =   "frmCIPIBarang.frx":0614
      Stretch         =   -1  'True
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Purchase Invoice"
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
      TabIndex        =   14
      Top             =   300
      Width           =   3075
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
Attribute VB_Name = "frmCIPIBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAddBarang As ADODB.Recordset
Dim rsInventory As ADODB.Recordset

Dim sSatuanBeli, sSupplier, sNamaSupplier As String
Dim iFSatuanKecil, iJumlahPO As Single

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
  rsInventory.Open "SELECT C_IPO.PO, C_IStockCard.Inventory, C_IStockCard.NamaInventory FROM C_IPO INNER JOIN C_IStockCard ON C_IPO.PO = C_IStockCard.PO WHERE (((C_IPO.PO)='" & frmCIPInvoice.txtFields(2).Text & "'))", db, adOpenStatic, adLockReadOnly
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
  sFormAktif = "FRMCIPINVOICE"
End Sub

Private Sub cmdUpdate_Click()
  '
  ' Cek Barang
  If dcInventory.Text = "" Then
    MsgBox "Data Inventory yang diterima tidak boleh kosong", vbCritical
    dcInventory.SetFocus
    Exit Sub
  End If
  iCounter = JumlahRecord("Select * From [C_IStockCard] Where Inventory = '" & dcInventory.BoundText & "' And PI = '" & frmCIPInvoice.txtFields(0).Text & "'", db)
  If iCounter > 0 Then
    MsgBox "Barang telah ada dalam daftar Invoice", vbCritical
    ClearForm
    Exit Sub
  End If
  '
  rsAddBarang.AddNew
  rsAddBarang!Tanggal = frmCIPInvoice.txtFields(1).Text
  rsAddBarang!Inventory = dcInventory.BoundText
  rsAddBarang!NamaInventory = dcInventory.Text
  rsAddBarang!IsInvoice = True
  rsAddBarang!PI = frmCIPInvoice.txtFields(0).Text
  rsAddBarang!JumlahPI = txtFields(0).Text
  rsAddBarang!SatuanBesar = lblSatuan(0).Caption
  rsAddBarang!Keterangan = frmCIPInvoice.txtFields(6).Text
  rsAddBarang!Harga = txtFields(1).Text
  rsAddBarang!UnitMasuk = txtFields(0).Text * iFSatuanKecil
  rsAddBarang!HargaSubTotal = txtFields(2).Text
  rsAddBarang.Update
  '
  ' Isi File History
  Dim rsHistoryDetail As New ADODB.Recordset
  rsHistoryDetail.Open "HistoryDetail", db, adOpenStatic, adLockOptimistic
  rsHistoryDetail.AddNew
  rsHistoryDetail!KodeRef = frmCIPInvoice.txtFields(0).Text
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
  ' Update File PO
  Dim rsPO As New ADODB.Recordset
  rsPO.Open "Select Inventory, PO, JumlahPO, SisaPO From [C_IStockCard] Where Inventory='" & dcInventory.BoundText & "' and PO='" & frmCIPInvoice.txtFields(2).Text & "'", db, adOpenStatic, adLockOptimistic
  rsPO!SisaPO = rsPO!SisaPO - CSng(txtFields(0).Text)
  rsPO.Update
  rsPO.Close
  Set rsPO = Nothing
  '
  ' Update File Inventory
  Dim rsItemInventory As New ADODB.Recordset
  rsItemInventory.Open "C_Inventory", db, adOpenStatic, adLockOptimistic
  rsItemInventory.Find "Inventory='" & dcInventory.BoundText & "'"
  rsItemInventory!QtyOnHand = rsItemInventory!QtyOnHand + txtFields(0).Text
  rsItemInventory!QtyOnOrder = rsItemInventory!QtyOnOrder - txtFields(0).Text
  rsItemInventory!JumlahItem = rsItemInventory!JumlahItem + (txtFields(0).Text * iFSatuanKecil)
  rsItemInventory!LastInvoiceDate = frmCIPInvoice.txtFields(1).Text
  rsItemInventory!LastInvoicePrice = txtFields(1).Text
  rsItemInventory!HargaPerUnit = (txtFields(1).Text / iFSatuanKecil)
  rsItemInventory!Supplier = sSupplier
  rsItemInventory!NamaSupplier = sNamaSupplier
  rsItemInventory.Update
  '
  ' Update file Stock Alert
  db.Execute "Delete From [C_StockAlert] Where Inventory='" & dcInventory.BoundText & "'"
  
  ' Update Harga Bahan Resep
  db.Execute "Update [C_ResepBahan] SET HargaPerUnit=" & CCur(txtFields(1).Text / iFSatuanKecil) & " Where Inventory = '" & dcInventory.BoundText & "'"
  db.Execute "Update [C_ResepBahan] SET SubTotal=HargaPerUnit*JumlahKonsumsi"
  
  ' Update Harga Bahan Menu
  db.Execute "Update [C_MenuBahan] SET HargaPerUnit=" & CCur(txtFields(1).Text / iFSatuanKecil) & " Where Inventory = '" & dcInventory.BoundText & "'"
  db.Execute "Update [C_MenuBahan] SET SubTotal=HargaPerUnit*JumlahKonsumsi"
  '
  ' Update File Resep
  Dim sKodeResep As String
  Dim sngYield As Single
  Dim rsResep As New ADODB.Recordset
  rsResep.Open "Select Resep, JumlahFisik, BiayaBuat, PersenYield, BiayaYield, BiayaFisik, BiayaPerUnit From [C_Resep]", db, adOpenStatic, adLockOptimistic
  If rsResep.RecordCount <> 0 Then
    rsResep.MoveFirst
    '
    Do While Not rsResep.EOF
      sKodeResep = rsResep!Resep
      sngYield = rsResep!PersenYield
      '
      ' Cari Biaya Buat Baru
      Dim rsBiayaBuat As New ADODB.Recordset
      Dim curBiayaBuat, curBiayaYield, curBiayaFisik, curBiayaPerUnit As Currency
      rsBiayaBuat.Open "SELECT sum(SubTotal) AS JumlahBiaya From [C_ResepBahan] WHERE Resep='" & sKodeResep & "'", db, adOpenStatic, adLockReadOnly
      curBiayaBuat = rsBiayaBuat!JumlahBiaya
      rsBiayaBuat.Close
      Set rsBiayaBuat = Nothing
      '
      ' Update Biaya Fisik & Biaya Per Unit
      rsResep!BiayaBuat = curBiayaBuat
      curBiayaYield = (sngYield / 100) * curBiayaBuat
      rsResep!BiayaYield = curBiayaYield
      curBiayaFisik = curBiayaBuat + curBiayaYield
      rsResep!BiayaFisik = curBiayaFisik
      curBiayaPerUnit = curBiayaFisik / rsResep!JumlahFisik
      rsResep!BiayaPerUnit = curBiayaPerUnit
      rsResep.Update
      '
      db.Execute "Update [C_ResepBahan] SET HargaPerUnit=" & curBiayaPerUnit & " Where Inventory = '" & sKodeResep & "'"
      db.Execute "Update [C_ResepBahan] SET SubTotal=HargaPerUnit*JumlahKonsumsi"
      '
      ' Update Resep di Inventory
      Dim rsUpdateInventory As New ADODB.Recordset
      rsUpdateInventory.Open "Select Inventory, HargaPerUnit From [C_Inventory] Where Inventory='" & sKodeResep & "'", db, adOpenStatic, adLockOptimistic
      rsUpdateInventory!HargaPerUnit = curBiayaPerUnit
      rsUpdateInventory.Update
      rsUpdateInventory.Close
      Set rsUpdateInventory = Nothing
      '
      db.Execute "Update [C_MenuBahan] SET HargaPerUnit=" & curBiayaPerUnit & " Where Inventory = '" & sKodeResep & "'"
      db.Execute "Update [C_MenuBahan] SET SubTotal=HargaPerUnit*JumlahKonsumsi"
      '
      rsResep.MoveNext
    Loop
    rsResep.Close
    Set rsResep = Nothing
  End If
  '
  ' Update File Menu
  Dim sKodeMenu As String
  Dim sngFaktor, sngMYield As Single
  Dim curMBiayaYield, curMBiaya As Currency
  Dim rsMenu As New ADODB.Recordset
  rsMenu.Open "Select * From [C_Menu]", db, adOpenStatic, adLockOptimistic
  If rsMenu.RecordCount <> 0 Then
    rsMenu.MoveFirst
    '
    Do While Not rsMenu.EOF
      sKodeMenu = rsMenu!Menu
      sngFaktor = rsMenu!FaktorMarkUp
      sngMYield = rsMenu!PersenYield
      '
      ' Cari Biaya Fisik Baru
      Dim rsMBiayaBuat As New ADODB.Recordset
      Dim curMBiayaBuat As Currency
      rsMBiayaBuat.Open "SELECT sum(SubTotal) AS JumlahBiaya From [C_MenuBahan] WHERE Menu='" & sKodeMenu & "'", db, adOpenStatic, adLockReadOnly
      curMBiayaBuat = rsMBiayaBuat!JumlahBiaya
      rsMBiayaBuat.Close
      Set rsMBiayaBuat = Nothing
      '
      ' Update Biaya Fisik & Biaya Per Unit
      rsMenu!JumlahBiaya = curMBiayaBuat
      curMBiayaYield = (sngMYield / 100) * curMBiayaBuat
      rsMenu!BiayaYield = curMBiayaYield
      curMBiaya = curMBiayaBuat + curMBiayaYield
      rsMenu!Biaya = curMBiaya
      rsMenu!GrossMargin = rsMenu!HargaJual - curMBiaya
      If curMBiaya <> 0 Then
        rsMenu!FaktorMarkUp = (rsMenu!HargaJual - curMBiaya) / curMBiaya * 100
      End If
      rsMenu.Update
      '
      ' Update File Price Alert
      If (rsMenu!GrossMargin / rsMenu!HargaJual) * 100 < 50 Then
        '
        'Tambahkan Item ke File Price Alert
        If JumlahRecord("Select Menu from [C_PriceAlert] Where Menu = '" & sKodeMenu & "'", db) = 0 Then
          Dim rsAlert As New ADODB.Recordset
          rsAlert.Open "C_PriceAlert", db, adOpenStatic, adLockOptimistic
          rsAlert.AddNew
          rsAlert!Menu = rsMenu!Menu
          rsAlert!NamaMenu = rsMenu!NamaMenu
          rsAlert!Untung = rsMenu!GrossMargin
          rsAlert!Harga = rsMenu!HargaJual
          rsAlert!PersenUntungOfSales = (rsMenu!GrossMargin / rsMenu!HargaJual) * 100
          rsAlert.Update
          rsAlert.Close
          Set rsAlert = Nothing
        Else
          db.Execute "Update [C_PriceAlert] Set Untung=" & rsMenu!GrossMargin & ", Harga=" & rsMenu!HargaJual & ", PersenUntungOfSales=" & (rsMenu!GrossMargin / rsMenu!HargaJual) * 100 & " Where Menu='" & rsMenu!Menu & "'"
        End If
      Else
        '
        db.Execute "Delete From [C_PriceAlert] Where Menu='" & rsMenu!Menu & "'"
        '
      End If
      '
      rsMenu.MoveNext
    Loop
    rsMenu.Close
    Set rsMenu = Nothing
    '
  End If
  
  ' Update File Supplier - Inventory
  db.Execute "Update [C_Supplier_Inventory] Set LastPrice=" & CCur(txtFields(1).Text) & " Where Supplier='" & frmCIPInvoice.txtFields(4).Text & "' and Inventory='" & dcInventory.BoundText & "'"
  '
  frmCIPInvoice.txtFields(7).Text = CCur(frmCIPInvoice.txtFields(7).Text) + CCur(txtFields(2).Text)
  MsgBox "Barang telah masuk dalam daftar Invoice", vbInformation
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
    Dim rsHargaSatuan As New ADODB.Recordset
    rsHargaSatuan.Open "SELECT C_IPO.PO, C_IPO.Supplier, C_IPO.NamaSupplier, C_IStockCard.Inventory, C_IStockCard.NamaInventory, C_IStockCard.SisaPO, C_IStockCard.Harga FROM C_IPO INNER JOIN C_IStockCard ON C_IPO.PO = C_IStockCard.PO WHERE (((C_IPO.PO)='" & frmCIPInvoice.txtFields(2).Text & "')) AND ((C_IStockCard.Inventory)='" & dcInventory.BoundText & "')", db, adOpenStatic, adLockReadOnly
    txtFields(1).Text = rsHargaSatuan!Harga
    sSupplier = rsHargaSatuan!Supplier
    sNamaSupplier = rsHargaSatuan!NamaSupplier
    iJumlahPO = rsHargaSatuan!SisaPO
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
        If CSng(txtFields(0).Text) > iJumlahPO Then
          MsgBox "Jumlah yang dipesan hanya sebanyak " & iJumlahPO, vbCritical
          txtFields(0).Text = 0
          txtFields(0).SetFocus
          Exit Sub
        End If
        txtFields(2).Text = txtFields(0).Text * txtFields(1).Text
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

