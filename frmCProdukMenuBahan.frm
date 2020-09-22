VERSION 5.00
Begin VB.Form frmCProdukMenuBahan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bahan Menu"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
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
   ScaleHeight     =   3330
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1755
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   6735
      Begin VB.CommandButton cmdLookUp 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00E0E0E0&
         DataField       =   "NamaGudang"
         Height          =   315
         Index           =   1
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Inventory"
         Height          =   315
         Index           =   0
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         DataField       =   "Inventory"
         Height          =   315
         Index           =   2
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         DataField       =   "LastInvoicePrice"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bahan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Pemakaian"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblSatuan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3540
         TabIndex        =   9
         Top             =   720
         Width           =   510
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Konsumsi Bahan"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   8
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label lblHarga 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0.00 )"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Simpan"
      Height          =   795
      Left            =   2640
      Picture         =   "frmCProdukMenuBahan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&elesai"
      Height          =   795
      Left            =   3480
      Picture         =   "frmCProdukMenuBahan.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bahan Menu"
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
      TabIndex        =   12
      Top             =   180
      Width           =   1665
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   180
      Picture         =   "frmCProdukMenuBahan.frx":0614
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   495
      Left            =   60
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmCProdukMenuBahan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSatuan, sNamaSatuan As String
Dim curHarga As Currency

Private Sub ClearForm()
  For iCounter = 0 To 1
    txtFields(iCounter).Text = ""
  Next
  For iCounter = 2 To 3
    txtFields(iCounter).Text = 0
  Next
  lblSatuan.Caption = "Satuan"
  lblHarga.Caption = "( 0.00 per Satuan )"
End Sub

Private Sub Form_Activate()
  If frmCProdukMenu.Tag = "2" Then
    Dim rsBahan As New ADODB.Recordset
    rsBahan.Open "Select Inventory, NamaInventory, SatuanKecil, NamaSatuanKecil, HargaPerUnit from [C_Inventory] Where Inventory='" & txtFields(0).Text & "'", db, adOpenStatic, adLockReadOnly
    If rsBahan.RecordCount <> 0 Then
      txtFields(1).Text = rsBahan!NamaInventory
      sSatuan = rsBahan!SatuanKecil
      sNamaSatuan = rsBahan!NamaSatuanKecil
      curHarga = rsBahan!HargaPerUnit
      '
      lblSatuan.Caption = sNamaSatuan
      lblHarga.Caption = "( " & Format(curHarga, "###,##0.00") & " per " & sNamaSatuan & " ) "
    End If
  End If
End Sub

Private Sub Form_Load()
  sFormAktif = Me.Name
  If frmCProdukMenu.Tag = "1" Then
    ClearForm
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  sFormAktif = "FRMCPRODUKMENU"
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  '
  If Index = 2 Then
    Dim strValid As String
    '
    strValid = "0123456789."
    '
    If KeyAscii > 26 Then
      If InStr(strValid, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
      End If
    End If
  End If
  '
  If KeyAscii = 13 Then
    Select Case Index
      Case 0
        sKontrolAktif = "ARM_1"
        Dim rsInventory As New ADODB.Recordset
        rsInventory.Open "Select Inventory, NamaInventory, SatuanKecil, NamaSatuanKecil, HargaPerUnit from [C_Inventory] Where Inventory='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsInventory.RecordCount <> 0 Then
          txtFields(1).Text = rsInventory!NamaInventory
          sSatuan = rsInventory!SatuanKecil
          sNamaSatuan = rsInventory!NamaSatuanKecil
          curHarga = rsInventory!HargaPerUnit
          '
          lblSatuan.Caption = sNamaSatuan
          lblHarga.Caption = "( " & Format(curHarga, "###,##0.00") & " per " & sNamaSatuan & " ) "
          '
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(1).Text = ""
          frmCari.Show vbModal
          txtFields(0).SetFocus
        End If
        rsInventory.Close
        sKontrolAktif = ""
      Case 2
        If txtFields(2).Text = "" Then txtFields(2).Text = 0
        txtFields(3).Text = txtFields(2).Text * curHarga
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub cmdUpdate_Click()
  '
  Select Case frmCProdukMenu.Tag
    Case "1"
      '
      ' Cek Barang
      If txtFields(0).Text = "" Then
        MsgBox "Data Bahan Menu tidak boleh kosong", vbCritical
        txtFields(0).SetFocus
        Exit Sub
      End If
      iCounter = JumlahRecord("Select * From [C_MenuBahan] Where Inventory = '" & txtFields(0).Text & "' And Menu='" & frmCProdukMenu.txtFields(0).Text & "'", db)
      If iCounter > 0 Then
        MsgBox txtFields(1).Text & " telah ada dalam daftar bahan menu " & frmCProdukMenu.txtFields(1).Text, vbCritical
        ClearForm
        Exit Sub
      End If
      '
      Dim rsBahanMenu As New ADODB.Recordset
      rsBahanMenu.Open "C_MenuBahan", db, adOpenStatic, adLockOptimistic
      rsBahanMenu.AddNew
      rsBahanMenu!Menu = frmCProdukMenu.txtFields(0).Text
      rsBahanMenu!Inventory = txtFields(0).Text
      rsBahanMenu!NamaInventory = txtFields(1).Text
      rsBahanMenu!Satuan = sSatuan
      rsBahanMenu!NamaSatuan = sNamaSatuan
      rsBahanMenu!HargaPerUnit = curHarga
      rsBahanMenu!JumlahKonsumsi = txtFields(2).Text
      rsBahanMenu!SubTotal = txtFields(3).Text
      rsBahanMenu.Update
      rsBahanMenu.Close
      Set rsBahanMenu = Nothing
      '
      frmCProdukMenu.txtFields(8).Text = CCur(frmCProdukMenu.txtFields(8).Text) + CCur(txtFields(3).Text)
      frmCProdukMenu.txtFields(10).Text = CCur((frmCProdukMenu.txtFields(9).Text / 100) * frmCProdukMenu.txtFields(8).Text)
      frmCProdukMenu.txtFields(4).Text = CCur(frmCProdukMenu.txtFields(8).Text) + CCur(frmCProdukMenu.txtFields(10).Text)
      frmCProdukMenu.txtFields(7).Text = CCur(frmCProdukMenu.txtFields(6).Text) - CCur(frmCProdukMenu.txtFields(4).Text)
      If CCur(frmCProdukMenu.txtFields(4).Text) = 0 Then
        frmCProdukMenu.txtFields(5).Text = 0
      Else
        frmCProdukMenu.txtFields(5).Text = Format((CCur(frmCProdukMenu.txtFields(7).Text) / CCur(frmCProdukMenu.txtFields(4).Text)) * 100, "#0.00")
      End If
      MsgBox txtFields(1) & " telah masuk dalam daftar bahan menu " & frmCProdukMenu.txtFields(1).Text, vbInformation
      ClearForm
      '
      txtFields(0).SetFocus
      '
    Case "2"
      '
      ' Update Bahan Menu
      Dim rsUpdateBahan As New Recordset
      rsUpdateBahan.Open "Select * From [C_MenuBahan] Where Menu='" & frmCProdukMenu.txtFields(0).Text & "' and Inventory='" & txtFields(0).Text & "'", db, adOpenStatic, adLockOptimistic
      rsUpdateBahan!JumlahKonsumsi = txtFields(2).Text
      rsUpdateBahan!SubTotal = txtFields(3).Text
      rsUpdateBahan.Update
      rsUpdateBahan.Close
      Set rsUpdateBahan = Nothing
      '
      frmCProdukMenu.txtFields(8).Text = CCur(frmCProdukMenu.txtFields(8).Text) + CCur(txtFields(3).Text)
      frmCProdukMenu.txtFields(10).Text = CCur((frmCProdukMenu.txtFields(9).Text / 100) * frmCProdukMenu.txtFields(8).Text)
      frmCProdukMenu.txtFields(4).Text = CCur(frmCProdukMenu.txtFields(8).Text) + CCur(frmCProdukMenu.txtFields(10).Text)
      frmCProdukMenu.txtFields(7).Text = CCur(frmCProdukMenu.txtFields(6).Text) - CCur(frmCProdukMenu.txtFields(4).Text)
      If CCur(frmCProdukMenu.txtFields(4).Text) = 0 Then
        frmCProdukMenu.txtFields(5).Text = 0
      Else
        frmCProdukMenu.txtFields(5).Text = Format((CCur(frmCProdukMenu.txtFields(7).Text) / CCur(frmCProdukMenu.txtFields(4).Text)) * 100, "#0.00")
      End If
      MsgBox txtFields(1) & " telah diupdate dalam daftar bahan menu " & frmCProdukMenu.txtFields(1).Text, vbInformation
      '
      txtFields(0).Enabled = True
      frmCProdukMenu.Tag = ""
      Unload Me
      frmCProdukMenu.cmdUpdate.Value = True
  End Select
  '
End Sub

Private Sub cmdClose_Click()
  frmCProdukMenu.txtFields(8).Text = CCur(frmCProdukMenu.txtFields(8).Text) + CCur(txtFields(3).Text)
  frmCProdukMenu.txtFields(10).Text = CCur((frmCProdukMenu.txtFields(9).Text / 100) * frmCProdukMenu.txtFields(8).Text)
  frmCProdukMenu.txtFields(4).Text = CCur(frmCProdukMenu.txtFields(8).Text) + CCur(frmCProdukMenu.txtFields(10).Text)
  frmCProdukMenu.txtFields(7).Text = CCur(frmCProdukMenu.txtFields(6).Text) - CCur(frmCProdukMenu.txtFields(4).Text)
  If CCur(frmCProdukMenu.txtFields(4).Text) = 0 Then
    frmCProdukMenu.txtFields(5).Text = 0
  Else
    frmCProdukMenu.txtFields(5).Text = Format((CCur(frmCProdukMenu.txtFields(7).Text) / CCur(frmCProdukMenu.txtFields(4).Text)) * 100, "#0.00")
  End If
  '
  txtFields(0).Enabled = True
  frmCProdukMenu.Tag = ""
  Unload Me
End Sub

Private Sub cmdLookUp_Click(Index As Integer)
  Select Case Index
    Case 0
      sKontrolAktif = "ARM_1"
      frmCari.Show vbModal
      txtFields(0).SetFocus
      If txtFields(0).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub
