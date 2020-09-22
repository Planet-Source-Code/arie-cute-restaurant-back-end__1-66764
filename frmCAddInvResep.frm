VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCAddInvResep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penambahan Inventory Resep"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
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
   ScaleHeight     =   4185
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   3435
      Begin TabDlg.SSTab TabIndek 
         Height          =   4035
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   7117
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Indeks"
         TabPicture(0)   =   "frmCAddInvResep.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtIndeks"
         Tab(0).Control(1)=   "lstIndeks"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCAddInvResep.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblLabels(11)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lstCari"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdTampil"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cboBy"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtCari"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCAddInvResep.frx":0038
            Left            =   240
            List            =   "frmCAddInvResep.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   1815
            ItemData        =   "frmCAddInvResep.frx":0080
            Left            =   240
            List            =   "frmCAddInvResep.frx":0087
            TabIndex        =   14
            Top             =   1860
            Width           =   2835
         End
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   -74760
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.ListBox lstIndeks 
            Height          =   2985
            Left            =   -74760
            TabIndex        =   1
            Top             =   840
            Width           =   2835
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berdasarkan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   32
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3195
      Left            =   3480
      TabIndex        =   15
      Top             =   60
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   2475
         Left            =   60
         ScaleHeight     =   2475
         ScaleWidth      =   7695
         TabIndex        =   16
         Top             =   660
         Width           =   7695
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "NilaiBuatAcc"
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
            Index           =   7
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   1740
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "KodeBuat"
            Height          =   315
            Index           =   0
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "1234567890"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            DataField       =   "Tanggal"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "d-M-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   5700
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "Karyawan"
            Height          =   315
            Index           =   2
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   660
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaKaryawan"
            Height          =   315
            Index           =   3
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   660
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaResep"
            Height          =   315
            Index           =   5
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1020
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "Resep"
            Height          =   315
            Index           =   4
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataField       =   "JumlahAcc"
            Height          =   315
            Index           =   6
            Left            =   3000
            TabIndex        =   8
            Top             =   1380
            Width           =   675
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Penambahan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   1140
            TabIndex        =   35
            Top             =   1800
            Width           =   1665
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Pembuatan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Buat"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4560
            TabIndex        =   27
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1140
            TabIndex        =   26
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Resep"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1140
            TabIndex        =   25
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblSatuan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            DataField       =   "Satuan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3840
            TabIndex        =   24
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah penambahan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1140
            TabIndex        =   23
            Top             =   1440
            Width           =   1470
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penambahan Inventory Resep"
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
         Left            =   780
         TabIndex        =   29
         Top             =   180
         Width           =   3825
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCAddInvResep.frx":0094
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   60
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   60
      TabIndex        =   33
      Top             =   3300
      Width           =   11235
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCAddInvResep.frx":095E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCAddInvResep.frx":0C68
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCAddInvResep.frx":0F72
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCAddInvResep.frx":127C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCAddInvResep.frx":1586
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Proses"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCAddInvResep.frx":1890
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCAddInvResep.frx":1B9A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Setuju"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCAddInvResep.frx":1EA4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCAddInvResep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim rsDetail As ADODB.Recordset
Dim rsIndeks As ADODB.Recordset
Dim rsCari As ADODB.Recordset

Dim PosisiRecord As Long

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub StatusFrame(bolStatus As Boolean)
  '
  Picture1.Enabled = bolStatus
  fraFind.Enabled = Not bolStatus
  '
End Sub

Private Sub Form_Load()
  '
  sFormAktif = Me.Name
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "Select * from [C_BuatResepSpesial] Order by KodeBuat", db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set lblSatuan.DataSource = adoPrimaryRS
  '
  RefreshIndeks
  lstCari.Clear
  TabIndek.Tab = 0
  '
  StatusFrame False
  '
  mbDataChanged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
  Screen.MousePointer = vbDefault
  '
  adoPrimaryRS.Close
  Set adoPrimaryRS = Nothing
  '
End Sub

Private Sub cmdAdd_Click()
  'On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Permintaan Pembuatan Resep yang telah disetujui kosong", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS!NeedTransfer = False Then
    MsgBox "Data Penambahan Inventory Resep telah diproses", vbInformation
    Exit Sub
  End If
  
  mvBookMark = adoPrimaryRS.Bookmark
  StatusFrame True
  txtFields(6).Text = adoPrimaryRS!Jumlah
  '
  db.BeginTrans
  mbEditFlag = True
  SetButtons False
  '
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS!JumlahAcc = txtFields(6).Text
  adoPrimaryRS!NeedTransfer = False
  adoPrimaryRS.UpdateBatch adAffectAll
  '
  ' Update StockCard Resep
  Dim rsStockCardResep As New ADODB.Recordset
  rsStockCardResep.Open "C_IStockCard", db, adOpenStatic, adLockOptimistic
  rsStockCardResep.AddNew
  rsStockCardResep!Tanggal = txtFields(1).Text
  rsStockCardResep!Inventory = txtFields(4).Text
  rsStockCardResep!NamaInventory = txtFields(5).Text
  rsStockCardResep!PI = txtFields(0).Text
  rsStockCardResep!JumlahPI = txtFields(6).Text
  rsStockCardResep!SatuanBesar = lblSatuan.Caption
  rsStockCardResep!Keterangan = "Penambahan Inventory Resep"
  If txtFields(6).Text = 0 Then
    rsStockCardResep!UnitMasuk = 0
    rsStockCardResep!Harga = 0
  Else
    rsStockCardResep!Harga = CCur(txtFields(7).Text / txtFields(6).Text)
    rsStockCardResep!UnitMasuk = txtFields(6).Text
  End If
  rsStockCardResep!HargaSubTotal = txtFields(7).Text
  rsStockCardResep.Update
  rsStockCardResep.Close
  Set rsStockCardResep = Nothing
  
  ' Update File Inventory Resep
  Dim rsItemInventory As New ADODB.Recordset
  Dim StockMin, QtyOnHandKecilBaru As Single
  rsItemInventory.Open "C_Inventory", db, adOpenStatic, adLockOptimistic
  rsItemInventory.Find "Inventory='" & txtFields(4).Text & "'"
  '
  StockMin = rsItemInventory!ReorderLevel
  QtyOnHandKecilBaru = rsItemInventory!QtyOnHandKecil + txtFields(6).Text
  '
  rsItemInventory!QtyOnHandKecil = QtyOnHandKecilBaru
  rsItemInventory!JumlahItem = rsItemInventory!JumlahItem + txtFields(6).Text
  rsItemInventory!LastInvoiceDate = txtFields(1).Text
  rsItemInventory.Update
  rsItemInventory.Close
  Set rsItemInventory = Nothing
  '
  ' Update Stock Alert
  If QtyOnHandKecilBaru > StockMin Then
    ' Hapus Item di File Stock Alert
    db.Execute "DELETE From [C_StockAlert] Where Inventory='" & txtFields(4).Text & "'"
  Else
    db.Execute "UPDATE [C_StockAlert] SET QtyOnHand =" & QtyOnHandKecilBaru & " Where Inventory='" & txtFields(4).Text & "'"
  End If
    
  db.CommitTrans

  sPesan = "Pembuatan Resep dengan Kode " & txtFields(0).Text & vbCrLf
  sPesan = sPesan & "telah ditambahkan dalam Inventory"
  MsgBox sPesan, vbInformation
  
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  RefreshIndeks
  StatusFrame False
  Exit Sub
  '
UpdateErr:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  adoPrimaryRS.Requery
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  '
  db.RollbackTrans
  '
  RefreshIndeks
  StatusFrame False
  mbDataChanged = False
  '
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error Resume Next
  '
  adoPrimaryRS.MoveFirst
  mbDataChanged = False
End Sub

Private Sub cmdLast_Click()
  On Error Resume Next
  '
  adoPrimaryRS.MoveLast
  mbDataChanged = False
End Sub

Private Sub cmdNext_Click()
  On Error Resume Next
  '
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False
End Sub

Private Sub cmdPrevious_Click()
  On Error Resume Next

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdClose.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub RefreshIndeks()
  '
  PosisiRecord = adoPrimaryRS.AbsolutePosition
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select KodeBuat FROM [C_BuatResepSpesial]", db, adOpenStatic, adLockOptimistic
  Indeks
  '
End Sub

Private Sub Indeks()
  '
  On Error Resume Next
  '
  lstIndeks.Clear
  Me.MousePointer = vbHourglass
  '
  Do While Not rsIndeks.EOF
    lstIndeks.AddItem rsIndeks!KodeBuat
    rsIndeks.MoveNext
  Loop
  
  If PosisiRecord = 0 Then PosisiRecord = 1
  lstIndeks.Selected(PosisiRecord - 1) = True
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub Cari()
  '
  On Error Resume Next
  '
  lstCari.Clear
  Me.MousePointer = vbHourglass
  '
  Do While Not rsCari.EOF
    lstCari.AddItem rsCari!KodeBuat
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select KodeBuat FROM [C_BuatResepSpesial] WHERE KodeBuat LIKE '%" & txtIndeks.Text & "%'", db, adOpenStatic, adLockOptimistic
  Indeks
  '
End Sub

Private Sub lstIndeks_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    lstIndeks_Click
  End If
  '
End Sub

Private Sub lstIndeks_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
      lstIndeks_Click
  End Select
End Sub

Private Sub lstIndeks_Click()
  On Error GoTo Handler
  '
  adoPrimaryRS.MoveFirst
  adoPrimaryRS.Find "KodeBuat ='" & lstIndeks.Text & "'"
  Exit Sub
  
Handler:
  MsgBox Err.Description
End Sub

Private Sub cmdTampil_Click()
  '
  If cboBy.ListIndex = -1 Then
    MsgBox "Tentukan kriteria pencarian terlebih dahulu", vbInformation
    cboBy.SetFocus
    Exit Sub
  End If
  '
  Set rsCari = New ADODB.Recordset
  Select Case cboBy.ListIndex
    Case 0
      rsCari.Open "Select * FROM [C_BuatResepSpesial] WHERE KodeBuat LIKE '%" & txtCari.Text & "%' and NeedTransfer=True Order By KodeBuat", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_BuatResepSpesial] WHERE KodeBuat LIKE '%" & txtCari.Text & "%' and NeedTransfer=false Order By KodeBuat", db, adOpenStatic, adLockOptimistic
  End Select
  '
  Cari
  rsCari.Close
  '
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    cmdTampil_Click
    lstCari.SetFocus
  End If
  '
End Sub

Private Sub cboBy_Change()
  '
  txtCari.Text = ""
  txtCari.SetFocus
  '
End Sub

Private Sub cboBy_KeyPress(KeyAscii As Integer)
  SendKeys "{TAB}"
End Sub

Private Sub lstCari_Click()
  '
  On Error GoTo Handler
  '
  adoPrimaryRS.MoveFirst
  adoPrimaryRS.Find "KodeBuat ='" & lstCari.Text & "'"
  Exit Sub

Handler:
  MsgBox Err.Description
  '
End Sub

Private Sub lstCari_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
      lstCari_Click
  End Select
End Sub
