VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHistory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documents History"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
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
   ScaleHeight     =   6090
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5115
      Left            =   3480
      TabIndex        =   18
      Top             =   60
      Width           =   7875
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   4395
         Left            =   60
         ScaleHeight     =   4395
         ScaleWidth      =   7755
         TabIndex        =   19
         Top             =   660
         Width           =   7755
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "Tanggal"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   5340
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "KodeRef"
            Height          =   315
            Index           =   0
            Left            =   2880
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "1234567890"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataField       =   "Nilai"
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
            Left            =   5940
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Keterangan"
            Height          =   315
            Index           =   4
            Left            =   2880
            Locked          =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   540
            Width           =   3675
         End
         Begin MSDataGridLib.DataGrid dgPI 
            Bindings        =   "frmHistory.frx":0000
            Height          =   2895
            Left            =   60
            TabIndex        =   5
            Top             =   1020
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   5106
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detail Informasi"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "Inventory"
               Caption         =   "Kode"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "NamaInventory"
               Caption         =   "Nama Inventory"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Jumlah"
               Caption         =   "Jumlah"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "Satuan"
               Caption         =   "Satuan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "HargaSatuan"
               Caption         =   "Harga Satuan"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "SubTotal"
               Caption         =   "Sub Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   2
               AllowSizing     =   0   'False
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1934.929
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   824.882
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   870.236
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1440
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4560
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Referensi"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1140
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Transaksi"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   4320
            TabIndex        =   23
            Top             =   4020
            Width           =   1410
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   1140
            TabIndex        =   22
            Top             =   600
            Width           =   840
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documents History"
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
         Left            =   900
         TabIndex        =   26
         Top             =   180
         Width           =   2445
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   240
         Picture         =   "frmHistory.frx":0015
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
         Width           =   7755
      End
   End
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   3375
      Begin VB.CommandButton cmdCetakDoc 
         Caption         =   "&Cetak Dokumen"
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
         Left            =   0
         Picture         =   "frmHistory.frx":08DF
         TabIndex        =   2
         Top             =   5580
         Width           =   3315
      End
      Begin TabDlg.SSTab TabIndek 
         Height          =   5535
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   9763
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
         TabPicture(0)   =   "frmHistory.frx":0BE9
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtIndeks"
         Tab(0).Control(1)=   "lstIndeks"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmHistory.frx":0C05
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblLabels(11)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "txtCari"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cboBy"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdTampil"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lstCari"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   -74760
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   3375
            ItemData        =   "frmHistory.frx":0C21
            Left            =   240
            List            =   "frmHistory.frx":0C28
            TabIndex        =   14
            Top             =   1860
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
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmHistory.frx":0C35
            Left            =   240
            List            =   "frmHistory.frx":0C54
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   2835
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ListBox lstIndeks 
            Height          =   4350
            Left            =   -74760
            TabIndex        =   1
            Top             =   900
            Width           =   2835
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Dokumen"
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
            TabIndex        =   17
            Top             =   480
            Width           =   1290
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   120
      TabIndex        =   27
      Top             =   5220
      Width           =   11235
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmHistory.frx":0D19
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   5940
         Picture         =   "frmHistory.frx":1023
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3420
         Picture         =   "frmHistory.frx":132D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4260
         Picture         =   "frmHistory.frx":1637
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5100
         Picture         =   "frmHistory.frx":1941
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmHistory"
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

Private Sub ProsesDetail()
  '
  Dim strSQL As String
  Dim strKey As String
  '
  If txtFields(0).Text = "" Then
    strKey = "0"
  Else
    strKey = Trim(txtFields(0).Text)
  End If
  '
  Set rsDetail = New ADODB.Recordset
  strSQL = "SELECT * FROM [HistoryDetail] WHERE KodeRef='" & strKey & "' Order By Inventory"
  rsDetail.Open strSQL, db, adOpenStatic, adLockReadOnly
  rsDetail.Requery
  '
  Set dgPI.DataSource = rsDetail
  dgPI.ReBind
  '
End Sub

Private Sub Form_Load()
  '
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "Select * from [History] Order by Tanggal", db, adOpenStatic, adLockReadOnly
  
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  '
  ProsesDetail
  '
  RefreshIndeks
  lstCari.Clear
  TabIndek.Tab = 0
  '
  mbDataChanged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
  Screen.MousePointer = vbDefault
  '
  adoPrimaryRS.Close
  rsDetail.Close
  Set adoPrimaryRS = Nothing
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
  '
  ProsesDetail
End Sub

Private Sub cmdLast_Click()
  On Error Resume Next
  '
  adoPrimaryRS.MoveLast
  mbDataChanged = False
  '
  ProsesDetail
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
  '
  ProsesDetail
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
  '
  ProsesDetail
End Sub

Private Sub SetButtons(bVal As Boolean)
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
  rsIndeks.Open "Select KodeRef FROM [History]", db, adOpenStatic, adLockOptimistic
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
    lstIndeks.AddItem rsIndeks!KodeRef
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
    lstCari.AddItem rsCari!KodeRef
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select KodeRef FROM [History] WHERE KodeRef LIKE '%" & txtIndeks.Text & "%'", db, adOpenStatic, adLockReadOnly
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
  adoPrimaryRS.Find "KodeRef ='" & lstIndeks.Text & "'"
  ProsesDetail
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
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 1
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='PO' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 2
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='PI' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 3
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='IA+' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 4
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='IA-' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 5
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='WASTE' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 6
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='PR' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 7
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='AI' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
    Case 8
      rsCari.Open "Select * FROM [History] WHERE KodeRef LIKE '%" & txtCari.Text & "%' and Jenis='KI' Order By Tanggal, KodeRef", db, adOpenStatic, adLockReadOnly
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
  adoPrimaryRS.Find "KodeRef ='" & lstCari.Text & "'"
  ProsesDetail
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

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  SendKeys "{TAB}"
End Sub

Private Sub cmdCetakDoc_Click()
  '
  Dim sSQL As String
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Tidak ada data yang tercetak", vbInformation
    Exit Sub
  End If
  '
  sSQL = "SELECT History.KodeRef, History.Tanggal, History.Keterangan, HistoryDetail.Inventory, HistoryDetail.NamaInventory, HistoryDetail.Jumlah, HistoryDetail.Satuan, HistoryDetail.HargaSatuan, HistoryDetail.SubTotal FROM [History] INNER JOIN [HistoryDetail] ON History.KodeRef = HistoryDetail.KodeRef WHERE (((History.KodeRef)='" & txtFields(0).Text & "'))"

  With frmUtama.datReport
    .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
    .RecordSource = sSQL
    .Refresh
  End With
  With frmUtama.crReport
    .Formulas(0) = ""
    .Formulas(1) = ""
    .Formulas(2) = ""
    .Formulas(3) = ""
    .Formulas(4) = ""
    .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
    .WindowTitle = "Dokumen Transaksi"
    .ReportFileName = App.Path & "\Report\Laporan History.Rpt"
    .Formulas(0) = "IDCompany= '" & ICNama & "'"
    .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
    .Formulas(2) = "IDKota= '" & ICKota & "'"
    .Action = 1
  End With
  '
End Sub
