VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCKembali 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pengembalian Item Inventory"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
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
   ScaleHeight     =   6135
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6195
      Left            =   60
      TabIndex        =   29
      Top             =   60
      Width           =   3375
      Begin VB.CommandButton cmdCetakKI 
         Caption         =   "&Cetak Form"
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
         Picture         =   "frmCKembali.frx":0000
         TabIndex        =   2
         Top             =   5640
         Width           =   3315
      End
      Begin TabDlg.SSTab TabIndek 
         Height          =   5595
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   9869
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
         TabPicture(0)   =   "frmCKembali.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtIndeks"
         Tab(0).Control(1)=   "lstIndeks"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCKembali.frx":0326
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
            TabIndex        =   15
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCKembali.frx":0342
            Left            =   240
            List            =   "frmCKembali.frx":0349
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   720
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   3570
            ItemData        =   "frmCKembali.frx":0360
            Left            =   240
            List            =   "frmCKembali.frx":0367
            TabIndex        =   17
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
            Height          =   4545
            Left            =   -74760
            TabIndex        =   1
            Top             =   900
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
            TabIndex        =   31
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   3480
      TabIndex        =   18
      Top             =   60
      Width           =   7935
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   60
         ScaleHeight     =   4455
         ScaleWidth      =   7815
         TabIndex        =   19
         Top             =   660
         Width           =   7815
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   4260
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "KodeBuat"
            Height          =   315
            Index           =   2
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "1234567890"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "NilaiKembali"
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
            Left            =   6060
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "KodeKembali"
            Height          =   315
            Index           =   0
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "1234567890"
            Top             =   240
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
            Left            =   5460
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBarang 
            Caption         =   "&Item Inventory"
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
            Left            =   2640
            Picture         =   "frmCKembali.frx":0374
            TabIndex        =   11
            Top             =   3840
            Width           =   2595
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaKaryawan"
            Height          =   315
            Index           =   4
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   960
            Width           =   2955
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "Karyawan"
            Height          =   315
            Index           =   3
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   960
            Width           =   675
         End
         Begin MSDataGridLib.DataGrid dgPI 
            Bindings        =   "frmCKembali.frx":067E
            Height          =   1875
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3307
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
            Caption         =   "Item Inventory yang dikembalikan"
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
               DataField       =   "UnitMasuk"
               Caption         =   "#Kembali"
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
               DataField       =   "Harga"
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
               DataField       =   "HargaSubTotal"
               Caption         =   "SubTotal"
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
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1920.189
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1319.811
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   1409.953
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Permintaan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1260
            TabIndex        =   33
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Pengembalian"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   4200
            TabIndex        =   27
            Top             =   3420
            Width           =   1725
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Pengembalian"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1260
            TabIndex        =   26
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4680
            TabIndex        =   25
            Top             =   300
            Width           =   615
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   24
            Top             =   1020
            Width           =   1170
         End
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCKembali.frx":0693
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pengembalian Item Inventory"
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
         Left            =   840
         TabIndex        =   28
         Top             =   180
         Width           =   3780
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   60
         Top             =   120
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   60
      TabIndex        =   32
      Top             =   5280
      Width           =   11355
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCKembali.frx":0F5D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCKembali.frx":1267
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCKembali.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCKembali.frx":187B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10500
         Picture         =   "frmCKembali.frx":1B85
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10500
         Picture         =   "frmCKembali.frx":1E8F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   9660
         Picture         =   "frmCKembali.frx":2199
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   9660
         Picture         =   "frmCKembali.frx":24A3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCKembali"
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

Public bDone As Boolean

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
  strSQL = "SELECT * FROM [C_AmbilKembaliDetail] WHERE PI='" & strKey & "' Order By Inventory"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgPI.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgPI.Caption = "Inventory yang dikembalikan berdasarkan kode " & adoPrimaryRS!KodeKembali
  End If
  dgPI.ReBind
  '
End Sub

Private Sub StatusFrame(bolStatus As Boolean)
  '
  Picture1.Enabled = bolStatus
  fraFind.Enabled = Not bolStatus
  '
End Sub

Private Sub Form_Load()
  '
  sFormAktif = Me.Name
  bDone = False
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "select * from [C_Kembali] Order by KodeKembali", db, adOpenStatic, adLockOptimistic
  
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
  StatusFrame False
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
  sFormAktif = ""
  '
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  '
  StatusFrame True
  bDone = False
  '
  db.BeginTrans
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    mbAddNewFlag = True
    SetButtons False
  End With
  '
  ProsesDetail
  '
  txtFields(7).Text = 0
  txtFields(1).Text = Date
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  ' Validasi Field
  If txtFields(0) = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(0).SetFocus
    Exit Sub
  End If
  If txtFields(2) = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(2).SetFocus
    Exit Sub
  End If
  
  If rsDetail.RecordCount = 0 Then
    MsgBox "Daftar harus berisi sedikitnya satu jenis inventory", vbCritical
    cmdBarang.SetFocus
    Exit Sub
  End If
  '
  db.Execute "Update [C_Minta] Set NeedReturn = False Where KodeBuat='" & txtFields(2).Text & "'"
  adoPrimaryRS!NeedCheck = True
  adoPrimaryRS.UpdateBatch adAffectAll
  '
  db.CommitTrans

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  ProsesDetail
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
  ProsesDetail
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
  rsIndeks.Open "Select KodeKembali FROM [C_Kembali]", db, adOpenStatic, adLockOptimistic
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
    lstIndeks.AddItem rsIndeks!KodeKembali
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
    lstCari.AddItem rsCari!KodeKembali
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select KodeKembali FROM [C_Kembali] WHERE KodeKembali LIKE '%" & txtIndeks.Text & "%'", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "KodeKembali ='" & lstIndeks.Text & "'"
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
      rsCari.Open "Select * FROM [C_Kembali] WHERE KodeKembali LIKE '%" & txtCari.Text & "%' Order By KodeKembali", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "KodeKembali ='" & lstCari.Text & "'"
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
  '
  If KeyAscii = 13 Then
    Select Case Index
      Case 0
        If JumlahRecord("Select KodeKembali From [C_Kembali] Where KodeKembali='" & txtFields(0).Text & "'", db) <> 0 Then
          MsgBox "Kode Dokumen " & txtFields(0).Text & " telah digunakan oleh dokumen lain", vbCritical
          txtFields(0).Text = ""
          txtFields(0).SetFocus
          Exit Sub
        End If
        If JumlahRecord("Select KodeRef From [History] Where KodeRef='" & txtFields(0).Text & "'", db) <> 0 Then
          MsgBox "Kode Dokumen " & txtFields(0).Text & " telah digunakan oleh dokumen lain", vbCritical
          txtFields(0).Text = ""
          txtFields(0).SetFocus
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case 2
        sKontrolAktif = "AKI_1"
        Dim rsMinta As New ADODB.Recordset
        rsMinta.Open "Select KodeBuat, Karyawan, NamaKaryawan, Selesai from [C_Minta] Where KodeBuat='" & txtFields(2).Text & "' and NeedReturn=True", db, adOpenStatic, adLockReadOnly
        If rsMinta.RecordCount <> 0 Then
          txtFields(3).Text = rsMinta!Karyawan
          txtFields(4).Text = rsMinta!NamaKaryawan
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(3).Text = ""
          txtFields(4).Text = ""
          frmCari.Show vbModal
          txtFields(2).SetFocus
        End If
        rsMinta.Close
        sKontrolAktif = ""
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
  '
End Sub

Private Sub cmdBarang_Click()
  '
  If (frmCKembali.txtFields(0).Text = "") Or (frmCKembali.txtFields(2).Text = "") Then
    MsgBox "Nomor Kode Pengembalian & Permintaan tidak boleh kosong", vbCritical
    txtFields(0).SetFocus
    Exit Sub
  End If
  '
  frmCKembaliDetail.Caption = "No. Pengembalian : " & adoPrimaryRS!KodeKembali & " - Inventory yang akan dikembalikan"
  frmCKembaliDetail.Show vbModal
  ProsesDetail
  '
End Sub

Private Sub cmdCetakKI_Click()
  '
  Dim sSQL As String
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Tidak ada data yang tercetak", vbInformation
    Exit Sub
  End If
  '
  sSQL = "SELECT C_Kembali.KodeKembali, C_Kembali.Tanggal, C_Kembali.NamaKaryawan, C_AmbilKembaliDetail.Inventory, C_AmbilKembaliDetail.NamaInventory, C_AmbilKembaliDetail.UnitMasuk, C_AmbilKembaliDetail.Satuan, C_AmbilKembaliDetail.Harga, C_AmbilKembaliDetail.HargaSubTotal FROM C_Kembali INNER JOIN C_AmbilKembaliDetail ON C_Kembali.KodeKembali = C_AmbilKembaliDetail.PI WHERE (((C_Kembali.KodeKembali)='" & txtFields(0).Text & "'))"

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
    .WindowTitle = "Form Pengembalian Inventory"
    .ReportFileName = App.Path & "\Report\Laporan Pengembalian Inventory.Rpt"
    .Formulas(0) = "IDCompany= '" & ICNama & "'"
    .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
    .Formulas(2) = "IDKota= '" & ICKota & "'"
    .Action = 1
  End With
  '
End Sub

Private Sub cmdLookUp_Click(Index As Integer)
  Select Case Index
    Case 0
      sKontrolAktif = "AKI_1"
      frmCari.Show vbModal
      txtFields(2).SetFocus
      If txtFields(2).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub
