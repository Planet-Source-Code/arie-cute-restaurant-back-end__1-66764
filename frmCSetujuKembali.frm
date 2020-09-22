VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCSetujuKembali 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pemeriksaan Pengembalian Inventory"
   ClientHeight    =   6270
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
   ScaleHeight     =   6270
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6195
      Left            =   60
      TabIndex        =   30
      Top             =   60
      Width           =   3375
      Begin VB.CommandButton cmdCetakBR 
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
         Picture         =   "frmCSetujuKembali.frx":0000
         TabIndex        =   2
         Top             =   5760
         Width           =   3315
      End
      Begin TabDlg.SSTab TabIndek 
         Height          =   5715
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   10081
         _Version        =   393216
         Tabs            =   2
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
         TabPicture(0)   =   "frmCSetujuKembali.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCSetujuKembali.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtCari"
         Tab(1).Control(1)=   "cboBy"
         Tab(1).Control(2)=   "cmdTampil"
         Tab(1).Control(3)=   "lstCari"
         Tab(1).Control(4)=   "lblLabels(11)"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   -74760
            TabIndex        =   12
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCSetujuKembali.frx":0342
            Left            =   -74760
            List            =   "frmCSetujuKembali.frx":034C
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   13
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   3570
            ItemData        =   "frmCSetujuKembali.frx":0392
            Left            =   -74760
            List            =   "frmCSetujuKembali.frx":0399
            TabIndex        =   14
            Top             =   1860
            Width           =   2835
         End
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.ListBox lstIndeks 
            Height          =   4545
            Left            =   240
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
            Left            =   -74760
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
      Height          =   5295
      Left            =   3480
      TabIndex        =   15
      Top             =   60
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   60
         ScaleHeight     =   4575
         ScaleWidth      =   7695
         TabIndex        =   16
         Top             =   660
         Width           =   7695
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "KodeBuat"
            Height          =   315
            Index           =   2
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "1234567890"
            Top             =   660
            Width           =   1095
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
            Index           =   8
            Left            =   5940
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3660
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "NilaiAcc"
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
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   4020
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "Karyawan"
            Height          =   315
            Index           =   3
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "123456"
            Top             =   1020
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaKaryawan"
            Height          =   315
            Index           =   4
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1020
            Width           =   2535
         End
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
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
            Left            =   5280
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "KodeKembali"
            Height          =   315
            Index           =   0
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "1234567890"
            Top             =   300
            Width           =   1095
         End
         Begin MSDataGridLib.DataGrid dgPI 
            Height          =   1875
            Left            =   60
            TabIndex        =   23
            Top             =   1680
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3307
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
            Caption         =   "Item Inventory yg dikembalikan"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "NamaInventory"
               Caption         =   "Nama Item Inventory"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Harga"
               Caption         =   "Harga"
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
            BeginProperty Column02 
               DataField       =   "UnitMasuk"
               Caption         =   "# Kembali"
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
               DataField       =   "UnitMasukAcc"
               Caption         =   "# Periksa"
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
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "HargaSubTotalAcc"
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
               MarqueeStyle    =   2
               ScrollBars      =   2
               AllowSizing     =   0   'False
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   1995.024
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   929.764
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   929.764
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1425.26
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
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Pengembalian"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   4020
            TabIndex        =   28
            Top             =   3720
            Width           =   1725
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Pemeriksaan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   4020
            TabIndex        =   27
            Top             =   4080
            Width           =   1635
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
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Buat"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4200
            TabIndex        =   25
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
            Left            =   1260
            TabIndex        =   24
            Top             =   1080
            Width           =   1170
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemeriksaan Pengembalian Inventory"
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
         Width           =   4890
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCSetujuKembali.frx":03A6
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
      Top             =   5400
      Width           =   11235
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Setuju"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCSetujuKembali.frx":0C70
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCSetujuKembali.frx":0F7A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCSetujuKembali.frx":1284
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCSetujuKembali.frx":158E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCSetujuKembali.frx":1898
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCSetujuKembali.frx":1BA2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "P&roses"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCSetujuKembali.frx":1EAC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCSetujuKembali.frx":21B6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCSetujuKembali"
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
  strSQL = "SELECT * FROM [C_AmbilKembaliDetail] WHERE PI='" & strKey & "' Order By Inventory"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgPI.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgPI.Caption = "Inventory yang dikembalikan berdasarkan kode " & adoPrimaryRS!KodeBuat
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
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "Select * from [C_Kembali] Order by KodeBuat", db, adOpenStatic, adLockOptimistic
  
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
  '
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Pengembalian Inventory kosong", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS!NeedCheck = False Then
    MsgBox "Data Pengembalian Inventory telah diproses", vbInformation
    Exit Sub
  End If
  
  mvBookMark = adoPrimaryRS.Bookmark
  StatusFrame True
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

  adoPrimaryRS!NeedCheck = False
  adoPrimaryRS.UpdateBatch adAffectAll
  '
  ' Isi File History
  Dim rsHistory As New ADODB.Recordset
  rsHistory.Open "History", db, adOpenStatic, adLockOptimistic
  rsHistory.AddNew
  rsHistory!KodeRef = txtFields(0).Text
  rsHistory!Tanggal = txtFields(1).Text
  rsHistory!Keterangan = "Pengembalian Inventory oleh " & txtFields(4).Text
  rsHistory!Nilai = txtFields(7).Text
  rsHistory!Jenis = "KI"
  rsHistory.Update
  rsHistory.Close
  Set rsHistory = Nothing
  
  ' Isi File Detail History
  Dim rsHistoryDetail As New ADODB.Recordset
  rsHistoryDetail.Open "HistoryDetail", db, adOpenStatic, adLockOptimistic
  rsDetail.MoveFirst
  Do While Not rsDetail.EOF
    rsHistoryDetail.AddNew
    rsHistoryDetail!KodeRef = txtFields(0).Text
    rsHistoryDetail!Inventory = rsDetail!Inventory
    rsHistoryDetail!NamaInventory = rsDetail!NamaInventory
    rsHistoryDetail!Jumlah = rsDetail!UnitMasukAcc
    rsHistoryDetail!Satuan = rsDetail!Satuan
    rsHistoryDetail!HargaSatuan = rsDetail!Harga
    rsHistoryDetail!SubTotal = rsDetail!HargaSubTotalAcc
    rsHistoryDetail.Update
    '
    rsDetail.MoveNext
  Loop
  rsHistoryDetail.Close
  Set rsHistoryDetail = Nothing
  
  ' Update File Inventory Bahan
  Dim NoInventory As String
  rsDetail.MoveFirst
  Do While Not rsDetail.EOF
    NoInventory = rsDetail!Inventory
    '
    Dim rsUpdateInventory As New ADODB.Recordset
    Dim NamaItem, SatuanPesan As String
    Dim Konversi, StockMin, JumlahItemBaru, QtyOnHandBaru, QtyOnHandKecilBaru As Single
    Dim CompStockKembali As Single
    Dim bolResep As Boolean
    '
    rsUpdateInventory.Open "Select Inventory, NamaInventory, IsResep, FSatuanKecil, ReorderLevel, QtyOnHand, QtyOnHandKecil, SatuanBesar, JumlahItem From [C_Inventory] Where Inventory='" & NoInventory & "'", db, adOpenStatic, adLockOptimistic
    NamaItem = rsUpdateInventory!NamaInventory
    bolResep = rsUpdateInventory!IsResep
    Konversi = rsUpdateInventory!FSatuanKecil
    StockMin = rsUpdateInventory!ReorderLevel
    SatuanPesan = rsUpdateInventory!SatuanBesar
    '
    CompStockKembali = rsDetail!UnitMasukAcc
    '
    If bolResep Then
      QtyOnHandKecilBaru = rsUpdateInventory!QtyOnHandKecil + CompStockKembali
      JumlahItemBaru = rsUpdateInventory!JumlahItem + CompStockKembali
      rsUpdateInventory!QtyOnHandKecil = QtyOnHandKecilBaru
    Else
      JumlahItemBaru = ((rsUpdateInventory!QtyOnHand * Konversi) + rsUpdateInventory!QtyOnHandKecil) + CompStockKembali
      QtyOnHandBaru = Int(JumlahItemBaru / Konversi)
      QtyOnHandKecilBaru = JumlahItemBaru - (QtyOnHandBaru * Konversi)
      rsUpdateInventory!QtyOnHand = QtyOnHandBaru
      rsUpdateInventory!QtyOnHandKecil = QtyOnHandKecilBaru
    End If
    '
    rsUpdateInventory!JumlahItem = JumlahItemBaru
    rsUpdateInventory.Update
    rsUpdateInventory.Close
    Set rsUpdateInventory = Nothing
    '
    ' Update Stock Alert
    If bolResep Then
      If QtyOnHandKecilBaru > StockMin Then
        db.Execute "DELETE From [C_StockAlert] Where Inventory='" & NoInventory & "'"
      Else
        db.Execute "UPDATE [C_StockAlert] SET QtyOnHand =" & QtyOnHandKecilBaru & " Where Inventory='" & NoInventory & "'"
      End If
    Else
      If QtyOnHandBaru > StockMin Then
        db.Execute "DELETE From [C_StockAlert] Where Inventory='" & NoInventory & "'"
      Else
        db.Execute "UPDATE [C_StockAlert] SET QtyOnHand =" & QtyOnHandBaru & " Where Inventory='" & NoInventory & "'"
      End If
    End If
    '
    ' Update StockCard Bahan
    Dim rsStockCardBahan As New ADODB.Recordset
    rsStockCardBahan.Open "C_IStockCard", db, adOpenStatic, adLockOptimistic
    rsStockCardBahan.AddNew
    rsStockCardBahan!Tanggal = txtFields(1).Text
    rsStockCardBahan!Inventory = NoInventory
    rsStockCardBahan!NamaInventory = NamaItem
    rsStockCardBahan!Keterangan = "Pengembalian Inventory"
    rsStockCardBahan!PI = txtFields(0).Text
    rsStockCardBahan!UnitMasuk = rsDetail!UnitMasukAcc
    rsStockCardBahan!SatuanBesar = rsDetail!Satuan
    rsStockCardBahan!Harga = rsDetail!Harga
    rsStockCardBahan!HargaSubTotal = rsDetail!HargaSubTotalAcc
    rsStockCardBahan.Update
    rsStockCardBahan.Close
    Set rsStockCardBahan = Nothing
    '
    ' Isi File Perbandingan In-Out
    Dim rsAnalisa As New ADODB.Recordset
    rsAnalisa.Open "Select Kode, Inventory, Kembali from [C_AnalisaInOut] Where Kode='" & txtFields(2).Text & "' and Inventory='" & NoInventory & "'", db, adOpenStatic, adLockOptimistic
    rsAnalisa!Kembali = CompStockKembali
    rsAnalisa.Update
    rsAnalisa.Close
    Set rsAnalisa = Nothing
    '
    ' Isi / Update File Cost Control
    Dim rsControl As New ADODB.Recordset
    rsControl.Open "Select * From [C_CostControl] Where Tanggal=#" & txtFields(1).Text & "# and Inventory='" & NoInventory & "'", db, adOpenStatic, adLockOptimistic
    If rsControl.RecordCount = 0 Then
      rsControl.AddNew
      rsControl!Tanggal = txtFields(1).Text
      rsControl!Inventory = NoInventory
      rsControl!NamaInventory = NamaItem
      rsControl!Kembali = CompStockKembali
      rsControl.Update
    Else
      rsControl!Kembali = rsControl!Kembali + CompStockKembali
      rsControl.Update
    End If
    rsControl.Close
    Set rsControl = Nothing
    '
    rsDetail.MoveNext
  Loop
  '
  db.Execute "Update [C_Minta] Set Done = True Where KodeBuat='" & txtFields(2).Text & "'"
  db.CommitTrans

  sPesan = "Pengembalian Inventory dengan Kode " & txtFields(0).Text & vbCrLf
  sPesan = sPesan & "telah diperiksa"
  MsgBox sPesan, vbInformation
  
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
  rsIndeks.Open "Select KodeKembali FROM [C_Kembali] WHERE KodeBuat LIKE '%" & txtIndeks.Text & "%'", db, adOpenStatic, adLockOptimistic
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
      rsCari.Open "Select * FROM [C_Kembali] WHERE KodeKembali LIKE '%" & txtCari.Text & "%' and NeedCheck=True Order By KodeBuat", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_Kembali] WHERE KodeKembali LIKE '%" & txtCari.Text & "%' and NeedCheck=False Order By KodeBuat", db, adOpenStatic, adLockOptimistic
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

Private Sub dgPI_AfterColEdit(ByVal ColIndex As Integer)
  If ColIndex = 3 Then
    '
    Dim rsKembali As New ADODB.Recordset
    Dim iBatas As Single
    '
    iBatas = 0
    rsKembali.Open "Select NamaInventory, PI, UnitKeluarAcc From [C_AmbilKembaliDetail] Where NamaInventory='" & dgPI.Columns(0).Value & "' and PI='" & txtFields(2).Text & "'", db, adOpenStatic, adLockReadOnly
    If rsKembali.RecordCount <> 0 Then iBatas = rsKembali!UnitKeluarAcc
    rsKembali.Close
    Set rsKembali = Nothing
    '
    If dgPI.Columns(3).Value > iBatas Then
      MsgBox "Jumlah Inventory " & dgPI.Columns(0).Value & " yang diminta hanya " & iBatas, vbInformation
      dgPI.Columns(3).Value = 0
      Exit Sub
    End If
    '
    dgPI.Columns(5).Value = dgPI.Columns(1).Value * dgPI.Columns(3).Value
    SendKeys "{DOWN}"
  End If
End Sub

Private Sub dgPI_AfterUpdate()
  '
  Dim rsJumlahBiaya As New ADODB.Recordset
  rsJumlahBiaya.Open "Select Sum(HargaSubTotalAcc) as Jumlah From [C_AmbilKembaliDetail] Where PI='" & txtFields(0).Text & "'", db, adOpenStatic, adLockReadOnly
  txtFields(7).Text = rsJumlahBiaya!Jumlah
  rsJumlahBiaya.Close
  Set rsJumlahBiaya = Nothing
  '
End Sub

Private Sub cmdCetakBR_Click()
  '
  Dim sSQL As String
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Tidak ada data yang tercetak", vbInformation
    Exit Sub
  End If
  If adoPrimaryRS!NeedCheck = True Then
    MsgBox "Data Pengembalian Inventory belum diproses", vbCritical
    Exit Sub
  End If
  '
  sSQL = "SELECT C_Kembali.KodeKembali, C_Kembali.Tanggal, C_Kembali.NamaKaryawan, C_IStockCard.Inventory, C_IStockCard.NamaInventory, C_IStockCard.UnitMasuk, C_IStockCard.SatuanBesar, C_IStockCard.Harga, C_IStockCard.HargaSubTotal FROM C_Kembali INNER JOIN C_IStockCard ON C_Kembali.KodeKembali = C_IStockCard.PI WHERE (((C_Kembali.KodeKembali)='" & txtFields(0).Text & "'))"

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
    .WindowTitle = "Form Pemeriksaan Pengembalian Inventory"
    .ReportFileName = App.Path & "\Report\Laporan Pemeriksaan Pengembalian Inventory.Rpt"
    .Formulas(0) = "IDCompany= '" & ICNama & "'"
    .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
    .Formulas(2) = "IDKota= '" & ICKota & "'"
    .Action = 1
  End With
  '
End Sub
