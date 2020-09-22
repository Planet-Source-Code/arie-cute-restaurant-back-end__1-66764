VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCIPInvoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faktur Penerimaan Barang (Purchase Invoice)"
   ClientHeight    =   6750
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
   ScaleHeight     =   6750
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   3480
      TabIndex        =   22
      Top             =   60
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   60
         ScaleHeight     =   5055
         ScaleWidth      =   7695
         TabIndex        =   23
         Top             =   660
         Width           =   7695
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   3900
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   660
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaSupplier"
            Height          =   315
            Index           =   5
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1020
            Width           =   3135
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "Supplier"
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   2820
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "123"
            Top             =   1020
            Width           =   615
         End
         Begin VB.CommandButton cmdBarang 
            Caption         =   "&Item Inventory Masuk..."
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
            Left            =   2580
            Picture         =   "frmCIPInvoice.frx":0000
            TabIndex        =   12
            Top             =   4440
            Width           =   2595
         End
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            DataField       =   "TanggalTerima"
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
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   300
            Width           =   1155
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "NamaPI"
            Height          =   555
            Index           =   6
            Left            =   2820
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1380
            Width           =   3795
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "PO"
            Height          =   315
            Index           =   2
            Left            =   2820
            TabIndex        =   10
            Top             =   660
            Width           =   1035
         End
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            DataField       =   "TanggalPesan"
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
            Index           =   3
            Left            =   5460
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   660
            Width           =   1155
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "PI"
            Height          =   315
            Index           =   0
            Left            =   2820
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "1234567890"
            Top             =   300
            Width           =   1035
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "NilaiPI"
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3960
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dgPI 
            Bindings        =   "frmCIPInvoice.frx":030A
            Height          =   1635
            Left            =   60
            TabIndex        =   27
            Top             =   2220
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   2884
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
            Caption         =   "Barang yang diterima"
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
               Caption         =   "Item Inventory"
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
               DataField       =   "JumlahPI"
               Caption         =   "# Terima"
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
               DataField       =   "SatuanBesar"
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
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1844.787
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   989.858
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   734.74
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1335.118
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Supplier"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   37
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal PO"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   4500
            TabIndex        =   35
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Invoice"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4200
            TabIndex        =   32
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   1080
            TabIndex        =   31
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. PO"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   720
            Width           =   510
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Purchase Invoice"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Purchase Invoice"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   3840
            TabIndex        =   28
            Top             =   4020
            Width           =   1965
         End
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCIPInvoice.frx":031F
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Invoice"
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
         TabIndex        =   33
         Top             =   180
         Width           =   2205
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
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6675
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   3375
      Begin TabDlg.SSTab TabIndek 
         Height          =   6195
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   10927
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
         TabPicture(0)   =   "frmCIPInvoice.frx":0BE9
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lstIndeks"
         Tab(0).Control(1)=   "txtIndeks"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCIPInvoice.frx":0C05
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblLabels(11)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lstCari"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "txtCari"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cboBy"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdTampil"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   -74760
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCIPInvoice.frx":0C21
            Left            =   240
            List            =   "frmCIPInvoice.frx":0C2B
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   2835
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ListBox lstIndeks 
            Height          =   5130
            Left            =   -74760
            TabIndex        =   1
            Top             =   840
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   4155
            ItemData        =   "frmCIPInvoice.frx":0C52
            Left            =   240
            List            =   "frmCIPInvoice.frx":0C59
            TabIndex        =   18
            Top             =   1860
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
            TabIndex        =   21
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdCetakPI 
         Caption         =   "&Cetak Invoice"
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
         Picture         =   "frmCIPInvoice.frx":0C66
         TabIndex        =   2
         Top             =   6240
         Width           =   3315
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   60
      TabIndex        =   34
      Top             =   5880
      Width           =   11235
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCIPInvoice.frx":0F70
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCIPInvoice.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCIPInvoice.frx":1584
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCIPInvoice.frx":188E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCIPInvoice.frx":1B98
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCIPInvoice.frx":1EA2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCIPInvoice.frx":21AC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCIPInvoice.frx":24B6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCIPInvoice"
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
  strSQL = "SELECT * FROM [C_IStockCard] WHERE PI='" & strKey & "' Order By StockCard"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgPI.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgPI.Caption = "Barang yg diterima dari " & adoPrimaryRS!NamaSupplier
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
  adoPrimaryRS.Open "select * from [C_IPI] Order by PI", db, adOpenStatic, adLockOptimistic
  
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
  On Error GoTo AddErr
  '
  StatusFrame True
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
  For iCounter = 0 To 2 Step 2
    If txtFields(iCounter) = "" Then
      MsgBox "Field ini dibutuhkan", vbCritical
      txtFields(iCounter).SetFocus
      Exit Sub
    End If
  Next
  
  If rsDetail.RecordCount = 0 Then
    MsgBox "Sebuah Purchase Invoice harus berisi sedikitnya satu jenis barang", vbCritical
    cmdBarang.SetFocus
    Exit Sub
  End If
  '
  adoPrimaryRS.UpdateBatch adAffectAll
  '
  ' Isi File History
  Dim rsHistory As New ADODB.Recordset
  rsHistory.Open "History", db, adOpenStatic, adLockOptimistic
  rsHistory.AddNew
  rsHistory!KodeRef = txtFields(0).Text
  rsHistory!Tanggal = txtFields(1).Text
  rsHistory!Keterangan = "Purchase Invoice dari PO : " & txtFields(2).Text
  rsHistory!Nilai = txtFields(7).Text
  rsHistory!Jenis = "PI"
  rsHistory.Update
  rsHistory.Close
  Set rsHistory = Nothing
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
  rsIndeks.Open "Select PI FROM [C_IPI]", db, adOpenStatic, adLockOptimistic
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
    lstIndeks.AddItem rsIndeks!PI
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
    lstCari.AddItem rsCari!PI
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select PI FROM [C_IPI] WHERE PI LIKE '%" & txtIndeks.Text & "%'", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "PI ='" & lstIndeks.Text & "'"
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
      rsCari.Open "Select * FROM [C_IPI] WHERE PI LIKE '%" & txtCari.Text & "%' Order By PI", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_IPI] WHERE NamaPI LIKE '%" & txtCari.Text & "%' Order By PI", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "PI ='" & lstCari.Text & "'"
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
        If JumlahRecord("Select KodeRef From [History] Where KodeRef='" & txtFields(0).Text & "'", db) <> 0 Then
          MsgBox "Kode Dokumen " & txtFields(0).Text & " telah digunakan oleh dokumen lain", vbCritical
          txtFields(0).Text = ""
          txtFields(0).SetFocus
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case 2
        sKontrolAktif = "API_1"
        Dim rsPO As New ADODB.Recordset
        rsPO.Open "Select PO, TanggalPesan, Supplier, NamaSupplier from [C_IPO] Where PO='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsPO.RecordCount <> 0 Then
          txtFields(3).Text = rsPO!TanggalPesan
          txtFields(4).Text = rsPO!Supplier
          txtFields(5).Text = rsPO!NamaSupplier
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(3).Text = ""
          txtFields(4).Text = ""
          txtFields(5).Text = ""
          frmCari.Show vbModal
          txtFields(2).SetFocus
        End If
        rsPO.Close
        sKontrolAktif = ""
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub cmdBarang_Click()
  '
  If (frmCIPInvoice.txtFields(0).Text = "") Or (frmCIPInvoice.txtFields(2).Text = "") Then
    MsgBox "Nomor Invoice atau Nomor PO tidak boleh kosong", vbCritical
    frmCIPInvoice.txtFields(0).SetFocus
    Exit Sub
  End If
  '
  frmCIPIBarang.Caption = "No. Invoice : " & adoPrimaryRS!PI & " - Terima Barang"
  frmCIPIBarang.Show vbModal
  ProsesDetail
  '
End Sub

Private Sub cmdCetakPI_Click()
  '
  Dim sSQL As String
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Tidak ada data yang tercetak", vbInformation
    Exit Sub
  End If
  '
  sSQL = "SELECT C_IPI.PI, C_IPI.PO, C_IPI.TanggalTerima, C_IPI.NamaSupplier, C_IStockCard.Inventory, C_IStockCard.NamaInventory, C_IStockCard.JumlahPI, C_IStockCard.SatuanBesar, C_IStockCard.Harga, C_IStockCard.HargaSubTotal FROM C_IPI INNER JOIN C_IStockCard ON C_IPI.PI = C_IStockCard.PI WHERE (((C_IPI.PI)='" & txtFields(0).Text & "'))"

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
    .WindowTitle = "Purchase Invoice"
    .ReportFileName = App.Path & "\Report\Laporan PI.Rpt"
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
      sKontrolAktif = "API_1"
      frmCari.Show vbModal
      txtFields(2).SetFocus
      If txtFields(2).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub

