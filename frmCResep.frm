VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCResep 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resep Standar"
   ClientHeight    =   6945
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
   ScaleHeight     =   6945
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5955
      Left            =   3480
      TabIndex        =   27
      Top             =   60
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   5235
         Left            =   60
         ScaleHeight     =   5235
         ScaleWidth      =   7695
         TabIndex        =   28
         Top             =   660
         Width           =   7695
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataField       =   "PersenYield"
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
            Index           =   9
            Left            =   2100
            TabIndex        =   14
            Text            =   "100.00"
            Top             =   4260
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "BiayaBuat"
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3900
            Width           =   1755
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "BiayaYield"
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
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   4260
            Width           =   1755
         End
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   5700
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   900
            Width           =   375
         End
         Begin VB.CommandButton cmdHapusBahan 
            Caption         =   "&Hapus Bahan"
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
            Left            =   5340
            Picture         =   "frmCResep.frx":0000
            TabIndex        =   17
            Top             =   4680
            Width           =   2055
         End
         Begin VB.CommandButton cmdEditBahan 
            Caption         =   "&Edit Bahan"
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
            Left            =   5340
            Picture         =   "frmCResep.frx":030A
            TabIndex        =   16
            Top             =   4260
            Width           =   2055
         End
         Begin VB.CommandButton cmdAddBahan 
            Caption         =   "&Tambah Bahan"
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
            Left            =   5340
            Picture         =   "frmCResep.frx":0614
            TabIndex        =   15
            Top             =   3840
            Width           =   2055
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "JumlahFisik"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2280
            TabIndex        =   13
            Top             =   1260
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaSatuan"
            Height          =   315
            Index           =   3
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   900
            Width           =   2115
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Satuan"
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   12
            Text            =   "1234567890"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "NamaResep"
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   11
            Top             =   540
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Resep"
            Height          =   315
            Index           =   0
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "1234567890"
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "BiayaPerUnit"
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
            Index           =   5
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   180
            Width           =   1755
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "BiayaFisik"
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
            Index           =   6
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   4620
            Width           =   1755
         End
         Begin MSDataGridLib.DataGrid dgBahanResep 
            Height          =   1935
            Left            =   60
            TabIndex        =   30
            Top             =   1800
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3413
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
            Caption         =   "Bahan-bahan"
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "Inventory"
               Caption         =   "Kode"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "NamaInventory"
               Caption         =   "Nama Bahan"
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
               DataField       =   "JumlahKonsumsi"
               Caption         =   "Jml Konsumsi"
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
               DataField       =   "NamaSatuan"
               Caption         =   ""
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
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2459.906
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   824.882
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1544.882
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Persentase Terbuang"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   420
            TabIndex        =   48
            Top             =   4320
            Width           =   1545
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   2760
            TabIndex        =   47
            Top             =   3960
            Width           =   495
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% ="
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   2940
            TabIndex        =   45
            Top             =   4320
            Width           =   330
         End
         Begin VB.Label lblSatuan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            DataField       =   "NamaSatuan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   3660
            TabIndex        =   42
            Top             =   1320
            Width           =   510
         End
         Begin VB.Label lblSatuan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            DataField       =   "NamaSatuan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   4740
            TabIndex        =   39
            Top             =   240
            Width           =   510
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Biaya per"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   4020
            TabIndex        =   38
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Fisik"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   20
            Left            =   1020
            TabIndex        =   35
            Top             =   1320
            Width           =   840
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1020
            TabIndex        =   34
            Top             =   960
            Width           =   510
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Resep"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   33
            Top             =   600
            Width           =   900
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Resep"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1020
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Biaya Pembuatan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   1620
            TabIndex        =   31
            Top             =   4680
            Width           =   1650
         End
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCResep.frx":091E
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resep Standar"
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
         TabIndex        =   36
         Top             =   180
         Width           =   1860
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
      Height          =   6855
      Left            =   60
      TabIndex        =   24
      Top             =   60
      Width           =   3375
      Begin TabDlg.SSTab TabIndek 
         Height          =   6795
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   11986
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
         TabPicture(0)   =   "frmCResep.frx":11E8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCResep.frx":1204
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstCari"
         Tab(1).Control(1)=   "cmdTampil"
         Tab(1).Control(2)=   "cboBy"
         Tab(1).Control(3)=   "txtCari"
         Tab(1).Control(4)=   "lblLabels(11)"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   4740
            ItemData        =   "frmCResep.frx":1220
            Left            =   -74760
            List            =   "frmCResep.frx":1227
            TabIndex        =   23
            Top             =   1860
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   22
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCResep.frx":1234
            Left            =   -74760
            List            =   "frmCResep.frx":123E
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   720
            Width           =   2835
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   -74760
            TabIndex        =   21
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ListBox lstIndeks 
            Height          =   5715
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
            TabIndex        =   26
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   60
      TabIndex        =   37
      Top             =   6060
      Width           =   11235
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   7860
         Picture         =   "frmCResep.frx":125A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCResep.frx":1564
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCResep.frx":186E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCResep.frx":1B78
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCResep.frx":1E82
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCResep.frx":218C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCResep.frx":2496
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Koreksi"
         Height          =   795
         Left            =   8700
         Picture         =   "frmCResep.frx":27A0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Hapus"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCResep.frx":2AAA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCResep.frx":2DB4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCResep"
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
  strSQL = "SELECT * FROM [C_ResepBahan] WHERE Resep='" & strKey & "'"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgBahanResep.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgBahanResep.Caption = "Bahan-bahan " & adoPrimaryRS!NamaResep
  End If
  dgBahanResep.ReBind
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
  adoPrimaryRS.Open "Select * from [C_Resep] Order by NamaResep", db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set lblSatuan(0).DataSource = adoPrimaryRS
  Set lblSatuan(1).DataSource = adoPrimaryRS
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
  iCounter = JumlahRecord("Select Resep From [C_Resep] Where Resep Like 'R%'", db)
  txtFields(0).Text = "R-" & iCounter + 1
  txtFields(6).Text = 0
  txtFields(7).Text = 0
  txtFields(8).Text = 0
  txtFields(9).Text = 0
  txtFields(1).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Resep kosong", vbCritical
    Exit Sub
  End If
  
  mvBookMark = adoPrimaryRS.Bookmark
  StatusFrame True
  '
  db.BeginTrans
  mbEditFlag = True
  SetButtons False
  '
  txtFields(0).Enabled = False
  txtFields(1).SetFocus
  SendKeys "{END}+{HOME}"
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Resep Standar kosong", vbCritical
    Exit Sub
  End If
  '
  If MsgBox("Hapus Resep Standar '" & adoPrimaryRS!NamaResep & "' ?", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
  Else
    '
    If JumlahRecord("Select Inventory From [C_MenuBahan] Where Inventory='" & txtFields(0).Text & "'", db) <> 0 Then
      MsgBox "Resep standar ini masih digunakan dalam produk menu", vbCritical
      Exit Sub
    End If
    
    db.BeginTrans
    db.Execute "DELETE FROM [C_RESEPBAHAN] WHERE RESEP='" & txtFields(0).Text & "'"
    db.Execute "DELETE FROM [C_INVENTORY] WHERE INVENTORY='" & txtFields(0).Text & "'"
    With adoPrimaryRS
      .Delete
      .MoveNext
      If .EOF Then .MoveLast
    End With
    db.CommitTrans
  End If
  '
  ProsesDetail
  RefreshIndeks
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  ' Validasi Field
  If txtFields(1).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(1).SetFocus
    Exit Sub
  End If
  If txtFields(2).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(2).SetFocus
    Exit Sub
  End If
  If txtFields(4).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(4).SetFocus
    Exit Sub
  End If
  '
  If rsDetail.RecordCount = 0 Then
    MsgBox "Bahan pembuat sebuah resep tidak boleh kosong", vbCritical
    cmdAddBahan.SetFocus
    Exit Sub
  End If
  '
  txtFields(0).Enabled = True
  adoPrimaryRS!BiayaPerUnit = txtFields(6).Text / txtFields(4).Text
  adoPrimaryRS.UpdateBatch adAffectAll
  
  Dim rsInventory As New ADODB.Recordset
  rsInventory.Open "C_Inventory", db, adOpenStatic, adLockOptimistic
  
  ' Resep Baru ?
  If mbAddNewFlag Then
    '
    ' Buat Item Inventory untuk resep
    rsInventory.AddNew
    rsInventory!Inventory = txtFields(0).Text
    rsInventory!NamaInventory = txtFields(1).Text
    rsInventory!IsResep = True
    rsInventory!HargaPerUnit = txtFields(5).Text
    rsInventory!SatuanBesar = txtFields(2).Text
    rsInventory!NamaSatuanBesar = txtFields(3).Text
    rsInventory!FSatuanKecil = 1
    rsInventory!SatuanKecil = txtFields(2).Text
    rsInventory!NamaSatuanKecil = txtFields(3).Text
    rsInventory.Update
    '
  End If
  '
  ' Edit Resep Lama ?
  If mbEditFlag Then
    '
    ' Update Harga Bahan Menu
    db.Execute "Update [C_MenuBahan] SET HargaPerUnit=" & adoPrimaryRS!BiayaPerUnit & " Where Inventory = '" & adoPrimaryRS!Resep & "'"
    db.Execute "Update [C_MenuBahan] SET SubTotal=HargaPerUnit*JumlahKonsumsi"
    
    ' Update File Menu
    Dim sKodeMenu As String
    Dim sngFaktor, sngMYield As Single
    Dim curMBiayaYield, curMBiaya As Currency
    Dim rsMenu As New ADODB.Recordset
    rsMenu.Open "Select * From [C_Menu]", db, adOpenStatic, adLockOptimistic
    '
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
    End If
    rsMenu.Close
    Set rsMenu = Nothing
    '
    rsInventory.Find "Inventory='" & txtFields(0).Text & "'"
    rsInventory!NamaInventory = txtFields(1).Text
    rsInventory!IsResep = True
    rsInventory!HargaPerUnit = txtFields(5).Text
    rsInventory!SatuanBesar = txtFields(2).Text
    rsInventory!NamaSatuanBesar = txtFields(3).Text
    rsInventory!FSatuanKecil = 1
    rsInventory!SatuanKecil = txtFields(2).Text
    rsInventory!NamaSatuanKecil = txtFields(3).Text
    rsInventory.Update
    '
  End If
  rsInventory.Close
  Set rsInventory = Nothing
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

  txtFields(0).Enabled = True
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
  Unload Me
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
  cmdEdit.Visible = bVal
  cmdDelete.Visible = bVal
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
  rsIndeks.Open "Select NamaResep FROM [C_Resep] Order By NamaResep", db, adOpenStatic, adLockOptimistic
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
    lstIndeks.AddItem rsIndeks!NamaResep
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
    lstCari.AddItem rsCari!NamaResep
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select NamaResep FROM [C_Resep] WHERE NamaResep LIKE '%" & txtIndeks.Text & "%' Order By NamaResep", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "NamaResep ='" & lstIndeks.Text & "'"
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
      rsCari.Open "Select * FROM [C_Resep] WHERE Resep LIKE '%" & txtCari.Text & "%' Order By Resep", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_Resep] WHERE NamaResep LIKE '%" & txtCari.Text & "%' Order By NamaResep", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "NamaResep ='" & lstCari.Text & "'"
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
      Case 2
        sKontrolAktif = "AR_1"
        Dim rsSatuan As New ADODB.Recordset
        rsSatuan.Open "Select Satuan, NamaSatuan from [C_Satuan] Where Satuan='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsSatuan.RecordCount <> 0 Then
          txtFields(3).Text = rsSatuan!NamaSatuan
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(3).Text = ""
          frmCari.Show vbModal
          txtFields(2).SetFocus
        End If
        rsSatuan.Close
        For iCounter = 0 To 1
          lblSatuan(iCounter).Caption = txtFields(3).Text
        Next
        sKontrolAktif = ""
      Case 4
        If Not IsNumeric(txtFields(4).Text) Then
          Exit Sub
        End If
        If txtFields(4).Text = "" Then txtFields(4).Text = 1
        txtFields(5).Text = CCur(txtFields(6).Text) / CCur(txtFields(4).Text)
        SendKeys "{TAB}"
      Case 9
        If Not IsNumeric(txtFields(9).Text) Then
          Exit Sub
        End If
        txtFields(7).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
        txtFields(6).Text = CCur(txtFields(8).Text) + CCur(txtFields(7).Text)
        txtFields(5).Text = CCur(txtFields(6).Text) / CCur(txtFields(4).Text)
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub cmdAddBahan_Click()
  '
  frmCResep.Tag = "1"
  frmCResepBahan.Show vbModal
  ProsesDetail
  '
End Sub

Private Sub cmdEditBahan_Click()
  '
  If rsDetail.RecordCount = 0 Then
    MsgBox "Data Bahan Resep kosong", vbCritical
    Exit Sub
  End If
  '
  frmCResep.Tag = "2"
  txtFields(8).Text = CCur(txtFields(8).Text) - CCur(dgBahanResep.Columns(4).Text)
  txtFields(7).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
  txtFields(6).Text = CCur(txtFields(8).Text) + CCur(txtFields(7).Text)
  With frmCResepBahan
    .txtFields(0).Text = dgBahanResep.Columns(0).Text
    .txtFields(0).Enabled = False
    .txtFields(1).Text = dgBahanResep.Columns(1).Text
    .txtFields(2).Text = dgBahanResep.Columns(2).Text
    .txtFields(3).Text = dgBahanResep.Columns(4).Text
    .cmdLookUp(0).Enabled = False
    .Show vbModal
  End With
  ProsesDetail
  '
End Sub

Private Sub cmdHapusBahan_Click()
  '
  If rsDetail.RecordCount = 0 Then
    MsgBox "Data Bahan Resep kosong", vbCritical
    Exit Sub
  End If
  '
  sPesan = dgBahanResep.Columns(1).Text & " akan dihapus dari daftar bahan resep " & txtFields(1).Text & "." & vbCrLf
  sPesan = sPesan & "Anda Yakin ?"
  If MsgBox(sPesan, vbExclamation + vbYesNo) = vbYes Then
    txtFields(8).Text = CCur(txtFields(8).Text) - CCur(dgBahanResep.Columns(4).Text)
    txtFields(7).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
    txtFields(6).Text = CCur(txtFields(8).Text) + CCur(txtFields(7).Text)
    txtFields(5).Text = CCur(txtFields(6).Text) / CCur(txtFields(4).Text)
    db.Execute "DELETE FROM [C_ResepBahan] WHERE Resep='" & txtFields(0).Text & "' and Inventory = '" & dgBahanResep.Columns(0).Text & "'"
    ProsesDetail
  End If
  '
  ProsesDetail
End Sub

Private Sub cmdLookUp_Click(Index As Integer)
  Select Case Index
    Case 0
      sKontrolAktif = "AR_1"
      frmCari.Show vbModal
      txtFields(2).SetFocus
      If txtFields(2).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub
