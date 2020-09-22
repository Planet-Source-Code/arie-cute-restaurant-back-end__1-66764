VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCInventory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
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
   ScaleHeight     =   7050
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6075
      Left            =   3480
      TabIndex        =   39
      Top             =   60
      Width           =   6675
      Begin TabDlg.SSTab tabInventory 
         Height          =   5295
         Left            =   60
         TabIndex        =   26
         Top             =   720
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   9340
         _Version        =   393216
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
         TabCaption(0)   =   "Informasi"
         TabPicture(0)   =   "frmCInventory.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Picture1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraFormula"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Supplier"
         TabPicture(1)   =   "frmCInventory.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Picture1(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Stock Card"
         TabPicture(2)   =   "frmCInventory.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Picture1(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame fraFormula 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   60
            TabIndex        =   59
            Top             =   4740
            Width           =   6375
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "FSatuanKecil"
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
               Index           =   15
               Left            =   3720
               TabIndex        =   18
               Top             =   60
               Width           =   1215
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   15
               Left            =   2280
               TabIndex        =   69
               Top             =   120
               Width           =   90
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Konversi Satuan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   10
               Left            =   480
               TabIndex        =   63
               Top             =   120
               Width           =   1380
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "="
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   11
               Left            =   3480
               TabIndex        =   62
               Top             =   120
               Width           =   120
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanBesar"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   3
               Left            =   2640
               TabIndex        =   61
               Top             =   120
               Width           =   495
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanKecil"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   4
               Left            =   5100
               TabIndex        =   60
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   4875
            Index           =   2
            Left            =   -74940
            ScaleHeight     =   4875
            ScaleWidth      =   6375
            TabIndex        =   58
            Top             =   360
            Width           =   6375
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
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
               Index           =   21
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   4380
               Width           =   1875
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               DataField       =   "HargaPerUnit"
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
               Index           =   18
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   77
               TabStop         =   0   'False
               Text            =   "999,999,999.00"
               Top             =   3660
               Width           =   1335
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               DataField       =   "JumlahItem"
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
               Index           =   20
               Left            =   1620
               Locked          =   -1  'True
               TabIndex        =   75
               TabStop         =   0   'False
               Text            =   "999,999,999.00"
               Top             =   4380
               Width           =   1515
            End
            Begin VB.TextBox txtFields 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               DataField       =   "NamaSupplier"
               Height          =   315
               Index           =   19
               Left            =   1620
               Locked          =   -1  'True
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   4020
               Width           =   2235
            End
            Begin MSDataGridLib.DataGrid dgStockCard 
               Height          =   3315
               Left            =   180
               TabIndex        =   68
               Top             =   120
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   5847
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
               Caption         =   "Stock Card"
               ColumnCount     =   6
               BeginProperty Column00 
                  DataField       =   "Tanggal"
                  Caption         =   "Tanggal"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "dd-MM-yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "PI"
                  Caption         =   "Ref #"
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
                  DataField       =   "UnitPesan"
                  Caption         =   "# Pesan"
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
                  DataField       =   "UnitMasuk"
                  Caption         =   "# Masuk"
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
                  DataField       =   "UnitKeluar"
                  Caption         =   "# Keluar"
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
                  DataField       =   "UnitRusak"
                  Caption         =   "# Rusak"
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
                     Alignment       =   2
                     ColumnAllowSizing=   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   1080
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column05 
                     Alignment       =   1
                     ColumnWidth     =   945.071
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nilai Item Inventory"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Index           =   18
               Left            =   4320
               TabIndex        =   79
               Top             =   4140
               Width           =   1710
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanKecil"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   7
               Left            =   3300
               TabIndex        =   76
               Top             =   4440
               Width           =   495
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jumlah Inventory"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   17
               Left            =   240
               TabIndex        =   74
               Top             =   4440
               Width           =   1260
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Supplier Terakhir"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   16
               Left            =   240
               TabIndex        =   73
               Top             =   4080
               Width           =   1200
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "per"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   13
               Left            =   3960
               TabIndex        =   66
               Top             =   3720
               Width           =   240
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Harga Item Inventory terakhir"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   12
               Left            =   240
               TabIndex        =   65
               Top             =   3720
               Width           =   2175
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanKecil"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   5
               Left            =   4260
               TabIndex        =   64
               Top             =   3720
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   4875
            Index           =   1
            Left            =   -74940
            ScaleHeight     =   4875
            ScaleWidth      =   6375
            TabIndex        =   56
            Top             =   360
            Width           =   6375
            Begin VB.CommandButton cmdDelSupp 
               Caption         =   "&Hapus Supplier"
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
               Left            =   3360
               Picture         =   "frmCInventory.frx":0054
               TabIndex        =   27
               Top             =   4020
               Width           =   2835
            End
            Begin VB.CommandButton cmdAddSupp 
               Caption         =   "Suppli&er Baru"
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
               Left            =   180
               Picture         =   "frmCInventory.frx":035E
               TabIndex        =   25
               Top             =   4020
               Width           =   3135
            End
            Begin MSDataGridLib.DataGrid dgSuppInv 
               Height          =   3675
               Left            =   180
               TabIndex        =   57
               Top             =   300
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   6482
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   16777215
               ForeColor       =   0
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
               Caption         =   "Supplier"
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "Supplier"
                  Caption         =   "Kode"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "NamaSupplier"
                  Caption         =   "Nama Supplier"
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
                  DataField       =   "LastPrice"
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   2
                  BeginProperty Column00 
                     ColumnAllowSizing=   -1  'True
                     ColumnWidth     =   989.858
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   2954.835
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                  EndProperty
               EndProperty
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   4395
            Index           =   0
            Left            =   60
            ScaleHeight     =   4395
            ScaleWidth      =   6375
            TabIndex        =   41
            Top             =   360
            Width           =   6375
            Begin VB.CommandButton cmdLookUp 
               Caption         =   "..."
               Height          =   315
               Index           =   3
               Left            =   5220
               TabIndex        =   84
               TabStop         =   0   'False
               Top             =   2040
               Width           =   375
            End
            Begin VB.CommandButton cmdLookUp 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   5220
               TabIndex        =   83
               TabStop         =   0   'False
               Top             =   1680
               Width           =   375
            End
            Begin VB.CommandButton cmdLookUp 
               Caption         =   "..."
               Height          =   315
               Index           =   1
               Left            =   5220
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   1320
               Width           =   375
            End
            Begin VB.CommandButton cmdLookUp 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   5220
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   960
               Width           =   375
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
               Index           =   17
               Left            =   4260
               Locked          =   -1  'True
               TabIndex        =   78
               TabStop         =   0   'False
               Text            =   "999,999,999.00"
               Top             =   2580
               Width           =   1335
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               DataField       =   "QtyOnHandKecil"
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
               Index           =   16
               Left            =   4260
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Top             =   3300
               Width           =   675
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Catatan"
               Height          =   675
               Index           =   10
               Left            =   480
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Top             =   2760
               Width           =   2175
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               DataField       =   "ReorderLevel"
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
               Index           =   14
               Left            =   4260
               TabIndex        =   17
               Top             =   4020
               Width           =   675
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               DataField       =   "QtyOnHand"
               Height          =   315
               Index           =   12
               Left            =   4260
               Locked          =   -1  'True
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   2940
               Width           =   675
            End
            Begin VB.TextBox txtFields 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               DataField       =   "QtyOnOrder"
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
               Index           =   13
               Left            =   4260
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   3660
               Width           =   675
            End
            Begin VB.TextBox txtFields 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               DataField       =   "LastInvoiceDate"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "d MMMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Index           =   11
               Left            =   480
               Locked          =   -1  'True
               TabIndex        =   36
               TabStop         =   0   'False
               Text            =   "99/99/9999"
               Top             =   3720
               Width           =   2175
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "SatuanKecil"
               Height          =   315
               Index           =   8
               Left            =   2220
               TabIndex        =   15
               Top             =   2040
               Width           =   1035
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00E0E0E0&
               DataField       =   "NamaSatuanKecil"
               Height          =   315
               Index           =   9
               Left            =   3300
               Locked          =   -1  'True
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   2040
               Width           =   1875
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "SatuanBesar"
               Height          =   315
               Index           =   6
               Left            =   2220
               TabIndex        =   14
               Top             =   1680
               Width           =   1035
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00E0E0E0&
               DataField       =   "NamaSatuanBesar"
               Height          =   315
               Index           =   7
               Left            =   3300
               Locked          =   -1  'True
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1680
               Width           =   1875
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00E0E0E0&
               DataField       =   "NamaKategori"
               Height          =   315
               Index           =   5
               Left            =   2940
               Locked          =   -1  'True
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   1320
               Width           =   2235
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Kategori"
               Height          =   315
               Index           =   4
               Left            =   2220
               TabIndex        =   13
               Top             =   1320
               Width           =   675
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00E0E0E0&
               DataField       =   "NamaGudang"
               Height          =   315
               Index           =   3
               Left            =   2940
               Locked          =   -1  'True
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   960
               Width           =   2235
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Gudang"
               Height          =   315
               Index           =   2
               Left            =   2220
               TabIndex        =   12
               Top             =   960
               Width           =   675
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "NamaInventory"
               Height          =   315
               Index           =   1
               Left            =   2220
               TabIndex        =   11
               Top             =   600
               Width           =   3375
            End
            Begin VB.TextBox txtFields 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Inventory"
               Height          =   315
               Index           =   0
               Left            =   2220
               MaxLength       =   10
               TabIndex        =   10
               Text            =   "1234567890"
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanKecil"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   6
               Left            =   5100
               TabIndex        =   71
               Top             =   3360
               Width           =   495
            End
            Begin VB.Label lblInv 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Harga Pembelian"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   14
               Left            =   2880
               TabIndex        =   67
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   20
               Left            =   540
               TabIndex        =   55
               Top             =   2520
               Width           =   585
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Invoice Terakhir"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   9
               Left            =   480
               TabIndex        =   54
               Top             =   3480
               Width           =   1770
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Satuan Pakai"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   8
               Left            =   840
               TabIndex        =   53
               Top             =   2100
               Width           =   930
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Satuan Beli"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   7
               Left            =   840
               TabIndex        =   52
               Top             =   1740
               Width           =   795
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanBesar"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   2
               Left            =   5100
               TabIndex        =   51
               Top             =   4080
               Width           =   495
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanBesar"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   1
               Left            =   5100
               TabIndex        =   50
               Top             =   3720
               Width           =   495
            End
            Begin VB.Label lblSatuan 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "satuan"
               DataField       =   "NamaSatuanBesar"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   5100
               TabIndex        =   49
               Top             =   3000
               Width           =   495
            End
            Begin VB.Label lblInv 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Reorder Level"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   6
               Left            =   2880
               TabIndex        =   48
               Top             =   4080
               Width           =   1215
            End
            Begin VB.Label lblInv 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Qty On Order"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   5
               Left            =   2880
               TabIndex        =   47
               Top             =   3720
               Width           =   1215
            End
            Begin VB.Label lblInv 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Qty On Hand"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   4
               Left            =   2880
               TabIndex        =   46
               Top             =   3000
               Width           =   1215
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kategori"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   3
               Left            =   840
               TabIndex        =   45
               Top             =   1380
               Width           =   600
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gudang"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   2
               Left            =   840
               TabIndex        =   44
               Top             =   1020
               Width           =   555
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   1
               Left            =   840
               TabIndex        =   43
               Top             =   660
               Width           =   840
            End
            Begin VB.Label lblInv 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Inventory"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   840
               TabIndex        =   42
               Top             =   300
               Width           =   1125
            End
         End
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   240
         Picture         =   "frmCInventory.frx":0668
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   900
         TabIndex        =   40
         Top             =   180
         Width           =   2025
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   60
         Top             =   120
         Width           =   6495
      End
   End
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   60
      TabIndex        =   29
      Top             =   60
      Width           =   3375
      Begin TabDlg.SSTab TabIndek 
         Height          =   6915
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   12197
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
         TabPicture(0)   =   "frmCInventory.frx":0F32
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCInventory.frx":0F4E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstCari"
         Tab(1).Control(1)=   "cmdTampil"
         Tab(1).Control(2)=   "cboBy"
         Tab(1).Control(3)=   "txtCari"
         Tab(1).Control(4)=   "lblLabels(11)"
         Tab(1).ControlCount=   5
         Begin VB.Frame Frame2 
            Caption         =   "Stok Filter"
            Height          =   1155
            Left            =   240
            TabIndex        =   85
            Top             =   5520
            Width           =   2835
            Begin VB.OptionButton optFilter 
               Caption         =   "Resep Standar"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   87
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton optFilter 
               Caption         =   "Raw Material"
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   86
               Top             =   360
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin VB.ListBox lstIndeks 
            Height          =   4545
            Left            =   240
            TabIndex        =   1
            Top             =   840
            Width           =   2835
         End
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   4740
            ItemData        =   "frmCInventory.frx":0F6A
            Left            =   -74760
            List            =   "frmCInventory.frx":0F71
            TabIndex        =   24
            Top             =   1860
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   23
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCInventory.frx":0F7E
            Left            =   -74760
            List            =   "frmCInventory.frx":0F88
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   720
            Width           =   2835
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   -74760
            TabIndex        =   22
            Top             =   1080
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
            TabIndex        =   31
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
      TabIndex        =   28
      Top             =   6180
      Width           =   10155
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Hapus"
         Height          =   795
         Left            =   8460
         Picture         =   "frmCInventory.frx":0FB1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   6780
         Picture         =   "frmCInventory.frx":12BB
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   9300
         Picture         =   "frmCInventory.frx":15C5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   5940
         Picture         =   "frmCInventory.frx":18CF
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3420
         Picture         =   "frmCInventory.frx":1BD9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4260
         Picture         =   "frmCInventory.frx":1EE3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5100
         Picture         =   "frmCInventory.frx":21ED
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   9300
         Picture         =   "frmCInventory.frx":24F7
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Koreksi"
         Height          =   795
         Left            =   7620
         Picture         =   "frmCInventory.frx":2801
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   8460
         Picture         =   "frmCInventory.frx":2B0B
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim rsDetail As ADODB.Recordset
Dim rsStockCard As ADODB.Recordset
Dim rsIndeks As ADODB.Recordset
Dim rsCari As ADODB.Recordset

Dim PosisiRecord As Long
Dim inDML As Boolean

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
  strSQL = "SELECT * FROM [C_Supplier_Inventory] WHERE Inventory='" & strKey & "' Order By Supplier"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgSuppInv.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgSuppInv.Caption = "Supplier " & adoPrimaryRS!NamaInventory
  End If
  dgSuppInv.ReBind
  '
  Set rsStockCard = New ADODB.Recordset
  strSQL = "SELECT * FROM [C_IStockCard] WHERE Inventory='" & strKey & "' Order By StockCard"
  rsStockCard.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsStockCard.Requery
  '
  Set dgStockCard.DataSource = rsStockCard
  If adoPrimaryRS.RecordCount <> 0 Then
    dgStockCard.Caption = "Stock Card " & adoPrimaryRS!NamaInventory & " ( " & adoPrimaryRS!NamaSatuanKecil & " ) "
  End If
  dgStockCard.ReBind
  If adoPrimaryRS.RecordCount <> 0 Then
    txtFields(21).Text = Format(adoPrimaryRS!HargaPerUnit * adoPrimaryRS!JumlahItem, "###,###,###,##0.00")
  End If
  '
End Sub

Private Sub StatusFrame(bolStatus As Boolean)
  '
  Picture1(0).Enabled = bolStatus
  fraFormula.Enabled = bolStatus
  fraFind.Enabled = Not bolStatus
  '
End Sub

Private Sub Form_Load()
  '
  sFormAktif = Me.Name
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "Select * from [C_Inventory] Order by NamaInventory", db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  For iCounter = 0 To 7
    Set lblSatuan(iCounter).DataSource = adoPrimaryRS
  Next
  '
  ProsesDetail
  '
  RefreshIndeks optFilter(0).Value
  lstCari.Clear
  TabIndek.Tab = 0
  tabInventory.Tab = 0
  optFilter(0).Value = True
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
  inDML = True
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
  TabIndek.Tab = 0
  tabInventory.Tab = 0
  '
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Inventory kosong", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS!IsResep = True Then
    For iCounter = 0 To 1
      txtFields(iCounter).Enabled = False
    Next
    For iCounter = 2 To 8 Step 2
      txtFields(iCounter).Enabled = False
    Next
    txtFields(15).Enabled = False
  End If
  
  mvBookMark = adoPrimaryRS.Bookmark
  StatusFrame True
  inDML = True
  '
  db.BeginTrans
  mbEditFlag = True
  SetButtons False
  '
  TabIndek.Tab = 0
  tabInventory.Tab = 0
  '
  txtFields(0).Enabled = False
  If adoPrimaryRS!IsResep = True Then
    txtFields(14).SetFocus
  Else
    txtFields(1).SetFocus
  End If
  '
  SendKeys "{END}+{HOME}"
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error Resume Next
  '
  If adoPrimaryRS!IsResep = True Then
    MsgBox "Data Resep Standar hanya bisa dihapus di modul Resep Standar", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Inventory kosong", vbCritical
    Exit Sub
  End If
  '
  If MsgBox("Hapus Inventory '" & adoPrimaryRS!NamaInventory & "' ?", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
  Else
    '
    If JumlahRecord("Select Inventory From [C_ResepBahan] Where Inventory='" & txtFields(0).Text & "'", db) <> 0 Then
      MsgBox "Item inventory ini masih digunakan dalam resep standar", vbCritical
      Exit Sub
    End If
    If JumlahRecord("Select Inventory From [C_MenuBahan] Where Inventory='" & txtFields(0).Text & "'", db) <> 0 Then
      MsgBox "Item inventory ini masih digunakan dalam produk menu", vbCritical
      Exit Sub
    End If
    If adoPrimaryRS!JumlahItem > 0 Then
      MsgBox "Stok Inventory masih ada", vbCritical
      Exit Sub
    End If
    db.BeginTrans
    db.Execute "DELETE FROM [C_SUPPLIER_INVENTORY] WHERE INVENTORY='" & txtFields(0).Text & "'"
    With adoPrimaryRS
      .Delete
      .MoveNext
      If .EOF Then .MoveLast
    End With
    db.CommitTrans
  End If
  '
  RefreshIndeks optFilter(0).Value
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  '
  ' Validasi Field
  If txtFields(1).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(1).SetFocus
    Exit Sub
  End If
  '
  If adoPrimaryRS!IsResep = False Then
    For iCounter = 2 To 8 Step 2
      If txtFields(iCounter) = "" Then
        MsgBox "Field ini dibutuhkan", vbCritical
        txtFields(iCounter).SetFocus
        Exit Sub
      End If
    Next
  End If
  '
  If txtFields(14).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(14).SetFocus
    Exit Sub
  End If
  If txtFields(15).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    txtFields(15).SetFocus
    Exit Sub
  End If
  '
  txtFields(0).Enabled = True
  adoPrimaryRS.UpdateBatch adAffectAll
  db.CommitTrans

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  If mbEditFlag = True Then
    For iCounter = 0 To 1
      txtFields(iCounter).Enabled = True
    Next
    For iCounter = 2 To 8 Step 2
      txtFields(iCounter).Enabled = True
    Next
    txtFields(15).Enabled = True
  End If
  '
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  ProsesDetail
  RefreshIndeks optFilter(0).Value
  StatusFrame False
  inDML = False
  '
  Exit Sub
  '
UpdateErr:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
  '
  txtFields(0).Enabled = True
  '
  SetButtons True
  '
  If mbEditFlag = True Then
    For iCounter = 0 To 1
      txtFields(iCounter).Enabled = True
    Next
    For iCounter = 2 To 8 Step 2
      txtFields(iCounter).Enabled = True
    Next
    txtFields(15).Enabled = True
  End If
  '
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
  RefreshIndeks optFilter(0).Value
  StatusFrame False
  inDML = False
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

Private Sub RefreshIndeks(bolRaw As Boolean)
  '
  PosisiRecord = adoPrimaryRS.AbsolutePosition
  '
  Set rsIndeks = New ADODB.Recordset
  If bolRaw = True Then
    rsIndeks.Open "Select NamaInventory FROM [C_Inventory] Where IsResep=False Order By NamaInventory", db, adOpenStatic, adLockOptimistic
  Else
    rsIndeks.Open "Select NamaInventory FROM [C_Inventory] Where IsResep=True Order By NamaInventory", db, adOpenStatic, adLockOptimistic
  End If
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
    lstIndeks.AddItem rsIndeks!NamaInventory
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
    lstCari.AddItem rsCari!NamaInventory
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  If optFilter(0).Value = True Then
    rsIndeks.Open "Select NamaInventory, IsResep FROM [C_Inventory] WHERE NamaInventory LIKE '%" & txtIndeks.Text & "%' and IsResep=False Order By NamaInventory", db, adOpenStatic, adLockOptimistic
  Else
    rsIndeks.Open "Select NamaInventory, IsResep FROM [C_Inventory] WHERE NamaInventory LIKE '%" & txtIndeks.Text & "%' and IsResep=True Order By NamaInventory", db, adOpenStatic, adLockOptimistic
  End If
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
  adoPrimaryRS.Find "NamaInventory ='" & lstIndeks.Text & "'"
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
      rsCari.Open "Select * FROM [C_Inventory] WHERE Inventory LIKE '%" & txtCari.Text & "%' Order By Inventory", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_Inventory] WHERE NamaInventory LIKE '%" & txtCari.Text & "%' Order By NamaInventory", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "NamaInventory ='" & lstCari.Text & "'"
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
        If JumlahRecord("Select Inventory From [C_Inventory] Where Inventory='" & txtFields(0).Text & "'", db) <> 0 Then
          MsgBox "Kode Inventory " & txtFields(0).Text & " telah digunakan", vbCritical
          txtFields(0).Text = ""
          txtFields(0).SetFocus
          Exit Sub
        End If
        If JumlahRecord("Select Inventory From [C_IStockcard] Where Inventory='" & txtFields(0).Text & "'", db) <> 0 Then
          MsgBox "Kode Inventory " & txtFields(0).Text & " telah digunakan", vbCritical
          txtFields(0).Text = ""
          txtFields(0).SetFocus
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case 2
        sKontrolAktif = "AI_1"
        Dim rsGudang As New ADODB.Recordset
        rsGudang.Open "Select Gudang, NamaGudang from [C_Gudang] Where Gudang='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsGudang.RecordCount <> 0 Then
          txtFields(3).Text = rsGudang!NamaGudang
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(3).Text = ""
          frmCari.Show vbModal
          txtFields(2).SetFocus
        End If
        rsGudang.Close
        sKontrolAktif = ""
      Case 4
        sKontrolAktif = "AI_2"
        Dim rsKategori As New ADODB.Recordset
        rsKategori.Open "Select Kategori, NamaKategori from [C_IKategori] Where Kategori='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsKategori.RecordCount <> 0 Then
          txtFields(5).Text = rsKategori!NamaKategori
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(5).Text = ""
          frmCari.Show vbModal
          txtFields(4).SetFocus
        End If
        rsKategori.Close
        sKontrolAktif = ""
      Case 6
        sKontrolAktif = "AI_3"
        Dim rsSatuan1 As New ADODB.Recordset
        rsSatuan1.Open "Select Satuan, NamaSatuan from [C_Satuan] Where Satuan='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsSatuan1.RecordCount <> 0 Then
          txtFields(7).Text = rsSatuan1!NamaSatuan
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(7).Text = ""
          frmCari.Show vbModal
          txtFields(6).SetFocus
        End If
        rsSatuan1.Close
        For iCounter = 0 To 3
          lblSatuan(iCounter).Caption = txtFields(7).Text
        Next
        sKontrolAktif = ""
      Case 8
        sKontrolAktif = "AI_4"
        Dim rsSatuan2 As New ADODB.Recordset
        rsSatuan2.Open "Select Satuan, NamaSatuan from [C_Satuan] Where Satuan='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsSatuan2.RecordCount <> 0 Then
          txtFields(9).Text = rsSatuan2!NamaSatuan
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(9).Text = ""
          frmCari.Show vbModal
          txtFields(8).SetFocus
        End If
        rsSatuan2.Close
        lblSatuan(4).Caption = txtFields(9).Text
        lblSatuan(6).Caption = txtFields(9).Text
        sKontrolAktif = ""
      Case 14
        If Not IsNumeric(txtFields(14).Text) Then
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case 15
        If Not IsNumeric(txtFields(15).Text) Then
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub cmdAddSupp_Click()
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Inventory kosong", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS!IsResep = True Then
    MsgBox "Item ini merupakan Standar Resep", vbCritical
    Exit Sub
  End If
  '
  If inDML Then
    MsgBox "Proses penambahan atau pengeditan data sedang berlangsung", vbCritical
    tabInventory.Tab = 0
    txtFields(1).SetFocus
    Exit Sub
  End If
  '
  frmCAddSupplier.Caption = " Menambah Supplier " & adoPrimaryRS!NamaInventory
  frmCAddSupplier.Show vbModal
  ProsesDetail
End Sub

Private Sub cmdDelSupp_Click()
  On Error Resume Next
  '
  ' Cek kondisi
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Inventory kosong", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS!IsResep = True Then
    MsgBox "Item ini merupakan Standar Resep", vbCritical
    Exit Sub
  End If
  '
  If inDML Then
    MsgBox "Proses penambahan atau pengeditan data sedang berlangsung", vbCritical
    tabInventory.Tab = 0
    txtFields(1).SetFocus
    Exit Sub
  End If
  If rsDetail.RecordCount = 0 Then
    MsgBox "Data Supplier kosong", vbCritical
    Exit Sub
  End If
  '
  If MsgBox("Hapus Supplier '" & rsDetail!NamaSupplier & "' ?", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
  Else
    '
    With rsDetail
      .Delete
      .MoveNext
      If .EOF Then .MoveLast
    End With
  End If
  '
  ProsesDetail
  '
End Sub

Private Sub cmdLookUp_Click(Index As Integer)
  Select Case Index
    Case 0
      sKontrolAktif = "AI_1"
      frmCari.Show vbModal
      txtFields(2).SetFocus
      If txtFields(2).Text <> "" Then SendKeys "{ENTER}"
    Case 1
      sKontrolAktif = "AI_2"
      frmCari.Show vbModal
      txtFields(4).SetFocus
      If txtFields(4).Text <> "" Then SendKeys "{ENTER}"
    Case 2
      sKontrolAktif = "AI_3"
      frmCari.Show vbModal
      txtFields(6).SetFocus
      If txtFields(6).Text <> "" Then SendKeys "{ENTER}"
    Case 3
      sKontrolAktif = "AI_4"
      frmCari.Show vbModal
      txtFields(8).SetFocus
      If txtFields(8).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub

Private Sub optFilter_Click(Index As Integer)
  If optFilter(0).Value = True Then
    RefreshIndeks True
  Else
    RefreshIndeks False
  End If
End Sub
