VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCProdukMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produk Menu"
   ClientHeight    =   7290
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
   ScaleHeight     =   7290
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   60
      TabIndex        =   34
      Top             =   60
      Width           =   3375
      Begin TabDlg.SSTab TabIndek 
         Height          =   7155
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   12621
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
         TabPicture(0)   =   "frmCProdukMenu.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCProdukMenu.frx":001C
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
            TabIndex        =   23
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCProdukMenu.frx":0038
            Left            =   -74760
            List            =   "frmCProdukMenu.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   720
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   24
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   4935
            ItemData        =   "frmCProdukMenu.frx":005C
            Left            =   -74760
            List            =   "frmCProdukMenu.frx":0063
            TabIndex        =   25
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
            Height          =   5910
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
            TabIndex        =   36
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   6315
      Left            =   3480
      TabIndex        =   27
      Top             =   60
      Width           =   7935
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   60
         ScaleHeight     =   5595
         ScaleWidth      =   7815
         TabIndex        =   28
         Top             =   660
         Width           =   7815
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
            Index           =   10
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3660
            Width           =   1635
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "JumlahBiaya"
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
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3300
            Width           =   1635
         End
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
            Left            =   1920
            TabIndex        =   13
            Text            =   "100.00"
            Top             =   3660
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "Biaya"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   4020
            Width           =   1635
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataField       =   "HargaJual"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   3180
            TabIndex        =   14
            Top             =   4740
            Width           =   1635
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "GrossMargin"
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
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   5100
            Width           =   1635
         End
         Begin VB.CommandButton cmdMarkUp 
            Caption         =   "&Faktor Mark Up"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            Picture         =   "frmCProdukMenu.frx":0070
            TabIndex        =   39
            Top             =   4380
            Width           =   1515
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "FaktorMarkUp"
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
            Left            =   3180
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   4380
            Width           =   915
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF8080&
            Caption         =   "&Non Aktif"
            DataField       =   "NonAktif"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5580
            TabIndex        =   18
            Top             =   4980
            Width           =   1635
         End
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   6240
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Menu"
            Height          =   315
            Index           =   0
            Left            =   2820
            MaxLength       =   14
            TabIndex        =   10
            Text            =   "12345678901234"
            Top             =   120
            Width           =   1635
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "NamaMenu"
            Height          =   315
            Index           =   1
            Left            =   2820
            TabIndex        =   11
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Departemen"
            Height          =   315
            Index           =   2
            Left            =   2820
            TabIndex        =   12
            Top             =   840
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaDepartemen"
            Height          =   315
            Index           =   3
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   840
            Width           =   2655
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
            Left            =   5400
            Picture         =   "frmCProdukMenu.frx":037A
            TabIndex        =   15
            Top             =   3300
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
            Left            =   5400
            Picture         =   "frmCProdukMenu.frx":0684
            TabIndex        =   16
            Top             =   3720
            Width           =   2055
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
            Left            =   5400
            Picture         =   "frmCProdukMenu.frx":098E
            TabIndex        =   17
            Top             =   4140
            Width           =   2055
         End
         Begin MSDataGridLib.DataGrid dgBahanMenu 
            Height          =   1875
            Left            =   120
            TabIndex        =   29
            Top             =   1320
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
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2670.236
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   824.882
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1425.26
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% ="
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   11
            Left            =   2760
            TabIndex        =   50
            Top             =   3720
            Width           =   330
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   10
            Left            =   2580
            TabIndex        =   49
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Persentase Terbuang"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   48
            Top             =   3720
            Width           =   1545
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Biaya"
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
            Index           =   12
            Left            =   1500
            TabIndex        =   45
            Top             =   4080
            Width           =   945
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Harga Jual"
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
            Index           =   4
            Left            =   1500
            TabIndex        =   43
            Top             =   4800
            Width           =   900
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gross Margin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   1500
            TabIndex        =   42
            Top             =   5160
            Width           =   1110
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Index           =   7
            Left            =   4200
            TabIndex        =   41
            Top             =   4440
            Width           =   195
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Menu"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1620
            TabIndex        =   32
            Top             =   180
            Width           =   795
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Menu"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   31
            Top             =   540
            Width           =   840
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departemen"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1620
            TabIndex        =   30
            Top             =   900
            Width           =   885
         End
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   240
         Picture         =   "frmCProdukMenu.frx":0C98
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produk Menu"
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
         TabIndex        =   33
         Top             =   180
         Width           =   1830
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
      TabIndex        =   26
      Top             =   6420
      Width           =   11355
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Hapus"
         Height          =   795
         Left            =   9660
         Picture         =   "frmCProdukMenu.frx":1562
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCProdukMenu.frx":186C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCProdukMenu.frx":1B76
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCProdukMenu.frx":1E80
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCProdukMenu.frx":218A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   7980
         Picture         =   "frmCProdukMenu.frx":2494
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Koreksi"
         Height          =   795
         Left            =   8820
         Picture         =   "frmCProdukMenu.frx":279E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10500
         Picture         =   "frmCProdukMenu.frx":2AA8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   9660
         Picture         =   "frmCProdukMenu.frx":2DB2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10500
         Picture         =   "frmCProdukMenu.frx":30BC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCProdukMenu"
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
  strSQL = "SELECT * FROM [C_MenuBahan] WHERE Menu='" & strKey & "'"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgBahanMenu.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgBahanMenu.Caption = "Bahan-bahan " & adoPrimaryRS!NamaMenu
  End If
  dgBahanMenu.ReBind
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
  adoPrimaryRS.Open "select * from [C_Menu] Order by NamaMenu", db, adOpenStatic, adLockOptimistic
  
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
  For iCounter = 4 To 10
    txtFields(iCounter).Text = 0
  Next
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Menu kosong", vbCritical
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
    MsgBox "Data Produk Menu kosong", vbCritical
    Exit Sub
  End If
  '
  If MsgBox("Hapus Produk Menu '" & adoPrimaryRS!NamaMenu & "' ?", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
  Else
    '
    db.BeginTrans
    db.Execute "DELETE FROM [C_MENUBAHAN] WHERE MENU='" & txtFields(0).Text & "'"
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
  '
  txtFields(0).Enabled = True
  '
  ' Validasi Field
  For iCounter = 0 To 2
    If txtFields(iCounter) = "" Then
      MsgBox "Field ini dibutuhkan", vbCritical
      txtFields(iCounter).SetFocus
      Exit Sub
    End If
  Next
  If (txtFields(6).Text = "") Or (Val(txtFields(6).Text) = 0) Then
    MsgBox "Tentukan Harga Penjualan terlebih dahulu", vbCritical
    txtFields(6).SetFocus
    Exit Sub
  End If
  '
  If rsDetail.RecordCount = 0 Then
    MsgBox "Bahan pembuat sebuah produk menu tidak boleh kosong", vbCritical
    cmdAddBahan.SetFocus
    Exit Sub
  End If
  '
  HitungCostBenefit
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
  rsIndeks.Open "Select NamaMenu FROM [C_Menu] Order By NamaMenu", db, adOpenStatic, adLockOptimistic
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
    lstIndeks.AddItem rsIndeks!NamaMenu
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
    lstCari.AddItem rsCari!NamaMenu
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select NamaMenu FROM [C_Menu] WHERE NamaMenu LIKE '%" & txtIndeks.Text & "%' Order By NamaMenu", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "NamaMenu ='" & lstIndeks.Text & "'"
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
      rsCari.Open "Select * FROM [C_Menu] WHERE Menu LIKE '%" & txtCari.Text & "%' Order By Menu", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_Menu] WHERE NamaMenu LIKE '%" & txtCari.Text & "%' Order By NamaMenu", db, adOpenStatic, adLockOptimistic
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
  adoPrimaryRS.Find "NamaMenu ='" & lstCari.Text & "'"
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
        If Not IsNumeric(txtFields(0).Text) Then
          txtFields(0).Text = ""
          Exit Sub
        End If
        txtFields(0).Enabled = True
        If Len(txtFields(0).Text) < 14 Then
          txtFields(0).Text = Format$(txtFields(0).Text, "0000000000000#")
        End If
        If JumlahRecord("SELECT MENU FROM [C_MENU] WHERE MENU='" & txtFields(0).Text & "'", db) > 0 Then
          MsgBox "Kode telah digunakan oleh produk menu lain", vbCritical
          txtFields(0).Text = ""
          txtFields(0).SetFocus
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case 2
        sKontrolAktif = "AM_1"
        Dim rsDepartemen As New ADODB.Recordset
        rsDepartemen.Open "Select Departemen, NamaDepartemen from [C_MDepartemen] Where Departemen='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsDepartemen.RecordCount <> 0 Then
          txtFields(3).Text = rsDepartemen!NamaDepartemen
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(3).Text = ""
          frmCari.Show vbModal
          txtFields(2).SetFocus
        End If
        rsDepartemen.Close
        sKontrolAktif = ""
      Case 6
        If Not IsNumeric(txtFields(6).Text) Then
          Exit Sub
        End If
        If CCur(txtFields(6).Text = 0) Then
          MsgBox "Harga Penjualan Produk menu tidak boleh bernilai nol", vbCritical
          Exit Sub
        End If
        txtFields(10).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
        txtFields(4).Text = CCur(txtFields(8).Text) + CCur(txtFields(10).Text)
        txtFields(7).Text = CCur(txtFields(6).Text) - CCur(txtFields(4).Text)
        If CCur(txtFields(4).Text) = 0 Then
          txtFields(5).Text = 0
        Else
          txtFields(5).Text = Format((CCur(txtFields(7).Text) / CCur(txtFields(4).Text)) * 100, "#0.00")
        End If
        '
        If (txtFields(7).Text / txtFields(6).Text) * 100 < 50 Then
          MsgBox "Margin Keuntungan (" & Format((txtFields(7).Text / txtFields(6).Text) * 100, "###,##0.00") & " %) kurang dari 50 %", vbExclamation
        End If
        SendKeys "{TAB}"
      Case 9
        If Not IsNumeric(txtFields(9).Text) Then
          Exit Sub
        End If
        SendKeys "{TAB}"
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub HitungCostBenefit()
  txtFields(10).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
  txtFields(4).Text = CCur(txtFields(8).Text) + CCur(txtFields(10).Text)
  txtFields(7).Text = CCur(txtFields(6).Text) - CCur(txtFields(4).Text)
  If CCur(txtFields(4).Text) = 0 Then
    txtFields(5).Text = 0
  Else
    txtFields(5).Text = Format((CCur(txtFields(7).Text) / CCur(txtFields(4).Text)) * 100, "#0.00")
  End If
  '
  If (txtFields(7).Text / txtFields(6).Text) * 100 < 50 Then
    ' Tambahkan Item ke File Price Alert
    If (txtFields(7).Text / txtFields(6).Text) * 100 < 0 Then
      MsgBox "Produk Menu ini memiliki kerugian sebesar " & Format(Abs((txtFields(7).Text / txtFields(6).Text) * 100), "###,##0.00") & " %", vbCritical
    Else
      MsgBox "Margin Keuntungan (" & Format((txtFields(7).Text / txtFields(6).Text) * 100, "###,##0.00") & " %) kurang dari 50 %", vbExclamation
    End If
    If JumlahRecord("Select Menu from [C_PriceAlert] Where Menu = '" & txtFields(0).Text & "'", db) = 0 Then
      Dim rsAlert As New ADODB.Recordset
      rsAlert.Open "C_PriceAlert", db, adOpenStatic, adLockOptimistic
      rsAlert.AddNew
      rsAlert!Menu = txtFields(0).Text
      rsAlert!NamaMenu = txtFields(1).Text
      rsAlert!Untung = txtFields(7).Text
      rsAlert!Harga = txtFields(6).Text
      rsAlert!PersenUntungOfSales = (txtFields(7).Text / txtFields(6).Text) * 100
      rsAlert.Update
      rsAlert.Close
      Set rsAlert = Nothing
    Else
      db.Execute "Update [C_PriceAlert] Set Untung=" & CCur(txtFields(7).Text) & ", Harga=" & CCur(txtFields(6).Text) & ", PersenUntungOfSales=" & CSng((txtFields(7).Text / txtFields(6).Text) * 100) & " Where Menu='" & txtFields(0).Text & "'"
    End If
  Else
    db.Execute "Delete From [C_PriceAlert] Where Menu='" & txtFields(0).Text & "'"
  End If
  '
  SendKeys "{TAB}"
End Sub

Private Sub cmdAddBahan_Click()
  '
  frmCProdukMenu.Tag = "1"
  frmCProdukMenuBahan.Show vbModal
  ProsesDetail
  'HitungCostBenefit
  '
End Sub

Private Sub cmdEditBahan_Click()
  '
  If rsDetail.RecordCount = 0 Then
    MsgBox "Data Bahan Menu kosong", vbCritical
    Exit Sub
  End If
  '
  frmCProdukMenu.Tag = "2"
  txtFields(8).Text = CCur(txtFields(8).Text) - CCur(dgBahanMenu.Columns(4).Text)
  txtFields(10).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
  txtFields(4).Text = CCur(txtFields(8).Text) + CCur(txtFields(10).Text)
  txtFields(7).Text = CCur(txtFields(6).Text) - CCur(txtFields(4).Text)
  If CCur(txtFields(4).Text) = 0 Then
    txtFields(5).Text = 0
  Else
    txtFields(5).Text = Format((CCur(txtFields(7).Text) / CCur(txtFields(4).Text)) * 100, "#0.00")
  End If
  With frmCProdukMenuBahan
    .txtFields(0).Text = dgBahanMenu.Columns(0).Text
    .txtFields(0).Enabled = False
    .txtFields(1).Text = dgBahanMenu.Columns(1).Text
    .txtFields(2).Text = dgBahanMenu.Columns(2).Text
    .txtFields(3).Text = dgBahanMenu.Columns(4).Text
    .cmdLookUp(0).Enabled = False
    .Show vbModal
  End With
  ProsesDetail
  'HitungCostBenefit
  '
End Sub

Private Sub cmdHapusBahan_Click()
  '
  If rsDetail.RecordCount = 0 Then
    MsgBox "Data Bahan Menu kosong", vbCritical
    Exit Sub
  End If
  '
  sPesan = dgBahanMenu.Columns(1).Text & " akan dihapus dari daftar bahan Menu " & txtFields(1).Text & "." & vbCrLf
  sPesan = sPesan & "Anda Yakin ?"
  If MsgBox(sPesan, vbExclamation + vbYesNo) = vbYes Then
    txtFields(8).Text = CCur(txtFields(8).Text) - CCur(dgBahanMenu.Columns(4).Text)
    txtFields(10).Text = CCur((txtFields(9).Text / 100) * txtFields(8).Text)
    txtFields(4).Text = CCur(txtFields(8).Text) + CCur(txtFields(10).Text)
    txtFields(7).Text = CCur(txtFields(6).Text) - CCur(txtFields(4).Text)
    If CCur(txtFields(4).Text) = 0 Then
      txtFields(5).Text = 0
    Else
      txtFields(5).Text = Format((CCur(txtFields(7).Text) / CCur(txtFields(4).Text)) * 100, "#0.00")
    End If
    db.Execute "DELETE FROM [C_MenuBahan] WHERE Menu='" & txtFields(0).Text & "' and Inventory = '" & dgBahanMenu.Columns(0).Text & "'"
    HitungCostBenefit
    ProsesDetail
  End If
  '
  'HitungCostBenefit
  ProsesDetail
End Sub

Private Sub cmdMarkUp_Click()
  frmCPMarkup.Show vbModal
End Sub

Private Sub cmdLookUp_Click(Index As Integer)
  Select Case Index
    Case 0
      sKontrolAktif = "AM_1"
      frmCari.Show vbModal
      txtFields(2).SetFocus
      If txtFields(2).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub
