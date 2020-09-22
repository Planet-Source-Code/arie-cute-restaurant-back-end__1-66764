VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCBuatResepModifier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Permintaan Pembuatan Resep"
   ClientHeight    =   6315
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
   ScaleHeight     =   6315
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5355
      Left            =   3480
      TabIndex        =   24
      Top             =   60
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   4635
         Left            =   60
         ScaleHeight     =   4635
         ScaleWidth      =   7695
         TabIndex        =   25
         Top             =   660
         Width           =   7695
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "..."
            Height          =   315
            Index           =   1
            Left            =   6960
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1020
            Width           =   375
         End
         Begin VB.CommandButton cmdLookUp 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   6960
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   660
            Width           =   375
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            DataField       =   "Jumlah"
            Height          =   315
            Index           =   6
            Left            =   3000
            TabIndex        =   12
            Top             =   1380
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Resep"
            Height          =   315
            Index           =   4
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaResep"
            Height          =   315
            Index           =   5
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1020
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "NamaKaryawan"
            Height          =   315
            Index           =   3
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   660
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "Karyawan"
            Height          =   315
            Index           =   2
            Left            =   3000
            TabIndex        =   10
            Top             =   660
            Width           =   1215
         End
         Begin VB.CommandButton cmdKomposisi 
            Caption         =   "&Hitung Komposisi"
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
            Left            =   5040
            Picture         =   "frmCBuatResepModifier.frx":0000
            TabIndex        =   13
            Top             =   1860
            Width           =   2595
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
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "99/99/9999"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00FFFFFF&
            DataField       =   "KodeBuat"
            Height          =   315
            Index           =   0
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "1234567890"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "NilaiBuat"
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
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3960
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dgPI 
            Height          =   1635
            Left            =   60
            TabIndex        =   14
            Top             =   2280
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
            Caption         =   "Komposisi Resep"
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "NamaInventory"
               Caption         =   "Nama Bahan Resep"
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
               DataField       =   "HargaPerUnit"
               Caption         =   "Harga Per Unit"
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
               DataField       =   "JumlahKonsumsi"
               Caption         =   "# Konsumsi"
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
               MarqueeStyle    =   2
               ScrollBars      =   2
               AllowSizing     =   0   'False
               RecordSelectors =   0   'False
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   2445.166
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1365.165
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1470.047
               EndProperty
            EndProperty
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah diminta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1260
            TabIndex        =   38
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label lblSatuan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
            DataField       =   "Satuan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3840
            TabIndex        =   37
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Resep"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1260
            TabIndex        =   36
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Karyawan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   34
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Buat"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4560
            TabIndex        =   30
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Pembuatan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   1260
            TabIndex        =   29
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Permintaan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   4200
            TabIndex        =   28
            Top             =   4020
            Width           =   1545
         End
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCBuatResepModifier.frx":030A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan Pembuatan Resep"
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
         TabIndex        =   31
         Top             =   180
         Width           =   3885
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
      Height          =   6255
      Left            =   60
      TabIndex        =   21
      Top             =   60
      Width           =   3435
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
         Picture         =   "frmCBuatResepModifier.frx":0BD4
         TabIndex        =   2
         Top             =   5820
         Width           =   3315
      End
      Begin TabDlg.SSTab TabIndek 
         Height          =   5775
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   10186
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
         TabPicture(0)   =   "frmCBuatResepModifier.frx":0EDE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCBuatResepModifier.frx":0EFA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblLabels(11)"
         Tab(1).Control(1)=   "txtCari"
         Tab(1).Control(2)=   "cboBy"
         Tab(1).Control(3)=   "cmdTampil"
         Tab(1).Control(4)=   "lstCari"
         Tab(1).ControlCount=   5
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
            Height          =   3570
            ItemData        =   "frmCBuatResepModifier.frx":0F16
            Left            =   -74760
            List            =   "frmCBuatResepModifier.frx":0F1D
            TabIndex        =   20
            Top             =   1860
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   19
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCBuatResepModifier.frx":0F2A
            Left            =   -74760
            List            =   "frmCBuatResepModifier.frx":0F31
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   720
            Width           =   2835
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   -74760
            TabIndex        =   18
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
            TabIndex        =   23
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
      TabIndex        =   32
      Top             =   5460
      Width           =   11235
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCBuatResepModifier.frx":0F46
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCBuatResepModifier.frx":1250
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCBuatResepModifier.frx":155A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCBuatResepModifier.frx":1864
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCBuatResepModifier.frx":1B6E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCBuatResepModifier.frx":1E78
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCBuatResepModifier.frx":2182
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCBuatResepModifier.frx":248C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCBuatResepModifier"
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

Private Function RekomendasiResep(KodeResep As String) As Long
  '
  ' Hitung Rekomendasi
  Dim Rekomendasi As Long
  Dim Batas() As Long
  Dim JumlahKonsumsi() As Double
  Dim Hasil() As Double
  Dim JumlahFisik As Double
  '
  Dim rsResep As New ADODB.Recordset
  rsResep.Open "Select Resep, JumlahFisik From [C_Resep] Where Resep='" & KodeResep & "'", db, adOpenStatic, adLockReadOnly
  JumlahFisik = rsResep!JumlahFisik
  rsResep.Close
  Set rsResep = Nothing
  
  Dim rsBahanResep As New ADODB.Recordset
  rsBahanResep.Open "Select * From [C_ResepBahan] Where Resep='" & KodeResep & "'", db, adOpenStatic, adLockReadOnly
  If rsBahanResep.RecordCount <> 0 Then
    ReDim Batas(rsBahanResep.RecordCount) As Long
    ReDim JumlahKonsumsi(rsBahanResep.RecordCount) As Double
    ReDim Hasil(rsBahanResep.RecordCount) As Double
    '
    rsBahanResep.MoveFirst
    iCounter = 0
    Do While Not rsBahanResep.EOF
      Dim rsItem As New ADODB.Recordset
      rsItem.Open "Select Inventory, JumlahItem From [C_Inventory] Where Inventory='" & rsBahanResep!Inventory & "'", db, adOpenStatic, adLockReadOnly
      '
      iCounter = iCounter + 1
      JumlahKonsumsi(iCounter) = rsBahanResep!JumlahKonsumsi / JumlahFisik
      Batas(iCounter) = rsItem!JumlahItem
      rsItem.Close
      Set rsItem = Nothing
      rsBahanResep.MoveNext
    Loop
  End If
  '
  Dim CariRekomendasi As Boolean
  Rekomendasi = 0
  CariRekomendasi = True
  Do While CariRekomendasi
    RekomendasiResep = Rekomendasi
    Rekomendasi = Rekomendasi + 1
    For iCounter = 1 To rsBahanResep.RecordCount
      Hasil(iCounter) = Rekomendasi * JumlahKonsumsi(iCounter)
      If Hasil(iCounter) >= Batas(iCounter) Then CariRekomendasi = False
    Next
  Loop
  '
  rsBahanResep.Close
  Set rsBahanResep = Nothing
  '
  RekomendasiResep = Rekomendasi
  '
End Function

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
  strSQL = "SELECT * FROM [C_BuatResepSpesialBahan] WHERE KodeBuat='" & strKey & "' Order By Inventory"
  rsDetail.Open strSQL, db, adOpenStatic, adLockOptimistic
  rsDetail.Requery
  '
  Set dgPI.DataSource = rsDetail
  If adoPrimaryRS.RecordCount <> 0 Then
    dgPI.Caption = "Komposisi bahan resep berdasarkan kode " & adoPrimaryRS!KodeBuat
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
  adoPrimaryRS.Open "select * from [C_BuatResepSpesial] Order by KodeBuat", db, adOpenStatic, adLockOptimistic
  
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set lblSatuan.DataSource = adoPrimaryRS
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
      rsCari.Open "Select * FROM [C_BuatResepSpesial] WHERE KodeBuat LIKE '%" & txtCari.Text & "%' Order By KodeBuat", db, adOpenStatic, adLockOptimistic
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
        If JumlahRecord("Select KodeBuat From [C_BuatResepSpesial] Where KodeBuat='" & txtFields(0).Text & "'", db) <> 0 Then
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
        sKontrolAktif = "ABRM_1"
        Dim rsKaryawan As New ADODB.Recordset
        rsKaryawan.Open "Select Karyawan, NamaKaryawan from [Karyawan] Where Karyawan='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsKaryawan.RecordCount <> 0 Then
          txtFields(3).Text = rsKaryawan!NamaKaryawan
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(3).Text = ""
          frmCari.Show vbModal
          txtFields(2).SetFocus
        End If
        rsKaryawan.Close
        sKontrolAktif = ""
      Case 4
        sKontrolAktif = "ABRM_2"
        Dim rsResep As New ADODB.Recordset
        rsResep.Open "Select Resep, NamaResep, NamaSatuan from [C_Resep] Where Resep='" & txtFields(Index).Text & "'", db, adOpenStatic, adLockReadOnly
        If rsResep.RecordCount <> 0 Then
          txtFields(5).Text = rsResep!NamaResep
          lblSatuan.Caption = rsResep!NamaSatuan
          KeyAscii = 0
          SendKeys "{TAB}"
        Else
          txtFields(5).Text = ""
          lblSatuan.Caption = "Satuan"
          frmCari.Show vbModal
          txtFields(4).SetFocus
        End If
        rsResep.Close
        sKontrolAktif = ""
      Case 6
        If Not IsNumeric(txtFields(6).Text) Then
          Exit Sub
        End If
        If txtFields(6).Text > RekomendasiResep(txtFields(4).Text) Then
          sPesan = "Dengan persediaan inventory yang ada hanya dapat dibuat" & vbCrLf
          sPesan = sPesan & "resep sebanyak " & RekomendasiResep(txtFields(4).Text) & " " & lblSatuan.Caption
          MsgBox sPesan, vbInformation
          txtFields(6).Text = 1
          Exit Sub
        End If
        cmdKomposisi_Click
        cmdUpdate.SetFocus
      Case Else
        SendKeys "{TAB}"
    End Select
  End If
End Sub

Private Sub cmdKomposisi_Click()
  '
  If (txtFields(0).Text = "") Then
    MsgBox "Nomor Kode Pembuatan Resep tidak boleh kosong", vbCritical
    txtFields(0).SetFocus
    Exit Sub
  End If
  If (txtFields(4).Text = "") Then
    MsgBox "Nomor Kode Resep tidak boleh kosong", vbCritical
    txtFields(4).SetFocus
    Exit Sub
  End If
  If (txtFields(6).Text = "") Then
    txtFields(6).Text = 1
  End If
  '
  Dim sngJumlahFisik As Single
  Dim rsJumlahFisikResep As New ADODB.Recordset
  rsJumlahFisikResep.Open "Select Resep, JumlahFisik From [C_Resep] Where Resep='" & txtFields(4).Text & "'", db, adOpenStatic, adLockReadOnly
  sngJumlahFisik = rsJumlahFisikResep!JumlahFisik
  rsJumlahFisikResep.Close
  Set rsJumlahFisikResep = Nothing
  '
  Dim rsKomposisi As New ADODB.Recordset
  rsKomposisi.Open "Select Resep, Inventory, NamaInventory, Satuan, NamaSatuan, HargaPerUnit, JumlahKonsumsi From [C_ResepBahan] Where Resep='" & txtFields(4).Text & "'", db, adOpenStatic, adLockReadOnly
  If rsKomposisi.RecordCount <> 0 Then
    rsKomposisi.MoveFirst
    Do While Not rsKomposisi.EOF
      If JumlahRecord("Select KodeBuat, Inventory From [C_BuatResepSpesialBahan] Where KodeBuat='" & txtFields(0).Text & "' and Inventory='" & rsKomposisi!Inventory & "'", db) <> 0 Then
        db.Execute "Delete From [C_BuatResepSpesialBahan] Where KodeBuat='" & txtFields(0).Text & "' and Inventory='" & rsKomposisi!Inventory & "'"
      End If
      '
      Dim rsAddKomposisi As New ADODB.Recordset
      rsAddKomposisi.Open "C_BuatResepSpesialBahan", db, adOpenStatic, adLockOptimistic
      rsAddKomposisi.AddNew
      rsAddKomposisi!KodeBuat = txtFields(0).Text
      rsAddKomposisi!Inventory = rsKomposisi!Inventory
      rsAddKomposisi!NamaInventory = rsKomposisi!NamaInventory
      rsAddKomposisi!Satuan = rsKomposisi!Satuan
      rsAddKomposisi!NamaSatuan = rsKomposisi!NamaSatuan
      rsAddKomposisi!HargaPerUnit = rsKomposisi!HargaPerUnit
      rsAddKomposisi!JumlahKonsumsi = Format((rsKomposisi!JumlahKonsumsi / sngJumlahFisik) * txtFields(6).Text, "###,###.##")
      rsAddKomposisi!SubTotal = rsAddKomposisi!JumlahKonsumsi * rsKomposisi!HargaPerUnit
      rsAddKomposisi!JumlahKonsumsiAcc = 0
      rsAddKomposisi!SubTotalAcc = 0
      rsAddKomposisi.Update
      rsAddKomposisi.Close
      Set rsAddKomposisi = Nothing
      '
      rsKomposisi.MoveNext
    Loop
  End If
  rsKomposisi.Close
  Set rsKomposisi = Nothing
  '
  Dim rsJumlahBiaya As New ADODB.Recordset
  rsJumlahBiaya.Open "Select Sum(Subtotal) as Jumlah From [C_BuatResepSpesialBahan] Where KodeBuat='" & txtFields(0).Text & "'", db, adOpenStatic, adLockReadOnly
  txtFields(7).Text = rsJumlahBiaya!Jumlah
  rsJumlahBiaya.Close
  Set rsJumlahBiaya = Nothing
  '
  ProsesDetail
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
  '
  sSQL = "SELECT C_BuatResepSpesial.KodeBuat, C_BuatResepSpesial.NamaResep, C_BuatResepSpesial.Jumlah, C_BuatResepSpesial.Satuan, C_BuatResepSpesialBahan.NamaInventory, C_BuatResepSpesialBahan.NamaSatuan, C_BuatResepSpesialBahan.HargaPerUnit, C_BuatResepSpesialBahan.JumlahKonsumsi FROM [C_BuatResepSpesial] INNER JOIN [C_BuatResepSpesialBahan] ON C_BuatResepSpesial.KodeBuat = C_BuatResepSpesialBahan.KodeBuat WHERE (((C_BuatResepSpesial.KodeBuat)='" & txtFields(0).Text & "'))"

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
    .WindowTitle = "Form Permintaan Pembuatan Resep"
    .ReportFileName = App.Path & "\Report\Laporan Minta Buat Resep.Rpt"
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
      sKontrolAktif = "ABRM_1"
      frmCari.Show vbModal
      txtFields(2).SetFocus
      If txtFields(2).Text <> "" Then SendKeys "{ENTER}"
    Case 1
      sKontrolAktif = "ABRM_2"
      frmCari.Show vbModal
      txtFields(4).SetFocus
      If txtFields(4).Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub
