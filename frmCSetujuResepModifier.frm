VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCSetujuResepModifier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Persetujuan Pembuatan Resep"
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
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   60
      TabIndex        =   32
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
         Picture         =   "frmCSetujuResepModifier.frx":0000
         TabIndex        =   2
         Top             =   5820
         Width           =   3315
      End
      Begin TabDlg.SSTab TabIndek 
         Height          =   5775
         Left            =   0
         TabIndex        =   33
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
         TabPicture(0)   =   "frmCSetujuResepModifier.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCSetujuResepModifier.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblLabels(11)"
         Tab(1).Control(1)=   "lstCari"
         Tab(1).Control(2)=   "cmdTampil"
         Tab(1).Control(3)=   "cboBy"
         Tab(1).Control(4)=   "txtCari"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   -74760
            TabIndex        =   10
            Top             =   1080
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCSetujuResepModifier.frx":0342
            Left            =   -74760
            List            =   "frmCSetujuResepModifier.frx":034C
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   720
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   11
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   3570
            ItemData        =   "frmCSetujuResepModifier.frx":038E
            Left            =   -74760
            List            =   "frmCSetujuResepModifier.frx":0395
            TabIndex        =   12
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
            TabIndex        =   34
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5355
      Left            =   3480
      TabIndex        =   13
      Top             =   60
      Width           =   7815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   4635
         Left            =   60
         ScaleHeight     =   4635
         ScaleWidth      =   7695
         TabIndex        =   14
         Top             =   660
         Width           =   7695
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
            Index           =   8
            Left            =   5940
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   3720
            Width           =   1455
         End
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
            Left            =   5940
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "999,999,999.00"
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BackColor       =   &H00E0E0E0&
            DataField       =   "KodeBuat"
            Height          =   315
            Index           =   0
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            DataField       =   "Jumlah"
            Height          =   315
            Index           =   6
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1380
            Width           =   675
         End
         Begin MSDataGridLib.DataGrid dgPI 
            Height          =   1755
            Left            =   60
            TabIndex        =   23
            Top             =   1920
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3096
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
            Caption         =   "Komposisi Resep"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "NamaInventory"
               Caption         =   "Nama Bahan Resep"
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
               DataField       =   "HargaPerUnit"
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
               DataField       =   "JumlahKonsumsi"
               Caption         =   "# Minta"
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
               DataField       =   "JumlahKonsumsiAcc"
               Caption         =   "# Setuju"
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
               DataField       =   "SubTotalAcc"
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
            Caption         =   "Total Nilai Permintaan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   4260
            TabIndex        =   39
            Top             =   3780
            Width           =   1545
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Nilai Pembuatan"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   12
            Left            =   4260
            TabIndex        =   30
            Top             =   4140
            Width           =   1545
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
            Caption         =   "Tanggal Buat"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   9
            Left            =   4560
            TabIndex        =   28
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
            TabIndex        =   27
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
            Left            =   1260
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   1440
            Width           =   510
         End
         Begin VB.Label lblInv 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah diminta"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   1260
            TabIndex        =   24
            Top             =   1440
            Width           =   1050
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Persetujuan Pembuatan Resep"
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
         Width           =   3915
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCSetujuResepModifier.frx":03A2
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
      TabIndex        =   35
      Top             =   5460
      Width           =   11235
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   5160
         Picture         =   "frmCSetujuResepModifier.frx":0C6C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   4320
         Picture         =   "frmCSetujuResepModifier.frx":0F76
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   3480
         Picture         =   "frmCSetujuResepModifier.frx":1280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCSetujuResepModifier.frx":158A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCSetujuResepModifier.frx":1894
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   10380
         Picture         =   "frmCSetujuResepModifier.frx":1B9E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "P&roses"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCSetujuResepModifier.frx":1EA8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Setuju"
         Height          =   795
         Left            =   9540
         Picture         =   "frmCSetujuResepModifier.frx":21B2
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmCSetujuResepModifier"
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

Dim bolResepNotAcc As Boolean

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
  adoPrimaryRS.Open "Select * from [C_BuatResepSpesial] Order by KodeBuat", db, adOpenStatic, adLockOptimistic
  
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
  On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Permintaan Pembuatan Resep kosong", vbCritical
    Exit Sub
  End If
  If adoPrimaryRS!IsNowAcc = True Then
    MsgBox "Data Permintaan Pembuatan Resep telah diproses", vbInformation
    Exit Sub
  End If
  '
  bolResepNotAcc = True
  '
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

  ' Cek Resep boleh disetujui atau tidak
  rsDetail.MoveFirst
  Do While Not rsDetail.EOF
    If rsDetail!JumlahKonsumsiAcc = 0 Then bolResepNotAcc = True
    rsDetail.MoveNext
  Loop
  If bolResepNotAcc Then
    MsgBox "Permintaan Resep tidak dapat disetujui karena ada item yang bernilai nol", vbCritical
    Exit Sub
  End If
  '
  adoPrimaryRS!IsNowAcc = True
  adoPrimaryRS!NeedTransfer = True
  adoPrimaryRS.UpdateBatch adAffectAll
  '
  ' Isi File History
  Dim rsHistory As New ADODB.Recordset
  rsHistory.Open "History", db, adOpenStatic, adLockOptimistic
  rsHistory.AddNew
  rsHistory!KodeRef = txtFields(0).Text
  rsHistory!Tanggal = txtFields(1).Text
  rsHistory!Keterangan = "Pembuatan Resep " & txtFields(5).Text
  rsHistory!Nilai = txtFields(7).Text
  rsHistory!Jenis = "PR"
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
    rsHistoryDetail!Jumlah = rsDetail!JumlahKonsumsiAcc
    rsHistoryDetail!Satuan = rsDetail!NamaSatuan
    rsHistoryDetail!HargaSatuan = rsDetail!HargaPerUnit
    rsHistoryDetail!SubTotal = rsDetail!SubTotalAcc
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
    Dim bolResep As Boolean
    Dim Konversi, StockMin, JumlahItemBaru, QtyOnHandBaru, QtyOnHandKecilBaru As Single
    '
    rsUpdateInventory.Open "Select Inventory, NamaInventory, IsResep, FSatuanKecil, ReorderLevel, QtyOnHand, QtyOnHandKecil, SatuanBesar, JumlahItem From [C_Inventory] Where Inventory='" & NoInventory & "'", db, adOpenStatic, adLockOptimistic
    NamaItem = rsUpdateInventory!NamaInventory
    bolResep = rsUpdateInventory!IsResep
    Konversi = rsUpdateInventory!FSatuanKecil
    StockMin = rsUpdateInventory!ReorderLevel
    SatuanPesan = rsUpdateInventory!SatuanBesar
    '
    If bolResep Then
      QtyOnHandKecilBaru = rsUpdateInventory!QtyOnHandKecil - rsDetail!JumlahKonsumsiAcc
      JumlahItemBaru = rsUpdateInventory!JumlahItem - rsDetail!JumlahKonsumsiAcc
      rsUpdateInventory!QtyOnHandKecil = QtyOnHandKecilBaru
    Else
      JumlahItemBaru = rsUpdateInventory!JumlahItem - rsDetail!JumlahKonsumsiAcc
      QtyOnHandBaru = Int(JumlahItemBaru / Konversi)
      QtyOnHandKecilBaru = JumlahItemBaru - (QtyOnHandBaru * Konversi)
      rsUpdateInventory!QtyOnHand = QtyOnHandBaru
      rsUpdateInventory!QtyOnHandKecil = QtyOnHandKecilBaru
    End If
    
    rsUpdateInventory!JumlahItem = JumlahItemBaru
    rsUpdateInventory.Update
    rsUpdateInventory.Close
    Set rsUpdateInventory = Nothing
    '
    'JumlahItemBaru = rsUpdateInventory!JumlahItem - rsDetail!JumlahKonsumsiAcc
    'QtyOnHandBaru = Int(JumlahItemBaru / Konversi)
    'QtyOnHandKecilBaru = JumlahItemBaru - (QtyOnHandBaru * Konversi)
    '
    'rsUpdateInventory!JumlahItem = JumlahItemBaru
    'rsUpdateInventory!QtyOnHand = QtyOnHandBaru
    'rsUpdateInventory!QtyOnHandKecil = QtyOnHandKecilBaru
    'rsUpdateInventory.Update
    'rsUpdateInventory.Close
    'Set rsUpdateInventory = Nothing
    '
    If QtyOnHandBaru <= StockMin Then
      ' Tambahkan Item ke File Stock Alert
      If JumlahRecord("Select Inventory from [C_StockAlert] Where Inventory = '" & NoInventory & "'", db) = 0 Then
        Dim rsAlert As New ADODB.Recordset
        rsAlert.Open "C_StockAlert", db, adOpenStatic, adLockOptimistic
        rsAlert.AddNew
        rsAlert!Inventory = NoInventory
        rsAlert!NamaInventory = NamaItem
        rsAlert!ReorderLevel = StockMin
        rsAlert!QtyOnHand = QtyOnHandBaru
        rsAlert!Satuan = SatuanPesan
        rsAlert.Update
        rsAlert.Close
        Set rsAlert = Nothing
      Else
        db.Execute "Update [C_StockAlert] Set QtyOnHand=" & QtyOnHandBaru & " Where Inventory='" & NoInventory & "'"
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
    rsStockCardBahan!Keterangan = "Pembuatan Resep"
    rsStockCardBahan!PI = txtFields(0).Text
    rsStockCardBahan!UnitKeluar = rsDetail!JumlahKonsumsiAcc
    rsStockCardBahan.Update
    rsStockCardBahan.Close
    Set rsStockCardBahan = Nothing
    '
    rsDetail.MoveNext
  Loop
  '
  db.CommitTrans

  sPesan = "Permintaan Pembuatan Resep dengan Kode " & txtFields(0).Text & vbCrLf
  sPesan = sPesan & "telah disetujui"
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
  '
  adoPrimaryRS!JumlahAcc = 0
  adoPrimaryRS!IsNowAcc = True
  adoPrimaryRS!NeedTransfer = False
  adoPrimaryRS.UpdateBatch adAffectAll
  
  adoPrimaryRS.Requery
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  '
  ' Isi File History
  Dim rsHistory As New ADODB.Recordset
  rsHistory.Open "History", db, adOpenStatic, adLockOptimistic
  rsHistory.AddNew
  rsHistory!KodeRef = txtFields(0).Text
  rsHistory!Tanggal = txtFields(1).Text
  rsHistory!Keterangan = "Pembuatan Resep " & txtFields(5).Text & " (ditolak)"
  rsHistory!Nilai = txtFields(7).Text
  rsHistory!Jenis = "PR"
  rsHistory.Update
  rsHistory.Close
  Set rsHistory = Nothing
  '
  ' Isi File Detail History
  Dim rsHistoryDetail As New ADODB.Recordset
  rsHistoryDetail.Open "HistoryDetail", db, adOpenStatic, adLockOptimistic
  rsDetail.MoveFirst
  Do While Not rsDetail.EOF
    rsHistoryDetail.AddNew
    rsHistoryDetail!KodeRef = txtFields(0).Text
    rsHistoryDetail!Inventory = rsDetail!Inventory
    rsHistoryDetail!NamaInventory = rsDetail!NamaInventory
    rsHistoryDetail!Jumlah = rsDetail!JumlahKonsumsiAcc
    rsHistoryDetail!Satuan = rsDetail!NamaSatuan
    rsHistoryDetail!HargaSatuan = rsDetail!HargaPerUnit
    rsHistoryDetail!SubTotal = rsDetail!SubTotalAcc
    rsHistoryDetail.Update
    '
    rsDetail.MoveNext
  Loop
  rsHistoryDetail.Close
  Set rsHistoryDetail = Nothing
  '
  db.CommitTrans
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
      rsCari.Open "Select * FROM [C_BuatResepSpesial] WHERE KodeBuat LIKE '%" & txtCari.Text & "%' and IsNowAcc=False Order By KodeBuat", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [C_BuatResepSpesial] WHERE KodeBuat LIKE '%" & txtCari.Text & "%' and IsNowAcc=True Order By KodeBuat", db, adOpenStatic, adLockOptimistic
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
    MsgBox "Nomor Kode Pembuatan Resep tidak boleh kosong", vbCritical
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
      rsAddKomposisi!JumlahKonsumsi = (rsKomposisi!JumlahKonsumsi / sngJumlahFisik) * txtFields(6).Text
      rsAddKomposisi!SubTotal = ((rsKomposisi!JumlahKonsumsi / sngJumlahFisik) * txtFields(6).Text) * rsKomposisi!HargaPerUnit
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
  rsJumlahBiaya.Open "Select Sum(SubtotalAcc) as Jumlah From [C_BuatResepSpesialBahan] Where KodeBuat='" & txtFields(0).Text & "'", db, adOpenStatic, adLockReadOnly
  txtFields(7).Text = rsJumlahBiaya!Jumlah
  rsJumlahBiaya.Close
  Set rsJumlahBiaya = Nothing
  '
  ProsesDetail
  '
End Sub

Private Sub dgPI_AfterColEdit(ByVal ColIndex As Integer)
  If ColIndex = 3 Then
    '
    Dim rsInventory As New ADODB.Recordset
    Dim iBatas As Single
    '
    iBatas = 0
    rsInventory.Open "Select NamaInventory, JumlahItem From [C_Inventory] Where NamaInventory='" & dgPI.Columns(0).Value & "'", db, adOpenStatic, adLockReadOnly
    If rsInventory.RecordCount <> 0 Then iBatas = rsInventory!JumlahItem
    rsInventory.Close
    Set rsInventory = Nothing
    '
    If dgPI.Columns(3).Value > iBatas Then
      MsgBox "Jumlah Inventory " & dgPI.Columns(0).Value & " yang ada tinggal " & iBatas, vbInformation
      dgPI.Columns(3).Value = 0
      Exit Sub
    End If
    '
    bolResepNotAcc = False
    dgPI.Columns(5).Value = dgPI.Columns(1).Value * dgPI.Columns(3).Value
    SendKeys "{DOWN}"
  End If
End Sub

Private Sub dgPI_AfterUpdate()
  '
  Dim rsJumlahBiaya As New ADODB.Recordset
  rsJumlahBiaya.Open "Select Sum(SubtotalAcc) as Jumlah From [C_BuatResepSpesialBahan] Where KodeBuat='" & txtFields(0).Text & "'", db, adOpenStatic, adLockReadOnly
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
  If adoPrimaryRS!IsNowAcc = False Then
    MsgBox "Data Permintaan Pembuatan Resep belum diproses", vbCritical
    Exit Sub
  End If
  '
  sSQL = "SELECT C_BuatResepSpesial.KodeBuat, C_BuatResepSpesial.NamaResep, C_BuatResepSpesial.Jumlah, C_BuatResepSpesial.Satuan, C_BuatResepSpesialBahan.NamaInventory, C_BuatResepSpesialBahan.NamaSatuan, C_BuatResepSpesialBahan.HargaPerUnit, C_BuatResepSpesialBahan.JumlahKonsumsiAcc FROM [C_BuatResepSpesial] INNER JOIN [C_BuatResepSpesialBahan] ON C_BuatResepSpesial.KodeBuat = C_BuatResepSpesialBahan.KodeBuat WHERE (((C_BuatResepSpesial.KodeBuat)='" & txtFields(0).Text & "'))"

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
    .WindowTitle = "Form Persetujuan Pembuatan Resep"
    .ReportFileName = App.Path & "\Report\Laporan Setuju Buat Resep.Rpt"
    .Formulas(0) = "IDCompany= '" & ICNama & "'"
    .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
    .Formulas(2) = "IDKota= '" & ICKota & "'"
    .Action = 1
  End With
  '
End Sub
