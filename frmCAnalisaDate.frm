VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCAnalisaDate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisa Pemakaian Inventory per Hari"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
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
   ScaleHeight     =   5265
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCetak 
      BackColor       =   &H8000000A&
      Caption         =   "&Cetak"
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
      Left            =   3300
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Keluar"
      Default         =   -1  'True
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
      Left            =   4800
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   0
      TabIndex        =   4
      Top             =   60
      Width           =   6315
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   3180
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dddd, d mmmm yyyy"
         Format          =   69533696
         CurrentDate     =   37160
      End
      Begin MSDataGridLib.DataGrid dgAnalisa 
         Bindings        =   "frmCAnalisaDate.frx":0000
         Height          =   3255
         Left            =   60
         TabIndex        =   1
         Top             =   1380
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   5741
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
         ColumnCount     =   8
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
            Caption         =   "Inventory"
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
            DataField       =   "NamaSatuanKecil"
            Caption         =   "Satuan*"
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
         BeginProperty Column03 
            DataField       =   "Ambil"
            Caption         =   "Ambil"
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
            DataField       =   "Kembali"
            Caption         =   "Kembali"
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
            DataField       =   "Pakai"
            Caption         =   "Aktual (X)"
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
         BeginProperty Column06 
            DataField       =   "Sales"
            Caption         =   "# Terpakai"
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
         BeginProperty Column07 
            DataField       =   "Varian"
            Caption         =   "Diff. (X - Y)"
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
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2324.977
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Object.Visible         =   0   'False
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCAnalisaDate.frx":0015
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Analisa Pemakaian Inventory"
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
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal "
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
         Index           =   9
         Left            =   2280
         TabIndex        =   5
         Top             =   900
         Width           =   720
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   60
         Top             =   180
         Width           =   6195
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   675
         Index           =   1
         Left            =   60
         Top             =   660
         Width           =   6195
      End
   End
   Begin VB.Label lblInv 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(*) Dalam satuan pakai / kecil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   2550
   End
End
Attribute VB_Name = "frmCAnalisaDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset

Private Sub TampilkanIsiGrid()
  '
  ' Isi Grid
  Set dgAnalisa.DataSource = Nothing
  If rsGrid.State = adStateOpen Then rsGrid.Close
  rsGrid.Open "SELECT C_CostControl.Tanggal, C_CostControl.Inventory, C_CostControl.NamaInventory, C_Inventory.NamaSatuanKecil, C_CostControl.Ambil, C_CostControl.Kembali, C_CostControl.Ambil-C_CostControl.Kembali AS Pakai, C_CostControl.Sales, Pakai-C_CostControl.Sales AS Varian FROM C_CostControl INNER JOIN C_Inventory ON C_CostControl.Inventory = C_Inventory.Inventory WHERE (((C_CostControl.Tanggal)=#" & dtpTanggal.Value & "#))", db, adOpenStatic, adLockReadOnly
  Set dgAnalisa.DataSource = rsGrid
  '
End Sub

Private Sub Form_Load()
  Set rsGrid = New ADODB.Recordset
End Sub

Private Sub Form_Activate()
  dtpTanggal.Value = Date
  TampilkanIsiGrid
End Sub

Private Sub dtpTanggal_Change()
  '
  TampilkanIsiGrid
  '
End Sub

Private Sub cmdCetak_Click()
  '
  Dim sSQL As String
  '
  sSQL = "SELECT C_CostControl.Tanggal, C_CostControl.Inventory, C_CostControl.NamaInventory, C_Inventory.NamaSatuanKecil, C_CostControl.Ambil, C_CostControl.Kembali, C_CostControl.Ambil-C_CostControl.Kembali AS Pakai, C_CostControl.Sales, Pakai-C_CostControl.Sales AS Varian FROM C_CostControl INNER JOIN C_Inventory ON C_CostControl.Inventory = C_Inventory.Inventory WHERE (((C_CostControl.Tanggal)=#" & dtpTanggal.Value & "#))"

  With frmUtama.datReport
    .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
    .RecordSource = sSQL
    .Refresh
  End With
  If frmUtama.datReport.Recordset.RecordCount = 0 Then
    MsgBox "Tidak ada data yang tercetak", vbInformation
    Exit Sub
  End If
  With frmUtama.crReport
    .Formulas(0) = ""
    .Formulas(1) = ""
    .Formulas(2) = ""
    .Formulas(3) = ""
    .Formulas(4) = ""
    .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
    .WindowTitle = "Analisa Pemakaian Inventory"
    .ReportFileName = App.Path & "\Report\Laporan Analisa Pemakaian Inventory.Rpt"
    .Formulas(0) = "Tanggal= '" & Format(dtpTanggal.Value, "dd/mm/yyyy") & "'"
    .Formulas(1) = "IDCompany= '" & ICNama & "'"
    .Formulas(2) = "IDAlamat= '" & ICAlamat & "'"
    .Formulas(3) = "IDKota= '" & ICKota & "'"
    .Formulas(4) = "Jam= '" & Format(Time, "HH:MM:SS") & "'"
    .Action = 1
  End With
  '
End Sub

Private Sub cmdOK_Click()
  If rsGrid.State = adStateOpen Then
    rsGrid.Close
    Set rsGrid = Nothing
  End If
  Unload Me
End Sub

