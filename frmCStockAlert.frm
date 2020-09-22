VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCStockAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Alert - Peringatan Persediaan Stok"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
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
   ScaleHeight     =   4815
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   7035
      Begin VB.CommandButton cmdCetak 
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
         Left            =   4020
         TabIndex        =   1
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Keluar"
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
         Left            =   5520
         TabIndex        =   2
         Top             =   4200
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dgAlert 
         Height          =   3495
         Left            =   60
         TabIndex        =   0
         Top             =   660
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   6165
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
         Caption         =   "Stock Alert"
         ColumnCount     =   5
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
            DataField       =   "ReorderLevel"
            Caption         =   "Min. Stok*"
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
            DataField       =   "QtyOnHand"
            Caption         =   "On Hand*"
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
            DataField       =   "Satuan"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2475.213
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
               ColumnWidth     =   1049.953
            EndProperty
         EndProperty
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(*) Dalam satuan pesan / besar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   2685
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmCStockAlert.frx":0000
         Top             =   60
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Alert"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   780
         TabIndex        =   4
         Top             =   180
         Width           =   1395
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   60
         Top             =   120
         Width           =   6915
      End
   End
End
Attribute VB_Name = "frmCStockAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsStokAlert As ADODB.Recordset

Private Sub Form_Load()
  Set rsStokAlert = New ADODB.Recordset
  rsStokAlert.Open "SELECT * FROM [C_StockAlert] ORDER BY INVENTORY", db, adOpenStatic, adLockReadOnly
  Set dgAlert.DataSource = rsStokAlert
End Sub

Private Sub cmdCetak_Click()
  If rsStokAlert.RecordCount = 0 Then
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
    .WindowTitle = "Laporan Stock Alert"
    .ReportFileName = App.Path & "\Report\Laporan Stock Alert.Rpt"
    .Formulas(0) = "IDCompany= '" & ICNama & "'"
    .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
    .Formulas(2) = "IDKota= '" & ICKota & "'"
    .Action = 1
  End With
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rsStokAlert.Close
  Set rsStokAlert = Nothing
End Sub
