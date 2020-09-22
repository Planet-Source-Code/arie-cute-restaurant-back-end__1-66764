VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCPriceAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price Alert - Peringatan Perubahan Harga Jual Produk"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8775
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
      Width           =   8655
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
         Left            =   6960
         TabIndex        =   2
         Top             =   4200
         Width           =   1455
      End
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
         Left            =   5460
         TabIndex        =   1
         Top             =   4200
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dgAlert 
         Height          =   3495
         Left            =   60
         TabIndex        =   0
         Top             =   660
         Width           =   8535
         _ExtentX        =   15055
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
         Caption         =   "Price Alert"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Menu"
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
            DataField       =   "NamaMenu"
            Caption         =   "Nama Produk Menu"
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
            DataField       =   "Harga"
            Caption         =   "Harga Jual"
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
            DataField       =   "Untung"
            Caption         =   "Gross Margin"
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
            DataField       =   "PersenUntungOfSales"
            Caption         =   "% Margin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
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
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2580.095
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
         EndProperty
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(*) Untuk % Margin dibawah 50%"
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
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price Alert*"
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
         Width           =   1530
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmCPriceAlert.frx":0000
         Top             =   60
         Width           =   480
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   60
         Top             =   120
         Width           =   8535
      End
   End
End
Attribute VB_Name = "frmCPriceAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPriceAlert As ADODB.Recordset

Private Sub Form_Load()
  Set rsPriceAlert = New ADODB.Recordset
  rsPriceAlert.Open "SELECT * FROM [C_PriceAlert] ORDER BY MENU", db, adOpenStatic, adLockReadOnly
  Set dgAlert.DataSource = rsPriceAlert
End Sub

Private Sub cmdCetak_Click()
  If rsPriceAlert.RecordCount = 0 Then
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
    .WindowTitle = "Laporan Price Alert"
    .ReportFileName = App.Path & "\Report\Laporan Price Alert.Rpt"
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
  rsPriceAlert.Close
  Set rsPriceAlert = Nothing
End Sub
