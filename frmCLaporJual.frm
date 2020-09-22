VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCLaporJual 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Penjualan"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
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
   ScaleHeight     =   3285
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3240
      TabIndex        =   6
      Top             =   2820
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
      Height          =   2655
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   6195
      Begin VB.OptionButton optSales 
         BackColor       =   &H00FF8080&
         Caption         =   "Satu hari"
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
         Index           =   2
         Left            =   780
         TabIndex        =   4
         Top             =   1860
         Width           =   1095
      End
      Begin VB.OptionButton optSales 
         BackColor       =   &H00FF8080&
         Caption         =   "Periodik"
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
         Index           =   1
         Left            =   780
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optSales 
         BackColor       =   &H00FF8080&
         Caption         =   "Seluruhnya sampai hari ini"
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
         Index           =   0
         Left            =   780
         TabIndex        =   0
         Top             =   1020
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         Top             =   1380
         Width           =   1395
         _ExtentX        =   2461
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   24379395
         CurrentDate     =   37160
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   3
         Top             =   1380
         Width           =   1395
         _ExtentX        =   2461
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   24379395
         CurrentDate     =   37160
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Top             =   1800
         Width           =   1395
         _ExtentX        =   2461
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   24379395
         CurrentDate     =   37160
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCLaporJual.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Index           =   2
         Left            =   3660
         TabIndex        =   10
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laporan Penjualan"
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
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   60
         Top             =   180
         Width           =   6075
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   1935
         Index           =   1
         Left            =   60
         Top             =   660
         Width           =   6075
      End
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
      Left            =   4740
      TabIndex        =   7
      Top             =   2820
      Width           =   1455
   End
End
Attribute VB_Name = "frmCLaporJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub StatusOption()
  If optSales(1).Value = True Then
    dtpTanggal(0).Enabled = True
    dtpTanggal(1).Enabled = True
    '
    dtpTanggal(2).Enabled = False
    '
  ElseIf optSales(2).Value = True Then
    dtpTanggal(2).Enabled = True
    '
    dtpTanggal(0).Enabled = False
    dtpTanggal(1).Enabled = False
    '
  Else
    dtpTanggal(0).Enabled = False
    dtpTanggal(1).Enabled = False
    dtpTanggal(2).Enabled = False
  End If
End Sub

Private Sub Form_Load()
  For iCounter = 0 To 2
    dtpTanggal(iCounter).Value = Date
  Next
End Sub

Private Sub Form_Activate()
  StatusOption
End Sub

Private Sub optSales_Click(Index As Integer)
  StatusOption
End Sub

Private Sub cmdCetak_Click()
  '
  If optSales(0).Value = True Then
    '
    Dim sSales As String
    sSales = "SELECT TanggalJual, Mesin, NamaMenu, Biaya, HargaJual, Sales, Biaya*Sales AS SumOfBiaya, HargaJual*Sales AS SumOfJual, SumOfJual-SumOfBiaya AS Profit From [C_SalesDetail]"
    '
    With frmUtama.datReport
      .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
      .RecordSource = sSales
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
      .WindowTitle = "Total Penjualan"
      .ReportFileName = App.Path & "\Report\Laporan Sales.Rpt"
      .Formulas(0) = "IDCompany= '" & ICNama & "'"
      .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
      .Formulas(2) = "IDKota= '" & ICKota & "'"
      .Action = 1
    End With
    '
  ElseIf optSales(1).Value = True Then
    '
    Dim sPeriodikSales As String
    sPeriodikSales = "SELECT TanggalJual, Mesin, NamaMenu, Biaya, HargaJual, Sales, Biaya*Sales AS SumOfBiaya, HargaJual*Sales AS SumOfJual, SumOfJual-SumOfBiaya AS Profit From [C_SalesDetail] WHERE (((TanggalJual)>=# " & Format(dtpTanggal(0).Value, "mm/dd/yyyy") & "# and (TanggalJual)<=#" & Format(dtpTanggal(1).Value, "mm/dd/yyyy") & "#))"
    '
    With frmUtama.datReport
      .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
      .RecordSource = sPeriodikSales
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
      .WindowTitle = "Penjualan per Periode"
      .ReportFileName = App.Path & "\Sales.Rpt"
      .Formulas(0) = "IDCompany= '" & ICNama & "'"
      .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
      .Formulas(2) = "IDKota= '" & ICKota & "'"
      .Action = 1
    End With
    '
  Else
    '
    Dim sDailySales As String
    sDailySales = "SELECT TanggalJual, Mesin, NamaMenu, Biaya, HargaJual, Sales, Biaya*Sales AS SumOfBiaya, HargaJual*Sales AS SumOfJual, SumOfJual-SumOfBiaya AS Profit From [C_SalesDetail] WHERE (((TanggalJual)=# " & Format(dtpTanggal(2).Value, "mm/dd/yyyy") & " #))"
    '
    With frmUtama.datReport
      .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
      .RecordSource = sDailySales
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
      .WindowTitle = "Penjualan Harian"
      .ReportFileName = App.Path & "\Sales.Rpt"
      .Formulas(0) = "IDCompany= '" & ICNama & "'"
      .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
      .Formulas(2) = "IDKota= '" & ICKota & "'"
      .Action = 1
    End With
    '
  End If
  '
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub
