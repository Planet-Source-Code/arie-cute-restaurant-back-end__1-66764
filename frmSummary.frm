VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSummary 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Summary"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
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
   ScaleHeight     =   2310
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   6
      Top             =   1860
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
      Height          =   1755
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6195
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   960
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
         Format          =   24641539
         CurrentDate     =   37160
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Index           =   1
         Left            =   4260
         TabIndex        =   3
         Top             =   960
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
         Format          =   24641539
         CurrentDate     =   37160
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   120
         Picture         =   "frmSummary.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sampai"
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
         Index           =   1
         Left            =   3420
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari tanggal :"
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
         Index           =   0
         Left            =   540
         TabIndex        =   7
         Top             =   1020
         Width           =   1125
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
         Left            =   780
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   1035
         Index           =   1
         Left            =   60
         Top             =   660
         Width           =   6075
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
         Left            =   3600
         TabIndex        =   4
         Top             =   1380
         Width           =   285
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
      Left            =   3240
      TabIndex        =   0
      Top             =   1860
      Width           =   1455
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  For iCounter = 0 To 1
    dtpTanggal(iCounter).Value = Date
  Next
End Sub

Private Sub cmdCetak_Click()
  '
  Dim sPeriodik As String
  '
  Select Case frmSummary.Tag
    Case "A"
      sPeriodik = "SELECT C_IStockcard.Inventory, C_IStockcard.NamaInventory, Sum(C_IStockcard.JumlahPI) AS TotalPI, C_IStockcard.SatuanBesar, Sum(C_IStockcard.UnitMasuk) AS TotalUnit, C_Inventory.NamaSatuanKecil FROM [C_IStockcard] INNER JOIN [C_Inventory] ON C_IStockcard.Inventory = C_Inventory.Inventory Where (((C_IStockcard.Tanggal)>=#" & Format(dtpTanggal(0).Value, "mm/dd/yyyy") & "# And (C_IStockcard.Tanggal)<=#" & Format(dtpTanggal(1).Value, "mm/dd/yyyy") & "#) And ((C_IStockcard.IsInvoice) = True)) GROUP BY C_IStockcard.Inventory, C_IStockcard.NamaInventory, C_IStockcard.SatuanBesar, C_Inventory.NamaSatuanKecil"
      '
      With frmUtama.datReport
        .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
        .RecordSource = sPeriodik
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
        .WindowTitle = "Ringkasan Pembelian per Periode (Purchase Summary)"
        .ReportFileName = App.Path & "\Report\Laporan Purchase Summary.Rpt"
        .Formulas(0) = "DateFrom= '" & Format(dtpTanggal(0).Value, "dd/mm/yyyy") & "'"
        .Formulas(1) = "DateTo= '" & Format(dtpTanggal(1).Value, "dd/mm/yyyy") & "'"
        .Formulas(2) = "IDCompany= '" & ICNama & "'"
        .Formulas(3) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(4) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
      Exit Sub
      '
    Case "B"
      sPeriodik = "SELECT C_Menu.NamaDepartemen, C_SalesDetail.NamaMenu, C_SalesDetail.Biaya AS Modal, C_SalesDetail.HargaJual AS Harga, Sum(C_SalesDetail.Sales) AS SumOfSales, Sum(C_SalesDetail.NilaiSales) AS SumOfNilaiSales FROM [C_SalesDetail] INNER JOIN [C_Menu] ON C_SalesDetail.Menu = C_Menu.Menu Where (((C_SalesDetail.TanggalJual)>= #" & Format(dtpTanggal(0).Value, "mm/dd/yyyy") & "# And (C_SalesDetail.TanggalJual)<= #" & Format(dtpTanggal(1).Value, "mm/dd/yyyy") & "#)) GROUP BY C_Menu.NamaDepartemen, C_SalesDetail.NamaMenu, C_SalesDetail.Biaya, C_SalesDetail.HargaJual"
      '
      With frmUtama.datReport
        .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
        .RecordSource = sPeriodik
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
        .WindowTitle = "Ringkasan Pembelian per Periode (Purchase Summary)"
        .ReportFileName = App.Path & "\Report\Laporan Sales Summary.Rpt"
        .Formulas(0) = "DateFrom= '" & Format(dtpTanggal(0).Value, "dd/mm/yyyy") & "'"
        .Formulas(1) = "DateTo= '" & Format(dtpTanggal(1).Value, "dd/mm/yyyy") & "'"
        .Formulas(2) = "IDCompany= '" & ICNama & "'"
        .Formulas(3) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(4) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
      Exit Sub
      '
    Case "C"
      sPeriodik = "SELECT C_Inventory.NamaKategori, C_CostControl.NamaInventory, C_Inventory.HargaPerUnit AS Harga, Sum(C_CostControl.Ambil) AS SumOfAmbil, Sum(C_CostControl.Kembali) AS SumOfKembali, SumOfAmbil-SumOfKembali AS Actual, Sum(C_CostControl.Sales) AS SumOfSales FROM [C_CostControl] INNER JOIN [C_Inventory] ON C_CostControl.Inventory = C_Inventory.Inventory Where (((C_CostControl.Tanggal)>= #" & Format(dtpTanggal(0).Value, "mm/dd/yyyy") & "# And (C_CostControl.Tanggal)<= #" & Format(dtpTanggal(1).Value, "mm/dd/yyyy") & "#)) GROUP BY C_Inventory.NamaKategori, C_CostControl.NamaInventory, C_Inventory.HargaPerUnit"
      '
      With frmUtama.datReport
        .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
        .RecordSource = sPeriodik
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
        .WindowTitle = "Ringkasan Penggunaan Inventory (Usage Summary)"
        .ReportFileName = App.Path & "\Report\Laporan Usage Summary.Rpt"
        .Formulas(0) = "DateFrom= '" & Format(dtpTanggal(0).Value, "dd/mm/yyyy") & "'"
        .Formulas(1) = "DateTo= '" & Format(dtpTanggal(1).Value, "dd/mm/yyyy") & "'"
        .Formulas(2) = "IDCompany= '" & ICNama & "'"
        .Formulas(3) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(4) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
      Exit Sub
      '
    Case "D"
      sPeriodik = "SELECT C_Inventory.NamaKategori, C_IStockCard.Inventory, C_IStockCard.NamaInventory, Sum(C_IStockCard.UnitRusak) AS Jumlah, C_Inventory.NamaSatuanKecil, C_Inventory.HargaPerUnit FROM [C_IStockCard] INNER JOIN [C_Inventory] ON C_IStockCard.Inventory = C_Inventory.Inventory Where (((C_IStockCard.Tanggal) >=#" & Format(dtpTanggal(0).Value, "mm/dd/yyyy") & "# And (C_IStockCard.Tanggal)<=#" & Format(dtpTanggal(1).Value, "mm/dd/yyyy") & "#) And ((C_IStockCard.Keterangan) = 'Waste/Shrinkage')) GROUP BY C_Inventory.NamaKategori, C_IStockCard.Inventory, C_IStockCard.NamaInventory, C_Inventory.NamaSatuanKecil, C_Inventory.HargaPerUnit ORDER BY C_IStockCard.Inventory"
      '
      With frmUtama.datReport
        .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
        .RecordSource = sPeriodik
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
        .WindowTitle = "Ringkasan Inventory Terbuang/Rusak (Wasted Summary)"
        .ReportFileName = App.Path & "\Report\Laporan Wasted Summary.Rpt"
        .Formulas(0) = "DateFrom= '" & Format(dtpTanggal(0).Value, "dd/mm/yyyy") & "'"
        .Formulas(1) = "DateTo= '" & Format(dtpTanggal(1).Value, "dd/mm/yyyy") & "'"
        .Formulas(2) = "IDCompany= '" & ICNama & "'"
        .Formulas(3) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(4) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
    
    Case "E"
      sPeriodik = "SELECT C_IStockCard.Inventory, C_IStockCard.NamaInventory, C_Inventory.SatuanKecil, Sum(C_IStockCard.UnitPesan) AS SumOfUnitPesan, Sum(C_IStockCard.UnitMasuk) AS SumOfUnitMasuk, Sum(C_IStockCard.UnitKeluar) AS SumOfUnitKeluar, Sum(C_IStockCard.UnitRusak) AS SumOfUnitRusak, C_Inventory.HargaPerUnit FROM [C_IStockCard] INNER JOIN [C_Inventory] ON C_IStockCard.Inventory = C_Inventory.Inventory Where (((C_IStockCard.Tanggal)>=#" & Format(dtpTanggal(0).Value, "mm/dd/yyyy") & "# And (C_IStockCard.Tanggal)<=#" & Format(dtpTanggal(1).Value, "mm/dd/yyyy") & "#)) GROUP BY C_IStockCard.Inventory, C_IStockCard.NamaInventory, C_Inventory.SatuanKecil, C_Inventory.HargaPerUnit"
      '
      With frmUtama.datReport
        .DatabaseName = sPathAplikasi & "\Hospitality.Mdb"
        .RecordSource = sPeriodik
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
        .WindowTitle = "Ringkasan StockCard Inventory (StockCard Summary)"
        .ReportFileName = App.Path & "\Report\Laporan StockCard Summary.Rpt"
        .Formulas(0) = "DateFrom= '" & Format(dtpTanggal(0).Value, "dd/mm/yyyy") & "'"
        .Formulas(1) = "DateTo= '" & Format(dtpTanggal(1).Value, "dd/mm/yyyy") & "'"
        .Formulas(2) = "IDCompany= '" & ICNama & "'"
        .Formulas(3) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(4) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
      Exit Sub
      '
  End Select
  '
'ErrHandler:
'  MsgBox "Error: Ada beberapa prosedur operasi yang belum dilaksanakan", vbCritical
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub
