VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm frmUtama 
   BackColor       =   &H8000000C&
   Caption         =   "Winsor F & B Control 2.0 - Single Store Edition"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6405
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport crReport 
      Bindings        =   "frmUtama.frx":030A
      Left            =   1860
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   2400
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0646
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":096A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0FB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":12CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":15EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsBitmaps 
      Left            =   540
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1C22
            Key             =   "mnuFileItem(1)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":1F3E
            Key             =   "mnuUtilItem(1)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":225A
            Key             =   "mnuJualItem(0)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":257A
            Key             =   "mnuPurchaseItem(0)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":2896
            Key             =   "mnuPurchaseItem(9)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":2BBA
            Key             =   "mnuUtilItem(6)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":2ED6
            Key             =   "mnuCostItem(7)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":31F6
            Key             =   "mnuCostItem(5)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":3516
            Key             =   "mnuBantuItem(2)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":3832
            Key             =   "mnuUtilItem(2)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":3B5A
            Key             =   "mnuCostItem(10)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":3E76
            Key             =   "mnuPurchaseItem(1)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":4192
            Key             =   "mnuUtilItem(5)"
         EndProperty
      EndProperty
   End
   Begin VB.Data datReport 
      Align           =   1  'Align Top
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   6405
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3945
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4233
            MinWidth        =   4233
            Text            =   " Nama User : Bagian Akunting"
            TextSave        =   " Nama User : Bagian Akunting"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:40"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4736
         EndProperty
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
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   1260
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2370
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4180
      ButtonWidth     =   2355
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setting Printer"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Item Inventory"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Resep Standar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Produk Menu"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Purchase Order"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Purchase Invoice"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Analisa"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "Perbandingan Ambil-Kembali Inventory"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   "Analisa Pemakaian Inventory"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "3"
                  Text            =   "Peringatan Stok (Stock Alert)"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "4"
                  Text            =   "Peringatan Harga Jual (Price Alert)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Laporan"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   13
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "6"
                  Text            =   "Supplier"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "7"
                  Text            =   "Item Inventory"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "8"
                  Text            =   "Resep Standar"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "9"
                  Text            =   "Produk Menu"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "10"
                  Text            =   "Penjualan"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "11"
                  Text            =   "Purchase Summary"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "12"
                  Text            =   "Period Sales Summary"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "13"
                  Text            =   "Usage Summary"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "14"
                  Text            =   "Wasted Summary"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "15"
                  Text            =   "Stock Card Summary"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileXX 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Informasi Perusahaan"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Setting Printer"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Keluar"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCostXX 
      Caption         =   "&Cost Control"
      Begin VB.Menu mnuCostItem 
         Caption         =   "&Supplier"
         Index           =   0
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "Satuan &Pengukuran"
         Index           =   1
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "&Gudang Penyimpanan"
         Index           =   3
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "&Kategori Inventory"
         Index           =   4
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "&Item Inventory "
         Index           =   5
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "&Resep Standar"
         Index           =   7
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "&Departemen"
         Index           =   9
      End
      Begin VB.Menu mnuCostItem 
         Caption         =   "Produk &Menu"
         Index           =   10
      End
   End
   Begin VB.Menu mnuPurchaseXX 
      Caption         =   "&Purchasing"
      Begin VB.Menu mnuPurchaseItem 
         Caption         =   "Purchase &Order"
         Index           =   0
      End
      Begin VB.Menu mnuPurchaseItem 
         Caption         =   "Purchase &Invoice"
         Index           =   1
      End
   End
   Begin VB.Menu mnuProduksiXX 
      Caption         =   "P&roduksi"
      Begin VB.Menu mnuProduksiItem 
         Caption         =   "Permintaan Pembuatan Resep "
         Index           =   0
      End
   End
   Begin VB.Menu mnuGudangXX 
      Caption         =   "&Gudang"
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Inventory Counting"
         Index           =   0
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Inventory Adjustment +"
         Index           =   2
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Inventory Adjustment -"
         Index           =   3
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Wasted Inventory"
         Index           =   4
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Persetujuan Pembuatan Resep"
         Index           =   6
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Penambahan Inventory Resep"
         Index           =   7
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Persetujuan Pengambilan Inventory"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGudangItem 
         Caption         =   "Pemeriksaan Pengembalian Inventory"
         Index           =   10
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDapurXX 
      Caption         =   "&Dapur"
      Visible         =   0   'False
      Begin VB.Menu mnuDapurItem 
         Caption         =   "Permintaan Item Inventory"
         Index           =   0
      End
      Begin VB.Menu mnuDapurItem 
         Caption         =   "Pengembalian Item Inventory"
         Index           =   1
      End
   End
   Begin VB.Menu mnuSalesXX 
      Caption         =   "&Sales"
      Begin VB.Menu mnuJualItem 
         Caption         =   "&Till Tape"
         Index           =   0
      End
      Begin VB.Menu mnuJualItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuJualItem 
         Caption         =   "Import dari &Micros"
         Index           =   2
      End
      Begin VB.Menu mnuJualItem 
         Caption         =   "Import dari &Optima Quorion 2510T (gITs)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuUtilXX 
      Caption         =   "&Manajer"
      Begin VB.Menu mnuUtilItem 
         Caption         =   "Set &Lokasi Database"
         Index           =   0
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "&Setup Cash Register"
         Index           =   1
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "&Wewenang User"
         Index           =   2
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "&Karyawan"
         Index           =   3
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "&Analisa"
         Index           =   5
         Begin VB.Menu mnuAnalisaItem 
            Caption         =   "&Perbandingan Ambil-Kembali Inventory"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAnalisaItem 
            Caption         =   "&Analisa Pemakaian Inventory"
            Index           =   1
         End
         Begin VB.Menu mnuAnalisaItem 
            Caption         =   "Peringatan &Stok (Stock Alert)"
            Index           =   2
         End
         Begin VB.Menu mnuAnalisaItem 
            Caption         =   "Peringatan &Harga Jual (Price Alert)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "&Laporan"
         Index           =   6
         Begin VB.Menu mnuLaporItem 
            Caption         =   "&Supplier"
            Index           =   0
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "&Item Inventory"
            Index           =   2
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "&Resep Standar"
            Index           =   3
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "&Produk Menu"
            Index           =   4
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "Penjualan"
            Index           =   6
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "Purchase Summary"
            Index           =   8
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "Period Sales Summary"
            Index           =   9
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "Usage Summary"
            Index           =   10
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "Wasted Summary"
            Index           =   11
         End
         Begin VB.Menu mnuLaporItem 
            Caption         =   "Stock Card Summary"
            Index           =   12
         End
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuUtilItem 
         Caption         =   "&Documents History"
         Index           =   8
      End
   End
   Begin VB.Menu mnuBantuXX 
      Caption         =   "&Bantuan"
      Begin VB.Menu mnuBantuItem 
         Caption         =   "Winsor Navigator"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBantuItem 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBantuItem 
         Caption         =   "Tentang..."
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit

Private Sub GetMenuIcon()
  On Error Resume Next
  Dim i As Long
   
End Sub

Private Sub MDIForm_Load()
  '
  GetMenuIcon
  '
  mnuCostXX.Enabled = bCost
  mnuPurchaseXX.Enabled = bPurchase
  mnuProduksiXX.Enabled = bProduksi
  mnuGudangXX.Enabled = bGudang
  mnuDapurXX.Enabled = bDapur
  mnuSalesXX.Enabled = bSales
  mnuUtilXX.Enabled = bGrant
  '
  MAX_CR = 5
  '
End Sub

Private Sub MDIForm_Resize()
  If Me.WindowState = 0 Then
    Me.WindowState = 2
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  sPesan = "Tindakan ini akan mengakhiri pemakaian aplikasi" & vbCrLf
  sPesan = sPesan & "Anda yakin ingin keluar ?"
  If MsgBox(sPesan, vbQuestion + vbYesNo) = vbNo Then
    Cancel = -1
  Else
    db.Close
    Set db = Nothing
    End
  End If
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmPerusahaan.Show vbModal
    Case 1
      On Error Resume Next
      '
      With dlgDialog
        .Flags = cdlPDPrintSetup
        .ShowPrinter
      End With
    Case 3
      Unload Me
  End Select
End Sub

Private Sub mnuCostItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmCSupplier.Show vbModal
    Case 1
      frmCSatuan.Show vbModal
    Case 3
      frmCGudang.Show vbModal
    Case 4
      frmCIKategori.Show vbModal
    Case 5
      frmCInventory.Show vbModal
    Case 7
      frmCResep.Show vbModal
    Case 9
      frmCDepartemen.Show vbModal
    Case 10
      frmCProdukMenu.Show vbModal
  End Select
End Sub

Private Sub mnuPurchaseItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmCIPOrder.Show vbModal
    Case 1
      frmCIPInvoice.Show vbModal
  End Select
End Sub

Private Sub mnuProduksiItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmCBuatResepModifier.Show vbModal
  End Select
End Sub

Private Sub mnuGudangItem_Click(Index As Integer)
  Select Case Index
    Case 0
      '
      ' Formulir Counting
      If JumlahRecord("Select Inventory From [C_Inventory]", db) = 0 Then
        MsgBox "Data Inventory kosong", vbInformation
        Exit Sub
      End If
      '
      With frmUtama.crReport
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = ""
        .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
        .WindowTitle = "Formulir Inventory Counting"
        .ReportFileName = App.Path & "\Report\Laporan Inventory Counting.Rpt"
        .Formulas(0) = "IDCompany= '" & ICNama & "'"
        .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(2) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
      '
    Case 2
      frmCManualMasuk.Show vbModal
    Case 3
      frmCManualKeluar.Show vbModal
    Case 4
      frmCIRusak.Show vbModal
    Case 6
      Dim iPermintaan As Long
      iPermintaan = JumlahRecord("Select KodeBuat, IsNowAcc From [C_BuatResepSpesial] Where IsNowAcc=False", db)
      If iPermintaan > 0 Then
        MsgBox "Ada " & iPermintaan & " Permintaan pembuatan resep", vbInformation
      End If
      '
      frmCSetujuResepModifier.Show vbModal
    Case 7
      Dim iTransfer As Long
      iTransfer = JumlahRecord("Select KodeBuat, NeedTransfer From [C_BuatResepSpesial] Where NeedTransfer=True", db)
      If iTransfer > 0 Then
        MsgBox "Ada " & iTransfer & " dokumen penambahan inventory resep", vbInformation
      End If
      '
      frmCAddInvResep.Show vbModal
    Case 9
      Dim iPengambilan As Long
      iPermintaan = JumlahRecord("Select KodeBuat, Selesai From [C_Minta] Where Selesai=False", db)
      If iPermintaan > 0 Then
        MsgBox "Ada " & iPermintaan & " permintaan pengambilan inventory", vbInformation
      End If
      '
      frmCSetujuMinta.Show vbModal
    Case 10
      Dim iPemeriksaan As Long
      iPemeriksaan = JumlahRecord("Select KodeBuat, NeedCheck From [C_Kembali] Where NeedCheck=True", db)
      If iPemeriksaan > 0 Then
        MsgBox "Ada " & iPemeriksaan & " pengembalian inventory yang harus diperiksa", vbInformation
      End If
      '
      frmCSetujuKembali.Show vbModal
  End Select
End Sub

Private Sub mnuDapurItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmCMinta.Show vbModal
    Case 1
      frmCKembali.Show vbModal
  End Select
End Sub

Private Sub mnuJualItem_Click(Index As Integer)
  '
  If JumlahRecord("Select * From [C_Mesin]", db) = 0 Then
    MsgBox "Tidak ada mesin yang terpasang", vbCritical
    Exit Sub
  End If
  If JumlahRecord("Select * From [C_Menu]", db) = 0 Then
    MsgBox "Produk menu belum ada", vbCritical
    Exit Sub
  End If
  '
  Select Case Index
    Case 0
      frmCSales.Show vbModal
    Case 2
      frmCMicros.Show vbModal
    Case 3
      frmCOptima2510T.Show vbModal
  End Select
  '
End Sub

Private Sub mnuUtilItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmSetDatabase.Show vbModal
    Case 1
      frmCSetupMesin.Show vbModal
    Case 2
      frmProfile.Show vbModal
    Case 3
      frmCKaryawan.Show vbModal
    Case 8
      frmHistory.Show vbModal
  End Select
End Sub

Private Sub mnuBantuItem_Click(Index As Integer)
  Select Case Index
    Case 0
      '
    Case 2
      frmTentang.Show vbModal
  End Select
End Sub

Private Sub mnuAnalisaItem_Click(Index As Integer)
  Select Case Index
    Case 0
      frmCAnalisaInOut.Show vbModal
    Case 1
      frmCAnalisaDate.Show vbModal
    Case 2
      frmCStockAlert.Show vbModal
    Case 3
      frmCPriceAlert.Show vbModal
  End Select
End Sub

Private Sub mnuLaporItem_Click(Index As Integer)
  '
  Select Case Index
    Case 0
      '
      If JumlahRecord("Select Supplier From [C_Supplier]", db) = 0 Then
        MsgBox "Tidak ada data yang tercetak", vbInformation
        Exit Sub
      End If
      '
      ' Laporan Supplier
      With frmUtama.crReport
        .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
        .WindowTitle = "Daftar Supplier"
        .ReportFileName = App.Path & "\Report\Laporan Supplier.Rpt"
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = ""
        .Action = 1
      End With
      '
    Case 2
      '
      If JumlahRecord("Select Inventory From [C_Inventory]", db) = 0 Then
        MsgBox "Tidak ada data yang tercetak", vbInformation
        Exit Sub
      End If
      '
      ' Laporan Inventory
      With frmUtama.crReport
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = ""
        .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
        .WindowTitle = "Item Inventory"
        .ReportFileName = App.Path & "\Report\Laporan Inventory.Rpt"
        .Formulas(0) = "IDCompany= '" & ICNama & "'"
        .Formulas(1) = "IDAlamat= '" & ICAlamat & "'"
        .Formulas(2) = "IDKota= '" & ICKota & "'"
        .Action = 1
      End With
      '
    Case 3
      '
      If JumlahRecord("Select Resep From [C_Resep]", db) = 0 Then
        MsgBox "Tidak ada data yang tercetak", vbInformation
        Exit Sub
      End If
      '
      ' Laporan Resep Standar
      With frmUtama.crReport
        .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
        .WindowTitle = "Resep Standar"
        .ReportFileName = App.Path & "\Report\Laporan Resep.Rpt"
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = ""
        .Action = 1
      End With
      '
    Case 4
      '
      If JumlahRecord("Select Menu From [C_Menu]", db) = 0 Then
        MsgBox "Tidak ada data yang tercetak", vbInformation
        Exit Sub
      End If
      '
      ' Laporan Produk Menu
      With frmUtama.crReport
        .DataFiles(0) = sPathAplikasi & "\Hospitality.Mdb"
        .WindowTitle = "Produk Menu"
        .ReportFileName = App.Path & "\Report\Laporan Menu.Rpt"
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        .Formulas(4) = ""
        .Action = 1
      End With
      '
    Case 6
      frmCLaporJual.Show vbModal
    Case 8
      frmSummary.Tag = "A"
      frmSummary.Caption = "Ringkasan Pembelian (Purchase Summary)"
      frmSummary.Label2.Caption = "Purchase Summary"
      frmSummary.Show vbModal
    Case 9
      frmSummary.Tag = "B"
      frmSummary.Caption = "Ringkasan Periodik Penjualan (Period Sales Summary)"
      frmSummary.Label2.Caption = "Period Sales Summary"
      frmSummary.Show vbModal
    Case 10
      frmSummary.Tag = "C"
      frmSummary.Caption = "Ringkasan Penggunaan Inventory (Usage Summary)"
      frmSummary.Label2.Caption = "Usage Summary"
      frmSummary.Show vbModal
    Case 11
      frmSummary.Tag = "D"
      frmSummary.Caption = "Ringkasan Inventory Terbuang/Rusak (Wasted Summary)"
      frmSummary.Label2.Caption = "Wasted Summary"
      frmSummary.Show vbModal
    Case 12
      frmSummary.Tag = "E"
      frmSummary.Caption = "Ringkasan StockCard Inventory (StockCard Summary)"
      frmSummary.Label2.Caption = "StockCard Summary"
      frmSummary.Show vbModal
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1
      mnuFileItem_Click (1)
    Case 3
      If Not bCost Then
        MsgBox "Akses ditolak", vbCritical
        Exit Sub
      Else
        mnuCostItem_Click (5)
      End If
    Case 4
      If Not bCost Then
        MsgBox "Akses ditolak", vbCritical
        Exit Sub
      Else
        mnuCostItem_Click (7)
      End If
    Case 5
      If Not bCost Then
        MsgBox "Akses ditolak", vbCritical
        Exit Sub
      Else
        mnuCostItem_Click (10)
      End If
    Case 7
      If Not bPurchase Then
        MsgBox "Akses ditolak", vbCritical
        Exit Sub
      Else
        mnuPurchaseItem_Click (0)
      End If
    Case 8
      If Not bPurchase Then
        MsgBox "Akses ditolak", vbCritical
        Exit Sub
      Else
        mnuPurchaseItem_Click (1)
      End If
  End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  If Not bGrant Then
    MsgBox "Akses ditolak", vbCritical
    Exit Sub
  Else
    Select Case ButtonMenu.Tag
      Case 1
        mnuAnalisaItem_Click (0)
      Case 2
        mnuAnalisaItem_Click (1)
      Case 3
        mnuAnalisaItem_Click (2)
      Case 4
        mnuAnalisaItem_Click (3)
      Case 6
        mnuLaporItem_Click (0)
      Case 7
        mnuLaporItem_Click (2)
      Case 8
        mnuLaporItem_Click (3)
      Case 9
        mnuLaporItem_Click (4)
      Case 10
        mnuLaporItem_Click (6)
      Case 11
        mnuLaporItem_Click (8)
      Case 12
        mnuLaporItem_Click (9)
      Case 13
        mnuLaporItem_Click (10)
      Case 14
        mnuLaporItem_Click (11)
      Case 15
        mnuLaporItem_Click (12)
    End Select
  End If
End Sub
