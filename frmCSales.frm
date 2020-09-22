VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCSales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Penjualan Harian"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
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
   ScaleHeight     =   6930
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&OK"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   6480
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
      Height          =   6315
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   8235
      Begin VB.ComboBox dcKode 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   22872064
         CurrentDate     =   37160
      End
      Begin MSDataGridLib.DataGrid dgAnalisa 
         Bindings        =   "frmCSales.frx":0000
         Height          =   4095
         Left            =   60
         TabIndex        =   1
         Top             =   1740
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   7223
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
            Caption         =   "Produk Menu"
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
            DataField       =   "HargaJual"
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
            DataField       =   "Sales"
            Caption         =   "# Jual"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "NilaiSales"
            Caption         =   "Nilai Penjualan"
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
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1755.213
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   240
         Picture         =   "frmCSales.frx":0015
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Tekan ENTER)"
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
         Left            =   3720
         TabIndex        =   9
         Top             =   1260
         Width           =   1230
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Mesin"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1260
         Width           =   960
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Index           =   1
         Left            =   60
         Top             =   660
         Width           =   8115
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Nilai Penjualan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   4260
         TabIndex        =   7
         Top             =   5940
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penjualan Harian (Till Tape)"
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
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   3780
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   60
         Top             =   180
         Width           =   8115
      End
   End
End
Attribute VB_Name = "frmCSales"
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
  '
  If JumlahRecord("Select * From [C_Sales] Where TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db) = 0 Then
    '
    db.BeginTrans
    '
    Dim rsTotalSales As New ADODB.Recordset
    rsTotalSales.Open "C_Sales", db, adOpenStatic, adLockOptimistic
    rsTotalSales.AddNew
    rsTotalSales!TanggalJual = dtpTanggal.Value
    rsTotalSales!Selesai = False
    rsTotalSales!Mesin = dcKode.Text
    rsTotalSales.Update
    rsTotalSales.Close
    Set rsTotalSales = Nothing
    '
    Dim rsSales As New ADODB.Recordset
    rsSales.Open "C_SalesDetail", db, adOpenStatic, adLockOptimistic
    Dim rsMenu As New ADODB.Recordset
    rsMenu.Open "Select * From [C_Menu] Where NonAktif=False", db, adOpenStatic, adLockReadOnly
    If rsMenu.RecordCount <> 0 Then
      rsMenu.MoveFirst
      Do While Not rsMenu.EOF
        '
        ' Tambah Record di file sales
        rsSales.AddNew
        rsSales!TanggalJual = dtpTanggal.Value
        rsSales!Menu = rsMenu!Menu
        rsSales!NamaMenu = rsMenu!NamaMenu
        rsSales!Biaya = rsMenu!Biaya
        rsSales!HargaJual = rsMenu!HargaJual
        rsSales!Mesin = dcKode.Text
        rsSales.Update
        '
        rsMenu.MoveNext
      Loop
    End If
    '
    rsSales.Close
    Set rsSales = Nothing
    rsMenu.Close
    Set rsMenu = Nothing
    '
    dcKode.Enabled = False
  Else
    dcKode.Enabled = True
  End If
  '
  rsGrid.Open "SELECT * From [C_SalesDetail] WHERE TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db, adOpenStatic, adLockOptimistic
  Set dgAnalisa.DataSource = rsGrid
  '
  Dim rsTotal As New ADODB.Recordset
  rsTotal.Open "Select Sum(NilaiSales) as TotalJual From [C_SalesDetail] WHERE TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db, adOpenStatic, adLockReadOnly
  txtTotal.Text = Format(rsTotal!TotalJual, "###,###,###,##0.00")
  rsTotal.Requery
  rsTotal.Close
  Set rsTotal = Nothing
  '
End Sub

Private Sub dgAnalisa_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
  On Error Resume Next
  '
  Dim rsPeriksa As New ADODB.Recordset
  rsPeriksa.Open "Select * From [C_Sales] WHERE TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db, adOpenStatic, adLockReadOnly
  If (rsPeriksa!Selesai = True) Or (rsPeriksa.RecordCount = 0) Then
    Cancel = True
  End If
  rsPeriksa.Close
  Set rsPeriksa = Nothing
End Sub

Private Sub dgAnalisa_AfterColEdit(ByVal ColIndex As Integer)
  If ColIndex = 3 Then
    dgAnalisa.Columns(4).Value = dgAnalisa.Columns(3).Value * dgAnalisa.Columns(2).Value
    SendKeys "{DOWN}"
  End If
End Sub

Private Sub dgAnalisa_AfterUpdate()
  '
  ' Hitung Jumlah Penjualan
  Dim rsJumlah As New ADODB.Recordset
  rsJumlah.Open "Select Sum(NilaiSales) as TotalJual From [C_SalesDetail] WHERE TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db, adOpenStatic, adLockReadOnly
  txtTotal.Text = Format(rsJumlah!TotalJual, "###,###,###,##0.00")
  rsJumlah.Requery
  rsJumlah.Close
  Set rsJumlah = Nothing
  '
End Sub

Private Sub Form_Load()
  Set rsGrid = New ADODB.Recordset
  '
  Dim rsKode As New ADODB.Recordset
  rsKode.Open "Select Mesin From [C_Mesin]", db, adOpenStatic, adLockReadOnly
  Do While Not rsKode.EOF
    dcKode.AddItem rsKode!Mesin
    rsKode.MoveNext
  Loop
  rsKode.Close
  Set rsKode = Nothing
End Sub

Private Sub Form_Activate()
  '
  dtpTanggal.Value = Date
  dcKode.SetFocus
  '
End Sub

Private Sub dcKode_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    TampilkanIsiGrid
  End If
  '
End Sub

Private Sub cmdSimpan_Click()
  Dim rsPeriksaStatus As New ADODB.Recordset
  rsPeriksaStatus.Open "Select * From [C_Sales] WHERE TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db, adOpenStatic, adLockReadOnly
  '
  If rsPeriksaStatus.RecordCount = 0 Then
    rsPeriksaStatus.Close
    Set rsPeriksaStatus = Nothing
    Unload Me
    Exit Sub
  End If
  '
  If rsPeriksaStatus!Selesai = False Then
    '
    db.Execute "Update [C_Sales] Set Selesai=True Where TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'"
    '
    Dim rsDetailSales As New ADODB.Recordset
    rsDetailSales.Open "Select * From [C_SalesDetail] Where TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & dcKode.Text & "'", db, adOpenStatic, adLockReadOnly
    rsDetailSales.MoveFirst
    Do While Not rsDetailSales.EOF
      '
      Dim rsBahanMenu As New ADODB.Recordset
      Dim JumlahTotal As Double
      Dim TotalSales As Double
      rsBahanMenu.Open "Select * From [C_MenuBahan] where Menu = '" & rsDetailSales!Menu & "'", db, adOpenStatic, adLockOptimistic
      Do While Not rsBahanMenu.EOF
        '
        ' Isi Cost Control
        Dim rsControl As New ADODB.Recordset
        rsControl.Open "Select * From [C_CostControl] Where Tanggal=#" & dtpTanggal.Value & "# and Inventory='" & rsBahanMenu!Inventory & "'", db, adOpenStatic, adLockOptimistic
        If rsControl.RecordCount = 0 Then
          rsControl.AddNew
          rsControl!Tanggal = dtpTanggal.Value
          rsControl!Inventory = rsBahanMenu!Inventory
          rsControl!NamaInventory = rsBahanMenu!NamaInventory
          rsControl.Update
        End If
        rsControl.Close
        Set rsControl = Nothing
        '
        JumlahTotal = Val(rsBahanMenu!JumlahKonsumsi) * rsDetailSales!Sales
        Dim rsUpdateCostControl As New ADODB.Recordset
        rsUpdateCostControl.Open "Select *  From [C_CostControl] Where (Inventory = '" & rsBahanMenu!Inventory & "') And (CDate(Tanggal) = '" & dtpTanggal.Value & "')", db, adOpenStatic, adLockOptimistic
        If rsUpdateCostControl.RecordCount <> 0 Then
          TotalSales = JumlahTotal + CDbl(rsUpdateCostControl!Sales)
          rsUpdateCostControl.Update "Sales", TotalSales
        End If
        '
        'MsgBox rsBahanMenu!NamaInventory
        'MsgBox TotalSales
        'MsgBox rsBahanMenu!Satuan
        'MsgBox rsBahanMenu!HargaPerUnit
        ' Update StockCard Bahan
        'Dim rsStockCardBahan As New ADODB.Recordset
        'rsStockCardBahan.Open "C_IStockCard", db, adOpenStatic, adLockOptimistic
        'rsStockCardBahan.AddNew
        'rsStockCardBahan!Tanggal = dtpTanggal.Value
        'rsStockCardBahan!Inventory = rsBahanMenu!Inventory
        'rsStockCardBahan!NamaInventory = rsBahanMenu!NamaInventory
        'rsStockCardBahan!Keterangan = "Sales"
        'rsStockCardBahan!PI = "SALES"
        'rsStockCardBahan!UnitKeluar = TotalSales
        'rsStockCardBahan!SatuanBesar = rsBahanMenu!Satuan
        'rsStockCardBahan!Harga = rsBahanMenu!HargaPerUnit
        'rsStockCardBahan!HargaSubTotal = rsBahanMenu!HargaPerUnit * TotalSales
        'rsStockCardBahan.Update
        'rsStockCardBahan.Close
        'Set rsStockCardBahan = Nothing
        '
        TotalSales = 0
        JumlahTotal = 0
        rsUpdateCostControl.Close
        rsBahanMenu.MoveNext
      Loop
      rsBahanMenu.Close
      '
      rsDetailSales.MoveNext
    Loop
    '
    rsDetailSales.Close
    Set rsDetailSales = Nothing
    '
    If MsgBox("Simpan Data Penjualan tanggal " & Format(dtpTanggal.Value, "d mmmm yyyy") & " dari Mesin " & dcKode.Text & " ?", vbQuestion + vbYesNo) = vbYes Then
      db.CommitTrans
      MsgBox "Data Penjualan tanggal " & Format(dtpTanggal.Value, "d mmmm yyyy") & " dari Mesin " & dcKode.Text & " disimpan", vbInformation
    Else
      db.RollbackTrans
    End If
    '
  End If
  rsPeriksaStatus.Close
  Set rsPeriksaStatus = Nothing
  '
  dgAnalisa.ReBind
  '
  dcKode.Enabled = True
  dcKode.SetFocus
  '
  Unload Me
End Sub

    
