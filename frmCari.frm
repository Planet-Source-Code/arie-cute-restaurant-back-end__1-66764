VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCari 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cari"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
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
   ScaleHeight     =   3075
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataList dtlCari 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4921
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCari As ADODB.Recordset

Private Sub Form_Load()
  Set rsCari = New ADODB.Recordset
  Select Case UCase$(sFormAktif)
    Case "FRMCINVENTORY"
      Select Case sKontrolAktif
        Case "AI_1"
          rsCari.Open "SELECT Gudang, NamaGudang FROM [C_Gudang] Order By NamaGudang", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaGudang"
            .BoundColumn = "Gudang"
          End With
        Case "AI_2"
          rsCari.Open "SELECT Kategori, NamaKategori FROM [C_IKategori] Order By NamaKategori", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKategori"
            .BoundColumn = "Kategori"
          End With
        Case "AI_3", "AI_4"
          rsCari.Open "SELECT Satuan, NamaSatuan FROM [C_Satuan] Order By NamaSatuan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaSatuan"
            .BoundColumn = "Satuan"
          End With
      End Select
    Case "FRMCADDSUPPLIER"
      Select Case sKontrolAktif
        Case "AIA_1"
          rsCari.Open "SELECT Supplier, NamaSupplier FROM [C_Supplier] Order By NamaSupplier", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaSupplier"
            .BoundColumn = "Supplier"
          End With
      End Select
    Case "FRMCIPORDER"
      Select Case sKontrolAktif
        Case "APO_1"
          rsCari.Open "SELECT Supplier, NamaSupplier FROM [C_Supplier] Order By NamaSupplier", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaSupplier"
            .BoundColumn = "Supplier"
          End With
        Case "APO_2"
          rsCari.Open "SELECT Karyawan, NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKaryawan"
            .BoundColumn = "Karyawan"
          End With
      End Select
    Case "FRMCIPINVOICE"
      Select Case sKontrolAktif
        Case "API_1"
          rsCari.Open "SELECT PO FROM [C_IPO]", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "PO"
            .BoundColumn = "PO"
          End With
      End Select
    Case "FRMCRESEP"
      Select Case sKontrolAktif
        Case "AR_1"
          rsCari.Open "SELECT Satuan, NamaSatuan FROM [C_Satuan] Order By NamaSatuan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaSatuan"
            .BoundColumn = "Satuan"
          End With
      End Select
    Case "FRMCRESEPBAHAN"
      Select Case sKontrolAktif
        Case "ARB_1"
          rsCari.Open "SELECT Inventory, NamaInventory FROM [C_Inventory] Where Inventory<>'" & frmCResep.txtFields(0).Text & "' Order By NamaInventory", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaInventory"
            .BoundColumn = "Inventory"
          End With
      End Select
    Case "FRMCPRODUKMENU"
      Select Case sKontrolAktif
        Case "AM_1"
          rsCari.Open "SELECT Departemen, NamaDepartemen FROM [C_MDepartemen] Order By NamaDepartemen", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaDepartemen"
            .BoundColumn = "Departemen"
          End With
      End Select
    Case "FRMCPRODUKMENUBAHAN"
      Select Case sKontrolAktif
        Case "ARM_1"
          rsCari.Open "SELECT Inventory, NamaInventory FROM [C_Inventory] Order By NamaInventory", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaInventory"
            .BoundColumn = "Inventory"
          End With
      End Select
    Case "FRMCMINTA"
      Select Case sKontrolAktif
        Case "AMI_1"
          rsCari.Open "SELECT Karyawan, NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKaryawan"
            .BoundColumn = "Karyawan"
          End With
      End Select
    Case "FRMCKEMBALI"
      Select Case sKontrolAktif
        Case "AKI_1"
          rsCari.Open "SELECT KodeBuat, Selesai FROM [C_Minta] Where NeedReturn=True", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "KodeBuat"
            .BoundColumn = "KodeBuat"
          End With
      End Select
    Case "FRMCMANUALMASUK"
      Select Case sKontrolAktif
        Case "AMM_1"
          rsCari.Open "SELECT Karyawan, NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKaryawan"
            .BoundColumn = "Karyawan"
          End With
      End Select
    Case "FRMCMANUALKELUAR"
      Select Case sKontrolAktif
        Case "AMK_1"
          rsCari.Open "SELECT Karyawan, NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKaryawan"
            .BoundColumn = "Karyawan"
          End With
      End Select
    Case "FRMCIRUSAK"
      Select Case sKontrolAktif
        Case "AIR_1"
          rsCari.Open "SELECT Karyawan, NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKaryawan"
            .BoundColumn = "Karyawan"
          End With
      End Select
    Case "FRMCOPTIMA2510T"
      Select Case sKontrolAktif
        Case "AIDM_2"
          rsCari.Open "SELECT Mesin, TipeMesin FROM [C_Mesin] Where TipeMesin='CR 2510T'", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "Mesin"
            .BoundColumn = "Mesin"
          End With
      End Select
    Case "FRMCMICROS"
      Select Case sKontrolAktif
        Case "AIDM_1"
          rsCari.Open "SELECT Mesin, TipeMesin FROM [C_Mesin] Where TipeMesin='Micros'", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "Mesin"
            .BoundColumn = "Mesin"
          End With
      End Select
    Case "FRMCBUATRESEPMODIFIER"
      Select Case sKontrolAktif
        Case "ABRM_1"
          rsCari.Open "SELECT Karyawan, NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaKaryawan"
            .BoundColumn = "Karyawan"
          End With
        Case "ABRM_2"
          rsCari.Open "SELECT Resep, NamaResep FROM [C_Resep] Order By NamaResep", db, adOpenStatic, adLockReadOnly
          Set dtlCari.RowSource = rsCari
          With dtlCari
            .ListField = "NamaResep"
            .BoundColumn = "Resep"
          End With
      End Select
  End Select
End Sub

Private Sub dtlCari_DblClick()
  Select Case UCase$(sFormAktif)
    Case "FRMCINVENTORY"
      Select Case sKontrolAktif
        Case "AI_1"
          frmCInventory.txtFields(2).Text = ""
          frmCInventory.txtFields(2).Text = dtlCari.BoundText
        Case "AI_2"
          frmCInventory.txtFields(4).Text = ""
          frmCInventory.txtFields(4).Text = dtlCari.BoundText
        Case "AI_3"
          frmCInventory.txtFields(6).Text = ""
          frmCInventory.txtFields(6).Text = dtlCari.BoundText
        Case "AI_4"
          frmCInventory.txtFields(8).Text = ""
          frmCInventory.txtFields(8).Text = dtlCari.BoundText
      End Select
    Case "FRMCADDSUPPLIER"
      Select Case sKontrolAktif
        Case "AIA_1"
          frmCAddSupplier.txtFields(0).Text = ""
          frmCAddSupplier.txtFields(0).Text = dtlCari.BoundText
      End Select
    Case "FRMCIPORDER"
      Select Case sKontrolAktif
        Case "APO_1"
          frmCIPOrder.txtFields(2).Text = ""
          frmCIPOrder.txtFields(2).Text = dtlCari.BoundText
        Case "APO_2"
          frmCIPOrder.txtFields(4).Text = ""
          frmCIPOrder.txtFields(4).Text = dtlCari.BoundText
      End Select
    Case "FRMCIPINVOICE"
      Select Case sKontrolAktif
        Case "API_1"
          frmCIPInvoice.txtFields(2).Text = ""
          frmCIPInvoice.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCRESEP"
      Select Case sKontrolAktif
        Case "AR_1"
          frmCResep.txtFields(2).Text = ""
          frmCResep.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCRESEPBAHAN"
      Select Case sKontrolAktif
        Case "ARB_1"
          frmCResepBahan.txtFields(0).Text = ""
          frmCResepBahan.txtFields(0).Text = dtlCari.BoundText
      End Select
    Case "FRMCPRODUKMENU"
      Select Case sKontrolAktif
        Case "AM_1"
          frmCProdukMenu.txtFields(2).Text = ""
          frmCProdukMenu.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCPRODUKMENUBAHAN"
      Select Case sKontrolAktif
        Case "ARM_1"
          frmCProdukMenuBahan.txtFields(0).Text = ""
          frmCProdukMenuBahan.txtFields(0).Text = dtlCari.BoundText
      End Select
    Case "FRMCMINTA"
      Select Case sKontrolAktif
        Case "AMI_1"
          frmCMinta.txtFields(2).Text = ""
          frmCMinta.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCKEMBALI"
      Select Case sKontrolAktif
        Case "AKI_1"
          frmCKembali.txtFields(2).Text = ""
          frmCKembali.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCMANUALMASUK"
      Select Case sKontrolAktif
        Case "AMM_1"
          frmCManualMasuk.txtFields(2).Text = ""
          frmCManualMasuk.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCMANUALKELUAR"
      Select Case sKontrolAktif
        Case "AMK_1"
          frmCManualKeluar.txtFields(2).Text = ""
          frmCManualKeluar.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCIRUSAK"
      Select Case sKontrolAktif
        Case "AIR_1"
          frmCIRusak.txtFields(2).Text = ""
          frmCIRusak.txtFields(2).Text = dtlCari.BoundText
      End Select
    Case "FRMCOPTIMA2510T"
      Select Case sKontrolAktif
        Case "AIDM_2"
          frmCOptima2510T.txtMesin.Text = ""
          frmCOptima2510T.txtMesin.Text = dtlCari.BoundText
      End Select
    Case "FRMCMICROS"
      Select Case sKontrolAktif
        Case "AIDM_1"
          frmCMicros.txtMesin.Text = ""
          frmCMicros.txtMesin.Text = dtlCari.BoundText
      End Select
    Case "FRMCBUATRESEPMODIFIER"
      Select Case sKontrolAktif
        Case "ABRM_1"
          frmCBuatResepModifier.txtFields(2).Text = ""
          frmCBuatResepModifier.txtFields(2).Text = dtlCari.BoundText
        Case "ABRM_2"
          frmCBuatResepModifier.txtFields(4).Text = ""
          frmCBuatResepModifier.txtFields(4).Text = dtlCari.BoundText
      End Select
  End Select
  '
  rsCari.Close
  Set rsCari = Nothing
  Unload Me
End Sub

Private Sub dtlCari_KeyPress(KeyAscii As Integer)
  '
  Select Case KeyAscii
    Case 13
      dtlCari_DblClick
    Case 27
      Unload Me
  End Select
  '
End Sub
