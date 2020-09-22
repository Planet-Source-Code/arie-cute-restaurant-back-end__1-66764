VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCMicros 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data dari Micros"
   ClientHeight    =   3570
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
   ScaleHeight     =   3570
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
      Left            =   2460
      TabIndex        =   4
      Top             =   3120
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
      Height          =   3015
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   6195
      Begin VB.CommandButton cmdLookUp 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox txtMesin 
         DataField       =   "Mesin"
         Height          =   315
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1260
         Width           =   735
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   2400
         Width           =   1395
      End
      Begin VB.CommandButton cmdBuka 
         Caption         =   "&Buka File..."
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   2400
         Width           =   1395
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   5595
      End
      Begin MSComCtl2.DTPicker dtpTanggal 
         Height          =   315
         Left            =   2940
         TabIndex        =   6
         Top             =   900
         Width           =   2895
         _ExtentX        =   5106
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
         Format          =   24510464
         CurrentDate     =   37160
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCMicros.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Label lblPesan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harap Tunggu..."
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
         Left            =   300
         TabIndex        =   11
         Top             =   2460
         Width           =   1320
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update dari Nomor Mesin "
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
         Left            =   300
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama File Micros (POS)"
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
         Left            =   300
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import Data untuk  tanggal"
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
         Left            =   300
         TabIndex        =   8
         Top             =   960
         Width           =   2325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MICROS"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1290
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
         Height          =   2295
         Index           =   1
         Left            =   60
         Top             =   660
         Width           =   6075
      End
   End
End
Attribute VB_Name = "frmCMicros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName, sInitDir As String

Private Sub Form_Load()
  sFormAktif = Me.Name
End Sub

Private Sub Form_Activate()
  dtpTanggal.Value = Date
  lblPesan.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
  sFormAktif = ""
End Sub

Private Sub cmdBuka_Click()
  On Error GoTo Handler
  '
  If txtMesin.Text = "" Then
    MsgBox "Tentukan dulu Nama Mesin yang akan diimport", vbCritical
    txtMesin.SetFocus
    Exit Sub
  End If
  '
  With frmUtama.dlgDialog
    .InitDir = sInitDir
    .DialogTitle = "Membuka File POS (Micros)"
    .Filter = "Semua File (*.*) |*.*"
    .Flags = cdlOFNPathMustExist Or cdlOFNFileMustExist
    .CancelError = True
    .ShowOpen
  End With
  strFileName = frmUtama.dlgDialog.FileName
  txtFileName.Text = strFileName
  Exit Sub

Handler:
  strFileName = ""
  txtFileName.Text = ""
End Sub

Private Sub cmdImport_Click()
  '
  If txtFileName.Text = "" Then
    MsgBox "Spesifikasikan dulu nama file yang akan diimport", vbInformation
    cmdBuka.SetFocus
    Exit Sub
  End If
  '
  If JumlahRecord("Select * From [C_Sales] Where TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & txtMesin.Text & "'", db) = 0 Then
    '
    Dim rsTotalSales As New ADODB.Recordset
    rsTotalSales.Open "C_Sales", db, adOpenStatic, adLockOptimistic
    rsTotalSales.AddNew
    rsTotalSales!TanggalJual = dtpTanggal.Value
    rsTotalSales!Selesai = True
    rsTotalSales!Mesin = txtMesin.Text
    rsTotalSales.Update
    rsTotalSales.Close
    Set rsTotalSales = Nothing
    '
  Else
    MsgBox "Data dari Mesin " & txtMesin.Text & " sudah ada", vbInformation
    Exit Sub
  End If
  '
  lblPesan.Caption = "Harap Tunggu..."
  '
  Dim fs, f, Baris
  Dim Element() As String
  Dim sKodeMenu As String
  Dim iTerjual As Long
  '
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set f = fs.OpenTextFile(strFileName, 1, TristateFalse)
  Do While f.AtEndOfLine <> True
    Baris = f.ReadLine
    Element = Split(Baris, ",")
    sKodeMenu = "0000000000" & Element(0)
    iTerjual = Element(1)
    '
    Dim rsSales As New ADODB.Recordset
    rsSales.Open "Select * From [C_SalesDetail] Where Menu='" & sKodeMenu & "' and TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & txtMesin.Text & "'", db, adOpenStatic, adLockOptimistic
    Dim rsMenu As New ADODB.Recordset
    rsMenu.Open "Select * From [C_Menu] Where Menu='" & sKodeMenu & "'", db, adOpenStatic, adLockReadOnly
    If rsMenu.RecordCount <> 0 Then
      '
      If rsSales.RecordCount = 0 Then
        ' Tambah Record di file sales
        rsSales.AddNew
        rsSales!TanggalJual = dtpTanggal.Value
        rsSales!Menu = sKodeMenu
        rsSales!NamaMenu = rsMenu!NamaMenu
        rsSales!Biaya = rsMenu!Biaya
        rsSales!HargaJual = rsMenu!HargaJual
        rsSales!Sales = iTerjual
        rsSales!NilaiSales = rsMenu!HargaJual * iTerjual
        rsSales!Mesin = txtMesin.Text
        rsSales.Update
      Else
        Dim iJual As Long
        iJual = rsSales!Sales
        db.Execute "Update [C_SalesDetail] Set Sales=" & iJual + iTerjual & ", NilaiSales=" & rsMenu!HargaJual * iTerjual & " Where Menu='" & sKodeMenu & "' and TanggalJual=#" & dtpTanggal.Value & "#"
      End If
    End If
    '
    rsSales.Close
    Set rsSales = Nothing
    rsMenu.Close
    Set rsMenu = Nothing
    '
  Loop
  f.Close
  Set f = Nothing
  Set fs = Nothing
  '
  Dim rsDetailSales As New ADODB.Recordset
  rsDetailSales.Open "Select * From [C_SalesDetail] Where TanggalJual=#" & dtpTanggal.Value & "# and Mesin='" & txtMesin.Text & "'", db, adOpenStatic, adLockReadOnly
  rsDetailSales.MoveFirst
  Do While Not rsDetailSales.EOF
    '
    Dim rsBahanMenu As New ADODB.Recordset
    Dim JumlahTotal As Double
    Dim TotalSales As Double
    rsBahanMenu.Open "Select * From [C_MenuBahan] where Menu = '" & rsDetailSales!Menu & "'", db, adOpenStatic, adLockOptimistic
    Do While Not rsBahanMenu.EOF
      JumlahTotal = Val(rsBahanMenu!JumlahKonsumsi) * rsDetailSales!Sales
      Dim rsUpdateCostControl As New ADODB.Recordset
      rsUpdateCostControl.Open "Select *  From [C_CostControl] Where (Inventory = '" & rsBahanMenu!Inventory & "') And (CDate(Tanggal) = '" & dtpTanggal.Value & "')", db, adOpenStatic, adLockOptimistic
      If rsUpdateCostControl.RecordCount <> 0 Then
        TotalSales = JumlahTotal + CDbl(rsUpdateCostControl!Sales)
        rsUpdateCostControl.Update "Sales", TotalSales
      End If
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
  lblPesan.Caption = ""
  MsgBox "Data telah sukses diimport", vbInformation
  '
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub txtMesin_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    sKontrolAktif = "AIDM_1"
    '
    Dim rsMicros As New ADODB.Recordset
    rsMicros.Open "Select * From [C_Mesin] Where Mesin='" & txtMesin.Text & "' And TipeMesin='Micros'", db, adOpenStatic, adLockReadOnly
    If rsMicros.RecordCount = 0 Then
      frmCari.Show vbModal
      txtMesin.SetFocus
    Else
      sInitDir = rsMicros!PathData
      txtFileName.Text = ""
      SendKeys "{TAB}"
    End If
    rsMicros.Close
    sKontrolAktif = ""
  End If
  '
End Sub

Private Sub cmdLookUp_Click(Index As Integer)
  Select Case Index
    Case 0
      sKontrolAktif = "AIDM_1"
      frmCari.Show vbModal
      txtMesin.SetFocus
      If txtMesin.Text <> "" Then SendKeys "{ENTER}"
  End Select
  sKontrolAktif = ""
End Sub
