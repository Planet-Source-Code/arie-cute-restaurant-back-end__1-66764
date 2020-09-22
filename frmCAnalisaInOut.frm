VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCAnalisaInOut 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Perbandingan Pengambilan & Pengembalian Inventory"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
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
      Left            =   5460
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4695
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6915
      Begin VB.TextBox txtFields 
         BackColor       =   &H00E0E0E0&
         DataField       =   "Tanggal"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "99/99/9999"
         Top             =   1260
         Width           =   2955
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00E0E0E0&
         DataField       =   "NamaKaryawan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1620
         Width           =   2955
      End
      Begin MSDataListLib.DataCombo dcKode 
         Height          =   315
         Left            =   2880
         TabIndex        =   0
         Top             =   900
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "(dcKode)"
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
      Begin MSDataGridLib.DataGrid dgAnalisa 
         Bindings        =   "frmCAnalisaInOut.frx":0000
         Height          =   2355
         Left            =   60
         TabIndex        =   4
         Top             =   2280
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   4154
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
         ColumnCount     =   6
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
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Pakai"
            Caption         =   "Terpakai"
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
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   180
         Picture         =   "frmCAnalisaInOut.frx":0015
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
         Left            =   5100
         TabIndex        =   10
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Permintaan Inventory"
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
         TabIndex        =   9
         Top             =   960
         Width           =   2355
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
         Left            =   300
         TabIndex        =   8
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lblInv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Karyawan"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   1545
         Index           =   1
         Left            =   60
         Top             =   660
         Width           =   6795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perbandingan In-Out Inventory"
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
         TabIndex        =   3
         Top             =   240
         Width           =   4080
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   60
         Top             =   180
         Width           =   6795
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
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   2550
   End
End
Attribute VB_Name = "frmCAnalisaInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsGrid As ADODB.Recordset
Dim rsKode As ADODB.Recordset
  
Private Sub Form_Load()
  Set rsGrid = New ADODB.Recordset
End Sub

Private Sub Form_Activate()
  dcKode.Text = ""
  txtFields(1).Text = ""
  txtFields(2).Text = ""
End Sub

Private Sub dcKode_GotFocus()
  Set rsKode = New ADODB.Recordset
  rsKode.Open "Select KodeBuat, Selesai From [C_Minta] Where Done=True", db, adOpenStatic, adLockReadOnly
  Set dcKode.RowSource = rsKode
  dcKode.ListField = "KodeBuat"
  dcKode.BoundColumn = "KodeBuat"
End Sub

Private Sub dcKode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If dcKode.Text = "" Then
      MsgBox "Masukkan kode permintaan terlebih dahulu", vbInformation
      dcKode.SetFocus
      Exit Sub
    End If
    '
    Dim rsDetail As New ADODB.Recordset
    rsDetail.Open "C_Minta", db, adOpenStatic, adLockReadOnly
    rsDetail.Find "KodeBuat='" & dcKode.Text & "'"
    If rsDetail.RecordCount <> 0 Then
      txtFields(1).Text = " " & Format(rsDetail!Tanggal, "dddd, d mmmm yyyy")
      txtFields(2).Text = " " & rsDetail!NamaKaryawan
    End If
    rsDetail.Close
    Set rsDetail = Nothing
    '
    ' Isi Grid
    Set dgAnalisa.DataSource = Nothing
    If rsGrid.State = adStateOpen Then rsGrid.Close
    rsGrid.Open "SELECT C_AnalisaInOut.Kode, C_AnalisaInOut.Inventory, C_AnalisaInOut.NamaInventory, C_Inventory.NamaSatuanKecil, C_AnalisaInOut.Awal, C_AnalisaInOut.Ambil, C_AnalisaInOut.Kembali, C_AnalisaInOut.Ambil-C_AnalisaInOut.Kembali AS Pakai FROM C_AnalisaInOut INNER JOIN C_Inventory ON C_AnalisaInOut.Inventory = C_Inventory.Inventory WHERE (((C_AnalisaInOut.Kode)='" & dcKode.Text & "'))", db, adOpenStatic, adLockReadOnly
    Set dgAnalisa.DataSource = rsGrid
    '
  End If
End Sub

Private Sub cmdOK_Click()
  If rsGrid.State = adStateOpen Then
    rsGrid.Close
    Set rsGrid = Nothing
  End If
  If rsKode.State = adStateOpen Then
    rsKode.Close
    Set rsKode = Nothing
  End If
  Unload Me
End Sub
