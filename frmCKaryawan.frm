VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCKaryawan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Karyawan"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
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
   ScaleHeight     =   4770
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFields 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   3480
      TabIndex        =   29
      Top             =   720
      Width           =   5055
      Begin TabDlg.SSTab tabKaryawan 
         Height          =   3015
         Left            =   60
         TabIndex        =   31
         Top             =   60
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5318
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Info Pekerjaan"
         TabPicture(0)   =   "frmCKaryawan.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Picture1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Info Pribadi"
         TabPicture(1)   =   "frmCKaryawan.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Picture1(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   2595
            Index           =   1
            Left            =   60
            ScaleHeight     =   2595
            ScaleWidth      =   4815
            TabIndex        =   33
            Top             =   360
            Width           =   4815
            Begin VB.TextBox txtFields 
               DataField       =   "Telepon"
               Height          =   315
               Index           =   9
               Left            =   1680
               TabIndex        =   18
               Top             =   1800
               Width           =   2895
            End
            Begin VB.TextBox txtFields 
               DataField       =   "TanggalLahir"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "d MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Index           =   8
               Left            =   1680
               TabIndex        =   14
               Text            =   "99/99/9999"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtFields 
               DataField       =   "KodePos"
               Height          =   315
               Index           =   7
               Left            =   1680
               TabIndex        =   17
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox txtFields 
               DataField       =   "Kota"
               Height          =   315
               Index           =   6
               Left            =   1680
               TabIndex        =   16
               Top             =   1080
               Width           =   2895
            End
            Begin VB.TextBox txtFields 
               DataField       =   "Alamat"
               Height          =   315
               Index           =   5
               Left            =   1680
               TabIndex        =   15
               Top             =   720
               Width           =   2895
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Telepon"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   9
               Left            =   360
               TabIndex        =   43
               Top             =   1860
               Width           =   570
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Lahir"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   8
               Left            =   360
               TabIndex        =   42
               Top             =   420
               Width           =   960
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Pos"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   7
               Left            =   360
               TabIndex        =   41
               Top             =   1500
               Width           =   660
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kota"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   6
               Left            =   360
               TabIndex        =   40
               Top             =   1140
               Width           =   330
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat Rumah"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   39
               Top             =   780
               Width           =   1035
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   2595
            Index           =   0
            Left            =   -74940
            ScaleHeight     =   2595
            ScaleWidth      =   4815
            TabIndex        =   32
            Top             =   360
            Width           =   4815
            Begin VB.TextBox txtFields 
               DataField       =   "Ekstension"
               Height          =   315
               Index           =   4
               Left            =   2100
               TabIndex        =   13
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtFields 
               DataField       =   "TanggalMasuk"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "d MMM yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   2100
               TabIndex        =   12
               Text            =   "99/99/9999"
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox txtFields 
               DataField       =   "Jabatan"
               Height          =   315
               Index           =   2
               Left            =   2100
               TabIndex        =   11
               Top             =   1080
               Width           =   2535
            End
            Begin VB.TextBox txtFields 
               DataField       =   "NamaKaryawan"
               Height          =   315
               Index           =   1
               Left            =   2100
               TabIndex        =   10
               Top             =   720
               Width           =   2535
            End
            Begin VB.TextBox txtFields 
               DataField       =   "Karyawan"
               Height          =   315
               Index           =   0
               Left            =   2100
               MaxLength       =   6
               TabIndex        =   9
               Text            =   "123456"
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ekstension"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   3
               Left            =   360
               TabIndex        =   38
               Top             =   1860
               Width           =   765
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Masuk"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   37
               Top             =   1500
               Width           =   1065
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jabatan"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   36
               Top             =   1140
               Width           =   585
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Karyawan"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   35
               Top             =   420
               Width           =   1125
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Lengkap"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   34
               Top             =   780
               Width           =   1050
            End
         End
      End
   End
   Begin VB.Frame fraFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   60
      TabIndex        =   26
      Top             =   720
      Width           =   3375
      Begin TabDlg.SSTab TabIndek 
         Height          =   3135
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   5530
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Indeks"
         TabPicture(0)   =   "frmCKaryawan.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtIndeks"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstIndeks"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cari"
         TabPicture(1)   =   "frmCKaryawan.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblLabels(11)"
         Tab(1).Control(1)=   "txtCari"
         Tab(1).Control(2)=   "cboBy"
         Tab(1).Control(3)=   "cmdTampil"
         Tab(1).Control(4)=   "lstCari"
         Tab(1).ControlCount=   5
         Begin VB.ListBox lstIndeks 
            Height          =   2010
            Left            =   240
            TabIndex        =   1
            Top             =   840
            Width           =   2835
         End
         Begin VB.TextBox txtIndeks 
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   2835
         End
         Begin VB.ListBox lstCari 
            Height          =   1035
            ItemData        =   "frmCKaryawan.frx":0070
            Left            =   -74760
            List            =   "frmCKaryawan.frx":0077
            TabIndex        =   24
            Top             =   1860
            Width           =   2835
         End
         Begin VB.CommandButton cmdTampil 
            Caption         =   "&Tampilkan"
            Height          =   375
            Left            =   -74760
            TabIndex        =   23
            Top             =   1440
            Width           =   2835
         End
         Begin VB.ComboBox cboBy 
            Height          =   315
            ItemData        =   "frmCKaryawan.frx":0084
            Left            =   -74760
            List            =   "frmCKaryawan.frx":008E
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   720
            Width           =   2835
         End
         Begin VB.TextBox txtCari 
            Height          =   315
            Left            =   -74760
            TabIndex        =   22
            Top             =   1080
            Width           =   2835
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berdasarkan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   -74760
            TabIndex        =   28
            Top             =   480
            Width           =   1080
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   60
      TabIndex        =   25
      Top             =   3900
      Width           =   8535
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Baru"
         Height          =   795
         Left            =   6000
         Picture         =   "frmCKaryawan.frx":00B0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "&Terakhir"
         Height          =   795
         Left            =   2520
         Picture         =   "frmCKaryawan.frx":03BA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "&Pertama"
         Height          =   795
         Left            =   0
         Picture         =   "frmCKaryawan.frx":06C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "M&undur"
         Height          =   795
         Left            =   840
         Picture         =   "frmCKaryawan.frx":09CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Maju"
         Height          =   795
         Left            =   1680
         Picture         =   "frmCKaryawan.frx":0CD8
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Batal"
         Height          =   795
         Left            =   7680
         Picture         =   "frmCKaryawan.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Selesai"
         Height          =   795
         Left            =   7680
         Picture         =   "frmCKaryawan.frx":12EC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Simpan"
         Height          =   795
         Left            =   6840
         Picture         =   "frmCKaryawan.frx":15F6
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Koreksi"
         Height          =   795
         Left            =   6840
         Picture         =   "frmCKaryawan.frx":1900
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   180
      Picture         =   "frmCKaryawan.frx":1C0A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Karyawan"
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
      TabIndex        =   30
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   0
      Left            =   60
      Top             =   180
      Width           =   8475
   End
End
Attribute VB_Name = "frmCKaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim rsIndeks As ADODB.Recordset
Dim rsCari As ADODB.Recordset

Dim PosisiRecord As Long

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub StatusFrame(bolStatus As Boolean)
  '
  Picture1(0).Enabled = bolStatus
  Picture1(1).Enabled = bolStatus
  fraFind.Enabled = Not bolStatus
  '
End Sub

Private Sub Form_Load()
  '
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "select * from [Karyawan] Order by NamaKaryawan", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  
  RefreshIndeks
  lstCari.Clear
  TabIndek.Tab = 0
  tabKaryawan.Tab = 0
  '
  StatusFrame False
  '
  mbDataChanged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
  Screen.MousePointer = vbDefault
  '
  adoPrimaryRS.Close
  Set adoPrimaryRS = Nothing
  '
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  '
  tabKaryawan.Tab = 0
  StatusFrame True
  '
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    mbAddNewFlag = True
    SetButtons False
  End With
  '
  txtFields(0).SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  '
  If adoPrimaryRS.RecordCount = 0 Then
    MsgBox "Data Karyawan kosong", vbCritical
    Exit Sub
  End If
  
  tabKaryawan.Tab = 0
  mvBookMark = adoPrimaryRS.Bookmark
  StatusFrame True
  '
  mbEditFlag = True
  SetButtons False
  '
  txtFields(1).SetFocus
  SendKeys "{END}+{HOME}"
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  ' Validasi Field
  For iCounter = 0 To 1
    If txtFields(iCounter) = "" Then
      MsgBox "Field ini dibutuhkan", vbCritical
      tabKaryawan.Tab = 0
      txtFields(iCounter).SetFocus
      Exit Sub
    End If
  Next
  If txtFields(3).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    tabKaryawan.Tab = 0
    txtFields(3).SetFocus
    Exit Sub
  End If
  If txtFields(8).Text = "" Then
    MsgBox "Field ini dibutuhkan", vbCritical
    tabKaryawan.Tab = 1
    txtFields(8).SetFocus
    Exit Sub
  End If
  For iCounter = 5 To 7
    If txtFields(iCounter) = "" Then
      MsgBox "Field ini dibutuhkan", vbCritical
      tabKaryawan.Tab = 1
      txtFields(iCounter).SetFocus
      Exit Sub
    End If
  Next
  '
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  RefreshIndeks
  StatusFrame False
  tabKaryawan.Tab = 0
  Exit Sub
  '
UpdateErr:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  '
  RefreshIndeks
  StatusFrame False
  mbDataChanged = False
  '
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error Resume Next
  '
  adoPrimaryRS.MoveFirst
  mbDataChanged = False
  '
End Sub

Private Sub cmdLast_Click()
  On Error Resume Next
  '
  adoPrimaryRS.MoveLast
  mbDataChanged = False
  '
End Sub

Private Sub cmdNext_Click()
  On Error Resume Next
  '
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False
  '
End Sub

Private Sub cmdPrevious_Click()
  On Error Resume Next

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False
  '
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdClose.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub RefreshIndeks()
  '
  PosisiRecord = adoPrimaryRS.AbsolutePosition
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select NamaKaryawan FROM [Karyawan] Order By NamaKaryawan", db, adOpenStatic, adLockOptimistic
  Indeks
  '
End Sub

Private Sub Indeks()
  '
  On Error Resume Next
  '
  lstIndeks.Clear
  Me.MousePointer = vbHourglass
  '
  Do While Not rsIndeks.EOF
    lstIndeks.AddItem rsIndeks!NamaKaryawan
    rsIndeks.MoveNext
  Loop
  
  If PosisiRecord = 0 Then PosisiRecord = 1
  lstIndeks.Selected(PosisiRecord - 1) = True
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub Cari()
  '
  On Error Resume Next
  '
  lstCari.Clear
  Me.MousePointer = vbHourglass
  '
  Do While Not rsCari.EOF
    lstCari.AddItem rsCari!NamaKaryawan
    rsCari.MoveNext
  Loop
  Me.MousePointer = vbDefault
  '
End Sub

Private Sub txtIndeks_Change()
  '
  Set rsIndeks = New ADODB.Recordset
  rsIndeks.Open "Select NamaKaryawan FROM [Karyawan] WHERE NamaKaryawan LIKE '%" & txtIndeks.Text & "%' Order By NamaKaryawan", db, adOpenStatic, adLockOptimistic
  Indeks
  '
End Sub

Private Sub lstIndeks_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    lstIndeks_Click
  End If
  '
End Sub

Private Sub lstIndeks_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
      lstIndeks_Click
  End Select
End Sub

Private Sub lstIndeks_Click()
  On Error GoTo Handler
  '
  adoPrimaryRS.MoveFirst
  adoPrimaryRS.Find "NamaKaryawan ='" & lstIndeks.Text & "'"
  Exit Sub
  
Handler:
  MsgBox Err.Description
End Sub

Private Sub cmdTampil_Click()
  '
  If cboBy.ListIndex = -1 Then
    MsgBox "Tentukan kriteria pencarian terlebih dahulu", vbInformation
    cboBy.SetFocus
    Exit Sub
  End If
  '
  Set rsCari = New ADODB.Recordset
  Select Case cboBy.ListIndex
    Case 0
      rsCari.Open "Select * FROM [Karyawan] WHERE Karyawan LIKE '%" & txtCari.Text & "%' Order By Karyawan", db, adOpenStatic, adLockOptimistic
    Case 1
      rsCari.Open "Select * FROM [Karyawan] WHERE NamaKaryawan LIKE '%" & txtCari.Text & "%' Order By NamaKaryawan", db, adOpenStatic, adLockOptimistic
  End Select
  '
  Cari
  rsCari.Close
  '
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    cmdTampil_Click
    lstCari.SetFocus
  End If
  '
End Sub

Private Sub cboBy_Change()
  '
  txtCari.Text = ""
  txtCari.SetFocus
  '
End Sub

Private Sub cboBy_KeyPress(KeyAscii As Integer)
  SendKeys "{TAB}"
End Sub

Private Sub lstCari_Click()
  '
  On Error GoTo Handler
  '
  adoPrimaryRS.MoveFirst
  adoPrimaryRS.Find "NamaKaryawan ='" & lstCari.Text & "'"
  Exit Sub

Handler:
  MsgBox Err.Description
  '
End Sub

Private Sub lstCari_KeyUp(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
      lstCari_Click
  End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub
