Attribute VB_Name = "MainModule"
' ---------------------------------------------------
' Winsor F & B Control 2.0 Single Store Edition
' Hak Cipta(c) 2001 Oleh PT Winsor Satria Persada
' Desain Interface & Kode Program oleh Kurnia Sembada
' ---------------------------------------------------
'
Option Explicit

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public bCost, bPurchase, bProduksi, bGudang, bDapur, bSales, bGrant As Boolean
Public MAX_CR As Integer

' Variabel ID Perusahaan
Public ICNama, ICAlamat, ICKota

Public sPesan As String
Public sUserName, sFormAktif, sKontrolAktif As String
Public iCounter As Integer

Public sPathAplikasi As String
Public db As ADODB.Connection

Function FileExist(NamaFile As String) As Boolean
  '
  On Error Resume Next
  FileExist = (Dir$(NamaFile) <> "")
  '
End Function

Public Function JumlahRecord(sSQL As String, sDB As ADODB.Connection) As Long
  '
  Dim rsTemp As ADODB.Recordset
  '
  Set rsTemp = New ADODB.Recordset
  rsTemp.Open sSQL, sDB, adOpenStatic, adLockOptimistic
  JumlahRecord = rsTemp.RecordCount
  '
  rsTemp.Close
  Set rsTemp = Nothing
  '
End Function

Sub Main()
  '
  On Error GoTo ErrHandler
  '
  ' Untuk Trial Version (30 Hari Pemakaian)
  'Dim fNum As Integer
  'fNum = FreeFile()
  'If Not FileExist(App.Path & "\Hospitality.dll") Then
  '  Open App.Path & "\Hospitality.dll" For Output As #fNum
  '  Print #fNum, Date + 30
  '  Close #fNum
  'Else
  '  Dim TanggalExpire As String
  '  '
  '  Open App.Path & "\Hospitality.dll" For Input As #fNum
  '  Input #fNum, TanggalExpire
  '  Close #fNum
  '  '
  '  If Date > CDate(TanggalExpire) Then
  '    sPesan = "Masa Pemakaian terbatas dari aplikasi ini telah berakhir" & vbCrLf & vbCrLf
  '    sPesan = sPesan & "Silahkan hubungi PT. Winsor Satria Persada (www.winsor.co.id)" & vbCrLf
  '    sPesan = sPesan & "untuk informasi selanjutnya."
  '    MsgBox sPesan, vbInformation
  '    End
  '  End If
  'End If
  '
  ' Path Database Aplikasi
  Dim fNum As Integer
  '
  fNum = FreeFile()
  If Not FileExist(App.Path & "\Hospitality.fnb") Then
    Open App.Path & "\Hospitality.fnb" For Output As #fNum
    Print #fNum, App.Path
    Close #fNum
  Else
    Open App.Path & "\Hospitality.fnb" For Input As #fNum
    Input #fNum, sPathAplikasi
    Close #fNum
  End If
  '
  ' Buat hubungan dengan database
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & sPathAplikasi & "\Hospitality.mdb;"
  '
  If db.State <> adStateOpen Then
    MsgBox "Gagal dalam membuka database", vbCritical
    End
  Else
    Load frmLogin
    frmLogin.Show
  End If
  Exit Sub

ErrHandler:
  MsgBox "Terjadi masalah dengan koneksi database aplikasi", vbCritical
  frmSetDatabase.Show vbModal
  End
  '
End Sub
