VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form foMain 
   Caption         =   "Norton Daily Antivirus Files Automatic Update Application"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "foMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.CheckBox tAcilistaCalissin 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   4260
      Width           =   255
   End
   Begin VB.CommandButton buKontrolEt 
      Caption         =   "UPDATE NOW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox sGuncellemeZamani 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton buGozat 
      Caption         =   "&BROWSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox sNortonKlasoru 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
   Begin InetCtlsObjects.Inet inBaglanti 
      Left            =   3240
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tiZamanlayici 
      Interval        =   1000
      Left            =   4200
      Top             =   4200
   End
   Begin VB.TextBox mStatu 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "foMain.frx":0ECA
      Top             =   2520
      Width           =   6375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Run when user logs in."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   4290
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UPDATE TIME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ANTIVIRUS FILES FOLDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6375
   End
   Begin VB.Image imEbit 
      Height          =   735
      Left            =   5520
      Picture         =   "foMain.frx":0ED5
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1020
   End
End
Attribute VB_Name = "foMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dZaman As Date

Private Sub buGozat_Click()
 foBrowse.Show vbModal
End Sub

Private Sub buKontrolEt_Click()
 Guncelle
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  KeyAscii = 0
  SendKeys "{TAB}"
 End If
End Sub

Private Sub Form_Load()

 'Check for setup file and get saved parameters if possible
 
 On Error GoTo DosyaYok
 Dim bBosNo As Byte, sKlasor As String, sZaman As String, sAcilis As String
 bBosNo = FreeFile
 Open App.Path & "\setup.txt" For Input As bBosNo
 Input #bBosNo, sKlasor
 Input #bBosNo, sZaman
 Input #bBosNo, sAcilis
 Close bBosNo
 sNortonKlasoru = sKlasor
 sGuncellemeZamani = sZaman
 tAcilistaCalissin = Val(sAcilis)
 Exit Sub
 
DosyaYok: Resume DevamEt
DevamEt: Close

 'If setup.txt file is not available create it with hard coded default values
 
 Open App.Path & "\setup.txt" For Output As bBosNo
 Write #bBosNo, "C:\Program Files\Common Files\Symantec Shared\VirusDefs\"
 Write #bBosNo, "12:00"
 Write #bBosNo, "1"
 Close bBosNo
 sNortonKlasoru = "C:\Program Files\Common Files\Symantec Shared\VirusDefs\"
 sGuncellemeZamani = "09:00"
 tAcilistaCalissin = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim bBosNo As Byte
 bBosNo = FreeFile
 
 'Save setup
 
 Open App.Path & "\setup.txt" For Output As bBosNo
 Write #bBosNo, sNortonKlasoru
 Write #bBosNo, sGuncellemeZamani
 Write #bBosNo, tAcilistaCalissin
 Close bBosNo
 
 'Depending on the choice of user add a registry key to run application after logon or delete it from registry.
 
 If tAcilistaCalissin Then
  SetStringValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "EbitNortonAutomaticUpdate", App.Path & "\NortonAutomaticUpdate.exe"
 Else
  DeleteValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run", "EbitNortonAutomaticUpdate"
 End If
End Sub

Private Sub imEbit_Click()
 Shell "explorer.exe http://www.e-bit.com.tr", vbMaximizedFocus
End Sub

Private Sub sGuncellemeZamani_LostFocus()
 
 'Check the format of update time
 
 If sGuncellemeZamani = "" Then
  sGuncellemeZamani = "09:00"
 Else
  If IsDate(sGuncellemeZamani) Then
   sGuncellemeZamani = Format(Hour(sGuncellemeZamani), "00") & ":" & Format(Minute(sGuncellemeZamani), "00")
  Else
   MsgBox "Please enter a valid time. Example: 03:15", vbCritical, "Warning"
   sGuncellemeZamani.SetFocus
  End If
 End If
End Sub

Private Sub tiZamanlayici_Timer()

 'Check if update time has come or not
 
 If (sNortonKlasoru = "") Or Not IsDate(sGuncellemeZamani) Then Exit Sub
 dZaman = Now()
 If Format(Hour(dZaman), "00") & ":" & Format(Minute(dZaman), "00") = sGuncellemeZamani Then Guncelle
End Sub

Sub Guncelle()
 Dim sHTML As String, sURL As String, sAntiVirusDosyasi As String, sDosyaBoyu As String
 Dim sYaratilmaTarihi As String, sDagitimTarihi As String, poUygulama As Long, sDosya As String
 Dim obDosyaSistemi As Object
 
 'Open the web page from Symantec's site to get latest update file information
 
 sURL = "http://securityresponse.symantec.com/avcenter/download/pages/US-N95.html"
 sHTML = inBaglanti.OpenURL(sURL)
 DoEvents
 mStatu = "Update started : " & Now() & vbNewLine & mStatu
 If sHTML <> "" Then
  If InStr(1, sHTML, ".EXE", vbTextCompare) > 0 Then
   sAntiVirusDosyasi = Left(sHTML, InStr(1, sHTML, ".EXE", vbTextCompare) + 3)
   sAntiVirusDosyasi = Right(sAntiVirusDosyasi, Len(sAntiVirusDosyasi) - InStrRev(sAntiVirusDosyasi, Chr(34)))
   sAntiVirusDosyasi = "http://securityresponse.symantec.com" & IIf(Left(sAntiVirusDosyasi, 1) = "/", "", "/") & sAntiVirusDosyasi
   sYaratilmaTarihi = Mid(sHTML, InStr(1, sHTML, ".EXE", vbTextCompare) + 4)
   sYaratilmaTarihi = Mid(sYaratilmaTarihi, InStr(1, sYaratilmaTarihi, "<TD><P>", vbTextCompare) + 7)
   sYaratilmaTarihi = Left(sYaratilmaTarihi, InStr(1, sYaratilmaTarihi, "</P>", vbTextCompare) - 1)
   sDagitimTarihi = Mid(sHTML, InStr(1, sHTML, sYaratilmaTarihi, vbTextCompare) + 4)
   sDagitimTarihi = Mid(sDagitimTarihi, InStr(1, sDagitimTarihi, "<TD><P>", vbTextCompare) + 7)
   sDagitimTarihi = Left(sDagitimTarihi, InStr(1, sDagitimTarihi, "</P>", vbTextCompare) - 1)
   sDosyaBoyu = Mid(sHTML, InStr(1, sHTML, sYaratilmaTarihi, vbTextCompare) + 4)
   sDosyaBoyu = Mid(sDosyaBoyu, InStr(1, sDosyaBoyu, sDagitimTarihi, vbTextCompare) + 4)
   sDosyaBoyu = Mid(sDosyaBoyu, InStr(1, sDosyaBoyu, "<TD><P>", vbTextCompare) + 7)
   sDosyaBoyu = Left(sDosyaBoyu, InStr(1, sDosyaBoyu, "</P>", vbTextCompare) - 1)
   mStatu = "Antivirus file : " & sAntiVirusDosyasi & vbNewLine & mStatu
   mStatu = "Create date : " & sYaratilmaTarihi & vbNewLine & mStatu
   mStatu = "Release date : " & sDagitimTarihi & vbNewLine & mStatu
   mStatu = "File size : " & sDosyaBoyu & vbNewLine & mStatu
   mStatu = "Downloading ..." & vbNewLine & mStatu
   
   'Download the latest file
   
   If foDownload.DownloadFile(sAntiVirusDosyasi, App.Path & "\antivirus.exe") Then
    mStatu = "File downloaded." & vbNewLine & mStatu
    mStatu = "Opening file contents ..." & vbNewLine & mStatu
    If Dir(App.Path & "\antivirus\ZDONE.DAT") <> "" Then Kill App.Path & "\antivirus\ZDONE.DAT"
    If Dir(App.Path & "\antivirus", vbDirectory) = "" Then MkDir App.Path & "\antivirus"
    poUygulama = Shell("""" & App.Path & "\antivirus.exe"" /extract """ & App.Path & "\antivirus""", vbNormalNoFocus)
    Do While Dir(App.Path & "\antivirus\ZDONE.DAT") = ""
     DoEvents
    Loop
    mStatu = "Opened file content." & vbNewLine & mStatu
    SendKeys "{ENTER}", True
    Me.SetFocus
    
    'Copying new files
    
    sDosya = Dir(App.Path & "\antivirus\*.*")
    Set obDosyaSistemi = CreateObject("Scripting.FileSystemObject")
    mStatu = "Copying files ..." & vbNewLine & mStatu
    Do
     If UCase(sDosya) <> "ZDONE.DAT" Then 'Save ZDONE.DAT to last because it should be copied after all other files are copied. Otherwise errors may occur.
      obDosyaSistemi.Copyfile App.Path & "\antivirus\" & sDosya, sNortonKlasoru & IIf(Right(sNortonKlasoru, 1) = "\", "", "\") & "Incoming\" & sDosya
      mStatu = sDosya & vbNewLine & mStatu
     End If
     sDosya = Dir()
    Loop Until sDosya = ""
    
    'Save ZDONE.DAT now to inform Norton that new files are put into incoming files folder
    
    obDosyaSistemi.Copyfile App.Path & "\antivirus\zdone.dat", sNortonKlasoru & IIf(Right(sNortonKlasoru, 1) = "\", "", "\") & "Incoming\" & "zdone.dat"
    Set obDosyaSistemi = Nothing
    mStatu = "Update operation is completed." & vbNewLine & mStatu
   Else
    mStatu = "File couldn't be downloaded." & vbNewLine & mStatu
   End If
   mStatu = "Last update time : " & Now() & vbNewLine & mStatu
   mStatu = "Ready..." & vbNewLine & mStatu
  End If
 End If
End Sub
