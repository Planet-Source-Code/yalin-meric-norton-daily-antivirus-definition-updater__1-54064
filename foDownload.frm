VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form foDownload 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "foDownload.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   645
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label RateLabel 
      AutoSize        =   -1  'True
      Caption         =   "RateLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2325
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label TransferRate 
      AutoSize        =   -1  'True
      Caption         =   "Transfer Rate :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label ToLabel 
      AutoSize        =   -1  'True
      Caption         =   "ToLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2325
      TabIndex        =   6
      Top             =   1260
      Width           =   585
   End
   Begin VB.Label DownloadTo 
      AutoSize        =   -1  'True
      Caption         =   "Target Folder :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   1035
   End
   Begin VB.Label TimeLabel 
      AutoSize        =   -1  'True
      Caption         =   "TimeLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2325
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label SourceLabel 
      Caption         =   "SourceLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5250
   End
   Begin VB.Label EstimatedTimeLeft 
      AutoSize        =   -1  'True
      Caption         =   "Estimated Time Left :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label StatusLabel 
      Caption         =   "StatusLabel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5235
   End
End
Attribute VB_Name = "foDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CancelSearch As Boolean


Public Function DownloadFile(strURL As String, _
                             strDestination As String, _
                             Optional tTip As Boolean = False, _
                             Optional UserName As String = Empty, _
                             Optional Password As String = Empty) _
                             As Boolean

' Funtion DownloadFile: Download a file via HTTP
'
' Author:   Jeff Cockayne
'
' Inputs:   strURL String; the source URL of the file
'           strDestination; valid Win95/NT path to where you want it
'           (i.e. "C:\Program Files\My Stuff\Purina.pdf")
'
' Returns:  Boolean; Was the download successful?

'Const CHUNK_SIZE As Long = 1024 ' Download chunk size
Const CHUNK_SIZE As Long = 4096 ' Download chunk size
Const ROLLBACK As Long = 8192   ' Bytes to roll back on resume
                                ' You can be less conservative,
                                ' and roll back less, but I
                                ' don't recommend it.
Dim bData() As Byte             ' Data var
Dim blnResume As Boolean        ' True if resuming download
Dim intFile As Integer          ' FreeFile var
Dim lngBytesReceived As Long    ' Bytes received so far
Dim lngFileLength As Long       ' Total length of file in bytes
Dim lngX                        ' Temp long var
Dim sglLastTime As Single          ' Time last chunk received
Dim sglRate As Single           ' Var to hold transfer rate
Dim sglTime As Single           ' Var to hold time remaining
Dim strFile As String           ' Temp filename var
Dim strHeader As String         ' HTTP header store
Dim strHost As String           ' HTTP Host
Dim sGelen() As String

'on local error GoTo InternetErrorHandler

' Start with Cancel flag = False
CancelSearch = False

' Get just filename (without dirs) for display
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
              
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty

' Show the download status form
Show
' Move form into view
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

StartDownload:

If blnResume Then
    StatusLabel = "Download proceeding..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Getting file information..."
End If
' Give the system time to update the form gracefully
DoEvents

' Download file
With Inet1
    .URL = strURL
    .UserName = UserName
    .Password = Password
    ' GET file, sending the magic resume input header...
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    ' While initiating connection, yield CPU to Windows
    While .StillExecuting
        DoEvents
        ' If user pressed Cancel button on StatusForm
        ' then fail, cancel, and exit this download
        If CancelSearch Then GoTo ExitDownload
    Wend

    StatusLabel = "Downloading : "
    SourceLabel = FitText(SourceLabel, "Downloading " & strHost & " from " & .RemoteHost)
    ToLabel = FitText(ToLabel, strDestination)

    ' Get first header ("HTTP/X.X XXX ...")
    strHeader = .GetHeader
End With

' Trap common HTTP response codes
Select Case Mid(strHeader, 10, 3)
    Case "200"  ' OK
        ' If resuming, however, this is a failure
        If blnResume Then
            ' Delete partially downloaded file
            Kill strDestination
            ' Prompt
            If MesajVer("Server cannot continue download job. Do you still want to continue?", 3) = vbYes Then
                    ' Yes - continue anyway:
                    ' Set resume flag to False
                    blnResume = False
                Else
                    ' No - cancel
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
            
    Case "206"  ' 206=Partial Content, which is GREAT when resuming!
    
    Case "204"  ' No content
        MesajVer "The file to download couldn't be found!", 2
        CancelSearch = True
        GoTo ExitDownload
        
    Case "401"  ' Not authorized
        MesajVer "Insufficient authorization!", 2
        CancelSearch = True
        GoTo ExitDownload
    
    Case "404"  ' File Not Found
        MesajVer "File not found!", 2
        CancelSearch = True
        GoTo ExitDownload
        
    Case vbCrLf ' Empty header
        MesajVer "Connection couldn't be made. Check your internet connection and try again.", 2
        CancelSearch = True
        GoTo ExitDownload
        
    Case Else
        ' Miscellaneous unexpected errors
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MesajVer "Server reported an error.", 2
        CancelSearch = True
        GoTo ExitDownload
End Select

' Get file length with "Content-Length" header request
If blnResume = False Then
    ' Set timer for gauging download speed
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

' Check for available disk space first...
' If on a physical or mapped drive. Can't with a UNC path.
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, _
                          InStr(strDestination, "\"))) < lngFileLength Then
        ' Not enough free space to download file
        MesajVer "Ther is not enough space in hard disk to svae this file. Please empty some space from the hard disk and try again.", 2
        GoTo ExitDownload
    End If
End If

' Prepare display
'
' Progress Bar
With ProgressBar
    .Value = 0
    .Max = lngFileLength
End With

' Give system a chance to show AVI
DoEvents

' Reset bytes received counter if not resuming
If blnResume = False Then lngBytesReceived = 0


'on local error GoTo FileErrorHandler

' Create destination directory, if necessary
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If

' If no errors occurred, then spank the file to disk
intFile = FreeFile()        ' Set intFile to an unused file.
' Open a file to write to.

'ASCII dosya ise
If tTip Then
 sGelen = Inet1.OpenURL(strURL, icString)
 Open strDestination For Output As #intFile
 'Print #1, sGelen()
 Close #intFile
Else
 Open strDestination For Binary Access Write As #intFile
 ' If resuming, then seek byte position in downloaded file
 ' where we last left off...
 If blnResume Then Seek #intFile, lngBytesReceived + 1
 Do
    ' Get chunks...
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    If UBound(bData, 1) > 0 Then Put #intFile, , bData   ' Put it into our destination file
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & _
                   " (" & _
                   FormatFileSize(lngBytesReceived) & _
                   " / " & _
                   FormatFileSize(lngFileLength) & _
                   " copied)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
    ProgressBar.Value = lngBytesReceived
    Me.Caption = strFile & " " & Format((lngBytesReceived / lngFileLength), "##0%") & _
                 " downloaded."
 Loop While UBound(bData, 1) > 0       ' Loop while there's still data...
 Close #intFile
End If

ExitDownload:
' Success if the # of bytes transferred = content length
If lngBytesReceived = lngFileLength Then
    StatusLabel = "Download complete!"
    DownloadFile = True
Else
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        ' Resume? (If not cancelled)
        If CancelSearch = False Then
            If MsgBox("Server connection is lost." & _
                      vbCr & vbCr & _
                      "Click ""Try Again"" button to try downloading again." & _
                      vbCr & "(Estimated time remaining: " & FormatTime(sglTime) & ")" & _
                      vbCr & vbCr & _
                      "Click ""Cancel"" button to stop downloading.", _
                      vbExclamation + vbRetryCancel, _
                      "Download Not Complete") = vbRetry Then
                    ' Yes
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    ' No or unresumable failure:
    ' Delete partially downloaded file
    If Not Dir(strDestination) = Empty Then Kill strDestination
    DownloadFile = False
End If

CleanUp:

' Make sure that the Internet connection is closed...
Inet1.Cancel
' ...and exit this function
Unload Me

Exit Function

InternetErrorHandler:
    ' Err# 9 occurs when UBound(bData,1) < 0
    If Err.Number = 9 Then Resume Next
    ' Other errors...
    MsgBox "Error occured: " & Err.Description, _
           vbCritical, _
           "File Download Error"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:
    MsgBox "File cannot br written to hard disk." & _
           vbCr & vbCr & _
           "Error: " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "File Download Error"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
    
End Function

Private Sub Inet1_StateChanged(ByVal State As Integer)
 Debug.Print State
End Sub

