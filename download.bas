Attribute VB_Name = "moDownload"
'****************************************************************
'Windows API/Global Declarations for :FreeDiskSpace
'****************************************************************
Private Declare Function GetDiskFreeSpace Lib "kernel32" _
                         Alias "GetDiskFreeSpaceA" _
                         (ByVal lpRootPathName As String, _
                          lpSectorsPerCluster As Long, _
                          lpBytesPerSector As Long, _
                          lpNumberOfFreeClusters As Long, _
                          lpTotalNumberOfClusters As Long) As Long
                          
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Function MesajVer(ByVal mMesaj As String, Optional bTip As Byte = 1, Optional iSure As Double = 10) As Long
 Dim sGecici As String, dBaslangic As Date
 
 Select Case bTip
  Case 1    'Göster kapa
   foMesaj.Show
   Do While mMesaj <> ""
    sGecici = mMesaj
    Do While foMesaj.TextWidth(sGecici) > (foMesaj.Width - 400)
     sGecici = Left(sGecici, Len(sGecici) - 1)
    Loop
    foMesaj.mCikti.Caption = foMesaj.mCikti.Caption & sGecici & vbNewLine
    mMesaj = Right(mMesaj, Len(mMesaj) - Len(sGecici))
   Loop
   foMesaj.Height = foMesaj.mCikti.Top + foMesaj.mCikti.Height + 510
   dBaslangic = Now()
   If iSure <= 0 Then iSure = 10
   Do While DateDiff("s", dBaslangic, Now()) < iSure
   Loop
   Unload foMesaj
  Case 2    'Göster bekle
   MsgBox mMesaj, vbOKOnly, "Sistem Mesajý"
  Case 3    'Göster cevap al
   MesajVer = MsgBox(mMesaj, vbYesNo, "Sistem Mesajý")
 End Select
End Function

Function SifreCoz(ByVal sGelenMetin) As String
 Dim sSifreliMetin As String, sHarf1, sHarf2, sHarf3, sHarf4, iDonguSayac
 Dim bKarakterKodu, bSifrelemeTipi As Byte
 
 If IsNull(sGelenMetin) Then GoTo bpBitir
 If sGelenMetin = "" Then GoTo bpBitir
 For iDonguSayac = 1 To Len(sGelenMetin)
  sHarf1 = ""
  sHarf2 = ""
  sHarf3 = ""
  sHarf4 = ""
  bSifrelemeTipi = Asc(Mid(sGelenMetin, iDonguSayac, 1))
  iDonguSayac = iDonguSayac + 1
  Select Case bSifrelemeTipi
  
   'Tip 1
  
   Case 1
    bKarakterKodu = Asc(Mid(sGelenMetin, iDonguSayac, 1)) - 5
    If bKarakterKodu < 0 Then bKarakterKodu = bKarakterKodu + 255
    sHarf1 = Chr(bKarakterKodu)
    If iDonguSayac < Len(sGelenMetin) Then
     iDonguSayac = iDonguSayac + 1
     bKarakterKodu = Asc(Mid(sGelenMetin, iDonguSayac, 1)) - 33
     If bKarakterKodu < 0 Then bKarakterKodu = bKarakterKodu + 255
     sHarf2 = Chr(bKarakterKodu)
    End If
    If iDonguSayac < Len(sGelenMetin) Then
     iDonguSayac = iDonguSayac + 1
     bKarakterKodu = Asc(Mid(sGelenMetin, iDonguSayac, 1)) - 7
     If bKarakterKodu < 0 Then bKarakterKodu = bKarakterKodu + 255
     sHarf3 = Chr(bKarakterKodu)
    End If
    If iDonguSayac < Len(sGelenMetin) Then
     iDonguSayac = iDonguSayac + 1
     bKarakterKodu = Asc(Mid(sGelenMetin, iDonguSayac, 1)) - 97
     If bKarakterKodu < 0 Then bKarakterKodu = bKarakterKodu + 255
     sHarf4 = Chr(bKarakterKodu)
    End If
  
   'Tip 2
  
   Case 2
    bKarakterKodu = Asc(Mid(sGelenMetin, iDonguSayac, 1)) - 38
    If bKarakterKodu < 0 Then bKarakterKodu = bKarakterKodu + 255
    sHarf1 = Chr(bKarakterKodu)
  
   'Tip 3
  
   Case 3
    bKarakterKodu = Asc(Mid(sGelenMetin, iDonguSayac, 1)) + 4
    If bKarakterKodu > 255 Then bKarakterKodu = bKarakterKodu - 255
    sHarf1 = Chr(bKarakterKodu)
    If iDonguSayac < Len(sGelenMetin) Then
     iDonguSayac = iDonguSayac + 1
     bKarakterKodu = RotateRight(Asc(Mid(sGelenMetin, iDonguSayac, 1)))
     sHarf2 = Chr(bKarakterKodu)
    End If
    If iDonguSayac < Len(sGelenMetin) Then
     iDonguSayac = iDonguSayac + 1
     bKarakterKodu = RotateLeft(RotateLeft(Asc(Mid(sGelenMetin, iDonguSayac, 1))))
     sHarf3 = Chr(bKarakterKodu)
    End If
  End Select
  sSifreliMetin = sSifreliMetin & sHarf1 & sHarf2 & sHarf3 & sHarf4
 Next
bpBitir: SifreCoz = sSifreliMetin
End Function

Function RotateLeft(ByVal gelen As Byte) As Byte
  'on error Resume Next
 Dim c As Byte
 If (gelen And 64) > 0 Then
  c = 128
 Else
  c = 0
 End If
 If (gelen And 32) > 0 Then
  c = c + 64
 End If
 If (gelen And 16) > 0 Then
  c = c + 32
 End If
 If (gelen And 8) > 0 Then
  c = c + 16
 End If
 If (gelen And 4) > 0 Then
  c = c + 8
 End If
 If (gelen And 2) > 0 Then
  c = c + 4
 End If
 If (gelen And 1) > 0 Then
  c = c + 2
 End If
 If (gelen And 128) > 0 Then
  c = c + 1
 End If
 RotateLeft = CByte(c)
End Function
Function RotateRight(ByVal gelen As Byte) As Byte
  'on error Resume Next
 Dim c As Byte
 If (gelen And 1) > 0 Then
  c = 128
 Else
  c = 0
 End If
 If (gelen And 128) > 0 Then
  c = c + 64
 End If
 If (gelen And 64) > 0 Then
  c = c + 32
 End If
 If (gelen And 32) > 0 Then
  c = c + 16
 End If
 If (gelen And 16) > 0 Then
  c = c + 8
 End If
 If (gelen And 8) > 0 Then
  c = c + 4
 End If
 If (gelen And 4) > 0 Then
  c = c + 2
 End If
 If (gelen And 2) > 0 Then
  c = c + 1
 End If
 RotateRight = CByte(c)
End Function

Public Function FitText(ByRef Ctl As Control, _
                        ByVal strCtlCaption) As String

' Function FitText
' Author:   Jeff Cockayne
'
' Fit the caption text passed in strCtlCaption
' to the width of the passed Control, Ctl.
' There are a few ways to blow this function, like
' passing a control without a Caption Property, but
' this Function is for internal use, so...
'
' Example:
' If "C:\Program Files\Test.TXT" was too wide, the
' returned string might be: "C:\Pro...\Test.TXT"

Dim lngCtlLeft As Long
Dim lngMaxWidth As Long
Dim lngTextWidth As Long
Dim lngX As Long

' Store frequently referenced values to increase
' performance (saves some OLE lookup)
lngCtlLeft = Ctl.Left
lngMaxWidth = Ctl.Width
lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)


lngX = (Len(strCtlCaption) \ 2) - 2
While lngTextWidth > lngMaxWidth And lngX > 3
    ' Text is too wide for Ctl's width;
    ' shrink the caption from the middle,
    ' replacing the 3 middlemost characters
    ' with ellipses (...)
    strCtlCaption = Left(strCtlCaption, lngX) & "..." & _
                    Right(strCtlCaption, lngX)
    lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
    lngX = lngX - 1
Wend

FitText = strCtlCaption

End Function

Public Function FormatFileSize(ByVal dblFileSize As Double, _
                               Optional ByVal strFormatMask As String) _
                               As String

' FormatFileSize:   Formats dblFileSize in bytes into
'                   X GB or X MB or X KB or X bytes depending
'                   on size (a la Win9x Properties tab)

Select Case dblFileSize
    Case 0 To 1023              ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575        ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823#       ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select

End Function

Public Function FormatTime(ByVal sglTime As Single) As String
                           
' FormatTime:   Formats time in seconds to time in
'               Hours and/or Minutes and/or Seconds

' Determine how to display the time
Select Case sglTime
    Case 0 To 59    ' Seconds
        FormatTime = Format(sglTime, "0") & " sn"
    Case 60 To 3599 ' Minutes Seconds
        FormatTime = Format(Int(sglTime / 60), "#0") & _
                     " dak " & _
                     Format(sglTime Mod 60, "0") & " sn"
    Case Else       ' Hours Minutes
        FormatTime = Format(Int(sglTime / 3600), "#0") & _
                     " saat " & _
                     Format(sglTime / 60 Mod 60, "0") & " dak"
End Select

End Function

Public Function DiskFreeSpace(strDrive As String) As Double

' DiskFreeSpace:    returns the amount of free space on a drive
'                   in Windows9x/2000/NT4+

Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim NumberOfFreeClusters As Long
Dim TotalNumberOfClusters As Long
Dim FreeBytes As Long
Dim spaceInt As Integer

strDrive = QualifyPath(strDrive)

' Call the API function
GetDiskFreeSpace strDrive, _
                 SectorsPerCluster, _
                 BytesPerSector, _
                 NumberOFreeClusters, _
                 TotalNumberOfClusters

' Calculate the number of free bytes
DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector

End Function


Public Function QualifyPath(strPath As String) As String

' Make sure the path ends in "\"
QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")

End Function


Public Function ReturnFileOrFolder(FullPath As String, _
                                   ReturnFile As Boolean, _
                                   Optional IsURL As Boolean = False) _
                                   As String

' ReturnFileOrFolder:   Returns the filename or path of an
'                       MS-DOS file or URL.
'
' Author:   Jeff Cockayne 4.30.99
'
' Inputs:   FullPath:   String; the full path
'           ReturnFile: Boolean; return filename or path?
'                       (True=filename, False=path)
'           IsURL:      Boolean; Pass True if path is a URL.
'
' Returns:  String:     the filename or path
'

Dim intDelimiterIndex As Integer

intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, _
                         Right(FullPath, Len(FullPath) - intDelimiterIndex), _
                         Left(FullPath, intDelimiterIndex))
End If

End Function

