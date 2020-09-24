Attribute VB_Name = "moRegistry"
Option Explicit
Global info

Global g
Global allcharacters
Global molestate()
Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
 
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Const REG_NONE = 0&
Const REG_EXPAND_SZ = 2&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Dim HKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

Const DisplayErrorMsg = False

Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then
      rtn = RegSetValueExA(HKey, Entry, 0, REG_DWORD, Value, 4)
      If Not rtn = ERROR_SUCCESS Then
         If DisplayErrorMsg = True Then
            MesajVer ErrorMsg(rtn), 2
         End If
      End If
      rtn = RegCloseKey(HKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user want errors displayed
         MesajVer ErrorMsg(rtn), 2 'display the error
      End If
   End If
End If

End Function
Function GetDWORDValue(SubKey As String, Entry As String)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegQueryValueExA(HKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(HKey)  'close the key
         GetDWORDValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         GetDWORDValue = "Error"  'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MesajVer ErrorMsg(rtn), 2       'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetDWORDValue = "Error"        'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MesajVer ErrorMsg(rtn), 2       'tell the user what was wrong
      End If
   End If
End If

End Function

Function SetBinaryValue(SubKey As String, Entry As String, Value As String)
  'On Error Resume Next
Dim i
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(Value, i, 1))
      Next
      rtn = RegSetValueExB(HKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         If DisplayErrorMsg = True Then 'if the user want errors displayed
            MesajVer ErrorMsg(rtn), 2       'display the error
         End If
      End If
      rtn = RegCloseKey(HKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MesajVer ErrorMsg(rtn), 2 'display the error
      End If
   End If
End If

End Function

Function GetBinaryValue(SubKey As String, Entry As String)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = RegQueryValueEx(HKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = RegQueryValueEx(HKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(HKey)  'close the key
         GetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetBinaryValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants to errors displayed
            MesajVer ErrorMsg(rtn), 2 'display the error to the user
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetBinaryValue = "Error" 'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants to errors displayed
         MesajVer ErrorMsg(rtn), 2 'display the error to the user
      End If
   End If
End If

End Function
Function DeleteKey(KeyName As String)
  'On Error Resume Next
Call ParseKey(KeyName, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, KeyName, 0, KEY_WRITE, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = RegDeleteKey(HKey, KeyName) 'delete the key
      rtn = RegCloseKey(HKey)  'close the key
   End If
End If

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long
  'On Error Resume Next
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006
   
Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Function ErrorMsg(lErrorCode As Long) As String
  'On Error Resume Next
    Dim GetErrorMsg
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
       Case 2, 1010
            GetErrorMsg = "Bad Key Name"
       Case 1011
            GetErrorMsg = "Can't Open Key"
       Case 4, 1012
            GetErrorMsg = "Can't Read Key"
       Case 5
            GetErrorMsg = "Access to this key is denied"
       Case 1013
            GetErrorMsg = "Can't Write Key"
       Case 8, 14
            GetErrorMsg = "Out of memory"
       Case 87
            GetErrorMsg = "Invalid Parameter"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Undefined Error Code:  " & str$(lErrorCode)
End Select

End Function

Function GetStringValue(SubKey As String, Entry As String)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(HKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(HKey)  'close the key
         sBuffer = Trim(sBuffer)
         GetStringValue = left(sBuffer, Len(sBuffer) - 1) 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetStringValue = "Error" 'return Error to the user
         If DisplayErrorMsg = True Then 'if the user wants errors displayed then
            MesajVer ErrorMsg(rtn), 2 'tell the user what was wrong
         End If
      End If
   Else 'otherwise, if the key couldnt be opened
      GetStringValue = "Error"       'return Error to the user
      If DisplayErrorMsg = True Then 'if the user wants errors displayed then
         MesajVer ErrorMsg(rtn), 2       'tell the user what was wrong
      End If
   End If
End If

End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)
  'On Error Resume Next
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If left(KeyName, 5) <> "HKEY_" Or right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MesajVer "Incorrect Format:" + Chr(10) + Chr(10) + KeyName, 2 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(left(KeyName, rtn - 1)) 'seperate the Keyname
   KeyName = right(KeyName, Len(KeyName) - rtn)
End If

End Sub
Function CreateKey(SubKey As String)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegCreateKey(MainKeyHandle, SubKey, HKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = RegCloseKey(HKey)  'close the key
   End If
End If

End Function
Function SetStringValue(SubKey As String, Entry As String, Value As String)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegSetValueEx(HKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MesajVer ErrorMsg(rtn), 2       'display the error
         End If
      End If
      rtn = RegCloseKey(HKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MesajVer ErrorMsg(rtn), 2       'display the error
      End If
   End If
End If
End Function

Function DeleteValue(SubKey As String, Entry As String)
  'On Error Resume Next
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, HKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = RegDeleteValue(HKey, Entry) 'delete the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user wants errors displayed
            MesajVer ErrorMsg(rtn), 2       'display the error
         End If
      End If
      rtn = RegCloseKey(HKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user wants errors displayed
         MesajVer ErrorMsg(rtn), 2       'display the error
      End If
   End If
End If

End Function
