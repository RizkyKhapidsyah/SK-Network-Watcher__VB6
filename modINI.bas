Attribute VB_Name = "modINI"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                "GetPrivateProfileStringA" (ByVal lpApplicationName _
                As String, ByVal lpKeyName As Any, ByVal lpDefault _
                As String, ByVal lpReturnedString As String, ByVal _
                nsize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                "WritePrivateProfileStringA" (ByVal lpApplicationName _
                As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
                ByVal lpFileName As String) As Long

Public Function sGetINI(sINIFile As String, sSection As String, sKey _
            As String, sDefault As String) As String
            
Dim sTemp As String * 256
Dim nLength As Integer

sTemp = Space$(256)
nLength = GetPrivateProfileString(sSection, sKey, sDefault, _
            sTemp, 255, sINIFile)
sGetINI = Left$(sTemp, nLength)
End Function

Public Sub WriteINI(sINIFile As String, sSection As String, _
            sKey As String, sValue As String)
            
Dim n As Integer
Dim sTemp As String

sTemp = sValue

'vervang CR/LF door spaties
For n = 1 To Len(sValue)
    If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf _
    Then Mid$(sValue, n) = " "
Next n

n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
            
End Sub

Public Sub SearchINIfile()
Dim cnt As Byte
Dim fso As Object
 Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists((CurDir & "\" & "SettingsNet.INI"))) Then
      'File Exist...no problem.. error
   Else
      MsgBox "This is maybe the first time that you use this application ! " & vbCrLf & " The settings values will be default ..." & _
      vbCrLf & "The next time it will use the previous settings ..", vbInformation
       'Create an ini file with settings up to 10 adaptors .... including the 0.0.0.0
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Info", "Creator", "Peter Verburgh"
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Info", "Email", "http://users.skynet.be/verburgh.peter"
       For cnt = 1 To 10
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "MaxUploadSpeed", 15
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "MaxDownloadSpeed", 100
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "TotalRecieved", 0
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "TotalSend", 0
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "ShowWindow", 1
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "posX", 500
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "posY", 500 + (cnt * 200)
       modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(cnt), "Skin", "ANALOGE"
       Next
   End If
End Sub

Sub WriteSettings(ByVal Adaptor As Integer, SubItem As String, data As String)
    modINI.WriteINI CurDir & "\" & "SettingsNet.INI", "Adaptor" & CStr(Adaptor), SubItem, data
End Sub
