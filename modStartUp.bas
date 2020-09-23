Attribute VB_Name = "modStartUp"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' EXPLODING FLOWER Screen Saver (start up module)
''' By Paul Bahlawan
'''
''' Most of this module courtesy of About.com
''' http://visualbasic.about.com/library/weekly/
'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public userQty As Long
Public userSize As Long
Public userWidth As Long
Public userPetals As Long
Public userForked As Long
Public userSpeed As Long
Public sMode As String

'Password API functions
Public Declare Sub PwdChangePassword Lib "mpr.dll" _
 Alias "PwdChangePasswordA" (ByVal lpProvider As String, _
 ByVal hwnd As Long, ByVal dwFlags1 As Long, _
 ByVal dwFlags2 As Long)

Public Declare Function VerifyScreenSavePwd Lib "password.cpl" _
 (ByVal hwnd As Long) As Boolean

'Registry Constants and API functions
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_DWORD = 4
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
 Alias "RegOpenKeyExA" (ByVal hKey As Long, _
 ByVal lpSubKey As String, ByVal ulOptions As Long, _
 ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
 Alias "RegQueryValueExA" (ByVal hKey As Long, _
 ByVal lpValueName As String, ByVal lpReserved As Long, _
 lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long


'Screen Saver API functions
Public Const SPI_SCREENSAVERRUNNING = 97

Public Declare Function ShowCursor Lib "user32" _
 (ByVal bShow As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" _
 Alias "SystemParametersInfoA" (ByVal uAction As Long, _
 ByVal uParam As Long, lpvParam As Any, _
 ByVal fuWinIni As Long) As Long


'Window related API functions
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function SetWindowPos Lib "user32" _
 (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
 ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
 ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)
Public Const WS_CHILD = &H40000000

Public Declare Function GetWindowLong Lib "user32" _
 Alias "GetWindowLongA" (ByVal hwnd As Long, _
 ByVal nIndex As Long) As Long
 
Public Declare Function SetWindowLong Lib "user32" _
 Alias "SetWindowLongA" (ByVal hwnd As Long, _
 ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetParent Lib "user32" _
 (ByVal hWndChild As Long, _
 ByVal hWndNewParent As Long) As Long
 
Public Declare Function FindWindow Lib "user32" _
 Alias "FindWindowA" (ByVal lpClassName As String, _
 ByVal lpWindowName As String) As Long
 
Public Declare Function GetActiveWindow Lib "user32" () As Long
 
 
Sub Main()
 Dim lhwndPreview As Long
 Dim lhwndConfig As Long
 Dim lResult As Long
 Dim lWindowStyle As Long
 
 'Get current settings from Windows Registry
 userQty = GetSetting("Exploding_Flower", "Config", "Qty", 5)
 userSize = GetSetting("Exploding_Flower", "Config", "Size", 5)
 userWidth = GetSetting("Exploding_Flower", "Config", "Width", 5)
 userPetals = GetSetting("Exploding_Flower", "Config", "Petals", 5)
 userForked = GetSetting("Exploding_Flower", "Config", "Forked", 5)
 userSpeed = GetSetting("Exploding_Flower", "Config", "Speed", 5)
    
 'Check command line parameters
 Select Case LCase$(Left$(Trim$(Command$()), 2))
  Case "/s", "-s", "s", ""
   sMode = "Exploding Flowers"
   
   lResult = FindWindow(vbNullString, sMode)
   
   If lResult = 0 Then
     frmScreenSaver.Show vbModal
   End If
   
  Case "/c", "-c", "c"
   frmConfig.Show vbModal
  
  Case "/p", "-p", "p"
  
   sMode = "Preview"
   Load frmScreenSaver
   
   lhwndPreview = Val(Right$(Left$(Trim$(Command$()), 7), 4))
   
   lWindowStyle = GetWindowLong(frmScreenSaver.hwnd, GWL_STYLE)
   lResult = SetWindowLong(frmScreenSaver.hwnd, GWL_STYLE, lWindowStyle Or WS_CHILD)
   
   If lResult <> 0 Then
    lResult = SetParent(frmScreenSaver.hwnd, lhwndPreview)
    If lResult <> 0 Then
     frmScreenSaver.Show vbModal
    Else
     Unload frmScreenSaver
    End If
   Else
    Unload frmScreenSaver
   End If
   
  Case "/a", "-a", "a"
   Call ChangePassword
  
  Case Else
   frmConfig.Show vbModeless
 End Select

End Sub

Public Sub ChangePassword()

 Dim lhwndParent As Long
 
 lhwndParent = GetForegroundWindow()
 Call PwdChangePassword("SCRSAVE", lhwndParent, 0, 0)
    
End Sub

Public Function bUsePassword() As Boolean

 Dim lKey As Long
 Dim lKeyData As Long
 Dim lKeyLength As Long
 Dim lKeyType As Long
 Dim sSubKey As String
 Dim sValue As String
 Dim lResult As Long

 bUsePassword = False

 lKeyLength = 4
 lKeyType = REG_DWORD
 sSubKey = "Control Panel\desktop"
 sValue = "ScreenSaveUsePassword"

 lResult = RegOpenKeyEx(HKEY_CURRENT_USER, sSubKey, 0, KEY_READ, lKey)
 If lResult = 0 Then
 
  lResult = RegQueryValueEx(lKey, sValue, 0, lKeyType, lKeyData, lKeyLength)
  If lResult = 0 And lKeyData <> 0 Then
   bUsePassword = True
  End If
  
  lResult = RegCloseKey(lKey)
 End If

End Function
