VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is my default class for anytime I want to use the registry

'declares for registry
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Public Enum HKeyConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Const REG_SZ = 1 'Unicode nul terminated string
Const REG_BINARY = 3 'Free form binary
Const REG_DWORD = 4 '32-bit number
Const ERROR_SUCCESS = 0&

Const SOUND_EVENTS As String = "AppEvents\Schemes\Apps\"
Const DEF_SOUND_EVENTS As String = "AppEvents\Schemes\Apps\.Default\"
Const RUN_ENTRY As String = "Software\Microsoft\Windows\CurrentVersion\Run"
Const WINDOWS_DEF_SND As String = ".Default"

'declares for playing sounds
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Const SND_ALIAS_START = 0
Const SND_ASYNC = &H1

Public Enum SoundEventConstants
    sndDefault
    sndAsterisk
    sndExclamation
    sndExit
    sndQuestion
    sndStart
    sndOpen
    sndClose
    sndMaximize
    sndMinimize
    sndAppGPFault
End Enum

Private m_DefaultRegEntry As String

Public Property Let DefaultRegEntry(NewValue As String)
    m_DefaultRegEntry = NewValue
End Property

Public Property Get DefaultRegEntry() As String
    DefaultRegEntry = m_DefaultRegEntry
End Property

Public Sub DeleteRegValue(ByVal strPath As String, ByVal strValue As String, Optional hKey As HKeyConstants = HKEY_LOCAL_MACHINE)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    
    lRegResult = RegDeleteValue(hCurKey, strValue)
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DeleteRegKey(ByVal strPath As String, Optional hKey As HKeyConstants = HKEY_LOCAL_MACHINE)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    
    lRegResult = RegDeleteKey(hCurKey, "")
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetRegSetting(strValue As String, Optional Default As String, Optional strPath As String, Optional hKey As HKeyConstants = HKEY_LOCAL_MACHINE) As String
    Dim hCurKey As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    
    If strPath = "" Then
        strPath = m_DefaultRegEntry
    End If
    
    'Set up default value
    If Not IsEmpty(Default) Then
        GetRegSetting = Default
    Else
        GetRegSetting = ""
    End If
    
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
        
        If lValueType = REG_SZ Then
        
            strBuffer = String(lDataBufferSize, " ")
            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            
            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetRegSetting = Left$(strBuffer, intZeroPos - 1)
            End If
        Else
            
            strBuffer = String(lDataBufferSize, 0)
            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
            If (lResult = ERROR_NONE) Then
            
                strBuffer = Left$(strBuffer, lDataBufferSize - 1)
                strBuffer2 = String(lDataBufferSize + 400, 0)
                lResult = ExpandEnvironmentStrings(strBuffer, strBuffer2, lDataBufferSize + 400)
                If (lResult > 0) Then
                    GetRegSetting = Left$(strBuffer2, lResult)
                End If
                
            Else
                'there is a problem
            End If

        End If
        
    Else
        'there is a problem
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveRegSetting(strValue As String, strData As String, Optional strPath As String, Optional hKey As HKeyConstants = HKEY_LOCAL_MACHINE)
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    If strPath = "" Then
        strPath = m_DefaultRegEntry
    End If
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
    ByVal strData, Len(strData))
    
    If lRegResult <> ERROR_SUCCESS Then
        'there is a problem
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Private Function GetSoundReg(strKey As String) As String
    hKey = HKEY_CURRENT_USER
    strPath = SOUND_EVENTS & strKey & "\.Current"
    strValue = ""
End Function

Public Sub PlayWindowsSound(SoundEvent As SoundEventConstants)
    Dim res As Long
    Dim strPath As String
    
    strPath = GetRegSetting("", , SOUND_EVENTS & ".Default\" & Cnv2Name(SoundEvent) & "\.Current", HKEY_CURRENT_USER)
    res = PlaySound(strPath, 0, SND_ASYNC)
End Sub

Public Sub PlayCustomSound(AppName As String, EventName As String)
    Dim res As Long
    Dim strPath As String
    
    strPath = GetRegSetting("", , SOUND_EVENTS & AppName & "\" & EventName & "\.Current", HKEY_CURRENT_USER)
    res = PlaySound(strPath, 0, SND_ASYNC)
End Sub

Public Sub SaveSoundEvent(FilePath As String, EventName As String, Optional AppName As String = WINDOWS_DEF_SND)
    SaveRegSetting "", FilePath, SOUND_EVENTS & AppName & "\" & EventName & "\.Current", HKEY_CURRENT_USER
End Sub

Public Function GetSoundEvent(EventName As String, Optional AppName As String = WINDOWS_DEF_SND)
    GetSoundEvent = GetRegSetting("", , SOUND_EVENTS & AppName & "\" & EventName & "\.Current", HKEY_CURRENT_USER)
End Function

Public Sub DeleteSoundEvent(EventName As String, Optional AppName As String = WINDOWS_DEF_SND)
    DeleteRegKey SOUND_EVENTS & AppName & "\" & EventName, HKEY_CURRENT_USER
End Sub

Public Sub SaveSoundEventApp(AppName As String, Optional AppDescription As String)
    SaveRegSetting "", AppDescription, SOUND_EVENTS & AppName, HKEY_CURRENT_USER
End Sub

Public Function GetSoundEventApp(AppName As String)
    GetSoundEventApp = GetRegSetting("", , SOUND_EVENTS & AppName, HKEY_CURRENT_USER)
End Function

Public Sub DeleteSoundEventApp(AppName As String)
    DeleteRegKey SOUND_EVENTS & AppName, HKEY_CURRENT_USER
End Sub

Private Function Cnv2Name(EnumValue As SoundEventConstants) As String
    Dim SndEvents As Variant

    SndEvents = Array(".Default", "SystemAsterisk", "SystemExclamation", "SystemExit", "SystemQuestion", "SystemStart", "Open", "Close", "Maximize", "Minimize", "AppGPFault")
    Cnv2Name = CStr(SndEvents(EnumValue))
End Function
