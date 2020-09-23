Attribute VB_Name = "mResMaker"
Option Explicit

'declares for the BuildPath and GetFileName functions
Private Declare Function PathAppend Lib "shlwapi.dll" Alias "PathAppendA" (ByVal pszPath As String, ByVal pMore As String) As Long
Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)

'obj for working with the registry and the default reg location
Public Reg As New CReg
Public Const USE_REG As String = "Software\BUILT4U\ResMaker\"

'global option variables
Public opt_UseAppPath As Boolean
Public opt_State As Long
Public opt_Height As Long
Public opt_Width As Long


Sub Main()
    'get options then show form
    Reg.DefaultRegEntry = USE_REG
    GetOptions
    frmResMaker.Show
On Error Resume Next
    'make form appear as last state and size
    frmResMaker.WindowState = opt_State
    frmResMaker.Height = opt_Height
    frmResMaker.Width = opt_Width
    
End Sub

Public Function BuildPath(ByVal Path As String, ByVal Name As String) As String
    'works the same as the FileSystemObject BuildPath method
    BuildPath = Path & String(255, 0)
    PathAppend BuildPath, Name
    BuildPath = Left$(BuildPath, InStr(BuildPath, Chr$(0)) - 1)
End Function

Public Function GetFileName(sFilePath As String) As String
    'gets the filename from a path
    GetFileName = sFilePath
    PathStripPath GetFileName
    GetFileName = Left$(GetFileName, InStr(1, GetFileName, Chr$(0)) - 1)
End Function

Public Sub GetOptions()
    'gets the options from the registry
    opt_UseAppPath = CBool(Reg.GetRegSetting("UseAppPath", "True"))
    opt_State = CLng(Reg.GetRegSetting("State", "0"))
    opt_Height = CLng(Reg.GetRegSetting("Height", "4395"))
    opt_Width = CLng(Reg.GetRegSetting("Width", "5325"))
End Sub

Public Sub SaveOptions()
    'save the options to the registry
    Reg.SaveRegSetting "UseAppPath", CStr(opt_UseAppPath)
    Reg.SaveRegSetting "State", CStr(opt_State)
    Reg.SaveRegSetting "Height", CStr(opt_Height)
    Reg.SaveRegSetting "Width", CStr(opt_Width)
End Sub


