VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)

Private m_ResName As String
Private m_ResLength As Long
Private m_ResData As String
Private m_ResFileName As String
Public Event NameChange(OldName As String, NewName As String, Cancel As Boolean)

Public Property Let ResName(NewValue As String)
    'the collection uses the res name as the key also so make sure its unique
    Dim oldres As String
    Dim docancel As Boolean
    docancel = False
    oldres = m_ResName
    
    m_ResName = NewValue
    
    'raises an event to handle any checks for duplicate keys or to handle the change or anything else
    RaiseEvent NameChange(oldres, NewValue, docancel)
    
    'if in the event the Cancel argument is changed to true then undo changes
    If docancel = True Then m_ResName = oldres
End Property

Public Property Get ResName() As String
    ResName = m_ResName
End Property

Public Property Get ResLength() As Long
    ResLength = m_ResLength
End Property

Public Property Get ResFileName() As String
    ResFileName = m_ResFileName
End Property

Public Property Let Fill(sAll As String)
    'quick fill to be used by the collection
    Dim tmp() As String
    Dim lTotLen As Long
    tmp = Split(sAll, "|")
On Error Resume Next 'just in case
    m_ResName = tmp(0)
    m_ResFileName = tmp(1)
    m_ResLength = CLng(tmp(2))
    lTotLen = 3 + Len(m_ResName) + Len(m_ResFileName) + Len(CStr(m_ResLength)) + 1
    m_ResData = Mid(sAll, lTotLen, m_ResLength)
End Property

Public Property Get Fill() As String
    'quick fill to be used by the collection
On Error Resume Next 'just in case
    Fill = m_ResName & "|"
    Fill = Fill & m_ResFileName & "|"
    Fill = Fill & CStr(m_ResLength) & "|"
    Fill = Fill & m_ResData
End Property

Public Sub ExportToFile(FilePath As String)
    Dim fFileNum As Integer
    fFileNum = FreeFile
    'get data
    Open FilePath For Binary As fFileNum
    Put #fFileNum, , m_ResData
    Close fFileNum
End Sub

Public Sub ImportFromFile(FilePath As String)
    Dim fFileNum As Integer
    fFileNum = FreeFile
    'get data
    
    m_ResFileName = GetFileName(FilePath)
    Open FilePath For Binary As fFileNum
    m_ResLength = LOF(fFileNum)
    m_ResData = Input(LOF(fFileNum), fFileNum)
    Close fFileNum
End Sub

Private Function GetFileName(sFilePath As String) As String
    'takes a path and returns just the filename
    GetFileName = sFilePath
    PathStripPath GetFileName
    GetFileName = Left$(GetFileName, InStr(1, GetFileName, Chr$(0)) - 1)
End Function

