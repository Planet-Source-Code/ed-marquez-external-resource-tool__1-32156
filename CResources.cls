VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CResource"
Attribute VB_Ext_KEY = "Member0" ,"CResource"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'api declarations
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'local variable to hold collection
Private mCol As Collection

Public Sub Delete(FilePath As String)
    'made public because it will be used often in the same project as the DLL
    DeleteFile FilePath
End Sub

Public Sub Save(FilePath As String)
    'sae all items in the collection to a single file
    Dim fFileNum As Integer
    Dim rCnt As Long
    Dim rMax As Long
    Dim sTemp As String
    
    fFileNum = FreeFile
    
    'if there is already a file by this name it doesn't delete past the end of the current file
    'so we just delete the old/current one first and remake it
    If FileExists(FilePath) = True Then
        DeleteFile FilePath
    End If
    
    'get data
    Open FilePath For Binary As fFileNum
    'write all items in collection to disk
        rMax = mCol.Count
        For rCnt = 1 To rMax
            'this line is the seperator it will split on when distinguishing items on load
            Put #fFileNum, , "*ITEM*"
            'this is all the properties of the obj in one long string
            sTemp = mCol.Item(rCnt).Fill
            Put #fFileNum, , sTemp
        Next rCnt
    Close fFileNum
End Sub

Public Sub Load(FilePath As String)
    'loads an archive file and and fills all the needed resource items in the collection
    Dim fFileNum As Integer
    Dim AllData As String
    Dim tmp() As String
    Dim tCnt As Long
    Dim tMin As Long
    Dim tMax As Long
    Dim obj As CResource
    Set obj = New CResource
    
    fFileNum = FreeFile
    'get data
    Open FilePath For Binary As fFileNum
    AllData = String(LOF(fFileNum), " ")
    Get #fFileNum, , AllData
    Close fFileNum
    
    'clear the old
    Me.Clear
    
    'fill obj
    'split based on the following line
    tmp = Split(AllData, "*ITEM*")
    tMin = LBound(tmp)
    tMax = UBound(tmp)
    For tCnt = tMin To tMax
        If Len(tmp(tCnt)) > 1 Then
            'fill temporary obj with all the properties then add it to the collection
            obj.Fill = tmp(tCnt)
            Me.Add obj
        End If
    Next tCnt
    
    Set obj = Nothing
End Sub

Public Function Add(ResObj As CResource) As CResource
On Error Resume Next 'just in case
    'create a new object
    Dim objNewMember As CResource
    Set objNewMember = New CResource

    'set the properties passed into the method
    objNewMember.Fill = ResObj.Fill
    
    'use the resource name as the key
    mCol.Add objNewMember, ResObj.ResName

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CResource
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
End Property

Public Sub ChangeKey(oldKey As String, newKey As String)
On Error Resume Next 'just in case
    Dim itm As CResource
    Set itm = New CResource
    'make a temp obj with all the properties of the one being changed
    itm.Fill = mCol(oldKey).Fill
    itm.ResName = newKey
    
    'remove the old one
    mCol.Remove oldKey
    
    'readd it with the new key
    Me.Add itm
    
    Set itm = Nothing
End Sub

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mCol.Remove vntIndexKey
End Sub

Public Sub Clear()
    'clear all items
    
    Dim itm As CResource
    Set itm = New CResource
    
    For Each itm In mCol
        mCol.Remove itm.ResName
    Next
    
    Set itm = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Private Function FileExists(sFilePath As String) As Boolean
    FileExists = CBool(PathFileExists(sFilePath))
End Function

