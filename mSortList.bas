Attribute VB_Name = "mSortList"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mostly taken from Randy Birch's code at http://www.VBNet.com
'A BIG thank you because the listview standard sorting is
'just DAMN annoying!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2001 VBnet, Randy Birch, All Rights Reserved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
  
'variable to hold the sort order (ascending or descending)
Public sOrder As Boolean

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type LV_FINDINFO
  flags As Long
  psz As String
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
'Constants
Public Const LVFI_PARAM As Long = &H1
Public Const LVIF_TEXT As Long = &H1

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
     
'API declarations
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function CompareValues(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long
     
  'CompareValues: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
  
   Dim val1 As Long
   Dim val2 As Long
     
  'Obtain the item names and values corresponding
  'to the input parameters
   val1 = ListView_GetItemValueStr(hWnd, lParam1)
   val2 = ListView_GetItemValueStr(hWnd, lParam2)
     
  'based on the Public variable sOrder set in the
  'columnheader click sub, sort the values appropriately:
   Select Case sOrder
      Case True: 'sort descending
            
            If val1 < val2 Then
                  CompareValues = 0
            ElseIf val1 = val2 Then
                  CompareValues = 1
            Else: CompareValues = 2
            End If
      
      Case Else: 'sort ascending
   
            If val1 > val2 Then
                  CompareValues = 0
            ElseIf val1 = val2 Then
                  CompareValues = 1
            Else: CompareValues = 2
            End If
   
   End Select

End Function



Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long) As Long

   Dim hIndex As Long
   Dim r As Long
  
  'Convert the input parameter to an index in the list view
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)
     
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.mask = LVIF_TEXT
   objItem.iSubItem = 2
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
  'get the string at subitem 2
  'and convert it into a long
   r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
   If r > 0 Then
      'i had to change it to remove the 'K' at the end
      ListView_GetItemValueStr = CLng(Left$(objItem.pszText, r - 1))
   End If

End Function

Public Function FARPROC(ByVal pfn As Long) As Long
  
  'A procedure that receives and returns
  'the value of the AddressOf operator.
  'This workaround is needed as you can't assign
  'AddressOf directly to an API when you are also
  'passing the value ByVal in the statement
  '(as is being done with SendMessage)
 
  FARPROC = pfn

End Function


