VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmResMaker 
   Caption         =   "Resource Manager"
   ClientHeight    =   3705
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResMaker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   3405
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Key             =   "count"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6191
            Key             =   "msg"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilstMain 
      Left            =   3600
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":1042
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":119C
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":1736
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":1CD0
            Key             =   "save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":1E2A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":21C4
            Key             =   "key"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":255E
            Key             =   "res"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":28F8
            Key             =   "res_add"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":2C92
            Key             =   "res_del"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResMaker.frx":302C
            Key             =   "res_ex"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   5205
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "tbarMain"
      MinHeight1      =   330
      Width1          =   435
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbarMain 
         Height          =   330
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ilstMain"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "add"
               Object.ToolTipText     =   "Add a File"
               ImageKey        =   "res_add"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "remove"
               Object.ToolTipText     =   "Remove a File"
               ImageKey        =   "res_del"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "export"
               Object.ToolTipText     =   "Export to File"
               ImageKey        =   "res_ex"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "open"
               Object.ToolTipText     =   "Open Archive"
               ImageKey        =   "open"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "save"
               Object.ToolTipText     =   "Save Archive"
               ImageKey        =   "save"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "help"
               Object.ToolTipText     =   "Help"
               ImageKey        =   "help"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lvContain 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilstMain"
      ColHdrIcons     =   "ilstMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "rName"
         Text            =   "Name"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "rFilename"
         Text            =   "Filename"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "rLength"
         Text            =   "Length"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileblank0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuResource 
      Caption         =   "&Resource"
      Begin VB.Menu mnuResourceAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuResourceRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuResourceBlank0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResourceExport 
         Caption         =   "Export To File"
      End
      Begin VB.Menu mnuResourceKey 
         Caption         =   "Change Key"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmResMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ResManager As New CResources
Private ArchName As String 'holds the current path of the active resource file

Private Sub RefreshBar()
    'refresh the status bar
    sbar.Panels("count").Text = ResManager.Count & " Item(s)"
    sbar.Panels("msg").Text = GetFileName(ArchName)
End Sub

Private Sub RefreshList()
    'refresh the listview
    Dim res As CResource
    Set res = New CResource
    
    lvContain.ListItems.Clear
    Dim itm As ListItem
    
    For Each res In ResManager
        With lvContain.ListItems
            Set itm = .Add(, , res.ResName, , "res")
            itm.SubItems(1) = res.ResFileName
            itm.SubItems(2) = Format((res.ResLength / 1000), "Standard") & " K"
        End With
    Next
    
    RefreshBar
    
    Set res = Nothing
End Sub

Private Sub Export(sKey As String)
    'export the selected resource to a file
    If opt_UseAppPath = True Then
        ResManager(sKey).ExportToFile BuildPath(App.Path, ResManager(sKey).ResFileName)
    Else
        cdlg.DialogTitle = "Export to File"
        cdlg.InitDir = App.Path
        cdlg.Filter = "All Files (*.*)|*.*"
        cdlg.ShowSave
        
        If cdlg.FileName <> "" Then
            ResManager(sKey).ExportToFile cdlg.FileName
        End If
    End If
End Sub

Private Sub RemoveResource(sKey As String)
    'remove a resource from the archive
    Dim resp As Long
    resp = MsgBox("Are you sure you want to remove " & lvContain.SelectedItem.SubItems(1) & " from the archive?", vbDefaultButton1 + vbYesNo, "Confirm")
    If resp = vbYes Then
        ResManager.Remove sKey
        
        RefreshList
    End If
End Sub

Private Sub AddResource()
    'add a new file to the archive
    Dim res As CResource
    Set res = New CResource
    Dim rName As String
    
    cdlg.DialogTitle = "Add File"
    cdlg.InitDir = App.Path
    cdlg.Filter = "All Files (*.*)|*.*"
    cdlg.ShowSave
    
    If cdlg.FileName <> "" Then
        res.ImportFromFile (cdlg.FileName)
        
        rName = InputBox("Enter Resource Name", "Key", "Key")
        If rName <> "" Then
            res.ResName = rName
            
            Screen.MousePointer = vbHourglass
            On Error GoTo noAdd
            ResManager.Add res
            Screen.MousePointer = vbDefault
        End If
    End If
    
    Set res = Nothing
    
    RefreshList
Exit Sub

noAdd:
    MsgBox "That key has already been used. Can NOT add this resource.", vbCritical, "Error"
    Set res = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub OpenArchive()
    'open an archive file
    ResManager.Clear
    
    cdlg.DialogTitle = "Open Archive"
    cdlg.InitDir = App.Path
    cdlg.Filter = "All Files (*.*)|*.*"
    cdlg.ShowOpen
    
    If cdlg.FileName <> "" Then
        Screen.MousePointer = vbHourglass
        ResManager.Load cdlg.FileName
        ArchName = cdlg.FileName
        RefreshBar
        Screen.MousePointer = vbDefault
        RefreshList
    End If
End Sub

Private Sub SaveAsArchive()
    'save an archive file as...
    cdlg.DialogTitle = "Save As"
    cdlg.InitDir = App.Path
    cdlg.Filter = "All Files (*.*)|*.*"
    cdlg.ShowSave

    If cdlg.FileName <> "" Then
        Screen.MousePointer = vbHourglass
        ResManager.Save cdlg.FileName
        ArchName = cdlg.FileName
        RefreshBar
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub SaveArchive()
    'save an archive file
    If ArchName <> "" Then
        Screen.MousePointer = vbHourglass
        ResManager.Save ArchName
        RefreshBar
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'save last size and positions
    opt_State = Me.WindowState
    opt_Height = Me.Height
    opt_Width = Me.Width
    SaveOptions
End Sub

Private Sub Form_Resize()
    With lvContain
        .Width = Me.Width - 100
        .Height = Me.Height - ((sbar.Height + cbar.Height) + 800)
        .ColumnHeaders("rName").Width = (.Width * 0.3)
        .ColumnHeaders("rFilename").Width = (.Width * 0.4)
        .ColumnHeaders("rLength").Width = (.Width * 0.3) - 100
    End With
End Sub

Private Sub lvContain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    'toggle the sort order for use in the CompareXX routines
    sOrder = Not sOrder
    
    Dim colh As ColumnHeader
    
    'show the up/down arrows in the column header
    For Each colh In lvContain.ColumnHeaders
        If colh Is ColumnHeader Then
            If sOrder = True Then colh.Icon = "up" Else colh.Icon = "down"
        Else
            colh.Icon = 0
            colh.Text = colh.Text
        End If
    Next
    
    'sort normal or by number
    lvContain.SortKey = ColumnHeader.Index - 1
    
    Select Case ColumnHeader.Index - 1
    Case 2
        'Use sort routine to sort by value
        lvContain.Sorted = False
        SendMessage lvContain.hWnd, LVM_SORTITEMS, lvContain.hWnd, ByVal FARPROC(AddressOf CompareValues)
        
    Case Else
        'Use default sorting to sort the items in the list
        lvContain.SortKey = 0
        lvContain.SortOrder = Abs(sOrder)
        lvContain.Sorted = True
        
    End Select
End Sub

Private Sub lvContain_DblClick()
    'export item on double click
    If Not lvContain.SelectedItem Is Nothing Then
        Export lvContain.SelectedItem.Text
    End If
End Sub

Private Sub mnufileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    OpenArchive
End Sub

Private Sub mnuFileSave_Click()
    If ArchName = "" Then
        SaveArchive
    Else
        SaveAsArchive
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveAsArchive
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuResourceAdd_Click()
    AddResource
End Sub

Private Sub mnuResourceExport_Click()
    If Not lvContain.SelectedItem Is Nothing Then
        Export lvContain.SelectedItem.Text
    End If
End Sub

Private Sub mnuResourceRemove_Click()
    If Not lvContain.SelectedItem Is Nothing Then
        RemoveResource lvContain.SelectedItem.Text
    End If
End Sub

Private Sub mnuToolsOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub tbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "help"
        frmAbout.Show vbModal
    Case "open"
        OpenArchive
    Case "save"
        If ArchName = "" Then
            SaveArchive
        Else
            SaveAsArchive
        End If
    Case "add"
        AddResource
    Case "remove"
        If Not lvContain.SelectedItem Is Nothing Then
            RemoveResource lvContain.SelectedItem.Text
        End If
    Case "export"
        If Not lvContain.SelectedItem Is Nothing Then
            Export lvContain.SelectedItem.Text
        End If
    End Select
End Sub
