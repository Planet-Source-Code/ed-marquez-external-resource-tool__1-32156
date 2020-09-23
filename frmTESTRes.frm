VERSION 5.00
Begin VB.Form frmTestRes 
   Caption         =   "Test"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   Icon            =   "frmTESTRes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   735
      Left            =   8640
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load From Res File"
      Height          =   735
      Left            =   8640
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTESTRes.frx":1272
      Height          =   4095
      Left            =   8640
      TabIndex        =   2
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7935
      Left            =   240
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmTestRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    'load no picture to clear it
    Image1.Picture = LoadPicture()
End Sub

Private Sub cmdLoad_Click()
    Dim res As CResources
    Set res = New CResources
    
    'load the archive to use normally just done once in an application globally
    res.Load (App.Path & "\MyRes.ede")
    'extract the file or resource we need in this case a picture
    'call it tmp.bmp for now
    'test1 is the key or resource name given to this item when it was added to the archive
    'we could also use the index number if we know that too
    res("test1").ExportToFile (App.Path & "\tmp.bmp")
    'load from the exported file
    Image1.Picture = LoadPicture(App.Path & "\tmp.bmp")
    'now just delete the file again so the user doesn't see it ot get to it
    res.Delete (App.Path & "\tmp.bmp")
    
    Set res = Nothing
End Sub
