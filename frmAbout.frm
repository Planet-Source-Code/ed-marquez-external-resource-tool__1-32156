VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "lblCopyright"
      Height          =   855
      Left            =   1080
      TabIndex        =   4
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblDescrip 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDescrip"
      Height          =   555
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   3945
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3930
      TabIndex        =   2
      Top             =   360
      Width           =   1140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1080
      X2              =   5040
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   1080
      X2              =   5040
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblAppTitle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":1042
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblTitle.Caption = App.Title
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & " (build " & App.Revision & ")"
    lblDescrip.Caption = App.FileDescription
    lblCopyright.Caption = App.CompanyName & vbNewLine & App.LegalCopyright & vbNewLine & vbNewLine & App.LegalTrademarks
End Sub
