VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmbarcodeweb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QR CODE PRINTING...."
   ClientHeight    =   8415
   ClientLeft      =   3330
   ClientTop       =   1725
   ClientWidth     =   13845
   Icon            =   "frmbarcodeweb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   13845
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      Picture         =   "frmbarcodeweb.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   5160
      Picture         =   "frmbarcodeweb.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser webbrowser1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      ExtentX         =   23945
      ExtentY         =   13361
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmbarcodeweb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
webbrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Enabled = True
frmbarcodeweb.webbrowser1.Navigate App.Path & "/barcode.html"
End Sub

