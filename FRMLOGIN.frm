VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMLOGIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USER LOGIN"
   ClientHeight    =   3210
   ClientLeft      =   6615
   ClientTop       =   2865
   ClientWidth     =   6495
   Icon            =   "FRMLOGIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6495
   Begin VB.ComboBox cbodb 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "MHV"
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CANCEL"
      Height          =   735
      Left            =   5160
      Picture         =   "FRMLOGIN.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOGIN"
      Height          =   735
      Left            =   5160
      Picture         =   "FRMLOGIN.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "&"
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtloginname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox TXTPROCYEAR 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo cbostaffcode 
      Bindings        =   "FRMLOGIN.frx":21BE
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   2280
      TabIndex        =   0
      Top             =   1800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "MODULE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   -360
      Picture         =   "FRMLOGIN.frx":21D3
      Top             =   -120
      Width           =   7575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "LOGIN NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PROCESSING YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATABASE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "FRMLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mpwd As String
Private Hash As New MD5Hash
Private bytBlock() As Byte
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub CBOMODULE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Label3.Visible = True
txtloginname.Visible = True
txtloginname.SetFocus
End If
End Sub

Private Sub CBOMODULE_LostFocus()
Mmodule = CBOMODULE.Text
End Sub

Private Sub cbostaffcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Label3.Visible = True
txtloginname.Visible = True
txtloginname.SetFocus
End If
End Sub

Private Sub cbostaffcode_LostFocus()
cbostaffcode.BackColor = vbWhite
MAINMODULEID = cbostaffcode.BoundText
MAINMODULENAME = cbostaffcode.Text
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select * FROM tblmodulemaster WHERE STATUS='ON' and moduleid='" & cbostaffcode.BoundText & "'", db
If rs.EOF <> True Then
Mlocation = rs!location
Else
MsgBox "Check Module Location."
End If

End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub



Private Sub cmdlogin_Click()
'On Error GoTo err
SysYear = TXTPROCYEAR
If txtloginname.Text = "" Or txtpassword.Text = "" Then
MsgBox "Please Provide Complete Login Information."
Exit Sub
End If

mpwd = ""
   bytBlock = StrConv(LCase(txtpassword.Text), vbFromUnicode)
   mpwd = Hash.HashBytes(bytBlock) ' MD5 Encruption

'CnnsecString = ""
'openDb




'mpwd = hizbiz(UCase(txtpassword))
Dim rs As New ADODB.Recordset
Set rs = Nothing

'rs.Open "select * from user", MHVDB
Set rs = Nothing
rs.Open "select * from tblsoftuser where username='" & txtloginname.Text & "'and password='" & mpwd & "' AND MAINMODULE='" & cbostaffcode.BoundText & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MUSER = txtloginname.Text
UserId = rs!UserId

Unload Me
frmMain.Show
Else
MsgBox "Inavlid Login Credential.!!"
txtloginname.SetFocus
End If
Exit Sub
'err:
'MsgBox "COULD NOT CONNECT TO THE DATABASE, PLEASE TRY AGAIN!"
'txtloginname.SetFocus

End Sub


Private Sub Command1_Click()

'

End Sub

Private Sub Form_Load()
'Shell App.Path + "\connMe.bat ", vbHide
Dim mm, tmp As String
TXTPROCYEAR.Text = Year(Date)
cbostaffcode.Enabled = True
cbostaffcode.BackColor = vbYellow

Dim MsvrName, Mdbname As String
Open App.Path + "\mhv.sys" For Input As #1
'Input #1, FPATH

While EOF(1) = 0
Line Input #1, tmp
FPATH = FPATH + tmp
Wend

Close #1
FPATH = Trim(FPATH)
mm = InStr(FPATH, "*")
If mm > 0 Then
   MsvrName = Trim(Mid(FPATH, 1, mm - 1))
   Mdbname = Trim(Mid(FPATH, mm + 1))
Else
   MsvrName = Trim(FPATH)
   Mdbname = "MHV"
End If
excelPath = "\\192.168.1.12\MhExcelMaster"
'excelPath = App.Path

'CnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
'                        & "SERVER=server_1;" _
'                        & " DATABASE=MHV;" _
'                        & "UID=admin;PWD=password; OPTION=3"
'
'
'OdkCnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
'                        & "SERVER=server_1;" _
'                        & " DATABASE=odk_prodlocal;" _
'                        & "UID=admin;PWD=password; OPTION=3"

                        
                        

CnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & MsvrName & ";" _
                        & " DATABASE=mhv;" _
                        & "UID=admin;PWD=password; OPTION=3"


OdkCnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & MsvrName & ";" _
                        & " DATABASE=odk_prodlocal;" _
                        & "UID=admin;PWD=password; OPTION=3"
                        
MhwebCnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & MsvrName & ";" _
                        & " DATABASE=mhweb;" _
                        & "UID=admin;PWD=password; OPTION=3"
                        
MhvhrCnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & MsvrName & ";" _
                        & " DATABASE=mhvhr;" _
                        & "UID=admin;PWD=password; OPTION=3"
                        
                        
                        
                        
Mserver = MsvrName
ODKDB.Open OdkCnnString
MHWEBDB.Open MhwebCnnString
MHVHRDB.Open MhvhrCnnString
openDb

Dim rs As New ADODB.Recordset

Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open CnnString
Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select concat(modulename, '  ',location) as modulename,moduleid,location FROM tblmodulemaster WHERE STATUS='ON' order by modulename", db
Set cbostaffcode.RowSource = rs
cbostaffcode.ListField = "modulename"
cbostaffcode.BoundColumn = "moduleid"

'cbostaffcode.Text = "REGISTRATION LMT"
MUSER = ""
UserId = ""
End Sub

Private Sub txtloginname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Label4.Visible = True
txtpassword.Visible = True
txtpassword.SetFocus
End If
End Sub

Private Sub txtloginname_LostFocus()
txtloginname.BackColor = vbWhite
End Sub

Private Sub txtpassword_GotFocus()
txtpassword.BackColor = vbYellow
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

cmdlogin.SetFocus
End If
End Sub

Private Sub txtpassword_LostFocus()
txtpassword.BackColor = vbWhite

End Sub
