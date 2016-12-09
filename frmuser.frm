VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USER MAINTAINENCE"
   ClientHeight    =   3120
   ClientLeft      =   7575
   ClientTop       =   2070
   ClientWidth     =   5310
   Icon            =   "frmuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5310
   Begin VB.TextBox TXTREPASS 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "&"
      TabIndex        =   8
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox TXTPASS 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "&"
      TabIndex        =   7
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox TXTUSER 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin MSDataListLib.DataCombo cbouser 
      Bindings        =   "frmuser.frx":0E42
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   5280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuser.frx":0E57
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuser.frx":11F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuser.frx":158B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuser.frx":2265
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuser.frx":26B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmuser.frx":2E71
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ADD"
            Key             =   "ADD"
            Object.ToolTipText     =   "ADDS NEW RECORD"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "New"
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Open"
                  Text            =   "Open"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OPEN"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SAVE"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "DELETE"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Key             =   "EXIT"
            Object.ToolTipText     =   "EXIT FROM THE FORM"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
   Begin MSDataListLib.DataCombo CBOMODULE 
      Bindings        =   "frmuser.frx":320B
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cbostaffcode 
      Bindings        =   "frmuser.frx":3220
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   2640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "STAFF NAME"
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
      Top             =   2760
      Width           =   1170
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
      TabIndex        =   10
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "USER ID"
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
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RETYPE PASSWORD"
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
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Label Label2 
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
      TabIndex        =   3
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "USER NAME"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   1110
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUs As New ADODB.Recordset
Private Hash As New MD5Hash
Private bytBlock() As Byte

Private Sub cbouser_LostFocus()
Dim rs As New ADODB.Recordset
CBOUSER.Enabled = False
Set rs = Nothing
rs.Open "SELECT * FROM tblsoftuser WHERE USERID='" & CBOUSER.BoundText & "'", MHVDB
If rs.EOF <> True Then
TXTUSER.Text = rs!UserName
End If
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsUs = Nothing

If rsUs.State = adStateOpen Then rsUs.Close
rsUs.Open "select concat(cast(userid as char) , ' ', username) as username,userid  from tblsoftuser order by userid", db
Set CBOUSER.RowSource = rsUs
CBOUSER.ListField = "username"
CBOUSER.BoundColumn = "userid"
Set rsUs = Nothing

If rsUs.State = adStateOpen Then rsUs.Close
rsUs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff order by STAFFCODE", db
Set cbostaffcode.RowSource = rsUs
cbostaffcode.ListField = "STAFFNAME"
cbostaffcode.BoundColumn = "STAFFCODE"




If UCase(MUSER) = "ADMIN" Then

TXTUSER.Enabled = True

Else

CBOUSER.Text = UserId
CBOUSER.Enabled = False
TXTUSER.Text = MUSER
TXTUSER.Enabled = False

End If

If rs.State = adStateOpen Then rs.Close
rs.Open "select *  from tblmodulemaster ", db
Set CBOMODULE.RowSource = rs
CBOMODULE.ListField = "modulename"
CBOMODULE.BoundColumn = "moduleid"


End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       If UCase(MUSER) = "ADMIN" Then
       CBOUSER.Enabled = False
        TB.Buttons(3).Enabled = True
       operation = "ADD"
       CLEARCONTROLL
       Dim rs As New ADODB.Recordset
       Set rs = Nothing
       rs.Open "SELECT MAX((USERID))+1 AS MaxID from tblsoftuser", MHVDB, adOpenForwardOnly, adLockOptimistic
       If rs.EOF <> True Then
       
       
        CBOUSER.Text = rs!MaxId
        
       Else
       CBOUSER.Text = "100"
       End If
       Label2.Caption = "PASSWORD"
       Else
        operation = "ADD"
       End If
       
       Case "OPEN"
       If UCase(MUSER) = "ADMIN" Then
       operation = "OPEN"
       CLEARCONTROLL
       CBOUSER.Enabled = True
      TB.Buttons(3).Enabled = True
       Else
       operation = "OPEN"
        TB.Buttons(3).Enabled = True
       ' Label2.Caption = "OLD PASSWORD"
       End If
       Case "SAVE"
       MNU_SAVE
        TB.Buttons(3).Enabled = False
     
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub CLEARCONTROLL()
CBOUSER.Text = ""
TXTUSER.Text = ""
TXTPASS.Text = ""
TXTREPASS.Text = ""
End Sub
Private Sub MNU_SAVE()
Dim pwd As String
On Error GoTo err
If Len(cbostaffcode.Text) = 0 Then
MsgBox "You must select staff to save this record."
Exit Sub
End If
If Len(TXTPASS.Text) = 0 Then Exit Sub
If TXTPASS.Text <> TXTREPASS.Text Then
MsgBox "Password Mismatched!"
Exit Sub
End If
If Len(CBOMODULE.Text) = 0 Then
MsgBox "Please Select The Module"
Exit Sub
End If
bytBlock = StrConv(LCase(TXTPASS.Text), vbFromUnicode)
pwd = Hash.HashBytes(bytBlock) ' MD5 Encruption
MHVDB.BeginTrans
If operation = "ADD" Then
MHVDB.Execute "delete from tblmodule where userid='" & CBOUSER.BoundText & "'"
MHVDB.Execute "insert into tblsoftuser (userid,username,password,type,status,mainmodule,staffcode) values('" & CBOUSER.Text & "','" & TXTUSER.Text & "','" & pwd & "','10','ON','" & CBOMODULE.BoundText & "','" & cbostaffcode.BoundText & "')"
AddModule
ElseIf operation = "OPEN" Then


   MHVDB.Execute "update tblsoftuser set username='" & TXTUSER.Text & "',password='" & pwd & "',mainmodule='" & CBOMODULE.BoundText & "',staffcode='" & cbostaffcode.BoundText & "' where userid='" & CBOUSER.BoundText & "' "
Else

End If
MHVDB.CommitTrans


Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans

End Sub
Private Sub AddModule()
Dim SQLSTR As String
SQLSTR = ""
SQLSTR = "insert into tblmodule(moduleid,modulename,userid,userrights,moduletype) (select moduleid,modulename,'" & CBOUSER.Text & "',0,moduletype from tblmodule where userid='100')"
MHVDB.Execute SQLSTR
End Sub
