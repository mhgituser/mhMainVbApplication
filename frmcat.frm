VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CATEGORY"
   ClientHeight    =   2535
   ClientLeft      =   7965
   ClientTop       =   1395
   ClientWidth     =   4905
   ClipControls    =   0   'False
   Icon            =   "frmcat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TXTTFLD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   7
      ToolTipText     =   "Maximum 10 Char"
      Top             =   2160
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DBCombo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Maximum 3 Charactar"
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox TXTTFLD 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   3
      ToolTipText     =   "Maximum 10 Char"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox TXTTFLD 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   2640
      Top             =   2760
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
            Picture         =   "frmcat.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcat.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcat.frx":0E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcat.frx":1B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcat.frx":1FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcat.frx":2784
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
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
   Begin VB.Label lblLabels 
      Caption         =   "Purchase Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sale Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db
Dim DATA1

Private Sub CBOTYPE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
If op = "Add" Then If KeyAscii = 32 Then KeyAscii = 0
End Sub

Private Sub DBCombo1_Validate(Cancel As Boolean)
Dim jg As New ADODB.Recordset
DBCombo1.Text = UCase(DBCombo1.BoundText)
jg.Open "select * from categoryfile where category = '" & DBCombo1.BoundText & "'", db
If jg.EOF Then
   If op = "Open" Then
      MsgBox "This code does not exists !!! "
      Cancel = True
      Exit Sub
   End If
Else
   If Operation = "ADD" Then
      MsgBox "This code already exists !!! "
      Operation = "OPEN"
   End If
   TXTTFLD(0) = IIf(IsNull(jg!Description), "", jg!Description)
   TXTTFLD(1) = IIf(IsNull(jg!acntcode), "", jg!acntcode)
   TXTTFLD(2) = IIf(IsNull(jg!purcode), "", jg!purcode)
   'CBOTYPE.ListIndex = IIf(IsNull(jg!usertype), 0, jg!usertype)
End If
jg.Close

TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = True
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set DATA1 = New ADODB.Recordset
DATA1.Open "SELECT * FROM categoryfile ORDER BY description,CATEGORY", db, adOpenDynamic, adLockReadOnly
Set DBCombo1.RowSource = DATA1
DBCombo1.ListField = "description"
DBCombo1.BoundColumn = "category"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuDelete_Click()
Dim jg As New ADODB.Recordset
If Operation <> "OPEN" Then
   Beep
   Beep
   Exit Sub
End If
If JU < 1 Then
   MsgBox "No rights! Contact System Administrator !"
   Exit Sub
End If
Set Item = New ADODB.Recordset
Item.Open "select * from invitems where category ='" & DBCombo1.BoundText & "'", db
If Not Item.EOF Then
   MsgBox "Category Exists in Item Master ! Cant delete !"
   Exit Sub
End If
If MsgBox("Delete it !!! Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo ERR
db.BeginTrans
db.Execute "delete  from categoryfile where category='" & DBCombo1.BoundText & "'"
db.CommitTrans

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = False
Operation = ""
DBCombo1 = ""
DATA1.Requery
DBCombo1.Refresh
Exit Sub
ERR:
MsgBox ERR.Description
ERR.Clear
db.RollbackTrans
End Sub

Private Sub mnuNew_Click()
Dim i As Long
Operation = "ADD"
For i = 0 To 2
    TXTTFLD(i) = ""
Next


TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = False
If AutoICode Then
    Set lastbill = New ADODB.Recordset
    lastbill.Open "select max(category) as lno from categoryfile", db, adOpenDynamic
    DBCombo1 = IIf(IsNull(lastbill!lno), 100, lastbill!lno + 1)
    Set lastbill = Nothing
    DBCombo1.Enabled = False
 Else
    DBCombo1 = ""
    DBCombo1.Enabled = True
    DBCombo1.SetFocus
 End If

End Sub

Private Sub mnuOpen_Click()
Operation = "OPEN"
DBCombo1 = ""
For i = 0 To 2
    TXTTFLD(i) = ""
Next
DBCombo1.Enabled = True
DBCombo1.SetFocus
End Sub

Private Sub mnuSave_Click()
Dim SQLSTR As String
If Not (Operation = "OPEN" Or Operation = "ADD") Then
   Beep
   MsgBox "No Operation Selected !!!!"
   Exit Sub
End If
On Error GoTo ERR
db.BeginTrans
If Operation = "ADD" Then
   SQLSTR = "insert into categoryfile (category,description,acntcode,purcode) values " _
         & " ('" & DBCombo1.BoundText & "','" & TXTTFLD(0) & "','" & TXTTFLD(1) & "','" & TXTTFLD(2) & "')"
   db.Execute SQLSTR
  ElseIf Operation = "OPEN" Then
   SQLSTR = "update categoryFILE set description='" & TXTTFLD(0) & "',acntcode='" & TXTTFLD(1) & "' " _
         & ",purcode='" & TXTTFLD(2) & "' where  category ='" & DBCombo1.BoundText & "'"
   db.Execute SQLSTR
Else
   Beep
   Beep
   Exit Sub
End If
db.CommitTrans
Operation = ""
DBCombo1 = ""
DATA1.Requery
DBCombo1.Refresh

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = False
Exit Sub
ERR:
MsgBox ERR.Description
ERR.Clear
db.RollbackTrans
End Sub

Private Sub Tb_ButtonClick(ByVal Button As msComctlLib.Button)
Select Case Button.Key
       Case "ADD"
       mnuNew_Click
       Case "OPEN"
       mnuOpen_Click
       Case "SAVE"
       mnuSave_Click
       Case "DELETE"
      ' mnuDelete_Click
       Case "EXIT"
       Unload Me
End Select
End Sub

Private Sub TXTTFLD_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub TXTTFLD_Validate(Index As Integer, Cancel As Boolean)
TXTTFLD(Index) = Trim(TXTTFLD(Index))
If Len(Trim(TXTTFLD(Index))) = 0 Then
   MsgBox "Can't be blank !!!"
   Cancel = True
End If
End Sub
