VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSupplMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SUPPLIER MASTER"
   ClientHeight    =   4140
   ClientLeft      =   6270
   ClientTop       =   1500
   ClientWidth     =   6555
   ClipControls    =   0   'False
   Icon            =   "frmSupplMaster.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TXTTFLD 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   1560
      TabIndex        =   12
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox TXTTFLD 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox TXTTFLD 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox TXTTFLD 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox TXTTFLD 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
   Begin VB.ComboBox CBOTYPE 
      Appearance      =   0  'Flat
      DataSource      =   "datPrimaryRS"
      Height          =   315
      ItemData        =   "frmSupplMaster.frx":076A
      Left            =   1560
      List            =   "frmSupplMaster.frx":0771
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo DBCombo1 
      Bindings        =   "frmSupplMaster.frx":077B
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
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
            Picture         =   "frmSupplMaster.frx":0790
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupplMaster.frx":0B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupplMaster.frx":0EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupplMaster.frx":1B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupplMaster.frx":1FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSupplMaster.frx":27AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
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
   Begin VB.Label Label1 
      Caption         =   "E-mail"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Suppl Type"
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
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "cont_Person:"
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
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phone_Nos:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address:"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Name:"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Suppl. Code:"
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
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmSupplMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
Dim rsbrbill As New ADODB.Recordset

Private Sub CBOTYPE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBCombo1_Validate(Cancel As Boolean)
Dim rs As ADODB.Recordset
Set rs = MHVDB.Execute("select * from supplier where SuplCode = '" & DBCombo1.BoundText & "'", dbOpenDynaset)
If rs.EOF Then
   If Operation = "OPEN" Then
      MsgBox "This code does not exists !!! "
      Cancel = True
      Exit Sub
   End If
Else
   If Operation = "ADD" Then
      MsgBox "This code already exists !!! "
      Operation = "OPEN"
   End If
   TXTTFLD(0) = IIf(IsNull(rs!Name), "", rs!Name)
   TXTTFLD(1) = IIf(IsNull(rs!Address), "", rs!Address)
   TXTTFLD(2) = IIf(IsNull(rs!phone_nos), "", rs!phone_nos)
   TXTTFLD(3) = IIf(IsNull(rs!cont_person), "", rs!cont_person)
   TXTTFLD(4) = IIf(IsNull(rs!e_mail), "", rs!e_mail)
   'CHK.Value = jg!crdtok
   CBOTYPE = IIf(IsNull(rs!PARTYTYPE), "", rs!PARTYTYPE)
End If
'mnuSave.Enabled = True
TB.Buttons(3).Enabled = True
'mnuDelete.Enabled = True
TB.Buttons(4).Enabled = True
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
If rsbrbill.State = adStateOpen Then rsbrbill.Close
 rsbrbill.Open ("SELECT * FROM supplier"), db

Set DBCombo1.RowSource = rsbrbill
DBCombo1.ListField = "SuplCode"
DBCombo1.BoundColumn = "SuplCode"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuDelete_Click()
Dim jg As ADODB.Recordset
If Operation <> "OPEN" Then
   Beep
   Beep
   Exit Sub
End If
If MsgBox("Delete it !!! Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo ERR
'MHVDB.BeginTrans
Set jg = MHVDB.Execute("select SuplCode from tranhdr where SuplCode='" & DBCombo1.BoundText & "'", dbOpenDynaset, DBReadOnly)
If jg.EOF Then
   MHVDB.Execute "delete from Supplier where SuplCode='" & DBCombo1.BoundText & "'", dbSeeChanges + dbFailOnError
Else
   MsgBox "Cant be deleted !!!"
   Exit Sub
End If
'MHVDB.CommitTrans
'mnuSave.Enabled = False
TB.Buttons(3).Enabled = False
'mnuCancel.Enabled = False
TB.Buttons(4).Enabled = False
Operation = ""
DBCombo = ""
Exit Sub
ERR:
MsgBox ERR.Description
ERR.Clear
'MHVDB.Rollback
End Sub

Private Sub mnuNew_Click()
Dim i As Long
Operation = "ADD"
For i = 0 To 3
    TXTTFLD(i) = ""
Next
DBCombo = ""
'mnuSave.Enabled = True
TB.Buttons(3).Enabled = False
'mnuDelete.Enabled = False
TB.Buttons(4).Enabled = True
DBCombo1.SetFocus
End Sub

Private Sub mnuOpen_Click()
Operation = "OPEN"
DBCombo = ""
DBCombo1.SetFocus
End Sub

Private Sub mnuSave_Click()
If Operation = "" Then
MsgBox "Opration Not Selected, Try again."
Exit Sub
End If
Dim rs As New ADODB.Recordset
Set rs = Nothing
'RS.Open "select * from SubAcntMast where acntcode='6080' and acntsubcode='" & DBCombo1.BoundText & "'", AccStr

Dim SQLSTR As String
'On Error GoTo ERR
'MHVDB.BeginTrans
If Operation = "ADD" Then
   SQLSTR = "insert into Supplier (SuplCode,name,address,phone_nos,cont_person,partytype,e_mail) values " _
         & " ('" & DBCombo1.BoundText & "','" & TXTTFLD(0) & "','" & TXTTFLD(1) & "','" & TXTTFLD(2) & "','" & TXTTFLD(3) & "','" & CBOTYPE & "','" & TXTTFLD(4) & "')"
   MHVDB.Execute SQLSTR, dbFailOnError
   
ElseIf Operation = "OPEN" Then
   SQLSTR = "update Supplier set name='" & TXTTFLD(0) & "',address='" & TXTTFLD(1) & "',phone_nos='" & TXTTFLD(2) & "',cont_person='" & TXTTFLD(3) & "',partytype='" & CBOTYPE & "',e_mail='" & TXTTFLD(4) & "' " _
         & " where  SuplCode ='" & DBCombo1.BoundText & "'"
   MHVDB.Execute SQLSTR, dbFailOnError
   Set rs = Nothing

 
Else
   Beep
   Beep
End If
'MHVDB.CommitTrans

Operation = ""
DBCombo = ""
'mnuSave.Enabled = False
TB.Buttons(3).Enabled = False
'mnuDelete.Enabled = False
TB.Buttons(4).Enabled = False
Exit Sub
ERR:
MsgBox ERR.Description
ERR.Clear
'MHVDB.Rollback
Exit Sub
JR:
MsgBox "Error in posing accounts subcode " + ERR.Description
ERR.Clear
db.RollbackTrans
db.Close
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
       'mnuDelete_Click
       Case "EXIT"
       Unload Me
End Select
End Sub

Private Sub TXTTFLD_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub
