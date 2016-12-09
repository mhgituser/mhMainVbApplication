VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEPARTMENT"
   ClientHeight    =   5220
   ClientLeft      =   8190
   ClientTop       =   1170
   ClientWidth     =   6750
   ClipControls    =   0   'False
   Icon            =   "frmDep.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TXTTFLD 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtremarks 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1680
      Width           =   3255
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
      Appearance      =   0
      Text            =   ""
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   5280
      Top             =   600
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
            Picture         =   "frmDep.frx":076A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDep.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDep.frx":0E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDep.frx":1B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDep.frx":1FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDep.frx":2784
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
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
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   6615
      _cx             =   11668
      _cy             =   4895
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   8438015
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDep.frx":2B1E
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblLabels 
      Caption         =   "Remarks"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
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
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Dept. ID"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
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
jg.Open "select * from tbldepartment where  deptid = '" & DBCombo1.BoundText & "'", db
If jg.EOF Then
   If operation = "OPEN" Then
      MsgBox "This code does not exists !!! "
      Cancel = True
      Exit Sub
   End If
Else
   If operation = "ADD" Then
      MsgBox "This code already exists !!! "
      op = "Open"
   End If
   TXTTFLD.Text = IIf(IsNull(jg!DeptName), "", jg!DeptName)
  txtremarks.Text = IIf(IsNull(jg!remarks), "", jg!remarks)
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
DATA1.Open "SELECT * FROM tbldepartment", db, adOpenDynamic, adLockReadOnly
Set DBCombo1.RowSource = DATA1
DBCombo1.ListField = "deptname"
DBCombo1.BoundColumn = "deptid"
FillGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuDelete_Click()
Dim jg As New ADODB.Recordset
If op <> "Open" Then
   Beep
   Beep
   Exit Sub
End If
If JU < 1 Then
   MsgBox "No rights! Contact System Administrator !"
   Exit Sub
End If
Set Item = New ADODB.Recordset
Item.Open "select * from tranhdr where suplcode ='" & DBCombo1.BoundText & "' and billtype='II'", db
If Not Item.EOF Then
   MsgBox "Material Issued to this Dept. ! Cant delete !"
   Exit Sub
End If
If MsgBox("Delete it !!! Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo err
db.BeginTrans
db.Execute "delete  from departments where deptcode='" & DBCombo1.BoundText & "'"
db.CommitTrans

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = False
op = ""
DBCombo1 = ""
DATA1.Requery
DBCombo1.Refresh
Exit Sub
err:
MsgBox err.Description
err.Clear
db.RollbackTrans
End Sub

Private Sub mnuNew_Click()
Dim i As Long
operation = "ADD"
TXTTFLD.Text = ""
txtremarks.Text = ""

    Set lastbill = New ADODB.Recordset
    lastbill.Open "select max(deptid) as lno from tbldepartment", db, adOpenDynamic
    DBCombo1 = IIf(IsNull(lastbill!lno), 100, lastbill!lno + 1)
    Set lastbill = Nothing
    DBCombo1.Enabled = False


TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = False
End Sub

Private Sub mnuOpen_Click()
operation = "OPEN"
DBCombo1 = ""
DBCombo1.Enabled = True
DBCombo1.SetFocus
End Sub

Private Sub mnuSave_Click()
Dim SQLSTR As String
If Not (operation = "OPEN" Or operation = "ADD") Then
   Beep
   MsgBox "No Operation Selected !!!!"
   Exit Sub
End If
On Error GoTo err
db.BeginTrans
If operation = "ADD" Then
   SQLSTR = "insert into tbldepartment (deptid,deptname,remarks) values " _
         & " ('" & DBCombo1.BoundText & "','" & TXTTFLD.Text & "','" & txtremarks.Text & "')"
   db.Execute SQLSTR
  ElseIf operation = "OPEN" Then
   SQLSTR = "update tbldepartment set deptname='" & TXTTFLD.Text & "',remarks='" & txtremarks.Text & "' where  deptid ='" & DBCombo1.BoundText & "'"
   db.Execute SQLSTR
Else
   Beep
   Beep
   MsgBox "NO OPERATION SELECTRD !!! CANT SAVE."
   Exit Sub
End If
db.CommitTrans
FillGrid
operation = ""
DBCombo1 = ""
DATA1.Requery
DBCombo1.Refresh

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = False
Exit Sub
err:
MsgBox err.Description
err.Clear
db.RollbackTrans
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub TXTTFLD_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = TXTTFLD.SelStart + 1
    Dim sText As String
    sText = Left$(TXTTFLD.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub TXTTFLD_Validate(Cancel As Boolean)
If Len(Trim(TXTTFLD.Text)) = 0 Then
   MsgBox "Can't be blank !!!"
   Cancel = True
End If
End Sub

Private Sub FillGrid()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Dept. ID|^Dept. Name|^remarks|^"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 825
Mygrid.ColWidth(2) = 2370
Mygrid.ColWidth(3) = 1650
Mygrid.ColWidth(4) = 660

rs.Open "select * from tbldepartment order by deptid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = rs!deptid
Mygrid.TextMatrix(i, 2) = rs!DeptName
Mygrid.TextMatrix(i, 3) = rs!remarks

rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
