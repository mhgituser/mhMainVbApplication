VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMROLE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ROLE MASTER"
   ClientHeight    =   5835
   ClientLeft      =   5835
   ClientTop       =   1185
   ClientWidth     =   10170
   Icon            =   "FRMROLE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   Begin VB.Frame itemcode 
      Caption         =   "ROLE INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   9975
      Begin VB.CheckBox CHKTSHOWOG 
         Caption         =   "TSHOWOG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox CHKGEWOG 
         Caption         =   "GEWOG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TXTDESC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   840
         Width           =   6975
      End
      Begin MSDataListLib.DataCombo CBOROLEID 
         Bindings        =   "FRMROLE.frx":0E42
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "OPTION TO GRAY OUT IN CONTACT DETAILS"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   4080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ROLE ID"
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
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
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
         Width           =   1275
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   9975
      _cx             =   17595
      _cy             =   4895
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12632256
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FRMROLE.frx":0E57
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
   Begin MSComctlLib.ImageList IMG 
      Left            =   2640
      Top             =   2400
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
            Picture         =   "FRMROLE.frx":0EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMROLE.frx":125E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMROLE.frx":15F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMROLE.frx":22D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMROLE.frx":2724
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMROLE.frx":2EDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
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
   Begin VB.Label Label4 
      Caption         =   "DATA SCREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "FRMROLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsROLE As New ADODB.Recordset
Dim mgraytype As String
'Mygrid.FormatString = "^SL.NO.|^DZONGKHAG ID|^DZONGKHAG NAME|^REMARKS"


Private Sub CBOROLEID_GotFocus()
   CBOROLEID.BackColor = vbYellow
End Sub

Private Sub CBOROLEID_LostFocus()

   On Error GoTo ERR
   CBOROLEID.BackColor = vbWhite
   CBOROLEID.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblrole where roleid='" & CBOROLEID.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   txtDesc.Text = rs!roledescription
   CHKGEWOG.Value = rs!GRAYGEWOG
   CHKTSHOWOG.Value = rs!GRAYTSHOWOG
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   
   Dim rschkrole As New ADODB.Recordset
Set rschkrole = Nothing
rschkrole.Open "select * from tblcontact where roleid='" & CBOROLEID.BoundText & "'", MHVDB
If rschkrole.EOF <> True Then
CHKGEWOG.Enabled = False
CHKTSHOWOG.Enabled = False
Else
CHKGEWOG.Enabled = True
CHKTSHOWOG.Enabled = True

End If
   
   
   
   
   Exit Sub
ERR:
   MsgBox ERR.Description
   'rs.Close
End Sub


Private Sub Form_Load()
On Error GoTo ERR
Operation = ""

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsROLE = Nothing

If rsROLE.State = adStateOpen Then rsROLE.Close
rsROLE.Open "select concat(roleid , ' ', roledescription) as roledesc,roleid  from tblrole order by roleid", db
Set CBOROLEID.RowSource = rsROLE
CBOROLEID.ListField = "roledesc"
CBOROLEID.BoundColumn = "roleid"
FillGrid
Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub Tb_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERR
Select Case Button.Key

       Case "ADD"
       CBOROLEID.Enabled = False
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       Dim rs As New ADODB.Recordset
       Set rs = Nothing
       rs.Open "SELECT MAX(SUBSTRING(ROLEID,2,2))+1 AS MaxID from tblrole", MHVDB, adOpenForwardOnly, adLockOptimistic
       If rs.EOF <> True Then
       
       
       
       If Len(rs!MaxID) = 1 Then
       CBOROLEID.Text = "R0" & IIf(IsNull(rs!MaxID), "01", rs!MaxID)
        Else
        CBOROLEID.Text = "R" & IIf(IsNull(rs!MaxID), "01", rs!MaxID)
        End If
       
        
        
       Else
       CBOROLEID.Text = "R01" & rs!MaxID
       End If
       
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       CBOROLEID.Enabled = True
      TB.Buttons(3).Enabled = True
       
       Case "SAVE"
       MNU_SAVE
        TB.Buttons(3).Enabled = False
        FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
ERR:
MsgBox ERR.Description
End Sub
Private Sub MNU_SAVE()
On Error GoTo ERR
Dim rs As New ADODB.Recordset

mgraytype = CHKGEWOG.Value & CHKTSHOWOG.Value
If mgraytype = "10" Then
MsgBox "ONLY GEWOG CANNOT BE GRAY TYPE.!"
Exit Sub
End If


MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute "INSERT INTO tblrole (roleid,roledescription,GRAYGEWOG,GRAYTSHOWOG) VALUEs('" & CBOROLEID.Text & "','" & txtDesc.Text & "','" & CHKGEWOG.Value & "','" & CHKTSHOWOG.Value & "')"

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblrole set roledescription='" & txtDesc.Text & "',GRAYGEWOG='" & CHKGEWOG.Value & "',GRAYTSHOWOG='" & CHKTSHOWOG.Value & "' where roleid='" & CBOROLEID.BoundText & "'"
Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans

Exit Sub

ERR:
MsgBox ERR.Description
MHVDB.RollbackTrans


End Sub
Private Sub CLEARCONTROLL()

txtDesc.Text = ""
CHKGEWOG.Value = 0
CHKTSHOWOG.Value = 0

End Sub
Private Sub FillGrid()
On Error GoTo ERR
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^SL.NO.|^ROLE ID|^DESCRIPTION"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 2835
Mygrid.ColWidth(2) = 6300


rs.Open "select * from tblrole order by ROLEID", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = rs!roleid
Mygrid.TextMatrix(i, 2) = rs!roledescription

rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
ERR:
MsgBox ERR.Description

End Sub
Private Sub TXTDESC_GotFocus()
txtDesc.BackColor = vbYellow
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtDesc.SelStart + 1
    Dim sText As String
    sText = Left$(txtDesc.Text, iPos)
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

Private Sub TXTDESC_LostFocus()
txtDesc.BackColor = vbWhite
End Sub






