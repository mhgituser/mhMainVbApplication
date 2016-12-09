VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMMODULEMAINTAINENCE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULE MENTAINANCE"
   ClientHeight    =   7950
   ClientLeft      =   6000
   ClientTop       =   1425
   ClientWidth     =   10005
   Icon            =   "FRMMODULEMAINTAINENCE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10005
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   6855
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   5655
      _cx             =   9975
      _cy             =   12091
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      FormatString    =   $"FRMMODULEMAINTAINENCE.frx":0E42
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
   Begin VB.Frame Frame2 
      Caption         =   "USER SELECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   3495
      Begin MSDataListLib.DataCombo cbouser 
         Bindings        =   "FRMMODULEMAINTAINENCE.frx":0ED9
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   360
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
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOAD"
      Height          =   615
      Left            =   3240
      Picture         =   "FRMMODULEMAINTAINENCE.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "MODULE CATEGORY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
      Begin VB.OptionButton optrpts 
         Caption         =   "REPORTS"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton opttrns 
         Caption         =   "TRANSECTION"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optmaster 
         Caption         =   "MASTER"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   600
      Top             =   1200
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
            Picture         =   "FRMMODULEMAINTAINENCE.frx":17B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMODULEMAINTAINENCE.frx":1B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMODULEMAINTAINENCE.frx":1EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMODULEMAINTAINENCE.frx":2BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMODULEMAINTAINENCE.frx":3018
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMMODULEMAINTAINENCE.frx":37D2
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
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin MSDataListLib.DataCombo cbostaffcode 
      Bindings        =   "FRMMODULEMAINTAINENCE.frx":3B6C
      DataField       =   "ItemCode"
      Height          =   420
      Left            =   960
      TabIndex        =   9
      Top             =   4560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   795
   End
End
Attribute VB_Name = "FRMMODULEMAINTAINENCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUs As New ADODB.Recordset

Private Sub cbouser_LostFocus()
CBOUSER.Enabled = False
End Sub

Private Sub Command1_Click()
FillGrid
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
MODULETYPE = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsUs = Nothing

If rsUs.State = adStateOpen Then rsUs.Close
rsUs.Open "select concat(cast(userid as char) , ' ', username) as username,userid  from tblsoftuser order by userid", db
Set CBOUSER.RowSource = rsUs
CBOUSER.ListField = "username"
CBOUSER.BoundColumn = "userid"


Dim rs As New ADODB.Recordset


Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select * FROM tblmodulemaster WHERE STATUS='ON'", db
Set cbostaffcode.RowSource = rs
cbostaffcode.ListField = "MODULENAME"
cbostaffcode.BoundColumn = "MODuLEID"


End Sub
Private Sub FillGrid()
On Error GoTo err

Dim rs As New ADODB.Recordset
Dim i As Integer
'If userid = "" Then Exit Sub
Set rs = Nothing
rs.Open "select * from  tblsoftuser where userid='" & CBOUSER.BoundText & "'", MHVDB
If rs.EOF <> True Then

Else
MsgBox "Not a valid user."
Exit Sub
End If
If Len(cbostaffcode.Text) = 0 Then
MsgBox "Please Select The Module"

Exit Sub
End If
Set rs = Nothing
mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^SL.NO.|^MODULE ID|^MODULENAME NAME|^ACCESS"
mygrid.ColWidth(0) = 600
mygrid.ColWidth(1) = 0
mygrid.ColWidth(2) = 2835
mygrid.ColWidth(3) = 960
mygrid.ColWidth(4) = 960

rs.Open "select moduleid,modulename,userrights from tblmodule WHERE MODULETYPE='" & MODULETYPE & "' AND MAINMODULE='" & cbostaffcode.BoundText & "' and userid='" & CBOUSER.BoundText & "'  union select moduleid,modulename,'0' from tblmodule where  moduletype='" & MODULETYPE & "' AND MAINMODULE='" & cbostaffcode.BoundText & "' AND moduleid not in (select moduleid from tblmodule WHERE  moduletype='" & MODULETYPE & "' AND MAINMODULE='" & cbostaffcode.BoundText & "' AND  userid='" & CBOUSER.BoundText & "')", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
mygrid.TextMatrix(i, 0) = i

mygrid.TextMatrix(i, 1) = rs!moduleid
mygrid.TextMatrix(i, 2) = rs!modulename
If Operation = "OPEN" Then
mygrid.TextMatrix(i, 3) = rs!userrights
Else
mygrid.TextMatrix(i, 3) = 0
End If
rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub mygrid_Click()
If mygrid.col = 3 Then
mygrid.Editable = flexEDKbdMouse
Else

mygrid.Editable = flexEDNone
End If
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()
MODULETYPE = "T"
End Sub

Private Sub Option3_Click()
MODULETYPE = "R"
End Sub

Private Sub optmaster_Click()
MODULETYPE = 0
End Sub

Private Sub optrpts_Click()
MODULETYPE = 2
End Sub

Private Sub opttrns_Click()
MODULETYPE = 1
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       CBOUSER.Enabled = True
        TB.buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
            
       
       Case "OPEN"
       Operation = "OPEN"
       'CLEARCONTROLL
       CBOUSER.Enabled = True
       TB.buttons(3).Enabled = True
       
       Case "SAVE"
       MNU_SAVE
        TB.buttons(3).Enabled = False
        'FillGrid
       
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
End Sub
Private Sub MNU_SAVE()

On Error GoTo err

Dim VKEY As String

If MODULETYPE = "" Then
MsgBox "Select The Proper Modle Type."
Exit Sub
End If
If optmaster.Value = True Then
MODULETYPE = 0
ElseIf opttrns.Value = True Then
MODULETYPE = 1
ElseIf optrpts.Value = True Then
MODULETYPE = 2

Else
MsgBox "Invalid Module Type Selection."
Exit Sub
End If


If Len(cbostaffcode.Text) = 0 Then
MsgBox "Please Select The Module."
Exit Sub
End If

Dim rr As String
Dim i As Integer
i = 1
Dim yy As String
yy = "delete from tblmodule where userid='" & CBOUSER.BoundText & "' and moduletype='" & MODULETYPE & "' AND MAINMODULE='" & cbostaffcode.BoundText & "'"
MHVDB.BeginTrans
MHVDB.Execute "delete from tblmodule where userid='" & CBOUSER.BoundText & "' and moduletype='" & MODULETYPE & "' AND MAINMODULE='" & cbostaffcode.BoundText & "'"

For i = 1 To mygrid.rows - 1
If Len(Trim(mygrid.TextMatrix(i, 1))) > 0 Then
        
            If mygrid.ValueMatrix(i, 3) * -1 = 1 Then
                           
               MHVDB.Execute "insert into tblmodule(moduleid,modulename,userid,userrights,moduletype,iconid,vkey,mainmodule) (select moduleid,modulename,'" & CBOUSER.BoundText & "',userrights,moduletype,iconid,vkey,mainmodule from tblmodule where moduleid='" & mygrid.TextMatrix(i, 1) & "' and userid='100' and MODULETYPE='" & MODULETYPE & "' )"
                             
            End If
       
   Else
   
   Exit For
      End If
  Next

MHVDB.CommitTrans
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub
