VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmmonitorassignment 
   Caption         =   "Monitor-Farmer Assignment"
   ClientHeight    =   8550
   ClientLeft      =   5490
   ClientTop       =   1605
   ClientWidth     =   12750
   Icon            =   "frmmonitorassignment.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12750
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   10560
      Picture         =   "frmmonitorassignment.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   3720
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Picture         =   "frmmonitorassignment.frx":11CC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit Farmer Detail"
         Top             =   1440
         Width           =   615
      End
      Begin VSFlex7Ctl.VSFlexGrid vgrid 
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3855
         _cx             =   6800
         _cy             =   2143
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
         FormatString    =   $"frmmonitorassignment.frx":1E96
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   45
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   9
      Top             =   720
      Width           =   4575
      Begin MSDataListLib.DataCombo cbodzongkhag 
         Bindings        =   "frmmonitorassignment.frx":1F2B
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   ""
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "frmmonitorassignment.frx":1F40
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   ""
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gewog"
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
         TabIndex        =   13
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dzongkhag"
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
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Farmer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   9015
      Begin VSFlex7Ctl.VSFlexGrid mygrid 
         Height          =   5415
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8655
         _cx             =   15266
         _cy             =   9551
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
         FormatString    =   $"frmmonitorassignment.frx":1F55
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
      Begin VB.Image imgBtnDn 
         Height          =   240
         Left            =   5040
         Picture         =   "frmmonitorassignment.frx":1FF9
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBtnUp 
         Height          =   240
         Left            =   5040
         Picture         =   "frmmonitorassignment.frx":2383
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tshowog"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   9720
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
      Begin VB.ListBox lstts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      Begin MSDataListLib.DataCombo cbomonitor 
         Bindings        =   "frmmonitorassignment.frx":270D
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   ""
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
      Begin MSDataListLib.DataCombo cbosupervisor 
         Bindings        =   "frmmonitorassignment.frx":2722
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   ""
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Active Monitor"
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
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Supervisor"
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
         Top             =   840
         Width           =   915
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   9000
      Top             =   720
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
            Picture         =   "frmmonitorassignment.frx":2737
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorassignment.frx":2AD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorassignment.frx":2E6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorassignment.frx":3B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorassignment.frx":3F97
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmonitorassignment.frx":4751
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   1164
      ButtonWidth     =   1191
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
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
            Caption         =   "View"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "EXIT"
            Object.ToolTipText     =   "EXIT FROM THE FORM"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Farmer under "
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
      TabIndex        =   16
      Top             =   8280
      Width           =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Farmer under "
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
      TabIndex        =   14
      Top             =   8040
      Width           =   1680
   End
End
Attribute VB_Name = "frmmonitorassignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim muk As Boolean
Private Sub cbodzongkhag_LostFocus()
On Error GoTo err
Dim rsGe As New ADODB.Recordset
cbodzongkhag.BackColor = vbWhite
If Len(cbodzongkhag.Text) = 0 Then
MsgBox "Please Select The Proper Dzongkhag First."
cbodzongkhag.SetFocus
Exit Sub
End If
cbodzongkhag.Enabled = False
If Operation = "ADD" Then
cbogewog.Enabled = True
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
cbogewog.Text = ""
If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cbodzongkhag.BoundText & "' order by gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"
'fillgewoglist

ElseIf Operation = "OPEN" Then

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub cbogewog_LostFocus()
filltshewog
End Sub

Private Sub CBOMONITOR_LostFocus()
If Len(cbomonitor.Text) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
cbomonitor.Enabled = False

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & cbomonitor.BoundText & "'", MHVDB
If rs.EOF <> True Then
If Len(rs!msupervisor) > 0 Then
FindsTAFF rs!msupervisor
End If
cbosupervisor.Text = rs!msupervisor & " " & sTAFF
End If
If Operation = "OPEN" Then
cbosupervisor.Enabled = False
cbodzongkhag.Enabled = False
cbogewog.Enabled = False
lstts.Enabled = False
TB.Buttons(3).Enabled = False
fillmain

Else
cbosupervisor.Enabled = True
cbodzongkhag.Enabled = True
cbogewog.Enabled = True
lstts.Enabled = True
mygrid.Enabled = True
TB.Buttons(3).Enabled = True
cbodzongkhag.Enabled = True
End If


End Sub

Private Sub Command1_Click()
If TB.Buttons(3).Enabled = False Then
MsgBox "Cannot release now!"
Exit Sub
End If

If MsgBox("Do you want to release all farmers from monitor " & cbomonitor.Text, vbYesNo) = vbYes Then
releaseall
End If

End Sub

Private Sub Command4_Click()
Frame3.Visible = False
End Sub

Private Sub Form_Load()
    fillmonitor
    fillms
    filldzongkhag
    muk = False
End Sub
Private Sub fillmonitor()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where moniter='1' order by STAFFCODE", db
Set cbomonitor.RowSource = Srs
cbomonitor.ListField = "STAFFNAME"
cbomonitor.BoundColumn = "STAFFCODE"


Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub fillms()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select distinct concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where MSUPERVISOR<>''  order by STAFFCODE", db
Set cbosupervisor.RowSource = Srs
cbosupervisor.ListField = "STAFFNAME"
cbosupervisor.BoundColumn = "STAFFCODE"


Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub filldzongkhag()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then rsDz.Close
Srs.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cbodzongkhag.RowSource = Srs
cbodzongkhag.ListField = "dzongkhagname"
cbodzongkhag.BoundColumn = "dzongkhagcode"
Exit Sub
err:
MsgBox err.Description
End Sub


Private Sub fillgewog()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then rsDz.Close
Srs.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cbodzongkhag.RowSource = Srs
cbodzongkhag.ListField = "dzongkhagname"
cbodzongkhag.BoundColumn = "dzongkhagcode"
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub lstge_Click()
filltshewog
End Sub

Private Sub lstts_Click()

fillfarmer

End Sub

Private Sub mygrid_Click()
mygrid.Editable = flexEDNone

If mygrid.col <> 1 And mygrid.row > 0 Then
Frame3.Visible = False
End If



End Sub
Private Sub fillvgrid()
Dim rs As New ADODB.Recordset
Dim i As Integer
Frame3.Visible = True
vgrid.Clear
vgrid.Rows = 1
vgrid.FormatString = "^Sl.No.|^Year|^Challan No.|^Plants|^"
vgrid.ColWidth(0) = 570
vgrid.ColWidth(1) = 960
vgrid.ColWidth(2) = 960
vgrid.ColWidth(3) = 960
vgrid.ColWidth(4) = 240

Set rs = Nothing
rs.Open "select farmercode,year,challanno,challanqty from tblplanted where farmercode='" & Mid(Trim(mygrid.TextMatrix(mygrid.row, 1)), 1, 14) & "'  order by year", MHVDB

i = 1
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
Label7.Caption = rs!farmercode & "  " & FAName
vgrid.Rows = vgrid.Rows + 1
vgrid.TextMatrix(i, 0) = i
vgrid.TextMatrix(i, 1) = rs!Year
vgrid.TextMatrix(i, 2) = rs!challanno
vgrid.TextMatrix(i, 3) = rs!challanqty

rs.MoveNext
i = i + 1
Loop

End Sub

Private Sub mygrid_DblClick()
If mygrid.col = 1 And mygrid.row > 0 Then

fillvgrid
Else
Frame3.Visible = False
End If
End Sub

Private Sub mygrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' only interesetd in left button
If Operation = "OPEN" Then Exit Sub
   On Error Resume Next
    If Button <> 1 Then Exit Sub
    
    ' get cell that was clicked
    Dim r&, c&
    r = mygrid.MouseRow
    c = mygrid.MouseCol
    
    ' make sure the click was on the sheet
    If r < 0 Or c < 0 Then Exit Sub
    
    ' make sure the click was on a cell with a button
    If mygrid.Cell(flexcpPicture, r, c) <> imgBtnUp Then Exit Sub
    
    ' make sure the click was on the button (not just on the cell)
    ' note: this works for right-aligned buttons
    Dim d!
    d = mygrid.Cell(flexcpLeft, r, c) + mygrid.Cell(flexcpWidth, r, c) - X
    If d > imgBtnDn.Width Then Exit Sub
    
    ' click was on a button: do the work
    mygrid.Cell(flexcpPicture, r, c) = imgBtnDn
    'MsgBox "Thanks for clicking my custom button!"
    'MsgBox "You have clicked on row=" & r & "  col=" & c
   
    If MsgBox("Make sure that you want to delink farmer " & mygrid.TextMatrix(r, 1) & " from monitor " & cbomonitor.Text & ".", vbYesNo) = vbYes Then
    removefarmer r, c
    End If
    
    mygrid.Cell(flexcpPicture, r, c) = imgBtnUp
    
    ' cancel default processing
    ' note: this is not strictly necessary in this case, because
    '       the dialog box already stole the focus etc, but let's be safe.
    Cancel = True
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       Operation = "ADD"
       CLEARCONTROLL
       cbomonitor.Enabled = True
       cbosupervisor.Enabled = True
       TB.Buttons(3).Enabled = False
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbomonitor.Enabled = True
       cbosupervisor.Enabled = True
       TB.Buttons(3).Enabled = False
       Case "SAVE"
       MNU_SAVE
        
        
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub MNU_SAVE()
Dim i As Integer
Dim rs As New ADODB.Recordset
If Len(cbosupervisor.Text) = 0 Then
MsgBox "Not a valid supervosor!"
Exit Sub
End If

MHVDB.BeginTrans
Set rs = Nothing
rs.Open "Select * from tblmhvstaff where staffcode='" & cbosupervisor.BoundText & "'", MHVDB
If rs.EOF <> True Then

Else
MsgBox "Not a valid supervosor!"
MHVDB.RollbackTrans
Exit Sub
End If

' update m supervisor
MHVDB.Execute "update tblmhvstaff set msupervisor='" & Mid(Trim(cbosupervisor.Text), 1, 5) & "' where staffcode='" & cbomonitor.BoundText & "'"
' update farmers with
For i = 1 To mygrid.Rows - 1
If Len(Trim(mygrid.TextMatrix(i, 1))) = 0 Then Exit For
MHVDB.Execute "update tblfarmer set monitor='" & cbomonitor.BoundText & "' where idfarmer='" & Mid(Trim(mygrid.TextMatrix(i, 1)), 1, 14) & "'"
Next
MHVDB.CommitTrans

TB.Buttons(3).Enabled = False
End Sub
Private Sub CLEARCONTROLL()
cbomonitor.Text = ""
cbosupervisor.Text = ""
cbodzongkhag.Text = ""
cbogewog.Text = ""
lstts.Clear

mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^Sl.No.|^Farmer(Registration)|^Farmer(Planted)|^|^"
mygrid.ColWidth(0) = 960
mygrid.ColWidth(1) = 3495
mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 330
mygrid.ColWidth(4) = 270


End Sub
Private Sub filltshewog()
Dim Dzstr As String
'Dzstr = ""
''SQLSTR = ""
'
'
'For i = 0 To lstge.ListCount - 1
'    If lstge.Selected(i) Then
'       Dzstr = Dzstr + "'" + Trim(Mid(lstge.List(i), InStr(1, lstge.List(i), "|") + 1)) + "',"
'       Mcat = lstge.List(i)
'       j = j + 1
'    End If
'    If RepName = "5" Then
'       If j > 1 Then
'          MsgBox "Gewog not selected !!!"
'            lstge.Clear
'          Exit Sub
'       End If
'    End If
'Next
'If Len(Dzstr) > 0 Then
'   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
'
'Else
'   MsgBox "Gewog not selected !!!"
'           lstts.Clear
'   Exit Sub
'End If




Dim rs As New ADODB.Recordset

Set rs = Nothing

rs.Open "select tshewogid,tshewogname from tbltshewog where dzongkhagid='" & cbodzongkhag.BoundText & "' and gewogid ='" & cbogewog.BoundText & "' Order by tshewogname", MHVDB, adOpenStatic
With rs
lstts.Clear
Do While Not .EOF
   lstts.AddItem Trim(!tshewogname) + " | " + !tshewogid
   .MoveNext
Loop
End With
End Sub

Private Sub fillfarmer()
Dim Dzstr As String
Dim Gestr As String
Dim SQLSTR As String
Dim misscnt As Integer
Dim rs As New ADODB.Recordset
Dim rsp As New ADODB.Recordset
Dzstr = ""
Gestr = ""
SQLSTR = ""


'For i = 0 To lstge.ListCount - 1
'    If lstge.Selected(i) Then
'       Gestr = Gestr + "'" + Trim(Mid(lstge.List(i), InStr(1, lstge.List(i), "|") + 1)) + "',"
'       Mcat = lstge.List(i)
'       j = j + 1
'    End If
'Next
'If Len(Gestr) > 0 Then
'   Gestr = "(" + Left(Gestr, Len(Gestr) - 1) + ")"
'
'Else
'   MsgBox "Gewog not selected !!!"
'
'   Exit Sub
'End If













    For i = 0 To lstts.ListCount - 1
    If lstts.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(lstts.List(i), InStr(1, lstts.List(i), "|") + 1)) + "',"
       Set rs = Nothing
  
       
       rs.Open "select * from tblfarmer where status='A' and substring(idfarmer,1,9)='" & cbodzongkhag.BoundText & cbogewog.BoundText & Trim(Mid(lstts.List(i), InStr(1, lstts.List(i), "|") + 1)) & "' and length(monitor)='5'", MHVDB
       
       
       If rs.EOF <> True Then
       If rs!monitor <> cbomonitor.BoundText Then
       FindsTAFF rs!monitor
       MsgBox "This Tshowog already assigned to monitor " & rs!monitor & "  " & sTAFF
       lstts.Selected(i) = False
       
       Exit Sub
       End If
       End If
       
       
       
       
       
       
       Mcat = lstts.List(i)
       j = j + 1
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
  ' MsgBox "Tshewog not selected !!!"
            mygrid.Clear
            mygrid.Rows = 1
           mygrid.FormatString = "^Sl.No.|^Farmer(Registration)|^Farmer(Planted)|^|^"
mygrid.ColWidth(0) = 960
mygrid.ColWidth(1) = 3495



mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 330
mygrid.ColWidth(4) = 270
            Label5.Caption = ""
            Label6.Caption = ""
   Exit Sub
End If

SQLSTR = "select * from tblfarmer where status='A' and substring(idfarmer,1,3)='" & cbodzongkhag.BoundText & "' and substring(idfarmer,4,3) ='" & cbogewog.BoundText & "'  and substring(idfarmer,7,3) in " & Dzstr



mygrid.Clear
misscnt = 0
mygrid.FormatString = "^Sl.No.|^Farmer(Registration)|^Farmer(Planted)|^|^"
mygrid.ColWidth(0) = 960
mygrid.ColWidth(1) = 3495
mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 330
mygrid.ColWidth(4) = 270

Set rs = Nothing
rs.Open SQLSTR, MHVDB
i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = rs!idfarmer & "  " & rs!farmername
Set rsp = Nothing
rsp.Open "select * from tblplanted where status<>'C' and farmercode='" & rs!idfarmer & "' group by farmercode", MHVDB
If rsp.EOF <> True Then
mygrid.TextMatrix(i, 2) = "" 'rs!idfarmer & "  " & rs!farmername
Else
mygrid.TextMatrix(i, 2) = "FARMER MISSING"
misscnt = misscnt + 1
End If



mygrid.ColAlignment(2) = flexAlignLeftTop
mygrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop
Label5.Caption = "Total farmer under " & cbomonitor.Text & "is " & i - 1
Label6.Caption = "Total missing farmer in planted list under " & cbomonitor.Text & "is " & misscnt
rs.Close
addbtn

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub addbtn()
' initialize grid
    mygrid.Editable = flexEDKbdMouse
    mygrid.AllowUserResizing = flexResizeNone

   
    Dim i%
    For i = 1 To mygrid.Rows - 1
        mygrid.Cell(flexcpPicture, i, 3) = imgBtnUp
        mygrid.Cell(flexcpPictureAlignment, i, 3) = flexAlignRightCenter
    Next
End Sub

Private Sub removefarmer(mrow As Long, MCOL As Long)
    Dim rs As New ADODB.Recordset
    MHVDB.Execute "update tblfarmer set monitor='' where idfarmer='" & Mid(Trim(mygrid.TextMatrix(mrow, 1)), 1, 14) & "'"
    mygrid.RemoveItem (mrow)
    mygrid.ComboList = ""
    mygrid.Editable = flexEDNone
    
    fillfarmer1
    End Sub
    Private Sub releaseall()
     Dim rs As New ADODB.Recordset
     Dim i As Integer
     
     For i = 1 To mygrid.Rows - 1
     If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
    MHVDB.Execute "update tblfarmer set monitor='' where idfarmer='" & Mid(Trim(mygrid.TextMatrix(i, 1)), 1, 14) & "'"
    mygrid.Editable = flexEDNone
    Next
    mygrid.Rows = 1
    End Sub
Private Sub fillmain()
On Error GoTo err
Dim misscnt As Integer
Dim rsp As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid.Clear
mygrid.Rows = 1
misscnt = 0
mygrid.FormatString = "^Sl.No.|^Farmer(Registration)|^Farmer(Planted)|^|^"
mygrid.ColWidth(0) = 960
mygrid.ColWidth(1) = 3495
mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 330
mygrid.ColWidth(4) = 270

Set rs = Nothing
rs.Open "select * from tblfarmer where status='A' and monitor='" & cbomonitor.BoundText & "' order by idfarmer", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = rs!idfarmer & "  " & rs!farmername

Set rsp = Nothing
rsp.Open "select * from tblplanted where status<>'C' and farmercode='" & rs!idfarmer & "' group by farmercode", MHVDB
If rsp.EOF <> True Then
mygrid.TextMatrix(i, 2) = "" 'rs!idfarmer & "  " & rs!farmername
Else
mygrid.TextMatrix(i, 2) = "FARMER MISSING"
misscnt = misscnt + 1
End If



mygrid.ColAlignment(2) = flexAlignLeftTop
mygrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop
Label5.Caption = "Total farmer under " & cbomonitor.Text & "is " & i - 1
Label6.Caption = "Total missing farmer in planted list under " & cbomonitor.Text & "is " & misscnt
rs.Close
addbtn

Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub fillgewoglist()
Dim rs As New ADODB.Recordset

Set rs = Nothing

rs.Open "select GewogId,Gewogname from tblgewog where DzongkhagId='" & cbodzongkhag.BoundText & "'  Order by Gewogname", MHVDB, adOpenStatic
With rs
lstge.Clear
Do While Not .EOF
   lstge.AddItem Trim(!gewogname) + " | " + !gewogid
   .MoveNext
Loop
End With
End Sub




Private Sub fillfarmer1()
  
For i = 1 To mygrid.Rows - 1
mygrid.TextMatrix(i, 0) = i
Next
Label5.Caption = "Total farmer under " & cbomonitor.Text & "is " & i - 1

Exit Sub
err:
MsgBox err.Description
End Sub
