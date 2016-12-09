VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmextensionupdate 
   Caption         =   "Territory"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Unselect All"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select All"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   8040
      Width           =   855
   End
   Begin VB.Frame Frame4 
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
      Height          =   5895
      Left            =   6000
      TabIndex        =   10
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
         TabIndex        =   11
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4815
      Begin MSDataListLib.DataCombo cbomonitor 
         Bindings        =   "frmextensionupdate.frx":0000
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSDataListLib.DataCombo cboregion 
         Bindings        =   "frmextensionupdate.frx":0015
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Region"
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
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Extension Officer"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1470
      End
   End
   Begin VB.Frame Frame5 
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
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   5775
      Begin VSFlex7Ctl.VSFlexGrid mygrid 
         Height          =   5415
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5295
         _cx             =   9340
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
         FormatString    =   $"frmextensionupdate.frx":002A
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
      Begin VB.Image imgBtnUp 
         Height          =   240
         Left            =   5040
         Picture         =   "frmextensionupdate.frx":00C2
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBtnDn 
         Height          =   240
         Left            =   5040
         Picture         =   "frmextensionupdate.frx":044C
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Territory"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   4575
      Begin MSDataListLib.DataCombo cbodzongkhag 
         Bindings        =   "frmextensionupdate.frx":07D6
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   10560
      Picture         =   "frmextensionupdate.frx":07EB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1095
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
            Picture         =   "frmextensionupdate.frx":0B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmextensionupdate.frx":0F0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmextensionupdate.frx":12A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmextensionupdate.frx":1F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmextensionupdate.frx":23D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmextensionupdate.frx":2B8F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
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
End
Attribute VB_Name = "frmextensionupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fillgewog()
Dim Dzstr As String


Dim rs As New ADODB.Recordset

Set rs = Nothing

rs.Open "select gewogid,gewogname from tblgewog where dzongkhagid='" & cbodzongkhag.BoundText & "'  Order by gewogid", MHVDB, adOpenStatic
With rs
lstts.Clear
Do While Not .EOF
   lstts.AddItem Trim(!gewogname) + " | " + !gewogid
   .MoveNext
Loop
End With
End Sub

Private Sub cbodzongkhag_LostFocus()
fillgewog
End Sub

Private Sub Form_Load()
fillext
fillextreg
filldzongkhag
End Sub
Private Sub fillext()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where dept='105' and status='ON' order by STAFFCODE", db
Set cbomonitor.RowSource = Srs
cbomonitor.ListField = "STAFFNAME"
cbomonitor.BoundColumn = "STAFFCODE"


Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
cbomonitor.Text = ""

cbodzongkhag.Text = ""
cboregion.Text = ""
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

Private Sub fillextreg()
On Error GoTo err
Operation = ""
Dim Srs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(regioncode , ' ', regionname) as regionname,regioncode  from tblextregion order by regioncode", db
Set cboregion.RowSource = Srs
cboregion.ListField = "regionname"
cboregion.BoundColumn = "regioncode"


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
Private Sub filltshowog()
Dim Dzstr As String
Dim Gestr As String
Dim SQLSTR As String
Dim misscnt As Integer
Dim rs As New ADODB.Recordset
Dim rsp As New ADODB.Recordset
Dzstr = ""
Gestr = ""
SQLSTR = ""


 mygrid.Clear
            mygrid.Rows = 1


    For i = 0 To lstts.ListCount - 1
    If lstts.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(lstts.List(i), InStr(1, lstts.List(i), "|") + 1)) + "',"
       Set rs = Nothing
  
       
       rs.Open "select * from tbltshewog where  dzongkhagid='" & cbodzongkhag.BoundText & "' and gewogid= '" & Trim(Mid(lstts.List(i), InStr(1, lstts.List(i), "|") + 1)) & "' and length(regioncode)>'0'", MHVDB
       
       
'       If rs.EOF <> True Then
'       If rs!monitor <> cbomonitor.BoundText Then
'       FindsTAFF rs!monitor
'       MsgBox "This Tshowog is already assigned to  " & rs!monitor & "  " & sTAFF
'       lstts.Selected(i) = False
'
'       Exit Sub
'       End If
'       End If
       
       
       
       
       
       
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
           
   Exit Sub
End If

SQLSTR = "select * from tbltshewog where  dzongkhagid='" & cbodzongkhag.BoundText & "' and gewogid in " & Dzstr



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
mygrid.TextMatrix(i, 1) = rs!dzongkhagid & rs!gewogid & rs!tshewogid & "  " & rs!tshewogname

mygrid.ColAlignment(2) = flexAlignLeftTop
mygrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop
rs.Close

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub lstts_Click()
filltshowog
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       Operation = "ADD"
       CLEARCONTROLL
       cbomonitor.Enabled = True
       cboregion.Enabled = True
       TB.Buttons(3).Enabled = True
       
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
If Len(cbomonitor.Text) = 0 Then
MsgBox "Select Valid extension officer"
Exit Sub
End If

If Len(cboregion.Text) = 0 Then
MsgBox "Select Valid region "
Exit Sub
End If

MHVDB.BeginTrans

Set rs = Nothing
rs.Open "Select * from tblmhvstaff where staffcode='" & cbomonitor.BoundText & "' and dept='105' and status='ON'", MHVDB
If rs.EOF <> True Then
Else
MsgBox "Select valid extension officer!"
MHVDB.RollbackTrans
Exit Sub
End If




' update m supervisor
MHVDB.Execute "update tblextregion set extensionofficercode='" & Mid(Trim(cbomonitor.Text), 1, 5) & "' where regioncode='" & cboregion.BoundText & "'"
' update farmers with
For i = 1 To mygrid.Rows - 1
If Len(Trim(mygrid.TextMatrix(i, 1))) = 0 Then Exit For
MHVDB.Execute "update tbltshewog set regioncode='" & cboregion.BoundText & "' where concat(dzongkhagid,gewogid,tshewogid)='" & Mid(Trim(mygrid.TextMatrix(i, 1)), 1, 9) & "'"
Next
MHVDB.CommitTrans

TB.Buttons(3).Enabled = False
End Sub
