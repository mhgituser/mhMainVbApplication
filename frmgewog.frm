VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmgewog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GEWOG MASTER"
   ClientHeight    =   9240
   ClientLeft      =   4800
   ClientTop       =   1680
   ClientWidth     =   9735
   Icon            =   "frmgewog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   9735
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   9615
      _cx             =   16960
      _cy             =   8493
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmgewog.frx":0E42
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
   Begin VB.Frame itemcode 
      Caption         =   "DZONGKHAG INFORMATION"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9495
      Begin VB.TextBox txtgewogname 
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
         TabIndex        =   3
         Top             =   1920
         Width           =   6735
      End
      Begin VB.TextBox txtremarks 
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
         Top             =   2400
         Width           =   6735
      End
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "frmgewog.frx":0EE9
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2520
         TabIndex        =   4
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "frmgewog.frx":0EFE
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
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
      Begin MSDataListLib.DataCombo cbodemkhong 
         Bindings        =   "frmgewog.frx":0F13
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   960
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
      Begin VB.Label Label6 
         Caption         =   "DEMKHONG"
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
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GEWOG ID"
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
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG"
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
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GEWOG NAME"
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
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "REMARKS"
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
         Top             =   2520
         Width           =   915
      End
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
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmgewog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsDz As New ADODB.Recordset
Dim rsDem As New ADODB.Recordset
Dim rsGe As New ADODB.Recordset

Private Sub cboDzongkhag_GotFocus()
cbodzongkhag.BackColor = vbYellow
End Sub

Private Sub cbodzongkhag_LostFocus()
On Error GoTo err
cbodzongkhag.BackColor = vbWhite
If Len(cbodzongkhag.Text) = 0 Then
MsgBox "Please Select The Proper Dzongkhag First."
cbodzongkhag.SetFocus
Exit Sub
End If
cbodzongkhag.Enabled = False
If Operation = "ADD" Then
cbogewog.Enabled = False
Dim rs As New ADODB.Recordset
       Set rs = Nothing
       rs.Open "SELECT MAX(SUBSTRING(gewogid,2,2))+1 AS MaxID from tblgewog where dzongkhagid='" & cbodzongkhag.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
       If rs.EOF <> True Then
       If Len(rs!MaxId) = 1 Then
       cbogewog.Text = "G0" & rs!MaxId
        Else
        cbogewog.Text = "G" & IIf(IsNull(rs!MaxId), "01", rs!MaxId)
        End If
       Else
       cbogewog.Text = "G01" & rs!MaxId
       End If
ElseIf Operation = "OPEN" Then
cbogewog.Enabled = True
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
cbogewog = ""
If rsDem.State = adStateOpen Then rsDem.Close
rsDem.Open "select concat(demkhongid , ' ', demkhongname) as demkhongname,demkhongid  from tbldemkhong where dzongkhag='" & cbodzongkhag.BoundText & "' order by demkhongid", db
Set cbodemkhong.RowSource = rsDem
cbodemkhong.ListField = "demkhongname"
cbodemkhong.BoundColumn = "demkhongid"

If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cbodzongkhag.BoundText & "' order by gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"


Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub cbogewog_GotFocus()
If Len(cbodzongkhag.Text) = 0 Then
MsgBox "Please Select The Dzongkhag FIrst."
cbodzongkhag.Enabled = True
cbodzongkhag.SetFocus
Exit Sub
End If
cbogewog.BackColor = vbYellow
End Sub

Private Sub cbogewog_LostFocus()
On Error GoTo err
Dim rs As New ADODB.Recordset
cbogewog.BackColor = vbWhite

Set rs = Nothing
rs.Open "select * from tblgewog where dzongkhagid='" & cbodzongkhag.BoundText & "' and gewogid='" & cbogewog.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic

If rs.EOF <> True Then
txtgewogname.Text = rs!gewogname
txtremarks.Text = rs!remarks
Else
MsgBox "Record Not Found."
Exit Sub
End If
cbogewog.Enabled = False
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Check1_Click()
'Mygrid.MergeCol(1) = True
'If Mygrid.MergeCells = 0 Then
'        Mygrid.MergeCells = 2
'    Else
'        Mygrid.MergeCells = 0
'    End If
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cbodzongkhag.RowSource = rsDz
cbodzongkhag.ListField = "dzongkhagname"
cbodzongkhag.BoundColumn = "dzongkhagcode"

If rsDem.State = adStateOpen Then rsDem.Close
rsDem.Open "select concat(demkhongid , ' ', demkhongname) as demkhongname,demkhongid  from tbldemkhong order by demkhongid", db
Set cbodemkhong.RowSource = rsDem
cbodemkhong.ListField = "demkhongname"
cbodemkhong.BoundColumn = "demkhongid"

If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog order by gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"


FillGrid
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FillGrid()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^SL.NO.|^DZONGKHAG ID|^GEWOG ID|^GEWOG NAME|^REMARKS"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 2760
mygrid.ColWidth(2) = 1050
mygrid.ColWidth(3) = 1980
mygrid.ColWidth(4) = 2895

rs.Open "select * from tblgewog order by dzongkhagid,gewogid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
mygrid.TextMatrix(i, 0) = i
FindDZ (rs!dzongkhagid)
mygrid.TextMatrix(i, 1) = Dzname
mygrid.TextMatrix(i, 2) = rs!gewogid

mygrid.TextMatrix(i, 3) = rs!gewogname
mygrid.TextMatrix(i, 4) = rs!remarks
rs.MoveNext
i = i + 1
Loop

rs.Close
mygrid.MergeCol(1) = True
mygrid.MergeCells = 1
  
Exit Sub
err:
MsgBox err.Descriptiont

End Sub
Private Sub FindDZ(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dzname = ""
Set rs = Nothing
rs.Open "select * from tbldzongkhag where dzongkhagcode='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Dzname = rs!DZONGKHAGNAME
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub






Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ADD"

       cbodzongkhag.Enabled = False
        TB.buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       cbodzongkhag.Enabled = True
       
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbodzongkhag.Enabled = True
       cbodemkhong.Enabled = True
       cbogewog.Enabled = True
      TB.buttons(3).Enabled = True
       
       Case "SAVE"
       MNU_SAVE
        TB.buttons(3).Enabled = False
        FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub CLEARCONTROLL()
cbodzongkhag.Text = ""
cbodemkhong.Text = ""
cbogewog.Text = ""
txtgewogname.Text = ""
txtremarks.Text = ""
End Sub
Private Sub MNU_SAVE()
Dim mainid As String
On Error GoTo err

MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute "INSERT INTO tblgewog (gewogid,gewogname,DZONGKHAGid,REMARKS,mainid) VALUEs('" & cbogewog.Text & "','" & txtgewogname.Text & "','" & cbodzongkhag.BoundText & "','" & txtremarks.Text & "','" & cbodzongkhag.BoundText & cbogewog.Text & "')"

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblgewog set gewogname='" & txtgewogname.Text & "',remarks='" & txtremarks.Text & "',mainid='" & cbodzongkhag.BoundText & cbogewog.BoundText & "', demkhongid='" & cbodemkhong.BoundText & "' where gewogid='" & cbogewog.BoundText & "' and dzongkhagid='" & cbodzongkhag.BoundText & "'"
Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub

Private Sub txtgewogname_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtgewogname.SelStart + 1
    Dim sText As String
    sText = Left$(txtgewogname.Text, iPos)
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
