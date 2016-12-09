VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDzongkhag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DZONGKHAG MASTER"
   ClientHeight    =   5805
   ClientLeft      =   3840
   ClientTop       =   1380
   ClientWidth     =   9960
   Icon            =   "frmDzongkhag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9960
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   2880
      Width           =   9615
      _cx             =   16960
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDzongkhag.frx":0E42
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
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   9975
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
         TabIndex        =   6
         Top             =   1320
         Width           =   6975
      End
      Begin VB.TextBox txtdzname 
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
         TabIndex        =   5
         Top             =   840
         Width           =   6975
      End
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "frmDzongkhag.frx":0ED3
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2520
         TabIndex        =   8
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
         TabIndex        =   4
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG NAME"
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
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG ID"
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
         Top             =   360
         Width           =   1440
      End
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
            Picture         =   "frmDzongkhag.frx":0EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDzongkhag.frx":1282
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDzongkhag.frx":161C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDzongkhag.frx":22F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDzongkhag.frx":2748
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDzongkhag.frx":2F02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
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
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmDzongkhag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDz As New ADODB.Recordset
'Mygrid.FormatString = "^SL.NO.|^DZONGKHAG ID|^DZONGKHAG NAME|^REMARKS"


Private Sub cboDzongkhag_GotFocus()
   cboDzongkhag.BackColor = vbYellow
End Sub

Private Sub cboDzongkhag_LostFocus()
   On Error GoTo err
   cboDzongkhag.BackColor = vbWhite
   cboDzongkhag.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tbldzongkhag where dzongkhagcode='" & cboDzongkhag.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   txtdzname.Text = rs!DZONGKHAGNAME
   txtremarks.Text = rs!remarks
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub

Private Sub Form_Load()
On Error GoTo err
operation = ""

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cboDzongkhag.RowSource = rsDz
cboDzongkhag.ListField = "dzongkhagname"
cboDzongkhag.BoundColumn = "dzongkhagcode"
FillGrid
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       cboDzongkhag.Enabled = False
        TB.Buttons(3).Enabled = True
       operation = "ADD"
       CLEARCONTROLL
       Dim rs As New ADODB.Recordset
       Set rs = Nothing
       rs.Open "SELECT MAX(SUBSTRING(dzongkhagcode,2,2))+1 AS MaxID from tbldzongkhag", MHVDB, adOpenForwardOnly, adLockOptimistic
       If rs.EOF <> True Then
       
       
        cboDzongkhag.Text = "D" & rs!MaxId
        
       Else
       cboDzongkhag.Text = "D01" & rs!MaxId
       End If
       
       
       Case "OPEN"
       operation = "OPEN"
       CLEARCONTROLL
       cboDzongkhag.Enabled = True
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
err:
MsgBox err.Description
End Sub
Private Sub MNU_SAVE()
Dim rs As New ADODB.Recordset
On Error GoTo err


MHVDB.BeginTrans
If operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tbldzongkhag (DZONGKHAGCODE,DZONGKHAGNAME,REMARKS) VALUEs('" & cboDzongkhag.Text & "','" & txtdzname.Text & "','" & txtremarks.Text & "')"
LogRemarks = "Inserted new record" & cboDzongkhag & "," & txtdzname.Text & "," & txtremarks
updatemhvlog Now, MUSER, LogRemarks, ""
ElseIf operation = "OPEN" Then
MHVDB.Execute "update tbldzongkhag set dzongkhagname='" & txtdzname.Text & "',remarks='" & txtremarks.Text & "' where dzongkhagcode='" & cboDzongkhag.BoundText & "'"
Set rs = Nothing

rs.Open "select * from tbldzongkhag where dzongkhagcode='" & cboDzongkhag.BoundText & "' ", MHVDB
If Trim(rs!DZONGKHAGNAME) <> Trim(txtdzname.Text) Then
LogRemarks = "Updated Dzongkhag Name from " & rs!DZONGKHAGNAME & "to" & txtdzname.Text
updatemhvlog Now, MUSER, LogRemarks, ""
End If

If Trim(rs!remarks) <> Trim(txtremarks.Text) Then
LogRemarks = "Updated Remarks from " & rs!remarks & "to" & txtremarks.Text
updatemhvlog Now, MUSER, LogRemarks, ""
End If


Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
'cboDzongkhag.Refreshy
Exit Sub

err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub
Private Sub CLEARCONTROLL()

txtdzname.Text = ""
txtremarks.Text = ""

End Sub
Private Sub FillGrid()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^SL.NO.|^DZONGKHAG ID|^DZONGKHAG NAME|^REMARKS"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 2835
Mygrid.ColWidth(2) = 2835
Mygrid.ColWidth(3) = 2895

rs.Open "select * from tbldzongkhag order by dzongkhagcode", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = rs!DZONGKHAGCODE
Mygrid.TextMatrix(i, 2) = rs!DZONGKHAGNAME
Mygrid.TextMatrix(i, 3) = rs!remarks

rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub txtdzname_GotFocus()
txtdzname.BackColor = vbYellow
End Sub

Private Sub txtdzname_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtdzname.SelStart + 1
    Dim sText As String
    sText = Left$(txtdzname.Text, iPos)
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

Private Sub txtdzname_LostFocus()
txtdzname.BackColor = vbWhite
End Sub

Private Sub txtremarks_GotFocus()
txtremarks.BackColor = vbYellow
End Sub

Private Sub txtremarks_LostFocus()
txtremarks.BackColor = vbWhite
End Sub
