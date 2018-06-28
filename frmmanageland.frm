VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmmanageland 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Land Management"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14505
   Icon            =   "frmmanageland.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
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
      Height          =   3855
      Left            =   9960
      TabIndex        =   15
      Top             =   1920
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
         Height          =   3435
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Location"
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
      Left            =   9960
      TabIndex        =   10
      Top             =   720
      Width           =   4455
      Begin MSDataListLib.DataCombo cbodzongkhag 
         Bindings        =   "frmmanageland.frx":0E42
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   11
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
         Bindings        =   "frmmanageland.frx":0E57
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   12
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
         TabIndex        =   14
         Top             =   480
         Width           =   975
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
      Height          =   7095
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   9735
      Begin VSFlex7Ctl.VSFlexGrid mygrid 
         Height          =   6615
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   9255
         _cx             =   16325
         _cy             =   11668
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmmanageland.frx":0E6C
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
         Picture         =   "frmmanageland.frx":0F3E
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBtnDn 
         Height          =   240
         Left            =   5040
         Picture         =   "frmmanageland.frx":12C8
         Top             =   960
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   4800
      TabIndex        =   1
      Top             =   2760
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
         Picture         =   "frmmanageland.frx":1652
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit Farmer Detail"
         Top             =   1440
         Width           =   615
      End
      Begin VSFlex7Ctl.VSFlexGrid vgrid 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
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
         FormatString    =   $"frmmanageland.frx":231C
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
         TabIndex        =   4
         Top             =   1560
         Width           =   45
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   10560
      Picture         =   "frmmanageland.frx":23B1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1095
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   9720
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
            Picture         =   "frmmanageland.frx":273B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanageland.frx":2AD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanageland.frx":2E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanageland.frx":3B49
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanageland.frx":3F9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanageland.frx":4755
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
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
      TabIndex        =   9
      Top             =   8040
      Width           =   1680
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
      TabIndex        =   8
      Top             =   8280
      Width           =   1680
   End
End
Attribute VB_Name = "frmmanageland"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub Form_Load()
filldzongkhag
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

Private Sub lstts_Click()
fillfarmer
End Sub

Private Sub mygrid_Click()
If mygrid.col = 5 Then
mygrid.Editable = flexEDKbdMouse

mygrid.ComboList = "New|Completely Planted|Partially Planted"
Else
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       Operation = "ADD"
       CLEARCONTROLL
       cbodzongkhag.Enabled = True
       TB.buttons(3).Enabled = True
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
             TB.buttons(3).Enabled = False
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


MHVDB.BeginTrans

For i = 1 To mygrid.rows - 1
If Len(Trim(mygrid.TextMatrix(i, 1))) = 0 Then Exit For
MHVDB.Execute "update tbllandreg set plantedstatus='" & Mid(Trim(mygrid.TextMatrix(i, 5)), 1, 1) & "' where farmerid='" & Mid(Trim(mygrid.TextMatrix(i, 2)), 1, 14) & "' and trnid='" & Val(mygrid.TextMatrix(i, 1)) & "'"
Next
MHVDB.CommitTrans

'TB.Buttons(3).Enabled = False
End Sub

Private Sub CLEARCONTROLL()
cbodzongkhag.Text = ""
cbogewog.Text = ""
lstts.Clear

mygrid.Clear
mygrid.rows = 1
mygrid.FormatString = "^Sl.No.|^Trn. Id|^Farmer|^Reg. Date|^Acre|^Planted Status|^"
mygrid.ColWidth(0) = 660
mygrid.ColWidth(1) = 960
mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 960
mygrid.ColWidth(4) = 750
mygrid.ColWidth(5) = 1860
mygrid.ColWidth(6) = 405
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


    For i = 0 To lstts.ListCount - 1
    If lstts.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(lstts.List(i), InStr(1, lstts.List(i), "|") + 1)) + "',"
       Set rs = Nothing
  
       
           
       
       
       
       
       Mcat = lstts.List(i)
       j = j + 1
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
  ' MsgBox "Tshewog not selected !!!"
            mygrid.Clear
            mygrid.rows = 1
mygrid.FormatString = "^Sl.No.|^Trn. Id|^Farmer|^Reg. Date|^Acre|^Planted Status|^"
mygrid.ColWidth(0) = 660
mygrid.ColWidth(1) = 960
mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 960
mygrid.ColWidth(4) = 750
mygrid.ColWidth(5) = 1860
mygrid.ColWidth(6) = 405
            Label5.Caption = ""
            Label6.Caption = ""
   Exit Sub
End If

SQLSTR = "select * from tbllandreg where status not in('D','R','C') and " _
& " substring(farmerid,1,3)='" & cbodzongkhag.BoundText & "' and " _
& " substring(farmerid,4,3) ='" & cbogewog.BoundText & "'  and substring(farmerid,7,3) in " & Dzstr & " and plantedstatus='N'" _
& " order by regdate desc"



mygrid.Clear
mygrid.rows = 1
misscnt = 0
mygrid.FormatString = "^Sl.No.|^Trn. Id|^Farmer|^Reg. Date|^Acre|^Planted Status|^"
mygrid.ColWidth(0) = 660
mygrid.ColWidth(1) = 960
mygrid.ColWidth(2) = 3495
mygrid.ColWidth(3) = 960
mygrid.ColWidth(4) = 750
mygrid.ColWidth(5) = 1860
mygrid.ColWidth(6) = 405

Set rs = Nothing
rs.Open SQLSTR, MHVDB
i = 1
Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = rs!trnid
FindFA rs!farmerid, "F"
mygrid.TextMatrix(i, 2) = rs!farmerid & "  " & FAName
mygrid.TextMatrix(i, 3) = Format(rs!regdate, "dd/MM/yyyy")
mygrid.TextMatrix(i, 4) = Format(rs!regland, "##0.00")
Select Case rs!plantedstatus
Case "N"
mygrid.TextMatrix(i, 5) = "New"
'mygrid.Cell(flexcpBackColor, 1, 5, i) = vbGreen 'RGB(255, 255, 0)

Case "C"
mygrid.TextMatrix(i, 5) = "Completely Planted"
'mygrid.Cell(flexcpBackColor, 1, 5, i) = vbRed 'RGB(255, 255, 0)
Case "P"
mygrid.TextMatrix(i, 5) = "Partially Planted"
'mygrid.Cell(flexcpBackColor, 1, 5, i) = vbYellow 'RGB(255, 255, 0)
End Select


mygrid.ColAlignment(2) = flexAlignLeftTop
mygrid.ColAlignment(1) = flexAlignLeftTop
rs.MoveNext
i = i + 1
Loop
'Label5.Caption = "Total farmer under " & cbomonitor.Text & "is " & i - 1
'Label6.Caption = "Total missing farmer in planted list under " & cbomonitor.Text & "is " & misscnt
rs.Close
'addbtn
mygrid.MergeCol(2) = True
mygrid.MergeCells = 1
Exit Sub
err:
MsgBox err.Description
End Sub
