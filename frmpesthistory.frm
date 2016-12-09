VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpesthistory 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   3390
   ClientTop       =   2400
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "FARMER DETAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Command3 
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
         Left            =   6600
         Picture         =   "frmpesthistory.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit Farmer Detail"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox TXTSTATUS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TXTFARMERNAME 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   5655
      End
      Begin VB.TextBox TXTAREA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TXTPLANTS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TXTDZ 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TXTGE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TXTTS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
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
         TabIndex        =   24
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "FARMER NAME"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ACRE REG."
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
         Left            =   2640
         TabIndex        =   22
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "NO. OF PLANTS"
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
         Left            =   4800
         TabIndex        =   21
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG "
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4320
         TabIndex        =   19
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton Command2 
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
         Left            =   5520
         Picture         =   "frmpesthistory.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "View Pest History"
         Top             =   840
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "frmpesthistory.frx":1474
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboparaid 
         Bindings        =   "frmpesthistory.frx":1489
         Height          =   315
         Left            =   7080
         TabIndex        =   7
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PARAMETER NAME"
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
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FARMER ID"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "PEST HISTORY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   14655
      Begin VSFlex7Ctl.VSFlexGrid mygrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   14535
         _cx             =   25638
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648384
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmpesthistory.frx":149E
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
   End
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
      Left            =   7200
      Picture         =   "frmpesthistory.frx":15B9
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Farmer Detail"
      Top             =   4680
      Width           =   615
   End
End
Attribute VB_Name = "frmpesthistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cbofarmerid_Click(Area As Integer)
Frame5.Visible = False
End Sub

Private Sub cbofarmerid_DblClick(Area As Integer)
If Len(cbofarmerid.Text) > 0 Then
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblfarmer   where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
TXTFARMERNAME.Text = ""
TXTDZ.Text = ""
TXTAREA.Text = ""
TXTGE.Text = ""
TXTTS.Text = ""
TXTSTATUS.Text = ""
TXTPLANTS.Text = ""
Frame5.Visible = True
TXTFARMERNAME.Text = rs!idfarmer & " " & rs!farmername
FindDZ Mid(rs!idfarmer, 1, 3)
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
TXTDZ.Text = Mid(rs!idfarmer, 1, 3) & " " & Dzname
TXTGE.Text = Mid(rs!idfarmer, 4, 3) & " " & GEname
TXTTS.Text = Mid(rs!idfarmer, 7, 3) & " " & TsName
If rs!Status = "A" Then
TXTSTATUS.Text = "ACTIVE"
ElseIf rs!Status = "R" Then
TXTSTATUS.Text = "REJECTED"
Else
TXTSTATUS.Text = "DROPPED OUT"
End If
Set rs1 = Nothing
rs1.Open "select sum(regland) as land from tbllandreg where farmerid='" & rs!idfarmer & "'", MHVDB
If rs1.EOF <> True Then
TXTAREA.Text = Format(rs1!land, "###0.00")
End If

Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as tr from tblplanted where farmercode='" & rs!idfarmer & "'", MHVDB
If rs1.EOF <> True Then
TXTPLANTS = rs1!tr
End If


End If
End If
End Sub

Private Sub Command2_Click()
If Len(cboparaid.Text) = 0 Then Exit Sub
If Len(cbofarmerid.Text) = 0 Then Exit Sub
filldetail
End Sub

Private Sub Command3_Click()
Frame5.Visible = False
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
Dim rsfr As New ADODB.Recordset


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set db1 = New ADODB.Connection
db1.CursorLocation = adUseClient
db1.Open OdkCnnString






Set rsfr = Nothing

If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer where status='A' AND IDFARMER IN(select distinct farmercode from odk_prodlocal.tblodkfollowuplog) order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open " select DISTINCT concat(cast(paraid as char),'  ',paraname,'  ',fstype,'  ',cast(value as char)) as description,paraid from tblodkalarmparameter where status='ON'", db1
Set cboparaid.RowSource = rsfr
cboparaid.ListField = "description"
cboparaid.BoundColumn = "paraid"

cbofarmerid.Text = pestFarmer
cboparaid.Text = pestparamname
     filldetail
       

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub filldetail()
On Error GoTo err
findParamDetails cboparaid.BoundText
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing

mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^SL.NO.|^ENTRY DATE|^START DATE|^" & paramName & "|^STAFF|^FD. CODE|^ACTION TAKEN|^RECOMMENDATION|^uri|^email|^"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 1320
mygrid.ColWidth(2) = 1665
mygrid.ColWidth(3) = 1125
mygrid.ColWidth(4) = 2460
'mygrid.ColWidth(5) = 2820
mygrid.ColWidth(5) = 960
mygrid.ColWidth(6) = 1530
mygrid.ColWidth(7) = 4395
mygrid.ColWidth(8) = 120
mygrid.ColWidth(9) = 1
mygrid.ColWidth(10) = 1



rs.Open "select * from tblodkfollowuplog where farmercode='" & cbofarmerid.BoundText & "'   and paraid='" & cboparaid.BoundText & "' order by odkstartdate desc", ODKDB

i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i
FindsTAFF rs!staffcode
mygrid.TextMatrix(i, 1) = Format(rs!entrydate, "dd/MM/yyyy")
mygrid.TextMatrix(i, 2) = Format(rs!odkStartDate, "dd/MM/yyyy")
If ispercentage Then
mygrid.TextMatrix(i, 3) = Format(rs!odkValue, "####0.00") & "%"
Else
mygrid.TextMatrix(i, 3) = Format(rs!odkValue, "####0.00")
End If
mygrid.ColAlignment(3) = flexAlignRightTop
mygrid.TextMatrix(i, 4) = rs!staffcode & " " & sTAFF
mygrid.ColAlignment(4) = flexAlignLeftTop
'mygrid.TextMatrix(i, 5) = rs!farmercode & " " & FAName
'mygrid.ColAlignment(5) = flexAlignLeftTop
mygrid.TextMatrix(i, 5) = rs!fieldcode
mygrid.TextMatrix(i, 6) = rs!actiontaken
mygrid.TextMatrix(i, 7) = rs!recommendation
mygrid.TextMatrix(i, 9) = rs!uri
mygrid.TextMatrix(i, 10) = rs!emailstatus
i = i + 1
rs.MoveNext
Loop
'addcells
Exit Sub
err:
MsgBox err.Description
End Sub

