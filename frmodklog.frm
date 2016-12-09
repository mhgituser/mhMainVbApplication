VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmodklog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODK LOG"
   ClientHeight    =   6465
   ClientLeft      =   3675
   ClientTop       =   2865
   ClientWidth     =   12855
   Icon            =   "frmodklog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   12855
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Picture         =   "frmodklog.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtlog 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5400
      Width           =   12615
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   4320
      Picture         =   "frmodklog.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "LOAD "
      Top             =   120
      Width           =   855
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12615
      _cx             =   22251
      _cy             =   8070
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
      BackColorAlternate=   16777215
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   3
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmodklog.frx":21BE
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
   Begin MSDataListLib.DataCombo CBOUSER 
      Bindings        =   "frmodklog.frx":227B
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "USER NAME"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmodklog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(CBOUSER.Text) = 0 Then
FillGrid "A"
Else
FillGrid ""
End If
End Sub
Private Sub FillGrid(ff As String)

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                     
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^SL.NO.|^TRN.ID|^_URI|^DATE|^DESCRIPTION|^TABLENAME"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 780
Mygrid.ColWidth(2) = 2835
Mygrid.ColWidth(3) = 1100
Mygrid.ColWidth(4) = 4500
Mygrid.ColWidth(5) = 2565

If ff = "A" Then
rs.Open "select * from tblodkmodificationlog order by date_of_modification", db, adOpenForwardOnly, adLockOptimistic

Else
rs.Open "select * from tblodkmodificationlog where modified_by='" & CBOUSER.Text & "'", db, adOpenForwardOnly, adLockOptimistic

End If
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = rs!TRNID 'rs!FARMERcode

Mygrid.TextMatrix(i, 2) = rs![_uri]
Mygrid.TextMatrix(i, 3) = Format(rs!date_of_modification, "dd/MM/yyyy")


Mygrid.TextMatrix(i, 4) = rs!Description

Mygrid.TextMatrix(i, 5) = rs!table_name
rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                       
'odk_prodLocal
Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct modified_by  from tblodkmodificationlog", db
Set CBOUSER.RowSource = rs
CBOUSER.ListField = "modified_by"
CBOUSER.BoundColumn = "modified_by"

 FillGrid "A"
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Mygrid_Click()
txtlog.Text = Mygrid.TextMatrix(Mygrid.row, 4)
End Sub
