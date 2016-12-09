VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmprequestionabsentee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P R E V I O U S      Q U E S T I O N S   "
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12090
   Icon            =   "frmprequestionabsentee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   3210
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11895
      _cx             =   20981
      _cy             =   5662
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorBkg    =   16777215
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
      GridLines       =   0
      GridLinesFixed  =   0
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
      FormatString    =   $"frmprequestionabsentee.frx":0E42
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
Attribute VB_Name = "frmprequestionabsentee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FillGrid()
On Error GoTo err

Dim RS As New ADODB.Recordset
Dim i As Integer
'If userid = "" Then Exit Sub

Set RS = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^SL.NO.|^A.ID|^Q.ID|^QUESTION|^ANSWER|^COMPLETED"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 0
Mygrid.ColWidth(2) = 0
Mygrid.ColWidth(3) = 4005
Mygrid.ColWidth(4) = 4005
Mygrid.ColWidth(5) = 1320
'Mygrid.ColWidth(6) = 960


RS.Open "SELECT * FROM TBLABSENTEEQUESTION where absenteeid='" & ANAME & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While RS.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i
Mygrid.TextMatrix(i, 1) = RS!ABSENTEEID
Mygrid.TextMatrix(i, 2) = RS!TRNID
Mygrid.TextMatrix(i, 3) = RS!QUESTION
Mygrid.TextMatrix(i, 4) = RS!ANSWER
Mygrid.TextMatrix(i, 5) = RS!ANSWERCOMPLETE

If IIf(IsNull(RS!ANSWERCOMPLETE), 0, RS!ANSWERCOMPLETE) = 0 Then
Mygrid.TextMatrix(i, 5) = "NO"
Else
Mygrid.TextMatrix(i, 5) = "YES"
End If

RS.MoveNext
i = i + 1
Loop

RS.Close
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub Form_Load()
FillGrid
End Sub
