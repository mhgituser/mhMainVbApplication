VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FRMLANDDETAILS 
   BorderStyle     =   0  'None
   Caption         =   "LAND DETAIL"
   ClientHeight    =   4395
   ClientLeft      =   4755
   ClientTop       =   3405
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6855
      _cx             =   12091
      _cy             =   5741
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FRMLANDDETAILS.frx":0000
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
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   2640
      Picture         =   "FRMLANDDETAILS.frx":009E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "LAND DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FRMLANDDETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'mFARID = ""
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim TOTL As Double
Dim i As Integer
i = 1
TOTL = 0
Set rs = Nothing

rs.Open "SELECT * FROM tblfarmer WHERE IDFARMER='" & mFARID & "' ", MHVDB
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1

Mygrid.TextMatrix(i, 0) = i
Mygrid.TextMatrix(i, 1) = "BY REGISTRATION"
Mygrid.TextMatrix(i, 2) = "NA"
Mygrid.TextMatrix(i, 3) = Format(IIf(IsNull(rs!REGAREA), 0, rs!REGAREA), "#####0.00")
TOTL = TOTL + rs!REGAREA
i = i + 1
rs.MoveNext
Loop
If TOTL = 0 Then
Mygrid.Rows = Mygrid.Rows - 1
i = 1
Else

End If
Set rs = Nothing
rs.Open "SELECT * FROM tbllandreg WHERE FARMERID='" & mFARID & "' ", MHVDB
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1

Mygrid.TextMatrix(i, 0) = i
Mygrid.TextMatrix(i, 1) = rs!trnid
Mygrid.TextMatrix(i, 2) = rs!regdate
Mygrid.TextMatrix(i, 3) = Format(IIf(IsNull(rs!regland), 0, rs!regland), "#####0.00")
TOTL = TOTL + IIf(IsNull(rs!regland), 0, rs!regland)
i = i + 1
rs.MoveNext
Loop
Label2.Caption = Format(TOTL, "#####0.00")
Exit Sub
err:
MsgBox err.Description
End Sub

