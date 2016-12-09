VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmdistributionreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DISTRIBUTION REPORT"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17100
   Icon            =   "frmdistributionreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   17100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "jonnes"
      Height          =   855
      Left            =   600
      TabIndex        =   31
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "survival farmer detail"
      Height          =   1095
      Left            =   4560
      TabIndex        =   30
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Temp STorage Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Picture         =   "frmdistributionreport.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5400
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker txtregdate 
      Height          =   375
      Left            =   10680
      TabIndex        =   27
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   104267777
      CurrentDate     =   41624
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Temp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Picture         =   "frmdistributionreport.frx":0ED4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Servival"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "frmdistributionreport.frx":163E
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Planted List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Picture         =   "frmdistributionreport.frx":1DA8
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "frmdistributionreport.frx":2512
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txttsplants 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   12840
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtgeplants 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Picture         =   "frmdistributionreport.frx":2C7C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Picture         =   "frmdistributionreport.frx":3946
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtplants 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtfarmer 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame4 
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
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox cbomnth 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmdistributionreport.frx":40B0
         Left            =   4080
         List            =   "frmdistributionreport.frx":40D8
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo cboyear 
         Bindings        =   "frmdistributionreport.frx":4118
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   3240
         TabIndex        =   23
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         TabIndex        =   12
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   3255
      Left            =   11400
      TabIndex        =   6
      Top             =   840
      Width           =   5535
      Begin VSFlex7Ctl.VSFlexGrid tsgrid 
         Height          =   2535
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   5415
         _cx             =   9551
         _cy             =   4471
         _ConvInfo       =   1
         Appearance      =   2
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmdistributionreport.frx":412D
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "frmdistributionreport.frx":420B
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
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
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3255
      Left            =   5760
      TabIndex        =   2
      Top             =   840
      Width           =   5535
      Begin VSFlex7Ctl.VSFlexGrid gegrid 
         Height          =   2535
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   5415
         _cx             =   9551
         _cy             =   4471
         _ConvInfo       =   1
         Appearance      =   2
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmdistributionreport.frx":4220
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
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "frmdistributionreport.frx":4300
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5535
      Begin VSFlex7Ctl.VSFlexGrid dzgrid 
         Height          =   2535
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   5415
         _cx             =   9551
         _cy             =   4471
         _ConvInfo       =   1
         Appearance      =   2
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmdistributionreport.frx":4315
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
   End
   Begin MSComCtl2.DTPicker txtregdate1 
      Height          =   375
      Left            =   10680
      TabIndex        =   28
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115277825
      CurrentDate     =   41624
   End
End
Attribute VB_Name = "frmdistributionreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbodzongkhag_LostFocus()
Dim rsGe As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
cbogewog = ""
If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cboDzongkhag.BoundText & "' order by gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"




End Sub

Private Sub Command1_Click()
If Len(cboyear.Text) = 0 Then Exit Sub

filldz
'If Len(cboDzongkhag.Text) > 0 Then
'fillge
'End If
'If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 Then
'fillts
'End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
showDetail
End Sub
Private Sub showDetail()
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object

    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Dzongkhag")
    excel_sheet.Cells(3, 2) = ProperCase("gewog")
    excel_sheet.Cells(3, 3) = ProperCase("tshowog")
    excel_sheet.Cells(3, 4) = ProperCase("old farmer")
    excel_sheet.Cells(3, 5) = ProperCase("new farmer")
    excel_sheet.Cells(3, 6) = ProperCase("# of plants")
    excel_sheet.Cells(3, 7) = ProperCase("Acre")
    
    i = 4
  
    SQLSTR = "select farmercode,sum((crateno * 35)+plno+crate) plants,area,newold from tblplantdistributiondetail where trnid in(select trnid from tblplantdistributionheader where status='ON') and subtotindicator not in ('S','T') and status<>'C' and year='" & cboyear.BoundText & "' and  senttofield <>'N' and farmercode in(select idfarmer from tblfarmer where status='A')group by farmercode order by farmercode"
   
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    FindDZ Mid(rs!farmercode, 1, 3)
    FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
    excel_sheet.Cells(i, 1) = Dzname
    excel_sheet.Cells(i, 2) = GEname
    excel_sheet.Cells(i, 3) = TsName
    If rs!newold = "O" Then
    excel_sheet.Cells(i, 4) = rs!farmercode
    Else
    
    excel_sheet.Cells(i, 5) = rs!farmercode
    End If
    excel_sheet.Cells(i, 6) = rs!plants
    excel_sheet.Cells(i, 7) = rs!area
    
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command4_Click()
showPlantedList
End Sub
Private Sub showPlantedList()
Dim frcode As String
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim mn As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
frcode = ""
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Dzongkhag")
    excel_sheet.Cells(3, 2) = ProperCase("gewog")
    excel_sheet.Cells(3, 3) = ProperCase("tshowog")
    'excel_sheet.Cells(3, 4) = ProperCase("farmers")
    
    excel_sheet.Cells(3, 4) = ProperCase("Planted List")
   ' excel_sheet.Cells(3, 5) = ProperCase("New Farmers(Distribution)")
    'excel_sheet.Cells(3, 6) = ProperCase("Old Farmers(Distribution)")
    
    excel_sheet.Cells(3, 7) = ProperCase("2011")
    excel_sheet.Cells(3, 8) = ProperCase("2012")
    excel_sheet.Cells(3, 9) = ProperCase("2013")
        excel_sheet.Cells(3, 10) = ProperCase("2014")
    'excel_sheet.Cells(3, 9) = ProperCase("2013( new farmers )")
   ' excel_sheet.Cells(3, 10) = ProperCase("2013( old farmers )")
    'excel_sheet.Cells(3, 11) = ProperCase("Acre(New)")
    excel_sheet.Cells(3, 11) = ProperCase("Acre")
    excel_sheet.Cells(3, 12) = ProperCase("2015")
    
    i = 4
  
   ' SQLSTR = "select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,,sum(acre) acre,sum(acreold) acreold ,cntplanted,cntdistnew,cntdistold ,year from tbldisttemp where  year in(2011,2012,2013) group by year,farmercode order by farmercode"
    SQLSTR = "select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,year from tblplanted where  year in(2011,2012,2013,2014,2015) group by year,farmercode order by farmercode"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do Until rs.EOF
    
  
    
    frcode = rs!farmercode
    FindDZ Mid(rs!farmercode, 1, 3)
    FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
    excel_sheet.Cells(i, 1) = Dzname
    excel_sheet.Cells(i, 2) = GEname
    excel_sheet.Cells(i, 3) = TsName
    'excel_sheet.Cells(i, 4) = rs!farmercode
    Set mn = Nothing
    mn.Open "select sum(regland) regland from tbllandreg where farmerid='" & rs!farmercode & "' group by farmerid", MHVDB
    If mn.EOF <> True Then
    excel_sheet.Cells(i, 11) = mn!regland
    End If
   
    
    
    Do While frcode = rs!farmercode
 
    If rs!Year = 2011 Then
    excel_sheet.Cells(i, 7) = rs!pplant
    excel_sheet.Cells(i, 4) = rs!farmercode
    ElseIf rs!Year = 2012 Then
    excel_sheet.Cells(i, 8) = rs!pplant
     excel_sheet.Cells(i, 4) = rs!farmercode
    ElseIf rs!Year = 2013 Then
    
      excel_sheet.Cells(i, 9) = rs!pplant
      
       excel_sheet.Cells(i, 4) = rs!farmercode
  
        ElseIf rs!Year = 2014 Then
    
      excel_sheet.Cells(i, 10) = rs!pplant
      
       excel_sheet.Cells(i, 4) = rs!farmercode
       
           ElseIf rs!Year = 2015 Then
    
      excel_sheet.Cells(i, 12) = rs!pplant
      
       excel_sheet.Cells(i, 4) = rs!farmercode
       
    End If
    
 
   
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
    
    
     sl = sl + 1
    i = i + 1
    Loop
    
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command5_Click()
Dim frcode As String
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim mn As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Dim SQLSTR As String
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""


frcode = ""
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Dzongkhag")
    excel_sheet.Cells(3, 2) = ProperCase("gewog")
    excel_sheet.Cells(3, 3) = ProperCase("tshowog")
    
    excel_sheet.Cells(2, 4) = ProperCase("Plants Distributed")
    excel_sheet.Cells(3, 4) = ProperCase("2011")
    excel_sheet.Cells(3, 5) = ProperCase("2012")
    excel_sheet.Cells(3, 6) = ProperCase("2013")
    excel_sheet.Cells(3, 7) = ProperCase("total")
    
    excel_sheet.Cells(2, 8) = ProperCase("Field")
    excel_sheet.Cells(3, 8) = ProperCase("Total trees")
    excel_sheet.Cells(3, 9) = ProperCase("dead")
    excel_sheet.Cells(3, 10) = ProperCase("survival %")
    
    excel_sheet.Cells(2, 11) = ProperCase("storage")
    excel_sheet.Cells(3, 11) = ProperCase("Total trees")
    excel_sheet.Cells(3, 12) = ProperCase("dead")
    excel_sheet.Cells(3, 13) = ProperCase("survival %")
    
    
     excel_sheet.Cells(2, 14) = ProperCase("Total")
    excel_sheet.Cells(3, 14) = ProperCase("Total Trees(Field +Storage)")
    excel_sheet.Cells(3, 15) = ProperCase("Dead(Field + Storage)")
    excel_sheet.Cells(3, 16) = ProperCase("survival %")
    
    i = 4
    
    
    
    
    
    
     GetTbl
        
    
SQLSTR = ""

           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,fstype,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,'F',0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
         
         
         
    ODKDB.Execute SQLSTR
    
    SQLSTR = ""
          SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,fstype,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,'S',0,0,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,0,0,0,0,ndtrees,wlogged,0,0,0," _
         & "0,adamage,'0' from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"
          
  ODKDB.Execute SQLSTR
    
    
    SQLSTR = ""
  
   ' SQLSTR = "select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,,sum(acre) acre,sum(acreold) acreold ,cntplanted,cntdistnew,cntdistold ,year from tbldisttemp where  year in(2011,2012,2013) group by year,farmercode order by farmercode"
    SQLSTR = "select substring(farmercode,1,9) dgt,sum(challanqty) challanqty,year from tblplanted where year in(2011,2012,2013) group by year,dgt order by dgt"
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do Until rs.EOF
    
  
    
    frcode = rs!dgt
    FindDZ Mid(rs!dgt, 1, 3)
    FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
    FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
    excel_sheet.Cells(i, 1) = Dzname
    excel_sheet.Cells(i, 2) = GEname
    excel_sheet.Cells(i, 3) = TsName
  
    Set mn = Nothing
    mn.Open "select sum(regland) regland from tbllandreg where substring(farmerid,1,9)='" & rs!dgt & "' group by substring(farmerid,1,9)", MHVDB
    If mn.EOF <> True Then
    excel_sheet.Cells(i, 11) = mn!regland
    End If
    Do While frcode = rs!dgt
    If rs!Year = 2011 Then
    excel_sheet.Cells(i, 4) = rs!challanqty
    ElseIf rs!Year = 2012 Then
    excel_sheet.Cells(i, 5) = rs!challanqty
    
    ElseIf rs!Year = 2013 Then
    excel_sheet.Cells(i, 6) = rs!challanqty
    End If
    excel_sheet.Cells(i, 7) = Val(excel_sheet.Cells(i, 4)) + Val(excel_sheet.Cells(i, 5)) + Val(excel_sheet.Cells(i, 6))
    
    
    'field
 Set rsF = Nothing
 rsF.Open "select sum(totaltrees) totaltrees,sum(tree_count_deadmissing) deadmissing from  " & Mtblname & " where fstype='F' and substring(farmercode,1,9)='" & rs!dgt & "' group by substring(farmercode,1,9)", ODKDB
 
 If rsF.EOF <> True Then
 excel_sheet.Cells(i, 8) = rsF!totaltrees
 excel_sheet.Cells(i, 9) = rsF!deadmissing
 excel_sheet.Cells(i, 10) = 1 - rsF!deadmissing / Val(excel_sheet.Cells(i, 7))
 Else
 
  excel_sheet.Cells(i, 8) = ""
 excel_sheet.Cells(i, 9) = ""
 excel_sheet.Cells(i, 10) = ""
 End If
 
 'storage
 
 Set rsF = Nothing
 rsF.Open "select sum(totaltrees) totaltrees,sum(tree_count_deadmissing) deadmissing from  " & Mtblname & " where fstype='S' and substring(farmercode,1,9)='" & rs!dgt & "' group by substring(farmercode,1,9)", ODKDB
 
 If rsF.EOF <> True Then
 excel_sheet.Cells(i, 11) = rsF!totaltrees
 excel_sheet.Cells(i, 12) = rsF!deadmissing
 excel_sheet.Cells(i, 13) = 1 - rsF!deadmissing / Val(excel_sheet.Cells(i, 7))
 Else
 
 excel_sheet.Cells(i, 11) = ""
 excel_sheet.Cells(i, 12) = ""
 excel_sheet.Cells(i, 13) = ""
 End If
 
 'total
 excel_sheet.Cells(i, 14) = Val(excel_sheet.Cells(i, 8)) + Val(excel_sheet.Cells(i, 11))
 If Val(excel_sheet.Cells(i, 14)) > 0 Then
 
 excel_sheet.Cells(i, 15) = Val(excel_sheet.Cells(i, 9)) + Val(excel_sheet.Cells(i, 12))
 excel_sheet.Cells(i, 16) = 1 - Val(excel_sheet.Cells(i, 15)) / Val(excel_sheet.Cells(i, 7))
 Else
  excel_sheet.Cells(i, 14) = ""
 excel_sheet.Cells(i, 15) = ""
 excel_sheet.Cells(i, 16) = ""
 
 End If
 
 
   
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
    
    
     sl = sl + 1
    i = i + 1
    Loop
    
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    
    
    ODKDB.Execute "drop table " & Mtblname & ""
    ODKDB.Close
    
    
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command6_Click()
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim rs As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object

    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Farmer Name")
  excel_sheet.Cells(3, 2) = ProperCase("Land Type")
    excel_sheet.Cells(3, 3) = ProperCase("Acre")
    
    i = 4
  
   'SQLSTR = "select idfarmer,farmername from tblfarmer where substring(regdate,1,10)>='" & Format(txtregdate.Value, "yyyy-MM-dd") & "' and substring(regdate,1,10)<='" & Format(txtregdate1.Value, "yyyy-MM-dd") & "' and status not in('D','R')"
    SQLSTR = "select farmerid,sum(regland) as regland from tbllandreg where substring(regdate,1,10)>='" & Format(txtregdate.Value, "yyyy-MM-dd") & "' and substring(regdate,1,10)<='" & Format(txtregdate1.Value, "yyyy-MM-dd") & "' and status not in('D','R') group by farmerid"
    
   
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
 Set rsm = Nothing
' rsm.Open "select sum(regland) as regland from tbllandreg where farmerid='" & rs!idfarmer & "' group by farmerid", MHVDB
'i f rsm.EOF <> True Then
        excel_sheet.Cells(i, 1) = rs!farmerid
  excel_sheet.Cells(i, 2) = Mid(rs!farmerid, 10, 1)
    excel_sheet.Cells(i, 3) = rs!regland
    
   ' End If
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command7_Click()
Dim frcode As String
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim mn As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
frcode = ""
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Dzongkhag")
    excel_sheet.Cells(3, 2) = ProperCase("gewog")
    excel_sheet.Cells(3, 3) = ProperCase("tshowog")
    excel_sheet.Cells(3, 4) = ProperCase("No. of crates")
    excel_sheet.Cells(3, 5) = ProperCase("qty. sent")
    excel_sheet.Cells(3, 6) = ProperCase("Monitor")
    

    
    i = 4
  
   ' SQLSTR = "select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,,sum(acre) acre,sum(acreold) acreold ,cntplanted,cntdistnew,cntdistold ,year from tbldisttemp where  year in(2011,2012,2013) group by year,farmercode order by farmercode"
    SQLSTR = "select * from tblqmssendtotempstoragehdr where status='ON' order by vehicleno"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do Until rs.EOF
     
    FindTempStorage Mid(rs!vehicleno, 1, 3)
    FindDZ Mid(tempStorageName, 1, 3)
    FindGE Mid(tempStorageName, 1, 3), Mid(tempStorageName, 4, 3)
    FindTs Mid(tempStorageName, 1, 3), Mid(tempStorageName, 4, 3), Mid(tempStorageName, 7, 3)
    excel_sheet.Cells(i, 1) = Mid(tempStorageName, 1, 3) & "  " & Dzname
    excel_sheet.Cells(i, 2) = Mid(tempStorageName, 4, 3) & "  " & GEname
    excel_sheet.Cells(i, 3) = Mid(tempStorageName, 7, 3) & "  " & TsName
   excel_sheet.Cells(i, 4) = rs!cratecount
    excel_sheet.Cells(i, 5) = rs!sendtofieldqty
    Set rs1 = Nothing
    rs1.Open "select distinct monitor from tblfarmer where substring(idfarmer,1,9)='" & tempStorageName & "' ", MHVDB
    If rs1.EOF <> True Then
    FindsTAFF rs1!monitor
    excel_sheet.Cells(i, 6) = rs1!monitor & "  " & sTAFF
    End If
    
    
    i = i + 1
    rs.MoveNext
    Loop
    
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command8_Click()
Dim frcode As String
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim mn As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Dim SQLSTR As String
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""


frcode = ""
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Dzongkhag")
    excel_sheet.Cells(3, 2) = ProperCase("gewog")
    excel_sheet.Cells(3, 3) = ProperCase("tshowog")
    
     excel_sheet.Cells(3, 4) = ProperCase("farmer")
     
     
    excel_sheet.Cells(2, 5) = ProperCase("Plants Distributed")
    excel_sheet.Cells(3, 5) = ProperCase("2011")
    excel_sheet.Cells(3, 6) = ProperCase("2012")
    excel_sheet.Cells(3, 7) = ProperCase("2013")
    excel_sheet.Cells(3, 8) = ProperCase("total")
    
    excel_sheet.Cells(2, 9) = ProperCase("Field")
    excel_sheet.Cells(3, 9) = ProperCase("Total trees")
    excel_sheet.Cells(3, 10) = ProperCase("dead")
    excel_sheet.Cells(3, 11) = ProperCase("survival %")
    
    excel_sheet.Cells(2, 12) = ProperCase("storage")
    excel_sheet.Cells(3, 12) = ProperCase("Total trees")
    excel_sheet.Cells(3, 13) = ProperCase("dead")
    excel_sheet.Cells(3, 14) = ProperCase("survival %")
    
    
     excel_sheet.Cells(2, 15) = ProperCase("Total")
    excel_sheet.Cells(3, 15) = ProperCase("Total Trees(Field +Storage)")
    excel_sheet.Cells(3, 16) = ProperCase("Dead(Field + Storage)")
    excel_sheet.Cells(3, 17) = ProperCase("survival %")
    
    i = 4
    
    
    
    
    
    
     GetTbl
        
    
SQLSTR = ""

           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,fstype,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,'F',0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
         
         
         
    ODKDB.Execute SQLSTR
    
    SQLSTR = ""
          SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,fstype,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,'S',0,0,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,0,0,0,0,ndtrees,wlogged,0,0,0," _
         & "0,adamage,'0' from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"
          
  ODKDB.Execute SQLSTR
    
    
    SQLSTR = ""
  
   ' SQLSTR = "select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,,sum(acre) acre,sum(acreold) acreold ,cntplanted,cntdistnew,cntdistold ,year from tbldisttemp where  year in(2011,2012,2013) group by year,farmercode order by farmercode"
    SQLSTR = "select substring(farmercode,1,14) dgt,sum(challanqty) challanqty,year from tblplanted where year in(2011,2012,2013) group by year,dgt order by dgt"
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do Until rs.EOF
    
  
    
    frcode = rs!dgt
    FindDZ Mid(rs!dgt, 1, 3)
    FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
    FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
    FindFA Mid(rs!dgt, 1, 14), "F"
    excel_sheet.Cells(i, 1) = Dzname
    excel_sheet.Cells(i, 2) = GEname
    excel_sheet.Cells(i, 3) = TsName
   excel_sheet.Cells(i, 3) = Mid(rs!dgt, 1, 14) & "  " & FAName
'   If rs!dgt = "D01G03T01F0019" Then
'
'   MsgBox "asdas"
'   End If
    Set mn = Nothing
    mn.Open "select sum(regland) regland from tbllandreg where substring(farmerid,1,14)='" & rs!dgt & "' group by substring(farmerid,1,14)", MHVDB
    If mn.EOF <> True Then
    excel_sheet.Cells(i, 11) = mn!regland
    End If
    Do While frcode = rs!dgt
    If rs!Year = 2011 Then
    excel_sheet.Cells(i, 5) = rs!challanqty
    ElseIf rs!Year = 2012 Then
    excel_sheet.Cells(i, 6) = rs!challanqty
    
    ElseIf rs!Year = 2013 Then
    excel_sheet.Cells(i, 7) = rs!challanqty
    End If
    excel_sheet.Cells(i, 8) = Val(excel_sheet.Cells(i, 5)) + Val(excel_sheet.Cells(i, 6)) + Val(excel_sheet.Cells(i, 7))
    
    
    'field
 Set rsF = Nothing
 rsF.Open "select sum(totaltrees) totaltrees,sum(tree_count_deadmissing) deadmissing from  " & Mtblname & " where fstype='F' and substring(farmercode,1,14)='" & rs!dgt & "' group by substring(farmercode,1,14)", ODKDB
 
 If rsF.EOF <> True Then
 excel_sheet.Cells(i, 9) = rsF!totaltrees
 excel_sheet.Cells(i, 10) = rsF!deadmissing

 Else
 
  excel_sheet.Cells(i, 9) = ""
 excel_sheet.Cells(i, 10) = ""
 excel_sheet.Cells(i, 11) = ""
 End If
 
 'storage
 
 Set rsF = Nothing
 rsF.Open "select sum(totaltrees) totaltrees,sum(tree_count_deadmissing) deadmissing from  " & Mtblname & " where fstype='S' and substring(farmercode,1,14)='" & rs!dgt & "' group by substring(farmercode,1,14)", ODKDB
 
 If rsF.EOF <> True Then
 excel_sheet.Cells(i, 12) = rsF!totaltrees
 excel_sheet.Cells(i, 13) = rsF!deadmissing

 Else
 
 excel_sheet.Cells(i, 11) = ""
 excel_sheet.Cells(i, 12) = ""
 excel_sheet.Cells(i, 13) = ""
 End If
 
 'total
 excel_sheet.Cells(i, 15) = Val(excel_sheet.Cells(i, 9)) + Val(excel_sheet.Cells(i, 12))
 If Val(excel_sheet.Cells(i, 15)) > 0 Then
 
 excel_sheet.Cells(i, 16) = Val(excel_sheet.Cells(i, 10)) + Val(excel_sheet.Cells(i, 13))
 
 Else
  excel_sheet.Cells(i, 15) = ""
 excel_sheet.Cells(i, 16) = ""
 excel_sheet.Cells(i, 17) = ""
 
 End If
 
 
   
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
    If Val(excel_sheet.Cells(i, 13)) > 0 Then
     excel_sheet.Cells(i, 14) = 1 - Val(excel_sheet.Cells(i, 13)) / Val(excel_sheet.Cells(i, 8))
     End If
     
     If Val(excel_sheet.Cells(i, 10)) > 0 Then
     excel_sheet.Cells(i, 11) = 1 - Val(excel_sheet.Cells(i, 10)) / Val(excel_sheet.Cells(i, 8))
     End If
      If Val(excel_sheet.Cells(i, 15)) > 0 Then
     excel_sheet.Cells(i, 17) = 1 - Val(excel_sheet.Cells(i, 16)) / Val(excel_sheet.Cells(i, 8))
     End If
     
     
     sl = sl + 1
    i = i + 1
    Loop
    
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    
    
    ODKDB.Execute "drop table " & Mtblname & ""
    ODKDB.Close
    
    
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command9_Click()
Dim frcode As String
Dim i, sl As Integer
Dim cntold, cntnew As Integer
Dim mn As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
frcode = ""
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("Dzongkhag")
    excel_sheet.Cells(3, 2) = ProperCase("gewog")
    excel_sheet.Cells(3, 3) = ProperCase("tshowog")
    'excel_sheet.Cells(3, 4) = ProperCase("farmers")
    
    excel_sheet.Cells(3, 4) = ProperCase("Planted List")
   ' excel_sheet.Cells(3, 5) = ProperCase("New Farmers(Distribution)")
    'excel_sheet.Cells(3, 6) = ProperCase("Old Farmers(Distribution)")
    
    excel_sheet.Cells(3, 7) = ProperCase("2011")
    excel_sheet.Cells(3, 8) = ProperCase("2012")
    excel_sheet.Cells(3, 9) = ProperCase("2013")
        excel_sheet.Cells(3, 10) = ProperCase("2014")
    'excel_sheet.Cells(3, 9) = ProperCase("2013( new farmers )")
   ' excel_sheet.Cells(3, 10) = ProperCase("2013( old farmers )")
    'excel_sheet.Cells(3, 11) = ProperCase("Acre(New)")
    excel_sheet.Cells(3, 11) = ProperCase("Acre")
    'excel_sheet.Cells(3, 12) = ProperCase("Acre")
    
    i = 4
  
   SQLSTR = "" '"select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,,sum(acre) acre,sum(acreold) acreold ,cntplanted,cntdistnew,cntdistold ,year from tbldisttemp where  year in(2011,2012,2013) group by year,farmercode order by farmercode"
   SQLSTR = "SELECT  FARMERID,substring(FARMERID,1,3) dz,substring(FARMERID,4,3) ge,substring(FARMERID,7,3) ts, FARMERID,plantedstatus,sum(REGLAND) regland FROM tbllandreg WHERE status not in('D','R','C') and REGDATE <='2014-10-01' group by plantedstatus,farmerid"
    'SQLSTR = "select substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts, farmercode,sum(challanqty) pplant,year from tblplanted where  year in(2011,2012,2013,2014) group by year,farmercode order by farmercode"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    
  
    
    frcode = rs!farmerid
    FindDZ Mid(rs!farmerid, 1, 3)
    FindGE Mid(rs!farmerid, 1, 3), Mid(rs!farmerid, 4, 3)
    FindTs Mid(rs!farmerid, 1, 3), Mid(rs!farmerid, 4, 3), Mid(rs!farmerid, 7, 3)
    excel_sheet.Cells(i, 1) = Dzname
    excel_sheet.Cells(i, 2) = GEname
    excel_sheet.Cells(i, 3) = TsName
    excel_sheet.Cells(i, 4) = rs!farmerid
    excel_sheet.Cells(i, 7) = rs!plantedstatus
    excel_sheet.Cells(i, 11) = rs!regland
    'Set mn = Nothing
    'mn.Open "select sum(challanqty) pplant from tblplanted where year in(2012,2013) and farmercode='" & rs!farmerid & "'", MHVDB
    'If mn.EOF <> True Then
   ' excel_sheet.Cells(i, 7) = mn!pplant
   ' End If
   
    
    
    
    rs.MoveNext
    sl = sl + 1
    i = i + 1
    Loop
    
    
   
    
    
    End If
  
    
   
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Distribution Report"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Dim tt As String
Dim rs As New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open CnnString
Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct  Year from tblplantdistributiondetail where trnid in(select trnid from tblplantdistributionheader where status='ON' )" _
& "  order by year", db
Set cboyear.RowSource = rs
cboyear.ListField = "Year"
cboyear.BoundColumn = "Year"

Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cboDzongkhag.RowSource = rs
cboDzongkhag.ListField = "dzongkhagname"
cboDzongkhag.BoundColumn = "dzongkhagcode"


End Sub
Private Sub filldz()
Dim totfarmer As Long
Dim rs As New ADODB.Recordset
Dim i As Integer
totfarmer = 0
i = 1
dzgrid.Rows = 1
Set rs = Nothing
If Len(cbomnth.Text) = 0 Then
rs.Open "select mid(farmercode,1,3) dz, mid(farmercode,4,3) ge, mid(farmercode,7,3) ts, farmercode,count(farmercode) cnt,sum((crateno * 35)+plno+crate) plants,sum(area) area from tblplantdistributiondetail where trnid in(select trnid from tblplantdistributionheader where status='ON') and subtotindicator not in ('S','T') and status<>'C' and year='" & cboyear.BoundText & "' and  senttofield <>'N' and farmercode in(select idfarmer from tblfarmer where status='A')group by substring(farmercode,1,3) order by substring(farmercode,1,3)", MHVDB
Else
rs.Open "select mid(farmercode,1,3) dz, mid(farmercode,4,3) ge, mid(farmercode,7,3) ts, farmercode,count(farmercode) cnt,sum((crateno * 35)+plno+crate) plants,sum(area) area from tblplantdistributiondetail where trnid in(select trnid from tblplantdistributionheader where status='ON') and subtotindicator not in ('S','T') and status<>'C' and year='" & cboyear.BoundText & "' and mnth='" & cbomnth.ListIndex + 1 & "' and  senttofield <>'N' and farmercode in(select idfarmer from tblfarmer where status='A')group by substring(farmercode,1,3) order by substring(farmercode,1,3)", MHVDB
End If
If rs.EOF <> True Then
Do While rs.EOF <> True
FindDZ rs!dz
dzgrid.Rows = dzgrid.Rows + 1
dzgrid.TextMatrix(i, 0) = Dzname
dzgrid.TextMatrix(i, 1) = CLng(rs!cnt)
dzgrid.TextMatrix(i, 2) = CLng(rs!plants)
dzgrid.TextMatrix(i, 3) = Format(rs!area, "####0.00")
totfarmer = totfarmer + CLng(rs!plants)
i = i + 1
rs.MoveNext
Loop

Else
MsgBox "Record Not Available"
End If
txtplants.Text = totfarmer

End Sub

Private Sub fillge()
Dim totfarmer As Long
Dim rs As New ADODB.Recordset
Dim i As Integer
totfarmer = 0
i = 1
gegrid.Rows = 1
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz, mid(farmercode,4,3) ge, mid(farmercode,7,3) ts, farmercode,count(farmercode) cnt,sum(totalplant) plants,sum(area) area from tblplantdistributiondetail where trnid in(select trnid from tblplantdistributionheader where status='ON') and subtotindicator not in ('S','T') and status<>'C' and year='" & cboyear.BoundText & "' and substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and senttofield <>'N' and farmercode in(select idfarmer from tblfarmer where status='A')group by substring(farmercode,1,3),substring(farmercode,4,3) order by substring(farmercode,1,3),substring(farmercode,4,3)", MHVDB
If rs.EOF <> True Then
Do While rs.EOF <> True
FindGE rs!dz, rs!ge
gegrid.Rows = gegrid.Rows + 1
gegrid.TextMatrix(i, 0) = GEname
gegrid.TextMatrix(i, 1) = CLng(rs!cnt)
gegrid.TextMatrix(i, 2) = CLng(rs!plants)
gegrid.TextMatrix(i, 3) = Format(rs!area, "####0.00")
totfarmer = totfarmer + CLng(rs!plants)
i = i + 1
rs.MoveNext
Loop

Else
MsgBox "Record Not Available"
End If
txtgeplants.Text = totfarmer

End Sub

Private Sub fillts()
Dim totfarmer As Long
Dim rs As New ADODB.Recordset
Dim i As Integer
totfarmer = 0
i = 1
tsgrid.Rows = 1
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz, mid(farmercode,4,3) ge, mid(farmercode,7,3) ts, farmercode,count(farmercode) cnt,sum(totalplant) plants,sum(area) area from tblplantdistributiondetail where trnid in(select trnid from tblplantdistributionheader where status='ON') and subtotindicator not in ('S','T') and status<>'C' and year='" & cboyear.BoundText & "' and substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' and senttofield <>'N' and farmercode in(select idfarmer from tblfarmer where status='A')group by substring(farmercode,1,3),substring(farmercode,4,3),substring(farmercode,7,3) order by substring(farmercode,1,3),substring(farmercode,4,3),substring(farmercode,7,3)", MHVDB
If rs.EOF <> True Then
Do While rs.EOF <> True
FindTs rs!dz, rs!ge, rs!ts
tsgrid.Rows = tsgrid.Rows + 1
tsgrid.TextMatrix(i, 0) = TsName
tsgrid.TextMatrix(i, 1) = CLng(rs!cnt)
tsgrid.TextMatrix(i, 2) = CLng(rs!plants)
tsgrid.TextMatrix(i, 3) = Format(rs!area, "####0.00")
totfarmer = totfarmer + CLng(rs!plants)
i = i + 1
rs.MoveNext
Loop

Else
MsgBox "Record Not Available"
End If
txttsplants.Text = totfarmer

End Sub

