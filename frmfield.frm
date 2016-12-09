VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmfield 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIELD REPORT"
   ClientHeight    =   6180
   ClientLeft      =   3660
   ClientTop       =   2955
   ClientWidth     =   9210
   Icon            =   "frmfield.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9210
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   120
      TabIndex        =   37
      Top             =   3720
      Width           =   495
   End
   Begin VB.CheckBox CHKMONTHLY 
      Caption         =   "MONTHLY"
      Height          =   195
      Left            =   3960
      TabIndex        =   36
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "DGT SELECTION OPTION"
      Height          =   1095
      Left            =   480
      TabIndex        =   34
      Top             =   4920
      Width           =   4815
      Begin MSDataListLib.DataCombo CBODGT 
         Bindings        =   "frmfield.frx":076A
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
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
   End
   Begin VB.CheckBox CHKRANGE 
      Caption         =   "RANGE"
      Height          =   195
      Left            =   3960
      TabIndex        =   31
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox CHKDGT 
      Caption         =   "DETAIL"
      Height          =   195
      Left            =   3960
      TabIndex        =   26
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "DATE SELECTION"
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OPTSEL 
         Caption         =   "SELECTIVE"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "RECORD SELECTION"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   5055
      Begin VB.OptionButton OPTTOPN 
         Caption         =   "LAST VISIT"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TXTRECORDNO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   21
         Top             =   350
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OPTALLVISIT 
         Caption         =   "ALL VISIT"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "VISIT"
         Height          =   195
         Left            =   3435
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SHOW"
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
      Left            =   480
      Picture         =   "frmfield.frx":077F
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   1575
   End
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
      Left            =   2160
      Picture         =   "frmfield.frx":0EE9
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   5055
      Begin VB.ComboBox CBODATE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmfield.frx":1BB3
         Left            =   1080
         List            =   "frmfield.frx":1BB5
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   79364097
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   79364097
         CurrentDate     =   41362
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MORE OPTION"
      Height          =   6135
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton OPTSLOWGROWING 
         Caption         =   "SLOW GROWING"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtrange2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   29
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtrange1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optmoist 
         Caption         =   "MOISTURE"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optdead 
         Caption         =   "DEAD"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox TXTVALUE 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton optrootpest 
         Caption         =   "ROOT PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton OPTSTEMPEST 
         Caption         =   "STEM PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton OPTLEAFPEST 
         Caption         =   "LEAF PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   1335
         Left            =   120
         TabIndex        =   33
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
         _cx             =   3413
         _cy             =   2355
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
         BackColorAlternate=   16777152
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
         Rows            =   5
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmfield.frx":1BB7
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
         Editable        =   2
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
         Caption         =   "____"
         Height          =   375
         Left            =   3000
         TabIndex        =   30
         Top             =   3720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "PERCENTAGE RANGE"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "PERCENTAGE VALUE GREATER THEN"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   3135
      End
   End
   Begin VB.CheckBox CHKMOREOPTION 
      Caption         =   "MORE OPTION"
      Height          =   195
      Left            =   3960
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   6120
   End
End
Attribute VB_Name = "frmfield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mfieldname As String
Dim mfname As String

Private Sub CHKDGT_Click()
If CHKDGT.Value = 1 Then
Frame4.Visible = True
Else
Frame4.Visible = False
End If
End Sub

Private Sub CHKMONTHLY_Click()
If CHKMONTHLY.Value = 1 Then
OPTSLOWGROWING.Value = True
Frame4.Visible = True
TXTVALUE.Text = 10
mfieldname = "tree_count_slowgrowing"
mfname = "SLOW GROWING"
End If

If CHKDGT.Value = 1 Then
Frame4.Visible = True
End If
End Sub

Private Sub CHKMOREOPTION_Click()
If CHKMOREOPTION.Value = 1 Then
frmfield.Width = 9435
OPTTOPN.Enabled = False
Else
frmfield.Width = 5595
OPTTOPN.Enabled = True
End If
OPTSLOWGROWING.Value = True
TXTVALUE.Text = 10
mfieldname = "tree_count_slowgrowing"
End Sub

Private Sub CHKRANGE_Click()
If CHKRANGE.Value = 1 Then

Mygrid.Visible = True
Else
Mygrid.Visible = False
End If
End Sub

Private Sub Command1_Click()
Dim SQLSTR As String
      SQLSTR = ""
      SQLSTR = "insert into tblfieldlastvisitrpt (START,end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,farmerbarcode,fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS) " _
      & "select START,n.end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,n.farmerbarcode,n.fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode, n.fdcode"
         
         
 ODKDB.Execute "delete from tblfieldlastvisitrpt"
 ODKDB.Execute SQLSTR
 ODKDB.Execute "update tblfieldlastvisitrpt set region_dcode=substring(farmerbarcode,1,3),region_gcode=substring(farmerbarcode,4,3), " _
 & " region=substring(farmerbarcode,7,3)"
MHVDB.Execute "update odk_prodlocal.tblfieldlastvisitrpt a ,tbldzongkhag b set region_dcode=concat(substring(region_dcode,1,3),'  ',DzongkhagName) where substring(region_dcode,1,3)=DzongkhagCode"
MHVDB.Execute "update odk_prodlocal.tblfieldlastvisitrpt a ,tblgewog b set region_gcode=concat(substring(region_gcode,1,3),'  ',GewogName) where concat(substring(region_dcode,1,3),substring(region_gcode,1,3))=concat(DzongkhagId,GewogId)"
MHVDB.Execute "update odk_prodlocal.tblfieldlastvisitrpt a ,tbltshewog b set region=concat(substring(region,1,3),'  ',TshewogName) where concat(substring(region_dcode,1,3),substring(region_gcode,1,3),substring(region,1,3))=concat(DzongkhagId,GewogId,TshewogId)"
MHVDB.Execute "update odk_prodlocal.tblfieldlastvisitrpt a ,tblfarmer b set farmerbarcode=concat(farmerbarcode,'  ',farmername) where farmerbarcode=idfarmer"
MHVDB.Execute "update odk_prodlocal.tblfieldlastvisitrpt a ,tblmhvstaff b set staffbarcode=concat(staffbarcode,'  ',staffname) where staffbarcode=staffcode"


End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
'fieldmonthlyvisit
'Exit Sub
If CHKMONTHLY.Value = 1 Then
RPTMONTHLY
Exit Sub
End If


If CHKRANGE.Value = 0 Then
allfieldsandstorage
'allfields
Else
If mfieldname = "" Then
MsgBox "SELECT ONE FIELD."
Exit Sub
End If

If CHKDGT.Value = 0 Then
allfieldsRANGE
Else
allfieldsRANGEDET
End If



End If



End Sub
Private Sub fieldDaily(receipient_id As Integer, nextemaildate As Date, frequency As Integer)
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsch As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                   
'GetTbl
mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1





SQLSTR = ""



         SQLSTR = "select _URI,start, tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,treeheight,comments1,other2,stems,management,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub6_core "
       SQLSTR = SQLSTR & "where status<>'BAD' and substring(start,1,10)>='" & Format(Now - 10, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(Now - 1, "yyyy-MM-dd") & "' order by cast(substring(staffbarcode,3,3) as unsigned integer)"
               
      

   'On Error Resume Next




Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer
Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "START DATE"
    excel_sheet.Cells(3, 3) = "TDATE"
    excel_sheet.Cells(3, 4) = "END DATE"
    excel_sheet.Cells(3, 5) = "STAFF CODE-NAME"
    excel_sheet.Cells(3, 6) = "DZONGKHAG"
    excel_sheet.Cells(3, 7) = "GEWOG"
    excel_sheet.Cells(3, 8) = "TSHOWOG"
    excel_sheet.Cells(3, 9) = UCase("Farmer ID")
    excel_sheet.Cells(3, 10) = UCase("average tree height")
    excel_sheet.Cells(3, 11) = UCase("Field ID")
    excel_sheet.Cells(3, 12) = UCase("Total Trees Distributed - Planted List")
    excel_sheet.Cells(3, 13) = UCase("Total Trees")
    excel_sheet.Cells(3, 14) = UCase("Good Moisture")
    excel_sheet.Cells(3, 15) = UCase("Poor Moisture")
    excel_sheet.Cells(3, 16) = UCase("Total Mositure Tally")
    excel_sheet.Cells(3, 17) = UCase("Dead Missing")
    excel_sheet.Cells(3, 18) = UCase("Slow Growing")
    excel_sheet.Cells(3, 19) = UCase("Dormant")
    excel_sheet.Cells(3, 20) = UCase("Active Growing")
    excel_sheet.Cells(3, 21) = UCase("Shock")
    excel_sheet.Cells(3, 22) = UCase("Nutrient Deficient")
    excel_sheet.Cells(3, 23) = UCase("WaterLogged")
    excel_sheet.Cells(3, 24) = UCase("average no. of stem")
    excel_sheet.Cells(3, 25) = UCase("Active leaf Pest")
    excel_sheet.Cells(3, 26) = UCase("Stem Pest")
    excel_sheet.Cells(3, 27) = UCase("Root Pest")
    excel_sheet.Cells(3, 28) = UCase("Animal Damage")
    excel_sheet.Cells(3, 29) = UCase("Monitor's comments")
    excel_sheet.Cells(3, 30) = UCase("follow up question")
    excel_sheet.Cells(3, 31) = UCase("management")
    excel_sheet.Cells(3, 32) = UCase("farmer's comments")
   i = 4
  Set rs = Nothing
  
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!Start
excel_sheet.Cells(i, 3) = "'" & rs!tdate
excel_sheet.Cells(i, 4) = "'" & rs!End  'rs.Fields(Mindex)
FindsTAFF rs!staffbarcode
excel_sheet.Cells(i, 5) = rs!staffbarcode & " " & sTAFF

FindDZ Mid(rs!farmerbarcode, 1, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmerbarcode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3)
excel_sheet.Cells(i, 7) = Mid(rs!farmerbarcode, 4, 3) & " " & GEname
FindTs Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3), Mid(rs!farmerbarcode, 7, 3)
excel_sheet.Cells(i, 8) = Mid(rs!farmerbarcode, 7, 3) & " " & TsName
FindFA IIf(IsNull(rs!farmerbarcode), "", rs!farmerbarcode), "F"
excel_sheet.Cells(i, 9) = IIf(IsNull(rs!farmerbarcode), "", rs!farmerbarcode) & " " & FAName


excel_sheet.Cells(i, 10) = IIf(IsNull(rs!treeheight), "", rs!treeheight)
excel_sheet.Cells(i, 11) = IIf(IsNull(rs!FDCODE), "", rs!FDCODE)

Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as nooftrees from tblplanted where farmercode='" & rs!farmerbarcode & "' group by farmercode ", MHVDB
If rs1.EOF <> True Then
excel_sheet.Cells(i, 12) = rs1!nooftrees

Else

excel_sheet.Cells(i, 12) = ""
End If
excel_sheet.Cells(i, 13) = IIf(IsNull(rs!tree_count_totaltrees), 0, rs!tree_count_totaltrees)
excel_sheet.Cells(i, 14) = IIf(IsNull(rs!qc_tally_goodmoisture), "", rs!qc_tally_goodmoisture)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!qc_tally_poormoisture), "", rs!qc_tally_poormoisture)
excel_sheet.Cells(i, 16) = IIf(IsNull(rs!qc_tally_goodmoisture), "", rs!qc_tally_goodmoisture) + IIf(IsNull(rs!qc_tally_poormoisture), "", rs!qc_tally_poormoisture)
excel_sheet.Cells(i, 17) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
excel_sheet.Cells(i, 18) = IIf(IsNull(rs!tree_count_slowgrowing), "", rs!tree_count_slowgrowing)
excel_sheet.Cells(i, 19) = IIf(IsNull(rs!tree_count_dor), "", rs!tree_count_dor)
excel_sheet.Cells(i, 20) = IIf(IsNull(rs!tree_count_activegrowing), "", rs!tree_count_activegrowing)
excel_sheet.Cells(i, 21) = IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.Cells(i, 22) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.Cells(i, 23) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.Cells(i, 24) = IIf(IsNull(rs!stems), "", rs!stems)
excel_sheet.Cells(i, 25) = IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.Cells(i, 26) = IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.Cells(i, 27) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.Cells(i, 28) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.Cells(i, 29) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)
If Mid(IIf(IsNull(rs!management), "", rs!management), 1, 1) = "y" Then
excel_sheet.Cells(i, 30) = "Yes"
Else
excel_sheet.Cells(i, 30) = "No"
End If

If UCase((IIf(IsNull(rs!management), "", rs!management))) = UCase("y") Then
Set rs1 = Nothing
rs1.Open "select * from phealthhub15_management1 where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tblfieldchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
If UCase(rsch!Label) = UCase("description9") Then
actstring = rs!other2
Else
actstring = IIf(IsNull(rsch!Label), "", rsch!Label) & " # " & actstring
End If
End If

rs1.MoveNext
Loop
If Len(actstring) > 0 Then
 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.Cells(i, 31) = actstring
End If
Else

excel_sheet.Cells(i, 31) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
Else
excel_sheet.Cells(i, 31) = ""
End If

excel_sheet.Cells(i, 32) = IIf(IsNull(rs!comments1), "", rs!comments1)
SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up




    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:ag3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


excel_sheet.Columns("A:ag").Select
 excel_app.Selection.ColumnWidth = 15
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
'excel_app.Visible = True


db.Close

'updateemaillog excel_app, receipient_id, nextemaildate, frequency


Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault




Exit Sub
err:
MsgBox err.Description
err.Clear

End Sub

Private Sub fieldmonthlyvisit(receipient_id As Integer, nextemaildate As Date, frequency As Integer)
Dim fdcnt As Integer
Dim excel_app As Object
Dim excel_sheet As Object
Dim Excel_WBook As Object
Dim exactrow As Integer
Dim Excel_Chart As Object
'On Error Resume Next
Dim jrow As Long
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
mchk = True
Dim i, Jmth, K As Integer
Dim j As Double
Dim mtot(1 To 14), jtot As Double
 Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim DT1 As Date
Dim DT2 As Date
intYear = CInt(Year(txtfrmdate.Value))
intMonth = 1
intDay = 1
 DT1 = DateSerial(intYear, intMonth, intDay)
txtfrmdate.Value = Format(DT1, "dd/MM/yyyy")
intYear = CInt(Year(txtfrmdate.Value))
intMonth = 12
intDay = 31
 DT2 = DateSerial(intYear, intMonth, intDay)
txttodate.Value = Format(DT2, "dd/MM/yyyy")



Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
db.Execute "delete from tempfarmernotinfield"
SQLSTR = " insert into tempfarmernotinfield(end,farmercode,fdcode,staffbarcode)" _
& "select '' as end,farmercode,'',monitor from mhv.tblplanted as a , mhv.tblfarmer as b  " _
& "where farmercode=idfarmer and farmercode not in (select farmerbarcode from  phealthhub15_core)"

db.Execute SQLSTR


fdcnt = 0
For i = 1 To 13
    mtot(i) = 0
Next
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_app.Caption = "mhv"
    excel_app.Visible = True
     jrow = 2
      excel_sheet.Cells(jrow, 1) = "FARMER"
      
     
    Set rs1 = Nothing
    rs1.Open "SELECT DISTINCT staffbarcode FROM phealthhub15_core", db
    Do While rs1.EOF <> True
    SQLSTR = ""
    SQLSTR = "select max(END) as end,farmerbarcode,concat(farmerbarcode,cast(fdcode as char))  as id ,fdcode,count(farmerbarcode) as jval,year(end) as procyear,month(end) " _
    & " as procmonth from odk_prodlocal.phealthhub15_core  where  staffbarcode='" & rs1!staffbarcode & "' and end between '2013-01-01' and '2013-12-31' group by year" _
    & " (end),month(end),farmerbarcode,fdcode union SELECT  STR_TO_DATE('2013-01-01 14:15:16', '%d/%m/%Y') as END , farmercode, farmercode AS id,0 AS fdcode, 0 AS jval," _
    & "  year('" & Format(DT1, "yyyy-MM-dd") & "') AS procyear,  month('" & Format(DT1, "yyyy-MM-dd") & "') as procmonth FROM tempfarmernotinfield  WHERE staffbarcode='" & rs1!staffbarcode & "'" _
    & " GROUP BY farmercode ORDER BY farmerbarcode, fdcode, YEAR(END) , MONTH(END) "


Set rs = Nothing
    rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    
    jCol = 5 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs1!staffbarcode
     
       jrow = jrow + 1
    excel_sheet.Cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
    excel_sheet.Cells(jrow, 2) = "LAST VISITED"
   excel_sheet.Cells(jrow, 3) = "FIELD CODE"
     
    K = 3
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(jrow, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.Cells(jrow, jCol) = UCase("Total")
    excel_sheet.Range(excel_sheet.Cells(jrow - 1, 1), _
    excel_sheet.Cells(jrow, 14)).Select
    excel_app.Selection.Font.Bold = True
    jtot = 0
    fdcnt = 0
  
    
     Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       FindFA Trim(rs!farmerbarcode), "F"
       excel_sheet.Cells(jrow, 1) = rs!farmerbarcode & " " & FAName
       jtot = 0
       j = 0
       
       Do While pyear = rs!id
          i = rs!procmonth + 4 - Month(txtfrmdate)
          
          j = IIf(rs!jval = "", 0, rs!jval)
          jtot = jtot + j
         
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, 2) = "'" & rs!End
          excel_sheet.Cells(jrow, 3) = CInt(rs!FDCODE)
          
          'fdcnt = fdcnt + 1
          excel_sheet.Cells(jrow, i) = Val(j)
         pyear = rs!id
          rs.MoveNext
         
          If rs.EOF Then Exit Do
          'jrow = jrow + 1
       Loop
       
     
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
       If Val(jtot) > 0 Then
        fdcnt = fdcnt + 1
       End If
       'rs.MoveNext
       'jtot = 0
    Loop
   jtot = 0
    excel_sheet.Cells(jrow + 1, 3) = fdcnt
    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
    For i = 4 To jCol - 1
        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
      excel_sheet.Range(excel_sheet.Cells(jrow + 1, 1), _
    excel_sheet.Cells(jrow + 1, 16)).Select
    excel_app.Selection.Font.Bold = True
    jtot = 0
    
     For i = 2 To jCol - 1
       mtot(i - 1) = 0
       
    Next
    jrow = jrow + 2
    rs1.MoveNext
    Loop
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(jrow + 1, jCol)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(3, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     'excel_sheet.Range("A1:Aa15").Font.Bold = True
    
   excel_sheet.Name = "Detail"

    Excel_WBook.Sheets("sheet2").Activate
  If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If

excel_sheet.Name = "Summary"
SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from  phealthhub15_core where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"


Set rs = Nothing
rs.Open SQLSTR, OdkCnnString
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs!id
    'excel_sheet.Cells(2, 1) = "MONTHLY ACTIVITY OF MONITOR " & rs!id & " " & sTAFF
    excel_sheet.Cells(3, 1) = "MONITOR"
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.Cells(3, jCol) = UCase("Total")
    excel_sheet.Cells(3, jCol + 1) = ("Detail")
    
    exactrow = 3
    jrow = 3
    Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       FindsTAFF Trim(rs!id)
       excel_sheet.Cells(jrow, 1) = rs!id & " " & sTAFF
       jtot = 0
       Do While pyear = rs!id
          i = rs!procmonth + 2 - Month(txtfrmdate)
          j = rs!jval
          jtot = jtot + j
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, i) = Val(j)
          rs.MoveNext
          If rs.EOF Then Exit Do
          exactrow = exactrow + 1
       Loop
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
       tt = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(P3&" & "!A:A" & "),MATCH(" & "A" & jrow & ",INDIRECT(P3&" & "!A:A" & "),0)))," & "Link" & ")"
       excel_sheet.Cells(jrow, jCol + 1).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
       '
      
       exactrow = exactrow + 1
    Loop
    jtot = 0
    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
    exactrow = exactrow + 1
    For i = 2 To jCol - 1
        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
    exactrow = exactrow + 1
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(jrow + 1, jCol)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    
     excel_sheet.Range("A1:o3").Font.Bold = True
    


'updateemaillog excel_app, receipient_id, nextemaildate, frequency



    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
    
 

End Sub


Private Sub RPTMONTHLYVisit()
Dim excel_app As Object
Dim excel_sheet As Object
Dim Excel_WBook As Object
Dim exactrow As Integer
Dim Excel_Chart As Object
'On Error Resume Next
Dim jrow As Long
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
mchk = True
Dim i, Jmth, K As Integer
Dim j As Double
Dim mtot(1 To 13), jtot As Double
 Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim DT1 As Date
Dim DT2 As Date
intYear = CInt(Year(txtfrmdate.Value))
intMonth = 1
intDay = 1
 DT1 = DateSerial(intYear, intMonth, intDay)
txtfrmdate.Value = Format(DT1, "dd/MM/yyyy")
intYear = CInt(Year(txtfrmdate.Value))
intMonth = 12
intDay = 31
 DT2 = DateSerial(intYear, intMonth, intDay)
txttodate.Value = Format(DT2, "dd/MM/yyyy")



Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
                  
'If Len(cbostaffcode.Text) <> 0 Then
'SQLSTR = "select value as id ,count(value) as jval,year(end) as procyear,month(end) as procmonthfrom dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and staffbarcode='" & cbostaffcode.BoundText & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),value order by convert(substring(value,9,2) ,unsigned integer),year(end),month(end)"
'Else
''SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from dailyacthub9_core  where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"
'End If
For i = 1 To 13
    mtot(i) = 0
Next
    Screen.MousePointer = vbHourglass
    'DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_app.Caption = "mhv"
    excel_app.Visible = False
     jrow = 2
    Set rs1 = Nothing
    rs1.Open "SELECT DISTINCT staffbarcode FROM dailyacthub9_core", db
    Do While rs1.EOF <> True
    SQLSTR = "select value as id ,count(value) as jval,year(end) as procyear,month(end) as procmonth from dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and staffbarcode='" & rs1!staffbarcode & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),value order by convert(substring(value,9,2) ,unsigned integer),year(end),month(end)"
    Set rs = Nothing
    rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs1!staffbarcode
      excel_sheet.Cells(jrow, 1) = "ACTIVITY"
       jrow = jrow + 1
    excel_sheet.Cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
   
  
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(jrow, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.Cells(jrow, jCol) = UCase("Total")
    excel_sheet.Range(excel_sheet.Cells(jrow - 1, 1), _
    excel_sheet.Cells(jrow, 14)).Select
    excel_app.Selection.Font.Bold = True
    jtot = 0
    
    Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       findActivity Trim(rs!id)
       excel_sheet.Cells(jrow, 1) = rs!id & "  :  " & ActName
       jtot = 0
       j = 0
       Do While pyear = rs!id
          i = rs!procmonth + 2 - Month(txtfrmdate)
          j = rs!jval
          jtot = jtot + j
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, i) = Val(j)
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
       'jtot = 0
    Loop
   jtot = 0
    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
    For i = 2 To jCol - 1
        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
      excel_sheet.Range(excel_sheet.Cells(jrow + 1, 1), _
    excel_sheet.Cells(jrow + 1, 14)).Select
    excel_app.Selection.Font.Bold = True
    jtot = 0
    
     For i = 2 To jCol - 1
       mtot(i - 1) = 0
       
    Next
    jrow = jrow + 2
    rs1.MoveNext
    Loop
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(jrow + 1, jCol)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     'excel_sheet.Range("A1:Aa15").Font.Bold = True
    
   excel_sheet.Name = "Detail"

    Excel_WBook.Sheets("sheet2").Activate
  If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If

excel_sheet.Name = "Summary"
SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from  dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"


Set rs = Nothing
rs.Open SQLSTR, OdkCnnString
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs!id
    'excel_sheet.Cells(2, 1) = "MONTHLY ACTIVITY OF MONITOR " & rs!id & " " & sTAFF
    excel_sheet.Cells(3, 1) = "MONITOR"
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.Cells(3, jCol) = UCase("Total")
    excel_sheet.Cells(3, jCol + 1) = ("Detail")
    
    exactrow = 3
    jrow = 3
    Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       FindsTAFF Trim(rs!id)
       excel_sheet.Cells(jrow, 1) = rs!id & " " & sTAFF
       jtot = 0
       Do While pyear = rs!id
          i = rs!procmonth + 2 - Month(txtfrmdate)
          j = rs!jval
          jtot = jtot + j
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, i) = Val(j)
          rs.MoveNext
          If rs.EOF Then Exit Do
          exactrow = exactrow + 1
       Loop
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
       tt = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(P3&" & "!A:A" & "),MATCH(" & "A" & jrow & ",INDIRECT(P3&" & "!A:A" & "),0)))," & "Link" & ")"
       excel_sheet.Cells(jrow, jCol + 1).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
       '
      
       exactrow = exactrow + 1
    Loop
    jtot = 0
    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
    exactrow = exactrow + 1
    For i = 2 To jCol - 1
        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
    exactrow = exactrow + 1
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(jrow + 1, jCol)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A1:o3").Font.Bold = True
    






    

    Screen.MousePointer = vbDefault

db.Close

'updateemaillog excel_app, receipient_id, nextemaildate, frequency

     
Screen.MousePointer = vbDefault
Set excel_sheet = Nothing
Set excel_app = Nothing
End Sub
Private Sub RPTMONTHLY()
Dim excel_app As Object
Dim excel_sheet As Object
Dim locstr As String
Dim muk As New ADODB.Recordset
Dim Excel_WBook As Object
Dim Excel_Chart As Object
'On Error Resume Next
Dim jrow As Long
Dim rs As New ADODB.Recordset
Dim comp As New ADODB.Recordset
mchk = True
Dim i, Jmth, K As Integer
Dim j As Double
Dim mtot(1 To 13), jtot As Double
 Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim DT1 As Date
Dim rng As Range
Dim DT2 As Date
intYear = CInt(Year(txtfrmdate.Value))
intMonth = 1
intDay = 1
 DT1 = DateSerial(intYear, intMonth, intDay)
txtfrmdate.Value = Format(DT1, "dd/MM/yyyy")
intYear = CInt(Year(txtfrmdate.Value))
intMonth = 12
intDay = 31
 DT2 = DateSerial(intYear, intMonth, intDay)
txttodate.Value = Format(DT2, "dd/MM/yyyy")



Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
              GetTbl
                  
SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  year(end),month(end) ,farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY year(end),month(end) , n.farmerbarcode, n.fdcode"
         
    db.Execute SQLSTR
    
    GetTbl1
   SQLSTR = ""
SQLSTR = "insert into " & Mtblname1 & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY  n.farmerbarcode, n.fdcode"
         
    db.Execute SQLSTR
If Len(CBODGT.Text) = 0 Then
                SQLSTR = "select COUNT(fdcode) as fcnt, SUBSTRING(farmercode,1,9) as id ,SUM(totaltrees) as tt,sum(" & mfieldname & ") as jval,(sum(" & mfieldname & ")/sum(totaltrees)*100) as percent,year(end) as procyear,month(end) as procmonth  from " & Mtblname & "   where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),substring(farmercode,1,9) order by substring(farmercode,1,9),year(end),month(end)"
        Else
                SQLSTR = "select COUNT(fdcode) as fcnt,SUBSTRING(farmercode,1,9) as id ,SUM(totaltrees) as tt,sum(" & mfieldname & ") as jval,(sum(" & mfieldname & ")/sum(totaltrees)*100) as percent,year(end) as procyear,month(end) as procmonth  from " & Mtblname & "   where substring(farmercode,1,9)='" & CBODGT.BoundText & "'  and  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),substring(farmercode,1,9) order by substring(farmercode,1,9),year(end),month(end)"
End If





For i = 1 To 13
    mtot(i) = 0
Next

rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    Screen.MousePointer = vbHourglass
    'DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_app.Caption = "mhv"
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    'FindsTAFF "S0" & rs!MM
    excel_sheet.Cells(2, 1) = "MONTHLY SUMMARY OF " & mfname
    
    excel_sheet.Cells(3, 1) = "DZONGKHAG  GEWOG  TSHOWOG"
     excel_sheet.Cells(3, 2) = "TOTAL TREES"
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 2
        excel_sheet.Cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
        K = K + 1
    Next
    'excel_sheet.Cells(3, jCol) = UCase("Total")
    
    
    
    Dim ll As Integer
    jrow = 4
    For ll = 3 To 38
    excel_sheet.Cells(jrow, ll) = "NO.OF FIELDS"
    excel_sheet.Cells(jrow, ll + 1) = "TREE NOS."
    
    excel_sheet.Cells(jrow, ll + 2) = mfname
    ll = ll + 2
    Next
       
    Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       locstr = ""
       FindDZ Mid(rs!id, 1, 3)
       FindGE Mid(rs!id, 1, 3), Mid(rs!id, 4, 3)
       FindTs Mid(rs!id, 1, 3), Mid(rs!id, 4, 3), Mid(rs!id, 7, 3)
       locstr = Mid(rs!id, 1, 3) & " " & Dzname & "  " & Mid(rs!id, 4, 3) & " " & GEname & " " & Mid(rs!id, 7, 3) & " " & TsName
       excel_sheet.Cells(jrow, 1) = locstr
       'jtot = 0
       Do While pyear = rs!id
       If rs!procmonth = 1 Then
        i = 3
       ElseIf rs!procmonth = 2 Then
       i = 6
        ElseIf rs!procmonth = 3 Then
         i = 9
         ElseIf rs!procmonth = 4 Then
          i = 12
          ElseIf rs!procmonth = 5 Then
           i = 15
           ElseIf rs!procmonth = 6 Then
            i = 18
            ElseIf rs!procmonth = 7 Then
             i = 21
             ElseIf rs!procmonth = 8 Then
              i = 24
              ElseIf rs!procmonth = 9 Then
               i = 27
               ElseIf rs!procmonth = 10 Then
                i = 30
                ElseIf rs!procmonth = 11 Then
                 i = 33
                Else
                 i = 36
                End If
'          i = 1
'       i = i * 3
Set muk = Nothing

'    If Mid(rs!id, 1, 6) = "D06G01" Then
'    MsgBox "UMMM"
'    End If
    muk.Open "select sum(totaltrees) as ttrees from " & Mtblname1 & " where substring(farmercode,1,9)='" & rs!id & "'", ODKDB
'i = rs!procmonth + 3 - Month(txtfrmdate)
          j = rs!jval
          'jtot = jtot + j
          'mtot(i - 1) = mtot(i - 1) + j
           excel_sheet.Cells(jrow, 2) = muk!ttrees
           jtot = jtot + muk!ttrees
          excel_sheet.Cells(jrow, i) = rs!FCNT
          excel_sheet.Cells(jrow, i + 1) = rs!tt
           
           excel_sheet.Cells(jrow, i + 2) = rs!jval

          
          
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
      ' excel_sheet.Cells(jrow, jCol) = Val(jtot)
    Loop
    'jtot = 0
'    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
'    For i = 2 To jCol - 1
'        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
'        jtot = jtot + mtot(i - 1)
'    Next
'    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
'MsgBox jtot
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(jrow + 1, jCol)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A1:al4").Font.Bold = True
    
    If Len(CBODGT.Text) <> 0 Then
     Set Excel_Chart = excel_app.Charts.Add 'Before:=Worksheets(Worksheets.Count)

    With excel_app.Charts("Chart1")
  
    '.chartType = 1
    .SeriesCollection.Add _
    Source:=excel_sheet.Range(excel_sheet.Cells(4, 4), _
        excel_sheet.Cells(jrow - 2, jCol - 2))
    '.PlotBy = xlRows
    .HasDataTable = True
    .DataTable.HasBorderOutline = True
    .DataTable.Font.Size = 10
   ' .DataTable.Font.Bold = True
    .HasTitle = True
    .ChartTitle.Text = Me.Caption
    .ChartTitle.Font.Size = 12
    '.ChartTitle.Font.Bold = True
     .Axes(xlCategory, xlPrimary).HasTitle = False
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "NO OF " & mfname
     .PageSetup.LeftFooter = "MHV"
     .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
     .PrintPreview
    End With
    
    End If
     excel_app.Visible = True
    
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
End Sub
Private Sub allfieldsRANGEDET()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim fcriteria As Integer
Dim mdgt, locstring As String
Dim m, n As Integer
Dim t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16 As Integer
Dim TOTTREES, MFLD, FCNT As Double
Dim MCOL As Integer
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                        
GetTbl


      SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode, n.fdcode"
    

db.Execute SQLSTR
On Error Resume Next




'Dim RS As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer

Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    
    
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "FARMER NAME"
    MCOL = 6
    For m = 1 To Mygrid.Rows - 1
    If Len(Mygrid.TextMatrix(m, 0)) = 0 Then Exit For
    If Len(Mygrid.TextMatrix(m, 1)) <> 0 Then
    excel_sheet.Cells(3, MCOL) = Mygrid.TextMatrix(m, 0) & " % To " & Mygrid.TextMatrix(m, 1) & " %"
    If MCOL = 6 Or MCOL = 9 Or MCOL = 12 Then
    excel_sheet.Range(excel_sheet.Cells(3, MCOL), _
                             excel_sheet.Cells(3, MCOL + 2)).Select
                             
                            
                            
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            'excel_sheet.Cells(3, MCOL) = Mygrid.TextMatrix(m, 0) & " To " & Mygrid.TextMatrix(m, 1)
    
    End If

    
    
    Else
      excel_sheet.Range(excel_sheet.Cells(3, MCOL), _
                             excel_sheet.Cells(3, MCOL + 2)).Select
                             
                            
                            
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.Cells(3, MCOL) = ">= Then " & Mygrid.TextMatrix(m, 0) & " %"
     'excel_sheet.Cells(3, MCOL) = "> Then" & Mygrid.TextMatrix(m, 0)
    End If
    MCOL = MCOL + 3
    Next
  
    
   i = 4
   MCOL = 5
   excel_sheet.Cells(i, 6) = "TOTAL TREES"
   excel_sheet.Cells(i, 7) = mfname
   excel_sheet.Cells(i, 8) = "FIELD COUNT"
   
   
   excel_sheet.Cells(i, 9) = "TOTAL TREES"
   excel_sheet.Cells(i, 10) = mfname
   excel_sheet.Cells(i, 11) = "FIELD COUNT"
   
   excel_sheet.Cells(i, 12) = "TOTAL TREES"
   excel_sheet.Cells(i, 13) = mfname
   excel_sheet.Cells(i, 14) = "FIELD COUNT"
   
    excel_sheet.Cells(i, 15) = "TOTAL TREES"
   excel_sheet.Cells(i, 16) = mfname
   excel_sheet.Cells(i, 17) = "FIELD COUNT"
  i = 5
   n = i
   SLNO = 1
   
   MCOL = 5
Set rs = Nothing
  If Len(CBODGT.Text) = 0 Then
   rs.Open "select (farmercode) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(" & mfieldname & ") as mfieldname,count(fdcode) as cnt from " & Mtblname & "  group by farmercode order by farmercode", db
   Else
    rs.Open "select (farmercode) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(" & mfieldname & ") as mfieldname,count(fdcode) as cnt from " & Mtblname & "  where substring(farmercode,1,9)='" & CBODGT.BoundText & "' group by farmercode order by farmercode", db
   End If

  Dim md, mg, mt As String
  Do Until rs.EOF
  mdgt = Mid(rs!dgt, 1, 9)
  md = Mid(rs!dgt, 1, 3)
   mg = Mid(rs!dgt, 4, 3)
    mt = Mid(rs!dgt, 7, 3)
     mchk = True
 FindDZ Mid(rs!dgt, 1, 3)
 FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
 FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
 
  TOTTREES = 0
  MFLD = 0
  FCNT = 0
  t5 = 0
  t6 = 0
  t7 = 0
  t8 = 0
  t9 = 0
  t10 = 0
  t11 = 0
  t12 = 0
  t13 = 0
  t14 = 0
  t15 = 0
  t16 = 0
Do While mdgt = Mid(rs!dgt, 1, 9)
'fcriteria here
'    excel_sheet.Cells(i, 1) = SLNO
'    excel_sheet.Cells(i, 2) = Mid(rs!dgt, 1, 3)
'    excel_sheet.Cells(i, 3) = Mid(rs!dgt, 4, 3)
'    excel_sheet.Cells(i, 4) = Mid(rs!dgt, 7, 3)
FindFA rs!farmercode, "F"
    excel_sheet.Cells(i, 5) = rs!farmercode & " " & FAName
    
    fcriteria = (rs!mfieldname / rs!totaltrees) * 100
    
    If fcriteria >= 0 And fcriteria <= 10 Then
    MCOL = 6
     t5 = t5 + rs!totaltrees
      t6 = t6 + rs!mfieldname
      t7 = t7 + rs!cnt
    ElseIf fcriteria >= 11 And fcriteria <= 20 Then
     t8 = t8 + rs!totaltrees
      t9 = t9 + rs!mfieldname
      t10 = t10 + rs!cnt
    MCOL = 9
    ElseIf fcriteria >= 21 And fcriteria <= 50 Then
      t11 = t11 + rs!totaltrees
      t12 = t12 + rs!mfieldname
      t13 = t13 + rs!cnt
    MCOL = 12
    Else
       t14 = t14 + rs!totaltrees
      t15 = t15 + rs!mfieldname
      t16 = t16 + rs!cnt
    MCOL = 15
    End If
    
   
     
      
      

    excel_sheet.Cells(i, MCOL) = rs!totaltrees

    excel_sheet.Cells(i, MCOL + 1) = rs!mfieldname


    excel_sheet.Cells(i, MCOL + 2) = rs!cnt
    
     
    
    
    
    
    
    
   
    
If Len(CBODGT.Text) <> 0 Then
SLNO = SLNO + 1
End If
i = i + 1
mdgt = Mid(rs!dgt, 1, 9)
If rs.EOF Then Exit Do
rs.MoveNext

Loop
locstring = ""
'If Len(CBODGT.Text) = 0 Then
'excel_sheet.Cells(n, 1) = SLNO

'    excel_sheet.Cells(i, 2) = md & " " & Dzname
'    excel_sheet.Cells(i, 3) = mg & " " & GEname
'    excel_sheet.Cells(i, 4) = mt & " " & TsName
'    excel_sheet.Cells(i, 5) = mt & " " & TsName
locstring = md & " " & Dzname & " " & mg & " " & GEname & " " & mt & " " & TsName
    excel_sheet.Cells(i, 6) = IIf(t5 <> 0, t5, "")
    excel_sheet.Cells(i, 7) = IIf(t6 <> 0, t6, "")
    excel_sheet.Cells(i, 8) = IIf(t7 <> 0, t7, "")
    excel_sheet.Cells(i, 9) = IIf(t8 <> 0, t8, "")
    excel_sheet.Cells(i, 10) = IIf(t9 <> 0, t9, "")
    excel_sheet.Cells(i, 11) = IIf(t10 <> 0, t10, "")
    excel_sheet.Cells(i, 12) = IIf(t11 <> 0, t11, "")
    excel_sheet.Cells(i, 13) = IIf(t12 <> 0, t12, "")
    excel_sheet.Cells(i, 14) = IIf(t13 <> 0, t13, "")
    excel_sheet.Cells(i, 15) = IIf(t14 <> 0, t14, "")
    excel_sheet.Cells(i, 16) = IIf(t15 <> 0, t15, "")
    excel_sheet.Cells(i, 17) = IIf(t16 <> 0, t16, "")
    
    
    
                             excel_sheet.Range(excel_sheet.Cells(n, 1), _
                             excel_sheet.Cells(i - 1, 1)).Select
                                With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            
                            
                            excel_sheet.Cells(n, 1) = SLNO
    
    
    excel_sheet.Range(excel_sheet.Cells(n, 2), _
                             excel_sheet.Cells(i - 1, 4)).Select
                                With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            
                            
                            excel_sheet.Cells(n, 2) = locstring
                            
                            
   excel_sheet.Range(excel_sheet.Cells(i, 2), _
                             excel_sheet.Cells(i, 4)).Select
                                With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            
                            
                            excel_sheet.Cells(i, 2) = "TOTAL"
    
     excel_sheet.Range(excel_sheet.Cells(i, 2), _
 excel_sheet.Cells(i, 17)).Select


                            excel_app.Selection.Font.Bold = True
    
    
    
SLNO = SLNO + 1




'SLNO = SLNO + 1

i = i + 1
n = i
   Loop
  'i = i + 1
'excel_sheet.Cells(i, 5) = "TOTAL"
'
'excel_sheet.Cells(i, 6).Formula = "=SUM(f" & 5 & ":f" & i - 1 & ")"
'excel_sheet.Cells(i, 7).Formula = "=SUM(g" & 5 & ":g" & i - 1 & ")"
'excel_sheet.Cells(i, 8).Formula = "=SUM(h" & 5 & ":h" & i - 1 & ")"
'excel_sheet.Cells(i, 9).Formula = "=SUM(i" & 5 & ":i" & i - 1 & ")"
'excel_sheet.Cells(i, 10).Formula = "=SUM(j" & 5 & ":j" & i - 1 & ")"
'excel_sheet.Cells(i, 11).Formula = "=SUM(k" & 5 & ":k" & i - 1 & ")"
'excel_sheet.Cells(i, 12).Formula = "=SUM(l" & 5 & ":l" & i - 1 & ")"
'excel_sheet.Cells(i, 13).Formula = "=SUM(m" & 5 & ":m" & i - 1 & ")"
'excel_sheet.Cells(i, 14).Formula = "=SUM(n" & 5 & ":n" & i - 1 & ")"
'excel_sheet.Cells(i, 15).Formula = "=SUM(o" & 5 & ":o" & i - 1 & ")"
'excel_sheet.Cells(i, 16).Formula = "=SUM(p" & 5 & ":p" & i - 1 & ")"
'excel_sheet.Cells(i, 17).Formula = "=SUM(q" & 5 & ":q" & i - 1 & ")"
'
' excel_sheet.Range(excel_sheet.Cells(i, 4), _
' excel_sheet.Cells(i, 17)).Select
'
'
'                             excel_app.Selection.Font.Bold = True
   'make up




    excel_sheet.Cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:q4").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


excel_sheet.Columns("A:aa").Select
 excel_app.Selection.ColumnWidth = 15
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
db.Execute "drop table " & Mtblname & ""
db.Close
Exit Sub
err:
db.Execute "drop table " & Mtblname & ""
MsgBox err.Description
err.Clear

End Sub
Private Sub allfieldsRANGE()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim fcriteria As Integer
Dim mdgt As String
Dim m, n As Integer
Dim t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16 As Integer
Dim TOTTREES, MFLD, FCNT As Double
Dim MCOL As Integer
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                        
GetTbl


SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode, n.fdcode"
    

db.Execute SQLSTR
On Error Resume Next




'Dim RS As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer

Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    
    
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
     excel_sheet.Cells(3, 5) = "TOTAL TREES"
    MCOL = 6
    For m = 1 To Mygrid.Rows - 1
    If Len(Mygrid.TextMatrix(m, 0)) = 0 Then Exit For
    If Len(Mygrid.TextMatrix(m, 1)) <> 0 Then
    excel_sheet.Cells(3, MCOL) = Mygrid.TextMatrix(m, 0) & " % To " & Mygrid.TextMatrix(m, 1) & " %"
    If MCOL = 6 Or MCOL = 9 Or MCOL = 12 Then
    excel_sheet.Range(excel_sheet.Cells(3, MCOL), _
                             excel_sheet.Cells(3, MCOL + 2)).Select
                             
                            
                            
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            'excel_sheet.Cells(3, MCOL) = Mygrid.TextMatrix(m, 0) & " To " & Mygrid.TextMatrix(m, 1)
    
    End If

    
    
    Else
      excel_sheet.Range(excel_sheet.Cells(3, MCOL), _
                             excel_sheet.Cells(3, MCOL + 2)).Select
                             
                            
                            
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.Cells(3, MCOL) = ">= Then " & Mygrid.TextMatrix(m, 0) & " %"
     'excel_sheet.Cells(3, MCOL) = "> Then" & Mygrid.TextMatrix(m, 0)
    End If
    MCOL = MCOL + 3
    Next
  
    
   i = 4
   MCOL = 6
   excel_sheet.Cells(i, 6) = "NO. OF  TREES"
   excel_sheet.Cells(i, 7) = mfname
   excel_sheet.Cells(i, 8) = "FIELD COUNT"
   
   
   excel_sheet.Cells(i, 9) = "NO. OF  TREES"
   excel_sheet.Cells(i, 10) = mfname
   excel_sheet.Cells(i, 11) = "FIELD COUNT"
   
   excel_sheet.Cells(i, 12) = "NO. OF  TREES"
   excel_sheet.Cells(i, 13) = mfname
   excel_sheet.Cells(i, 14) = "FIELD COUNT"
   
    excel_sheet.Cells(i, 15) = "NO. OF  TREES"
   excel_sheet.Cells(i, 16) = mfname
   excel_sheet.Cells(i, 17) = "FIELD COUNT"
  i = 5
   n = i
   SLNO = 1
   
   MCOL = 5
Set rs = Nothing
  
   rs.Open "select substring(farmercode,1,9) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(" & mfieldname & ") as mfieldname,count(fdcode) as cnt from " & Mtblname & "  group by farmercode order by farmercode", db

  Dim md, mg, mt As String
  Do Until rs.EOF
  mdgt = rs!dgt
  md = Mid(rs!dgt, 1, 3)
   mg = Mid(rs!dgt, 4, 3)
    mt = Mid(rs!dgt, 7, 3)
     mchk = True
 FindDZ Mid(rs!dgt, 1, 3)
 FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
 FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
  TOTTREES = 0
  MFLD = 0
  FCNT = 0
  t5 = 0
  t6 = 0
  t7 = 0
  t8 = 0
  t9 = 0
  t10 = 0
  t11 = 0
  t12 = 0
  t13 = 0
  t14 = 0
  t15 = 0
  t16 = 0
Do While mdgt = rs!dgt
'fcriteria here
'    excel_sheet.Cells(i, 1) = SLNO
'    excel_sheet.Cells(i, 2) = Mid(rs!dgt, 1, 3)
'    excel_sheet.Cells(i, 3) = Mid(rs!dgt, 4, 3)
'    excel_sheet.Cells(i, 4) = Mid(rs!dgt, 7, 3)
    
    
    fcriteria = (rs!mfieldname / rs!totaltrees) * 100
    
    If fcriteria >= Val(Mygrid.TextMatrix(0, 1)) And fcriteria <= Val(Mygrid.TextMatrix(1, 1)) Then
    MCOL = 5
     t5 = t5 + rs!totaltrees
      t6 = t6 + rs!mfieldname
      t7 = t7 + rs!cnt
    ElseIf fcriteria >= Val(Mygrid.TextMatrix(0, 2)) And fcriteria <= Val(Mygrid.TextMatrix(1, 2)) Then
     t8 = t8 + rs!totaltrees
      t9 = t9 + rs!mfieldname
      t10 = t10 + rs!cnt
    MCOL = 8
    ElseIf fcriteria >= Val(Mygrid.TextMatrix(0, 3)) And fcriteria <= Val(Mygrid.TextMatrix(1, 3)) Then
      t11 = t11 + rs!totaltrees
      t12 = t12 + rs!mfieldname
      t13 = t13 + rs!cnt
    MCOL = 11
    Else
       t14 = t14 + rs!totaltrees
      t15 = t15 + rs!mfieldname
      t16 = t16 + rs!cnt
    MCOL = 14
    End If
    
   
     
      
      
'
'    excel_sheet.Cells(i, MCOL) = rs!totaltrees
'
'    excel_sheet.Cells(i, MCOL + 1) = rs!mfieldname
'
'
'    excel_sheet.Cells(i, MCOL + 2) = rs!cnt
    
     
    
    
    
    
    
    
   
    
    
'SLNO = SLNO + 1
'i = i + 1
mdgt = rs!dgt
If rs.EOF Then Exit Do
rs.MoveNext

Loop
 excel_sheet.Cells(i, 1) = SLNO

    excel_sheet.Cells(i, 2) = md & " " & Dzname
    excel_sheet.Cells(i, 3) = mg & " " & GEname
    excel_sheet.Cells(i, 4) = mt & " " & TsName
    excel_sheet.Cells(i, 5) = IIf(t5 + t8 + t11 + t14 <> 0, t5 + t8 + t11 + t14, "")
    excel_sheet.Cells(i, 6) = IIf(t5 <> 0, t5, "")
    excel_sheet.Cells(i, 7) = IIf(t6 <> 0, t6, "")
    excel_sheet.Cells(i, 8) = IIf(t7 <> 0, t7, "")
    excel_sheet.Cells(i, 9) = IIf(t8 <> 0, t8, "")
    excel_sheet.Cells(i, 10) = IIf(t9 <> 0, t9, "")
    excel_sheet.Cells(i, 11) = IIf(t10 <> 0, t10, "")
    excel_sheet.Cells(i, 12) = IIf(t11 <> 0, t11, "")
    excel_sheet.Cells(i, 13) = IIf(t12 <> 0, t12, "")
    excel_sheet.Cells(i, 14) = IIf(t13 <> 0, t13, "")
    excel_sheet.Cells(i, 15) = IIf(t14 <> 0, t14, "")
    excel_sheet.Cells(i, 16) = IIf(t15 <> 0, t15, "")
    excel_sheet.Cells(i, 17) = IIf(t16 <> 0, t16, "")
 
 
'excel_sheet.Range(excel_sheet.Cells(m, 5),
'excel_sheet.Cells(i - 1, 17)).Select
'excel_sheet.FormulaR1C1 = "=SUM(R[-" & m & "]C:R[-1]C)"     '"=SUM(e" & m & ":q" & i & ")"


SLNO = SLNO + 1
i = i + 1

   Loop
  
excel_sheet.Cells(i, 4) = "TOTAL"
excel_sheet.Cells(i, 5).Formula = "=SUM(E" & 5 & ":E" & i - 1 & ")"
excel_sheet.Cells(i, 6).Formula = "=SUM(F" & 5 & ":F" & i - 1 & ")"
excel_sheet.Cells(i, 7).Formula = "=SUM(G" & 5 & ":G" & i - 1 & ")"
excel_sheet.Cells(i, 8).Formula = "=SUM(H" & 5 & ":H" & i - 1 & ")"
excel_sheet.Cells(i, 9).Formula = "=SUM(I" & 5 & ":I" & i - 1 & ")"
excel_sheet.Cells(i, 10).Formula = "=SUM(J" & 5 & ":J" & i - 1 & ")"
excel_sheet.Cells(i, 11).Formula = "=SUM(K" & 5 & ":K" & i - 1 & ")"
excel_sheet.Cells(i, 12).Formula = "=SUM(L" & 5 & ":L" & i - 1 & ")"
excel_sheet.Cells(i, 13).Formula = "=SUM(M" & 5 & ":M" & i - 1 & ")"
excel_sheet.Cells(i, 14).Formula = "=SUM(N" & 5 & ":N" & i - 1 & ")"
excel_sheet.Cells(i, 15).Formula = "=SUM(O" & 5 & ":O" & i - 1 & ")"
excel_sheet.Cells(i, 16).Formula = "=SUM(P" & 5 & ":P" & i - 1 & ")"
excel_sheet.Cells(i, 17).Formula = "=SUM(Q" & 5 & ":Q" & i - 1 & ")"

 excel_sheet.Range(excel_sheet.Cells(i, 4), _
 excel_sheet.Cells(i, 16)).Select
                             
                          
                             excel_app.Selection.Font.Bold = True
   'make up




    excel_sheet.Cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:Q4").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


excel_sheet.Columns("A:aa").Select
 excel_app.Selection.ColumnWidth = 15
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
db.Execute "drop table " & Mtblname & ""
db.Close
Exit Sub
err:
db.Execute "drop table " & Mtblname & ""
MsgBox err.Description
err.Clear


End Sub
Private Sub allfields()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                   
GetTbl
If OPTALL.Value = True Then
Mindex = 51
End If
mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1





SQLSTR = ""


If OPTTOPN.Value = False Then
         SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,start, end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core "
         
                If OPTALL.Value = True Then
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD'"
                SQLSTR = SQLSTR & "where   status<>'BAD'"
                Else
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                 SQLSTR = SQLSTR & "where status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by cast(substring(staffbarcode,3,3) as unsigned integer)"
                End If
      
Else
      SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,start,n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode, n.fdcode"
    
End If
         
  
  db.Execute SQLSTR
  SQLSTR = ""
  SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select '1900-01-01','1900-01-01','1900-01-01','',substring(farmercode,1,3)," _
         & "substring(farmercode,4,3),substring(farmercode,7,3),farmercode,0,0,0,0,0,0," _
         & "0,0,0,0,0,0,0,0,0,0," _
         & "0,0,0 from mhv.tblplanted where farmercode not in(select farmerbarcode from phealthhub15_core)"
  
  
  
   db.Execute SQLSTR
  
  Set rss = Nothing
  SQLSTR = ""

         
SQLSTR = ""
If CHKMOREOPTION.Value = 0 Then
   SQLSTR = "select * from " & Mtblname & " "
Else
    If optmoist.Value = True Then
    
        SQLSTR = "select * from " & Mtblname & " where (poormoisture/totaltally)*100>'" & Val(TXTVALUE.Text) & "'"
      
    ElseIf optrootpest.Value = True Then
    If CHKRANGE.Value = 0 Then
        SQLSTR = "select * from " & Mtblname & " where (rootpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    
    SQLSTR = "select * from " & Mtblname & " where (rootpest/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (rootpest/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
    End If
    
    ElseIf OPTSTEMPEST.Value = True Then
     If CHKRANGE.Value = 0 Then
        SQLSTR = "select * from " & Mtblname & " where (stempest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    SQLSTR = "select * from " & Mtblname & " where (stempest/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (stempest/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
        
    End If
    
    ElseIf OPTLEAFPEST.Value = True Then
    If CHKRANGE.Value = 0 Then
    
        SQLSTR = "select * from " & Mtblname & " where (leafpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    SQLSTR = "select * from " & Mtblname & " where (leafpest/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (leafpest/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
    
    End If
    
    ElseIf optdead.Value = True Then
    If CHKRANGE.Value = 0 Then
    
        SQLSTR = "select * from " & Mtblname & " where (tree_count_deadmissing/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
        Else
        SQLSTR = "select * from " & Mtblname & " where (tree_count_deadmissing/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (deadmissing/totaltrees)*100<='" & Val(txtrange1.Text) & "'"
        End If
        
        Else
        
        If CHKRANGE.Value = 0 Then
        SQLSTR = "select * from " & Mtblname & " where (tree_count_slowgrowing/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    
    SQLSTR = "select * from " & Mtblname & " where (slowgrowing/totaltrees)*100>='" & Val(txtrange1.Text) & "'  and (deadmissing/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
    End If
    
    End If

End If


SQLSTR = SQLSTR & " order by id,end,farmercode,fdcode"

'If CHKDGT.Value = 1 Then
'SQLSTR = ""
'SQLSTR = "select '' as end,id,substring(farmercode,1,9) as farmercode,sum(treesreceived) as treesreceived ,'' as fdcode,sum" _
'& "(totaltrees) totaltrees,sum(goodmoisture) as goodmoisture,sum(poormoisture) as poormoisture,sum(totaltally) as" _
'& "totaltally,sum(deadmissing) as deadmissing,sum(slowgrowing)as slowgrowing,sum(dor) dor,sum(activegrowing) as" _
'& "activegrowing,sum(shock)as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(leafpest) as leafpest,sum" _
'& "(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage,'' as monitorcomments " _
'& "from " & Mtblname & " group by substring(farmercode,1,9) order by substring(farmercode,1,9)"
'
'End If


'On Error Resume Next




'Dim RS As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer
Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("START DATE")
    excel_sheet.Cells(3, 3) = ProperCase("END DATE")
    excel_sheet.Cells(3, 4) = ProperCase("STAFF CODE-NAME")
    excel_sheet.Cells(3, 5) = ProperCase("DZONGKHAG")
    excel_sheet.Cells(3, 6) = ProperCase("GEWOG")
    excel_sheet.Cells(3, 7) = ProperCase("TSHOWOG")
    excel_sheet.Cells(3, 8) = ProperCase("Farmer ID")
    excel_sheet.Cells(3, 9) = ProperCase("Total Distributed")
    excel_sheet.Cells(3, 10) = ProperCase("Field ID")
    excel_sheet.Cells(3, 11) = ProperCase("Total Trees Distributed - Planted List")
    excel_sheet.Cells(3, 12) = ProperCase("Total Trees")
    excel_sheet.Cells(3, 13) = ProperCase("Good Moisture")
    excel_sheet.Cells(3, 14) = ProperCase("Poor Moisture")
    excel_sheet.Cells(3, 15) = ProperCase("Total Mositure Tally")
    excel_sheet.Cells(3, 16) = ProperCase("Dead Missing")
    excel_sheet.Cells(3, 17) = ProperCase("Slow Growing")
    excel_sheet.Cells(3, 18) = ProperCase("Dormant")
    excel_sheet.Cells(3, 19) = ProperCase("Active Growing")
    excel_sheet.Cells(3, 20) = ProperCase("Shock")
    excel_sheet.Cells(3, 21) = ProperCase("Nutrient Deficient")
    excel_sheet.Cells(3, 22) = ProperCase("Water Logg")
    excel_sheet.Cells(3, 23) = ProperCase("Leaf Pest")
    excel_sheet.Cells(3, 24) = ProperCase("Active Pest")
    excel_sheet.Cells(3, 25) = ProperCase("Stem Pest")
    excel_sheet.Cells(3, 26) = ProperCase("Root Pest")
    excel_sheet.Cells(3, 27) = ProperCase("Animal Damage")
    excel_sheet.Cells(3, 28) = ProperCase("comments")
   i = 4
  Set rs = Nothing
  
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!Start
excel_sheet.Cells(i, 3) = "'" & rs!End  'rs.Fields(Mindex)
FindsTAFF rs!id
excel_sheet.Cells(i, 4) = rs!id & " " & sTAFF

FindDZ Mid(rs!farmercode, 1, 3)
excel_sheet.Cells(i, 5) = Mid(rs!farmercode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmercode, 4, 3) & " " & GEname
FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
excel_sheet.Cells(i, 7) = Mid(rs!farmercode, 7, 3) & " " & TsName
FindFA IIf(IsNull(rs!farmercode), "", rs!farmercode), "F"
excel_sheet.Cells(i, 8) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & " " & FAName


excel_sheet.Cells(i, 9) = IIf(IsNull(rs!treesreceived), "", rs!treesreceived)
excel_sheet.Cells(i, 10) = IIf(IsNull(rs!FDCODE), "", rs!FDCODE)
excel_sheet.Cells(i, 11) = ""
excel_sheet.Cells(i, 12) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.Cells(i, 13) = IIf(IsNull(rs!goodmoisture), "", rs!goodmoisture)
excel_sheet.Cells(i, 14) = IIf(IsNull(rs!poormoisture), "", rs!poormoisture)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.Cells(i, 16) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
excel_sheet.Cells(i, 17) = IIf(IsNull(rs!tree_count_slowgrowing), "", rs!tree_count_slowgrowing)
excel_sheet.Cells(i, 18) = IIf(IsNull(rs!tree_count_dor), "", rs!tree_count_dor)
excel_sheet.Cells(i, 19) = IIf(IsNull(rs!tree_count_activegrowing), "", rs!tree_count_activegrowing)
excel_sheet.Cells(i, 20) = IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.Cells(i, 21) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.Cells(i, 22) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.Cells(i, 23) = IIf(IsNull(rs!leafpest), "", rs!leafpest)
excel_sheet.Cells(i, 24) = IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.Cells(i, 25) = IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.Cells(i, 26) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.Cells(i, 27) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.Cells(i, 28) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)


SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up




    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:Ab3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


excel_sheet.Columns("A:ab").Select
 excel_app.Selection.ColumnWidth = 15
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
db.Execute "drop table " & Mtblname & ""
db.Close
Exit Sub
err:
db.Execute "drop table " & Mtblname & ""
MsgBox err.Description
err.Clear


End Sub


Private Sub Form_Load()
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
frmfield.Width = 5595
'populatedate "phealthhub15_core", 11

Dim rs As New ADODB.Recordset
Dim da As New ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.CursorLocation = adUseClient
CONNLOCAL.Open OdkCnnString
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct substring(farmerbarcode,1,9) as dgt  from phealthhub15_core group by substring(farmerbarcode,1,9) order by substring(farmerbarcode,1,9)", CONNLOCAL
Set CBODGT.RowSource = rs
CBODGT.ListField = "dgt"
CBODGT.BoundColumn = "dgt"
Mname = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

End Sub
Private Sub populatedate(tt As String, fc As Integer)
Dim i, j, fcount As Integer
Operation = ""
Mindex = 0
'Mygrid.Visible = False
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                      
db.Open OdkCnnString
                     
Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & fc & "' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)
CBODATE.Clear
Set rs = Nothing
'"SELECT * FROM " & LCase(rs!tblname) & " WHERE  " & rs!Key & "='" & RsRemote.Fields(0) & "' ", CONNLOCAL  'IN(SELECT " & RS!Key & " FROM  " & RS!tblname & " )  ", CONNLOCAL
rs.Open "SELECT * FROM " & tt & " where 1", CONNLOCAL
For j = 0 To fcount - 1
If rs.Fields(j).Type = 135 Then

CBODATE.AddItem rs.Fields(j).Name
End If
Next

Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub OPTALLRECORDS_Click()
TXTRECORDNO.Text = ""
End Sub

Private Sub OPTALL_Click()
Frame1.Enabled = False
End Sub

Private Sub OPTALLVISIT_Click()
If OPTALLVISIT.Value = True Then
CHKMOREOPTION.Enabled = True
Else
CHKMOREOPTION.Enabled = False
End If
End Sub

Private Sub optdead_Click()
TXTVALUE.Text = 20
mfieldname = "tree_count_deadmissing"
mfname = "DEAD/MISSING"
End Sub

Private Sub OPTLEAFPEST_Click()
TXTVALUE.Text = 5
mfieldname = "activepest"
mfname = "LEAF PEST"
End Sub

Private Sub optmoist_Click()
TXTVALUE.Text = 30
mfieldname = ""
End Sub

Private Sub optpestdamage_Click()
TXTVALUE.Text = 5
End Sub

Private Sub optrootpest_Click()
TXTVALUE.Text = 5
mfieldname = "rootpest"
mfname = "ROOT PEST"
End Sub

Private Sub OPTSEL_Click()

Frame1.Enabled = True

End Sub

Private Sub OPTSLOWGROWING_Click()
mfieldname = "tree_count_slowgrowing"
mfname = "SLOW GROWING"
End Sub

Private Sub OPTSTEMPEST_Click()
TXTVALUE.Text = 5
mfieldname = "stempest"
mfname = "STEM PEST"
End Sub

Private Sub OPTTOPN_Click()
'If OPTTOPN.Value = True Then
'CHKMOREOPTION.Enabled = False
'Else
'CHKMOREOPTION.Enabled = True
'
'End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TXTRECORDNO_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TXTRECORDNO_LostFocus()
If Val(TXTRECORDNO.Text) > 1 Then
Label5.Caption = "VISITS"
ElseIf Val(TXTRECORDNO.Text) <= 0 Then
MsgBox "IN VALID INPUT!"
TXTRECORDNO.SetFocus
Else
Label5.Caption = "VISIT"
End If
End Sub


Private Sub allfieldsandstorage()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                   
GetTbl
If OPTALL.Value = True Then
Mindex = 51
End If
mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1





SQLSTR = ""



         SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,fstype,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,start, end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,'F',0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core "
         
            
                SQLSTR = SQLSTR & "where   status<>'BAD'"
                  

         
  
  db.Execute SQLSTR
  SQLSTR = ""
 
    SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,fstype,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
         & "waterlog,leafpest,animaldamage,monitorcomments) select start,tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,'S',totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,ndtrees,wlogged,pdamage,adamage,monitorcomments from storagehub6_core "
         SQLSTR = SQLSTR & "where   status<>'BAD'"
               
                
db.Execute SQLSTR
SQLSTR = ""
SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,fstype,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select distinct '1900-01-01','1900-01-01','1900-01-01','',substring(farmercode,1,3)," _
         & "substring(farmercode,4,3),substring(farmercode,7,3),farmercode,'N',0,0,0,0,0,0," _
         & "0,0,0,0,0,0,0,0,0,0," _
         & "0,0,0 from mhv.tblplanted where farmercode not in(select farmercode from " & Mtblname & " )"
         db.Execute SQLSTR
   
   
  Set rss = Nothing
  SQLSTR = ""

         
SQLSTR = ""
If CHKMOREOPTION.Value = 0 Then
   SQLSTR = "select * from " & Mtblname & " "
Else
    If optmoist.Value = True Then
    
        SQLSTR = "select * from " & Mtblname & " where (poormoisture/totaltally)*100>'" & Val(TXTVALUE.Text) & "'"
      
    ElseIf optrootpest.Value = True Then
    If CHKRANGE.Value = 0 Then
        SQLSTR = "select * from " & Mtblname & " where (rootpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    
    SQLSTR = "select * from " & Mtblname & " where (rootpest/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (rootpest/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
    End If
    
    ElseIf OPTSTEMPEST.Value = True Then
     If CHKRANGE.Value = 0 Then
        SQLSTR = "select * from " & Mtblname & " where (stempest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    SQLSTR = "select * from " & Mtblname & " where (stempest/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (stempest/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
        
    End If
    
    ElseIf OPTLEAFPEST.Value = True Then
    If CHKRANGE.Value = 0 Then
    
        SQLSTR = "select * from " & Mtblname & " where (leafpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    SQLSTR = "select * from " & Mtblname & " where (leafpest/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (leafpest/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
    
    End If
    
    ElseIf optdead.Value = True Then
    If CHKRANGE.Value = 0 Then
    
        SQLSTR = "select * from " & Mtblname & " where (tree_count_deadmissing/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
        Else
        SQLSTR = "select * from " & Mtblname & " where (tree_count_deadmissing/totaltrees)*100>='" & Val(txtrange1.Text) & "' and (deadmissing/totaltrees)*100<='" & Val(txtrange1.Text) & "'"
        End If
        
        Else
        
        If CHKRANGE.Value = 0 Then
        SQLSTR = "select * from " & Mtblname & " where (tree_count_slowgrowing/totaltrees)*100>'" & Val(TXTVALUE.Text) & "'"
    Else
    
    SQLSTR = "select * from " & Mtblname & " where (slowgrowing/totaltrees)*100>='" & Val(txtrange1.Text) & "'  and (deadmissing/totaltrees)*100<='" & Val(txtrange2.Text) & "'"
    End If
    
    End If

End If


SQLSTR = SQLSTR & " order by id,end,farmercode,fdcode"

'If CHKDGT.Value = 1 Then
'SQLSTR = ""
'SQLSTR = "select '' as end,id,substring(farmercode,1,9) as farmercode,sum(treesreceived) as treesreceived ,'' as fdcode,sum" _
'& "(totaltrees) totaltrees,sum(goodmoisture) as goodmoisture,sum(poormoisture) as poormoisture,sum(totaltally) as" _
'& "totaltally,sum(deadmissing) as deadmissing,sum(slowgrowing)as slowgrowing,sum(dor) dor,sum(activegrowing) as" _
'& "activegrowing,sum(shock)as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(leafpest) as leafpest,sum" _
'& "(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage,'' as monitorcomments " _
'& "from " & Mtblname & " group by substring(farmercode,1,9) order by substring(farmercode,1,9)"
'
'End If


'On Error Resume Next




'Dim RS As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer
Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("START DATE")
    excel_sheet.Cells(3, 3) = ProperCase("END DATE")
    excel_sheet.Cells(3, 4) = ProperCase("STAFF CODE-NAME")
    excel_sheet.Cells(3, 5) = ProperCase("DZONGKHAG")
    excel_sheet.Cells(3, 6) = ProperCase("GEWOG")
    excel_sheet.Cells(3, 7) = ProperCase("TSHOWOG")
    excel_sheet.Cells(3, 8) = ProperCase("Farmer ID")
    excel_sheet.Cells(3, 9) = ProperCase("Total Distributed")
    excel_sheet.Cells(3, 10) = ProperCase("Field ID")
    excel_sheet.Cells(3, 11) = ProperCase("Total Trees Distributed - Planted List")
    excel_sheet.Cells(3, 12) = ProperCase("Total Trees")
    excel_sheet.Cells(3, 13) = ProperCase("Good Moisture")
    excel_sheet.Cells(3, 14) = ProperCase("Poor Moisture")
    excel_sheet.Cells(3, 15) = ProperCase("Total Mositure Tally")
    excel_sheet.Cells(3, 16) = ProperCase("Dead Missing")
    excel_sheet.Cells(3, 17) = ProperCase("Slow Growing")
    excel_sheet.Cells(3, 18) = ProperCase("Dormant")
    excel_sheet.Cells(3, 19) = ProperCase("Active Growing")
    excel_sheet.Cells(3, 20) = ProperCase("Shock")
    excel_sheet.Cells(3, 21) = ProperCase("Nutrient Deficient")
    excel_sheet.Cells(3, 22) = ProperCase("Water Logg")
    excel_sheet.Cells(3, 23) = ProperCase("Leaf Pest")
    excel_sheet.Cells(3, 24) = ProperCase("Active Pest")
    excel_sheet.Cells(3, 25) = ProperCase("Stem Pest")
    excel_sheet.Cells(3, 26) = ProperCase("Root Pest")
    excel_sheet.Cells(3, 27) = ProperCase("Animal Damage")
    excel_sheet.Cells(3, 28) = ProperCase("comments")
    excel_sheet.Cells(3, 29) = ProperCase("Field")
     excel_sheet.Cells(3, 30) = ProperCase("Storage")
      excel_sheet.Cells(3, 31) = ProperCase("No Visit")
   i = 4
  Set rs = Nothing
  
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!Start
excel_sheet.Cells(i, 3) = "'" & rs!End  'rs.Fields(Mindex)
FindsTAFF rs!id
excel_sheet.Cells(i, 4) = rs!id & " " & sTAFF

FindDZ Mid(rs!farmercode, 1, 3)
excel_sheet.Cells(i, 5) = Mid(rs!farmercode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmercode, 4, 3) & " " & GEname
FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
excel_sheet.Cells(i, 7) = Mid(rs!farmercode, 7, 3) & " " & TsName
FindFA IIf(IsNull(rs!farmercode), "", rs!farmercode), "F"
excel_sheet.Cells(i, 8) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & " " & FAName


excel_sheet.Cells(i, 9) = IIf(IsNull(rs!treesreceived), "", rs!treesreceived)
excel_sheet.Cells(i, 10) = IIf(IsNull(rs!FDCODE), "", rs!FDCODE)
excel_sheet.Cells(i, 11) = ""
excel_sheet.Cells(i, 12) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.Cells(i, 13) = IIf(IsNull(rs!goodmoisture), "", rs!goodmoisture)
excel_sheet.Cells(i, 14) = IIf(IsNull(rs!poormoisture), "", rs!poormoisture)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.Cells(i, 16) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
excel_sheet.Cells(i, 17) = IIf(IsNull(rs!tree_count_slowgrowing), "", rs!tree_count_slowgrowing)
excel_sheet.Cells(i, 18) = IIf(IsNull(rs!tree_count_dor), "", rs!tree_count_dor)
excel_sheet.Cells(i, 19) = IIf(IsNull(rs!tree_count_activegrowing), "", rs!tree_count_activegrowing)
excel_sheet.Cells(i, 20) = IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.Cells(i, 21) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.Cells(i, 22) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.Cells(i, 23) = IIf(IsNull(rs!leafpest), "", rs!leafpest)
excel_sheet.Cells(i, 24) = IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.Cells(i, 25) = IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.Cells(i, 26) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.Cells(i, 27) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.Cells(i, 28) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)
If rs!fstype = "F" Then
excel_sheet.Cells(i, 29) = "F"
ElseIf rs!fstype = "S" Then
excel_sheet.Cells(i, 30) = "S"
Else
excel_sheet.Cells(i, 31) = "N"

End If


SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up




    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:Ab3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


excel_sheet.Columns("A:ab").Select
 excel_app.Selection.ColumnWidth = 15
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
db.Execute "drop table " & Mtblname & ""
db.Close
Exit Sub
err:
db.Execute "drop table " & Mtblname & ""
MsgBox err.Description
err.Clear


End Sub


