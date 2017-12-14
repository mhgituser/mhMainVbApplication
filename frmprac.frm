VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmprac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUKTI"
   ClientHeight    =   6405
   ClientLeft      =   4950
   ClientTop       =   2145
   ClientWidth     =   11715
   Icon            =   "frmprac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11715
   Begin VB.CommandButton Command35 
      Caption         =   "farmer not assigned"
      Height          =   855
      Left            =   10320
      TabIndex        =   44
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command34 
      Caption         =   "daily act vs field and storage"
      Height          =   735
      Left            =   8520
      TabIndex        =   43
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command33 
      Caption         =   "create kml"
      Height          =   855
      Left            =   5640
      TabIndex        =   42
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton Command32 
      Caption         =   "googledoc"
      Height          =   495
      Left            =   3720
      TabIndex        =   41
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command31 
      Caption         =   "pivot"
      Height          =   615
      Left            =   1440
      TabIndex        =   39
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command30 
      Caption         =   "graph"
      Height          =   615
      Left            =   8040
      TabIndex        =   38
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command29 
      Caption         =   "week day"
      Height          =   375
      Left            =   1320
      TabIndex        =   37
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command28 
      Caption         =   "chk farmer code error"
      Height          =   615
      Left            =   120
      TabIndex        =   36
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command27 
      Caption         =   "chilkat"
      Height          =   615
      Left            =   9960
      TabIndex        =   35
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command26 
      Caption         =   "email me"
      Height          =   495
      Left            =   6360
      TabIndex        =   34
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command25 
      Caption         =   "email"
      Height          =   735
      Left            =   4920
      TabIndex        =   33
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "odk error check"
      Height          =   855
      Left            =   600
      TabIndex        =   32
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command23 
      Caption         =   "update variety in pbatch detail"
      Height          =   615
      Left            =   10080
      TabIndex        =   31
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command22 
      Caption         =   "update chemicaltrade name id"
      Height          =   495
      Left            =   6120
      TabIndex        =   30
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command21 
      Caption         =   "update qms boxes"
      Height          =   495
      Left            =   3240
      TabIndex        =   29
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command20 
      Caption         =   "all reports"
      Height          =   495
      Left            =   9000
      TabIndex        =   27
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "update all"
      Height          =   735
      Left            =   8400
      TabIndex        =   26
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "field pest summary"
      Height          =   735
      Left            =   3480
      TabIndex        =   20
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton Command17 
      Caption         =   "field and storage"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command16 
      Caption         =   "dist update"
      Height          =   615
      Left            =   2640
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "first monday"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   "update staffbarcode"
      Height          =   855
      Left            =   3720
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Total no farmers and acres registered - field"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "update dailyact"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "drop out"
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "update storage"
      Height          =   735
      Left            =   9720
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "up date field"
      Height          =   495
      Left            =   8400
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "temp bar code"
      Height          =   855
      Left            =   1440
      TabIndex        =   10
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "odk vs local"
      Height          =   975
      Left            =   10080
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   855
      Left            =   9600
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5880
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   8280
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1695
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   2990
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   1335
      Left            =   1440
      TabIndex        =   21
      Top             =   1320
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
      FormatString    =   $"frmprac.frx":06EA
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
   Begin MSComCtl2.DTPicker txtfrmdate 
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130940929
      CurrentDate     =   41362
   End
   Begin MSComCtl2.DTPicker txttodate 
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130940929
      CurrentDate     =   41362
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130940929
      CurrentDate     =   41362
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   130940929
      CurrentDate     =   41362
   End
   Begin VSFlex7Ctl.VSFlexGrid mygrid1 
      Height          =   1335
      Left            =   3360
      TabIndex        =   28
      Top             =   2520
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
      FormatString    =   $"frmprac.frx":074D
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
   Begin VB.Label Label1 
      Caption         =   "web link"
      Height          =   375
      Left            =   960
      TabIndex        =   40
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   720
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "frmprac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mymax As Integer


Private Sub Command10_Click()
Dim SQLSTR As String

Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
Dim frcode As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
  

db.Execute "delete from mtemp"


SQLSTR = ""
   SQLSTR = "insert into mtemp SELECT _URI, region_dcode, region_gcode, region,fcode FROM storagehub6_core where farmerbarcode=''"
  ODKDB.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = UCase(Mid(rss!dcode, 1, 3))
  mgcode = UCase(Mid(rss!gcode, 1, 3))
  mtcode = UCase(Mid(rss!tcode, 1, 3))
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  Set rsF = Nothing
  
  db.Execute "update storagehub6_core set farmerbarcode='" & mfcode & "' where region_dcode='" & rss!dcode & "' and region_gcode='" & rss!gcode & "' and region='" & rss!tcode & "' and fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"

frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
LogRemarks = "table storagehub6_core updated successfully.farmerbarcode updated(" & frcode & ")"
updateodklog "no uri", Now, MUSER, LogRemarks, "storagehub6_core"


frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM storagehub6_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update storagehub6_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
  LogRemarks = "table storagehub6_core updated successfully.staffbarcode updated(" & frcode & ")"
  updateodklog "no uri", Now, MUSER, LogRemarks, "storagehub6_core"

  MsgBox "done"
  
  'storage
  
 
                        
                        
                        
End Sub

Private Sub Command12_Click()
mchk = True
chkred = True
Dim SQLSTR As String
Dim myphone As String
Dim TOTLAND As Double
Dim totadd As Double
TOTLAND = 0
totadd = 0
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Set rsadd = Nothing
'Dim sqlstr As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
'db.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ODKLOCAL;Initial Catalog=odk_prodLocal" ' local connection
'odk_prodLocal
db.Open OdkCnnString
                      
'db.Open tempstr
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM dailyacthub9_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update dailyacthub9_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"

  rss.MoveNext
  Loop
  
  
  'field
  
 db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM storagehub6_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update storagehub6_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 
  rss.MoveNext
  Loop
 ' LogRemarks = "table dailyacthub9_core updated successfully.staffbarcode updated(" & frcode & ")"
'updateodklog "no uri", Now, MUSER, LogRemarks, "dailyacthub9_core"

                        
                        
                        
                        

MsgBox "done"
End Sub

Private Sub Command13_Click()
Dim rs As New ADODB.Recordset

'GetTblmhv
MHVDB.Execute "insert into " & Mtblname & " (farmercode) select distinct farmerbarcode from odk_prodlocal.phealthhub15_core" ' where substring(end,1,10)<='2013-05-31'"
'MHVDB.Execute "insert into " & Mtblname & " (farmercode) select distinct farmerbarcode from odk_prodlocal.storagehub6_core where substring(end,1,10)<='2013-04-30'"


Set rs = Nothing

rs.Open "select distinct idfarmer,sum(regland) from tblfarmer as a,tbllandreg as b where idfarmer=farmerid and a.status='A' and " _
& " idfarmer  in(select farmercode from " & Mtblname & ") group by idfarmer", MHVDB

MsgBox "hjags"

End Sub

Private Sub Command14_Click()
updateField
'updateStorage
'updateDailyact
updateregistration
End Sub



Private Sub Command16_Click()
Dim SQLSTR As String

Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
Dim frcode As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                       

db.Execute "delete from mtemp"



frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM distribution3_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update distribution3_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 
  rss.MoveNext
  Loop
 

End Sub

Private Sub Command17_Click()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim muk As New ADODB.Recordset
Dim actstring As String
Dim CrtStr As String
Dim totregland As Double
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
totregland = 0
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                    
'If OPTALL.Value = True Then
'Mindex = 51
'End If

Dim SQLSTR As String
SQLSTR = ""
SLNO = 1
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""

SQLSTR = ""
   
    
    GetTbl
        
    
SQLSTR = ""

           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
          
  db.Execute SQLSTR
  Set rss = Nothing
  Set rs1 = Nothing
  rs1.Open "select sum(regland) as regland from mhv.tbllandreg where farmerid in(select farmercode from " & Mtblname & ")", ODKDB
totregland = rs1!regland
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""


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
    excel_app.Visible = False
     excel_sheet.cells(1, 1) = ProperCase("Field")
     
    excel_sheet.cells(2, 1) = ProperCase("Total No. of hazelnut field")
    excel_sheet.cells(3, 1) = ProperCase("Total No. of trees in the field")
    excel_sheet.cells(4, 1) = ProperCase("Total acres")
    excel_sheet.cells(5, 1) = ""
    excel_sheet.cells(6, 1) = ProperCase("Slow growing")
    excel_sheet.cells(7, 1) = ProperCase("Dormant")
    excel_sheet.cells(8, 1) = ProperCase("Dead ")
    excel_sheet.cells(9, 1) = ProperCase("Active growing")
    excel_sheet.cells(10, 1) = ""
    excel_sheet.cells(11, 1) = ProperCase("Shock")
    excel_sheet.cells(12, 1) = ProperCase("Nutrient deficeint")
    excel_sheet.cells(13, 1) = ProperCase("Waterlog")
    excel_sheet.cells(14, 1) = ProperCase("Leafpest")
    excel_sheet.cells(15, 1) = ProperCase("Active pest")
    excel_sheet.cells(16, 1) = ProperCase("Stem pest")
    excel_sheet.cells(17, 1) = ProperCase("Root pest")
    excel_sheet.cells(18, 1) = ProperCase("Animal Damage")
    excel_sheet.cells(1, 2) = "All"
    excel_sheet.cells(1, 3) = "%"
    excel_sheet.cells(1, 4) = "Field"
    excel_sheet.cells(1, 5) = "%"
    excel_sheet.cells(1, 6) = "Storage All"
    excel_sheet.cells(1, 7) = "%"
    excel_sheet.cells(1, 8) = "Storage Only"
    excel_sheet.cells(1, 9) = "%"
    excel_sheet.cells(1, 10) = "Storage (but have some trees in field)"
     excel_sheet.cells(1, 11) = "%"
    
    
    Set rs = Nothing
    rs.Open SQLSTR, ODKDB
   Call fillcell(excel_sheet, 4, rs, Round(totregland, 0))
    
    
    
    SQLSTR = ""
ODKDB.Execute "delete from " & Mtblname & ""
           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,0,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,0,0,0,0,ndtrees,wlogged,0,0,0," _
         & "0,adamage,'0' from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"
          
  db.Execute SQLSTR
  Set rss = Nothing
  Set rs1 = Nothing
  totregland = 0
  rs1.Open "select sum(regland) as regland from mhv.tbllandreg where farmerid in(select farmercode from " & Mtblname & ")", ODKDB
totregland = rs1!regland
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""

 Set rs = Nothing
    rs.Open SQLSTR, ODKDB
    Call fillcell(excel_sheet, 6, rs, Round(totregland, 0))
    
    
    
    
    
  db.Execute "delete from tempfarmernotinfield"
db.Execute " insert into tempfarmernotinfield(end,farmercode,staffbarcode)" _
           & "select distinct '' as end, farmerbarcode,staffbarcode from storagehub6_core " _
           & "where farmerbarcode not in (select farmerbarcode from phealthhub15_core) group by farmerbarcode"
           
SQLSTR = " delete from tempfarmernotinfield  where farmercode  in(" _
& "select farmercode from mhv.tblplanted as a , mhv.tblfarmer as b  " _
& "where farmercode=idfarmer)"

db.Execute SQLSTR
    
    SQLSTR = ""
ODKDB.Execute "delete from " & Mtblname & ""
           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,0,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,0,0,0,0,ndtrees,wlogged,0,0,0," _
         & "0,adamage,'0' from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and n.farmerbarcode in(select farmercode from tempfarmernotinfield ) GROUP BY n.farmerbarcode"
          
  db.Execute SQLSTR
  Set rss = Nothing
  Set rs1 = Nothing
  totregland = 0
  rs1.Open "select sum(regland) as regland from mhv.tbllandreg where farmerid in(select farmercode from " & Mtblname & ")", ODKDB
totregland = IIf(IsNull(rs1!regland), 0, rs1!regland)
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""

 Set rs = Nothing
    rs.Open SQLSTR, ODKDB
    Call fillcell(excel_sheet, 8, rs, Round(totregland, 0))
    
    
    
    
   
    
  For i = 2 To 18
  If i <> 5 Or i <> 10 Then
  excel_sheet.cells(i, 2) = excel_sheet.cells(i, 4) + excel_sheet.cells(i, 6)
  excel_sheet.cells(i, 2).NumberFormat = "0"
  End If
  Next
 For i = 2 To 18
  If i <> 5 Or i <> 10 Then
  excel_sheet.cells(i, 10) = excel_sheet.cells(i, 6) - excel_sheet.cells(i, 8)
  excel_sheet.cells(i, 2).NumberFormat = "0"
  End If
  Next
For i = 6 To 18
  If i <> 10 Then
  excel_sheet.cells(i, 3) = (excel_sheet.cells(i, 2) / excel_sheet.cells(3, 2))
   excel_sheet.cells(i, 3).NumberFormat = "0.00%"
   excel_sheet.cells(i, 5) = (excel_sheet.cells(i, 4) / excel_sheet.cells(3, 4))
   excel_sheet.cells(i, 5).NumberFormat = "0.00%"
   excel_sheet.cells(i, 7) = (excel_sheet.cells(i, 6) / excel_sheet.cells(3, 6))
   excel_sheet.cells(i, 7).NumberFormat = "0.00%"
   excel_sheet.cells(i, 9) = (excel_sheet.cells(i, 8) / excel_sheet.cells(3, 8))
   excel_sheet.cells(i, 9).NumberFormat = "0.00%"
   excel_sheet.cells(i, 11) = (excel_sheet.cells(i, 10) / excel_sheet.cells(3, 10))
   excel_sheet.cells(i, 11).NumberFormat = "0.00%"
  End If
  Next


    


    With excel_sheet
    
     'excel_sheet.Range("a1:b15").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Field and Storage Summary"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("A:a").Select
 excel_app.selection.columnWidth = 32
With excel_app.selection
.HorizontalAlignment = xlLeft
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

excel_sheet.Columns("b:k").Select
 excel_app.selection.columnWidth = 14
With excel_app.selection
.HorizontalAlignment = xlRight
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With





With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With

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
Private Function fillcell(excel_sheet As Object, col As Integer, rs As Object, totregland As Double)
excel_sheet.cells(2, col) = rs!fieldcode
    excel_sheet.cells(3, col) = rs!totaltrees
    excel_sheet.cells(4, col) = totregland
    excel_sheet.cells(5, col) = ""
    excel_sheet.cells(6, col) = rs!slowgrowing
    excel_sheet.cells(7, col) = rs!dor
    excel_sheet.cells(8, col) = rs!dead
    excel_sheet.cells(9, col) = rs!activegrowing
    excel_sheet.cells(10, col) = ""
    excel_sheet.cells(11, col) = rs!shock
    excel_sheet.cells(12, col) = rs!nutrient
    excel_sheet.cells(13, col) = rs!waterlog
    excel_sheet.cells(14, col) = rs!leafpest
    excel_sheet.cells(15, col) = rs!activepest
    excel_sheet.cells(16, col) = rs!stempest
    excel_sheet.cells(17, col) = rs!rootpest
    excel_sheet.cells(18, col) = rs!animaldamage
End Function

Private Sub allfieldpestdetail()
Dim SLNO As Integer
Dim fdno As Integer
Dim rs As New ADODB.Recordset
Dim marray
marray = Array("0% to 10%", "11% to 20%", "21% to 50%", ">50%")
Dim fcriteria As Integer
Dim mdgt, locstring As String
Dim m, n As Integer
Dim t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16 As Integer
Dim TOTTREES, MFLD, FCNT As Double
Dim MCOL As Integer
Dim muk As New ADODB.Recordset
fdno = 0
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
    Dim sl As Integer
    sl = 1
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 3) = ProperCase("GEWOG")
    excel_sheet.cells(3, 4) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 5) = ProperCase("FARMER NAME")
    MCOL = 6
    For m = 1 To Mygrid.rows - 1
    If Len(Mygrid.TextMatrix(m, 0)) = 0 Then Exit For
    If Len(Mygrid.TextMatrix(m, 1)) <> 0 Then
    excel_sheet.cells(3, MCOL) = Mygrid.TextMatrix(m, 0) & " % To " & Mygrid.TextMatrix(m, 1) & " %"
    If MCOL = 6 Or MCOL = 9 Or MCOL = 12 Then
    excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
    End If

    
    
    Else
      excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.cells(3, MCOL) = ">= Then " & Mygrid.TextMatrix(m, 0) & " %"
      End If
    MCOL = MCOL + 3
    Next
   i = 4
   MCOL = 5
   excel_sheet.cells(i, 6) = ProperCase("Last Visited")
   excel_sheet.cells(i, 7) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 8) = ProperCase("SLowgrowing")
   excel_sheet.cells(i, 9) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 10) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 11) = ProperCase("SLowgrowing")
   excel_sheet.cells(i, 12) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 13) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 14) = ProperCase("SLowgrowing")
   excel_sheet.cells(i, 15) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 16) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 17) = ProperCase("SLowgrowing")
   excel_sheet.cells(i, 18) = ProperCase("FIELD CODE")
  i = 5
   n = i
   SLNO = 1
   MCOL = 5
Set rs = Nothing
rs.Open "select end, fdcode,(farmercode) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(tree_count_slowgrowing) as mfieldname,count(fdcode) as cnt from " & Mtblname & "  group by farmercode,fdcode order by farmercode,fdcode,end", db
Dim md, mg, mt As String
  Do While rs.EOF <> True
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
  fdno = 0
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
FindFA rs!farmercode, "F"
    excel_sheet.cells(i, 5) = rs!farmercode & " " & FAName
      excel_sheet.cells(i, 6) = "'" & rs!end
      
    fcriteria = (rs!mfieldname / rs!totaltrees) * 100
    If fcriteria >= Val(Mygrid.TextMatrix(1, 0)) And fcriteria <= Val(Mygrid.TextMatrix(1, 1)) Then
    MCOL = 7
     t5 = t5 + rs!totaltrees
      t6 = t6 + rs!mfieldname
      t7 = t7 + rs!cnt
      If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    ElseIf fcriteria >= Val(Mygrid.TextMatrix(2, 0)) And fcriteria <= Val(Mygrid.TextMatrix(2, 1)) Then
     t8 = t8 + rs!totaltrees
      t9 = t9 + rs!mfieldname
      t10 = t10 + rs!cnt
         If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 10
    ElseIf fcriteria >= Val(Mygrid.TextMatrix(3, 0)) And fcriteria <= Val(Mygrid.TextMatrix(3, 1)) Then
      t11 = t11 + rs!totaltrees
      t12 = t12 + rs!mfieldname
      t13 = t13 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 13
    Else
       t14 = t14 + rs!totaltrees
      t15 = t15 + rs!mfieldname
      t16 = t16 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 16
    End If
  
    excel_sheet.cells(i, MCOL) = rs!totaltrees
    excel_sheet.cells(i, MCOL + 1) = rs!mfieldname
    excel_sheet.cells(i, MCOL + 2) = rs!FDCODE
  

i = i + 1
mdgt = Mid(rs!dgt, 1, 9)
If rs.EOF Then Exit Do
rs.MoveNext

Loop
locstring = ""
locstring = md & " " & Dzname & " " & mg & " " & GEname & " " & mt & " " & TsName

    excel_sheet.cells(i, 7) = IIf(t5 <> 0, t5, "")
    excel_sheet.cells(i, 8) = IIf(t6 <> 0, t6, "")
    excel_sheet.cells(i, 9) = IIf(t7 <> 0, t7, "")
    excel_sheet.cells(i, 10) = IIf(t8 <> 0, t8, "")
    excel_sheet.cells(i, 11) = IIf(t9 <> 0, t9, "")
    excel_sheet.cells(i, 12) = IIf(t10 <> 0, t10, "")
    excel_sheet.cells(i, 13) = IIf(t11 <> 0, t11, "")
    excel_sheet.cells(i, 14) = IIf(t12 <> 0, t12, "")
    excel_sheet.cells(i, 15) = IIf(t13 <> 0, t13, "")
    excel_sheet.cells(i, 16) = IIf(t14 <> 0, t14, "")
    excel_sheet.cells(i, 17) = IIf(t15 <> 0, t15, "")
    excel_sheet.cells(i, 18) = IIf(t16 <> 0, t16, "")
    excel_sheet.Range(excel_sheet.cells(n, 1), _
                             excel_sheet.cells(i - 1, 1)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 1) = SLNO
    excel_sheet.Range(excel_sheet.cells(n, 2), _
                             excel_sheet.cells(i - 1, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 2) = locstring
                            excel_sheet.Range(excel_sheet.cells(i, 2), _
                             excel_sheet.cells(i, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(i, 2) = "TOTAL"
    excel_sheet.Range(excel_sheet.cells(i, 2), _
 excel_sheet.cells(i, 17)).Select
excel_app.selection.Font.Bold = True
SLNO = SLNO + 1
i = i + 1
n = i
fdno = 0
   Loop
  excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:r4").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
excel_sheet.Columns("A:aa").Select
 excel_app.selection.columnWidth = 15
With excel_app.selection
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
       
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB
End With

db.Execute "drop table " & Mtblname & ""
excel_sheet.name = "Slowgrowing_Detail"
Excel_WBook.Sheets("sheet2").Activate
  If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
excel_sheet.name = "Slowgrowing_Summary"
mchk = True
Dim Jmth, K As Integer
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
             SQLSTR = ""
SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate,n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  year(end),month(end) ,farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY year(end),month(end) , n.farmerbarcode, n.fdcode"
         
    db.Execute SQLSTR
    
    GetTbl1
   SQLSTR = ""
SQLSTR = "insert into " & Mtblname1 & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate, n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY  n.farmerbarcode, n.fdcode"

    db.Execute SQLSTR

                SQLSTR = "select COUNT(fdcode) as fcnt, SUBSTRING(farmercode,1,9) as id ,SUM(totaltrees) as tt,sum(tree_count_slowgrowing) as jval,(sum(tree_count_slowgrowing)/sum(totaltrees)*100) as percent,year(end) as procyear,month(end) as procmonth  from " & Mtblname & "   where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),substring(farmercode,1,9) order by substring(farmercode,1,9),year(end),month(end)"
   For i = 1 To 13
    mtot(i) = 0
Next
Set rs = Nothing
rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    Screen.MousePointer = vbHourglass
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    excel_sheet.cells(2, 1) = ProperCase("MONTHLY SUMMARY OF Slowgrowing")
    excel_sheet.cells(3, 1) = ProperCase("DZONGKHAG  GEWOG  TSHOWOG")
     excel_sheet.cells(3, 2) = ProperCase("TOTAL TREES")
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 2
        excel_sheet.cells(3, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value)) & " " & " Slowgrowing"
       ' excel_sheet.Cells(3, K + 1) = ProperCase("Slowgrowing")
       
        K = K + 2
    Next
Dim ll As Integer
    jrow = 4
    For ll = 3 To 50
    excel_sheet.cells(jrow, ll) = ProperCase("NO.OF FIELDS")
    excel_sheet.cells(jrow, ll + 1) = ProperCase("TREE NOS.")
    excel_sheet.cells(jrow, ll + 2) = ProperCase("No.")
    excel_sheet.cells(jrow, ll + 3) = ProperCase("%")
    ll = ll + 3
    Next
    excel_sheet.cells(jrow, ll) = "Slowgrowing_Detail"
    Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       locstr = ""
       FindDZ Mid(rs!id, 1, 3)
       FindGE Mid(rs!id, 1, 3), Mid(rs!id, 4, 3)
       FindTs Mid(rs!id, 1, 3), Mid(rs!id, 4, 3), Mid(rs!id, 7, 3)
       locstr = Mid(rs!id, 1, 3) & " " & Dzname & " " & Mid(rs!id, 4, 3) & " " & GEname & " " & Mid(rs!id, 7, 3) & " " & TsName
       excel_sheet.cells(jrow, 1) = locstr
       Do While pYear = rs!id
    If rs!procmonth = 1 Then
    i = 3
    ElseIf rs!procmonth = 2 Then
    i = 7
    ElseIf rs!procmonth = 3 Then
    i = 11
    ElseIf rs!procmonth = 4 Then
    i = 15
    ElseIf rs!procmonth = 5 Then
    i = 19
    ElseIf rs!procmonth = 6 Then
    i = 23
    ElseIf rs!procmonth = 7 Then
    i = 27
    ElseIf rs!procmonth = 8 Then
    i = 31
    ElseIf rs!procmonth = 9 Then
    i = 35
    ElseIf rs!procmonth = 10 Then
    i = 39
    ElseIf rs!procmonth = 11 Then
    i = 43
    Else
    i = 47
    End If
Set muk = Nothing
muk.Open "select sum(totaltrees) as ttrees from " & Mtblname1 & " where substring(farmercode,1,9)='" & rs!id & "'", ODKDB
j = rs!jval
     excel_sheet.cells(jrow, 2) = muk!ttrees
           jtot = jtot + muk!ttrees
          excel_sheet.cells(jrow, i) = rs!FCNT
          excel_sheet.cells(jrow, i + 1) = rs!tt
           
           excel_sheet.cells(jrow, i + 2) = rs!jval
           excel_sheet.cells(jrow, i + 3) = (rs!jval / rs!tt)
           excel_sheet.cells(jrow, i + 3).NumberFormat = "0.00%"
          excel_sheet.cells(jrow, 51) = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
      
    Loop
    'jtot = 0
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, 51)).Select
    excel_app.selection.Columns.AutoFit
    'excel_sheet.Columns("b:ag").Select
excel_app.selection.columnWidth = 15
    excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
   excel_sheet.Range("A1:aY4").Font.Bold = True
   db.Execute "drop table " & Mtblname & ""
   db.Execute "drop table " & Mtblname1 & ""
   ' slow growing ends here
   'active leaf pest starts here
   
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



Excel_WBook.Sheets("sheet3").Activate
  If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
excel_sheet.name = "Active_Leaf_Pest_Detail"

   
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 3) = ProperCase("GEWOG")
    excel_sheet.cells(3, 4) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 5) = ProperCase("FARMER NAME")
    MCOL = 6
    For m = 1 To Mygrid.rows - 1
    If Len(Mygrid.TextMatrix(m, 0)) = 0 Then Exit For
    If Len(Mygrid.TextMatrix(m, 1)) <> 0 Then
    excel_sheet.cells(3, MCOL) = Mygrid.TextMatrix(m, 0) & " % To " & Mygrid.TextMatrix(m, 1) & " %"
    If MCOL = 6 Or MCOL = 9 Or MCOL = 12 Then
    excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
    End If

    
    
    Else
      excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.cells(3, MCOL) = ">= Then " & Mygrid.TextMatrix(m, 0) & " %"
      End If
    MCOL = MCOL + 3
    Next
   i = 4
   MCOL = 5
   excel_sheet.cells(i, 6) = ProperCase("Last Visited")
   excel_sheet.cells(i, 7) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 8) = ProperCase("Active Leaf Pest")
   excel_sheet.cells(i, 9) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 10) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 11) = ProperCase("Active Leaf Pest")
   excel_sheet.cells(i, 12) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 13) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 14) = ProperCase("Active Leaf Pest")
   excel_sheet.cells(i, 15) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 16) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 17) = ProperCase("Active Leaf Pest")
   excel_sheet.cells(i, 18) = ProperCase("FIELD CODE")
  i = 5
   n = i
   SLNO = 1
   MCOL = 5
Set rs = Nothing
rs.Open "select end, fdcode,(farmercode) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(activepest) as mfieldname,count(fdcode) as cnt from " & Mtblname & "  group by farmercode,fdcode order by farmercode,fdcode,end", db

  Do While rs.EOF <> True
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
  fdno = 0
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
FindFA rs!farmercode, "F"
    excel_sheet.cells(i, 5) = rs!farmercode & " " & FAName
      excel_sheet.cells(i, 6) = "'" & rs!end
      
    fcriteria = (rs!mfieldname / rs!totaltrees) * 100
    If fcriteria >= Val(Mygrid.TextMatrix(1, 0)) And fcriteria <= Val(Mygrid.TextMatrix(1, 1)) Then
    MCOL = 7
     t5 = t5 + rs!totaltrees
      t6 = t6 + rs!mfieldname
      t7 = t7 + rs!cnt
      If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    ElseIf fcriteria >= Val(Mygrid.TextMatrix(2, 0)) And fcriteria <= Val(Mygrid.TextMatrix(2, 1)) Then
     t8 = t8 + rs!totaltrees
      t9 = t9 + rs!mfieldname
      t10 = t10 + rs!cnt
         If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 10
    ElseIf fcriteria >= Val(Mygrid.TextMatrix(3, 0)) And fcriteria <= Val(Mygrid.TextMatrix(3, 1)) Then
      t11 = t11 + rs!totaltrees
      t12 = t12 + rs!mfieldname
      t13 = t13 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 13
    Else
       t14 = t14 + rs!totaltrees
      t15 = t15 + rs!mfieldname
      t16 = t16 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 16
    End If
  
    excel_sheet.cells(i, MCOL) = rs!totaltrees
    excel_sheet.cells(i, MCOL + 1) = rs!mfieldname
    excel_sheet.cells(i, MCOL + 2) = rs!FDCODE
  

i = i + 1
mdgt = Mid(rs!dgt, 1, 9)
If rs.EOF Then Exit Do
rs.MoveNext

Loop
locstring = ""
locstring = md & " " & Dzname & " " & mg & " " & GEname & " " & mt & " " & TsName

    excel_sheet.cells(i, 7) = IIf(t5 <> 0, t5, "")
    excel_sheet.cells(i, 8) = IIf(t6 <> 0, t6, "")
    excel_sheet.cells(i, 9) = IIf(t7 <> 0, t7, "")
    excel_sheet.cells(i, 10) = IIf(t8 <> 0, t8, "")
    excel_sheet.cells(i, 11) = IIf(t9 <> 0, t9, "")
    excel_sheet.cells(i, 12) = IIf(t10 <> 0, t10, "")
    excel_sheet.cells(i, 13) = IIf(t11 <> 0, t11, "")
    excel_sheet.cells(i, 14) = IIf(t12 <> 0, t12, "")
    excel_sheet.cells(i, 15) = IIf(t13 <> 0, t13, "")
    excel_sheet.cells(i, 16) = IIf(t14 <> 0, t14, "")
    excel_sheet.cells(i, 17) = IIf(t15 <> 0, t15, "")
    excel_sheet.cells(i, 18) = IIf(t16 <> 0, t16, "")
    excel_sheet.Range(excel_sheet.cells(n, 1), _
                             excel_sheet.cells(i - 1, 1)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 1) = SLNO
    excel_sheet.Range(excel_sheet.cells(n, 2), _
                             excel_sheet.cells(i - 1, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 2) = locstring
                            excel_sheet.Range(excel_sheet.cells(i, 2), _
                             excel_sheet.cells(i, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(i, 2) = "TOTAL"
    excel_sheet.Range(excel_sheet.cells(i, 2), _
 excel_sheet.cells(i, 17)).Select
excel_app.selection.Font.Bold = True
SLNO = SLNO + 1
i = i + 1
n = i
fdno = 0
   Loop
  excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:r4").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
excel_sheet.Columns("A:aa").Select
 excel_app.selection.columnWidth = 15
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With

With excel_sheet.PageSetup
       
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB
End With
db.Execute "drop table " & Mtblname & ""

Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Active_Leaf_Pest_Summary"

mchk = True

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

CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
              GetTbl
             SQLSTR = ""
SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate,n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  year(end),month(end) ,farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY year(end),month(end) , n.farmerbarcode, n.fdcode"
         
    db.Execute SQLSTR
    
    GetTbl1
   SQLSTR = ""
SQLSTR = "insert into " & Mtblname1 & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate, n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY  n.farmerbarcode, n.fdcode"

    db.Execute SQLSTR

                SQLSTR = "select COUNT(fdcode) as fcnt, SUBSTRING(farmercode,1,9) as id ,SUM(totaltrees) as tt,sum(activepest) as jval,(sum(activepest)/sum(totaltrees)*100) as percent,year(end) as procyear,month(end) as procmonth  from " & Mtblname & "   where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),substring(farmercode,1,9) order by substring(farmercode,1,9),year(end),month(end)"
   For i = 1 To 13
    mtot(i) = 0
Next
Set rs = Nothing
rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    Screen.MousePointer = vbHourglass
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    excel_sheet.cells(2, 1) = ProperCase("MONTHLY SUMMARY OF  Active Leaf Pest")
    excel_sheet.cells(3, 1) = ProperCase("DZONGKHAG  GEWOG  TSHOWOG")
     excel_sheet.cells(3, 2) = ProperCase("TOTAL TREES")
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 2
        excel_sheet.cells(3, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value)) & " " & " Active Leaf Pest"
        'excel_sheet.Cells(3, K + 1) = ProperCase("Active Leaf Pest")
        K = K + 2
    Next

    jrow = 4
    For ll = 3 To 50
    excel_sheet.cells(jrow, ll) = ProperCase("NO.OF FIELDS")
    excel_sheet.cells(jrow, ll + 1) = ProperCase("TREE NOS.")
    excel_sheet.cells(jrow, ll + 2) = ProperCase("No.")
    excel_sheet.cells(jrow, ll + 3) = ProperCase("%")
    ll = ll + 3
    Next
    excel_sheet.cells(jrow, ll) = "Active_Leaf_Pest_Detail"
    Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       locstr = ""
       FindDZ Mid(rs!id, 1, 3)
       FindGE Mid(rs!id, 1, 3), Mid(rs!id, 4, 3)
       FindTs Mid(rs!id, 1, 3), Mid(rs!id, 4, 3), Mid(rs!id, 7, 3)
       locstr = Mid(rs!id, 1, 3) & " " & Dzname & " " & Mid(rs!id, 4, 3) & " " & GEname & " " & Mid(rs!id, 7, 3) & " " & TsName
       excel_sheet.cells(jrow, 1) = locstr
       Do While pYear = rs!id
    If rs!procmonth = 1 Then
    i = 3
    ElseIf rs!procmonth = 2 Then
    i = 7
    ElseIf rs!procmonth = 3 Then
    i = 11
    ElseIf rs!procmonth = 4 Then
    i = 15
    ElseIf rs!procmonth = 5 Then
    i = 19
    ElseIf rs!procmonth = 6 Then
    i = 23
    ElseIf rs!procmonth = 7 Then
    i = 27
    ElseIf rs!procmonth = 8 Then
    i = 31
    ElseIf rs!procmonth = 9 Then
    i = 35
    ElseIf rs!procmonth = 10 Then
    i = 39
    ElseIf rs!procmonth = 11 Then
    i = 43
    Else
    i = 47
    End If
Set muk = Nothing
muk.Open "select sum(totaltrees) as ttrees from " & Mtblname1 & " where substring(farmercode,1,9)='" & rs!id & "'", ODKDB
j = rs!jval
     excel_sheet.cells(jrow, 2) = muk!ttrees
           jtot = jtot + muk!ttrees
          excel_sheet.cells(jrow, i) = rs!FCNT
          excel_sheet.cells(jrow, i + 1) = rs!tt
           
           excel_sheet.cells(jrow, i + 2) = rs!jval
           excel_sheet.cells(jrow, i + 3) = (rs!jval / rs!tt)
           excel_sheet.cells(jrow, i + 3).NumberFormat = "0.00%"
          excel_sheet.cells(jrow, 51) = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
      
    Loop
    'jtot = 0
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, 51)).Select
    excel_app.selection.Columns.AutoFit
    excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
   excel_sheet.Range("A1:aY4").Font.Bold = True
   db.Execute "drop table " & Mtblname & ""
   db.Execute "drop table " & Mtblname1 & ""
   ' active leaf pest ends here
   ' root pest starts here
   
   
   
   Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Root_Pest_Detail"
   
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




  
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 3) = ProperCase("GEWOG")
    excel_sheet.cells(3, 4) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 5) = ProperCase("FARMER NAME")
    MCOL = 6
    For m = 1 To mygrid1.rows - 1
    If Len(mygrid1.TextMatrix(m, 0)) = 0 Then Exit For
    If Len(mygrid1.TextMatrix(m, 1)) <> 0 Then
    excel_sheet.cells(3, MCOL) = mygrid1.TextMatrix(m, 0) & " % To " & mygrid1.TextMatrix(m, 1) & " %"
    If MCOL = 6 Or MCOL = 9 Or MCOL = 12 Then
    excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
    End If

    
    
    Else
      excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.cells(3, MCOL) = ">= Then " & mygrid1.TextMatrix(m, 0) & " %"
      End If
    MCOL = MCOL + 3
    Next
   i = 4
   MCOL = 5
   excel_sheet.cells(i, 6) = ProperCase("Last Visited")
   excel_sheet.cells(i, 7) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 8) = ProperCase("root Pest")
   excel_sheet.cells(i, 9) = ProperCase("FILD CODE")
   excel_sheet.cells(i, 10) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 11) = ProperCase("root Pest")
   excel_sheet.cells(i, 12) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 13) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 14) = ProperCase("root Pest")
   excel_sheet.cells(i, 15) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 16) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 17) = ProperCase("root Pest")
   excel_sheet.cells(i, 18) = ProperCase("FIELD CODE")
  i = 5
   n = i
   SLNO = 1
   MCOL = 5
Set rs = Nothing
rs.Open "select end, fdcode,(farmercode) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(rootpest) as mfieldname,count(fdcode) as cnt from " & Mtblname & "  group by farmercode,fdcode order by farmercode,fdcode,end", db

  Do While rs.EOF <> True
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
  fdno = 0
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
FindFA rs!farmercode, "F"
    excel_sheet.cells(i, 5) = rs!farmercode & " " & FAName
      excel_sheet.cells(i, 6) = "'" & rs!end
      
    fcriteria = (rs!mfieldname / rs!totaltrees) * 100
    If fcriteria >= Val(mygrid1.TextMatrix(1, 0)) And fcriteria <= Val(mygrid1.TextMatrix(1, 1)) Then
    MCOL = 7
     t5 = t5 + rs!totaltrees
      t6 = t6 + rs!mfieldname
      t7 = t7 + rs!cnt
      If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    ElseIf fcriteria >= Val(mygrid1.TextMatrix(2, 0)) And fcriteria <= Val(mygrid1.TextMatrix(2, 1)) Then
     t8 = t8 + rs!totaltrees
      t9 = t9 + rs!mfieldname
      t10 = t10 + rs!cnt
         If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 10
    ElseIf fcriteria >= Val(mygrid1.TextMatrix(3, 0)) And fcriteria <= Val(mygrid1.TextMatrix(3, 1)) Then
      t11 = t11 + rs!totaltrees
      t12 = t12 + rs!mfieldname
      t13 = t13 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 13
    Else
       t14 = t14 + rs!totaltrees
      t15 = t15 + rs!mfieldname
      t16 = t16 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 16
    End If
  
    excel_sheet.cells(i, MCOL) = rs!totaltrees
    excel_sheet.cells(i, MCOL + 1) = rs!mfieldname
    excel_sheet.cells(i, MCOL + 2) = rs!FDCODE
  

i = i + 1
mdgt = Mid(rs!dgt, 1, 9)
If rs.EOF Then Exit Do
rs.MoveNext

Loop
locstring = ""
locstring = md & " " & Dzname & " " & mg & " " & GEname & " " & mt & " " & TsName

    excel_sheet.cells(i, 7) = IIf(t5 <> 0, t5, "")
    excel_sheet.cells(i, 8) = IIf(t6 <> 0, t6, "")
    excel_sheet.cells(i, 9) = IIf(t7 <> 0, t7, "")
    excel_sheet.cells(i, 10) = IIf(t8 <> 0, t8, "")
    excel_sheet.cells(i, 11) = IIf(t9 <> 0, t9, "")
    excel_sheet.cells(i, 12) = IIf(t10 <> 0, t10, "")
    excel_sheet.cells(i, 13) = IIf(t11 <> 0, t11, "")
    excel_sheet.cells(i, 14) = IIf(t12 <> 0, t12, "")
    excel_sheet.cells(i, 15) = IIf(t13 <> 0, t13, "")
    excel_sheet.cells(i, 16) = IIf(t14 <> 0, t14, "")
    excel_sheet.cells(i, 17) = IIf(t15 <> 0, t15, "")
    excel_sheet.cells(i, 18) = IIf(t16 <> 0, t16, "")
    excel_sheet.Range(excel_sheet.cells(n, 1), _
                             excel_sheet.cells(i - 1, 1)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 1) = SLNO
    excel_sheet.Range(excel_sheet.cells(n, 2), _
                             excel_sheet.cells(i - 1, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 2) = locstring
                            excel_sheet.Range(excel_sheet.cells(i, 2), _
                             excel_sheet.cells(i, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(i, 2) = "TOTAL"
    excel_sheet.Range(excel_sheet.cells(i, 2), _
 excel_sheet.cells(i, 17)).Select
excel_app.selection.Font.Bold = True
SLNO = SLNO + 1
i = i + 1
n = i
fdno = 0
   Loop
  excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:r4").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
excel_sheet.Columns("A:aa").Select
 excel_app.selection.columnWidth = 15
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With

With excel_sheet.PageSetup
       
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB
End With
db.Execute "drop table " & Mtblname & ""

Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Root_Pest_Summary"

mchk = True

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

CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
              GetTbl
             SQLSTR = ""
SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate,n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  year(end),month(end) ,farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY year(end),month(end) , n.farmerbarcode, n.fdcode"
         
    db.Execute SQLSTR
    
    GetTbl1
   SQLSTR = ""
SQLSTR = "insert into " & Mtblname1 & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate, n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY  n.farmerbarcode, n.fdcode"

    db.Execute SQLSTR

                SQLSTR = "select COUNT(fdcode) as fcnt, SUBSTRING(farmercode,1,9) as id ,SUM(totaltrees) as tt,sum(rootpest) as jval,(sum(activepest)/sum(totaltrees)*100) as percent,year(end) as procyear,month(end) as procmonth  from " & Mtblname & "   where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),substring(farmercode,1,9) order by substring(farmercode,1,9),year(end),month(end)"
   For i = 1 To 13
    mtot(i) = 0
Next
Set rs = Nothing
rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    Screen.MousePointer = vbHourglass
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    excel_sheet.cells(2, 1) = ProperCase("MONTHLY SUMMARY OF  root Pest")
    excel_sheet.cells(3, 1) = ProperCase("DZONGKHAG  GEWOG  TSHOWOG")
     excel_sheet.cells(3, 2) = ProperCase("TOTAL TREES")
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 2
        excel_sheet.cells(3, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value)) & " " & " Root Pest"
        'excel_sheet.Cells(3, K + 1) = ProperCase("Active Leaf Pest")
        K = K + 2
    Next

    jrow = 4
    For ll = 3 To 50
    excel_sheet.cells(jrow, ll) = ProperCase("NO.OF FIELDS")
    excel_sheet.cells(jrow, ll + 1) = ProperCase("TREE NOS.")
    excel_sheet.cells(jrow, ll + 2) = ProperCase("No.")
    excel_sheet.cells(jrow, ll + 3) = ProperCase("%")
    ll = ll + 3
    Next
    excel_sheet.cells(jrow, ll) = "Root_Pest_Detail"
    Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       locstr = ""
       FindDZ Mid(rs!id, 1, 3)
       FindGE Mid(rs!id, 1, 3), Mid(rs!id, 4, 3)
       FindTs Mid(rs!id, 1, 3), Mid(rs!id, 4, 3), Mid(rs!id, 7, 3)
       locstr = Mid(rs!id, 1, 3) & " " & Dzname & " " & Mid(rs!id, 4, 3) & " " & GEname & " " & Mid(rs!id, 7, 3) & " " & TsName
       excel_sheet.cells(jrow, 1) = locstr
       Do While pYear = rs!id
    If rs!procmonth = 1 Then
    i = 3
    ElseIf rs!procmonth = 2 Then
    i = 7
    ElseIf rs!procmonth = 3 Then
    i = 11
    ElseIf rs!procmonth = 4 Then
    i = 15
    ElseIf rs!procmonth = 5 Then
    i = 19
    ElseIf rs!procmonth = 6 Then
    i = 23
    ElseIf rs!procmonth = 7 Then
    i = 27
    ElseIf rs!procmonth = 8 Then
    i = 31
    ElseIf rs!procmonth = 9 Then
    i = 35
    ElseIf rs!procmonth = 10 Then
    i = 39
    ElseIf rs!procmonth = 11 Then
    i = 43
    Else
    i = 47
    End If
Set muk = Nothing
muk.Open "select sum(totaltrees) as ttrees from " & Mtblname1 & " where substring(farmercode,1,9)='" & rs!id & "'", ODKDB
j = rs!jval
     excel_sheet.cells(jrow, 2) = muk!ttrees
           jtot = jtot + muk!ttrees
          excel_sheet.cells(jrow, i) = rs!FCNT
          excel_sheet.cells(jrow, i + 1) = rs!tt
           
           excel_sheet.cells(jrow, i + 2) = rs!jval
           excel_sheet.cells(jrow, i + 3) = (rs!jval / rs!tt)
           excel_sheet.cells(jrow, i + 3).NumberFormat = "0.00%"
          excel_sheet.cells(jrow, 51) = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
      
    Loop
    'jtot = 0
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, 51)).Select
    excel_app.selection.Columns.AutoFit
    excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
   excel_sheet.Range("A1:aY4").Font.Bold = True
   db.Execute "drop table " & Mtblname & ""
   db.Execute "drop table " & Mtblname1 & ""
' root pest ends here
'stem pest starts here
   Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Stem_Pest_Detail"
   
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





   
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 3) = ProperCase("GEWOG")
    excel_sheet.cells(3, 4) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 5) = ProperCase("FARMER NAME")
    MCOL = 6
    For m = 1 To mygrid1.rows - 1
    If Len(mygrid1.TextMatrix(m, 0)) = 0 Then Exit For
    If Len(mygrid1.TextMatrix(m, 1)) <> 0 Then
    excel_sheet.cells(3, MCOL) = mygrid1.TextMatrix(m, 0) & " % To " & mygrid1.TextMatrix(m, 1) & " %"
    If MCOL = 6 Or MCOL = 9 Or MCOL = 12 Then
    excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
    End If

    
    
    Else
      excel_sheet.Range(excel_sheet.cells(3, MCOL), _
                             excel_sheet.cells(3, MCOL + 2)).Select
                                 With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.cells(3, MCOL) = ">= Then " & mygrid1.TextMatrix(m, 0) & " %"
      End If
    MCOL = MCOL + 3
    Next
   i = 4
   MCOL = 5
   excel_sheet.cells(i, 6) = ProperCase("Last Visited")
   excel_sheet.cells(i, 7) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 8) = ProperCase("Stem Pest")
   excel_sheet.cells(i, 9) = ProperCase("FILD CODE")
   excel_sheet.cells(i, 10) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 11) = ProperCase("Stem Pest")
   excel_sheet.cells(i, 12) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 13) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 14) = ProperCase("Stem Pest")
   excel_sheet.cells(i, 15) = ProperCase("FIELD CODE")
   excel_sheet.cells(i, 16) = ProperCase("TOTAL TREES")
   excel_sheet.cells(i, 17) = ProperCase("Stem Pest")
   excel_sheet.cells(i, 18) = ProperCase("FIELD CODE")
  i = 5
   n = i
   SLNO = 1
   MCOL = 5
Set rs = Nothing
rs.Open "select end, fdcode,(farmercode) as dgt,farmercode,sum(totaltrees) as totaltrees,sum(stempest) as mfieldname,count(fdcode) as cnt from " & Mtblname & "  group by farmercode,fdcode order by farmercode,fdcode,end", db

  Do While rs.EOF <> True
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
  fdno = 0
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
FindFA rs!farmercode, "F"
    excel_sheet.cells(i, 5) = rs!farmercode & " " & FAName
      excel_sheet.cells(i, 6) = "'" & rs!end
      
    fcriteria = (rs!mfieldname / rs!totaltrees) * 100
    If fcriteria >= Val(mygrid1.TextMatrix(1, 0)) And fcriteria <= Val(mygrid1.TextMatrix(1, 1)) Then
    MCOL = 7
     t5 = t5 + rs!totaltrees
      t6 = t6 + rs!mfieldname
      t7 = t7 + rs!cnt
      If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    ElseIf fcriteria >= Val(mygrid1.TextMatrix(2, 0)) And fcriteria <= Val(mygrid1.TextMatrix(2, 1)) Then
     t8 = t8 + rs!totaltrees
      t9 = t9 + rs!mfieldname
      t10 = t10 + rs!cnt
         If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 10
    ElseIf fcriteria >= Val(mygrid1.TextMatrix(3, 0)) And fcriteria <= Val(mygrid1.TextMatrix(3, 1)) Then
      t11 = t11 + rs!totaltrees
      t12 = t12 + rs!mfieldname
      t13 = t13 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 13
    Else
       t14 = t14 + rs!totaltrees
      t15 = t15 + rs!mfieldname
      t16 = t16 + rs!cnt
        If rs!cnt > 0 Then
        fdno = fdno + 1
   End If
    MCOL = 16
    End If
  
    excel_sheet.cells(i, MCOL) = rs!totaltrees
    excel_sheet.cells(i, MCOL + 1) = rs!mfieldname
    excel_sheet.cells(i, MCOL + 2) = rs!FDCODE
  

i = i + 1
mdgt = Mid(rs!dgt, 1, 9)
If rs.EOF Then Exit Do
rs.MoveNext

Loop
locstring = ""
locstring = md & " " & Dzname & " " & mg & " " & GEname & " " & mt & " " & TsName

    excel_sheet.cells(i, 7) = IIf(t5 <> 0, t5, "")
    excel_sheet.cells(i, 8) = IIf(t6 <> 0, t6, "")
    excel_sheet.cells(i, 9) = IIf(t7 <> 0, t7, "")
    excel_sheet.cells(i, 10) = IIf(t8 <> 0, t8, "")
    excel_sheet.cells(i, 11) = IIf(t9 <> 0, t9, "")
    excel_sheet.cells(i, 12) = IIf(t10 <> 0, t10, "")
    excel_sheet.cells(i, 13) = IIf(t11 <> 0, t11, "")
    excel_sheet.cells(i, 14) = IIf(t12 <> 0, t12, "")
    excel_sheet.cells(i, 15) = IIf(t13 <> 0, t13, "")
    excel_sheet.cells(i, 16) = IIf(t14 <> 0, t14, "")
    excel_sheet.cells(i, 17) = IIf(t15 <> 0, t15, "")
    excel_sheet.cells(i, 18) = IIf(t16 <> 0, t16, "")
    excel_sheet.Range(excel_sheet.cells(n, 1), _
                             excel_sheet.cells(i - 1, 1)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 1) = SLNO
    excel_sheet.Range(excel_sheet.cells(n, 2), _
                             excel_sheet.cells(i - 1, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(n, 2) = locstring
                            excel_sheet.Range(excel_sheet.cells(i, 2), _
                             excel_sheet.cells(i, 4)).Select
                                With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(i, 2) = "TOTAL"
    excel_sheet.Range(excel_sheet.cells(i, 2), _
 excel_sheet.cells(i, 17)).Select
excel_app.selection.Font.Bold = True
SLNO = SLNO + 1
i = i + 1
n = i
fdno = 0
   Loop
  excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:r4").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
excel_sheet.Columns("A:aa").Select
 excel_app.selection.columnWidth = 15
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With

With excel_sheet.PageSetup
       
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB
End With
db.Execute "drop table " & Mtblname & ""

Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Stem_Pest_Summary"

mchk = True

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

CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
              GetTbl
             SQLSTR = ""
SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate,n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  year(end),month(end) ,farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY year(end),month(end) , n.farmerbarcode, n.fdcode"
         
    db.Execute SQLSTR
    
    GetTbl1
   SQLSTR = ""
SQLSTR = "insert into " & Mtblname1 & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select start,tdate, n.end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY  farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY  n.farmerbarcode, n.fdcode"

    db.Execute SQLSTR

                SQLSTR = "select COUNT(fdcode) as fcnt, SUBSTRING(farmercode,1,9) as id ,SUM(totaltrees) as tt,sum(rootpest) as jval,(sum(activepest)/sum(totaltrees)*100) as percent,year(end) as procyear,month(end) as procmonth  from " & Mtblname & "   where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),substring(farmercode,1,9) order by substring(farmercode,1,9),year(end),month(end)"
   For i = 1 To 13
    mtot(i) = 0
Next
Set rs = Nothing
rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    Screen.MousePointer = vbHourglass
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    excel_sheet.cells(2, 1) = ProperCase("MONTHLY SUMMARY OF  Stem Pest")
    excel_sheet.cells(3, 1) = ProperCase("DZONGKHAG  GEWOG  TSHOWOG")
     excel_sheet.cells(3, 2) = ProperCase("TOTAL TREES")
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 2
        excel_sheet.cells(3, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value)) & " " & " Stem Pest"
        'excel_sheet.Cells(3, K + 1) = ProperCase("Active Leaf Pest")
        K = K + 2
    Next

    jrow = 4
    For ll = 3 To 50
    excel_sheet.cells(jrow, ll) = ProperCase("NO.OF FIELDS")
    excel_sheet.cells(jrow, ll + 1) = ProperCase("TREE NOS.")
    excel_sheet.cells(jrow, ll + 2) = ProperCase("No.")
    excel_sheet.cells(jrow, ll + 3) = ProperCase("%")
    ll = ll + 3
    Next
    excel_sheet.cells(jrow, ll) = "Stem_Pest_Detail"
    Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       locstr = ""
       FindDZ Mid(rs!id, 1, 3)
       FindGE Mid(rs!id, 1, 3), Mid(rs!id, 4, 3)
       FindTs Mid(rs!id, 1, 3), Mid(rs!id, 4, 3), Mid(rs!id, 7, 3)
       locstr = Mid(rs!id, 1, 3) & " " & Dzname & " " & Mid(rs!id, 4, 3) & " " & GEname & " " & Mid(rs!id, 7, 3) & " " & TsName
       excel_sheet.cells(jrow, 1) = locstr
       Do While pYear = rs!id
    If rs!procmonth = 1 Then
    i = 3
    ElseIf rs!procmonth = 2 Then
    i = 7
    ElseIf rs!procmonth = 3 Then
    i = 11
    ElseIf rs!procmonth = 4 Then
    i = 15
    ElseIf rs!procmonth = 5 Then
    i = 19
    ElseIf rs!procmonth = 6 Then
    i = 23
    ElseIf rs!procmonth = 7 Then
    i = 27
    ElseIf rs!procmonth = 8 Then
    i = 31
    ElseIf rs!procmonth = 9 Then
    i = 35
    ElseIf rs!procmonth = 10 Then
    i = 39
    ElseIf rs!procmonth = 11 Then
    i = 43
    Else
    i = 47
    End If
Set muk = Nothing
muk.Open "select sum(totaltrees) as ttrees from " & Mtblname1 & " where substring(farmercode,1,9)='" & rs!id & "'", ODKDB
j = rs!jval
     excel_sheet.cells(jrow, 2) = muk!ttrees
           jtot = jtot + muk!ttrees
          excel_sheet.cells(jrow, i) = rs!FCNT
          excel_sheet.cells(jrow, i + 1) = rs!tt
           
           excel_sheet.cells(jrow, i + 2) = rs!jval
           excel_sheet.cells(jrow, i + 3) = (rs!jval / rs!tt)
           excel_sheet.cells(jrow, i + 3).NumberFormat = "0.00%"
          excel_sheet.cells(jrow, 51) = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(AY4&" & Chr(34) & "!B:B" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
      
    Loop
    'jtot = 0
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, 51)).Select
    excel_app.selection.Columns.AutoFit
    excel_sheet.cells(5, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
   excel_sheet.Range("A1:aY4").Font.Bold = True
   db.Execute "drop table " & Mtblname & ""
   db.Execute "drop table " & Mtblname1 & ""
   'stem pest ends here
   

Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

db.Close
Exit Sub
err:

MsgBox err.Description
err.Clear

End Sub

Private Sub Command18_Click()
allfieldpestdetail
End Sub

Private Sub Command19_Click()
updateField
updateStorage
updateDailyact
updatesiatribution
MsgBox "done"
End Sub

Private Sub Command20_Click()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsch As New ADODB.Recordset
Dim actstring As String
Dim mstaff As String
Dim tt As String
Dim SQLSTR As String
Dim sl As Integer
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
mstaff = ""
db.Open OdkCnnString
mchk = True
mchk = True
SQLSTR = ""
SLNO = 1
SQLSTR = "select * from dailyacthub9_core where  SUBSTRING( start ,1,10)>='" & Format(Now - 9, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now - 1, "yyyy-MM-dd") & "'  order by staffbarcode "
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
sl = 1
excel_sheet.name = "Daily Activity"
excel_app.Visible = True
excel_sheet.cells(3, 1) = "Sl.No."
excel_sheet.cells(3, 2) = "Start Date"
excel_sheet.cells(3, 3) = "Tdate"
excel_sheet.cells(3, 4) = ProperCase("End Date")
excel_sheet.cells(3, 5) = ProperCase("DZONGKHAG")
excel_sheet.cells(3, 6) = ProperCase("GEWOG")
excel_sheet.cells(3, 7) = ProperCase("TSHOWOG")
excel_sheet.cells(3, 8) = ProperCase("SURVEYOR ID")
excel_sheet.cells(3, 9) = ProperCase("NAME")
excel_sheet.cells(3, 10) = ProperCase("Yesterday's Activity")
excel_sheet.cells(3, 11) = ProperCase("No. of field visits")
excel_sheet.cells(3, 12) = ProperCase("No. of field failed")
excel_sheet.cells(3, 13) = ProperCase("Reason failed")
excel_sheet.cells(3, 14) = ProperCase("No. of storage visits")
excel_sheet.cells(3, 15) = ProperCase("No. of storage failed")
excel_sheet.cells(3, 16) = ProperCase("Reason failed")
excel_sheet.cells(3, 17) = ProperCase("Farmer registered")
excel_sheet.cells(3, 18) = ProperCase("Acre registered")
excel_sheet.cells(3, 19) = ProperCase("Travelling from")
excel_sheet.cells(3, 20) = ProperCase("Travelling to")
excel_sheet.cells(3, 21) = ProperCase("comments")
i = 4
Set rs = Nothing
rs.Open SQLSTR, db
        Do While rs.EOF <> True
                chkred = False
                excel_sheet.cells(i, 1) = SLNO
                excel_sheet.cells(i, 2) = "'" & rs!start
                excel_sheet.cells(i, 3) = "'" & rs!tdate
                excel_sheet.cells(i, 4) = "'" & rs!end
                excel_sheet.cells(i, 5) = ProperCase(ValidateLocationString(IIf(IsNull(rs!location_dcode), "", rs!location_dcode)))
                excel_sheet.cells(i, 6) = ProperCase(ValidateLocationString(IIf(IsNull(rs!location_gcode), "", rs!location_gcode)))
                excel_sheet.cells(i, 7) = ProperCase(ValidateLocationString(IIf(IsNull(rs!location), "", rs!location)))
                excel_sheet.cells(i, 8) = IIf(IsNull(rs!staffbarcode), "", rs!staffbarcode)
                FindsTAFF excel_sheet.cells(i, 8)
                excel_sheet.cells(i, 9) = sTAFF
                Dim chcnt As Integer
                Set rs1 = Nothing
                rs1.Open "select * from dailyacthub9_activities where _parent_auri='" & rs![_uri] & "' ", db
                actstring = ""
                    If rs1.EOF <> True Then
                                Do While rs1.EOF <> True
                                        Set rsch = Nothing
                                        rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
                                        If rsch.EOF <> True Then
                                            actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
                                        End If
                                        rs1.MoveNext
                                Loop
                                If Len(actstring) > 0 Then
                                    actstring = Left(actstring, Len(actstring) - 3)
                                    excel_sheet.cells(i, 10) = actstring
                                End If
                    Else
                                excel_sheet.cells(i, 10) = ""
                    End If
                'activity ends here
                excel_sheet.cells(i, 11) = IIf(IsNull(rs!field), "", rs!field)
                excel_sheet.cells(i, 12) = IIf(IsNull(rs!nofailed), "", rs!nofailed)
                'qc failed
                Set rs1 = Nothing
                rs1.Open "select * from dailyacthub9_qcfailed where _parent_auri='" & rs![_uri] & "' ", db
                actstring = ""
                    If rs1.EOF <> True Then
                        Do While rs1.EOF <> True
                            Set rsch = Nothing
                            rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
                                If rsch.EOF <> True Then
                                actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
                                End If
                            rs1.MoveNext
                        Loop
                        actstring = Left(actstring, Len(actstring) - 3)
                        excel_sheet.cells(i, 13) = actstring
                    Else
                        excel_sheet.cells(i, 13) = ""
                    End If
                ' qc failed ends here
                excel_sheet.cells(i, 14) = IIf(IsNull(rs!storage), "", rs!storage)
                excel_sheet.cells(i, 15) = IIf(IsNull(rs!nofailed1), "", rs!nofailed1)
                ' storage  failed
                Set rs1 = Nothing
                rs1.Open "select * from dailyacthub9_qcfailed1 where _parent_auri='" & rs![_uri] & "' ", db
                actstring = ""
                If rs1.EOF <> True Then
                Do While rs1.EOF <> True
                Set rsch = Nothing
                rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
                If rsch.EOF <> True Then
                actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
                End If
                rs1.MoveNext
                Loop
                actstring = Left(actstring, Len(actstring) - 3)
                excel_sheet.cells(i, 16) = actstring
                Else
                excel_sheet.cells(i, 16) = ""
                End If
                'storage failed ends here
                excel_sheet.cells(i, 17) = IIf(IsNull(rs!registered), "", rs!registered)
                excel_sheet.cells(i, 18) = IIf(IsNull(rs!privateland), "", rs!privateland)
                excel_sheet.cells(i, 19) = IIf(IsNull(rs!travel1), "", rs!travel1)
                excel_sheet.cells(i, 20) = IIf(IsNull(rs!travel2), "", rs!travel2)
                excel_sheet.cells(i, 21) = IIf(IsNull(rs!Comments), "", rs!Comments)
                SLNO = SLNO + 1
                i = i + 1
                rs.MoveNext
        Loop
'make up
excel_sheet.cells(4, 2).Select
excel_app.ActiveWindow.FreezePanes = True
excel_sheet.cells(1, 1).Select
With excel_sheet
excel_sheet.Range("A3:u3").Font.Bold = True
.PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
.PageSetup.CenterFooter = ProperCase("DAILY ACTIVITY")
.PageSetup.LeftFooter = "MHV"
.PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
.PageSetup.PrintGridlines = True
End With
excel_sheet.Columns("A:t").Select
excel_app.selection.columnWidth = 12
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
excel_sheet.Columns("u:u").Select
excel_app.selection.columnWidth = 80
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
With excel_sheet
    .PageSetup.Orientation = xlLandscape
End With
Dim PB As Integer
With excel_sheet.PageSetup
      PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
     .Zoom = False
     .FitToPagesWide = 1
     '.FitToPagesTall = PB
End With


' here ends daily act report
'starts monthly daily act report

Dim jrow As Long
mchk = True
Dim Jmth, K As Integer
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
For i = 1 To 13
    mtot(i) = 0
Next
Screen.MousePointer = vbHourglass
DoEvents

Excel_WBook.Sheets("sheet2").Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Detail"
jrow = 2
Set rs1 = Nothing
rs1.Open "SELECT DISTINCT staffbarcode FROM dailyacthub9_core", db
Do While rs1.EOF <> True
SQLSTR = "select value as id ,count(value) as jval,year(end) as procyear,month(end) as procmonth from dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and staffbarcode='" & rs1!staffbarcode & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),value order by convert(substring(value,9,2) ,unsigned integer),year(end),month(end)"
Set rs = Nothing
rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
FindsTAFF rs1!staffbarcode
excel_sheet.cells(jrow, 1) = ProperCase("ACTIVITY")
jrow = jrow + 1
excel_sheet.cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
K = 1
For i = Month(txtfrmdate) To Month(txttodate)
K = K + 1
excel_sheet.cells(jrow, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
Next
excel_sheet.cells(jrow, jCol) = ProperCase("Total")
excel_sheet.Range(excel_sheet.cells(jrow - 1, 1), _
excel_sheet.cells(jrow, 14)).Select
excel_app.selection.Font.Bold = True
jtot = 0
Do Until rs.EOF
jrow = jrow + 1
pYear = rs!id
findActivity Trim(rs!id)
excel_sheet.cells(jrow, 1) = rs!id & "  :  " & ActName
jtot = 0
j = 0
Do While pYear = rs!id
i = rs!procmonth + 2 - Month(txtfrmdate)
j = rs!jval
jtot = jtot + j
mtot(i - 1) = mtot(i - 1) + j
excel_sheet.cells(jrow, i) = Val(j)
rs.MoveNext
If rs.EOF Then Exit Do
Loop
excel_sheet.cells(jrow, jCol) = Val(jtot)
Loop
jtot = 0
excel_sheet.cells(jrow + 1, 1) = ProperCase("Total")
For i = 2 To jCol - 1
excel_sheet.cells(jrow + 1, i) = mtot(i - 1)
jtot = jtot + mtot(i - 1)
Next
excel_sheet.cells(jrow + 1, jCol) = Val(jtot)
excel_sheet.Range(excel_sheet.cells(jrow + 1, 1), _
excel_sheet.cells(jrow + 1, 14)).Select
excel_app.selection.Font.Bold = True
jtot = 0
For i = 2 To jCol - 1
       mtot(i - 1) = 0
Next
jrow = jrow + 2
rs1.MoveNext
Loop
'make up
excel_sheet.Range(excel_sheet.cells(3, 1), _
excel_sheet.cells(jrow + 1, jCol)).Select
excel_app.selection.Columns.AutoFit
excel_sheet.cells(4, 2).Select
excel_app.ActiveWindow.FreezePanes = True
excel_sheet.cells(1, 1).Select
With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
End With
Excel_WBook.Sheets("sheet3").Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Summary"
SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from  dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"
Set rs = Nothing
rs.Open SQLSTR, OdkCnnString
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
FindsTAFF rs!id
excel_sheet.cells(3, 1) = ProperCase("MONITOR")
K = 1
For i = Month(txtfrmdate) To Month(txttodate)
K = K + 1
excel_sheet.cells(3, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
Next
excel_sheet.cells(3, jCol) = ProperCase("Total")
excel_sheet.cells(3, jCol + 1) = ProperCase("Detail")
exactrow = 3
jrow = 3
Do Until rs.EOF
jrow = jrow + 1
pYear = rs!id
FindsTAFF Trim(rs!id)
excel_sheet.cells(jrow, 1) = rs!id & " " & sTAFF
jtot = 0
Do While pYear = rs!id
i = rs!procmonth + 2 - Month(txtfrmdate)
j = rs!jval
jtot = jtot + j
mtot(i - 1) = mtot(i - 1) + j
excel_sheet.cells(jrow, i) = Val(j)
rs.MoveNext
If rs.EOF Then Exit Do
exactrow = exactrow + 1
Loop
excel_sheet.cells(jrow, jCol) = Val(jtot)
tt = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(P3&" & "!A:A" & "),MATCH(" & "A" & jrow & ",INDIRECT(P3&" & "!A:A" & "),0)))," & "Link" & ")"
excel_sheet.cells(jrow, jCol + 1).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
exactrow = exactrow + 1
Loop
jtot = 0
excel_sheet.cells(jrow + 1, 1) = ProperCase("Total")
exactrow = exactrow + 1
For i = 2 To jCol - 1
excel_sheet.cells(jrow + 1, i) = mtot(i - 1)
jtot = jtot + mtot(i - 1)
Next
excel_sheet.cells(jrow + 1, jCol) = Val(jtot)
exactrow = exactrow + 1
'make up
excel_sheet.Range(excel_sheet.cells(3, 1), _
excel_sheet.cells(jrow + 1, jCol)).Select
excel_app.selection.Columns.AutoFit
excel_sheet.cells(4, 2).Select
excel_app.ActiveWindow.FreezePanes = True
excel_sheet.cells(1, 1).Select
With excel_sheet
    .PageSetup.LeftFooter = "mhv"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
End With
excel_sheet.Range("A1:o3").Font.Bold = True
' end dailyact summary
' starts field daily

Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(Excel_WBook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Daily Field"

Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim mdcode, mgcode, mtcode, mfcode As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
mchk = True
SQLSTR = ""
SLNO = 1
SQLSTR = ""
SQLSTR = "select _URI,start, tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,treeheight,comments1,other2,stems,management,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core "
         SQLSTR = SQLSTR & "where status<>'BAD' and substring(start,1,10)>='" & Format(Now - 9, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(Now - 1, "yyyy-MM-dd") & "' order by cast(substring(staffbarcode,3,3) as unsigned integer)"
sl = 1

    
    excel_sheet.cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.cells(3, 2) = ProperCase("START DATE")
    excel_sheet.cells(3, 3) = ProperCase("TDATE")
    excel_sheet.cells(3, 4) = ProperCase("END DATE")
    excel_sheet.cells(3, 5) = ProperCase("STAFF CODE-NAME")
    excel_sheet.cells(3, 6) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 7) = ProperCase("GEWOG")
    excel_sheet.cells(3, 8) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 9) = ProperCase("Farmer ID")
    excel_sheet.cells(3, 10) = ProperCase("average tree height")
    excel_sheet.cells(3, 11) = ProperCase("Field ID")
    excel_sheet.cells(3, 12) = ProperCase("Total Trees Distributed - Planted List")
    excel_sheet.cells(3, 13) = ProperCase("Total Trees")
    excel_sheet.cells(3, 14) = ProperCase("Good Moisture")
    excel_sheet.cells(3, 15) = ProperCase("Poor Moisture")
    excel_sheet.cells(3, 16) = ProperCase("Total Mositure Tally")
    excel_sheet.cells(3, 17) = ProperCase("Dead Missing")
    excel_sheet.cells(3, 18) = ProperCase("Slow Growing")
    excel_sheet.cells(3, 19) = ProperCase("Dormant")
    excel_sheet.cells(3, 20) = ProperCase("Active Growing")
    excel_sheet.cells(3, 21) = ProperCase("Shock")
    excel_sheet.cells(3, 22) = ProperCase("Nutrient Deficient")
    excel_sheet.cells(3, 23) = ProperCase("WaterLogged")
    excel_sheet.cells(3, 24) = ProperCase("average no. of stem")
    excel_sheet.cells(3, 25) = ProperCase("Active leaf Pest")
    excel_sheet.cells(3, 26) = ProperCase("Stem Pest")
    excel_sheet.cells(3, 27) = ProperCase("Root Pest")
    excel_sheet.cells(3, 28) = ProperCase("Animal Damage")
    excel_sheet.cells(3, 29) = ProperCase("Monitor's comments")
    excel_sheet.cells(3, 30) = ProperCase("follow up question")
    excel_sheet.cells(3, 31) = ProperCase("management")
    excel_sheet.cells(3, 32) = ProperCase("farmer's comments")
   i = 4
  Set rs = Nothing
  rs.Open SQLSTR, db
  Do While rs.EOF <> True
excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = "'" & rs!start
excel_sheet.cells(i, 3) = "'" & rs!tdate
excel_sheet.cells(i, 4) = "'" & rs!end  'rs.Fields(Mindex)
FindsTAFF rs!staffbarcode
excel_sheet.cells(i, 5) = rs!staffbarcode & " " & sTAFF
FindDZ Mid(rs!farmerbarcode, 1, 3)
excel_sheet.cells(i, 6) = Mid(rs!farmerbarcode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3)
excel_sheet.cells(i, 7) = Mid(rs!farmerbarcode, 4, 3) & " " & GEname
FindTs Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3), Mid(rs!farmerbarcode, 7, 3)
excel_sheet.cells(i, 8) = Mid(rs!farmerbarcode, 7, 3) & " " & TsName
FindFA IIf(IsNull(rs!farmerbarcode), "", rs!farmerbarcode), "F"
excel_sheet.cells(i, 9) = IIf(IsNull(rs!farmerbarcode), "", rs!farmerbarcode) & " " & FAName
excel_sheet.cells(i, 10) = IIf(IsNull(rs!treeheight), "", rs!treeheight)
excel_sheet.cells(i, 11) = IIf(IsNull(rs!FDCODE), "", rs!FDCODE)
Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as nooftrees from tblplanted where farmercode='" & rs!farmerbarcode & "' group by farmercode ", MHVDB
If rs1.EOF <> True Then
excel_sheet.cells(i, 12) = rs1!nooftrees
Else
excel_sheet.cells(i, 12) = ""
End If
excel_sheet.cells(i, 13) = IIf(IsNull(rs!tree_count_totaltrees), 0, rs!tree_count_totaltrees)
excel_sheet.cells(i, 14) = IIf(IsNull(rs!qc_tally_goodmoisture), "", rs!qc_tally_goodmoisture)
excel_sheet.cells(i, 15) = IIf(IsNull(rs!qc_tally_poormoisture), "", rs!qc_tally_poormoisture)
excel_sheet.cells(i, 16) = IIf(IsNull(rs!qc_tally_goodmoisture), "", rs!qc_tally_goodmoisture) + IIf(IsNull(rs!qc_tally_poormoisture), "", rs!qc_tally_poormoisture)
excel_sheet.cells(i, 17) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
excel_sheet.cells(i, 18) = IIf(IsNull(rs!tree_count_slowgrowing), "", rs!tree_count_slowgrowing)
excel_sheet.cells(i, 19) = IIf(IsNull(rs!tree_count_dor), "", rs!tree_count_dor)
excel_sheet.cells(i, 20) = IIf(IsNull(rs!tree_count_activegrowing), "", rs!tree_count_activegrowing)
excel_sheet.cells(i, 21) = IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.cells(i, 22) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.cells(i, 23) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.cells(i, 24) = IIf(IsNull(rs!stems), "", rs!stems)
excel_sheet.cells(i, 25) = IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.cells(i, 26) = IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.cells(i, 27) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.cells(i, 28) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.cells(i, 29) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)
If Mid(IIf(IsNull(rs!management), "", rs!management), 1, 1) = "y" Then
excel_sheet.cells(i, 30) = "Yes"
Else
excel_sheet.cells(i, 30) = "No"
End If
If UCase((IIf(IsNull(rs!management), "", rs!management))) = UCase("yes") Then
Set rs1 = Nothing
rs1.Open "select * from phealthhub15_management1 where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tblfieldchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
If UCase(rsch!label) = UCase("description9") Then
actstring = rs!other2
Else
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If
End If

rs1.MoveNext
Loop
If Len(actstring) > 0 Then
 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 31) = actstring
End If
Else

excel_sheet.cells(i, 31) = ""
End If
Else
excel_sheet.cells(i, 31) = ""
End If

excel_sheet.cells(i, 32) = IIf(IsNull(rs!comments1), "", rs!comments1)
SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    excel_sheet.Range("A3:ag3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
excel_sheet.Columns("A:ag").Select
excel_app.selection.columnWidth = 15
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
With excel_sheet
    .PageSetup.Orientation = xlLandscape
End With
With excel_sheet.PageSetup
      PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
     .Zoom = False
     .FitToPagesWide = 1
     '.FitToPagesTall = PB
End With

' daily fields ends here
' field summary starts here
Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Field Detail"
Dim fdcnt As Integer

mchk = True






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
               
db.Open OdkCnnString
db.Execute "delete from tempfarmernotinfield"
SQLSTR = " insert into tempfarmernotinfield(end,farmercode,fdcode,staffbarcode)" _
& "select '' as end,farmercode,'',monitor from MHV.tblplanted as a , MHV.tblfarmer as b  " _
& "where farmercode=idfarmer and farmercode not in (select farmerbarcode from  phealthhub15_core)"

db.Execute SQLSTR


fdcnt = 0
For i = 1 To 13
    mtot(i) = 0
Next
    
    excel_app.Caption = "mhv"
     jrow = 2
      excel_sheet.cells(jrow, 1) = ProperCase("FARMER")
      
     
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
    excel_sheet.cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
    excel_sheet.cells(jrow, 2) = ProperCase("LAST VISITED")
   excel_sheet.cells(jrow, 3) = ProperCase("FIELD CODE")
     
    K = 3
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.cells(jrow, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.cells(jrow, jCol) = ProperCase("Total")
    excel_sheet.Range(excel_sheet.cells(jrow - 1, 1), _
    excel_sheet.cells(jrow, 14)).Select
    excel_app.selection.Font.Bold = True
    jtot = 0
    fdcnt = 0
  
    
     Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       FindFA Trim(rs!farmerbarcode), "F"
       excel_sheet.cells(jrow, 1) = rs!farmerbarcode & " " & FAName
       jtot = 0
       j = 0
       
       Do While pYear = rs!id
          i = rs!procmonth + 4 - Month(txtfrmdate)
          
          j = IIf(rs!jval = "", 0, rs!jval)
          jtot = jtot + j
         
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.cells(jrow, 2) = "'" & rs!end
          excel_sheet.cells(jrow, 3) = CInt(rs!FDCODE)
          excel_sheet.cells(jrow, i) = Val(j)
         pYear = rs!id
           rs.MoveNext
         
          If rs.EOF Then Exit Do
          Loop
       
     
       excel_sheet.cells(jrow, jCol) = Val(jtot)
       If Val(jtot) > 0 Then
        fdcnt = fdcnt + 1
       End If
       'rs.MoveNext
       'jtot = 0
    Loop
   jtot = 0
    excel_sheet.cells(jrow + 1, 3) = fdcnt
    excel_sheet.cells(jrow + 1, 1) = UCase("Total")
    For i = 4 To jCol - 1
        excel_sheet.cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.cells(jrow + 1, jCol) = Val(jtot)
      excel_sheet.Range(excel_sheet.cells(jrow + 1, 1), _
    excel_sheet.cells(jrow + 1, 16)).Select
    excel_app.selection.Font.Bold = True
    jtot = 0
    
     For i = 2 To jCol - 1
       mtot(i - 1) = 0
       
    Next
    jrow = jrow + 2
    rs1.MoveNext
    Loop
    
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, jCol)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(3, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     'excel_sheet.Range("A1:Aa15").Font.Bold = True
    
   Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Field Summary"
SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from  phealthhub15_core where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"


Set rs = Nothing
rs.Open SQLSTR, OdkCnnString
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs!id
    'excel_sheet.Cells(2, 1) = "MONTHLY ACTIVITY OF MONITOR " & rs!id & " " & sTAFF
    excel_sheet.cells(3, 1) = ProperCase("MONITOR")
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.cells(3, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.cells(3, jCol) = ProperCase("Total")
    excel_sheet.cells(3, jCol + 1) = ("Field Detail")
    
    exactrow = 3
    jrow = 3
    Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       FindsTAFF Trim(rs!id)
       excel_sheet.cells(jrow, 1) = rs!id & " " & sTAFF
       jtot = 0
       Do While pYear = rs!id
          i = rs!procmonth + 2 - Month(txtfrmdate)
          j = rs!jval
          jtot = jtot + j
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.cells(jrow, i) = Val(j)
          rs.MoveNext
          If rs.EOF Then Exit Do
          exactrow = exactrow + 1
       Loop
       excel_sheet.cells(jrow, jCol) = Val(jtot)
       tt = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(P3&" & "!A:A" & "),MATCH(" & "A" & jrow & ",INDIRECT(P3&" & "!A:A" & "),0)))," & "Link" & ")"
       excel_sheet.cells(jrow, jCol + 1).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
       '
      
       exactrow = exactrow + 1
    Loop
    jtot = 0
    excel_sheet.cells(jrow + 1, 1) = ProperCase("Total")
    exactrow = exactrow + 1
    For i = 2 To jCol - 1
        excel_sheet.cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.cells(jrow + 1, jCol) = Val(jtot)
    exactrow = exactrow + 1
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, jCol)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    
     excel_sheet.Range("A1:o3").Font.Bold = True
     ' field summary ends here
     ' storage daily starts
     
     
     Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Daily Storage"




Set db = New ADODB.Connection
db.CursorLocation = adUseClient


mchk = True
db.Open OdkCnnString
                      



SQLSTR = ""


         SQLSTR = "select _uri,start,tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,totaltrees,other,gmoisture,pmoisture,gmoisture+pmoisture as totaltally," _
         & "dtrees,ndtrees,wlogged,pdamage,adamage,monitorcomments from storagehub6_core "
         SQLSTR = SQLSTR & "where status<>'BAD' and substring(start,1,10)>='" & Format(Now - 9, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(Now - 1, "yyyy-MM-dd") & "' order by staffbarcode"
               
                
                
                
                
         
 
  db.Execute SQLSTR

    sl = 1
 
    excel_sheet.cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.cells(3, 2) = ProperCase("start date")
    excel_sheet.cells(3, 3) = ProperCase("tdate")
    excel_sheet.cells(3, 4) = ProperCase("end date")
    excel_sheet.cells(3, 5) = ProperCase("STAFF CODE - NAME")
    excel_sheet.cells(3, 6) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 7) = ProperCase("GEWOG")
    excel_sheet.cells(3, 8) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 9) = ProperCase("Farmer code - name")
    excel_sheet.cells(3, 10) = ProperCase("storage condition")
    excel_sheet.cells(3, 11) = ProperCase("storage problem")
    excel_sheet.cells(3, 12) = ProperCase("action recommended")
    excel_sheet.cells(3, 13) = ProperCase("Total Trees Distributed - Planted List")
    excel_sheet.cells(3, 14) = ProperCase("Total Trees")
    excel_sheet.cells(3, 15) = ProperCase("Good Moisture")
    excel_sheet.cells(3, 16) = ProperCase("Poor Moisture")
    excel_sheet.cells(3, 17) = ProperCase("Total Mositure Tally")
    excel_sheet.cells(3, 18) = ProperCase("Dead Missing")
    'excel_sheet.Cells(3, 17) = ProperCase("Slow Growing")
    'excel_sheet.Cells(3, 18) = ProperCase("Dormant")
    'excel_sheet.Cells(3, 19) = ProperCase("Active Growing")
    'excel_sheet.Cells(3, 20) = ProperCase("Shock")
    excel_sheet.cells(3, 19) = ProperCase("Nutrient Deficient")
    excel_sheet.cells(3, 20) = ProperCase("Water Logg")
    excel_sheet.cells(3, 21) = ProperCase("pest damage")
    'excel_sheet.Cells(3, 20) = ProperCase("Active Pest")
    'excel_sheet.Cells(3, 21) = ProperCase("Stem Pest")
    'excel_sheet.Cells(3, 22) = ProperCase("Root Pest")
    excel_sheet.cells(3, 22) = ProperCase("Animal Damage")
    excel_sheet.cells(3, 23) = ProperCase("comments")
   i = 4
  Set rs = Nothing
  SLNO = 1
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
mchk = True
excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = "'" & rs!start
excel_sheet.cells(i, 3) = "'" & rs!tdate
excel_sheet.cells(i, 4) = "'" & rs!end
FindsTAFF rs!staffbarcode
excel_sheet.cells(i, 5) = rs!staffbarcode & " " & sTAFF

FindDZ Mid(rs!farmerbarcode, 1, 3)
excel_sheet.cells(i, 6) = Mid(rs!farmerbarcode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3)
excel_sheet.cells(i, 7) = Mid(rs!farmerbarcode, 4, 3) & " " & GEname
FindTs Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3), Mid(rs!farmerbarcode, 7, 3)
excel_sheet.cells(i, 8) = Mid(rs!farmerbarcode, 7, 3) & " " & TsName

FindFA rs!farmerbarcode, "F"
excel_sheet.cells(i, 9) = rs!farmerbarcode & " " & FAName




Set rs1 = Nothing
rs1.Open "select * from storagehub6_scondition where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tblstoragechoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If
rs1.MoveNext
Loop

If Len(actstring) > 0 Then
 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 10) = actstring
End If
Else

excel_sheet.cells(i, 10) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If



Set rs1 = Nothing
rs1.Open "select * from storagehub6_treeproblem where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tblstoragechoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If
rs1.MoveNext
Loop

If Len(actstring) > 0 Then
 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 11) = actstring
End If
Else

excel_sheet.cells(i, 11) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If


Set rs1 = Nothing
rs1.Open "select * from storagehub6_draction where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tblstoragechoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
If UCase(rsch!name) = UCase("action7") Then
actstring = rs!OTHER
Else
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If
End If
rs1.MoveNext
Loop

If Len(actstring) > 0 Then
 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 12) = actstring
End If
Else

excel_sheet.cells(i, 12) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If

Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as nooftrees from tblplanted where farmercode='" & rs!farmerbarcode & "' group by farmercode ", MHVDB
If rs1.EOF <> True Then
excel_sheet.cells(i, 13) = rs1!nooftrees

Else

excel_sheet.cells(i, 13) = ""
End If

excel_sheet.cells(i, 14) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.cells(i, 15) = IIf(IsNull(rs!gmoisture), "", rs!gmoisture)
excel_sheet.cells(i, 16) = IIf(IsNull(rs!pmoisture), "", rs!pmoisture)
excel_sheet.cells(i, 17) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.cells(i, 18) = IIf(IsNull(rs!dtrees), "", rs!dtrees)
'excel_sheet.Cells(i, 17) = "" 'IIf(IsNull(rs!slowgrowing), "", rs!slowgrowing)
'excel_sheet.Cells(i, 18) = "" 'IIf(IsNull(rs!dor), "", rs!dor)
'excel_sheet.Cells(i, 19) = "" 'IIf(IsNull(rs!activegrowing), "", rs!activegrowing)
'excel_sheet.Cells(i, 20) = "" 'IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.cells(i, 19) = IIf(IsNull(rs!ndtrees), "", rs!ndtrees)
excel_sheet.cells(i, 20) = IIf(IsNull(rs!wlogged), "", rs!wlogged)
excel_sheet.cells(i, 21) = IIf(IsNull(rs!pdamage), "", rs!pdamage)
'excel_sheet.Cells(i, 20) = "" 'IIf(IsNull(rs!activepest), "", rs!activepest)
'excel_sheet.Cells(i, 21) = "" 'IIf(IsNull(rs!stempest), "", rs!stempest)
'excel_sheet.Cells(i, 22) = "" 'IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.cells(i, 22) = IIf(IsNull(rs!adamage), "", rs!adamage)
excel_sheet.cells(i, 23) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)




SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up




'   excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(i, 15)).Select
'    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:aa3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL STORAGE"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
  



' excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(3, 15)).Select
excel_sheet.Columns("A:aa").Select
 excel_app.selection.columnWidth = 15
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With

db.Close

With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
'     .FitToPagesTall = PB

End With
' storage daily ends here
'storage detail/summary starts


Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Storage Detail"



mchk = True


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
              
db.Open OdkCnnString



db.Execute "delete from tempfarmernotinfield"
db.Execute " insert into tempfarmernotinfield(end,farmercode,staffbarcode)" _
           & "select distinct '' as end, farmerbarcode,staffbarcode from storagehub6_core " _
           & "where farmerbarcode not in (select farmerbarcode from phealthhub15_core) group by farmerbarcode"
           
SQLSTR = " delete from tempfarmernotinfield  where farmercode  in(" _
& "select farmercode from MHV.tblplanted as a , MHV.tblfarmer as b  " _
& "where farmercode=idfarmer)"

db.Execute SQLSTR


fdcnt = 0
For i = 1 To 13
    mtot(i) = 0
Next
    
    excel_app.Caption = "mhv"
   
     jrow = 2
      excel_sheet.cells(jrow, 1) = ProperCase("FARMER")
      
     
    Set rs1 = Nothing
    rs1.Open "SELECT DISTINCT staffbarcode FROM storagehub6_core", db
    Do While rs1.EOF <> True
    SQLSTR = ""
    SQLSTR = "select max(END) as end,farmerbarcode,farmerbarcode  as id ,count(farmerbarcode) as jval,year(end) as procyear,month(end) " _
    & " as procmonth from odk_prodlocal.storagehub6_core  where  staffbarcode='" & rs1!staffbarcode & "' and end between '2013-01-01' and '2013-12-31' group by year" _
    & " (end),month(end),farmerbarcode union SELECT  STR_TO_DATE('2013-01-01 14:15:16', '%d/%m/%Y') as END , farmercode, farmercode AS id, 0 AS jval," _
    & "  year('" & Format(DT1, "yyyy-MM-dd") & "') AS procyear,  month('" & Format(DT1, "yyyy-MM-dd") & "') as procmonth FROM tempfarmernotinfield  WHERE staffbarcode='" & rs1!staffbarcode & "'" _
    & " GROUP BY farmercode ORDER BY farmerbarcode, YEAR(END) , MONTH(END) "


Set rs = Nothing
    rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    
    jCol = 4 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs1!staffbarcode
     
       jrow = jrow + 1
    excel_sheet.cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
    excel_sheet.cells(jrow, 2) = ProperCase("LAST VISITED")
       
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.cells(jrow, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.cells(jrow, jCol) = ProperCase("Total")
    excel_sheet.Range(excel_sheet.cells(jrow - 1, 1), _
    excel_sheet.cells(jrow, 14)).Select
    excel_app.selection.Font.Bold = True
    jtot = 0
    fdcnt = 0
  
    
     Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       FindFA Trim(rs!farmerbarcode), "F"
       excel_sheet.cells(jrow, 1) = rs!farmerbarcode & " " & FAName
       jtot = 0
       j = 0
       
       Do While pYear = rs!id
          i = rs!procmonth + 3 - Month(txtfrmdate)
          
          j = IIf(rs!jval = "", 0, rs!jval)
          jtot = jtot + j
         
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.cells(jrow, 2) = "'" & rs!end
         
          
          'fdcnt = fdcnt + 1
          excel_sheet.cells(jrow, i) = IIf(Val(j) = 0, "", Val(j))
          

                        
                         
          
         pYear = rs!id
          rs.MoveNext
         
          If rs.EOF Then Exit Do
          'jrow = jrow + 1
       Loop
       
     
       excel_sheet.cells(jrow, jCol) = Val(jtot)
       If Val(jtot) = 0 Then
                            excel_sheet.Range(excel_sheet.cells(jrow, 1), _
                             excel_sheet.cells(jrow, 1)).Select
                             excel_app.selection.Interior.ColorIndex = 15
                             End If
       'rs.MoveNext
       'jtot = 0
    Loop
   jtot = 0
    'excel_sheet.Cells(jrow + 1, 3) = fdcnt
    excel_sheet.cells(jrow + 1, 1) = ProperCase("Total")
    For i = 3 To jCol - 1
        excel_sheet.cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.cells(jrow + 1, jCol) = Val(jtot)
      excel_sheet.Range(excel_sheet.cells(jrow + 1, 1), _
    excel_sheet.cells(jrow + 1, 16)).Select
    excel_app.selection.Font.Bold = True
    jtot = 0
    
     For i = 2 To jCol - 1
       mtot(i - 1) = 0
       
    Next
    jrow = jrow + 2
    rs1.MoveNext
    Loop
    
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, jCol)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(3, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     'excel_sheet.Range("A1:Aa15").Font.Bold = True
    
Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Storage Summary"
SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from  storagehub6_core where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"


Set rs = Nothing
rs.Open SQLSTR, OdkCnnString
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs!id
    'excel_sheet.Cells(2, 1) = "MONTHLY ACTIVITY OF MONITOR " & rs!id & " " & sTAFF
    excel_sheet.cells(3, 1) = ProperCase("MONITOR")
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.cells(3, jCol) = ProperCase("Total")
    excel_sheet.cells(3, jCol + 1) = ("Storage Detail")
    
    exactrow = 3
    jrow = 3
    Do Until rs.EOF
       jrow = jrow + 1
       pYear = rs!id
       FindsTAFF Trim(rs!id)
       excel_sheet.cells(jrow, 1) = rs!id & " " & sTAFF
       jtot = 0
       Do While pYear = rs!id
          i = rs!procmonth + 2 - Month(txtfrmdate)
          j = rs!jval
          jtot = jtot + j
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.cells(jrow, i) = Val(j)
          rs.MoveNext
          If rs.EOF Then Exit Do
          exactrow = exactrow + 1
       Loop
       excel_sheet.cells(jrow, jCol) = Val(jtot)
       tt = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(P3&" & "!A:A" & "),MATCH(" & "A" & jrow & ",INDIRECT(P3&" & "!A:A" & "),0)))," & "Link" & ")"
       excel_sheet.cells(jrow, jCol + 1).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),MATCH(" & "A" & jrow & ",INDIRECT(O3&" & Chr(34) & "!A:A" & Chr(34) & "),0)))," & Chr(34) & "Click here for detail" & Chr(34) & ")"
       '
      
       exactrow = exactrow + 1
    Loop
    jtot = 0
    excel_sheet.cells(jrow + 1, 1) = ProperCase("Total")
    exactrow = exactrow + 1
    For i = 2 To jCol - 1
        excel_sheet.cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.cells(jrow + 1, jCol) = Val(jtot)
    exactrow = exactrow + 1
    'make up
    excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(jrow + 1, jCol)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    
     excel_sheet.Range("A1:o3").Font.Bold = True

' ends storage detail/summary
' starts distribution





Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(Excel_WBook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Distribution"




Set db = New ADODB.Connection
db.CursorLocation = adUseClient
mstaff = ""
db.Open OdkCnnString
                        
Dim totfarmer, totturnedup, totdropout, totnewfarmers As Integer
Dim btypedistribyted, btypereturned, etypedistribyted, etypereturned, ptypedistribyted, ptypereturned As Double
Dim atypedistribyted, atypereturned, l1typedistribyted, l1typereturned, l2typedistribyted, l2typereturned As Double
Dim p1typedistribyted, p1typereturned, ntypedistribyted, ntypereturned, dnarea, droparea, newarea As Double



 totfarmer = 0
 totturnedup = 0
 totdropout = 0
 totnewfarmers = 0
 btypedistribyted = 0
 btypereturned = 0
 etypedistribyted = 0
 etypereturned = 0
 ptypedistribyted = 0
 ptypereturned = 0
 atypedistribyted = 0
 atypereturned = 0
 l1typedistribyted = 0
 l1typereturned = 0
 l2typedistribyted = 0
 l2typereturned = 0
 p1typedistribyted = 0
 p1typereturned = 0
 ntypedistribyted = 0
 ntypereturned = 0
 dnarea = 0
 droparea = 0
 newarea = 0

mchk = True
SQLSTR = ""
SLNO = 1


SQLSTR = "SELECT start,tdate,END , staffid,staffbarcode, staff_name, delivery_no, location_dcode," _
      & "location_gcode, location, totalfarmers_dist_list,actual_farmer_present, dropout, newfarmers," _
      & "dn_area,new_area,drop_area,b_type_distributed, b_type_returned, e_type_distributed," _
      & "e_type_returned, p_type_distributed,p_type_returned," _
      & "a_type_distributed, a_type_returned,p1_type_distributed," _
      & "p1_type_returned, n_type_distributed, n_type_returned," _
      & "l1_type_distributed, l1_type_returned, l2_type_distributed," _
      & "l2_type_returned, p1_type_distributed, p1_type_returned,monitorscomments" _
      & " FROM  distribution4_core where status<>'BAD' AND SUBSTRING(end,1,10)>='" & Format("01/06/2013", "yyyy-MM-dd") & "' and SUBSTRING(end ,1,10)<='" & Format(Now - 1, "yyyy-MM-dd") & "' ORDER BY end "




    sl = 1
    
   
    excel_sheet.cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.cells(3, 2) = ProperCase("START DATE")
     excel_sheet.cells(3, 3) = ProperCase("TDATE")
      excel_sheet.cells(3, 4) = ProperCase("END DATE")
    excel_sheet.cells(3, 5) = ProperCase("MONITOR ID")
    excel_sheet.cells(3, 6) = ProperCase("MONITOR NAME")
    excel_sheet.cells(3, 7) = ProperCase("DELIVERY NO.")
    excel_sheet.cells(3, 8) = ProperCase("DZONGKHAG")
    excel_sheet.cells(3, 9) = ProperCase("GEWOG")
    excel_sheet.cells(3, 10) = ProperCase("TSHOWOG")
    excel_sheet.cells(3, 11) = ProperCase("TOTAL FARMERS")
    excel_sheet.cells(3, 12) = ProperCase("TOTAL FARMERS TURNED UP")
    excel_sheet.cells(3, 13) = ProperCase("TOTAL ACRE")
    excel_sheet.cells(3, 14) = ProperCase("DROPOUT")
    excel_sheet.cells(3, 15) = ProperCase("ACRE OF DROPOUT")
    excel_sheet.cells(3, 16) = ProperCase("NEW FARMERS")
    excel_sheet.cells(3, 17) = ProperCase("ACRE OF NEW FARMERS")
    excel_sheet.cells(3, 18) = ProperCase("B- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 19) = ProperCase("B- TYPE RETURNED")
    excel_sheet.cells(3, 20) = ProperCase("E- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 21) = ProperCase("E- TYPE RETURNED")
    excel_sheet.cells(3, 22) = ProperCase("P- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 23) = ProperCase("P- TYPE RETURNED")
    excel_sheet.cells(3, 24) = ProperCase("A- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 25) = ProperCase("A- TYPE RETURNED")
    excel_sheet.cells(3, 26) = ProperCase("L1- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 27) = ProperCase("L1-TYPE RETURNED")
    excel_sheet.cells(3, 28) = ProperCase("L2- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 29) = ProperCase("L2- TYPE RETURNED")
    excel_sheet.cells(3, 30) = ProperCase("P1- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 31) = ProperCase("P1- TYPE RETURNED")
    excel_sheet.cells(3, 32) = ProperCase("N- TYPE DISTRIBUTED")
    excel_sheet.cells(3, 33) = ProperCase("N- TYPE RETURNED")
    excel_sheet.cells(3, 34) = ProperCase("MONITOR COMMENTS")
   i = 4
  Set rs = Nothing
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  
  excel_sheet.cells(i, 1) = sl
  excel_sheet.cells(i, 2) = "'" & rs!start
  excel_sheet.cells(i, 3) = "'" & rs!tdate
    excel_sheet.cells(i, 4) = "'" & rs!end
    excel_sheet.cells(i, 5) = rs!staffbarcode
    FindsTAFF rs!staffbarcode
    excel_sheet.cells(i, 6) = sTAFF
    excel_sheet.cells(i, 7) = rs!delivery_no
    
    excel_sheet.cells(i, 8) = ProperCase(ValidateLocationString(rs!location_dcode))
    excel_sheet.cells(i, 9) = ProperCase(ValidateLocationString(rs!location_gcode))
    excel_sheet.cells(i, 10) = ProperCase(ValidateLocationString(rs!location))
   
    excel_sheet.cells(i, 11) = rs!totalfarmers_dist_list
    excel_sheet.cells(i, 12) = rs!actual_farmer_present
    excel_sheet.cells(i, 13) = IIf(IsNull(rs!dn_area), "", rs!dn_area)
    excel_sheet.cells(i, 14) = IIf(rs!dropout = 0, "", rs!dropout)
    excel_sheet.cells(i, 15) = IIf(IsNull(rs!drop_area), "", rs!drop_area)
    excel_sheet.cells(i, 16) = IIf(rs!newfarmers = 0, "", rs!newfarmers)
    excel_sheet.cells(i, 17) = IIf(IsNull(rs!new_area), "", rs!new_area)
    excel_sheet.cells(i, 18) = IIf(rs!b_type_distributed = 0, "", rs!b_type_distributed)
    excel_sheet.cells(i, 19) = IIf(rs!b_type_returned = 0, "", rs!b_type_returned)
    excel_sheet.cells(i, 20) = IIf(rs!e_type_distributed = 0, "", rs!e_type_distributed)
    excel_sheet.cells(i, 21) = IIf(rs!e_type_returned = 0, "", rs!e_type_returned)
    excel_sheet.cells(i, 22) = IIf(rs!p_type_distributed = 0, "", rs!p_type_distributed)
    excel_sheet.cells(i, 23) = IIf(rs!p_type_returned = 0, "", rs!p_type_returned)
    excel_sheet.cells(i, 24) = IIf(rs!a_type_distributed = 0, "", rs!a_type_distributed)
    excel_sheet.cells(i, 25) = IIf(rs!a_type_returned = 0, "", rs!a_type_returned)
    excel_sheet.cells(i, 26) = IIf(rs!l1_type_distributed = 0, "", rs!l1_type_distributed)
    excel_sheet.cells(i, 27) = IIf(rs!l1_type_returned = 0, "", rs!l1_type_returned)
    excel_sheet.cells(i, 28) = IIf(rs!l2_type_distributed = 0, "", rs!l2_type_distributed)
    excel_sheet.cells(i, 29) = IIf(rs!l2_type_returned = 0, "", rs!l2_type_returned)
    excel_sheet.cells(i, 30) = IIf(rs!p1_type_distributed = 0, "", rs!p1_type_distributed)
    excel_sheet.cells(i, 31) = IIf(rs!p1_type_returned = 0, "", rs!p1_type_returned)
    excel_sheet.cells(i, 32) = IIf(rs!n_type_distributed = 0, "", rs!n_type_distributed)
    excel_sheet.cells(i, 33) = IIf(rs!n_type_returned = 0, "", rs!n_type_returned)
   excel_sheet.cells(i, 34) = rs!monitorscomments
  
  
  totfarmer = totfarmer + rs!totalfarmers_dist_list
 totturnedup = totturnedup + rs!actual_farmer_present
 totdropout = totdropout + IIf(rs!dropout = 0, 0, rs!dropout)
 totnewfarmers = totnewfarmers + IIf(rs!newfarmers = 0, 0, rs!newfarmers)
 btypedistribyted = btypedistribyted + IIf(rs!b_type_distributed = 0, 0, rs!b_type_distributed)
 btypereturned = btypereturned + IIf(rs!b_type_returned = 0, 0, rs!b_type_returned)
 etypedistribyted = etypedistribyted + IIf(rs!e_type_distributed = 0, 0, rs!e_type_distributed)
 etypereturned = etypereturned + IIf(rs!e_type_returned = 0, 0, rs!e_type_returned)
 ptypedistribyted = ptypedistribyted + IIf(rs!p_type_distributed = 0, 0, rs!p_type_distributed)
 ptypereturned = ptypereturned + IIf(rs!p_type_returned = 0, 0, rs!p_type_returned)
 atypedistribyted = atypedistribyted + IIf(rs!a_type_distributed = 0, 0, rs!a_type_distributed)
 atypereturned = atypereturned + IIf(rs!a_type_returned = 0, 0, rs!a_type_returned)
 l1typedistribyted = l1typedistribyted + IIf(rs!l1_type_distributed = 0, 0, rs!l1_type_distributed)
 l1typereturned = l1typereturned + IIf(rs!l1_type_returned = 0, 0, rs!l1_type_returned)
 l2typedistribyted = l2typedistribyted + IIf(rs!l2_type_distributed = 0, 0, rs!l2_type_distributed)
 l2typereturned = l2typereturned + IIf(rs!l2_type_returned = 0, 0, rs!l2_type_returned)
 p1typedistribyted = p1typedistribyted + IIf(rs!p1_type_distributed = 0, 0, rs!p1_type_distributed)
 p1typereturned = p1typereturned + IIf(rs!p1_type_returned = 0, 0, rs!p1_type_returned)
 ntypedistribyted = ntypedistribyted + IIf(rs!n_type_distributed = 0, 0, rs!n_type_distributed)
 ntypereturned = ntypereturned + IIf(rs!n_type_returned = 0, 0, rs!n_type_returned)
   dnarea = dnarea + IIf(IsNull(rs!dn_area), 0, rs!dn_area)
 droparea = droparea + IIf(IsNull(rs!drop_area), 0, rs!drop_area)
 newarea = newarea + IIf(IsNull(rs!new_area), 0, rs!new_area)
  
  
  sl = sl + 1
  i = i + 1
rs.MoveNext
   Loop


    excel_sheet.cells(i, 8) = "TOTAL"
   
    excel_sheet.cells(i, 11) = totfarmer
    excel_sheet.cells(i, 12) = totturnedup
    excel_sheet.cells(i, 13) = dnarea
    excel_sheet.cells(i, 14) = totdropout
    excel_sheet.cells(i, 15) = droparea
    excel_sheet.cells(i, 16) = totnewfarmers
    excel_sheet.cells(i, 17) = newarea
    excel_sheet.cells(i, 18) = btypedistribyted
    excel_sheet.cells(i, 19) = btypereturned
    excel_sheet.cells(i, 20) = etypedistribyted
    excel_sheet.cells(i, 21) = etypereturned
    excel_sheet.cells(i, 22) = ptypedistribyted
    excel_sheet.cells(i, 23) = ptypereturned
    excel_sheet.cells(i, 24) = atypedistribyted
    excel_sheet.cells(i, 25) = atypereturned
    excel_sheet.cells(i, 26) = l1typedistribyted
    excel_sheet.cells(i, 27) = l1typereturned
    excel_sheet.cells(i, 28) = l2typedistribyted
    excel_sheet.cells(i, 29) = l2typereturned
    excel_sheet.cells(i, 30) = p1typedistribyted
    excel_sheet.cells(i, 31) = p1typereturned
    excel_sheet.cells(i, 32) = ntypedistribyted
    excel_sheet.cells(i, 33) = ntypereturned
                             excel_sheet.Range(excel_sheet.cells(i, 2), _
                             excel_sheet.cells(i, 33)).Select
                             
                             excel_app.selection.Font.Bold = True




   'make up


    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:Ah3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = ProperCase("distribution report")
        .PageSetup.LeftFooter = ProperCase("MHV")
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("A:Ag").Select
 excel_app.selection.columnWidth = 12
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

excel_sheet.Columns("Ah:Ah").Select
 excel_app.selection.columnWidth = 60
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




 
'MsgBox CountOfBreaks


With excel_sheet.PageSetup
              PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With




' distribution ends here
' qc summary start here

Excel_WBook.Sheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
Excel_WBook.Sheets(ActiveWorkbook.Sheets.Count).Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "OM Orchard QC ODK"


Dim CrtStr As String
Dim totregland As Double

totregland = 0
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                    
'If OPTALL.Value = True Then
'Mindex = 51
'End If


SQLSTR = ""
SLNO = 1
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""

SQLSTR = ""
   
    
    GetTbl
        
    
SQLSTR = ""

           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
          
  db.Execute SQLSTR
  Set rss = Nothing
  Set rs1 = Nothing
  rs1.Open "select sum(regland) as regland from MHV.tbllandreg where farmerid in(select farmercode from " & Mtblname & ")", ODKDB
totregland = rs1!regland
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""


'On Error Resume Next






    sl = 1
    'excel_app.DisplayFullScreen = True
   
     excel_sheet.cells(1, 1) = ProperCase("Field")
     
    excel_sheet.cells(2, 1) = ProperCase("Total No. of hazelnut field")
    excel_sheet.cells(3, 1) = ProperCase("Total No. of trees in the field")
    excel_sheet.cells(4, 1) = ProperCase("Total acres")
    excel_sheet.cells(5, 1) = ""
    excel_sheet.cells(6, 1) = ProperCase("Slow growing")
    excel_sheet.cells(7, 1) = ProperCase("Dormant")
    excel_sheet.cells(8, 1) = ProperCase("Dead ")
    excel_sheet.cells(9, 1) = ProperCase("Active growing")
    excel_sheet.cells(10, 1) = ""
    excel_sheet.cells(11, 1) = ProperCase("Shock")
    excel_sheet.cells(12, 1) = ProperCase("Nutrient deficeint")
    excel_sheet.cells(13, 1) = ProperCase("Waterlog")
    excel_sheet.cells(14, 1) = ProperCase("Leafpest")
    excel_sheet.cells(15, 1) = ProperCase("Active pest")
    excel_sheet.cells(16, 1) = ProperCase("Stem pest")
    excel_sheet.cells(17, 1) = ProperCase("Root pest")
    excel_sheet.cells(18, 1) = ProperCase("Animal Damage")
    excel_sheet.cells(1, 2) = "All"
    excel_sheet.cells(1, 3) = "%"
    excel_sheet.cells(1, 4) = "Field"
    excel_sheet.cells(1, 5) = "%"
    excel_sheet.cells(1, 6) = "Storage All"
    excel_sheet.cells(1, 7) = "%"
    excel_sheet.cells(1, 8) = "Storage Only"
    excel_sheet.cells(1, 9) = "%"
    excel_sheet.cells(1, 10) = "Storage (but have some trees in field)"
     excel_sheet.cells(1, 11) = "%"
    
    
    Set rs = Nothing
    rs.Open SQLSTR, ODKDB
   Call fillcell(excel_sheet, 4, rs, Round(totregland, 0))
    
    
    
    SQLSTR = ""
ODKDB.Execute "delete from " & Mtblname & ""
           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,0,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,0,0,0,0,ndtrees,wlogged,0,0,0," _
         & "0,adamage,'0' from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"
          
  db.Execute SQLSTR
  Set rss = Nothing
  Set rs1 = Nothing
  totregland = 0
  rs1.Open "select sum(regland) as regland from MHV.tbllandreg where farmerid in(select farmercode from " & Mtblname & ")", ODKDB
totregland = IIf(IsNull(rs1!regland), 0, rs1!regland)
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""

 Set rs = Nothing
    rs.Open SQLSTR, ODKDB
    Call fillcell(excel_sheet, 6, rs, Round(totregland, 0))
    
    
    
    
    
  db.Execute "delete from tempfarmernotinfield"
db.Execute " insert into tempfarmernotinfield(end,farmercode,staffbarcode)" _
           & "select distinct '' as end, farmerbarcode,staffbarcode from storagehub6_core " _
           & "where farmerbarcode not in (select farmerbarcode from phealthhub15_core) group by farmerbarcode"
           
SQLSTR = " delete from tempfarmernotinfield  where farmercode  in(" _
& "select farmercode from MHV.tblplanted as a , MHV.tblfarmer as b  " _
& "where farmercode=idfarmer)"

db.Execute SQLSTR
    
    SQLSTR = ""
ODKDB.Execute "delete from " & Mtblname & ""
           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,0,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,0,0,0,0,ndtrees,wlogged,0,0,0," _
         & "0,adamage,'0' from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and n.farmerbarcode in(select farmercode from tempfarmernotinfield ) GROUP BY n.farmerbarcode"
          
  db.Execute SQLSTR
  Set rss = Nothing
  Set rs1 = Nothing
  totregland = 0
  rs1.Open "select sum(regland) as regland from MHV.tbllandreg where farmerid in(select farmercode from " & Mtblname & ")", ODKDB
totregland = rs1!regland
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""

 Set rs = Nothing
    rs.Open SQLSTR, ODKDB
    Call fillcell(excel_sheet, 8, rs, Round(totregland, 0))
      
  For i = 2 To 18
  If i <> 5 Or i <> 10 Then
  If i = 2 Or i = 4 Then
  excel_sheet.cells(i, 2) = excel_sheet.cells(i, 4) + excel_sheet.cells(i, 8)
  Else
    excel_sheet.cells(i, 2) = excel_sheet.cells(i, 4) + excel_sheet.cells(i, 6)
  End If
  excel_sheet.cells(i, 2).NumberFormat = "0"
  End If
  Next
 For i = 2 To 18
  If i <> 5 Or i <> 10 Then
  excel_sheet.cells(i, 10) = excel_sheet.cells(i, 6) - excel_sheet.cells(i, 8)
  excel_sheet.cells(i, 2).NumberFormat = "0"
  End If
  Next
For i = 6 To 18
  If i <> 10 Then
  excel_sheet.cells(i, 3) = (excel_sheet.cells(i, 2) / excel_sheet.cells(3, 2))
   excel_sheet.cells(i, 3).NumberFormat = "0.00%"
   excel_sheet.cells(i, 5) = (excel_sheet.cells(i, 4) / excel_sheet.cells(3, 4))
   excel_sheet.cells(i, 5).NumberFormat = "0.00%"
   excel_sheet.cells(i, 7) = (excel_sheet.cells(i, 6) / excel_sheet.cells(3, 6))
   excel_sheet.cells(i, 7).NumberFormat = "0.00%"
   excel_sheet.cells(i, 9) = (excel_sheet.cells(i, 8) / excel_sheet.cells(3, 8))
   excel_sheet.cells(i, 9).NumberFormat = "0.00%"
   excel_sheet.cells(i, 11) = (excel_sheet.cells(i, 10) / excel_sheet.cells(3, 10))
   excel_sheet.cells(i, 11).NumberFormat = "0.00%"
  End If
  Next


    


    With excel_sheet
    
     'excel_sheet.Range("a1:b15").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "Field and Storage Summary"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("A:a").Select
 excel_app.selection.columnWidth = 32
With excel_app.selection
.HorizontalAlignment = xlLeft
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

excel_sheet.Columns("b:k").Select
 excel_app.selection.columnWidth = 14
With excel_app.selection
.HorizontalAlignment = xlRight
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With





With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With

'excel_app.Visible = False


db.Execute "drop table " & Mtblname & ""
db.Close


' ends reporting here


'db.Close
Screen.MousePointer = vbDefault
Set excel_sheet = Nothing
Set excel_app = Nothing
End Sub
'Private Function fillcell(excel_sheet As Object, col As Integer, rs As Object, totregland As Double)
'excel_sheet.Cells(2, col) = rs!fieldcode
'    excel_sheet.Cells(3, col) = rs!totaltrees
'    excel_sheet.Cells(4, col) = totregland
'    excel_sheet.Cells(5, col) = ""
'    excel_sheet.Cells(6, col) = rs!slowgrowing
'    excel_sheet.Cells(7, col) = rs!dor
'    excel_sheet.Cells(8, col) = rs!dead
'    excel_sheet.Cells(9, col) = rs!activegrowing
'    excel_sheet.Cells(10, col) = ""
'    excel_sheet.Cells(11, col) = rs!shock
'    excel_sheet.Cells(12, col) = rs!nutrient
'    excel_sheet.Cells(13, col) = rs!waterlog
'    excel_sheet.Cells(14, col) = rs!leafpest
'    excel_sheet.Cells(15, col) = rs!activepest
'    excel_sheet.Cells(16, col) = rs!stempest
'    excel_sheet.Cells(17, col) = rs!rootpest
'    excel_sheet.Cells(18, col) = rs!animaldamage
'End Function


Private Sub Command21_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsplantbatchhdr", MHVDB
Do While rs.EOF <> True
Set rs1 = Nothing
rs1.Open "select max(convert(boxno,unsigned integer)) as max from tblqmsplantbatchdetail where trnid='" & rs!trnid & "'", MHVDB
MHVDB.Execute "update tblqmsplantbatchhdr set noofboxes='" & rs1!max & "' where trnid='" & rs!trnid & "'"


rs.MoveNext
Loop

End Sub

Private Sub Command22_Click()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmschemicalhdr", MHVDB
Do While rs.EOF <> True

MHVDB.Execute "update tblqmschemicalhdr set chemicalid='" & Right("00000" & rs!chemicalid, 3) & "' where chemicalid='" & rs!chemicalid & "'"
rs.MoveNext
Loop
MsgBox "yeah!"
End Sub

Private Sub Command23_Click()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsplantbatchdetail", MHVDB
Do While rs.EOF <> True

MHVDB.Execute "update tblqmsplantbatchdetail set plantvariety='" & Right("00000" & rs!plantvariety, 3) & "' where plantvariety='" & rs!plantvariety & "'"
rs.MoveNext
Loop
MsgBox "yeah!"
End Sub

Private Sub Command24_Click()
Dim rs As New ADODB.Recordset
Dim rsodk As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblodkalarmparameter where status='ON' and paraid not in('14')", ODKDB
Do While rs.EOF <> True
findtablename rs!odktable
Set rsodk = Nothing
rsodk.Open "select *," & rs!Formula & " as odkvalue from " & odkTableName & " where " & rs!Formula & " > '" & rs!Value & "' and substring(start,1,10)>='" & Format(rs!applicablefrom, "yyyy-MM-dd") & "' and status<>'BAD' ", ODKDB
Do While rsodk.EOF <> True
findfieldindes rs!odktable, odkTableName, rs!paraname
If rs!isstaffcode = 1 And rs!isfarmercode = 1 And rs!isfieldcode = 1 Then
updatefollowuplog rs!paraId, rsodk![_uri], odkTableName, rs!paraname, Now, rsodk!start, rsodk!odkValue, rsodk!staffbarcode, rsodk!farmerbarcode, rsodk!FDCODE
ElseIf rs!isstaffcode = 1 And rs!isfarmercode = 1 And rs!isfieldcode = 0 Then
updatefollowuplog rs!paraId, rsodk![_uri], odkTableName, rs!paraname, Now, rsodk!start, rsodk!odkValue, rsodk!staffbarcode, rsodk!farmerbarcode, 0
ElseIf rs!isstaffcode = 1 And rs!isfarmercode = 0 And rs!paraname = 0 Then
updatefollowuplog rs!paraId, rsodk![_uri], odkTableName, rs!paraname, Now, rsodk!start, rsodk!odkValue, rsodk!staffbarcode, 0, 0
ElseIf rs!isfarmercode = 0 And rs!isstaffcode = 0 And rs!isfieldcode = 0 Then
updatefollowuplog rs!paraId, rsodk![_uri], odkTableName, rs!paraname, Now, rsodk!start, rsodk!odkValue, 0, 0, 0
End If

rsodk.MoveNext
Loop
rs.MoveNext
Loop

End Sub

Private Sub Command25_Click()
'frmsendodkerroremail.Show 1
Dim bodymsg As String
Dim htmlbody As String
Dim tblheader As String
Dim htmlfoot As String
Dim tropen, trclose As String
emailMessageString = ""
Dim emailmsg As String
Dim recordno As Integer
Dim param As Integer
Dim rs As New ADODB.Recordset
mchk = True
Dim oSmtp As New EASendMailObjLib.Mail
  
    oSmtp.LicenseCode = "TryIt"
    oSmtp.FromAddr = "ODKerror@mhv.com"
    oSmtp.AddRecipientEx "muktitcc@gmail.com", 0
    
    oSmtp.ServerAddr = "smtp.tashicell.com"
    oSmtp.BodyFormat = 1
    tblheader = "<tr>" _
& "<th  bgcolor=""yellow"">S/N</th>" _
& "<th  bgcolor=""yellow"">Date</th>" _
& "<th  bgcolor=""yellow"">Dzongkhag</th>" _
& "<th  bgcolor=""yellow"">Gewog</th>" _
& "<th  bgcolor=""yellow"">Tshowog</th>" _
& "<th  bgcolor=""yellow"">Farmer Code</th>" _
& "<th  bgcolor=""yellow"">Monitor</th>" _
& "<th  bgcolor=""yellow"">Field/Storage</th>" _
& "</tr>"
tropen = "<tr>"
trclose = "</tr>"
  bodymsg = ProperCase("The List furnished below is the list of farmers having no registration record. Confirm from the concern monitor and rectify the records as soon as possible. Also update follow up log.")
'       htmlbody = "<tr>" _
'                & "<td>mukti</td><td>sikka</td><td>nirvana</td>" _
'                & "</tr>" _
'                & "</table>" _
'                 &
'
    
Set rs = Nothing
rs.Open "select * from tblodkfollowuplog where emailstatus='ON' and paraid in(14) order by paraid", ODKDB

    Do Until rs.EOF
       
       param = rs!paraId
       findParamDetails rs!paraId
       
      recordno = 0
       Do While param = rs!paraId
       FindsTAFF rs!staffcode
      
       FindDZ Mid(rs!farmercode, 1, 3)
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
       FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
      
      
          recordno = recordno + 1

            emailMessageString = "<td>" & recordno & "</td>"
            emailMessageString = emailMessageString & "<td>" & Format(rs!odkStartDate, "dd/MM/yyyy") & "</td>"
              emailMessageString = emailMessageString & "<td>" & Mid(rs!farmercode, 1, 3) & " " & Dzname & "</td>"
            emailMessageString = emailMessageString & "<td>" & Mid(rs!farmercode, 4, 3) & " " & GEname & "</td>"
            emailMessageString = emailMessageString & "<td>" & Mid(rs!farmercode, 7, 3) & " " & TsName & "</td>"
               emailMessageString = emailMessageString & "<td>" & rs!farmercode & "</td>"
                      
            emailMessageString = emailMessageString & "<td>" & rs!staffcode & " " & sTAFF & "</td>"
            If rs!fieldcode = 0 Then
            emailMessageString = emailMessageString & "<td>Field</td>"
             Else
                 emailMessageString = emailMessageString & "<td>Storage</td>"
             End If
         
           
            
            
            emailmsg = emailmsg & tropen & emailMessageString & trclose
            
            
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
        oSmtp.Subject = Format(Now, "yyyyMMdd") & " " & "ODK Threshold Exceeded On " & ProperCase(paramDesc) & "(" & ProperCase(paramFieldStorage) & ")"
        oSmtp.BodyText = "<html><head><title></title></head><body><br><h4>" & bodymsg & "</h4><br><TABLE border=""1"" cellspacing=""0"" >" & tblheader & emailmsg & "</TABLE></body></html>"
       '<h3>" & bodymsg & "</h3>
       If oSmtp.SendMail() = 0 Then
       ' send ok
      ODKDB.Execute "update tblodkfollowuplog set emailstatus='C' where emailstatus='ON' and paraid='" & param & "' "
      emailMessageString = ""
      emailmsg = ""
    Else
        'MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    emailMessageString = ""
    emailmsg = ""
    End If
Loop
End Sub

Private Sub Command26_Click()
Dim bodymsg As String
Dim htmlbody As String
Dim tblheader As String
Dim htmlfoot As String
Dim tropen, trclose As String
emailMessageString = ""
Dim emailmsg As String
Dim recordno As Integer
Dim param As Integer
Dim rs As New ADODB.Recordset
Dim oSmtp As New EASendMailObjLib.Mail
  
    oSmtp.LicenseCode = "TryIt"
    oSmtp.FromAddr = "ODKerror@mhv.com"
    oSmtp.AddRecipientEx "muktitcc@gmail.com", 0
    
    oSmtp.ServerAddr = "smtp.tashicell.com"
    oSmtp.BodyFormat = 1
    tblheader = "<tr>" _
& "<th  bgcolor=""yellow"">S/N</th>" _
& "<th  bgcolor=""yellow"">Date</th>" _
& "<th  bgcolor=""yellow"">Parameter of Concern</th>" _
& "<th  bgcolor=""yellow"">Threshold Value</th>" _
& "<th  bgcolor=""yellow"">ODK Value</th>" _
& "<th  bgcolor=""yellow"">Report Name</th>" _
& "<th  bgcolor=""yellow"">Monitor</th>" _
& "<th  bgcolor=""yellow"">Farmer</th>" _
& "<th  bgcolor=""yellow"">Dzongkhag</th>" _
& "<th bgcolor=""yellow"">Gewog </th>" _
& "<th  bgcolor=""yellow"">Tshowg </th>" _
& "<th  bgcolor=""yellow"">Field/Storage</th>" _
& "<th  bgcolor=""yellow"">Field Code</th>" _
& "</tr>"
tropen = "<tr>"
trclose = "</tr>"
bodymsg = "This Record Is Available In Follow Up Log.regularly Up Date The Follow Up Log."
    
'       htmlbody = "<tr>" _
'                & "<td>mukti</td><td>sikka</td><td>nirvana</td>" _
'                & "</tr>" _
'                & "</table>" _
'                 &
'
    
Set rs = Nothing
rs.Open "select * from tblodkfollowuplog where emailstatus='ON' and paraid not in(14) order by paraid", ODKDB

    Do Until rs.EOF
       
       param = rs!paraId
       findParamDetails rs!paraId
       
      recordno = 0
       Do While param = rs!paraId
       FindsTAFF rs!staffcode
       FindFA rs!farmercode, "F"
       FindDZ Mid(rs!farmercode, 1, 3)
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
       FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
      
      
          recordno = recordno + 1

            emailMessageString = "<td>" & recordno & "</td>"
            emailMessageString = emailMessageString & "<td>" & Format(rs!odkStartDate, "dd/MM/yyyy") & "</td>"
            emailMessageString = emailMessageString & "<td>" & paramName & "</td>"
            emailMessageString = emailMessageString & "<td>" & acceptableThresholdValue & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!odkValue & "</td>"
            emailMessageString = emailMessageString & "<td>" & ReportName & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!staffcode & " " & sTAFF & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!farmercode & " " & FAName & "</td>"
            emailMessageString = emailMessageString & "<td>" & Mid(rs!farmercode, 1, 3) & " " & Dzname & "</td>"
            emailMessageString = emailMessageString & "<td>" & Mid(rs!farmercode, 4, 3) & " " & GEname & "</td>"
            emailMessageString = emailMessageString & "<td>" & Mid(rs!farmercode, 7, 3) & " " & TsName & "</td>"
            emailMessageString = emailMessageString & "<td>" & paramFieldStorage & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!fieldcode & "</td>"
            
            
            emailmsg = emailmsg & tropen & emailMessageString & trclose
            
            
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
        oSmtp.Subject = Format(Now, "yyyyMMdd") & " " & "ODK Threshold Exceeded On " & ProperCase(paramDesc) & "(" & ProperCase(paramFieldStorage) & ")"
        oSmtp.BodyText = "<html><head><title></title></head><body><br><h5>" & bodymsg & "</h5><br><TABLE border=""1"" cellspacing=""0"" >" & tblheader & emailmsg & "</TABLE></body></html>"
       If oSmtp.SendMail() = 0 Then
       ' send ok
      ODKDB.Execute "update tblodkfollowuplog set emailstatus='C' where emailstatus='ON' and paraid='" & param & "' "
      emailMessageString = ""
      emailmsg = ""
    Else
        'MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    emailMessageString = ""
    emailmsg = ""
    End If
Loop
End Sub

Private Sub Command28_Click()
Dim sl As Integer
Dim i As Integer
Dim SQLSTR As String
Dim rsodk As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString

mchk = True
SQLSTR = ""
    
        GetTbl
    
          SQLSTR = "insert into " & Mtblname & " (_URI,odktable,end,dcode,gcode,tcode,farmercode,fstype,id,fname) select _uri,'55',n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,'0',staffbarcode,fname from phealthhub15_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
         db.Execute SQLSTR
    
         SQLSTR = "insert into " & Mtblname & " (_URI,odktable,end,dcode,gcode,tcode,farmercode,fstype,id,fname) select _uri,'59', n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,'1',staffbarcode,fname from storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
         db.Execute SQLSTR
   
    Set rsodk = Nothing
    SQLSTR = ""
    SQLSTR = "SELECT _URI,odktable,fstype,end,id,fname,farmercode,count(farmercode) as cnt FROM " & Mtblname & " WHERE farmercode not in(select idfarmer from MHV.tblfarmer) group by farmercode"
    rsodk.Open SQLSTR, db


    Do While rsodk.EOF <> True
    findtablename rsodk!odktable
    findfieldindes rsodk!odktable, odkTableName, "FARMERBARCODE"
    updatefollowuplog 14, rsodk![_uri], odkTableName, "FARMERBARCODE", Now, rsodk!end, 0, rsodk!id, rsodk!farmercode, rsodk!fstype
    rsodk.MoveNext
   Loop



db.Execute "drop table " & Mtblname & ""
db.Close
Exit Sub
err:
db.Execute "drop table " & Mtblname & ""
MsgBox err.Description
err.Clear
End Sub

Private Sub Command29_Click()
Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim i As Integer
Dim mydt As Date
i = 1
mydt = Format(Now, "dd/MM/yyyy")
intYear = CInt(Year(mydt))
intMonth = Month(mydt) + 1
intDay = i
For i = 1 To 7
If UCase(WDayName(DateSerial(intYear, intMonth, i), 2)) = "MON" Then

firstMondayOfTheMonth = DateSerial(intYear, intMonth, i)
Exit For
End If


Next

End Sub

Private Sub Command3_Click()

Dim newmatched, newnotmatched As Integer
Dim rsnew As New ADODB.Recordset
Dim newchk As Boolean
Dim fchk As Boolean
Dim schk As Boolean
mchk = True
chkred = True
Dim SQLSTR As String
Dim myphone As String
Dim TOTLAND As Double
Dim totadd As Double
TOTLAND = 0
totadd = 0
newmatched = 0
newnotmatched = 0
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Set rsadd = Nothing
'Dim sqlstr As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                      

    db.Execute "delete from tbltemp"

SQLSTR = ""
   SQLSTR = "insert into tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,fdcode) SELECT max(END), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode FROM phealthhub15_core where farmerbarcode<>'' group  by farmerbarcode,fdcode"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT max(END ) as end , dcode, gcode, tcode, fcode, farmerbarcode, (totaltrees),fdcode FROM phealthhub15_core WHERE farmerbarcode ='' GROUP BY dcode, gcode, tcode, fcode,fdcode", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  Set rsF = Nothing
  
  rsF.Open "select * from tbltemp where var6='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
  If rsF.EOF <> True Then
    
  If rsF!var1 > rss!end Then
  db.Execute "update tbltemp set var1='" & Format(rsF!var1, "yyyy-MM-dd") & "' , var7='" & rss!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,fdcode)values('" & Format(rss!end, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "') "
  
  End If
  
  rss.MoveNext
  Loop
  
  
  'storage
  SQLSTR = ""
   SQLSTR = "insert into tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,fdcode) SELECT max(END), dcode, gcode, tcode,fcode, scanlocation, (totaltrees),'S'  as fs,'' FROM storagehub6_core where scanlocation<>'' group  by scanlocation"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT max(END ) as end , dcode, gcode, tcode, fcode, scanlocation, (totaltrees)FROM storagehub6_core WHERE scanlocation ='' GROUP BY dcode, gcode, tcode, fcode", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  Set rsF = Nothing
  
  rsF.Open "select * from tbltemp where var6='" & mfcode & "'", db
  If rsF.EOF <> True Then
    
  If rsF!var1 > rss!end Then
  db.Execute "update tbltemp set var1='" & Format(rsF!var1, "yyyy-MM-dd") & "' , var7='" & rss!totaltrees & "' where var6='" & mfcode & "' and fs='S' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='S' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs)values('" & Format(rss!end, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','S') "
  
  End If
  
  rss.MoveNext
  Loop
                        
                        
                        
                        

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
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    i = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = "DZONGKHAG"
    excel_sheet.cells(3, 3) = "GEWOG"
    excel_sheet.cells(3, 4) = "TSHOWOG"
     excel_sheet.cells(3, 5) = "FARMER CODE"
    excel_sheet.cells(3, 6) = "FAMER"
    excel_sheet.cells(3, 7) = "REG. LAND (ACRE)"
    excel_sheet.cells(3, 8) = "PLANTED(ACRE)"
    excel_sheet.cells(3, 9) = "ACTUAL DISTRIBUTED"
    excel_sheet.cells(3, 10) = "TREES(FIELD)"
    excel_sheet.cells(3, 11) = "REES(STORAGE"
      i = 4
                        
                        
                        SQLSTR = ""
                    
                    SQLSTR = "select distinct farmercode from allfarmersexdropped where type='A'"
                        
                        
                            Set rs = Nothing
                            rs.Open SQLSTR, db
                            If rs.EOF <> True Then
                            Do While rs.EOF <> True
                            chkred = False
                            fchk = False
schk = False
                            excel_sheet.cells(i, 1) = sl
                            FindDZ Mid(rs!farmercode, 1, 3)
     excel_sheet.cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
     FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
   excel_sheet.cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
   FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
    excel_sheet.cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
  FindFA rs!farmercode, "F"
  
  
    excel_sheet.cells(i, 5) = rs!farmercode
  excel_sheet.cells(i, 6) = FAName
  
  
  
  Set rsF = Nothing
  rsF.Open "SELECT farmerid,sum(regland) as rl  from tbllandreg   farmerid where farmerid='" & rs!farmercode & "'", MHVDB
  If rsF.EOF <> True Then
  excel_sheet.cells(i, 7) = IIf(IsNull(rsF!rl), "", rsF!rl)
  
  End If
  
  
  
  
  Set rss = Nothing
  rss.Open "SELECT sum(acreplanted) as pl ,sum(nooftrees) as t FROM tblplanted where farmercode='" & rs!farmercode & "' GROUP BY farmercode", db

  If rss.EOF <> True Then
  
  excel_sheet.cells(i, 8) = rss!PL
  excel_sheet.cells(i, 9) = rss!T
  
  End If
  
  
 newchk = False
 fchk = False
  schk = False
  
  Set rss = Nothing
  rss.Open "SELECT max( var1 ) , var2, var3, var4, var5, var6, sum( var7 ) AS var7, fs, fdcode FROM tbltemp where fs='F' and var6='" & rs!farmercode & "' GROUP BY var6", db

  If rss.EOF <> True Then
  
    excel_sheet.cells(i, 10) = rss!var7
    newchk = True
    
    
    
    Else
    
    excel_sheet.cells(i, 10) = ""
'   excel_sheet.Range(excel_sheet.Cells(i, 1), _
'                             excel_sheet.Cells(i, 11)).Select
'                             excel_app.Selection.Font.Color = vbBlue

  fchk = True
    
  End If
  
 
   Set rss = Nothing
  rss.Open "SELECT max( var1 ) , var2, var3, var4, var5, var6, sum( var7 ) AS var7, fs, fdcode FROM tbltemp where fs='S' and var6='" & rs!farmercode & "' GROUP BY var6", db

  If rss.EOF <> True Then
  
    excel_sheet.cells(i, 11) = rss!var7
    newchk = True
  
    
    Else
     excel_sheet.cells(i, 11) = ""
'   excel_sheet.Range(excel_sheet.Cells(i, 1), _
'                             excel_sheet.Cells(i, 11)).Select
'                             excel_app.Selection.Font.Color = vbGreen
                             schk = True
      
  End If
  
  
  
  
  
  If newchk = True Then
    Set rsnew = Nothing
    rsnew.Open "select * from newfarmer where farmercode='" & rs!farmercode & "'", db
    If rsnew.EOF <> True Then
    newmatched = newmatched + 1
    Else
               Set rsnew = Nothing
    rsnew.Open "select * from newfarmer where farmercode='" & rs!farmercode & "'", db
    If rsnew.EOF <> True Then
    newnotmatched = newnotmatched + 1
    
    End If
    
    End If
    End If
  
  
  If fchk = True And schk = True Then
  
     excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 11)).Select
                             excel_app.selection.Font.Color = vbRed
                  
  End If
  
  
  
  
  fchk = False
  schk = False
  newchk = False
  
  
  

  
  
  
    i = i + 1
      sl = sl + 1
    
    
    
    
    
  
    
 
                            
                            rs.MoveNext
                            Loop
End If

 
                i = i + 1
                   excel_sheet.cells(i, 11) = "Match=" & newmatched & "Not Match=" & newnotmatched
                            'make up
   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 11)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:k3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "PLANTED LIST"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault

'excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

End Sub


Private Sub Command1_Click()

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                      
db.Open OdkCnnString
                      
Set rs = Nothing
rs.Open "select * from tbltable where tblid='8' ", db


Open "abc.htm" For Output As #3 ' creates file "abc.htm"
Print #3, rs.GetString(, 100, vbTab, "<br>", " "); 'prints selected records in the file
Close #3 'closes the file
WebBrowser1.Navigate (App.Path + "abc.htm")

End Sub

Private Sub createhtml()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim html As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                        
'If OPTALL.Value = True Then
'Exit Sub
'
'Else
'If CHKALLFIELD.Value = 0 Then
'If Len(CBOFARMER.Text) <> 0 Then
'rs.Open "select * from tempgoogle where substring(var1,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(var1,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and var7='" & Right(CBOMONITOR.BoundText, 3) & "' AND LAT<>0 AND LNG<>0 and var6='" & CBOFARMER.BoundText & "' and fdcode='" & CBOFDCODE.BoundText & "'", db
'Else
'rs.Open "select * from tempgoogle where substring(var1,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(var1,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and var7='" & Right(CBOMONITOR.BoundText, 3) & "' AND LAT<>0 AND LNG<>0 ", db
'End If
'Else
'rs.Open "select * from tempgoogle where substring(var1,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(var1,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and var7='" & Right(CBOMONITOR.BoundText, 3) & "' AND LAT<>0 AND LNG<>0 and var6='" & CBOFARMER.BoundText & "' ", db
'End If
'End If





' Build KML Feature
Dim FileNum As Integer
    FileNum = FreeFile
KmlFileName = App.Path & "\tempfile1.html"

    Open KmlFileName For Output As #FileNum

Print #FileNum, "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><title>QR code sheet</title></head><body><table border=""0"" cellspacing=""20"">"


Print #FileNum, "<tr><td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00038&size=150x150""><small><br>Y19G-ZZDK-2012-00038<br>3?1?(Y19-ZZDK-2012-00065</small></td>"
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00066&size=150x150"><small><br>Y19G-ZZDK-2012-00066<br>6?1?(Y19-ZZDK-2012-00103)</small></td>
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00010&size=150x150"><small><br>Y19G-ZZDK-2012-00010<br>4?2(Y19-ZZDK-2012-00014)</small></td>
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00020&size=150x150"><small><br>Y19G-ZZDK-2012-00020<br>6?4(Y19-ZZDK-2012-00024)</small></td>
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00041&size=150x150"><small><br>Y19G-ZZDK-2012-00041<br>3?2?(Y19-ZZDK-2012-00068</small></td>
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00034&size=150x150"><small><br>Y19G-ZZDK-2012-00034<br>2?1(Y19-ZZDK-2012-00057)</small></td>
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00021&size=150x150"><small><br>Y19G-ZZDK-2012-00021<br>7?1(Y19-ZZDK-2012-00025)</small></td>
'<td align="center"><img src="http://api.qrserver.com/v1/create-qr-code/?data=Y19G-ZZDK-2012-00004&size=150x150"><small><br>Y19G-ZZDK-2012-00004<br>1?4(Y19-ZZDK-2012-00004)</small></td>
Print #FileNum, "</tr>"

Print #FileNum, "</table></body></html>"


'Set objExplorer = New InternetExplorer
'objExplorer.Visible = True
'objExplorer.Navigate (App.Path + "tempfile1.html")

'Print #FileNum, tempfile1.html
'frmprac.webbrowser1.Navigate App.Path & "/tempfile1.html"
'WebBrowser1.Navigate (App.Path + "tempfile1.html")
'WebBrowser1.LocationURL = (App.Path + "tempfile1.html")

'Print #FileNum, "<Document>"
'Print #FileNum, "<name>" & CBOMONITOR.Text & "</name>"
''Print #FileNum, "<Placemark>"
''
''
''Print #FileNum, " <name>mukti</name>"
''Print #FileNum, "<description><![CDATA["
''
''Print #FileNum, " ]]>"
''Print #FileNum, " </description>"
''Print #FileNum, " <Point>"
''Print #FileNum, "    <coordinates>91.5920023400, 27.1888975300, 0</coordinates>"
''Print #FileNum, "  </Point>"
''Print #FileNum, "</Placemark>"
'If rs.EOF <> True Then
'mchk = True
'retVal = True
'Do While rs.EOF <> True
'Print #FileNum, " <Placemark>"
'Print #FileNum, "   <name> " & rs!var6 & "(" & rs!fdcode & ")" & " </name>"
'Print #FileNum, "  <description><![CDATA["
'     Print #FileNum, "    DATE VISITED: " & Format(rs!var1, "yyyy-MM-dd")
'     FindsTAFF "S0" & rs!var7
'     Print #FileNum, "    MONITOR: " & rs!var7 & " " & sTAFF
'     FindFA rs!var6, "F"
'     Print #FileNum, "    FARMER: " & FAName
'Print #FileNum, "    ]]>"
'Print #FileNum, " </description>"
'Print #FileNum, " <Point>"
'Print #FileNum, "   <coordinates>" & rs!LNG & "," & rs!LAT & " </coordinates>"
'
'
'
'Print #FileNum, " </Point>"
'Print #FileNum, " <TimeStamp>"
'Print #FileNum, "<when>" & Format(rs!var1, "yyyy-MM-ddThh:mm:ssZ") & "</when>"
'Print #FileNum, " </TimeStamp>"
'Set rschk = Nothing
'rschk.Open "select * from tblfarmer where idfarmer='" & rs!var6 & "'", MHVDB
'If rschk.EOF <> True Then
'Print #FileNum, " <Style id=""yellow"">"
' Print #FileNum, "  <IconStyle>"
'  Print #FileNum, "   <Icon>"
' Print #FileNum, "      <href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png </href>"
'   Print #FileNum, "  </Icon>"
'  Print #FileNum, " </IconStyle>"
'Print #FileNum, " </Style>"
'
'Else
'Print #FileNum, " <Style id=""red"">"
' Print #FileNum, "  <IconStyle>"
'  Print #FileNum, "   <Icon>"
' Print #FileNum, "      <href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png </href>"
'   Print #FileNum, "  </Icon>"
'  Print #FileNum, " </IconStyle>"
'Print #FileNum, " </Style>"
'
'End If
'
'Print #FileNum, " </Placemark>"
'rs.MoveNext
'Loop
'Else
'retVal = False
'End If
'mchk = False
''Print #FileNum, "<Placemark>"
''Print #FileNum, " <name>nirvana</name>"
''Print #FileNum, " <description>This is location 3</description>"
''Print #FileNum, "  <Point>"
''Print #FileNum, "   <coordinates>91.5920650800,27.1890155200</coordinates>"
''Print #FileNum, " </Point>"
''Print #FileNum, " </Placemark>"
'
'
'Print #FileNum, "</Document>"
'Print #FileNum, "</kml>"

Close #FileNum
    
End Sub

Private Sub Command30_Click()
 Dim oXL As Object        ' Excel application
    Dim oBook As Object      ' Excel workbook
    Dim oSheet As Object     ' Excel Worksheet
    Dim oChart As Object     ' Excel Chart
    
    Dim iRow As Integer      ' Index variable for the current Row
    Dim iCol As Integer      ' Index variable for the current Row
    
    Const cNumCols = 12      ' Number of points in each Series
    Const cNumRows = 4      ' Number of Series

    
    ReDim aTemp(1 To cNumRows, 1 To cNumCols)
    
    'Start Excel and create a new workbook
    Set oXL = CreateObject("Excel.application")
    Set oBook = oXL.Workbooks.Add
    Set oSheet = oBook.Worksheets.Item(1)
    
    ' Insert Random data into Cells for the two Series:
    Randomize Now()
    For iRow = 1 To cNumRows
       For iCol = 1 To cNumCols
          aTemp(iRow, (iCol)) = Int(Rnd * 50) + 1
       Next iCol
    Next iRow
    oSheet.Range("A1").Resize(cNumRows, cNumCols).Value = aTemp
    
    'Add a chart object to the first worksheet
    Set oChart = oSheet.ChartObjects.Add(50, 40, 300, 200).Chart
    oChart.SetSourceData Source:=oSheet.Range("A1").Resize(cNumRows, cNumCols)

    ' Make Excel Visible:
    oXL.Visible = True
'oChart.Legend.Clear
    oXL.UserControl = True
    oChart.ChartArea.Select
oChart.ChartArea.Copy
Image1.Picture = Clipboard.GetData(vbCFBitmap)

End Sub

Private Sub Command31_Click()
PivotTest
End Sub
Sub PivotTest()

Dim PTCache As PivotCache
Dim Table As PivotTable
Dim MyRange As Range
    
    xlWs.cells(2, 1).CopyFromRecordset rsXcl
    Set MyRange = ActiveSheet.UsedRange
    
    Set PTCache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:=MyRange)
    Set Table = PTCache.CreatePivotTable(TableDestination:="", TableName:="pt1")
    
With Table
    'Assign Pivot format here!
End With

End Sub

Private Sub Command32_Click()
 
createhtml1
End Sub
Private Sub createhtml1()


Dim sURL As String

sURL = "https://docs.google.com/spreadsheet/ccc?key=tDjrAUfxsOezCf3-ketogRA#gid=7"
Shell "C:\Program Files\Google\Chrome\Application\chrome.exe " & sURL, vbMaximizedFocus
End Sub

Private Sub Command33_Click()
createkml
End Sub
Private Sub createkml()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim SQLSTR As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                        
GetTbl
        
    
SQLSTR = ""

           
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area,gpslat,gpslng) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0',GPS_COORDINATES_LAT,GPS_COORDINATES_LNG from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
          
  db.Execute SQLSTR
  
  
SQLSTR = "select *  from   " & Mtblname & " where gpslat>0 and gpslng>0"
rs.Open SQLSTR, ODKDB


' Build KML Feature
Dim FileNum As Integer
    FileNum = FreeFile
KmlFileName = App.Path & "\tempfile.kml"

    Open KmlFileName For Output As #FileNum

Print #FileNum, "<kml xmlns=""http://earth.google.com/kml/2.0"">"
Print #FileNum, "<Document>"
Print #FileNum, "<name>"; "FIELD"; "</name>"
'Print #FileNum, "<Placemark>"
'
'
'Print #FileNum, " <name>mukti</name>"
'Print #FileNum, "<description><![CDATA["
'
'Print #FileNum, " ]]>"
'Print #FileNum, " </description>"
'Print #FileNum, " <Point>"
'Print #FileNum, "    <coordinates>91.5920023400, 27.1888975300, 0</coordinates>"
'Print #FileNum, "  </Point>"
'Print #FileNum, "</Placemark>"
If rs.EOF <> True Then
mchk = True
retVal = True
Do While rs.EOF <> True
Print #FileNum, " <Placemark>"
Print #FileNum, "   <name> " & rs!farmercode & " </name>"
Print #FileNum, "  <description><![CDATA["
     Print #FileNum, "    DATE VISITED: " & Format(rs!end, "yyyy-MM-dd")
     'FindsTAFF rs!staffbarcode
     'Print #FileNum, "    MONITOR: " & rs!staffbarcode & " " & sTAFF
     FindFA rs!farmercode, "F"
     Print #FileNum, "    FARMER: " & FAName
     Print #FileNum, "    FIELD: " & rs!FDCODE
Print #FileNum, "    ]]>"
Print #FileNum, " </description>"
Print #FileNum, " <Point>"
Print #FileNum, "   <coordinates>" & rs!gpslng & "," & rs!gpslat & " </coordinates>"



Print #FileNum, " </Point>"
Print #FileNum, " <TimeStamp>"
Print #FileNum, "<when>" & Format(rs!start, "yyyy-MM-ddThh:mm:ssZ") & "</when>"
Print #FileNum, " </TimeStamp>"
Set rschk = Nothing
rschk.Open "select * from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
If rschk.EOF <> True Then
Print #FileNum, " <Style id=""yellow"">"
 Print #FileNum, "  <IconStyle>"
  Print #FileNum, "   <Icon>"
 Print #FileNum, "      <href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png </href>"
   Print #FileNum, "  </Icon>"
  Print #FileNum, " </IconStyle>"
Print #FileNum, " </Style>"

Else
Print #FileNum, " <Style id=""red"">"
 Print #FileNum, "  <IconStyle>"
  Print #FileNum, "   <Icon>"
 Print #FileNum, "      <href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png </href>"
   Print #FileNum, "  </Icon>"
  Print #FileNum, " </IconStyle>"
Print #FileNum, " </Style>"

End If

Print #FileNum, " </Placemark>"
rs.MoveNext
Loop
Else
retVal = False
End If
mchk = False
'Print #FileNum, "<Placemark>"
'Print #FileNum, " <name>nirvana</name>"
'Print #FileNum, " <description>This is location 3</description>"
'Print #FileNum, "  <Point>"
'Print #FileNum, "   <coordinates>91.5920650800,27.1890155200</coordinates>"
'Print #FileNum, " </Point>"
'Print #FileNum, " </Placemark>"


Print #FileNum, "</Document>"
Print #FileNum, "</kml>"

Close #FileNum
    
End Sub

Private Sub Command34_Click()
Dim SLNO As Integer
Dim mFarmercode As String
Dim rs As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim mremarks As String
Dim visitcnt As Integer
Dim mfdcode As Integer
Dim loopcnt As Integer
Dim actstring As String
Dim mstaff As String
Dim tt As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
mstaff = ""
mFarmercode = ""
db.Open OdkCnnString
  mchk = True



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
   Dim sl As Integer
    sl = 1
 
    excel_app.Visible = False
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "Sl.No."
    excel_sheet.cells(3, 2) = "Start Date"
    excel_sheet.cells(3, 3) = "Tdate"
    excel_sheet.cells(3, 4) = ProperCase("End Date")
    excel_sheet.cells(3, 5) = ProperCase("SURVEYOR ID")
    excel_sheet.cells(3, 6) = ProperCase("NAME")
    excel_sheet.cells(3, 7) = "Field Visit(Daily Act)"
    excel_sheet.cells(3, 8) = "Field Visit(Field)"
    excel_sheet.cells(3, 9) = "Storage Visit(Daily Act)"
    excel_sheet.cells(3, 10) = "Storage Visit(Storage)"
    excel_sheet.cells(3, 11) = "Remarks"





Dim SQLSTR As String
mchk = True
SQLSTR = ""
SLNO = 1

SQLSTR = "select * from dailyacthub9_core where  (SUBSTRING( start ,1,10)>='" & Format(Now - 1 - backlogged, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "')  order by staffbarcode "
'(field>0 or storage>0) and
'SQLSTR = "select * from dailyacthub9_core where  SUBSTRING( start ,1,10)>='" & Format(Now - 90 - backlogged, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now - 89, "yyyy-MM-dd") & "'  order by staffbarcode "
'On Error Resume Next





'Dim RS As New ADODB.Recordset
i = 4
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
  mremarks = ""
  
  
  excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = "'" & rs!start
excel_sheet.cells(i, 3) = "'" & rs!tdate
excel_sheet.cells(i, 4) = "'" & rs!end
excel_sheet.cells(i, 5) = "'" & rs!staffbarcode
FindsTAFF rs!staffbarcode
excel_sheet.cells(i, 6) = sTAFF
excel_sheet.cells(i, 7) = IIf(IsNull(rs!field), "", rs!field)
excel_sheet.cells(i, 9) = IIf(IsNull(rs!storage), "", rs!storage)
Set rsF = Nothing
' field visit
visitcnt = 0
rsF.Open "select * from phealthhub15_core where staffbarcode='" & rs!staffbarcode & "'  and SUBSTRING( start ,1,10)>='" & Format(Now - 1 - backlogged, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' ", db
'rsF.Open "select * from phealthhub15_core where staffbarcode='" & rs!staffbarcode & "'  and SUBSTRING( start ,1,10)>='" & Format(Now - 90 - backlogged, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now - 89, "yyyy-MM-dd") & "' ", db
Do Until rsF.EOF
       mFarmercode = rsF!farmerbarcode
       mfdcode = rsF!FDCODE
       visitcnt = visitcnt + 1
       loopcnt = 0
       mremarks = ""
       
       Do While mFarmercode = rsF!farmerbarcode And mfdcode = rsF!FDCODE
            loopcnt = loopcnt + 1
            If loopcnt > 1 Then
            FindFA rsF!farmerbarcode, "F"
            mremarks = "Same farmer's record ( " & rsF!farmerbarcode & "  " & FAName & "  being sent twice." & " # " & mremarks
            End If
            
          rsF.MoveNext
          If rsF.EOF Then Exit Do
       Loop
      
    Loop
excel_sheet.cells(i, 8) = visitcnt
If Len(mremarks) > 0 Then
mremarks = Left(mremarks, Len(mremarks) - 3)
End If
excel_sheet.cells(i, 11) = mremarks

mremarks = ""

' storage visit
Set rsF = Nothing
visitcnt = 0
rsF.Open "select * from storagehub6_core where staffbarcode='" & rs!staffbarcode & "'  and SUBSTRING( start ,1,10)>='" & Format(Now - 1 - backlogged, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' ", db
'rsF.Open "select * from storagehub6_core where staffbarcode='" & rs!staffbarcode & "'  and SUBSTRING( start ,1,10)>='" & Format(Now - 90 - backlogged, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now - 89, "yyyy-MM-dd") & "' ", db
Do Until rsF.EOF
       mFarmercode = rsF!farmerbarcode
       visitcnt = visitcnt + 1
       loopcnt = 0
       mremarks = ""
       Do While mFarmercode = rsF!farmerbarcode
            loopcnt = loopcnt + 1
            
            If loopcnt > 1 Then
            FindFA rsF!farmerbarcode, "F"
            mremarks = "Same farmer's record ( " & rsF!farmerbarcode & "  " & FAName & "  being sent twice." & " # " & mremarks
            End If
            
          rsF.MoveNext
          If rsF.EOF Then Exit Do
       Loop
      
    Loop
excel_sheet.cells(i, 10) = visitcnt



If Len(mremarks) > 0 Then
mremarks = Left(mremarks, Len(mremarks) - 3)
excel_sheet.cells(i, 11) = excel_sheet.cells(i, 11) & "," & mremarks
End If


SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

 

   
 'make up


    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet

     excel_sheet.Range("A3:u3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = ProperCase("Daily Actity Vs Field and Storage Visit")
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("A:t").Select
 excel_app.selection.columnWidth = 15
'With excel_app.Selection
'
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'excel_sheet.Columns("u:u").Select
' excel_app.Selection.ColumnWidth = 80
With excel_app.selection

.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




 
'MsgBox CountOfBreaks

Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With
'

db.Close

'updateemaillog excel_app, receipient_id, nextemaildate, frequency
     
Screen.MousePointer = vbDefault
Set excel_sheet = Nothing
Set excel_app = Nothing






End Sub

Private Sub Command4_Click()
createhtml
frmbarcodeweb.Show 1
End Sub

Private Sub Command5_Click()
Dim SQLSTR As String
Dim monitor As Integer
Dim i, j, K, l, m As Integer
Dim c As Integer
Dim mycnt As Integer
Dim Y As Integer
Dim tempstr As String
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim mdcode, mgcode, mtcode, mfcode As String
Dim rst As New ADODB.Recordset
Dim inf As Boolean
j = 0
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
tempstr = ""
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                  

db.Execute "delete from tbltemp"

   SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,deadmissing,slowgrowing,dor,activegrowing,id) SELECT END, dcode, gcode, tcode,fcode, farmerbarcode, totaltrees,'F' as fs,fdcode,deadmissing,slowgrowing,dor,activegrowing,id FROM phealthhub15_core where farmerbarcode<>'' group by end,farmerbarcode,fdcode "
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT END , dcode, gcode, tcode, fcode, farmerbarcode, totaltrees,fdcode ,deadmissing,slowgrowing,dor,activegrowing,id FROM phealthhub15_core WHERE farmerbarcode =''group by end,dcode,gcode,tcode,fcode,fdcode ", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  
  
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,deadmissing,slowgrowing,dor,activegrowing,id)values('" & Format(rss!end, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "','" & rss!deadmissing & "','" & rss!slowgrowing & "','" & rss!dor & "','" & rss!activegrowing & "','" & rss!id & "') "
  
  
  
  rss.MoveNext
  Loop
  
  
  'storage
  SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,deadmissing,id) SELECT END, dcode, gcode, tcode,fcode, scanlocation, (totaltrees),'S'  as fs,'',dtrees,id FROM storagehub6_core where scanlocation<>'' group by end, scanlocation "
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT END , dcode, gcode, tcode, fcode, scanlocation, totaltrees,dtrees,id FROM storagehub6_core WHERE scanlocation ='' group by end,dcode,gcode,tcode", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  Set rsF = Nothing
    
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,deadmissing,id)values('" & Format(rss!end, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','S','" & rss!dtrees & "','" & rss!id & "') "
   
  rss.MoveNext
  Loop
  
  'fillin
  SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,fs,fdcode,btree1,etree1,ptree1,id) SELECT end,dcode,gcode,tcode,fcode,farmercode,'M',fdcode,btree1,etree1,ptree1,id FROM fillin where farmercode<>'' group by end,farmercode,fdcode "
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT end,dcode,gcode,tcode,fcode,farmercode,fdcode,btree1,etree1,ptree1,id FROM fillin WHERE farmercode ='' group by end,dcode,gcode,tcode,fcode", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(IIf(IsNull(rss!fcode), "", rss!fcode))
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  Set rsF = Nothing
    
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,fs,fdcode,btree1,etree1,ptree1,id)values('" & Format(rss!end, "yyyy-MM-dd") & "'," _
            & " '99','99','99','99','" & mfcode & "','M','" & rss!FDCODE & "','" & rss!btree1 & "','" & rss!etree1 & "','" & rss!ptree1 & "','" & rss!id & "') "
   
  rss.MoveNext
  Loop
  
  
  
  
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
   
Dim sl As Integer
sl = 1
i = 1
   
    excel_app.Visible = True
    
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = "MONITOR"
    excel_sheet.cells(3, 3) = "M. SUPERVISOR"
    excel_sheet.cells(3, 4) = "DZONGKHAG"
    excel_sheet.cells(3, 5) = "GEWOG"
    excel_sheet.cells(3, 6) = "TSHEWOG"
    
    excel_sheet.cells(3, 7) = "FARMER CODE"
    excel_sheet.cells(3, 8) = "2011"
    excel_sheet.cells(3, 9) = "2012"
    excel_sheet.cells(3, 10) = "TOTAL TREES"   ' NEW
    excel_sheet.cells(3, 11) = "FAMER NAME"
    excel_sheet.cells(3, 12) = "REG. LAND (ACRE)"
    excel_sheet.cells(3, 13) = "PRODUCTION" 'NEW
    excel_sheet.cells(3, 14) = "POLLEN" ' NEW
    excel_sheet.cells(3, 15) = "EST TOTAL)"  ' NEW
    excel_sheet.cells(3, 16) = "PLANTED(ACRE)"
    excel_sheet.cells(3, 17) = "DATE VISITED"
    excel_sheet.cells(3, 18) = "FIELD ID"
    excel_sheet.cells(3, 19) = "TOTAL TREES"
    excel_sheet.cells(3, 20) = "DEAD/MISSING"
    excel_sheet.cells(3, 21) = "ACTIVE TREES"
    excel_sheet.cells(3, 22) = "DATE VISIT"
    excel_sheet.cells(3, 23) = "TOTAL TREES"
    excel_sheet.cells(3, 24) = "DEAD/MISSING"
    excel_sheet.cells(3, 25) = "ACTIVE TREES"  'NEW
    excel_sheet.cells(3, 26) = "TOTAL ALIVE"  'NEW
    excel_sheet.cells(3, 27) = "DATE"
     excel_sheet.cells(3, 28) = "FIELD CODE"
    excel_sheet.cells(3, 29) = "E TYPE"
    excel_sheet.cells(3, 30) = "B TYPE"
    excel_sheet.cells(3, 31) = "P TYPE"
    excel_sheet.cells(3, 32) = "TOTAL FILLIN ESTIMATE"   'NEW
    
     excel_sheet.cells(3, 33) = "TOTAL AREA REGISTERED"   'NEW
      excel_sheet.cells(3, 34) = UCase("Total Area Planted to Date")   'NEW
       excel_sheet.cells(3, 35) = UCase("Registered area to plant this year")   'NEW
        excel_sheet.cells(3, 36) = UCase("Field Plants Alive")   'NEW
         excel_sheet.cells(3, 37) = UCase("Stored Plants Alive")   'NEW
      i = 4
     
      mchk = True
                        
    SQLSTR = ""
    SQLSTR = "select farmercode from allfarmersexdropped where type='A' order by farmercode"
    Set rs = Nothing
    rs.Open SQLSTR, db
   
    
    If rs.EOF <> True Then
    Do While rs.EOF <> True
   mycnt = i
    excel_sheet.cells(i, 1) = sl                        '"SL.NO."
    
    
    
    
     Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Mid
    excel_sheet.cells(i, 2) = rst!Mid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(i, 2) = ""
    End If
    
    
    
       Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Msupid
    excel_sheet.cells(i, 3) = rst!Msupid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(i, 3) = ""
    End If
    
    FindDZ Mid(rs!farmercode, 1, 3)
    
      excel_sheet.cells(i, 4) = Mid(rs!farmercode, 1, 3) & " " & Dzname
      
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
        excel_sheet.cells(i, 5) = Mid(rs!farmercode, 4, 3) & " " & GEname
        FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
          excel_sheet.cells(i, 6) = Mid(rs!farmercode, 7, 3) & " " & TsName
    
    
    excel_sheet.cells(i, 7) = rs!farmercode
    
    
    
    
    
    
    
    ' "FARMER CODE"
    Set rss = Nothing
    rss.Open "select sum(nooftrees) as t from tblplanted where farmercode='" & rs!farmercode & "' and year='2011' ", MHVDB
    If rss.EOF <> True Then
    excel_sheet.cells(i, 8) = IIf(IsNull(rss!T), "", rss!T)   '"REG. LAND (ACRE)"
    Else
    excel_sheet.cells(i, 8) = ""
    
    End If
    
    Set rss = Nothing
    rss.Open "select sum(nooftrees) as t from tblplanted where farmercode='" & rs!farmercode & "' and year='2012' ", MHVDB
    If rss.EOF <> True Then
    excel_sheet.cells(i, 9) = IIf(IsNull(rss!T), "", rss!T)   '"REG. LAND (ACRE)"
    Else
    excel_sheet.cells(i, 9) = ""
    
    End If
    
    'excel_sheet.Cells(i, 4) = ""                        '"2011"
    ' excel_sheet.Cells(i, 5) = ""                       ' "2012"
     FindFA rs!farmercode, "F"
    excel_sheet.cells(i, 10) = excel_sheet.cells(i, 8) + excel_sheet.cells(i, 9) 'FAName                    '"FAMER NAME"
    excel_sheet.cells(i, 11) = FAName
    Set rss = Nothing
    rss.Open "select sum(regland) as rl from tbllandreg where farmerid='" & rs!farmercode & "'", MHVDB
    If rss.EOF <> True Then
    excel_sheet.cells(i, 12) = IIf(IsNull(rss!rl), "", rss!rl)   '"REG. LAND (ACRE)"
    Else
    excel_sheet.cells(i, 12) = ""
    
    End If
    
    
    excel_sheet.cells(i, 13) = excel_sheet.cells(i, 12) * (450 * 0.84)                    '"PLANTED(ACRE)"
    excel_sheet.cells(i, 14) = excel_sheet.cells(i, 12) * 450 * 0.16
    excel_sheet.cells(i, 15) = excel_sheet.cells(i, 12) * 450
    
    excel_sheet.cells(i, 16) = ""
'    If rs!farmercode = "nocode1" Then
'    MsgBox "sdjb"
'    End If

   
    
    
    inf = False
    K = i
      Set rss = Nothing
      rss.Open "select * from tbltemp where  fs='F' and farmercode='" & rs!farmercode & "' order by fdcode,end desc limit 3", db
    
    If rss.EOF <> True Then
    Do While rss.EOF <> True
    inf = True
    Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Mid
    excel_sheet.cells(K, 2) = rst!Mid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(K, 2) = ""
    End If
    
    
    
       Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Msupid
    excel_sheet.cells(K, 3) = rst!Msupid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(K, 3) = ""
    End If
    
    FindDZ Mid(rs!farmercode, 1, 3)
    
      excel_sheet.cells(K, 4) = Mid(rs!farmercode, 1, 3) & " " & Dzname
      
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
        excel_sheet.cells(K, 5) = Mid(rs!farmercode, 4, 3) & " " & GEname
        FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
          excel_sheet.cells(K, 6) = Mid(rs!farmercode, 7, 3) & " " & TsName
    
    
    excel_sheet.cells(K, 7) = rs!farmercode
    
    
    
    
    
    
    
    
    
    
    
     excel_sheet.cells(K, 17) = "'" & rss!end '"DATE VISITED"
     excel_sheet.cells(K, 18) = rss!FDCODE
     excel_sheet.cells(K, 19) = rss!totaltrees '"TOTAL TREES"
     excel_sheet.cells(K, 20) = rss!deadmissing ' "dead missing"
     excel_sheet.cells(K, 21) = rss!slowgrowing + rss!dor + rss!activegrowing ' "ACTIVE TREES"
     excel_sheet.cells(K, 26).Formula = "=U" & K & "+Q" & K
     K = K + 1
     rss.MoveNext
    Loop
    
    Else
        
    excel_sheet.cells(K, 17) = "" ' "DATE VISITED"
    excel_sheet.cells(K, 18) = ""
    excel_sheet.cells(K, 19) = "" '"TOTAL TREES"
    excel_sheet.cells(K, 20) = "" '"dead missing"
    excel_sheet.cells(K, 21) = "" '"ACTIVE TREES"
    End If
    
   
    
    
    
    
    
    
    
    
    
    
    
    
    
    l = i
    Set rss = Nothing
    rss.Open "select * from tbltemp where  fs='S' and farmercode='" & rs!farmercode & "' order by end desc limit 3", db
    If rss.EOF <> True Then
    Do While rss.EOF <> True
    
    
    If inf = False Then
    
    Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Mid
    excel_sheet.cells(l, 2) = rst!Mid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(l, 2) = ""
    End If
    
    
    
       Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Msupid
    excel_sheet.cells(l, 3) = rst!Msupid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(l, 3) = ""
    End If
    
    FindDZ Mid(rs!farmercode, l, 3)
    
      excel_sheet.cells(l, 4) = Mid(rs!farmercode, 1, 3) & " " & Dzname
      
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
        excel_sheet.cells(l, 5) = Mid(rs!farmercode, 4, 3) & " " & GEname
        FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
          excel_sheet.cells(l, 6) = Mid(rs!farmercode, 7, 3) & " " & TsName
    
    
    excel_sheet.cells(l, 7) = rs!farmercode
    
    
    
    End If
    
    
    
    
     excel_sheet.cells(l, 22) = "'" & rss!end   '"DATE VISITED"
     excel_sheet.cells(l, 23) = rss!totaltrees  '"TOTAL TREES"
     excel_sheet.cells(l, 24) = rss!deadmissing ' "ACTIVE TREES"
     excel_sheet.cells(l, 24) = excel_sheet.cells(l, 19) - excel_sheet.cells(l, 20)
      excel_sheet.cells(l, 26).Formula = "=U" & l & "+Q" & l
     l = l + 1
     rss.MoveNext
    Loop
    
    Else
        
    excel_sheet.cells(l, 22) = "" ' "DATE VISITED"
    excel_sheet.cells(l, 23) = "" '"TOTAL TREES"
    excel_sheet.cells(l, 24) = "" '"ACTIVE TREES"
    excel_sheet.cells(l, 25) = ""
  
    End If
    'excel_sheet.Cells(l, 22) = Val(excel_sheet.Cells(l, 21)) + Val(excel_sheet.Cells(l, 17))
   'excel_sheet.Cells(i, 22).Formula = "=ROUND(U" & i & "*0.6,0)"
    
    m = i
    
    Set rss = Nothing
    rss.Open "select * from tbltemp where  fs='M' and farmercode='" & rs!farmercode & "' order by end desc limit 3", db
    If rss.EOF <> True Then
    Do While rss.EOF <> True
    
    
    
    If inf = False Then
    
    Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Mid
    excel_sheet.cells(m, 2) = rst!Mid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(m, 2) = ""
    End If
    
    
    
       Set rst = Nothing
    rst.Open "select * from mastreg  where farmercode='" & rs!farmercode & "'", db
    If rst.EOF <> True Then
    FindsTAFF rst!Msupid
    excel_sheet.cells(m, 3) = rst!Msupid & " " & sTAFF                        ' "MONITOR"
    Else
     excel_sheet.cells(m, 3) = ""
    End If
    
    FindDZ Mid(rs!farmercode, m, 3)
    
      excel_sheet.cells(m, 4) = Mid(rs!farmercode, 1, 3) & " " & Dzname
      
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
        excel_sheet.cells(m, 5) = Mid(rs!farmercode, 4, 3) & " " & GEname
        FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
          excel_sheet.cells(m, 6) = Mid(rs!farmercode, 7, 3) & " " & TsName
    
    
    excel_sheet.cells(m, 7) = rs!farmercode
    
    
    
    End If
    
    
    
    
    
     excel_sheet.cells(m, 27) = "'" & rss!end ' "DATE"
     excel_sheet.cells(m, 28) = rss!FDCODE
    excel_sheet.cells(m, 29) = rss!btree1 '"E TYPE"
     '"E TYPE"
    excel_sheet.cells(m, 30) = rss!etree1 ' "B TYPE"
    excel_sheet.cells(m, 31) = rss!ptree1 '"P TYPE"
    excel_sheet.cells(m, 32) = Val(excel_sheet.cells(m, 29)) + Val(excel_sheet.cells(m, 30)) + Val(excel_sheet.cells(m, 31))
     m = m + 1
     rss.MoveNext
    Loop
    
    Else
        
    excel_sheet.cells(m, 27) = "" ' "DATE"
     excel_sheet.cells(m, 28) = "" 'rss!FDCODE
    excel_sheet.cells(m, 29) = "" '"E TYPE"
    excel_sheet.cells(m, 30) = "" ' "B TYPE"
    excel_sheet.cells(m, 31) = "" '"P TYPE"
     excel_sheet.cells(m, 32) = ""
    End If
    

    
    
    max K, l, m
    
    
    
    i = mymax
    If (K - l) - (m - i) = 0 Then
    i = i + 1
    End If
    

   If sl Mod 2 = 0 Then
    excel_sheet.Range(excel_sheet.cells(mycnt, 1), _
                             excel_sheet.cells(i - 1, 37)).Select
                             excel_app.selection.Interior.ColorIndex = 15
   Else
   
      excel_sheet.Range(excel_sheet.cells(mycnt, 1), _
                             excel_sheet.cells(i - 1, 37)).Select
                             excel_app.selection.Interior.ColorIndex = 6
   End If
    
    
    
    sl = sl + 1
    rs.MoveNext
    Loop
    End If

    'MAKE UP
'
'
'                            excel_sheet.Range(excel_sheet.Cells(2, 2), _
'                             excel_sheet.Cells(2, 6)).Select
'
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(2, 2) = "ACTIVE FARMER"
'
'
'     excel_sheet.Range(excel_sheet.Cells(2, 7), _
'                             excel_sheet.Cells(2, 12)).Select
'
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(2, 7) = "REGISTRATION"
'
'
'
'      excel_sheet.Range(excel_sheet.Cells(2, 13), _
'                             excel_sheet.Cells(2, 17)).Select
'
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(2, 13) = "FIELD"
'
'
'    excel_sheet.Range(excel_sheet.Cells(2, 18), _
'                             excel_sheet.Cells(2, 21)).Select
'
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(2, 18) = "STORAGE"
'
'
'     excel_sheet.Range(excel_sheet.Cells(2, 23), _
'                             excel_sheet.Cells(2, 28)).Select
'
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(2, 23) = "FILLIN"
    
    
    
     excel_sheet.Range(excel_sheet.cells(3, 1), _
     excel_sheet.cells(i, 38)).Select
     excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
     excel_sheet.cells(4, 2).Select
     excel_app.ActiveWindow.FreezePanes = True
     excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:Ah3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "PLANTED LIST"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault
    
  
  
  
  
  
  
  
  
  
  
End Sub
Function max(ByVal a As Integer, ByVal b As Integer, ByVal c As Integer) As Integer
Dim maxAB As Integer
mymax = 0
Dim T As Integer
T = IIf(a > b, a, b)
mymax = IIf(T > c, T, c)

End Function
Private Sub Command6_Click()
mchk = True
chkred = True
Dim SQLSTR As String
Dim myphone As String
Dim TOTLAND As Double
Dim totadd As Double
TOTLAND = 0
totadd = 0
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Set rsadd = Nothing
'Dim sqlstr As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                       

    db.Execute "delete from tbltemp"

SQLSTR = ""
   SQLSTR = "insert into tbltemp SELECT max(END), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode FROM phealthhub15_core where farmerbarcode<>'' group  by farmerbarcode,fdcode"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT max(END ) as end , dcode, gcode, tcode, fcode, farmerbarcode, (totaltrees),fdcode FROM phealthhub15_core WHERE farmerbarcode ='' GROUP BY dcode, gcode, tcode, fcode,fdcode", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  Set rsF = Nothing
  
  rsF.Open "select * from tbltemp where var6='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
  If rsF.EOF <> True Then
    
  If rsF!var1 > rss!end Then
  db.Execute "update tbltemp set var1='" & Format(rsF!var1, "yyyy-MM-dd") & "' , var7='" & rss!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,fdcode)values('" & Format(rss!end, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "') "
  
  End If
  
  rss.MoveNext
  Loop
  
  
  'storage
  SQLSTR = ""
   SQLSTR = "insert into tbltemp SELECT max(END), dcode, gcode, tcode,fcode, scanlocation, (totaltrees),'S',''  as fs FROM storagehub6_core where scanlocation<>'' group  by scanlocation"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT max(END ) as end , dcode, gcode, tcode, fcode, scanlocation, (totaltrees)FROM storagehub6_core WHERE scanlocation ='' GROUP BY dcode, gcode, tcode, fcode", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
  Set rsF = Nothing
  
  rsF.Open "select * from tbltemp where var6='" & mfcode & "'", db
  If rsF.EOF <> True Then
    
'  If RSF!var1 > RSS!End Then
'  db.Execute "update tbltemp set var1='" & Format(RSF!var1, "yyyy-MM-dd") & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='S' "
'  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='S' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs)values('" & Format(rss!end, "yyyy-MM-dd") & "','" & mdcode & "','" & mgcode & "','" & mtcode & "','99','" & mfcode & "','" & rss!totaltrees & "','S') "
  
  End If
  
  rss.MoveNext
  Loop
                        
                        
                        
                        

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
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    i = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = "date last visited"
    excel_sheet.cells(3, 3) = "farmercode"
'    excel_sheet.Cells(3, 4) = "TSHOWOG"
'     excel_sheet.Cells(3, 5) = "FARMER CODE"
'    excel_sheet.Cells(3, 6) = "FAMER"
'    excel_sheet.Cells(3, 7) = "REG. LAND (ACRE)"
'    excel_sheet.Cells(3, 8) = "PLANTED(ACRE)"
'    excel_sheet.Cells(3, 9) = "ACTUAL DISTRIBUTED"
'    excel_sheet.Cells(3, 10) = "TREES(FIELD)"
'    excel_sheet.Cells(3, 11) = "REES(STORAGE"
      i = 4
                        
                        
                        SQLSTR = ""
SQLSTR = "SELECT max(var1) as var1,var2,var3,var4,var5,var6 from tbltemp  group by var6 "
                        
                        
                            Set rs = Nothing
                            rs.Open SQLSTR, db
                            If rs.EOF <> True Then
                            Do While rs.EOF <> True
                            chkred = False
                            excel_sheet.cells(i, 1) = sl
                           
     excel_sheet.cells(i, 2) = "'" & rs!var1
   
   excel_sheet.cells(i, 3) = rs!var6
  
  
  
  
  
    i = i + 1
      sl = sl + 1
    
    
    
    
    
  
    
 
                            
                            rs.MoveNext
                            Loop
End If

      
                            
                            
                            'make up
   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 11)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:k3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "PLANTED LIST"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault

'excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

End Sub

Private Sub Command7_Click()
mchk = True
chkred = True
Dim SQLSTR As String
Dim myphone As String
Dim TOTLAND As Double
Dim totadd As Double
TOTLAND = 0
totadd = 0
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Set rsadd = Nothing
'Dim sqlstr As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
'db.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ODKLOCAL;Initial Catalog=odk_prodLocal" ' local connection
'odk_prodLocal
db.Open OdkCnnString
                       
'db.Open tempstr
db.Execute "delete from tbltemp"


SQLSTR = ""
   SQLSTR = "insert into tbltemp SELECT (END), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode,id,sname,fname FROM phealthhub15_core where farmerbarcode<>''"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT (END ) as end , dcode, gcode, tcode, fcode, farmerbarcode, (totaltrees),fdcode,id,sname,fname FROM phealthhub15_core WHERE farmerbarcode =''", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  

  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,fdcode,id,sname,fname)values('" & Format(rss!end, "yyyy-MM-dd") & "','" & mdcode & "','" & mgcode & "','" & mtcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "','" & rss!id & "','" & rss!sname & "','" & rss!fname & "') "

  
  rss.MoveNext
  Loop
  
  
  'storage
  SQLSTR = ""
   SQLSTR = "insert into tbltemp SELECT (END), dcode, gcode, tcode,fcode, scanlocation, (totaltrees),'S',''  as fs,id,sname,fname FROM storagehub6_core where scanlocation<>''"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT (END ) as end , dcode, gcode, tcode, fcode, scanlocation, (totaltrees) ,id,sname,fname FROM storagehub6_core WHERE scanlocation ='' ", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  
 
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='S' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,id,sname,fname)values('" & Format(rss!end, "yyyy-MM-dd") & "','" & mdcode & "','" & mgcode & "','" & mtcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','S','" & rss!id & "','" & rss!sname & "','" & rss!fname & "') "
  
  
  
  rss.MoveNext
  Loop
                        
                        
                        
                        

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
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    i = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
     excel_sheet.cells(3, 2) = "DATE VISITED"
    excel_sheet.cells(3, 3) = "DZONGKHAG"
    excel_sheet.cells(3, 4) = "GEWOG"
    excel_sheet.cells(3, 5) = "TSHOWOG"
     excel_sheet.cells(3, 6) = "FARMER CODE"
    excel_sheet.cells(3, 7) = "FAMER"
    excel_sheet.cells(3, 8) = "S. ID"
    excel_sheet.cells(3, 9) = "S. NAME"
      i = 4
                        
                        
                        SQLSTR = ""
'SQLSTR = "SELECT farmercode,sum(acreplanted)as pl,sum(nooftrees) as t from tblplanted group by farmercode"
         '
               
              SQLSTR = "select * from tbltemp where substring(var1,1,2)<>0 and var6 not in(select farmercode from allfarmersexdropped where FarmerCode is not null)  order by fname desc"
                            Set rs = Nothing
                            rs.Open SQLSTR, db
                            If rs.EOF <> True Then
                            Do While rs.EOF <> True
                            chkred = False
                            excel_sheet.cells(i, 1) = sl
                  excel_sheet.cells(i, 2) = "'" & rs!var1
                excel_sheet.cells(i, 3) = "'" & rs!VAR2
                excel_sheet.cells(i, 4) = "'" & rs!VAR3
                excel_sheet.cells(i, 5) = "'" & rs!VAR4
                If sl = 126 Then
                
                MsgBox "asdlcjads"
                End If
                excel_sheet.cells(i, 6) = rs!var6
                excel_sheet.cells(i, 7) = rs!fname
    
   excel_sheet.cells(i, 8) = rs!id
 excel_sheet.cells(i, 9) = rs!sname
  
  
  
    i = i + 1
      sl = sl + 1
    
    
    
    
    
  
    
 
                            
                            rs.MoveNext
                            Loop
End If

      
                            
                            
                            'make up
   excel_sheet.Range(excel_sheet.cells(sl, 1), _
    excel_sheet.cells(i, 11)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:k3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "PLANTED LIST"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault

'excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub Command9_Click()
mchk = True
chkred = True
Dim SQLSTR As String
Dim myphone As String
Dim TOTLAND As Double
Dim totadd As Double
TOTLAND = 0
totadd = 0
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Set rsadd = Nothing
'Dim sqlstr As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
'db.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ODKLOCAL;Initial Catalog=odk_prodLocal" ' local connection
'odk_prodLocal
db.Open OdkCnnString
                      
'db.Open tempstr
db.Execute "delete from mtemp"


SQLSTR = ""
   SQLSTR = "insert into mtemp SELECT _URI, region_dcode, region_gcode, region,fcode FROM phealthhub15_core where farmerbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "0000" & rss!dcode
  mdcode = "D" & Right(mdcode, 2)
  mgcode = "00000" & rss!gcode
  mgcode = "G" & Right(mgcode, 2)
  mtcode = "0000" & rss!tcode
  mtcode = "T" & Right(mtcode, 2)
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  Set rsF = Nothing
  
  db.Execute "update phealthhub15_core set farmerbarcode='" & mfcode & "' where dcode='" & rss!dcode & "' and gcode='" & rss!gcode & "' and tcode='" & rss!tcode & "' and fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"


  rss.MoveNext
  Loop
  
  
  'storage
  
 
                        
                        
                        
                        

MsgBox "done"
   
End Sub

Private Sub Form_Load()
Mname = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

End Sub

Private Sub Label1_Click()
Dim sURL As String

sURL = "http://www.google.com"
Shell "C:\Program Files\Internet Explorer\iexplore.exe " & sURL, vbMaximizedFocus
End Sub

