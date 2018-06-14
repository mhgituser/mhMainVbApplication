VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmmedia 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command42 
      Caption         =   "Command42"
      Height          =   375
      Left            =   9120
      TabIndex        =   43
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command41 
      Caption         =   "zero"
      Height          =   495
      Left            =   8400
      TabIndex        =   42
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command40 
      Caption         =   "update faremerbarcode"
      Height          =   495
      Left            =   4560
      TabIndex        =   41
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command39 
      Caption         =   "inet test"
      Height          =   375
      Left            =   7920
      TabIndex        =   40
      Top             =   5880
      Width           =   735
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5880
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command38 
      Caption         =   "dschool image"
      Height          =   615
      Left            =   10920
      TabIndex        =   39
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command37 
      Caption         =   "dschool farmercode fix"
      Height          =   615
      Left            =   10800
      TabIndex        =   38
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command36 
      Caption         =   "sean"
      Height          =   615
      Left            =   5640
      TabIndex        =   37
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command35 
      Caption         =   "monthly monitoring performance"
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
      Left            =   2880
      TabIndex        =   36
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command34 
      Caption         =   "create dashboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   35
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command33 
      Caption         =   "fix person registering"
      Height          =   495
      Left            =   4680
      TabIndex        =   34
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton fileupdate 
      Caption         =   "fileupdate"
      Height          =   735
      Left            =   9120
      TabIndex        =   33
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command32 
      Caption         =   "testing image download from vps"
      Height          =   975
      Left            =   3000
      TabIndex        =   32
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "tset ea send mail"
      Height          =   615
      Left            =   2880
      TabIndex        =   31
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command30 
      Caption         =   "st altitude"
      Height          =   735
      Left            =   9000
      TabIndex        =   30
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "phi"
      Height          =   495
      Left            =   4200
      TabIndex        =   29
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command28 
      Caption         =   "last field visit"
      Height          =   495
      Left            =   840
      TabIndex        =   28
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command27 
      Caption         =   "planted-hdr-detail"
      Height          =   495
      Left            =   8520
      TabIndex        =   27
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton Command26 
      Caption         =   "dailyactemail"
      Height          =   735
      Left            =   6600
      TabIndex        =   26
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command25 
      Caption         =   "mail test-smtp"
      Height          =   615
      Left            =   2640
      TabIndex        =   25
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "image"
      Height          =   615
      Left            =   5280
      TabIndex        =   24
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command23 
      Caption         =   "dist planning tool"
      Height          =   375
      Left            =   3240
      TabIndex        =   23
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command22 
      Caption         =   "45 days prod act"
      Height          =   615
      Left            =   5040
      TabIndex        =   22
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton Command21 
      Caption         =   "random"
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command20 
      Caption         =   "storage mortality"
      Height          =   735
      Left            =   10560
      TabIndex        =   20
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command19 
      Caption         =   "field vs last visit gps"
      Height          =   615
      Left            =   9720
      TabIndex        =   19
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   615
      Left            =   6960
      TabIndex        =   18
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "last storage visit"
      Height          =   855
      Left            =   8400
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "height analysis"
      Height          =   735
      Left            =   7080
      TabIndex        =   16
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton Command15 
      Caption         =   "hr noti email"
      Height          =   975
      Left            =   3120
      TabIndex        =   15
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      Caption         =   "dist confirmation to monitors"
      Height          =   1095
      Left            =   10080
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "bod-field storage"
      Height          =   1095
      Left            =   3600
      TabIndex        =   13
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command12 
      Caption         =   "dist plan assment"
      Height          =   855
      Left            =   8640
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Monitor status"
      Height          =   975
      Left            =   5280
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "phealth-ok"
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H8000000D&
      Caption         =   "daily meeting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "extension-ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "regrpt"
      Height          =   855
      Left            =   10080
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "qcqc-ok"
      Height          =   1095
      Left            =   5520
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "dist plan"
      Height          =   1215
      Left            =   6480
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "m acre"
      Height          =   1095
      Left            =   6960
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   6960
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Farmers Without Photo"
      Height          =   855
      Left            =   2880
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Farmers Code"
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _cx             =   4895
      _cy             =   11245
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
      Rows            =   5000
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
Attribute VB_Name = "frmmedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long
Public Function ComputerName() As String
  Dim sBuffer As String
  
  Dim lAns As Long
 
  sBuffer = Space$(255)
  lAns = GetComputerName(sBuffer, 255)
  If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
   Else
        err.Raise err.LastDllError, , _
          "A system call returned an error code of " _
           & err.LastDllError
   End If

End Function
Private Sub Command1_Click()
ListFiles "\\192.168.1.12\MhGeneral\MHV Media Library\FARMERS", "" 'mpg files
End Sub

Private Sub ListFiles(strPath As String, Optional Extention As String)
'Leave Extention blank for all files
    Dim File As String
    Dim i As Integer
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
       
    If Trim$(Extention) = "" Then
        Extention = "*.*"
    ElseIf Left$(Extention, 2) <> "*." Then
        Extention = "*." & Extention
    End If
       
    File = Dir$(strPath & Extention)
    i = 1
    VSFlexGrid1.Clear
    Do While Len(File)
        'List1.AddItem File
        VSFlexGrid1.TextMatrix(i, 1) = Mid(File, 1, 14)
        File = Dir$
       i = i + 1
    Loop
End Sub

Private Sub Command10_Click()

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsch As New ADODB.Recordset

Dim actstring As String

ODKDB.Execute "insert into phealthhub15_coreext select *,'' from phealthhub15_core where _uri not in(select _uri from phealthhub15_coreext)"


Set rs = Nothing
rs.Open "select * from phealthhub15_core where   _uri  in(select _uri from phealthhub15_coreext where management='yes' and length(mgmtcomments)=0) ", ODKDB
Do While rs.EOF <> True

            Set rs1 = Nothing
            rs1.Open "select * from phealthhub15_management1 where _parent_auri='" & rs![_uri] & "' ", ODKDB
            
          
                        Do While rs1.EOF <> True
                        Set rsch = Nothing
                        rsch.Open "select * from tblfieldchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", ODKDB
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
            
            ODKDB.Execute "update phealthhub15_coreext set mgmtcomments='" & actstring & "' where _uri='" & rs![_uri] & "' "
            
            End If
            
            actstring = ""
            rs.MoveNext
Loop

End Sub

Private Sub Command11_Click()

'updatemonitor
'updatemonitorassignedfarmer
'updatemonitorteritory
'updatemonitoracreregistered
'updatemonitor45daysvisit
'updatemonitorzeruvisit
'updatemonitorqcofqc
'updatemonitorfieldstoragevisit
'updatemonitormortality 'This function shifted to edok

End Sub
Private Sub updatemonitorfieldstoragevisit()
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

MHVDB.Execute "update tblmonitorstatus set productiveactivity='0',fieldvisit=0,storagevisit=0"
 'SQLSTR = " select _uri,n.start,n.staffbarcode from dailyacthub9_core n INNER JOIN (SELECT staffbarcode,MAX(start )" _
         & "lastEdit FROM dailyacthub9_core GROUP BY staffbarcode)x ON " _
         & "n.staffbarcode = x.staffbarcode AND n.start = x.LastEdit " _
         & "AND STATUS <>  'BAD' and (SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' " _
         & "and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "') GROUP BY n.staffbarcode"



SQLSTR = "select * from dailyacthub9_core where  (SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "')  order by staffbarcode "

'SQLSTR = "select * from dailyacthub9_core where  (SUBSTRING( start ,1,10)>='2014-01-17' and SUBSTRING( start ,1,10)<='2014-01-18')  order by staffbarcode "
Set rs = Nothing
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
Set rs1 = Nothing
Set rs1 = Nothing
rs1.Open "select count(*) cnt from dailyacthub9_activities where _PARENT_AURI='" & rs![_uri] & "' and value in( " _
& "select name from tbldailyactchoices where isproductive='1')", ODKDB
If IIf(IsNull(rs1!cnt), 0, rs1!cnt) > 0 Then
MHVDB.Execute "update tblmonitorstatus set productiveactivity='1' where substring(monitor,1,5)='" & Mid(rs!staffbarcode, 1, 5) & "'"
End If
 

rs.MoveNext
Loop


SQLSTR = "select staffbarcode,count(*) as cnt from phealthhub15_core where  (SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "') group by staffbarcode order by staffbarcode "

'SQLSTR = "select staffbarcode,count(*) as cnt from phealthhub15_core where  (SUBSTRING( start ,1,10)>='2014-01-17' and SUBSTRING( start ,1,10)<='2014-01-18') group by staffbarcode  order by staffbarcode "
Set rs = Nothing
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True

MHVDB.Execute "update tblmonitorstatus set fieldvisit='" & rs!cnt & "' where substring(monitor,1,5)='" & Mid(rs!staffbarcode, 1, 5) & "'"
rs.MoveNext
Loop

SQLSTR = "select staffbarcode,count(*) as cnt from storagehub6_core where  (SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "') group by staffbarcode order by staffbarcode "

'SQLSTR = "select staffbarcode,count(*) as cnt from storagehub6_core where  (SUBSTRING( start ,1,10)>='2014-01-17' and SUBSTRING( start ,1,10)<='2014-01-18') group by staffbarcode order by staffbarcode "
Set rs = Nothing
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True

MHVDB.Execute "update tblmonitorstatus set storagevisit='" & rs!cnt & "' where substring(monitor,1,5)='" & Mid(rs!staffbarcode, 1, 5) & "'"
rs.MoveNext
Loop


End Sub
Private Sub updatemonitor()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select concat(staffcode,' ',staffname) staffname from tblmhvstaff where " _
& " staffcode not in(select substring(monitor,1,5) monitor from tblmonitorstatus) and moniter='1'", MHVDB

Do While rs.EOF <> True

MHVDB.Execute "insert into tblmonitorstatus(monitor,mobilestatus,backitudestatus,qcofqcfail) values( " _
& "'" & rs!staffname & "','G','R','Y')"
rs.MoveNext
Loop



End Sub

Private Sub updatemonitorassignedfarmer()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select monitor,count(idfarmer) as cnt from tblfarmer where status not in('D','R') group by monitor order by monitor", MHVDB

Do While rs.EOF <> True

MHVDB.Execute "update tblmonitorstatus set assignedfarmers='" & rs!cnt & "' where substring(monitor,1,5)='" & rs!monitor & "'"

rs.MoveNext
Loop


End Sub
Private Sub updatemonitorteritory()
MHVDB.Execute "update tblmonitorstatus a,tblmhvstaff b set a.teritory=mteritory where substring(monitor,1,5)=staffcode"
End Sub

Private Sub updatemonitoracreregistered()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select staffcode as monitor,sum(regland) as regland from tblregistrationrpt where stype='M' group by staffcode", MHVDB

Do While rs.EOF <> True

MHVDB.Execute "update tblmonitorstatus set acrereg='" & rs!regland & "' where substring(monitor,1,5)='" & rs!monitor & "'"

rs.MoveNext
Loop


End Sub

Private Sub updatemonitor45daysvisit()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select staffbarcode as monitor,percentage from tblfarmvisit ", ODKDB

Do While rs.EOF <> True

MHVDB.Execute "update tblmonitorstatus set last45daysvisit='" & rs!percentage & "' where substring(monitor,1,5)='" & rs!monitor & "'"

rs.MoveNext
Loop


End Sub
Private Sub updatemonitorqcofqc()
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
  SQLSTR = " select _uri,n.end,n.mid from monitor_qcv5_core n INNER JOIN (SELECT mid,MAX(END )" _
         & "lastEdit FROM monitor_qcv5_core GROUP BY mid)x ON " _
         & "n.mid = x.mid AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.mid"
  rs.Open SQLSTR, ODKDB
   
  Do While rs.EOF <> True
  Set rs1 = Nothing
  rs1.Open "select * from monitor_qcv5_field where _PARENT_AURI='" & rs![_uri] & "'", ODKDB
  If rs1.EOF <> True Then
  Set RS2 = Nothing
  RS2.Open "select count(*) cnt from monitor_qcv5_field where _PARENT_AURI='" & rs![_uri] & "' and qcpass<>'pass'", ODKDB
  If IIf(IsNull(RS2!cnt), 0, RS2!cnt) > 0 Then
   MHVDB.Execute "update tblmonitorstatus set qcofqcfail='R' where substring(monitor,1,5)='" & Mid(rs!Mid, 1, 5) & "'"
  Else
  MHVDB.Execute "update tblmonitorstatus set qcofqcfail='G' where substring(monitor,1,5)='" & Mid(rs!Mid, 1, 5) & "'"
  End If
 
  End If
  
  rs.MoveNext
  Loop
         
         
         
End Sub

Private Sub updatemonitorzeruvisit()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select staffcode as monitor,count(*) as cnt from tblzerovisit group by staffcode ", MHVDB

Do While rs.EOF <> True

MHVDB.Execute "update tblmonitorstatus set zerovisit='" & rs!cnt & "' where substring(monitor,1,5)='" & rs!monitor & "'"

rs.MoveNext
Loop


End Sub


Private Sub updatemonitormortality()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim SQLSTR As String
SQLSTR = ""
Set rs = Nothing
rs.Open "SELECT DISTINCT staffcode as monitor FROM  tblmhvstaff WHERE moniter='1'" _
& " order by cast(substring(staffcode,2,4) as unsigned)", MHVDB

Do While rs.EOF <> True
Set rs1 = Nothing
SQLSTR = "SELECT SUM( TREE_COUNT_TOTALTREES ) totaltrees, SUM( TREE_COUNT_DEADMISSING ) deadmissing, " _
& " (Sum(TREE_COUNT_DEADMISSING) * 100) / SUM( TREE_COUNT_TOTALTREES ) mortality From " _
& " tblfieldlastvisitrpt Where SUBSTRING(farmerbarcode, 1, 9) IN (" _
& " SELECT SUBSTRING( idfarmer, 1, 9 )From mhv.tblfarmer " _
& " WHERE SUBSTRING( monitor, 1, 5 ) =  '" & rs!monitor & "' ) "
rs1.Open SQLSTR, ODKDB

If rs1.EOF <> True Then
MHVDB.Execute "update tblmonitorstatus set totaltrees='" & rs1!totaltrees & "', " _
& " deadmissing='" & rs1!deadmissing & "',mortality='" & rs1!mortality & "' where substring(monitor,1,5)='" & rs!monitor & "'"
End If
rs.MoveNext
Loop



End Sub


Private Sub Command12_Click()
Dim rs As New ADODB.Recordset
Dim SQLSTR As String
SQLSTR = ""
MHVDB.Execute "delete from tbldistplanassesment"
SQLSTR = "insert into tbldistplanassesment SELECT SUBSTRING(IDFARMER,1,3)dzcode,SUBSTRING(IDFARMER,4,3)gecode, " _
& " SUBSTRING(IDFARMER,7,3)tscode,IDFARMER,FARMERNAME,0 REGLAND, " _
& " village,phone1,'' platedstatus,'' dgt,'',1,'','',0,0,0,0,0,0,0,0,0,0,0,monitor,0,'','' FROM tblfarmer  WHERE status not in('D','R')"

MHVDB.Execute SQLSTR

MHVDB.Execute "delete from tbldistplanassesment where idfarmer not in(select farmercode from tblplanted)"


Set rs = Nothing
rs.Open "select farmerid,sum(regland) regland from tbllandreg " _
& " where farmerid in(select farmercode from tblplanted where status<>'C') group by farmerid", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistplanassesment set regland='" & rs!regland & "' " _
& "  where idfarmer='" & rs!farmerid & "' "
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select farmercode,sum(challanqty) chhalanqty from tblplanted where status<>'C' group by farmercode", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistplanassesment set plantsalreadydistributed='" & rs!chhalanqty & "' " _
& "  where idfarmer='" & rs!farmercode & "' "
rs.MoveNext
Loop






'Set rs = Nothing
'rs.Open "select * from tbllandregdetail where plantedstatus in('N')", MHVDB
'Do While rs.EOF <> True
'MHVDB.Execute "update tbldistplan set regland='" & rs!acre & "' where  plantedstatus='P' and idfarmer='" & rs!farmercode & "' "
'rs.MoveNext
'Loop



Set rs = Nothing
rs.Open "select * from tbldistplanassesment ", MHVDB
Do While rs.EOF <> True
FindDZ rs!dzcode
FindGE rs!dzcode, rs!GECODE
FindTs rs!dzcode, rs!GECODE, rs!tscode
MHVDB.Execute "update tbldistplanassesment set dzcode='" & rs!dzcode & "  " & Dzname & "', " _
& " gecode='" & rs!GECODE & "  " & GEname & "', " _
& " tscode='" & rs!tscode & "  " & TsName & "', " _
& "dgt='" & rs!dzcode & "  " & Dzname & "  " & rs!GECODE & "  " & GEname & "  " & rs!tscode & "  " & TsName & "'" _
& "where idfarmer='" & rs!idfarmer & "' "
rs.MoveNext
Loop

SQLSTR = ""

  SQLSTR = "select n.farmerbarcode ,sum(TREE_COUNT_TOTALTREES) totaltrees, " _
  & " sum(tree_count_deadmissing) deadmissing,end from  " _
  & " phealthhub15_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"

Set rs = Nothing
rs.Open SQLSTR, ODKDB

Do While rs.EOF <> True
MHVDB.Execute "update tbldistplanassesment set fieldtottrees='" & rs!totaltrees & "', " _
& " fielddeadmissing='" & rs!deadmissing & "',fieldstoragetype='F',regdate= '" & Format(rs!end, "yyyy-MM-dd") & "' where idfarmer='" & rs!farmerbarcode & "' "
rs.MoveNext
Loop


SQLSTR = ""

  SQLSTR = "select n.farmerbarcode ,sum(TOTALTREES) totaltrees, " _
  & " sum(DTREES) deadmissing,end from  " _
  & " storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"

Set rs = Nothing
rs.Open SQLSTR, ODKDB

Do While rs.EOF <> True
MHVDB.Execute "update tbldistplanassesment set storagetottrees='" & rs!totaltrees & "', " _
& " storagedeadmissing='" & rs!deadmissing & "',fieldstoragetype=concat(fieldstoragetype,'S') where idfarmer='" & rs!farmerbarcode & "' "
rs.MoveNext
Loop



MHVDB.Execute "update tbldistplanassesment set expectedplants=regland*420"
MHVDB.Execute "update tbldistplanassesment set fieldassesment=fieldtottrees-fielddeadmissing"
MHVDB.Execute "update tbldistplanassesment set storageassesment=storagetottrees-storagedeadmissing"
MHVDB.Execute "update tbldistplanassesment set actualassesment=expectedplants-(fieldassesment+storageassesment)"
MHVDB.Execute "update tbldistplanassesment set recordage=0"
MHVDB.Execute "update tbldistplanassesment set recordage=datediff(CURDATE(),regdate)"
MHVDB.Execute "update tbldistplanassesment set landtype='Private' where substring(idfarmer,10,1)='F'"
MHVDB.Execute "update tbldistplanassesment set landtype='GRF/SRF' where substring(idfarmer,10,1)='G'"
MHVDB.Execute "update tbldistplanassesment set landtype='CF' where substring(idfarmer,10,1)='C'"
MHVDB.Execute "update tbldistplanassesment a ,tblmhvstaff b set staff=concat(staff,'  ',staffname) where staff=staffcode"
MHVDB.Execute "update tbldistplanassesment a ,tblplanted b set previousdisttype='2011-2012' where " _
& " idfarmer in(select farmercode from tblplanted where year in(2011,2012)) "
MHVDB.Execute "update tbldistplanassesment a ,tblplanted b set previousdisttype='2013-New and Additional' where " _
& " length(previousdisttype)=0 "

End Sub

Private Sub Command13_Click()
Dim rs As New ADODB.Recordset
Dim SQLSTR As String
SQLSTR = ""

  SQLSTR = "select count(n.farmerbarcode) fcnt ,sum(TREE_COUNT_TOTALTREES) totaltrees, " _
  & " sum(tree_count_deadmissing) deadmissing,end from  " _
  & " phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and staffbarcode in(select staffcode from " _
         & " mhv.tblmhvstaff where moniter=1 and status<>'C')"

Set rs = Nothing
rs.Open SQLSTR, ODKDB

Do While rs.EOF <> True
ODKDB.Execute "update tblbodfieldstorage set totaltrees='" & rs!totaltrees & "', " _
& " deadmissing='" & rs!deadmissing & "',farmerscount='" & rs!FCNT & "',mortality= '" & (rs!deadmissing * 100) / rs!totaltrees & "' where id=1 "
rs.MoveNext
Loop


SQLSTR = ""

  SQLSTR = "select count(n.farmerbarcode) fcnt ,sum(TOTALTREES) totaltrees, " _
  & " sum(DTREES) deadmissing,end from  " _
  & " storagehub6_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and staffbarcode in(select staffcode from " _
         & " mhv.tblmhvstaff where moniter=1 and status<>'C')"

Set rs = Nothing
rs.Open SQLSTR, ODKDB

Do While rs.EOF <> True
ODKDB.Execute "update tblbodfieldstorage set totaltrees='" & rs!totaltrees & "', " _
& " deadmissing='" & rs!deadmissing & "',farmerscount='" & rs!FCNT & "',mortality= '" & (rs!deadmissing * 100) / rs!totaltrees & "' where id=2 "
rs.MoveNext
Loop


End Sub

Private Sub Command14_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
 MHVDB.Execute "update tbldistplan set monitor='',email=''"
Set rs = Nothing
rs.Open "select distinct substring(idfarmer,1,9) as dgt from tbldistplan", MHVDB
Do While rs.EOF <> True
Set rs1 = Nothing
rs1.Open "select monitor from tblfarmer where substring(idfarmer,1,9)='" & rs!dgt & "'", MHVDB
If rs1.EOF <> True Then
FindsTAFF rs1!monitor

MHVDB.Execute "update tbldistplan set monitor='" & rs1!monitor & "  " & sTAFF & "',email='" & emailAddress & "' " _
& " where substring(idfarmer,1,9)='" & rs!dgt & "'"

End If
rs.MoveNext
Loop
send_email

End Sub
Private Sub send_email()
Dim bodymsg As String
Dim htmlbody As String
Dim FCNT As Integer
Dim montot As Double
Dim tblheader As String
Dim htmlfoot As String
Dim Dzstr As String
Dim mmonitor, memail As String
Dim tropen, trclose As String
Dim mstaffcode As String
emailMessageString = ""
Dim emailmsg As String
Dim recordno As Integer
Dim param As String
Dim mdt As Date
Dim rs As New ADODB.Recordset
Dim oSmtp As New EASendMailObjLib.Mail
  
    oSmtp.LicenseCode = "TryIt"
    oSmtp.FromAddr = "tmindrup@mountainhazelnuts.com"
    
    
    oSmtp.ServerAddr = "smtp.tashicell.com"
    oSmtp.BodyFormat = 1
    tblheader = "<tr>" _
& "<th  bgcolor=""yellow"">S/N</th>" _
& "<th  bgcolor=""yellow"">Dzongkhag</th>" _
& "<th  bgcolor=""yellow"">Gewog</th>" _
& "<th  bgcolor=""yellow"">Tshowog Value</th>" _
& "<th  bgcolor=""yellow"">Farmer Code</th>" _
& "<th  bgcolor=""yellow"">Farmer Name</th>" _
& "<th  bgcolor=""yellow"">Land Type</th>" _
& "<th  bgcolor=""yellow"">Farmer Type</th>" _
& "<th  bgcolor=""yellow"">Acre Registered</th>" _
& "</tr>"
tropen = "<tr>"
trclose = "</tr>"
'bodymsg = "This Record Is Available In Follow Up Log.regularly Up Date The Follow Up Log."
    
Set rs = Nothing
rs.Open "select * from tbldistplanconfirmation where status='No'", MHVDB
Do While rs.EOF <> True
bodymsg = rs!emailmsg
Dzstr = Dzstr + "'" + rs!email + "',"
rs.MoveNext
Loop
             
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
Set rs = Nothing
rs.Open "select monitor,email,dzcode,gecode,tscode,idfarmer,farmername,landtype,farmertype,sum(regland) regland from " _
& " tbldistplan where length(email)>0 and email in " & Dzstr & " group by  monitor,email,dzcode,gecode,tscode,idfarmer,farmername,landtype,farmertype order by  " _
& " monitor,email,dzcode,gecode,tscode,idfarmer,farmername,landtype,farmertype ", MHVDB
Else
Exit Sub
End If
   


    Do Until rs.EOF
       mmonitor = ""
       memail = ""
       emailId = ""
       param = ""
       mstaffcode = ""
       FCNT = 0
       montot = 0
       param = rs!email
        mmonitor = rs!monitor
        mstaffcode = Mid(Trim(rs!monitor), 1, 5)
        memail = rs!email
        emailId = rs!email
      recordno = 0
       Do While param = rs!email
       
          recordno = recordno + 1

            emailMessageString = "<td>" & recordno & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!dzcode & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!GECODE & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!tscode & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!idfarmer & "</td>"
            emailMessageString = emailMessageString & "<td>" & rs!farmername & "</td>"
             emailMessageString = emailMessageString & "<td>" & rs!LANDTYPE & "</td>"
              emailMessageString = emailMessageString & "<td>" & rs!farmertype & "</td>"
            emailMessageString = emailMessageString & "<td>" & Format(rs!regland, "####0.00") & "</td>"
            montot = montot + rs!regland
            FCNT = FCNT + 1
            emailmsg = emailmsg & tropen & emailMessageString & trclose
            
            
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       
       
       
           tblheaderfooter = "<tr>" _
& "<th  bgcolor=""yellow"" colspan=""5"">No. of Farmers:" & FCNT & "</th>" _
& "<th  bgcolor=""yellow""></th>" _
& "<th  bgcolor=""yellow""></th>" _
& "<th  bgcolor=""yellow""></th>" _
& "<th  bgcolor=""yellow""></th>" _
& "<th  bgcolor=""yellow"" colspan=""3"">Total Acre</th>" _
& "<th  bgcolor=""yellow""></th>" _
& "<th  bgcolor=""yellow""></th>" _
& "<th  bgcolor=""yellow"">" & Format(montot, "####0.00") & "</th>" _
& "</tr>"
       
       
        oSmtp.AddRecipientEx "muktitcc@gmail.com", 0
        'emailId
        oSmtp.Subject = Format(Now, "yyyyMMdd") & " " & "Distribution Confirmation " & mmonitor
        oSmtp.BodyText = "<html><head><title></title></head><body><br><h5>" & bodymsg & "</h5><br><TABLE border=""1"" cellspacing=""0"" >" & tblheader & emailmsg & tblheaderfooter & "</TABLE></body></html>"
       If oSmtp.SendMail() = 0 Then
       ' send ok
      ODKDB.Execute "update tbldistplanconfirmation set status='Yes' where email='" & emailId & "' sentdate='" & Format(Now, "yyyy-MM-dd") & "' "
      emailMessageString = ""
      emailmsg = ""
    Else
        'MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    emailMessageString = ""
    emailmsg = ""
    End If
Loop




End Sub


Private Sub Command15_Click()
Dim bodymsg As String
Dim Dzstr As String
Dim memail As String
Dim emailmsg As String

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim oSmtp As New EASendMailObjLib.Mail
  
    oSmtp.LicenseCode = "TryIt"
    oSmtp.FromAddr = "ytamang@mountainhazelnuts.com"
    oSmtp.ServerAddr = "smtp.tashicell.com"
    oSmtp.BodyFormat = 1
 
Set rs = Nothing
rs.Open "select * from tblnotificationheader where emailsent='No' and notificationtype='Daily' and mhour='" & Format(Now, "HH") & "'", MHVHRDB
If rs.EOF <> True Then

Else
Set rs = Nothing
rs.Open "select * from tblnotificationheader where emailsent='No' and notificationtype='Weekly' and mday='" & Weekday(Now, vbMonday) & "' and mhour='" & Format(Now, "HH") & "'", MHVHRDB
If rs.EOF <> True Then

Else

Set rs = Nothing
rs.Open "select * from tblnotificationheader where emailsent='No' and notificationtype='Monthly' and mmonth='" & Month(Now) & "' and mdate='" & Day(Now) & "' and mhour='" & Format(Now, "HH") & "'", MHVHRDB
If rs.EOF <> True Then

Else
Exit Sub
End If

End If

End If




Do While rs.EOF <> True
bodymsg = rs!NotificationMessage

Set rs1 = Nothing
rs1.Open "select * from tblnotificationdetail where headerid='" & rs!trnid & "'", MHVHRDB
Do While rs1.EOF <> True
Dzstr = Dzstr + rs1!email + ","
rs1.MoveNext
Loop
             
If Len(Dzstr) > 0 Then
   Dzstr = Left(Dzstr, Len(Dzstr) - 1)
   Else
   Exit Sub
End If


      
       
oSmtp.AddRecipientEx Dzstr, 0
oSmtp.Subject = "Reminder: Weekly Pulse Report "
oSmtp.BodyText = "<html><head><title></title></head><body><br><h5>" & bodymsg & "</h5><br></body></html>"
If oSmtp.SendMail() = 0 Then
 ' send ok
   MHVHRDB.Execute "update tblnotificationheader set emailsent='Yes' where trnid='" & rs!trnid & "'"
   emailMessageString = ""
   emailmsg = ""
    Else
   'MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()

End If

bodymsg = ""
Dzstr = ""

rs.MoveNext
Loop

End Sub

Private Sub Command16_Click()
Dim rs As New ADODB.Recordset
ODKDB.Execute "delete from tblplantheightanalysisrpt"
SQLSTR = ""
      SQLSTR = "insert into tblplantheightanalysisrpt(farmercode,fieldcode,datevisitpre) " _
         & " select n.farmerbarcode,n.fdcode,n.end from phealthhub15_core   n INNER JOIN " _
         & "(SELECT farmerbarcode,fdcode, MAX(END)" _
         & "lastEdit FROM phealthhub15_core where substring(end,1,10)>='20101231'  and substring(end,1,10)<='20121231' GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'   GROUP BY n.farmerbarcode, n.fdcode"
         
           'SQLSTR = "insert into tblplantheightanalysisrpt(farmercode,datevisitpre) select n.end,n.farmerbarcode from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' GROUP BY n.farmerbarcode"
         
   ODKDB.Execute SQLSTR
   
   
         SQLSTR = " select n.farmerbarcode,n.fdcode,n.end,TREE_COUNT_TOTALTREES,TREE_COUNT_DEADMISSING,TREEHEIGHT from phealthhub15_core   n INNER JOIN " _
         & "(SELECT farmerbarcode,fdcode, MAX(END)" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'   GROUP BY n.farmerbarcode, n.fdcode"
   
   Set rs = Nothing
   rs.Open SQLSTR, ODKDB
   
   Do While rs.EOF <> True
   ODKDB.Execute "update tblplantheightanalysisrpt set datevisitcurrent='" & Format(rs!end, "yyyy-MM-dd") & "', " _
   & " recordage='" & DateDiff("d", rs!end, Now) & "',totaltrees='" & rs!tree_count_totaltrees & "'," _
   & " deadmissing='" & rs!tree_count_deadmissing & "',avgheight='" & rs!treeheight & "' where " _
   & " farmercode='" & rs!farmerbarcode & "' and fieldcode='" & rs!FDCODE & "'"
    rs.MoveNext
   Loop
   
    Set rs = Nothing
   rs.Open "select * from tblfarmer", MHVDB
    Do While rs.EOF <> True
   ODKDB.Execute "update tblplantheightanalysisrpt set farmername='" & rs!farmername & "', " _
   & " currentfarmerstatus='" & rs!status & "' where " _
   & " farmercode='" & rs!idfarmer & "'"
    rs.MoveNext
   Loop

End Sub

Private Sub Command17_Click()
Dim SQLSTR As String
      SQLSTR = ""
      SQLSTR = "insert into tblstoragedlastvisitrpt(START,end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,farmerbarcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS) " _
      & "select START,n.end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,n.farmerbarcode," _
      & " DTREES,0,0,ddamage,pdamage, " _
      & " ndtrees,GPS_LAT,GPS_LNG,0, " _
      & "0,TOTALTREES,wlogged,0,adamage, " _
      & "monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
         
      'ddamage(deasease damage)=STEMPEST
      'pdamage(Number of trees with pest damage)=ROOTPEST
      'ndtrees(Number of nutrient deficient trees)=ACTIVEPEST
      
 ODKDB.Execute "delete from tblstoragedlastvisitrpt"
 ODKDB.Execute SQLSTR
 ODKDB.Execute "update tblstoragedlastvisitrpt set region_dcode=substring(farmerbarcode,1,3),region_gcode=substring(farmerbarcode,4,3), " _
 & " region=substring(farmerbarcode,7,3)"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tbldzongkhag b set region_dcode=concat(substring(region_dcode,1,3),'  ',DzongkhagName) where substring(region_dcode,1,3)=DzongkhagCode"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tblgewog b set region_gcode=concat(substring(region_gcode,1,3),'  ',GewogName) where concat(substring(region_dcode,1,3),substring(region_gcode,1,3))=concat(DzongkhagId,GewogId)"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tbltshewog b set region=concat(substring(region,1,3),'  ',TshewogName) where concat(substring(region_dcode,1,3),substring(region_gcode,1,3),substring(region,1,3))=concat(DzongkhagId,GewogId,TshewogId)"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tblfarmer b set farmerbarcode=concat(farmerbarcode,'  ',farmername) where farmerbarcode=idfarmer"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tblmhvstaff b set staffbarcode=concat(staffbarcode,'  ',staffname) where staffbarcode=staffcode"

End Sub

Private Sub Command18_Click()
MsgBox "reg"
End Sub

Private Sub Command19_Click()

Dim rs As New ADODB.Recordset

Set rs = Nothing
rs.Open "select farmercode,end,lat,lng from tblextensionmortality where lat>0 and lng>0", ODKDB
Do While rs.EOF <> True
MHVDB.Execute "update tblfarmer set flat='" & rs!lat & "',flng='" & rs!lng & "'," _
& " lastvisiteddate='" & Format(rs!end, "yyyy-MM-dd") & "' where idfarmer='" & Mid(rs!farmercode, 1, 14) & "'"
rs.MoveNext
Loop

End Sub

Private Sub Command2_Click()
rinzin
End Sub

Private Sub rinzin()
Dim excel_app As Object
Dim excel_sheet As Object
Dim farmerstr As String
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset


For i = 1 To VSFlexGrid1.rows - 1
If Len(VSFlexGrid1.TextMatrix(i, 1)) = 0 Then Exit For
farmerstr = farmerstr + "'" + Trim(VSFlexGrid1.TextMatrix(i, 1)) + "',"

Next

If Len(farmerstr) > 0 Then
 farmerstr = "(" + Left(farmerstr, Len(farmerstr) - 1) + ")"
End If















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
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = "DZONGKHAG"
    excel_sheet.cells(3, 3) = "GEWOG"
    excel_sheet.cells(3, 4) = "TSHOWOG"
    excel_sheet.cells(3, 5) = "FARMER CODE"
    excel_sheet.cells(3, 6) = "FARMER NAME"
    excel_sheet.cells(3, 7) = "MONITOR"
    i = 4





Set rs = Nothing
rs.Open "SELECT distinct farmercode FROM tblplanted where farmercode not in" & farmerstr & " and status<>'C' group by farmercode order by farmercode ", MHVDB

If rs.EOF <> True Then

Do While rs.EOF <> True

excel_sheet.cells(i, 1) = sl
FindDZ Mid(rs!farmercode, 1, 3)
    excel_sheet.cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
   FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
    excel_sheet.cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
  excel_sheet.cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
  
  excel_sheet.cells(i, 5) = rs!farmercode
  FindFA rs!farmercode, "F"
  excel_sheet.cells(i, 6) = rs!farmercode & " " & FAName
  
  
  Set rss = Nothing
rss.Open "select  monitor from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
If rss.EOF <> True Then
FindsTAFF rss!monitor
excel_sheet.cells(i, 7) = rss!monitor & " " & sTAFF
Else
excel_sheet.cells(i, 7) = ""
End If


i = i + 1
sl = sl + 1

rs.MoveNext
Loop











Else

'MsgBox "uuummmm"
End If


'make up
   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 7)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:G3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ERROR LISTING"
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
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear

End Sub

Private Sub Command20_Click()
updateextensionmortalitystorage
End Sub

Private Sub updateextensionmortalitystorage()
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
ODKDB.Execute "delete from tblextensionmortalitystorage"
        SQLSTR = ""

           
         SQLSTR = "insert into tblextensionmortalitystorage (dzongkhag,gewog,tshowog,farmercode,fieldcode,lat,lng,alt,acrereg,acrecultivated,totaltrees,deadmissing,percent,end,STAFFBARCODE) " _
         & " select substring(n.farmerbarcode,1,3)," _
         & "substring(n.farmerbarcode,4,3),substring(n.farmerbarcode,7,3),n.farmerbarcode,0,GPS_LAT,GPS_LNG," _
         & "0,0,0,TOTALTREES,DTREES,((DTREES*100)/TOTALTREES),end,STAFFBARCODE from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END)" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and  substring(STAFFBARCODE,1,5) in(select staffcode from mhv.tblmhvstaff where moniter=1 and status<>'C') GROUP BY n.farmerbarcode"
         
 ODKDB.Execute SQLSTR

    'ddamage(deasease damage)=STEMPEST
      'pdamage(Number of trees with pest damage)=ROOTPEST
      'ndtrees(Number of nutrient deficient trees)=ACTIVEPEST
'DTREES=deadmissing


 
 
' Set rs = Nothing
' rs.Open "select * from tblextensionmortalitystorage", ODKDB
' Do While rs.EOF <> True
' Set rs1 = Nothing
' rs1.Open "select sum(regland) as regland from tbllandreg where  farmerid='" & rs!farmercode & "'", MHVDB
'  ODKDB.Execute "update tblextensionmortality set acrereg='" & IIf(IsNull(rs1!regland), 0, rs1!regland) & "' where farmercode='" & rs!farmercode & "'"
' rs.MoveNext
' Loop
 
 
 ODKDB.Execute "update tblextensionmortalitystorage set grpgewog=concat(dzongkhag,gewog)"
 ODKDB.Execute "update tblextensionmortalitystorage set grptshowog=concat(dzongkhag,gewog,tshowog)"
 
 
 ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tbldzongkhag b set " _
 & " dzongkhag= concat(dzongkhagcode,'  ',dzongkhagname) where dzongkhag=b.dzongkhagcode"
 
   ODKDB.Execute " update tblextensionmortalitystorage a ,mhv.tbltshewog b set" _
 & " a.region= b.regioncode where concat(DzongkhagId,GewogId,TshewogId)=substring(farmercode,1,9)"
 
    ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tblextregion b set " _
 & " a.region= extensionofficercode where a.region=b.regioncode"
  
 ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tblgewog b set " _
 & " grpgewog= concat(grpgewog,'  ',gewogname) where grpgewog=concat(dzongkhagid,gewogid)"
 
  ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tbltshewog b set " _
 & " grptshowog= concat(grptshowog,'  ',tshewogname) where grptshowog=concat(dzongkhagid,gewogid,tshewogid)"
 
 ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tblfarmer b set farmercode=concat(farmercode,'  ',farmername)" _
  & "  where farmercode=idfarmer"
  
   ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tblmhvstaff b set STAFFBARCODE=concat(STAFFBARCODE,'  ',staffname)" _
  & "  where staffcode=STAFFBARCODE"
  
    ODKDB.Execute "update tblextensionmortalitystorage a ,mhv.tblmhvstaff b set region=concat(staffcode,'  ',staffname)" _
  & "  where staffcode=region"
End Sub

Private Sub Command21_Click()
'MsgBox randomKey
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblmhvstaff where length(mlockkey)=0 and length(munlockkey)=0 and length(mphonekey)=0 ", MHVDB

Do While rs.EOF <> True
MHVDB.Execute "update tblmhvstaff set mlockkey='" & randomKey & "' where staffcode='" & rs!staffcode & "' and length(mlockkey)=0"
MHVDB.Execute "update tblmhvstaff set munlockkey='" & randomKey & "' where staffcode='" & rs!staffcode & "' and length(munlockkey)=0"
MHVDB.Execute "update tblmhvstaff set mphonekey='" & randomKey & "' where staffcode='" & rs!staffcode & "' and length(mphonekey)=0"
rs.MoveNext
Loop
End Sub

Private Sub Command22_Click()
Dim i, j As Integer
Dim dt As Date
Dim mact As String
Dim mstaff As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim ra As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim SQLSTR As String

ODKDB.Execute "delete from odk_prodlocal.tblproductiveactivitylog"
Set rsm = Nothing
SQLSTR = "insert into odk_prodlocal.tblproductiveactivitylog(monitor) SELECT concat(staffcode,' ',staffname) as staffcode from tblmhvstaff where moniter='1'"
rsm.Open SQLSTR, MHVDB
Set rsm = Nothing
rsm.Open "select monitor from odk_prodlocal.tblproductiveactivitylog", ODKDB

Do While rsm.EOF <> True

Set rs1 = Nothing

rs1.Open "select VALUE,count(*) cnt from dailyacthub9_activities where   _PARENT_AURI in(SELECT _URI FROM `dailyacthub9_core` WHERE STAFFID='" & Mid(rsm!monitor, 1, 5) & "'  and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "') group by VALUE", ODKDB

Do While rs1.EOF <> True
ODKDB.Execute "update odk_prodlocal.tblproductiveactivitylog set " & rs1!Value & " = '" & rs1!cnt & "' where substring(monitor,1,5)='" & Mid(rsm!monitor, 1, 5) & "'"
rs1.MoveNext
Loop
rsm.MoveNext
Loop

    
         
         




End Sub

Private Sub idealdelivery()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset

Dim startdate As Date
Dim enddate As Date
startdate = Year(Now) & "-01-01"
enddate = Year(Now) & "-12-31"

MHVDB.Execute "delete from tblidealdelivery"

MHVDB.Execute "insert into tblidealdelivery(dgtcode,demand) " _
& " select upper(substring(farmerid,1,9)) as dgt,sum(round(regland,0)*420) regland from " _
& " tbllandreg where status not in ('C') and plantedstatus='N' " _
& "  group by substring(farmerid,1,9) having sum(round(regland,0)*420)>0 order by substring(farmerid,1,9) "

Set rs = Nothing

rs.Open "select * from tbltshewog", MHVDB
Do While rs.EOF <> True
            Set rschk = Nothing
            rschk.Open "select * from tbldistributionallow", MHVDB
            Do While rschk.EOF <> True
                    Select Case rschk!id
                            Case 1
                                        If rschk!allow = "Yes" Then
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " mansooncritical='" & Format(rs!mansooncritical, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                        
                                        Else
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " mansooncritical='" & Format(rs!mansoonstart, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                        
                                        End If
                            
                            
                            
                            Case 2
                            
                                        If rschk!allow = "Yes" Then
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " farmercritical='" & Format(enddate, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                        
                                        Else
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " farmercritical='" & Format(rs!farmerbusystart, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                        
                                        End If
                            
                            
                            Case 3
                            
                                        If rschk!allow = "Yes" And rs!irrigationavailable = "Yes" Then
                                            
                                            
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " plantingcritical='" & Format(enddate, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                    
                                                                    
                                        End If
                            
                        Case 4
                        
                                        If rschk!allow = "Yes" And rs!irrigationavailable = 0 Then
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " plantingcritical='" & Format(rs!rainyseasonend, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                        
                                        ElseIf rschk!allow = "No" Then
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " plantingcritical='" & Format(rs!idealplanting, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                                        End If
                            
                            
                           If rschk!allow = "Yes" And rs!irrigationavailable = "Yes" Then ' need to check for min date
                                            
                                            
                                            MHVDB.Execute "update tblidealdelivery set " _
                                            & " diststart1='" & Format(startdate, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                            
                            Else
                            
                                MHVDB.Execute "update tblidealdelivery set " _
                                            & " diststart1='" & Format(rainyseasonstart, "yyyy-MM-dd") & "' " _
                                            & " where dgtcode='" & rs!dzongkhagid & rs!gewogid & rs!tshewogid & "'"
                            
                                                                    
                           End If

                                
                    
                    
                    End Select
                    rschk.MoveNext
            Loop
            
            
            rs.MoveNext
Loop



End Sub

Private Sub Command23_Click()
idealdelivery
End Sub

Private Sub Command24_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim connRemote As New ADODB.Connection
Dim LastUpdateDate As Date
LastUpdateDate = Now
 connRemote.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=206.217.141.114;" _
                        & " DATABASE=odk_prod;" _
                        & "UID=odk_user;PWD=none; OPTION=3"
                        
                        

connRemote.Open

Set rs1 = Nothing

rs1.Open "select _URI from T01_DAILY_CHECK_LIST_CORE where substring(_CREATION_DATE,1,10)>='" & Format(LastUpdateDate - 60, "yyyy-MM-dd") & "'", connRemote, adOpenDynamic
                       
Do While (rs1.EOF <> True)
Set rs = Nothing
rs.Open "select * from T01_DAILY_CHECK_LIST_PICTURE_BLB where _TOP_LEVEL_AURI='" & rs1![_uri] & "'", connRemote, adOpenStatic, adLockOptimistic
If rs.EOF <> True Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
If Len(rs!Value) > 0 Then
mystream.Write rs!Value
mystream.SaveToFile "C:\xampp\htdocs\mhweb\transPic\" & Mid(rs![_TOP_LEVEL_AURI], 6, 600) & ".jpg", adSaveCreateOverWrite
mystream.Close
End If
End If

rs1.MoveNext
Loop
End Sub
Private Sub downloadTransImage(LastUpdateDate As Date)
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim connRemote As New ADODB.Connection
 connRemote.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=206.217.141.114;" _
                        & " DATABASE=odk_prod;" _
                        & "UID=odk_user;PWD=none; OPTION=3"
                        
                        

connRemote.Open

Set rs1 = Nothing

rs1.Open "select _URI from T01_DAILY_CHECK_LIST_CORE where substring(_CREATION_DATE,1,10)>='" & Format(LastUpdateDate - 60, "yyyy-MM-dd") & "'", connRemote, adOpenDynamic
                       
Do While (rs1.EOF <> True)
Set rs = Nothing
rs.Open "select * from T01_DAILY_CHECK_LIST_PICTURE_BLB where _TOP_LEVEL_AURI='" & rs1![_uri] & "'", connRemote, adOpenStatic, adLockOptimistic
If rs.EOF <> True Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
If Len(rs!Value) > 0 Then
mystream.Write rs!Value
mystream.SaveToFile "C:\xampp\htdocs\mhweb\transPic\" & Mid(rs![_TOP_LEVEL_AURI], 6, 600) & ".jpg", adSaveCreateOverWrite
mystream.Close
End If
End If

rs1.MoveNext
Loop
End Sub


Private Sub Command25_Click()
Dim mCONNECTION As String
Dim retVal          As String
Dim emailmessage As String

Dim EMAILIDS As String
EMAILIDS = "muktitcc@gmail.com"
emailmessage = "fuck"
mCONNECTION = "smtp.tashicell.com"

 retVal = SendMail(EMAILIDS, "STATUS ON ODK DATA TRANSFER.", "test@MHV.COM", _
    emailmessage, mCONNECTION, 25, _
    "habizabi", "habizabi", "", CBool(False))
  
If retVal = "ok" Then
MsgBox "done"
Else
MsgBox "Please Check Internet Connection " & retVal
End If
End Sub

Private Sub Command26_Click()
reocreatefirstpassofdailyactemail
createodkdailyactsummary
createphonedailyactsummary
'createfielddeadmissingsummary
End Sub
Private Sub createfielddeadmissingsummary()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim SQLSTR As String
Dim mact As String
Dim actmsg As String
Dim othermsg As String
Dim dm, ap, rp, sp, ad As Double

Set rs = Nothing
rs.Open "select * from tblphithreshold", ODKDB
If rs.EOF <> True Then
dm = rs!deadmissing
ap = rs!activepest
rp = rs!rootpest
sp = rs!stempest
ad = rs!animaldamage
End If


ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='DM' and fieldtype='F' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,TREE_COUNT_DEADMISSING, " _
& " (2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING)) as mper,'No','F','DM' " _
& " from  tblfieldlastvisitrpt  where (2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING)) >= '" & dm & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"

ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='SP' and fieldtype='F' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,STEMPEST, " _
& " (2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST)) as mper,'No','F','SP' " _
& " from  tblfieldlastvisitrpt  where (2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST)) >= '" & sp & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"

ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='RP' and fieldtype='F' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,ROOTPEST, " _
& " (2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST)) as mper,'No','F','RP' " _
& " from  tblfieldlastvisitrpt  where (2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST)) >= '" & rp & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"

ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='AP' and fieldtype='F' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,ACTIVEPEST, " _
& " (2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST)) as mper,'No','F','AP' " _
& " from  tblfieldlastvisitrpt  where (2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST)) >= '" & ap & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"

ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='AD' and fieldtype='F' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,ANIMALDAMAGE, " _
& " (2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE)) as mper,'No','F','AD' " _
& " from  tblfieldlastvisitrpt  where (2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE)) >= '" & ad & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"





ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='DM' and fieldtype='S' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,TREE_COUNT_DEADMISSING, " _
& " (2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING)) as mper,'No','S','DM' " _
& " from  tblstoragedlastvisitrpt  where (2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING)) >= '" & dm & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"


ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='ND' and fieldtype='S' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,ACTIVEPEST, " _
& " (2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST)) as mper,'No','S','ND' " _
& " from  tblstoragedlastvisitrpt  where (2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST)) >= '" & ap & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"



ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='PD' and fieldtype='S' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,ROOTPEST, " _
& " (2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST)) as mper,'No','S','PD' " _
& " from  tblstoragedlastvisitrpt  where (2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST)) >= '" & rp & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"


ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='DD' and fieldtype='S' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,STEMPEST, " _
& " (2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST)) as mper,'No','S','DD' " _
& " from  tblstoragedlastvisitrpt  where (2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST)) >= '" & sp & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"



ODKDB.Execute "delete from tbldailyemailfield where emailsent='No' and paratype='AD' and fieldtype='S' "
ODKDB.Execute "INSERT INTO tbldailyemailfield(uri,  visitdate, emaildate, staffcode, tshowog, " _
& " farmer, totaltrees, affected, mper, emailsent, fieldtype, paratype) " _
& " select _URI,'" & Format(Now - 1, "yyyy-MM-dd") & "','' emaildate,substring(staffbarcode,1,5),substring(farmerbarcode,1,9) tshowog, " _
& " substring(farmerbarcode,1,14) farmer,TREE_COUNT_TOTALTREES,ANIMALDAMAGE, " _
& " (2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE)) as mper,'No','S','AD' " _
& " from  tblstoragedlastvisitrpt  where (2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE)) >= '" & ad & "' and " _
& " SUBSTRING( start ,1,10)>='" & Format(Now - 1, "yyyy-MM-dd") & "' and " _
& " SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' and status<>'BAD' " _
& " and  _uri not in(select uri from tbldailyemailfield)"


End Sub
Private Sub createphonedailyactsummary()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim SQLSTR As String
Dim mact As String
Dim actmsg As String
Dim othermsg As String
'ODKDB.Execute "delete from tbldailyactivitysummary where emailsent='No' and "
Set rs = Nothing
rs.Open "select name from tbldailyactchoices where listname='dayactivity'", ODKDB
Do While rs.EOF <> True
actmsg = ""
If rs!name = "activity4" Then
Set rs1 = Nothing
rs1.Open "select count(*) as cnt,sum(fqc) fqc  from odk_prodlocal.tbldailyactivityemail where fqc>0 and (act1='" & rs!name & "' or act2='" & rs!name & "' or act3='" & rs!name & "')", ODKDB
If rs1.EOF <> True Then
If IIf(IsNull(rs1!cnt), 0, rs1!cnt) > 0 Then
    actmsg = rs1!fqc & " Fields by " & rs1!cnt & " Monitors"
End If
    Set RS2 = Nothing
    RS2.Open "select count(*) as cnt  from odk_prodlocal.tbldailyactivitysummary where activity='" & rs!name & "' and activitydate='" & Format(Now - 1, "yyyy-MM-dd") & "'", ODKDB
    If RS2!cnt > 0 Then
       ODKDB.Execute "update  odk_prodlocal.tbldailyactivitysummary set phonecomments='" & actmsg & "'where activity='" & rs!name & "' and activitydate='" & Format(Now - 1, "yyyy-MM-dd") & "'"
    Else
    ODKDB.Execute "insert into odk_prodlocal.tbldailyactivitysummary(activitydate,activity,phonecomments) values('" & Format(Now - 1, "yyyy-MM-dd") & "','" & rs!name & "','" & actmsg & "')"
    End If
End If
    
Set rs1 = Nothing
rs1.Open "select count(*) as cnt,sum(sqc) fqc  from odk_prodlocal.tbldailyactivityemail where sqc>0 and (act1='" & rs!name & "' or act2='" & rs!name & "' or act3='" & rs!name & "')", ODKDB
If rs1.EOF <> True Then
If IIf(IsNull(rs1!cnt), 0, rs1!cnt) > 0 Then
    actmsg = actmsg & "<br>" & rs1!fqc & " Storages by " & rs1!cnt & " Monitors"
    End If
    Set RS2 = Nothing
    RS2.Open "select count(*) as cnt  from odk_prodlocal.tbldailyactivitysummary where activity='" & rs!name & "' and activitydate='" & Format(Now - 1, "yyyy-MM-dd") & "'", ODKDB
    If RS2!cnt > 0 Then
    ODKDB.Execute "update  odk_prodlocal.tbldailyactivitysummary set phonecomments='" & actmsg & "'where activity='" & rs!name & "' and activitydate='" & Format(Now - 1, "yyyy-MM-dd") & "'"
    Else
    ODKDB.Execute "insert into odk_prodlocal.tbldailyactivitysummary(activitydate,activity,phonecomments) values('" & Format(Now - 1, "yyyy-MM-dd") & "','" & rs!name & "','" & actmsg & "')"
    End If

End If









Else

Set rs1 = Nothing
rs1.Open "select count(*) as cnt from odk_prodlocal.tbldailyactivityemail where  (act1='" & rs!name & "' or act2='" & rs!name & "' or act3='" & rs!name & "')", ODKDB
If rs1.EOF <> True Then
If IIf(IsNull(rs1!cnt), 0, rs1!cnt) > 0 Then
    actmsg = rs1!cnt & " Monitors"
    End If
    Set RS2 = Nothing
    RS2.Open "select count(*) as cnt  from odk_prodlocal.tbldailyactivitysummary where activity='" & rs!name & "' and activitydate='" & Format(Now - 1, "yyyy-MM-dd") & "'", ODKDB
    If RS2!cnt > 0 Then
    ODKDB.Execute "update  odk_prodlocal.tbldailyactivitysummary set phonecomments='" & actmsg & "'where activity='" & rs!name & "' and activitydate='" & Format(Now - 1, "yyyy-MM-dd") & "'"
    Else
    ODKDB.Execute "insert into odk_prodlocal.tbldailyactivitysummary(activitydate,activity,phonecomments) values('" & Format(Now - 1, "yyyy-MM-dd") & "','" & rs!name & "','" & actmsg & "')"
    End If

End If

End If



rs.MoveNext
Loop
ODKDB.Execute "delete from  tbldailyactivitysummary where length(odkcomments)=0 and length(phonecomments)=0"
'ODKDB.Execute "update tbldailyactchoices a, tbldailyactivitysummary b set activity=label where name=activity"
End Sub
Private Sub createodkdailyactsummary()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim SQLSTR As String
Dim mact As String
Dim actmsg As String
Dim othermsg As String
ODKDB.Execute "delete from tbldailyactivitysummary where emailsent='No'"
Set rs = Nothing
rs.Open "select value,count(value) cnt from dailyacthub9_activities where _PARENT_AURI in(select uri from tbldailyactivityemail where emailsent='No' and actualvisitdate='" & Format(Now - 1, "yyyy-MM-dd") & "') group by value  ", ODKDB
Do While rs.EOF <> True
actmsg = ""
If rs!Value = "activity4" Then
Set rs1 = Nothing
rs1.Open "select sum(field) fd ,count(staffbarcode) cn  from dailyacthub9_core where _uri in(select uri from tbldailyactivityemail where emailsent='No' and actualvisitdate='" & Format(Now - 1, "yyyy-MM-dd") & "'  and field>0 and activity like '%activity4%')", ODKDB
If rs1.EOF <> True Then
actmsg = IIf(IsNull(rs1!fd), 0, rs1!fd) & " Field by " & IIf(IsNull(rs1!cn), 0, rs1!cn) & " Monitors"
End If

Set rs1 = Nothing
rs1.Open "select sum(storage) st ,count(staffbarcode) cn from dailyacthub9_core where _uri in(select uri from tbldailyactivityemail  where emailsent='No' and actualvisitdate='" & Format(Now - 1, "yyyy-MM-dd") & "' and storage>0 and activity like '%activity4%')", ODKDB
If rs1.EOF <> True Then
actmsg = actmsg & "<br>" & IIf(IsNull(rs1!st), 0, rs1!st) & " Storage by " & IIf(IsNull(rs1!cn), 0, rs1!cn) & " Monitors"
End If

Else
actmsg = rs!cnt & " Monitors"
End If
ODKDB.Execute "insert into tbldailyactivitysummary(activitydate,activity,odkcomments) values('" & Format(Now - 1, "yyyy-MM-dd") & "','" & rs!Value & "','" & actmsg & "') "
rs.MoveNext
Loop
End Sub

Private Sub reocreatefirstpassofdailyactemail()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim SQLSTR As String
Dim mact As String
Dim mstsffcode As String
Dim actmsg As String
Dim othermsg As String
Dim activityday As String
tt = ""
ODKDB.Execute "delete from tblreodailyactivityemail where emailsent='No'"
SQLSTR = "select n.staffcode,start,_uri,activityday,OTHERCOMMENTS from reofvisit_core n INNER JOIN (SELECT staffcode,MAX(start)" _
         & " lastEdit FROM reofvisit_core where  _uri not in(select uri from tblreodailyactivityemail) GROUP BY staffcode)x ON " _
         & " n.staffcode = X.staffcode And n.Start = X.LastEdit " _
         & " AND STATUS <>  'BAD'  and _uri not in(select uri from tblreodailyactivityemail) " _
         & "  and SUBSTRING( start ,1,10)>='" & Format(Now - 20, "yyyy-MM-dd") & "' and SUBSTRING( start ,1,10)<='" & Format(Now, "yyyy-MM-dd") & "' " _
         & " GROUP BY n.staffcode order by n.staffcode,activityday"


Set rs = Nothing
rs.Open SQLSTR, ODKDB
If rs.EOF <> True Then
    Do Until rs.EOF
    mstsffcode = rs!staffcode
    activityday = rs!activityday
    If (activityday = "TOD") Then
    activityday = "Yesterday's Task"
    Else
     activityday = "Today's Task"
    End If
        
   Do While mstsffcode = rs!staffcode
        
         mact = ""
        actmsg = ""
        Set rs1 = Nothing
        rs1.Open "select * from reofvisit_activities where _PARENT_AURI='" & rs![_uri] & "'", ODKDB
        Do While rs1.EOF <> True
        mact = mact & rs1!Value & ","
        
        ' field visit
        If Trim(rs1!Value) = "activity1" Then
        actmsg = actmsg & "No. of farmers distributed: " & yesno(rs!distribution)
        If yesno(rs!distribution) = "No" Then
        actmsg = actmsg & ", Farmers left: " & rs!notvisited
        End If
        actmsg = actmsg & "<br>"
        End If
        
        'Compile weekly report
        If Trim(rs1!Value) = "activity2" Then
        actmsg = actmsg & "Actual number of farmers that received trees and should be present for the meeting: " & rs!NOREPRESENTATIVE & " and number of farmers who receives trees but did not have representative present for demo planting " & rs!NOREPRESENTATIVE1
        actmsg = actmsg & "<br>"
        End If
        
        'Give training
        If Trim(rs1!Value) = "activity3" Then
        actmsg = actmsg & " Individual Field Training: " & rs!individual
        actmsg = actmsg & "<br>"
        End If
        
        'Receive training
        If Trim(rs1!Value) = "activity4" Then
        If (rs!field) > 0 Then
        actmsg = actmsg & "Field QC: " & rs!field
        End If
        
        If rs!nofailed > 0 Then
        actmsg = actmsg & "," & " No. of failed field " & rs!nofailed
        End If
        
        If (rs!storage) > 0 Then
        actmsg = actmsg & " Storage QC: " & rs!storage
        End If
        
        If rs!nofailed1 > 0 Then
        actmsg = actmsg & "," & " No. of failed storage " & rs!nofailed1
        End If
        
        actmsg = actmsg & "<br>"
        End If
        
        If Trim(rs1!Value) = "activity5" Then
        actmsg = actmsg & " No. of farmers: " & rs!registered & " and acre registered" & rs!privateland
        actmsg = actmsg & "<br>"
        End If
        
        'On leave
        If Trim(rs1!Value) = "activity5" Then
        actmsg = actmsg & " Individual Field Training: " & rs!individual
        actmsg = actmsg & "<br>"
        End If
        
        'Other
        If Trim(rs1!Value) = "activity6" Then
        actmsg = actmsg & " Individual Field Training: " & rs!individual
        actmsg = actmsg & "<br>"
        End If
 
        rs1.MoveNext
        Loop
        

        
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
    
    
    
    Loop
    
End If

 Set rs = Nothing
 rs.Open "select staffcode from tblmhvstaff where status not in ('D','R','C') and dept='105' and staffcode not in (select monitor from odk_prodlocal.tblreodailyactivityemail where actualvisitdate='" & Format(Now - 1, "yyyy-MM-dd") & "' and emailsent='No')", MHVDB
 Do While rs.EOF <> True
 ODKDB.Execute "insert into tblreodailyactivityemail(actualvisitdate,monitor,entrytype)values('" & Format(Now - 1, "yyyy-MM-dd") & "','" & rs!staffcode & "','M')"
 rs.MoveNext
 Loop

 'ODKDB.Execute "update mhv.tblmhvstaff as a,odk_prodlocal.tbldailyactivityemail b set ms=MSUPERVISOR where staffcode=monitor"
' ODKDB.Execute "update odk_prodlocal.tbldailyactivityemail  set ms='' where length(ms)<>5"
 
End Sub
Function yesno(yn As String) As String
If Mid(UCase(yn), 1, 1) = "N" Then
yesno = "No"
Else
yesno = "Yes"
End If
End Function




Private Sub Command27_Click()
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select trnId,sum(crateqty) as qty from mhv.tblplanteddetail group by trnId", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update mhv.tblplanted set challanqty='" & rs!qty & "' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop
End Sub

Private Sub Command28_Click()
Dim SQLSTR As String
      SQLSTR = ""
      SQLSTR = "insert into tblfieldlastvisitrpt (_URI,START,end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,farmerbarcode,fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS,TREEHEIGHT) " _
      & "select _URI,START,n.end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,n.farmerbarcode,n.fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS,TREEHEIGHT from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
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

Private Sub Command29_Click()
Dim atleastone As Boolean
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim dm, ap, rp, sp, ad, cdm, cap, crp, csp, cad As Double
Dim dm1, ap1, rp1, sp1, ad1 As Double
Dim dm2, ap2, rp2, sp2, ad2 As Double
Dim SQLSTR As String
Dim mfarmerbarcode As String
Dim mfdcode As Integer
Set rs = Nothing
rs.Open "select * from tblphithreshold", ODKDB
If rs.EOF <> True Then
dm = rs!deadmissing
ap = rs!activepest
rp = rs!rootpest
sp = rs!stempest
ad = rs!animaldamage
cdm = rs!deadmissingchange
cap = rs!activepestchange
crp = rs!rootpestchange
csp = rs!stempestchange
cad = rs!animaldamagechange
End If



      SQLSTR = ""
      SQLSTR = "insert into tblfieldtofollowup (_URI,START,end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,farmerbarcode,fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS,TREEHEIGHT) " _
      & "select _URI,START,n.end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,n.farmerbarcode,n.fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS,TREEHEIGHT from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode, n.fdcode"
         
         
 ODKDB.Execute "delete from tblfieldtofollowup"
 ODKDB.Execute SQLSTR
 ODKDB.Execute "update tblfieldtofollowup set region_dcode=substring(farmerbarcode,1,3),region_gcode=substring(farmerbarcode,1,6), " _
 & " region=substring(farmerbarcode,1,9)"
 
 
 
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set  recordage=round(datediff(CURDATE(),END),0)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set startenddiff=round(datediff(END,START),0)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set activepercent=(TREE_COUNT_ACTIVEGROWING*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set deadpercent=(TREE_COUNT_DEADMISSING*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set stempestpercent=(STEMPEST*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set rootpestpercent=(ROOTPEST*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set activepestpercent=(ACTIVEPEST*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set adamagepercent=(ANIMALDAMAGE*100)/(TREE_COUNT_TOTALTREES)")

    
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phideadmissing=2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phistempest=2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phirootpest=2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phiactivepest=2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phianimaldamage=2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE)")


ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set dm='Y' where phideadmissing>='" & dm & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set sp='Y' where phistempest>='" & sp & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set rp='Y' where  phirootpest>='" & rp & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set ap='Y' where  phiactivepest>='" & ap & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set ad='Y'where phianimaldamage>='" & ad & "'")

'check records never show
Set rs = Nothing
rs.Open "select * from odk_prodlocal.tblfieldtofollowup where concat(farmerbarcode,fdcode) in(select concat(farmerbarcode,fdcode) from tblfieldfollowupnevershow)", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "delete from odk_prodlocal.tblfieldtofollowup where FARMERBARCODE='" & rs!farmerbarcode & "' and FDCODE='" & rs!FDCODE & "'"
rs.MoveNext
Loop

'check records not to show for this visit

Set rs = Nothing
rs.Open "select * from odk_prodlocal.tblfieldtofollowup where _URI in(select _URI from tblfieldfollowupdonotshowincurrentvisit)", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "delete from odk_prodlocal.tblfieldtofollowup where _URI='" & rs![_uri] & "'"
rs.MoveNext
Loop




'check records already flagged
dm1 = 0
ap1 = 0
rp1 = 0
sp1 = 0
ad1 = 0
mfdcode = 0
mfarmerbarcode = ""
Set rs = Nothing
rs.Open "select _URI,farmerbarcode,fdcode, " _
& " 2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING) as dm1, " _
& " 2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST) as sp1, " _
& " 2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST) as rp1, " _
& " 2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST) as ap1, " _
& " 2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE) as ad1 " _
& " from odk_prodlocal.tblfieldtofollowup where concat(farmerbarcode,fdcode) in(select concat(farmerbarcode,fdcode) from tblfieldforfollowup where status='ON')", ODKDB
Do While rs.EOF <> True
atleastone = False
mfarmerbarcode = rs!farmerbarcode
mfdcode = rs!FDCODE
dm1 = IIf(IsNull(rs!dm1), 0, rs!dm1)
ap1 = IIf(IsNull(rs!ap1), 0, rs!ap1)
rp1 = IIf(IsNull(rs!rp1), 0, rs!rp1)
sp1 = IIf(IsNull(rs!sp1), 0, rs!sp1)
ad1 = IIf(IsNull(rs!ad1), 0, rs!ad1)
dm2 = 0
ap2 = 0
rp2 = 0
sp2 = 0
ad2 = 0
Set rs1 = Nothing
rs1.Open "select * from tblfieldforfollowup where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "' and status='ON'", ODKDB
If rs1.EOF <> True Then
dm2 = dm1 - rs1!dm
ap2 = ap1 - rs1!ap
rp2 = rp1 - rs1!rp
sp2 = sp1 - rs1!sp
ad2 = ad1 - rs1!ad

If dm2 >= cdm And rs1!dm >= dm Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phideadmissing='" & rs1!dm & "', dmc='" & dm1 & "',dm='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If dm2 <= (-1 * cdm) And rs1!dm >= dm Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phideadmissing='" & rs1!dm & "', dmc='" & dm1 & "',dm='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If ap2 >= cap And rs1!ap >= ap Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phiactivepest='" & rs1!ap & "',apc='" & dm1 & "',ap='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If ap2 <= (-1 * cap) And rs1!ap >= ap Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phiactivepest='" & rs1!ap & "',apc='" & dm1 & "',ap='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If rp2 >= crp And rs1!rp >= rp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phirootpest='" & rs1!rp & "',rpc='" & dm1 & "',rp='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If rp2 <= (-1 * crp) And rs1!rp >= rp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phirootpest='" & rs1!rp & "',rpc='" & dm1 & "',rp='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If sp2 >= csp And rs1!sp >= sp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phistempest='" & rs1!sp & "',spc='" & dm1 & "',sp='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If sp2 <= (-1 * csp) And rs1!sp >= sp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phistempest='" & rs1!sp & "',spc='" & dm1 & "',sp='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If


If ad2 >= cad And rs1!ad >= ad Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phianimaldamage='" & rs1!ad & "', adc='" & ad1 & "',ad='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If ad2 <= (-1 * cad) And rs1!ad >= ad Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phianimaldamage='" & rs1!ad & "', adc='" & ad1 & "',ad='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If atleastone = False Then
ODKDB.Execute "delete from  tblfieldtofollowup where _URI='" & rs![_uri] & "' and farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
End If

End If

rs.MoveNext
Loop






ODKDB.Execute ("delete from  odk_prodlocal.tblfieldtofollowup where phideadmissing<'" & dm & "' and phistempest<'" & sp & "' and  phirootpest<'" & rp & "' and phiactivepest<'" & ap & "' and phianimaldamage<'" & ad & "' and nodelete<>'1'")
 


End Sub

Private Sub Command3_Click()
Dim rs As New ADODB.Recordset
Dim mm As String
Set rs = Nothing
MHVDB.Execute "delete from tbldistributionchecklist"

MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,distplannedqty)" _
& " SELECT year,mnth,distno, round((SUM( bcrate*35 ) + SUM( ecrate*35 ) + SUM( bno*35 ) + SUM( plno ) + SUM( crate " _
& " )),0) distqty  FROM  `tblplantdistributiondetail` WHERE  subtotindicator='' and status not in ('C','F') and " _
& " distno>0 and assignedAtField='' and trnid in (select trnid from tblplantdistributionheader where status='ON') " _
& " GROUP BY year ,mnth,distno"

MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,distplannedqty) values('2011','7',0,0)"
MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,distplannedqty) values('2012','8',0,0)"
MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,distplannedqty) values('2010','1',0,0)"
MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,distplannedqty) values('2013','1',0,0)"
MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,distplannedqty) values('2013','6',6,0)"
Set rs = Nothing
rs.Open "SELECT year,distno, round((SUM( bcrate*35 ) + SUM(ecrate*35 ) + SUM( bno*35 ) + SUM( plno ) + SUM( crate " _
& " )),0) distqty  FROM  `tblplantdistributiondetail` WHERE  subtotindicator='' and status not in ('C','F') and " _
& " distno>0 and assignedAtField='Y' and trnid in (select trnid from tblplantdistributionheader where status='ON') " _
& " GROUP BY year ,distno", MHVDB

Do While rs.EOF <> True

MHVDB.Execute "update tbldistributionchecklist set distunplannedqty='" & rs!distqty & "' where " _
& " distno='" & rs!distno & "' and year='" & rs!Year & "'"

rs.MoveNext
Loop


Set rs = Nothing
rs.Open "SELECT year,DistributionNo as distno, round((SUM(sendtofieldqty ) + SUM(shortexcessqty)) ,0) distqty  FROM  `tblqmssendtofieldhdr` WHERE  status not in ('C') " _
& " GROUP BY year ,DistributionNo", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistributionchecklist set billofladdingqty='" & rs!distqty & "' where " _
& " distno='" & rs!distno & "' and year='" & rs!Year & "'"
rs.MoveNext
Loop

'select * from tblqmsplanttransaction
Set rs = Nothing
rs.Open "SELECT distyear as year,distributionno as distno, round((SUM(credit)) ,0) distqty  FROM  `tblqmsplanttransaction` WHERE  status not in ('C') and  " _
& " transactiontype=4  GROUP BY year ,DistributionNo", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistributionchecklist set senttofieldqty='" & rs!distqty & "' where " _
& " distno='" & rs!distno & "' and year='" & rs!Year & "'"
rs.MoveNext
Loop



Set rs = Nothing
rs.Open "SELECT distyear as year,distributionno as distno, round((SUM(debit)) ,0) distqty  FROM  `tblqmsplanttransaction` WHERE  status not in ('C') and  " _
& " transactiontype=5  GROUP BY year ,DistributionNo", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistributionchecklist set backtonurseryqty    ='" & rs!distqty & "' where " _
& " distno='" & rs!distno & "' and year='" & rs!Year & "'"
rs.MoveNext
Loop


Set rs = Nothing
rs.Open "SELECT year,dno as distno, round((SUM(challanqty)) ,0) distqty  FROM  `tblplanted`  " _
& " GROUP BY year ,dno", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistributionchecklist set challanqty    ='" & rs!distqty & "' where " _
& " distno='" & rs!distno & "' and year='" & rs!Year & "'"
rs.MoveNext
Loop


MHVDB.Execute "insert into tbldistributionchecklist(year,mnth,distno,sendtostorage) " _
& "select '2013','12','999',sum(sendtofieldqty) from tblqmssendtotempstoragehdr group by year"
Set rs = Nothing
rs.Open "select distinct mnth as mnth from tbldistributionchecklist", MHVDB
Do While rs.EOF <> True
mm = MonthName(rs!mnth, True)
MHVDB.Execute "update tbldistributionchecklist set mnthname='" & mm & "' where mnth='" & rs!mnth & "'"
rs.MoveNext
Loop

End Sub

Private Sub Command30_Click()
Dim SQLSTR As String
      SQLSTR = ""
      SQLSTR = "insert into tblstoragedlastvisitrpt(_URI,START,end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,farmerbarcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS) " _
      & "select _URI, START,n.end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,n.farmerbarcode," _
      & " DTREES,0,0,ddamage,pdamage, " _
      & " ndtrees,GPS_LAT,GPS_LNG,GPS_ALT, " _
      & "GPS_ACC,TOTALTREES,wlogged,0,adamage, " _
      & "monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
         
      'ddamage(deasease damage)=STEMPEST
      'pdamage(Number of trees with pest damage)=ROOTPEST
      'ndtrees(Number of nutrient deficient trees)=ACTIVEPEST
      
 ODKDB.Execute "delete from tblstoragedlastvisitrpt"
 ODKDB.Execute SQLSTR
 ODKDB.Execute "update tblstoragedlastvisitrpt set region_dcode=substring(farmerbarcode,1,3),region_gcode=substring(farmerbarcode,4,3), " _
 & " region=substring(farmerbarcode,7,3)"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tbldzongkhag b set region_dcode=concat(substring(region_dcode,1,3),'  ',DzongkhagName) where substring(region_dcode,1,3)=DzongkhagCode"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tblgewog b set region_gcode=concat(substring(region_gcode,1,3),'  ',GewogName) where concat(substring(region_dcode,1,3),substring(region_gcode,1,3))=concat(DzongkhagId,GewogId)"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tbltshewog b set region=concat(substring(region,1,3),'  ',TshewogName) where concat(substring(region_dcode,1,3),substring(region_gcode,1,3),substring(region,1,3))=concat(DzongkhagId,GewogId,TshewogId)"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tblfarmer b set farmerbarcode=concat(farmerbarcode,'  ',farmername) where farmerbarcode=idfarmer"
MHVDB.Execute "update odk_prodlocal.tblstoragedlastvisitrpt a ,tblmhvstaff b set staffbarcode=concat(staffbarcode,'  ',staffname) where staffbarcode=staffcode"

End Sub

Private Sub Command31_Click()


Dim oSmtp As New EASendMailObjLib.Mail
    Set oSmtp = Nothing
    oSmtp.LicenseCode = "TryIt"
    oSmtp.FromAddr = "rigden506@gmail.com "
    oSmtp.ServerAddr = "smtp.tashicell.com"
    oSmtp.BodyFormat = 1
    oSmtp.ServerPort = 587
    oSmtp.SSL_starttls = 1
    oSmtp.SSL_init
    
     oSmtp.AddRecipientEx "mchhetri@mountainhazelnuts.com", 0
        oSmtp.Subject = Format(Now, "yyyyMMdd") & " " & "Monitoring report"
        oSmtp.BodyText = "<html><head><title></title></head><body>" & bodymsg & tbldailyactivitysummary & tbldailyemailfield & tbldailyemailstorage & "<br><br> <p>testing group mail on monitoring</p>Best Regards<br><br>Rinzin Lhamo" & "</body></html>"
       '<h3>" & bodymsg & "</h3>
       If oSmtp.SendMail() = 0 Then
       ' send ok
    MsgBox "Ok"
    Else
      MsgBox "Not Ok"
  
    End If
End Sub

Private Sub Command32_Click()
'............
Dim LastUpdateDate As Date

Dim connRemote As New ADODB.Connection
Dim CONNLOCAL As New ADODB.Connection
Dim ConnectionString As String

LastUpdateDate = Format(Now, "yyyy-MM-dd")
 connRemote.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=192.157.233.175;" _
                        & " DATABASE=odk_prod;" _
                        & "UID=odk_user;PWD=none; OPTION=3"
                        
OdkCnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & MsvrName & ";" _
                        & " DATABASE=odk_prodlocal;" _
                        & "UID=admin;PWD=password; OPTION=3"
                        
                        
CnnString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=" & MsvrName & ";" _
                        & " DATABASE=" & Mdbname & ";" _
                        & "UID=admin;PWD=password; OPTION=3"
                        
                        
                        
 connRemote.Open
 CONNLOCAL.Open OdkCnnString


'............

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs1 = Nothing
Dim imagecounter As Integer
Dim FPATH As String
FPATH = "C:\xampp\htdocs\mhweb\reoPic"

rs1.Open "select _URI from REOFVISIT_CORE  where substring(_CREATION_DATE,1,10)>='" & Format(LastUpdateDate - 10, "yyyy-MM-dd") & "'", connRemote, adOpenForwardOnly

Do While (rs1.EOF <> True)
imagecounter = 1
Set rs = Nothing
rs.Open "select * from REOFVISIT_FIELDIMAGE_BLB where _TOP_LEVEL_AURI='" & rs1![_uri] & "'", connRemote, adOpenStatic, adLockOptimistic
If rs.EOF <> True Then



Do While rs.EOF <> True

If Not Dir$(FPATH + "\" + Mid(rs![_TOP_LEVEL_AURI], 6, 600) & imagecounter & ".jpg") <> vbNullString Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
If Len(rs!Value) > 0 Then
mystream.Write rs!Value
mystream.SaveToFile "C:\xampp\htdocs\mhweb\reoPic\" & Mid(rs![_TOP_LEVEL_AURI], 6, 600) & imagecounter & ".jpg", adSaveCreateOverWrite
mystream.Close
End If

End If


imagecounter = imagecounter + 1
rs.MoveNext




Loop




End If

rs1.MoveNext
Loop


'.................

End Sub

Private Sub Command33_Click()
fixpersonregistering
End Sub
Private Sub fixpersonregistering()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim mreggroup As Integer
Dim mreggroupwith As Integer
Dim msupport As String
'fix individual
Set rs = Nothing
rs.Open "SELECT * FROM tbllandreg WHERE individual<>''and regdate>='2013-11-01' and regtype=''", MHVDB
Do While rs.EOF <> True
    Set rs1 = Nothing
    rs1.Open "select * from tblmhvstaff where staffcode='" & rs!individual & "'", MHVDB
    If rs1.EOF <> True Then
    mreggroup = rs1!reggroup
    Else
    mreggroup = "NA"
    End If
    
MHVDB.Execute "update tbllandreg set reggroup='" & mreggroup & "', reggroupwith='" & mreggroup & "',regtype='I' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop
'individual end

'fix outreach monitor
Set rs = Nothing
rs.Open "SELECT * FROM tbllandreg WHERE monitor<>''and outreach<>'' and  regdate>='2013-11-01' and regtype=''", MHVDB
Do While rs.EOF <> True

    Set rs1 = Nothing
    rs1.Open "select * from tblmhvstaff where staffcode='" & rs!monitor & "'", MHVDB
    If rs1.EOF <> True Then
    mreggroup = rs1!reggroup
    Else
    mreggroup = "NA"
    End If
    
     Set rs1 = Nothing
    rs1.Open "select * from tblmhvstaff where staffcode='" & rs!outreach & "'", MHVDB
    If rs1.EOF <> True Then
    mreggroupwith = rs1!reggroup
    Else
    mreggroupwith = "NA"
    End If
    
    
    
MHVDB.Execute "update tbllandreg set individual='" & rs!monitor & "',leadstaff='" & rs!outreach & "',reggroup='" & mreggroup & "', reggroupwith='" & mreggroupwith & "',regtype='S' where trnid='" & rs!trnid & "'"
MHVDB.Execute "update tbllandreg set monitor='',outreach='' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop


'fix leadstaff support
Set rs = Nothing
rs.Open "SELECT * FROM `tbllandreg` WHERE leadstaff<>'' and regdate>='2013-11-01' and regtype=''", MHVDB
Do While rs.EOF <> True

    Set rs1 = Nothing
    rs1.Open "select * from tblmhvstaff where staffcode='" & rs!LEADSTAFF & "'", MHVDB
    If rs1.EOF <> True Then
    mreggroup = rs1!reggroup
    Else
    mreggroup = "NA"
    End If
    
     Set rs1 = Nothing
    rs1.Open "select * from tblmhvstaff where staffcode='" & rs!SUPPORT1 & "'", MHVDB
    If rs1.EOF <> True Then
    mreggroupwith = rs1!reggroup
    Else
    mreggroupwith = "NA"
    End If

msupport = rs!SUPPORT1
If Len(rs!SUPPORT2) > 0 Then
msupport = msupport & "," & rs!SUPPORT2
End If
If Len(rs!SUPPORT3) > 0 Then
msupport = msupport & "," & rs!SUPPORT3
End If
If Len(rs!SUPPORT4) > 0 Then
msupport = msupport & "," & rs!SUPPORT4
End If
If Len(rs!SUPPORT5) > 0 Then
msupport = msupport & "," & rs!SUPPORT5
End If

MHVDB.Execute "update tbllandreg set individual='" & rs!LEADSTAFF & "',SUPPORT1='" & msupport & "',reggroup='" & mreggroup & "', reggroupwith='" & mreggroupwith & "',regtype='G' where trnid='" & rs!trnid & "'"
MHVDB.Execute "update tbllandreg set leadstaff='',SUPPORT2='',SUPPORT3='',SUPPORT4='',SUPPORT5='' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop



MsgBox "Done"
End Sub

Private Sub Command34_Click()
Exit Sub
'createnewdashboardworkspace
End Sub

Private Sub Command35_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim mdays As Integer
Dim preceived As Integer
Dim prepreceived As Integer
Dim premonth As Integer
Dim preyear As Integer
Dim mchange As Integer
Dim currmonth, curryear As Integer
currmonth = Month(Now)
curryear = Year(Now)

ODKDB.Execute "delete from tblmonitormonthlyperformance where myear='" & curryear & "' and mmonth='" & currmonth & "'"
ODKDB.Execute "insert into tblmonitormonthlyperformance(id,myear,mmonth,staffid,ms) select id,myear, " _
& " mmonth,staffid,ms from vmonitormonthlyperformance where myear='" & curryear & "' and mmonth='" & currmonth & "'"


ODKDB.Execute "update tblmonitormonthlyperformance set  " _
& " farmerassigned=(select count(*) cnt from mhv.tblfarmer where " _
& " MONITOR=staffid)"

Set rs = Nothing



Set rs = Nothing
rs.Open "select * from tblmonitormonthlyperformance where myear='" & curryear & "' and mmonth='" & currmonth & "' order by myear,mmonth,staffid", ODKDB
Do While rs.EOF <> True

If rs!mmonth = 1 Then
premonth = 12
preyear = rs!myear - 1
Else
premonth = rs!mmonth - 1
preyear = rs!myear
End If


prepreceived = 0
mdays = 0
preceived = 0
mchange = 0
Set RS2 = Nothing

Set RS2 = ODKDB.Execute("select count(*) cnt from dailyacthub9_core n INNER JOIN " _
& "(SELECT staffbarcode,MAX(end) lastEdit FROM dailyacthub9_core " _
& " where year(end)='" & preyear & "' and month(end)='" & premonth & "' and staffbarcode='" & rs!staffid & "' " _
& " GROUP BY staffbarcode,date(end))x ON n.staffbarcode = X.staffbarcode And n.end = X.LastEdit " _
& " AND STATUS <> 'BAD'")
If RS2.EOF <> True Then
prepreceived = Round((RS2!cnt * 100) / Days(premonth, preyear), 0)
End If

'ODKDB.Execute "delete from tblmonitormonthlyperformance where myear='" & rs!myear & "' and mmonth='" & rs!mmonth & "' and staffid='" & rs!staffid & "'"
Set rs1 = Nothing

Set rs1 = ODKDB.Execute("select count(*) cnt from dailyacthub9_core n INNER JOIN " _
& "(SELECT staffbarcode,MAX(end) lastEdit FROM dailyacthub9_core " _
& " where year(end)='" & rs!myear & "' and month(end)='" & rs!mmonth & "' and staffbarcode='" & rs!staffid & "' " _
& " GROUP BY staffbarcode,date(end))x ON n.staffbarcode = X.staffbarcode And n.end = X.LastEdit " _
& " AND STATUS <> 'BAD'")


If rs1.EOF <> True Then
mdays = Days(rs!mmonth, rs!myear) - rs1!cnt
preceived = Round((rs1!cnt * 100) / Days(rs!mmonth, rs!myear), 0)
mchange = IIf(prepreceived = 0, 0, IIf(preceived - prepreceived < 0, preceived - prepreceived, "+" & preceived - prepreceived))
ODKDB.Execute "update tblmonitormonthlyperformance set  " _
& " dailyactivityreceived='" & rs1!cnt & "' ,dailyactivitynotreceived='" & mdays & "', pdailyactivityreceived='" & preceived & "',pdailyactivityreceivedchange='" & mchange & "' where myear='" & rs!myear & "' and mmonth='" & rs!mmonth & "' and staffid='" & rs!staffid & "'"
End If


'land reg


prepreceived = 0
mdays = 0
preceived = 0
mchange = 0

Set RS2 = Nothing

Set RS2 = MHVDB.Execute("select target,sum(regland) regland  from tblregistrationrpt where year(regdate)='" & preyear & "' and month(regdate)='" & premonth & "' and staffcode='" & rs!staffid & "'")
If RS2.EOF <> True Then
If RS2!Target > 0 Then
prepreceived = Round((RS2!regland * 100) / RS2!Target, 0)
End If

End If



Set rs1 = Nothing
Set rs1 = MHVDB.Execute("select target,sum(regland) regland  from tblregistrationrpt where year(regdate)='" & rs!myear & "' and month(regdate)='" & rs!mmonth & "' and staffcode='" & rs!staffid & "'")
If rs1.EOF <> True Then
If rs1!Target > 0 Then
preceived = Round((rs1!regland * 100) / rs1!Target, 0)
End If
mchange = IIf(prepreceived = 0, 0, IIf(preceived - prepreceived < 0, preceived - prepreceived, "+" & preceived - prepreceived))
ODKDB.Execute "update tblmonitormonthlyperformance set  " _
& " target='" & rs1!Target & "' ,acrereg='" & rs1!regland & "', pacrereg='" & preceived & "',pacreregchange='" & mchange & "' where myear='" & rs!myear & "' and mmonth='" & rs!mmonth & "' and staffid='" & rs!staffid & "'"
End If



prepreceived = 0
mdays = 0
preceived = 0
mchange = 0

If rs!myear >= 2014 Then


Set rs1 = Nothing
Set rs1 = ODKDB.Execute("select fieldvisit,storagevisit,nooffarmersvisted,farmersnotvisited,percentage,nooffarmersvisted  from tblfarmvisit where myear='" & rs!myear & "' and mmonth='" & rs!mmonth & "' and staffbarcode='" & rs!staffid & "'")
If rs1.EOF <> True Then

mchange = IIf(prepreceived = 0, 0, IIf(preceived - prepreceived < 0, preceived - prepreceived, "+" & preceived - prepreceived))
ODKDB.Execute "update tblmonitormonthlyperformance set  " _
& " nooffieldvisited='" & rs1!fieldvisit & "' , " _
& " noofstoragevisited='" & rs1!storagevisit & "', " _
& " totalvisited='" & rs1!nooffarmersvisted & "', " _
& " notvisited='" & rs1!farmersnotvisited & "', " _
& " pvisitoverdue='" & 100 - rs1!percentage & "', " _
& " pfieldvisitcoverage='" & rs1!percentage & "' " _
& " where myear='" & rs!myear & "' and mmonth='" & rs!mmonth & "' and staffid='" & rs!staffid & "'"
End If




End If








rs.MoveNext
Loop




End Sub

Private Function Days(pMonth As Integer, pYear As Integer) As Long

  Select Case pMonth

    Case 2
      If (pYear Mod 4 = 0) Then Days = 29 Else Days = 28

    Case 4, 6, 9, 11
      Days = 30

    Case Else
      Days = 31

  End Select

End Function
Private Sub farmvisit(myear As Integer, mmonth As Integer, staffid As String)
Dim rs As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim dgtstr As String
Dim fieldvisit As Integer

dgtstr = ""
fieldvisit = 0
Set rsm = Nothing
rsm.Open "select distinct(substring(idfarmer,1,9)) as dgt from tblfarmer where monitor='" & staffid & "'", MHVDB
Do While rsm.EOF <> True


dgtstr = dgtstr + "'" + Trim(rsm!dgt) + "',"
rsm.MoveNext
Loop
If Len(dgtstr) > 0 Then
dgtstr = "(" + Left(dgtstr, Len(dgtstr) - 1) + ")"
Else
dgtstr = "(" + "'" + A99 & "'" & ")"
End If



Set rs = Nothing
rs.Open "select distinct visitfrequency from tblfarmer where monitor='" & staffid & "'"
Do While rs.EOF <> True

Set rs1 = Nothing
'rs1.Open "select count(distinct farmerbarcode) cnt from phealthhub15_core where substring(farmerbarcode,1,9) in " & dgtstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs1.Open "select count(distinct farmerbarcode) cnt from phealthhub15_core where substring(farmerbarcode,1,9) in " & dgtstr & " and round(datediff(CURDATE(),END)-2,0)<='" & rs!visitfrequency & "'", ODKDB
If rs1.EOF <> True Then
fieldvisit = fieldvisit + rs1!cnt
End If



rs.MoveNext
Loop





End Sub

Private Sub Command36_Click()
Dim rs As New ADODB.Recordset
Dim rscheckearlier As New ADODB.Recordset
Dim rsnewacre As New ADODB.Recordset
Dim rsdist As New ADODB.Recordset
Set rs = Nothing
Dim tempfarmer, newfarmer, oldfarmer As String
Dim newacre, oldacre, newdist, olddist, filling As Double
Dim tempyear As Integer
MHVDB.Execute "delete from tblseantemp"
rs.Open "select distinct year,farmercode from tblplanted order by year", MHVDB
Do While rs.EOF <> True
tempyear = rs!Year
tempfarmer = rs!farmercode
Set rscheckearlier = Nothing
rscheckearlier.Open "select * from tblseantemp where newfarmers='" & tempfarmer & "'", MHVDB
If rscheckearlier.EOF <> True Then
oldfarmer = tempfarmer
newfarmer = ""
Set rsdist = Nothing
rsdist.Open "select sum(challanqty) challanqty from tblplanted where farmercode='" & tempfarmer & "' and year='" & tempyear & "'", MHVDB
If rsdist.EOF <> True Then
olddist = rsdist!challanqty
newdist = 0
End If
Else
newfarmer = tempfarmer
oldfarmer = ""
Set rsdist = Nothing
rsdist.Open "select sum(challanqty) challanqty from tblplanted where farmercode='" & tempfarmer & "' and year='" & tempyear & "'", MHVDB
If rsdist.EOF <> True Then
newdist = rsdist!challanqty
olddist = 0
End If

End If

Set rsnewacre = Nothing
rsnewacre.Open "select sum(regland) regland from tbllandreg where FARMERID='" & tempfarmer & "'", MHVDB
If rsnewacre.EOF <> True Then
newacre = rsnewacre!regland
oldacre = 0
End If

MHVDB.Execute "insert into tblseantemp(myear,newfarmers,newacre," _
& " newdistribute,oldfarmers,oldacre, " _
& " olddistribute,fillings) values('" & tempyear & "','" & newfarmer & "','" & newacre & "'," _
& " '" & newdist & "','" & oldfarmer & "',0,'" & olddist & "',0)"


rs.MoveNext

Loop


End Sub

Private Sub Command37_Click()

Dim SQLSTR As String
Dim frcode As String
Dim j As Integer
Dim rs As New ADODB.Recordset

Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
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
   SQLSTR = "insert into mtemp SELECT _URI, substring(tcode,1,3) region_dcode, substring(tcode,4,3) region_gcode, substring(tcode,7,3) region,fcode FROM dschoolsurvey_core where farmerbarcode=''"
  db.Execute SQLSTR
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
    db.Execute "update dschoolsurvey_core set farmerbarcode='" & mfcode & "' where " _
    & " substring(tcode,1,9)='" & rss!dcode & rss!gcode & rss!tcode & "' and " _
    & " fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"
  
  rss.MoveNext
  Loop



End Sub

Private Sub Command38_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs1 = Nothing
Dim FPATH As String
Dim mfilename As String
Dim i As Integer
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
FPATH = "C:\xampp\htdocs\mhweb\dsPic"

rs1.Open "select _URI,farmerbarcode from dschoolsurvey_core", db, adOpenDynamic
                       
Do While (rs1.EOF <> True)
mfilename = rs1!farmerbarcode
Set rs = Nothing
rs.Open "select * from dschoolsurvey_houseimage_blb where _TOP_LEVEL_AURI='" & rs1![_uri] & "'", db, adOpenStatic, adLockOptimistic
If rs.EOF <> True Then

If Not Dir$(FPATH + "\" + mfilename & ".jpg") <> vbNullString Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
If Len(rs!Value) > 0 Then
mystream.Write rs!Value
mystream.SaveToFile "C:\xampp\htdocs\mhweb\dsPic\" & mfilename & ".jpg", adSaveCreateOverWrite
mystream.Close
End If
Else


Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
If Len(rs!Value) > 0 Then
mystream.Write rs!Value
mystream.SaveToFile "C:\xampp\htdocs\mhweb\dsPic\" & mfilename & "-" & Mid(rs![_TOP_LEVEL_AURI], 6, 600) & ".jpg", adSaveCreateOverWrite
mystream.Close
End If
End If

End If

rs1.MoveNext
Loop
End Sub

Private Sub Command4_Click()
Dim excel_app As Object
Dim excel_sheet As Object
Dim farmerstr As String
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim rsnewold As New ADODB.Recordset
Dim rsgp As New ADODB.Recordset
Dim rss As New ADODB.Recordset

farmerstr = ""



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
    excel_sheet.cells(3, 1) = "SL.NO."
    excel_sheet.cells(3, 2) = "DZONGKHAG"
    excel_sheet.cells(3, 3) = "GEWOG"
    excel_sheet.cells(3, 4) = "TSHOWOG"
    excel_sheet.cells(3, 5) = "FARMER NAME"
    excel_sheet.cells(3, 6) = "gnew"
    excel_sheet.cells(3, 7) = "gold"
    excel_sheet.cells(3, 8) = "gtot"
    excel_sheet.cells(3, 9) = "pnew"
    excel_sheet.cells(3, 10) = "pold"
    excel_sheet.cells(3, 11) = "ptot"
    excel_sheet.cells(3, 12) = "tot"
    i = 4





Set rs = Nothing
rs.Open "SELECT farmerid, SUM( regland ) regland, plantedstatus" _
& " FROM  `tbllandreg` where farmerid in(select idfarmer from tblfarmer where status not in('D','R'))" _
& " GROUP BY farmerid, plantedstatus " _
& " ORDER BY farmerid", MHVDB



    Do Until rs.EOF
    
  
    
    farmerstr = rs!farmerid
    FindDZ Mid(rs!farmerid, 1, 3)
    FindGE Mid(rs!farmerid, 1, 3), Mid(rs!farmerid, 4, 3)
    FindTs Mid(rs!farmerid, 1, 3), Mid(rs!farmerid, 4, 3), Mid(rs!farmerid, 7, 3)
    FindFA rs!farmerid, "F"
    excel_sheet.cells(i, 2) = Dzname
    excel_sheet.cells(i, 3) = GEname
    excel_sheet.cells(i, 4) = TsName
    excel_sheet.cells(i, 5) = FAName
    
    
    Do While farmerstr = rs!farmerid
 
    Set rsnewold = Nothing
    rsnewold.Open "select * from tblplanted where farmercode='" & rs!farmerid & "'", MHVDB
    If rsnewold.EOF <> True And rs!plantedstatus = "C" Then
    'old
    If Mid(rs!farmerid, 10, 1) = "G" Then
    ' grf
    excel_sheet.cells(i, 7) = rs!regland
    Else
    'private
    excel_sheet.cells(i, 10) = rs!regland
    End If
    
    
    Else
    'new
      If Mid(rs!farmerid, 10, 1) = "G" Then
    ' grf
    excel_sheet.cells(i, 6) = rs!regland
    Else
    'private
    excel_sheet.cells(i, 9) = rs!regland
    End If
    End If
    
 
   
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
    
    
     sl = sl + 1
    i = i + 1
    Loop




'make up
   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 7)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:G3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ERROR LISTING"
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

Private Sub Command40_Click()
updatePollen
End Sub
Public Sub updatePollen()

Dim SQLSTR As String
Dim frcode As String
Dim j As Integer
Dim rs As New ADODB.Recordset

Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
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
Dim tempstr, updateStr As String
SQLSTR = ""
updateStr = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                    

db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp SELECT _URI, dcode, gcode, tcode,fcode,FARMERTYPE FROM tbleconomicsurvey_core where farmerbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = UCase(Mid(rss!dcode, 1, 3))
  mgcode = UCase(Mid(rss!gcode, 4, 3))
  mtcode = UCase(Mid(rss!tcode, 7, 3))
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = rss!fType & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  Set rsF = Nothing
  'updateStr = "update tblfieldqc_core set farmerbarcode='" & mfcode & "' where dcode='" & rss!dcode & "' and gcode='" & rss!gcode & "' and tcode='" & rss!tcode & "' and fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"
   updateStr = "update tbleconomicsurvey_core set farmerbarcode='" & mfcode & "',dcode=substring(farmerbarcode,1,3),gcode=substring(farmerbarcode,1,6),tcode=substring(farmerbarcode,1,9) where   _URI='" & rss![_uri] & "'"
  ' MsgBox updateStr
    db.Execute updateStr
  frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
'db.Execute "update tblpolination_core set dcode=substring(farmerbarcode,1,3),gcode=substring(farmerbarcode,1,6),tcode=substring(farmerbarcode,1,9) "
MsgBox "completed"






End Sub

Private Sub Command41_Click()
Dim SQLSTR As String
Dim farmerstr As String
Dim mcode As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString

    GetTbl
    SQLSTR = ""
    SQLSTR = "insert into " & Mtblname & " (start,tdate,end,farmercode,fdcode,fs) select n.end, n.end,n.end,n.farmerbarcode,n.fdcode,'F' " _
            & " from tblfieldqc_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
            & "lastEdit FROM tblfieldqc_core GROUP BY farmerbarcode,fdcode)x ON " _
            & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
            & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
            db.Execute SQLSTR
            
     SQLSTR = ""
     farmerstr = ""
Set rs = Nothing
rs.Open "select distinct farmercode from " & Mtblname & " ", db
Do While rs.EOF <> True
farmerstr = farmerstr + "'" + Trim(rs!farmercode) + "',"
rs.MoveNext
Loop
If Len(farmerstr) > 0 Then
farmerstr = "(" + Left(farmerstr, Len(farmerstr) - 1) + ")"
Else
farmerstr = "(" + "'" + A99 & "'" & ")"
End If



     SQLSTR = "insert into " & Mtblname & " (start,tdate,end,farmercode,fs)  select n.end,n.end, n.end,n.farmerbarcode,'S' from " _
            & "tblstorageqc_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
            & "lastEdit FROM tblstorageqc_core GROUP BY farmerbarcode)x ON " _
            & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
            & "AND STATUS <>  'BAD' and n.farmerbarcode  not in " & farmerstr & " GROUP BY n.farmerbarcode"
        
        
     db.Execute SQLSTR
     

SQLSTR = ""
farmerstr = ""
Set rs = Nothing

rs.Open "select distinct farmercode from " & Mtblname & " ", db
Do While rs.EOF <> True
farmerstr = farmerstr + "'" + Trim(rs!farmercode) + "',"
rs.MoveNext
Loop
If Len(farmerstr) > 0 Then
farmerstr = "(" + Left(farmerstr, Len(farmerstr) - 1) + ")"
Else
farmerstr = "(" + "'" + A99 & "'" & ")"
End If
SQLSTR = "select * from tblplanted where farmercode not in " & farmerstr & ""
Set rs = Nothing
rs.Open SQLSTR, MHVDB
Do While rs.EOF <> True

Set rs1 = Nothing
rs1.Open "select * from tblzerovisit where farmercode='" & rs!farmercode & "'", MHVDB
If rs1.EOF <> True Then

'nothing
Else
'insert

Set rsm = Nothing
rsm.Open "select * from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
    If rsm.EOF <> True Then
        FindFA rs!farmercode, "F"
        FindsTAFF rsm!monitor
            If Len(sTAFF) = 0 Then
            sTAFF = "Monitor Not Assigned"
            End If
        MHVDB.Execute "insert into tblzerovisit(farmercode,farmername,staffcode,staffname,cnt,status)" _
        & "values('" & rs!farmercode & "','" & FAName & "','" & rsm!monitor & "','" & sTAFF & "','1','Active')"
    
    End If

End If
rs.MoveNext
Loop




Set rs = Nothing
rs.Open "select * from tblzerovisit", MHVDB
Do While rs.EOF <> True
Set rs1 = Nothing

rs1.Open "select distinct farmerbarcode from tblfieldqc_core where farmerbarcode='" & rs!farmercode & "'", db
If rs1.EOF <> True Then
'delete from zero visit

MHVDB.Execute "delete from tblzerovisit where farmercode='" & rs1!farmerbarcode & "'"
End If


rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select * from tblzerovisit", MHVDB
Do While rs.EOF <> True
Set rs1 = Nothing
rs1.Open "select distinct farmerbarcode from tblstorageqc_core where farmerbarcode='" & rs!farmercode & "'", db
If rs1.EOF <> True Then
'delete from zero visit
MHVDB.Execute "delete from tblzerovisit where farmercode='" & rs1!farmerbarcode & "'"
End If
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select * from tblzerovisit where staffname='Monitor Not Assigned'", MHVDB
Do While rs.EOF <> True
Set rs1 = Nothing
rs1.Open "select * from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
If rs1.EOF <> True Then

        FindsTAFF rs1!monitor
            If Len(sTAFF) = 0 Then
            sTAFF = "Monitor Not Assigned"
            End If
        MHVDB.Execute "update tblzerovisit set staffcode='" & rs1!monitor & "' ,staffname='" & sTAFF & "' where farmercode='" & rs!farmercode & "'"
    
End If
rs.MoveNext
Loop



MHVDB.Execute "delete from tblzerovisit where farmercode in(select idfarmer from tblfarmer where status in('D','R'))"
MsgBox "comp"
End Sub

Private Sub Command42_Click()
MsgBox Format(Now(), "ww", vbSunday, vbFirstFullWeek)
End Sub

Private Sub Command5_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim SQLSTR As String
SQLSTR = ""
MHVDB.Execute "delete from tbldistplan"
SQLSTR = "insert into tbldistplan SELECT SUBSTRING(IDFARMER,1,3)dzcode,SUBSTRING(IDFARMER,4,3)gecode, " _
& " SUBSTRING(IDFARMER,7,3)tscode,IDFARMER,FARMERNAME,0 REGLAND, " _
& " village,phone1,'' platedstatus,'' dgt,'',1,'','','','' FROM tblfarmer  WHERE status not in('D','R')"

MHVDB.Execute SQLSTR

Set rs = Nothing
rs.Open "select farmerid,regdate,plantedstatus,sum(regland) regland from tbllandreg " _
& " where plantedstatus in('N','P') and status not in('R','D') group by farmerid", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistplan set regland='" & rs!regland & "', " _
& " plantedstatus='" & rs!plantedstatus & "',regdate='" & Format(rs!regdate, "yyyy-MM-dd") & "' where idfarmer='" & rs!farmerid & "' "
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select * from tbllandregdetail where plantedstatus in('N')", MHVDB
Do While rs.EOF <> True
MHVDB.Execute "update tbldistplan set regland='" & rs!acre & "' where  plantedstatus='P' and idfarmer='" & rs!farmercode & "' "
rs.MoveNext
Loop
MHVDB.Execute "delete from tbldistplan where regland=0"

Set rs = Nothing
rs.Open "select * from tbldistplan ", MHVDB
Do While rs.EOF <> True
FindDZ rs!dzcode
FindGE rs!dzcode, rs!GECODE
FindTs rs!dzcode, rs!GECODE, rs!tscode
MHVDB.Execute "update tbldistplan set dzcode='" & rs!dzcode & "  " & Dzname & "', " _
& " gecode='" & rs!GECODE & "  " & GEname & "', " _
& " tscode='" & rs!tscode & "  " & TsName & "', " _
& "dgt='" & rs!dzcode & "  " & Dzname & "  " & rs!GECODE & "  " & GEname & "  " & rs!tscode & "  " & TsName & "'" _
& "where idfarmer='" & rs!idfarmer & "' "
rs.MoveNext
Loop

MHVDB.Execute "update tbldistplan set landtype='Private' where substring(idfarmer,10,1)='F'"
MHVDB.Execute "update tbldistplan set landtype='GRF/SRF' where substring(idfarmer,10,1)='G'"
MHVDB.Execute "update tbldistplan set landtype='CF' where substring(idfarmer,10,1)='C'"
MHVDB.Execute "update tbldistplan set farmertype='New'"
MHVDB.Execute "update tbldistplan set farmertype='Old' where idfarmer in(select farmercode from tblplanted)"



 MHVDB.Execute "update tbldistplan set monitor='',email=''"
Set rs = Nothing
rs.Open "select distinct substring(idfarmer,1,9) as dgt from tbldistplan", MHVDB
Do While rs.EOF <> True
Set rs1 = Nothing
rs1.Open "select monitor from tblfarmer where substring(idfarmer,1,9)='" & rs!dgt & "'", MHVDB
If rs1.EOF <> True Then
FindsTAFF rs1!monitor

MHVDB.Execute "update tbldistplan set monitor='" & rs1!monitor & "  " & sTAFF & "',email='" & emailAddress & "' " _
& " where substring(idfarmer,1,9)='" & rs!dgt & "'"

End If
rs.MoveNext
Loop

End Sub

Private Sub updateqcofqc()
Dim SQLSTR As String
Dim frcode As String
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Dim MSTR As String
Set rsadd = Nothing
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
   SQLSTR = "insert into mtemp SELECT _URI, dcode, gcode, tcode,fcode FROM monitor_qcv5_field where farmerbarcode=''"
  db.Execute SQLSTR
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
    db.Execute "update monitor_qcv5_field set farmerbarcode='" & mfcode & "' where dcode='" & rss!dcode & "' and gcode='" & rss!gcode & "' and tcode='" & rss!tcode & "' and fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"
  frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
LogRemarks = "table monitor_qcv5_field updated successfully.farmerbarcode updated(" & frcode & ")"
updateodklog "no uri", Now, MUSER, LogRemarks, "phealthhub15_core"

Set rs = Nothing
rs.Open "select * from monitor_qcv5_core where length(mid)='3'", ODKDB
Do While rs.EOF <> True
MSTR = ""
MSTR = "S0" & rs!Mid
FindsTAFF MSTR
ODKDB.Execute "update monitor_qcv5_core set mid='" & MSTR & "  " & sTAFF & "' where mid='" & rs!Mid & "'"
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select * from monitor_qcv5_core where length(sid)='3'", ODKDB
Do While rs.EOF <> True
MSTR = ""
MSTR = "S0" & rs!sid
FindsTAFF MSTR
ODKDB.Execute "update monitor_qcv5_core set sid='" & MSTR & "  " & sTAFF & "' where sid='" & rs!sid & "'"
rs.MoveNext
Loop



Set rs = Nothing
rs.Open "select * from monitor_qcv5_field where length(farmerbarcode)='14'", ODKDB
Do While rs.EOF <> True
MSTR = ""
MSTR = rs!farmerbarcode
FindFA MSTR, "F"
ODKDB.Execute "update monitor_qcv5_field set farmerbarcode='" & MSTR & "  " & FAName & "' where farmerbarcode='" & rs!farmerbarcode & "'"
rs.MoveNext
Loop
End Sub

Private Sub Command6_Click()
updateqcofqc
End Sub

Private Sub Command7_Click()
Dim rs1 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim farmerstr As String
Dim monstr As String
Dim regdate As Date
Dim cnt As Integer


Set rs = Nothing
rs.Open "select * from tblregistrationrpt where length(farmercode)>1 order by farmercode", MHVDB

Do Until rs.EOF
    
    cnt = 1
    regdate = Format(rs!regdate, "yyyy-MM-dd")
    farmerstr = Mid(rs!farmercode, 1, 14)
    monstr = Mid(rs!staffcode, 1, 5)
    
    Do While farmerstr = Mid(rs!farmercode, 1, 14)
    Set rs1 = Nothing
      
    Set rs1 = Nothing
    rs1.Open "select * from tblregistrationrpt where substring(farmercode,1,14)='" & Mid(rs!farmercode, 1, 14) & "' and cnt='1'", MHVDB
    If rs1.EOF <> True Then
   
    Else
     MHVDB.Execute "update tblregistrationrpt set cnt='1' where substring(farmercode,1,14)='" & Mid(rs!farmercode, 1, 14) & "' and substring(staffcode,1,14)='" & Mid(rs!staffcode, 1, 14) & "' and regdate='" & Format(rs!regdate, "yyyy-MM-dd") & "' "
    
    End If
    
    
   
    
    
        
    
    
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
    
    
    
    Loop

End Sub

Private Sub Command8_Click()
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
ODKDB.Execute "delete from tblextensionmortality"
        SQLSTR = ""

           
         SQLSTR = "insert into tblextensionmortality (dzongkhag,gewog,tshowog,farmercode,fieldcode,lat,lng,alt,acrereg,acrecultivated,totaltrees,deadmissing,percent,end) " _
         & " select substring(n.farmerbarcode,1,3)," _
         & "substring(n.farmerbarcode,4,3),substring(n.farmerbarcode,7,3),n.farmerbarcode,n.fdcode,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG," _
         & "GPS_COORDINATES_ALT,0,TREE_COUNT_TOTALTREES/420,TREE_COUNT_TOTALTREES,tree_count_deadmissing,((tree_count_deadmissing*100)/TREE_COUNT_TOTALTREES),end from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
         
 ODKDB.Execute SQLSTR

 
 Set rs = Nothing
 rs.Open "select * from tblextensionmortality", ODKDB
 Do While rs.EOF <> True
 Set rs1 = Nothing
 rs1.Open "select sum(regland) as regland from tbllandreg where  farmerid='" & rs!farmercode & "'", MHVDB
  ODKDB.Execute "update tblextensionmortality set acrereg='" & IIf(IsNull(rs1!regland), 0, rs1!regland) & "' where farmercode='" & rs!farmercode & "'"
 rs.MoveNext
 Loop
 
 
 ODKDB.Execute "update tblextensionmortality set grpgewog=concat(dzongkhag,gewog)"
 ODKDB.Execute "update tblextensionmortality set grptshowog=concat(dzongkhag,gewog,tshowog)"
 
 
 ODKDB.Execute "update tblextensionmortality a ,mhv.tbldzongkhag b set " _
 & " dzongkhag= concat(dzongkhagcode,'  ',dzongkhagname) where dzongkhag=b.dzongkhagcode"
 
  
 ODKDB.Execute "update tblextensionmortality a ,mhv.tblgewog b set " _
 & " grpgewog= concat(grpgewog,'  ',gewogname) where grpgewog=concat(dzongkhagid,gewogid)"
 
  ODKDB.Execute "update tblextensionmortality a ,mhv.tbltshewog b set " _
 & " grptshowog= concat(grptshowog,'  ',tshewogname) where grptshowog=concat(dzongkhagid,gewogid,tshewogid)"
 
 ODKDB.Execute "update tblextensionmortality a ,mhv.tblfarmer b set farmercode=concat(farmercode,'  ',farmername)" _
  & "  where farmercode=idfarmer"
         
         
         
Set rs = Nothing
rs.Open "select farmerbarcode,fdcode,avg(GPS_COORDINATES_ALT) as alt  from phealthhub15_core where GPS_COORDINATES_ALT>0 group by farmerbarcode,fdcode order by farmerbarcode,fdcode", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "update tblextensionmortality set alt='" & rs!alt & "' where substring(farmercode,1,14)='" & rs!farmerbarcode & "' and fieldcode='" & rs!FDCODE & "'"
rs.MoveNext
Loop
End Sub

Private Sub Command9_Click()
'createweek
createnewtask
End Sub
Private Sub createnewtask()
On Error GoTo merr
Dim mweek As Integer
Dim i As Integer
Dim oldweek As Integer
Dim newweek As Integer
Dim chkyear As Integer
Dim myear As Integer
Dim oldtrnno As Integer
Dim rs As New ADODB.Recordset
mweek = Format(Now(), "ww", vbSunday, vbFirstFullWeek) + 1 'DatePart("ww", Now())
Set rs = Nothing
rs.Open "select * from tblweek where status='1'", MHWEBDB
If rs.EOF <> True Then
chkyear = rs!myear
If rs!weekno = 52 Then
oldweek = rs!weekno
newweek = 1
myear = rs!myear + 1
Else
oldweek = rs!weekno
newweek = rs!weekno + 1
myear = rs!myear
End If
End If



If oldweek <> mweek And oldweek + 1 = mweek And Weekday(Now, vbMonday) = 7 And chkyear = Year(Now) Then
Else
MsgBox oldweek & " " & mweek & "  " & Weekday(Now, vbMonday) & "  " & chkyear
Exit Sub
End If






MHVDB.BeginTrans

MHWEBDB.Execute "delete from tbldailymeettasktrn where length(taskdescription)=0"
MHWEBDB.Execute "delete from tbldailymorningmeetingissues where length(issuedescription)=0"
MHWEBDB.Execute "delete from tbldailymorningmeetingactionitem where length(actionitem)=0"
MHWEBDB.Execute "delete from tblncdailymeettasktrn where length(taskdescription)=0"
MHWEBDB.Execute "update tbldailymeettasktrn set completiondate='',reviseddate='' where isdailytask='Yes'"

MHWEBDB.Execute "delete from tbldailymeettasktrn where   weekno>'" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "delete from tblncdailymeettasktrn where   weekno>'" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "delete from tbldailymorningmeetingissues where   weekno>'" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "delete from tbldailymorningmeetingactionitem where   weekno>'" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "delete from tbldailymorningmeetingadminactionitem where   weekno>'" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "delete from tblncattendance where   weekno>'" & newweek & "' and myear='" & myear & "'"

' shifting daily task first lot
Set rs = Nothing
rs.Open "select * from tbldailymeettasktrn where isdailytask='Yes' and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymeettasktrn(weekno,taskdescription,personresponsible," _
& "completiondate,remarks,status,myear,isbacklogged,backloggedremarks,deptcode,isdailytask," _
& " datemon,datetue,datewed,datethu,datefri,datesat)" _
& " values(" _
& "'" & newweek & "','" & ValidateString(rs!taskdescription) & "','" & ValidateString(rs!personresponsible) & "'," _
& "'" & Format(rs!completiondate, "yyyy-MM-dd") & "','" & ValidateString(rs!remarks) & "','" & rs!status & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "','Yes','" & rs!datemon & "','" & rs!datetue & "','" & rs!datewed & "','" & rs!datethu & "','" & rs!datefri & "','" & rs!datesat & "')"
rs.MoveNext
Loop



' shifting daily task second lot,nursery construction
Set rs = Nothing

rs.Open "select * from tblncdailymeettasktrn where isdailytask='Yes' and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True

MHWEBDB.Execute "insert into tblncdailymeettasktrn " _
& " (weekno,taskdescription,weeklytarget,timeinmotion, " _
& " remarks,status, " _
& " mmon,mtue,mwed,mthu,mfri,msat,msun, " _
& " deptcode,myear,isbacklogged,backloggedremarks," _
& " isdailytask," _
& " pmon,ptue,pwed,pthu,pfri,psat,psun," _
& " amon,atue,awed,athu,afri,asat,asun," _
& " cmon,ctue,cwed,cthu,cfri,csat,csun) " _
& " values( " _
& "'" & newweek & "','" & ValidateString(rs!taskdescription) & "', " _
& " '" & ValidateString(rs!weeklytarget) & "','" & ValidateString(rs!timeinmotion) & "', " _
& " '" & ValidateString(rs!remarks) & "','" & rs!status & "'," _
& " '" & ValidateString(rs!mmon) & "','" & ValidateString(rs!mtue) & "','" & ValidateString(rs!mwed) & "','" & ValidateString(rs!mthu) & "','" & ValidateString(rs!mfri) & "','" & ValidateString(rs!msat) & "','" & ValidateString(rs!msun) & "', " _
& " '" & rs!deptcode & "','" & myear & "','1','" & oldweek & "','Yes', " _
& " '" & ValidateString(rs!pmon) & "','" & ValidateString(rs!ptue) & "','" & ValidateString(rs!pwed) & "','" & ValidateString(rs!pthu) & "','" & ValidateString(rs!pfri) & "','" & ValidateString(rs!psat) & "','" & ValidateString(rs!psun) & "', " _
& " '" & ValidateString(rs!amon) & "','" & ValidateString(rs!atue) & "','" & ValidateString(rs!awed) & "','" & ValidateString(rs!athu) & "','" & ValidateString(rs!afri) & "','" & ValidateString(rs!asat) & "','" & ValidateString(rs!asun) & "', " _
& " '" & ValidateString(rs!cmon) & "','" & ValidateString(rs!ctue) & "','" & ValidateString(rs!cwed) & "','" & ValidateString(rs!cthu) & "','" & ValidateString(rs!cfri) & "','" & ValidateString(rs!csat) & "','" & ValidateString(rs!csun) & "')"

rs.MoveNext
Loop





' shifting Y and R to next week
Set rs = Nothing

rs.Open "select * from tbldailymeettasktrn where status in('Y','R') and isdailytask not in('Yes') and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymeettasktrn(weekno,taskdescription,personresponsible," _
& "completiondate,reviseddate,remarks,status,myear,isbacklogged,backloggedremarks,deptcode)values(" _
& "'" & newweek & "','" & ValidateString(rs!taskdescription) & "','" & ValidateString(rs!personresponsible) & "'," _
& "'" & Format(rs!completiondate, "yyyy-MM-dd") & "','" & Format(rs!reviseddate, "yyyy-MM-dd") & "','" & ValidateString(rs!remarks) & "','" & rs!status & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "')"
MHWEBDB.Execute "update tbldailymeettasktrn set istaskshifted='1' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop


' shifting Y and R to next week for nursery and cons
Set rs = Nothing
rs.Open "select * from tblncdailymeettasktrn where status in('Y','R') and  isdailytask not in('Yes') and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tblncdailymeettasktrn " _
& " (weekno,taskdescription,weeklytarget,timeinmotion, " _
& " remarks,status, " _
& " mmon,mtue,mwed,mthu,mfri,msat,msun, " _
& " deptcode,myear,isbacklogged,backloggedremarks," _
& " isdailytask," _
& " pmon,ptue,pwed,pthu,pfri,psat,psun," _
& " amon,atue,awed,athu,afri,asat,asun," _
& " cmon,ctue,cwed,cthu,cfri,csat,csun)" _
& " values(" _
& "'" & newweek & "','" & ValidateString(rs!taskdescription) & "', " _
& " '" & ValidateString(rs!weeklytarget) & "','" & ValidateString(rs!timeinmotion) & "', " _
& " '" & ValidateString(rs!remarks) & "','" & rs!status & "'," _
& " '" & ValidateString(rs!mmon) & "','" & ValidateString(rs!mtue) & "','" & ValidateString(rs!mwed) & "','" & ValidateString(rs!mthu) & "','" & ValidateString(rs!mfri) & "','" & ValidateString(rs!msat) & "','" & ValidateString(rs!msun) & "', " _
& " '" & rs!deptcode & "','" & myear & "','1','" & oldweek & "','Yes'," _
& " '" & ValidateString(rs!pmon) & "','" & ValidateString(rs!ptue) & "','" & ValidateString(rs!pwed) & "','" & ValidateString(rs!pthu) & "','" & ValidateString(rs!pfri) & "','" & ValidateString(rs!psat) & "','" & ValidateString(rs!psun) & "', " _
& " '" & ValidateString(rs!amon) & "','" & ValidateString(rs!atue) & "','" & ValidateString(rs!awed) & "','" & ValidateString(rs!athu) & "','" & ValidateString(rs!afri) & "','" & ValidateString(rs!asat) & "','" & ValidateString(rs!asun) & "', " _
& " '" & ValidateString(rs!cmon) & "','" & ValidateString(rs!ctue) & "','" & ValidateString(rs!cwed) & "','" & ValidateString(rs!cthu) & "','" & ValidateString(rs!cfri) & "','" & ValidateString(rs!csat) & "','" & ValidateString(rs!csun) & "')"


MHWEBDB.Execute "update tbldailymeettasktrn set personresponsible1=personresponsible"
rs.MoveNext
Loop










'shifting completion date ahead of sunday

Set rs = Nothing
rs.Open "select * from tbldailymeettasktrn where status in('G') and isdailytask not in('Yes') and  weekno='" & oldweek & "'  and completiondate>'" & Format(Now, "yyyy-MM-dd") & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymeettasktrn(weekno,taskdescription,personresponsible," _
& "completiondate,remarks,status,myear,isbacklogged,backloggedremarks,deptcode)values(" _
& "'" & newweek & "','" & ValidateString(rs!taskdescription) & "','" & ValidateString(rs!personresponsible) & "'," _
& "'" & Format(rs!completiondate, "yyyy-MM-dd") & "','" & ValidateString(rs!remarks) & "','" & rs!status & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "')"
MHWEBDB.Execute "update tbldailymeettasktrn set istaskshifted='1' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop

'shifting revisied date ahead of sunday

Set rs = Nothing
rs.Open "select * from tbldailymeettasktrn where status in('G') and isdailytask not in('Yes') and  weekno='" & oldweek & "' and reviseddate>'" & Format(Now, "yyyy-MM-dd") & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymeettasktrn(weekno,taskdescription,personresponsible," _
& "completiondate,reviseddate,remarks,status,myear,isbacklogged,backloggedremarks,deptcode)values(" _
& "'" & newweek & "','" & ValidateString(rs!taskdescription) & "','" & ValidateString(rs!personresponsible) & "'," _
& "'" & Format(rs!completiondate, "yyyy-MM-dd") & "','" & Format(rs!reviseddate, "yyyy-MM-dd") & "','" & ValidateString(rs!remarks) & "','" & rs!status & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "')"
MHWEBDB.Execute "update tbldailymeettasktrn set istaskshifted='1' where trnid='" & rs!trnid & "'"
rs.MoveNext
Loop




' shifting unsolved issue
Set rs = Nothing
rs.Open "select * from tbldailymorningmeetingissues where resolved not in('Yes') and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymorningmeetingissues(weekno,issuedescription,resolved," _
& "myear,isbacklogged,backloggedremarks,deptcode,deadline,responsibleperson)values(" _
& "'" & newweek & "','" & ValidateString(rs!issuedescription) & "','" & rs!resolved & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "','" & Format(rs!deadline, "yyyy-MM-dd") & "','" & rs!responsibleperson & "')"

MHWEBDB.Execute "update tbldailymorningmeetingissues set istaskshifted='1' where trnid='" & rs!trnid & "'"

rs.MoveNext


Loop



'shifting unsolved action
Set rs = Nothing
rs.Open "select * from tbldailymorningmeetingactionitem where resolved not in('Yes') and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymorningmeetingactionitem(weekno,actionitem,deadline,responsibleperson," _
& "resolved," _
& "myear,isbacklogged,backloggedremarks,deptcode)values(" _
& "'" & newweek & "','" & ValidateString(rs!actionitem) & "','" & rs!deadline & "','" & ValidateString(rs!responsibleperson) & "', " _
& "  '" & rs!resolved & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "')"
MHWEBDB.Execute "update tbldailymorningmeetingactionitem set istaskshifted='1' where trnid='" & rs!trnid & "'"

rs.MoveNext


Loop

' admin action item

Set rs = Nothing
rs.Open "select * from tbldailymorningmeetingadminactionitem where resolved not in('Yes') and  weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
MHWEBDB.Execute "insert into tbldailymorningmeetingadminactionitem(weekno,actionitem,deadline,responsibleperson," _
& "resolved," _
& "myear,isbacklogged,backloggedremarks,deptcode)values(" _
& "'" & newweek & "','" & ValidateString(rs!actionitem) & "','" & rs!deadline & "','" & ValidateString(rs!responsibleperson) & "', " _
& "  '" & rs!resolved & "'," _
& "'" & myear & "','1','" & oldweek & "','" & rs!deptcode & "')"
MHWEBDB.Execute "update tbldailymorningmeetingadminactionitem set istaskshifted='1' where trnid='" & rs!trnid & "'"

rs.MoveNext


Loop




' shift attendance

Set rs = Nothing
rs.Open "select * from tblncattendance where weekno='" & oldweek & "' and myear='" & myear & "'", MHWEBDB
Do While rs.EOF <> True
If rs!trnid = 1 Or rs!trnid = 16 Or rs!trnid = 20 Then
MHWEBDB.Execute "insert into tblncattendance(trnid,weekno,mparticular,datemon,datetue, " _
& " datewed,datethu,datefri,datesat,datesun, " _
& " deptcode,myear,mcolor,remarks,noedit,morder) " _
& " values('" & rs!trnid & "','" & newweek & "','" & rs!mparticular & "', " _
& " '" & rs!datemon & "','" & rs!datetue & "','" & rs!datewed & "', " _
& " '" & rs!datethu & "','" & rs!datefri & "','" & rs!datesat & "', " _
& " '" & rs!datesun & "','" & rs!deptcode & "','" & myear & "', " _
& " '" & rs!mcolor & "','" & rs!remarks & "','" & rs!noedit & "', " _
& " '" & rs!morder & "')"
Else
MHWEBDB.Execute "insert into tblncattendance(trnid,weekno,mparticular," _
& " deptcode,myear,mcolor,noedit,morder) " _
& " values('" & rs!trnid & "','" & newweek & "','" & rs!mparticular & "', " _
& " '" & rs!deptcode & "','" & myear & "', " _
& " '" & rs!mcolor & "','" & rs!noedit & "', " _
& " '" & rs!morder & "')"
End If

rs.MoveNext


Loop







MHWEBDB.Execute "update tblweek set status=0"
MHWEBDB.Execute "update tblweek set status=1 where weekno='" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "update tblweek set morder=0"
MHWEBDB.Execute "update tblweek set morder=99 where  weekno='" & newweek & "' and myear='" & myear & "'"
MHWEBDB.Execute "update tblweek set morder=98 where  weekno='" & newweek - 1 & "' and myear='" & myear & "'"
MHWEBDB.Execute "update tblweek set morder=97 where  weekno='" & newweek + 1 & "' and myear='" & myear & "'"
MHVDB.CommitTrans


 retVal = SendMail("muktitcc@gmail.com,swangchuk@mountainhazelnuts.com", "dailyTaskCreatinInfo", "mukti@mountainhazelnuts.com", _
          "New task for the week " & mweek & " is successfully created! Cheers Bro!", "smtp.tashicell.com", 25, _
          "habizabi", "habizabi", _
           "", CBool(False))
 



If retVal = "ok" Then
'MsgBox "ok"
Else
'MsgBox "NOT ok"
End If



Exit Sub
merr:
    MHVDB.RollbackTrans
    MsgBox err.Description




End Sub



Private Sub createweek()
Dim i As Integer
For i = 1 To 52
MHWEBDB.Execute "insert into tblweek(weekno,status) values('" & i & "',0)"
Next
End Sub

Private Sub phi()
Dim atleastone As Boolean
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim dm, ap, rp, sp, ad, cdm, cap, crp, csp, cad As Double
Dim dm1, ap1, rp1, sp1, ad1 As Double
Dim dm2, ap2, rp2, sp2, ad2 As Double
Dim SQLSTR As String
Dim mfarmerbarcode As String
Dim mfdcode As Integer
Set rs = Nothing
rs.Open "select * from tblphithreshold", ODKDB
If rs.EOF <> True Then
dm = rs!deadmissing
ap = rs!activepest
rp = rs!rootpest
sp = rs!stempest
ad = rs!animaldamage
cdm = rs!deadmissingchange
cap = rs!activepestchange
crp = rs!rootpestchange
csp = rs!stempestchange
cad = rs!animaldamagechange
End If



      SQLSTR = ""
      SQLSTR = "insert into tblfieldtofollowup (_URI,START,end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,farmerbarcode,fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS,TREEHEIGHT) " _
      & "select _URI,START,n.end,REGION_GCODE,REGION_DCODE,REGION, " _
      & "staffbarcode,n.farmerbarcode,n.fdcode," _
      & " TREE_COUNT_DEADMISSING,TREE_COUNT_ACTIVEGROWING,TREE_COUNT_SLOWGROWING,STEMPEST,ROOTPEST, " _
      & " ACTIVEPEST,GPS_COORDINATES_LAT,GPS_COORDINATES_LNG,GPS_COORDINATES_ALT, " _
      & "GPS_COORDINATES_ACC,TREE_COUNT_TOTALTREES,WATERLOG,TREE_COUNT_DOR,ANIMALDAMAGE,MONITORCOMMENTS,TREEHEIGHT from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode, n.fdcode"
         
         
 ODKDB.Execute "delete from tblfieldtofollowup"
 ODKDB.Execute SQLSTR
 ODKDB.Execute "update tblfieldtofollowup set region_dcode=substring(farmerbarcode,1,3),region_gcode=substring(farmerbarcode,1,6), " _
 & " region=substring(farmerbarcode,1,9)"
 
 
 
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set  recordage=round(datediff(CURDATE(),END),0)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set startenddiff=round(datediff(END,START),0)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set activepercent=(TREE_COUNT_ACTIVEGROWING*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set deadpercent=(TREE_COUNT_DEADMISSING*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set stempestpercent=(STEMPEST*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set rootpestpercent=(ROOTPEST*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set activepestpercent=(ACTIVEPEST*100)/(TREE_COUNT_TOTALTREES)")
'ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set adamagepercent=(ANIMALDAMAGE*100)/(TREE_COUNT_TOTALTREES)")

    
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phideadmissing=2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phistempest=2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phirootpest=2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phiactivepest=2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST)")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set phianimaldamage=2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE)")


ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set dm='Y' where phideadmissing>='" & dm & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set sp='Y' where phistempest>='" & sp & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set rp='Y' where  phirootpest>='" & rp & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set ap='Y' where  phiactivepest>='" & ap & "'")
ODKDB.Execute ("update odk_prodlocal.tblfieldtofollowup set ad='Y'where phianimaldamage>='" & ad & "'")

'check records never show
Set rs = Nothing
rs.Open "select * from odk_prodlocal.tblfieldtofollowup where concat(farmerbarcode,fdcode) in(select concat(farmerbarcode,fdcode) from tblfieldfollowupnevershow)", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "delete from odk_prodlocal.tblfieldtofollowup where FARMERBARCODE='" & rs!farmerbarcode & "' and FDCODE='" & rs!FDCODE & "'"
rs.MoveNext
Loop

'check records not to show for this visit

Set rs = Nothing
rs.Open "select * from odk_prodlocal.tblfieldtofollowup where _URI in(select _URI from tblfieldfollowupdonotshowincurrentvisit)", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "delete from odk_prodlocal.tblfieldtofollowup where _URI='" & rs![_uri] & "'"
rs.MoveNext
Loop




'check records already flagged
dm1 = 0
ap1 = 0
rp1 = 0
sp1 = 0
ad1 = 0
mfdcode = 0
mfarmerbarcode = ""
Set rs = Nothing
rs.Open "select _URI,farmerbarcode,fdcode, " _
& " 2*((TREE_COUNT_DEADMISSING)/(TREE_COUNT_TOTALTREES)) + log(TREE_COUNT_DEADMISSING) as dm1, " _
& " 2*((STEMPEST)/(TREE_COUNT_TOTALTREES)) + log(STEMPEST) as sp1, " _
& " 2*((ROOTPEST)/(TREE_COUNT_TOTALTREES)) + log(ROOTPEST) as rp1, " _
& " 2*((ACTIVEPEST)/(TREE_COUNT_TOTALTREES)) + log(ACTIVEPEST) as ap1, " _
& " 2*((ANIMALDAMAGE)/(TREE_COUNT_TOTALTREES)) + log(ANIMALDAMAGE) as ad1 " _
& " from odk_prodlocal.tblfieldtofollowup where concat(farmerbarcode,fdcode) in(select concat(farmerbarcode,fdcode) from tblfieldforfollowup where status='ON')", ODKDB
Do While rs.EOF <> True
atleastone = False
mfarmerbarcode = rs!farmerbarcode
mfdcode = rs!FDCODE
dm1 = IIf(IsNull(rs!dm1), 0, rs!dm1)
ap1 = IIf(IsNull(rs!ap1), 0, rs!ap1)
rp1 = IIf(IsNull(rs!rp1), 0, rs!rp1)
sp1 = IIf(IsNull(rs!sp1), 0, rs!sp1)
ad1 = IIf(IsNull(rs!ad1), 0, rs!ad1)
dm2 = 0
ap2 = 0
rp2 = 0
sp2 = 0
ad2 = 0
Set rs1 = Nothing
rs1.Open "select * from tblfieldforfollowup where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "' and status='ON'", ODKDB
If rs1.EOF <> True Then
dm2 = dm1 - rs1!dm
ap2 = ap1 - rs1!ap
rp2 = rp1 - rs1!rp
sp2 = sp1 - rs1!sp
ad2 = ad1 - rs1!ad

If dm2 >= cdm And rs1!dm >= dm Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phideadmissing='" & rs1!dm & "', dmc='" & dm1 & "',dm='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If dm2 <= (-1 * cdm) And rs1!dm >= dm Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phideadmissing='" & rs1!dm & "', dmc='" & dm1 & "',dm='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If ap2 >= cap And rs1!ap >= ap Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phiactivepest='" & rs1!ap & "',apc='" & dm1 & "',ap='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If ap2 <= (-1 * cap) And rs1!ap >= ap Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phiactivepest='" & rs1!ap & "',apc='" & dm1 & "',ap='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If rp2 >= crp And rs1!rp >= rp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phirootpest='" & rs1!rp & "',rpc='" & dm1 & "',rp='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If rp2 <= (-1 * crp) And rs1!rp >= rp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phirootpest='" & rs1!rp & "',rpc='" & dm1 & "',rp='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If sp2 >= csp And rs1!sp >= sp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phistempest='" & rs1!sp & "',spc='" & dm1 & "',sp='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If sp2 <= (-1 * csp) And rs1!sp >= sp Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phistempest='" & rs1!sp & "',spc='" & dm1 & "',sp='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If


If ad2 >= cad And rs1!ad >= ad Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phianimaldamage='" & rs1!ad & "', adc='" & ad1 & "',ad='R' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If ad2 <= (-1 * cad) And rs1!ad >= ad Then
ODKDB.Execute "update tblfieldtofollowup set nodelete='1',phianimaldamage='" & rs1!ad & "', adc='" & ad1 & "',ad='G' where farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
atleastone = True
End If

If atleastone = False Then
ODKDB.Execute "delete from  tblfieldtofollowup where _URI='" & rs![_uri] & "' and farmerbarcode='" & mfarmerbarcode & "' and fdcode='" & mfdcode & "'"
End If

End If

rs.MoveNext
Loop






ODKDB.Execute ("delete from  odk_prodlocal.tblfieldtofollowup where phideadmissing<'" & dm & "' and phistempest<'" & sp & "' and  phirootpest<'" & rp & "' and phiactivepest<'" & ap & "' and phianimaldamage<'" & ad & "' and nodelete<>'1'")
 
End Sub

Private Sub fileupdate_Click()
updatephpfiles
End Sub
Private Sub updatephpfiles()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim pass1 As Boolean
Dim pass2 As Boolean
Set rs = Nothing
rs.Open "select * from tblfilemaster where entrytype in('New','Update')", MHWEBDB
If rs.EOF <> True Then
    Do While rs.EOF <> True
    pass1 = False
    pass2 = False
    Set rs1 = Nothing
    rs1.Open "select * from tblaccesslog where pageUrl='" & rs!FileName & "' and memberId='3' and " _
    & " memberId in(select uid from tblfileaccessrights where fid in (select fileid from tblfilemaster where filename='" & rs!FileName & "')and uid not in(1,3,5)) or  memberId in(select id from tbluser where dept='113')", MHWEBDB
    If rs1.EOF <> True Then
    pass1 = True
    End If
    
    If pass1 = True Then
    MHWEBDB.Execute "update tblfilemaster set entrytype='Normal' where filename='" & rs!FileName & "'"
    End If
     pass1 = False
    
    rs.MoveNext
    Loop

End If
End Sub



Private Sub createnewdashboardworkspace()
'On Error GoTo merr
Dim mweek As Integer
Dim i As Integer
Dim myear As Integer
Dim rsd As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim validmonth As Integer
Dim oldmonth As Integer
Dim oldyear As Integer
Dim validyear As Integer
Dim tchk As Integer
Set rs = Nothing

rs.Open "select * from tblmonth where status=1", MHWEBDB
If rs.EOF <> True Then
oldmonth = rs!monthno

If rs!monthno = 12 Then
validmonth = 1
validyear = rs!myear + 1
Else
validmonth = oldmonth + 1
validyear = rs!myear
End If

oldyear = rs!myear
End If
tchk = Month(Now) - oldmonth
If tchk = 2 Then
Else
'MsgBox "You cannot run this process"
'Exit Sub
End If



MHVDB.BeginTrans

MHWEBDB.Execute "delete from tbldashboardmonitoringissues where length(issues)=0"
MHWEBDB.Execute "delete from tbldashboardmonitoringactions where length(actionitem)=0"

MHWEBDB.Execute "delete from tbldashboardmonitoringsummary where   monthno='" & validmonth & "' and myear='" & validyear & "'"
MHWEBDB.Execute "delete from tbldashboardmonitoringissues where   monthno='" & validmonth & "' and myear='" & validyear & "'"
MHWEBDB.Execute "delete from tbldashboardmonitoringactions where   monthno='" & validmonth & "' and myear='" & validyear & "'"
MHWEBDB.Execute "delete from tbldashboardmonitoringhrneed where   monthno='" & validmonth & "' and myear='" & validyear & "'"
MHWEBDB.Execute "delete from tbldashboardworkplan where   monthno='" & validmonth & "' and myear='" & validyear & "'"
MHWEBDB.Execute "delete from tblconsthroperation where   monthno='" & validmonth & "' and myear='" & validyear & "'"


Set rsd = Nothing
rsd.Open "select deptid from tbldepartment where dashboard='1'", MHVDB
Do While rsd.EOF <> True
' processes tbldashboardmonitoringsummary
Set rs = Nothing
rs.Open "SELECT * FROM tbldashboardmonitoringsummary where monthno='" & oldmonth & "' and myear='" & oldyear & "' and deptcode='" & rsd!deptid & "'", MHWEBDB
    Do While rs.EOF <> True
    MHWEBDB.Execute "insert into tbldashboardmonitoringsummary(headerid,particulars,status,deptcode,monthno,myear)" _
    & " values('" & rs!headerid & "','" & rs!PARTICULARS & "','G','" & rs!deptcode & "','" & validmonth & "','" & validyear & "') "
    rs.MoveNext
    Loop


' processes tbldashboardmonitoringissues
For i = 1 To 10
    MHWEBDB.Execute "insert into tbldashboardmonitoringissues(headerid,deptcode,monthno,myear)" _
    & " values('1','" & rsd!deptid & "','" & validmonth & "','" & validyear & "') "
Next


' processes tbldashboardmonitoringhrneed
Set rs = Nothing
rs.Open "SELECT * FROM tbldashboardmonitoringhrneed where monthno='" & oldmonth & "' and myear='" & oldyear & "' and deptcode='" & rsd!deptid & "'", MHWEBDB
    Do While rs.EOF <> True
    MHWEBDB.Execute "insert into tbldashboardmonitoringhrneed(headerid,particulars,deptcode,monthno,myear)" _
    & " values('" & rs!headerid & "','" & rs!PARTICULARS & "','" & rs!deptcode & "','" & validmonth & "','" & validyear & "') "
    rs.MoveNext
    Loop


' processes tbldashboardmonitoringissues
For i = 1 To 10
    MHWEBDB.Execute "insert into tbldashboardmonitoringactions(headerid,deptcode,monthno,myear)" _
    & " values('1','" & rsd!deptid & "','" & validmonth & "','" & validyear & "') "
Next



' processes work plan
Set rs = Nothing
rs.Open "SELECT * FROM tbldashboardworkplan where monthno='" & oldmonth & "' and myear='" & oldyear & "' and " _
& " deptcode='" & rsd!deptid & "' and status in('Y','R') union SELECT * FROM tbldashboardworkplan where " _
& " monthno='" & oldmonth & "' and myear='" & oldyear & "' and deptcode='" & rsd!deptid & "' and  status in('G') and " _
& " completiondate>'" & Format(Now, "yyyy-MM-dd") & "' ", MHWEBDB
    Do While rs.EOF <> True
    MHWEBDB.Execute "insert into tbldashboardworkplan(monthno,myear,deptcode,taskdescription,personresponsible,completiondate,exportedtoweeklytask,status,maintaskid)" _
    & " values('" & validmonth & "','" & validyear & "','" & rs!deptcode & "','" & rs!taskdescription & "','" & rs!personresponsible & "','" & Format(rs!completiondate, "yyyy-MM-dd") & "','" & rs!exportedtoweeklytask & "','" & rs!status & "','" & rs!maintaskid & "') "
    rs.MoveNext
    Loop
    
    
    
If rsd!deptid = 110 Then
   ' tblconsthroperation
MHWEBDB.Execute "insert into tblconsthroperation (`monthno`, `taskdescription`, `personresponsible`, " _
& " `completiondate`, `reviseddate`, `remarks`, `status`, `week1`, `week2`, `week3`, `week4`, `week5`, " _
& " `week6`, `week7`, `week8`, `deptcode`, `myear`, `isbacklogged`, `backloggedremarks`, " _
& " `istaskshifted`, `del`, `isdailytask`, `istaskcompleted`, `maintaskid`, `subtaskid`, `nextmonthreq`)" _
& " select  '" & validmonth & "' `monthno`, `taskdescription`, `personresponsible`, `completiondate`, " _
& " `reviseddate`, `remarks`, `status`, `week1`, `week2`, `week3`, `week4`, " _
& " `week5`, `week6`, `week7`, `week8`, '" & rsd!deptid & "' `deptcode`, '" & validyear & "' `myear`, `isbacklogged`, " _
& " `backloggedremarks`, `istaskshifted`, `del`, `isdailytask`, `istaskcompleted`, " _
& " `maintaskid`, `subtaskid`, `nextmonthreq` from tblconsthroperation " _
& " where monthno='" & oldmonth & "' and myear='" & oldyear & "' and deptcode='" & rsd!deptid & "'"
    
End If
rsd.MoveNext

Loop





MHWEBDB.Execute "update tblmonth set status=0"
MHWEBDB.Execute "update tblmonth set status=1 where monthno='" & validmonth & "' and myear='" & validyear & "'"
MHVDB.CommitTrans


 retVal = SendMail("muktitcc@gmail.com,swangchuk@mountainhazelnuts.com", "dashboardCreatinInfo", "mukti@mountainhazelnuts.com", _
          "New task for the month " & MonthName(validmonth, True) & " is successfully created! Cheers Bro!", "smtp.tashicell.com", 25, _
          "habizabi", "habizabi", _
           "", CBool(False))
 



If retVal = "ok" Then
'MsgBox "ok"
Else
'MsgBox "NOT ok"
End If



Exit Sub
merr:
    MHVDB.RollbackTrans
    MsgBox err.Description




End Sub


