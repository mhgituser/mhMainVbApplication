VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmdashboardnursery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NURSERY DASH BOARD"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   Icon            =   "frmdashboardnursery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   1560
      TabIndex        =   24
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VSFlex7Ctl.VSFlexGrid mygrid 
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   2775
      _cx             =   4895
      _cy             =   1508
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
      Rows            =   200
      Cols            =   100
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmdashboardnursery.frx":0E42
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
   Begin VB.Frame Frame2 
      Caption         =   "Email List"
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
      Left            =   10320
      TabIndex        =   13
      Top             =   0
      Width           =   4215
      Begin VB.ListBox DZLIST 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   0
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ok"
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
         Left            =   3000
         Picture         =   "frmdashboardnursery.frx":0F50
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   5640
      TabIndex        =   9
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mail me"
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
         Left            =   3360
         Picture         =   "frmdashboardnursery.frx":1602
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Message"
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
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
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
         TabIndex        =   11
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         TabIndex        =   10
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command7 
         Caption         =   "dashbord"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "php script"
         Height          =   495
         Left            =   3960
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin MSComCtl2.DTPicker txtfromdate 
         Height          =   375
         Left            =   1080
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78774273
         CurrentDate     =   41520
      End
      Begin VB.CheckBox chkmailme 
         Caption         =   "Mail"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   18
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox cbomnth 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmdashboardnursery.frx":1D6C
         Left            =   4080
         List            =   "frmdashboardnursery.frx":1D94
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load Graph"
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
         Left            =   1080
         Picture         =   "frmdashboardnursery.frx":1DD4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
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
         Left            =   2400
         Picture         =   "frmdashboardnursery.frx":253E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cboyear 
         Bindings        =   "frmdashboardnursery.frx":3208
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSDataListLib.DataCombo cbodept 
         Bindings        =   "frmdashboardnursery.frx":321D
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   840
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
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3240
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   78774273
         CurrentDate     =   41520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Department"
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
         Width           =   990
      End
      Begin VB.Label Label2 
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
         TabIndex        =   2
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   405
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1680
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmdashboardnursery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shipmentsize, healthyplant, oversize, undersize, weakreceived, icedamaged, totreceived, deadin10 As Long
Dim healthyin10, healthyplantsaftereye, hardenplants, servival1 As Long
Dim totbillladding, totreceivedin, totoversize, totundersize, totweak, toticedamaged, totreceivedex, totdead, tothealthy10, tothealthyeye, tothardenplants As Double

Private Sub Chart1_GotFocus()

End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkmailme_Click()
If chkmailme.Value = 1 Then
frmdashboardnursery.Width = 10395
Else
frmdashboardnursery.Width = 5685
End If
End Sub

Private Sub Command1_Click()
'On Error GoTo err


Select Case cbodept.BoundText
          Case 1
                nursery
          Case 2
                finance
          Case 5
                monitoring
          Case 4
                outreach
          Case 6
                field
            Case 7
                storage
End Select
'Exit Sub
'err:
'   MsgBox err.Description & " Please close previously opened file."
'
'   Shell "taskkill.exe /f /t /im Excel.exe"
End Sub
Private Sub storage()
'    Dim xl As Excel.Application
'    Dim rs As New ADODB.Recordset
'    Dim SQLSTR As String
'    Dim var As Variant
'    Dim i, j As Integer
'    Set xl = CreateObject("excel.Application")
'    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm") <> vbNullString Then
'    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
'    End If
'
'     Set rs = Nothing
'    rs.Open "select * from tbldashbordtrn where trnid='4'", MHVDB
'    If rs.EOF <> True Then
'    getSheet 5, rs!FileName
'    End If
'    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
'    If Dir$(dashBoardName) <> vbNullString Then
'        Kill dashBoardName
'    End If
'
'
'    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
'    xl.Sheets("Data").Select
'    xl.Visible = False
'    GetTbl
'
'
'SQLSTR = ""
'
'
'         SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,totaltrees," _
'         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
'         & "waterlog,activepest,animaldamage,monitorcomments)  select start,tdate,end,staffbarcode,region_dcode," _
'         & "region_gcode,region,n.farmerbarcode,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
'         & "dtrees,ndtrees,wlogged,pdamage,adamage,monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
'         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
'         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
'         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
'
'
'  ODKDB.Execute SQLSTR
'
' SQLSTR = "select * from " & Mtblname & " where totaltrees>=100 "
'  i = 2
'  mchk = True
'  Set rs = Nothing
'    rs.Open SQLSTR, ODKDB
'  Do While rs.EOF <> True
'
'    FindDZ Mid(rs!farmercode, 1, 3)
'    xl.Cells(i, 1) = Mid(rs!farmercode, 1, 3) & " " & Dzname
'
'    FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
'    xl.Cells(i, 2) = Mid(rs!farmercode, 4, 3) & " " & GEname
'
'    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
'    xl.Cells(i, 3) = Mid(rs!farmercode, 7, 3) & " " & TsName
'
'    FindFA IIf(IsNull(rs!farmercode), "", rs!farmercode), "F"
'    xl.Cells(i, 4) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & " " & FAName
'
'    xl.Cells(i, 5) = "'" & rs!End
'    xl.Cells(i, 6) = DateDiff("d", rs!End, Now - 2)
'
'    xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
'    xl.Cells(i, 8) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
'    xl.Cells(i, 9) = IIf(IsNull(rs!tree_count_activegrowing), "", rs!tree_count_activegrowing)
'    xl.Cells(i, 10) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
'    xl.Cells(i, 11) = IIf(IsNull(rs!activepest), "", rs!activepest)
'    xl.Cells(i, 12) = IIf(IsNull(rs!stempest), "", rs!stempest)
'    xl.Cells(i, 13) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
'    xl.Cells(i, 14) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
'
'i = i + 1
'rs.MoveNext
'   Loop
' xl.Sheets("10 Worst Farmers").Select
'storageworst10 xl
'
'
'xl.Sheets("10 Worst Tshowog").Select
'storageworst10tshowog xl
'
'
'
'ODKDB.Execute "drop table " & Mtblname & ""
'xl.Visible = True
'Set xl = Nothing
'Screen.MousePointer = vbDefault

 Dim xl As Excel.Application
 Dim rr As String
    Dim rs As New ADODB.Recordset
    Dim SQLSTR As String
    Dim var As Variant
    Dim i, j As Integer
    Set xl = CreateObject("excel.Application")
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    End If
    
     Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='4'", MHVDB
    If rs.EOF <> True Then
    getSheet 5, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Sheets("Mortality Data").Select
    xl.Visible = False
    updatemortalityrecord
    fillmortalitybyalt xl
    fillmortalitybyacre xl
    rr = ""
    Set rs = Nothing
    rs.Open "select distinct region from tblextensionmortality order by region", ODKDB
    Do While rs.EOF <> True
    rr = IIf(Len(rs!region) = 0, "Not Assigned", rs!region)
    xl.Sheets(rr).Select
    fillmortalitybyreg xl, rr
    rs.MoveNext
    Loop
      
'
' xl.Sheets("10 Worst Farmers").Select
'storageworst10 xl
'
'
'xl.Sheets("10 Worst Tshowog").Select
'storageworst10tshowog xl
'
'
'
'ODKDB.Execute "drop table " & Mtblname & ""
xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault



End Sub
Private Sub fillmortalitybyreg(xl As Object, region As String)
Dim rs As New ADODB.Recordset
Dim ge As String
Dim ttrees As Long
Dim i As Integer
Set rs = Nothing
i = 3
rs.Open "SELECT grptshowog ,count(farmercode) cnt,sum(deadmissing) dm,sum(totaltrees) totaltrees " _
& " FROM tblextensionmortality where region='" & IIf(region = "Not Assigned", "", region) & "' group by grptshowog order by grptshowog", ODKDB
Do While rs.EOF <> True
         
xl.Cells(i, 1) = IIf(IsNull(rs!grptshowog), "", rs!grptshowog)
xl.Cells(i, 2) = IIf(IsNull(rs!cnt), 0, rs!cnt)
xl.Cells(i, 3) = IIf(IsNull(rs!dm), 0, rs!dm)
xl.Cells(i, 5) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
i = i + 1
rs.MoveNext

Loop


End Sub



Private Sub fillmortalitybyacre(xl As Object)
Dim rs As New ADODB.Recordset
Dim ge As String
Dim ttrees As Long
Dim i As Integer
Set rs = Nothing
i = 50
rs.Open "SELECT `region`,sum(acrereg) acrereg,rangeid FROM `tblextensionmortality` " _
& " group by region,rangeid order by  region,rangeid", ODKDB
Do Until rs.EOF
         ge = rs!region
         
             xl.Cells(i, 9) = IIf(Len(rs!region) = 0, "Not Assigned", rs!region)
          ttrees = 0
Do While ge = rs!region

Select Case rs!rangeid
Case 1
xl.Cells(i, 10) = IIf(IsNull(rs!acrereg), 0, rs!acrereg)
Case 2
xl.Cells(i, 11) = IIf(IsNull(rs!acrereg), 0, rs!acrereg)
Case 3
xl.Cells(i, 12) = IIf(IsNull(rs!acrereg), 0, rs!acrereg)
Case 4
xl.Cells(i, 13) = IIf(IsNull(rs!acrereg), 0, rs!acrereg)


End Select


rs.MoveNext
If rs.EOF Then Exit Do
Loop


i = i + 1
Loop




End Sub

Private Sub fillmortalitybyalt(xl As Object)
Dim rs As New ADODB.Recordset
Dim ge As String
Dim ttrees As Long
Dim i As Integer
Set rs = Nothing
i = 16
rs.Open "SELECT `grpgewog`,sum(totaltrees) totaltrees,region,`rangeidalt`,sum(deadmissing)/sum(totaltrees) per FROM `tblextensionmortality` " _
& " group by grpgewog,rangeidalt order by  grpgewog,rangeidalt", ODKDB
Do Until rs.EOF
         ge = rs!grpgewog
          xl.Cells(i, 1) = IIf(IsNull(rs!grpgewog), "", rs!grpgewog)
             xl.Cells(i, 2) = IIf(Len(rs!region) = 0, "Not Assigned", rs!region)
          ttrees = 0
Do While ge = rs!grpgewog

Select Case rs!rangeidalt
Case 2
xl.Cells(i, 4) = IIf(IsNull(rs!per), 0, rs!per)
Case 3
xl.Cells(i, 5) = IIf(IsNull(rs!per), 0, rs!per)
Case 4
xl.Cells(i, 6) = IIf(IsNull(rs!per), 0, rs!per)
Case 5
xl.Cells(i, 7) = IIf(IsNull(rs!per), 0, rs!per)


End Select

ttrees = ttrees + rs!totaltrees
rs.MoveNext
If rs.EOF Then Exit Do
Loop

xl.Cells(i, 3) = ttrees
i = i + 1
Loop




End Sub
Private Sub updatemortalityrecord()
Dim rs As New ADODB.Recordset

ODKDB.Execute "update odk_prodlocal.tblextensionmortality set rangeid='0'"
ODKDB.Execute "update odk_prodlocal.tblextensionmortality set rangeidalt='0'"

Set rs = Nothing
rs.Open "SELECT * from odk_prodlocal.tblextensionmortalityrange order by rangeid", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "update odk_prodlocal.tblextensionmortality set rangeid='" & rs!rangeid & "' where percent>='" & rs!minrange & "' and percent<'" & rs!maxrange & "'"
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "SELECT * from odk_prodlocal.tblextensionmortalityrangealt order by rangeid", ODKDB
Do While rs.EOF <> True
ODKDB.Execute "update odk_prodlocal.tblextensionmortality set rangeidalt='" & rs!rangeid & "' where alt>='" & rs!minrange & "' and alt<'" & rs!maxrange & "'"
rs.MoveNext
Loop

ODKDB.Execute "update odk_prodlocal.tblextensionmortality a ,mhv.tbltshewog b set region=regioncode where concat(dzongkhagid,gewogid,tshewogid)=substring(farmercode,1,9)"



End Sub
Private Sub field()
    Dim xl As Excel.Application
    Dim rs As New ADODB.Recordset
    Dim rschk As New ADODB.Recordset
    Dim SQLSTR As String
    Dim var As Variant
    Dim i, j As Integer
    
    
    
    Set rschk = Nothing
rschk.Open "select * from tblworstfieldsettings where isnextupdate='0' ", ODKDB
If rschk.EOF <> True Then
MsgBox "Cannot Proceed."
Exit Sub
End If

Set rschk = Nothing
rschk.Open "select * from tblworsttshowogsettings where isnextupdate='0'", ODKDB
If rschk.EOF <> True Then
MsgBox "Cannot Proceed."
Exit Sub
End If


Set rschk = Nothing
rschk.Open "select * from tblworstfieldsettings where nextScriptRunDate='" & Format(Now, "yyyy-MM-dd") & "'  ", ODKDB
If rschk.EOF <> True Then

Else
MsgBox "Cannot Proceed."
Exit Sub
End If

Set rschk = Nothing
rschk.Open "select * from tblworsttshowogsettings where nextScriptRunDate='" & Format(Now, "yyyy-MM-dd") & "'", ODKDB
If rschk.EOF <> True Then
Else
MsgBox "Cannot Proceed."
Exit Sub
End If
    
    
   
    Set xl = CreateObject("excel.Application")
  
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    End If
    
     Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='4'", MHVDB
    If rs.EOF <> True Then
    getSheet 4, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    xl.Sheets("Data").Select
    xl.Visible = False
    GetTbl
        
    
        SQLSTR = ""

           
         SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,sname,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,staffbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
           
           
  ODKDB.Execute SQLSTR

 SQLSTR = "select * from " & Mtblname & " where totaltrees>=100   "
  i = 2
  mchk = True
  Set rs = Nothing
    rs.Open SQLSTR, ODKDB
  Do While rs.EOF <> True
  
    FindDZ Mid(rs!farmercode, 1, 3)
    xl.Cells(i, 1) = Mid(rs!farmercode, 1, 3) & " " & Dzname
    
    FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
    xl.Cells(i, 2) = Mid(rs!farmercode, 4, 3) & " " & GEname

    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
    xl.Cells(i, 3) = Mid(rs!farmercode, 7, 3) & " " & TsName
    
    FindFA IIf(IsNull(rs!farmercode), "", rs!farmercode), "F"
    xl.Cells(i, 4) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & " " & FAName
    
    xl.Cells(i, 5) = "'" & rs!End
    xl.Cells(i, 6) = DateDiff("d", rs!End, Now - 2)
    
    xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
    xl.Cells(i, 8) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
    xl.Cells(i, 9) = IIf(IsNull(rs!tree_count_activegrowing), "", rs!tree_count_activegrowing)
    xl.Cells(i, 10) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
    xl.Cells(i, 11) = IIf(IsNull(rs!activepest), "", rs!activepest)
    xl.Cells(i, 12) = IIf(IsNull(rs!stempest), "", rs!stempest)
    xl.Cells(i, 13) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
    xl.Cells(i, 14) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
     
i = i + 1
rs.MoveNext
   Loop
' xl.Sheets("10 Worst Farmers").Select
'worst10 xl
'
'xl.Sheets("10 Worst Tshowog").Select
'worst10tshowog xl
'
'
'
'ODKDB.Execute "drop table " & Mtblname & ""
'
'ODKDB.Execute "insert into tblworstfieldsettings(lastScriptRunDate,generatReport,fieldstorage) " _
'& "('" & Format(Now, "yyyy-MM-dd") & "','No','F')"

xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault
End Sub
Private Sub worst10(xl As Object)
Dim rs As New ADODB.Recordset
Dim SQLSTR As String

'dead missing
i = 3
Set rs = Nothing
SQLSTR = "select farmercode,sname,fdcode,sum(totaltrees) as totaltrees,sum(tree_count_deadmissing) " _
& " affected,sum(tree_count_deadmissing/totaltrees) as percent,datediff(CURDATE(),end) as recordage " _
& " from " & Mtblname & " where  totaltrees>=100 and farmercode not in(select farmercode from " _
& " tblworstfarmers where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='DM') " _
& " group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
FindsTAFF rs!sname
xl.Cells(i, 1) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='DM' and farmercode='" & rs!farmercode & "'"

ODKDB.Execute "update tblworstfarmers set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='DM'"

ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','DM','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Dead Missing','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"


i = i + 1
rs.MoveNext
Loop

'water logged
i = 3
Set rs = Nothing
SQLSTR = "select farmercode,sname,fdcode,sum(totaltrees) as totaltrees,sum(waterlog) affected, " _
& " sum(waterlog/totaltrees) as percent,datediff(CURDATE(),end) as recordage from " & Mtblname & " " _
& " where totaltrees>=100 and farmercode not in(select farmercode from " _
& " tblworstfarmers where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='WL') " _
& " group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
FindsTAFF rs!sname
xl.Cells(i, 5) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 6) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 8) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='WL' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "update tblworstfarmers set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='WL'"

ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','WL','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Waterlogged','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"
i = i + 1
rs.MoveNext
Loop


'active pest
i = 3
Set rs = Nothing
SQLSTR = "select farmercode,sname,fdcode,sum(totaltrees) as totaltrees,sum(activepest) affected, " _
& " sum(activepest/totaltrees) as percent,datediff(CURDATE(),end) as recordage from " & Mtblname & " " _
& " where totaltrees>=100 and farmercode not in(select farmercode from " _
& " tblworstfarmers where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='AP') " _
& " group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
FindsTAFF rs!sname
xl.Cells(i, 9) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 10) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 11) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 12) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AP' and farmercode='" & rs!farmercode & "'"

ODKDB.Execute "update tblworstfarmers set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AP'"

ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AP','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Active Pest' ,'" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "')"

i = i + 1
rs.MoveNext
Loop

'Stem pest
i = 15
Set rs = Nothing
SQLSTR = "select farmercode,sname,fdcode,sum(totaltrees) as totaltrees,sum(stempest) affected, " _
& " sum(stempest/totaltrees) as percent,datediff(CURDATE(),end) as recordage from " & Mtblname & " " _
& " where totaltrees>=100 and farmercode not in(select farmercode from " _
& " tblworstfarmers where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='SP') " _
& " group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
FindsTAFF rs!sname
xl.Cells(i, 1) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='SP' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "update tblworstfarmers set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='SP'"

ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','SP','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Stem Pest', " _
& " '" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"


i = i + 1
rs.MoveNext
Loop

'Root pest
i = 15
Set rs = Nothing
SQLSTR = "select farmercode,sname,fdcode,sum(totaltrees) as totaltrees,sum(rootpest) affected, " _
& " sum(rootpest/totaltrees) as percent,datediff(CURDATE(),end) as recordage from " & Mtblname & " " _
& " where totaltrees>=100 and farmercode not in(select farmercode from " _
& " tblworstfarmers where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='RP')" _
& " group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
FindsTAFF rs!sname
xl.Cells(i, 5) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 6) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 8) = IIf(IsNull(rs!Percent), "", rs!Percent)


ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='RP' and farmercode='" & rs!farmercode & "'"

ODKDB.Execute "update tblworstfarmers set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='RP'"


ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','RP','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Root Pest','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"



i = i + 1
rs.MoveNext
Loop

'Animal Damage
i = 15
Set rs = Nothing
SQLSTR = "select farmercode,sname,fdcode,sum(totaltrees) as totaltrees,sum(animaldamage) affected, " _
& " sum(animaldamage/totaltrees) as percent,datediff(CURDATE(),end) as recordage from " & Mtblname & " " _
& " where totaltrees>=100 and farmercode not in(select farmercode from " _
& " tblworstfarmers where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='AD') group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
FindsTAFF rs!sname
xl.Cells(i, 9) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 10) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 11) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 12) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AD' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "update tblworstfarmers set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AD'"

ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AD','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Animal Damage','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"



i = i + 1
rs.MoveNext
Loop


End Sub
Private Sub storageworst10(xl As Object)
Dim rs As New ADODB.Recordset
Dim SQLSTR As String

'dead missing
i = 3
Set rs = Nothing
SQLSTR = "select farmercode,fdcode,sum(totaltrees) as totaltrees,sum(tree_count_deadmissing) as affected,sum(tree_count_deadmissing/totaltrees) as percent,datediff(CURDATE()-2,end) as recordage from " & Mtblname & " where  totaltrees>=100 group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
xl.Cells(i, 1) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='DM' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','DM','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Dead Missing' )"



i = i + 1
rs.MoveNext
Loop

'water logged
i = 3
Set rs = Nothing
SQLSTR = "select farmercode,fdcode,sum(totaltrees) as totaltrees,sum(waterlog) affected,sum(waterlog/totaltrees) as percent,datediff(CURDATE()-2,end) as recordage from " & Mtblname & "  where totaltrees>=100 group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
xl.Cells(i, 5) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 6) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 8) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='WL' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','WL','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Waterlogged' )"




i = i + 1
rs.MoveNext
Loop


'active pest
i = 3
Set rs = Nothing
SQLSTR = "select farmercode,fdcode,sum(totaltrees) as totaltrees,sum(activepest) affected,sum(activepest/totaltrees) as percent,datediff(CURDATE()-2,end) as recordage from " & Mtblname & "  where totaltrees>=100 group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
xl.Cells(i, 9) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 10) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 11) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 12) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='AP' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AP','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Pest Damage' )"

i = i + 1
rs.MoveNext
Loop

'animaldamage
i = 15
Set rs = Nothing
SQLSTR = "select farmercode,fdcode,sum(totaltrees) as totaltrees,sum(animaldamage) affected,sum(animaldamage/totaltrees) as percent,datediff(CURDATE()-2,end) as recordage from " & Mtblname & "  where totaltrees>=100 group by farmercode,fdcode order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
xl.Cells(i, 1) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & "  " & FAName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworstfarmers where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='AD' and farmercode='" & rs!farmercode & "'"


ODKDB.Execute "insert into tblworstfarmers (entrydate,parametertype,farmercode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AD','" & rs!farmercode & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Animal Damage' )"


i = i + 1
rs.MoveNext
Loop


End Sub

Private Sub worst10tshowog(xl As Object)
Dim rs As New ADODB.Recordset
Dim SQLSTR As String

'dead missing
i = 3
Set rs = Nothing
SQLSTR = "select sname,substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees, " _
& "sum(tree_count_deadmissing) affected,sum(tree_count_deadmissing)/sum(totaltrees) as percent, " _
& "avg(datediff(CURDATE(),end)) as " _
& "recordage from " & Mtblname & "  where  substring(farmercode,1,9) not in(select dgtcode from " _
& " tblworsttshowog where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='AD') group by " _
& " substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"


rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindsTAFF rs!sname
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)

xl.Cells(i, 1) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName

xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='DM' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "update tblworsttshowog set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AD'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','DM','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Dead Missing','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"




i = i + 1
rs.MoveNext
Loop

'water logged
i = 3
Set rs = Nothing
SQLSTR = "select sname,substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(waterlog) affected,sum(waterlog)/sum(totaltrees) " _
& "as percent,avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & " where  substring(farmercode,1,9) not in(select dgtcode from " _
& " tblworsttshowog where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='WL')  " _
& " group by substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindsTAFF rs!sname
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 5) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 6) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 8) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='WL' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "update tblworsttshowog set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='WL'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','WL','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Waterlogged','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"




i = i + 1
rs.MoveNext
Loop


'active pest
i = 3
Set rs = Nothing
SQLSTR = "select sname,substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(activepest) affected,sum(activepest)/sum(totaltrees)" _
& " as percent,avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  where  substring(farmercode,1,9) not in(select dgtcode from " _
& " tblworsttshowog where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='AP') " _
& " group by substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindsTAFF rs!sname
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 9) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 10) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 11) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 12) = IIf(IsNull(rs!Percent), "", rs!Percent)


ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AP' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "update tblworsttshowog set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AP'"


ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AP','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Active Pest' ,'" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "')"



i = i + 1
rs.MoveNext
Loop

'Stem pest
i = 15
Set rs = Nothing
SQLSTR = "select sname,substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(stempest) affected,sum(stempest)/sum(totaltrees) as percent, " _
& " avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  where  substring(farmercode,1,9) not in(select dgtcode from " _
& " tblworsttshowog where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='SP')  " _
& " group by substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindsTAFF rs!sname
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 1) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='SP' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "update tblworsttshowog set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='SP'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','SP','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Stem Pest','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"


i = i + 1
rs.MoveNext
Loop

'Root pest
i = 15
Set rs = Nothing
SQLSTR = "select sname,substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(rootpest) affected, " _
& " sum(rootpest)/sum(totaltrees) as percent, " _
& " avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  where  substring(farmercode,1,9) not in(select dgtcode from " _
& " tblworsttshowog where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='RP')  " _
& " group by " _
& " substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindsTAFF rs!sname
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 5) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 6) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 8) = IIf(IsNull(rs!Percent), "", rs!Percent)



ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='RP' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "update tblworsttshowog set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='RP'"


ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','RP','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Root Pest' ,'" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "')"


i = i + 1
rs.MoveNext
Loop

'Animal Damage
i = 15
Set rs = Nothing
SQLSTR = "select sname,substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(animaldamage) affected,sum(animaldamage)/sum(totaltrees) as percent, " _
& " avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  where  substring(farmercode,1,9) not in(select dgtcode from " _
& " tblworsttshowog where nextdate>'" & Format(Now, "yyyy-MM-dd") & "' and parametertype='AD') " _
& " group by " _
& " substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"





rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindsTAFF rs!sname
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 9) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 10) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 11) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 12) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AD' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "update tblworsttshowog set islast='1' where entrydate<'" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='F' and parametertype='AD'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay,monitor,nextdate) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AD','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','F','No','Animal Damage','" & rs!sname & "  " & sTAFF & "','" & Format(DateAdd("d", 90, Format(Now, "yyyy-MM-dd")), "yyyy-MM-dd") & "' )"



i = i + 1
rs.MoveNext
Loop

End Sub
Private Sub storageworst10tshowog(xl As Object)

Dim rs As New ADODB.Recordset
Dim SQLSTR As String

'dead missing
i = 3
Set rs = Nothing
SQLSTR = "select substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees, " _
& "sum(tree_count_deadmissing) affected,sum(tree_count_deadmissing)/sum(totaltrees) as percent,avg(datediff(CURDATE(),end)) as " _
& "recordage from " & Mtblname & "  group by " _
& " substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)

xl.Cells(i, 1) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='DM' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','DM','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Dead Missing' )"


i = i + 1
rs.MoveNext
Loop

'water logged
i = 3
Set rs = Nothing
SQLSTR = "select substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(waterlog) affected,sum(waterlog)/sum(totaltrees) " _
& "as percent,avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  " _
& " group by substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"





rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)

xl.Cells(i, 5) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 6) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 7) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 8) = IIf(IsNull(rs!Percent), "", rs!Percent)


ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='WL' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','WL','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Waterlogged' )"


i = i + 1
rs.MoveNext
Loop


'active pest
i = 3
Set rs = Nothing
SQLSTR = "select substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(activepest) affected,sum(activepest)/sum(totaltrees)" _
& " as percent,avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  " _
& " group by substring(farmercode,1,9)  HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"
rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 9) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 10) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 11) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 12) = IIf(IsNull(rs!Percent), "", rs!Percent)

ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='AP' and dgtcode='" & rs!dgt & "'"



ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AP','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Pest Damage' )"


i = i + 1
rs.MoveNext
Loop



'Animal Damage
i = 15
Set rs = Nothing
SQLSTR = "select substring(farmercode,1,9) dgt,sum(totaltrees) as totaltrees,sum(animaldamage) affected,sum(animaldamage)/sum(totaltrees) as percent, " _
& " avg(datediff(CURDATE(),end)) as recordage from " & Mtblname & "  where (totaltrees)>=1000 group by " _
& " substring(farmercode,1,9) HAVING SUM(totaltrees) >=1000 order by percent desc limit 10"


rs.Open SQLSTR, ODKDB
Do While rs.EOF <> True
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)
xl.Cells(i, 1) = Mid(rs!dgt, 1, 3) & "  " & Dzname & "  " & Mid(rs!dgt, 4, 3) & "  " & GEname & "  " & Mid(rs!dgt, 7, 3) & "  " & TsName
xl.Cells(i, 2) = IIf(IsNull(rs!recordage), "", rs!recordage)
xl.Cells(i, 3) = IIf(IsNull(rs!totaltrees), "", rs!totaltrees)
xl.Cells(i, 4) = IIf(IsNull(rs!Percent), "", rs!Percent)



ODKDB.Execute "delete from tblworsttshowog where entrydate='" & Format(Now, "yyyy-MM-dd") & "' " _
& " and fieldstorage='S' and parametertype='AD' and dgtcode='" & rs!dgt & "'"

ODKDB.Execute "insert into tblworsttshowog (entrydate,parametertype,dgtcode,recordage,totaltrees, " _
& " totalaffected,percent,fieldstorage,status,paradisplay) values " _
& " ('" & Format(Now, "yyyy-MM-dd") & "','AD','" & rs!dgt & "','" & rs!recordage & "', " _
& " '" & rs!totaltrees & "','" & rs!affected & "','" & rs!Percent & "','S','No','Animal Damage' )"


i = i + 1
rs.MoveNext
Loop



End Sub


Private Sub outreach()
Dim SQLSTR As String
Dim marray
Dim i, colcnt As Integer
Dim rs As New ADODB.Recordset



Dim xl As Excel.Application
    Dim var As Variant
    Set xl = CreateObject("excel.Application")

    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    End If
    
      Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='3'", MHVDB
    If rs.EOF <> True Then
    getSheet 3, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    xl.Sheets("Monitor Data").Select
    xl.Visible = False
    
    Select Case cbomnth.ListIndex
     Case 0
     
     Case 1
     
     Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
     
    End Select
    
    monitor xl
    xl.Sheets("Advocate Data").Select
ADVOCATE xl
'
 xl.Sheets("Outreach").Select
outreachsummary xl
    xl.Visible = True









End Sub
Private Sub outreachsummary(xl As Object)
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim tr As New ADODB.Recordset
Dim rsytd As New ADODB.Recordset
Dim trstr As String
Dim selop As Integer
Dim SQLSTR As String
Dim i, j As Integer
Dim currmonth, premonth, lastmonth, curryear, preyear, lastyear As Integer









'filltbl
selop = Month(Now) - 1
  Select Case selop
     Case 0
        curryear = Year(Now)
        preyear = Year(Now) - 1
        lastyear = Year(Now) - 1
        currmonth = 1
        premonth = 12
        lastmonth = 11
     
    Case 1
        curryear = Year(Now)
        preyear = Year(Now)
        lastyear = Year(Now) - 1
        currmonth = 2
        premonth = 1
        lastmonth = 12
     
    Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
        curryear = Year(Now)
        preyear = Year(Now)
        lastyear = Year(Now)
        currmonth = Month(Now)
        premonth = Month(Now) - 1
        lastmonth = Month(Now) - 2
     End Select
    xl.Cells(2, 17) = MonthName(lastmonth, True)
    xl.Cells(2, 18) = MonthName(premonth, True)
    xl.Cells(2, 19) = MonthName(currmonth, True)
          
'monitors
i = 0
Set rs = Nothing
SQLSTR = ""
SQLSTR = "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & lastyear & "' " _
& " and month(regdate) in('" & lastmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14) not " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & preyear & "' " _
& " and month(regdate) in('" & premonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14) not " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & curryear & "' " _
& " and month(regdate) in('" & currmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14) not " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) "


rs.Open SQLSTR, MHVDB

Do While rs.EOF <> True
For j = 0 To 2
If Mid(UCase(xl.Cells(2, 17 + j)), 1, 3) = UCase(MonthName(rs!mnth, True)) Then
xl.Cells(5, 17 + j) = Format(rs!regland, "##0.00")
End If
Next
i = i + 1
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select sum(regland) regland from tblregistrationrpt where " _
& " substring(farmercode,10,1)='F' and staffcode " _
& " in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14) not in(select farmercode " _
& " from tblplanted) ", MHVDB
Do While rs.EOF <> True
xl.Cells(5, 20) = Format(rs!regland, "##0.00")
rs.MoveNext
Loop



i = 0
Set rs = Nothing
SQLSTR = ""
SQLSTR = "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & lastyear & "' " _
& " and month(regdate) in('" & lastmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14)  " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & preyear & "' " _
& " and month(regdate) in('" & premonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14)  " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & curryear & "' " _
& " and month(regdate) in('" & currmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14)  " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) "


rs.Open SQLSTR, MHVDB




Do While rs.EOF <> True
For j = 0 To 2
If UCase(xl.Cells(2, 17 + j)) = UCase(MonthName(rs!mnth, True)) Then
xl.Cells(6, 17 + j) = Format(rs!regland, "##0.00")
End If
Next
i = i + 1
rs.MoveNext
Loop
Set rs = Nothing
rs.Open "select sum(regland) regland from tblregistrationrpt where " _
& "substring(farmercode,10,1)='F'  " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') and substring(farmercode,1,14)  in(select farmercode " _
& "from tblplanted)  ", MHVDB
Do While rs.EOF <> True
xl.Cells(6, 20) = Format(rs!regland, "##0.00")
rs.MoveNext
Loop

i = 0

SQLSTR = ""
SQLSTR = "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)<>'F' and year(regdate)='" & lastyear & "' " _
& " and month(regdate) in('" & lastmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)<>'F' and year(regdate)='" & preyear & "' " _
& " and month(regdate) in('" & premonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)<>'F' and year(regdate)='" & curryear & "' " _
& " and month(regdate) in('" & currmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where moniter='1') group by year(regdate),month(regdate) "

Set rs = Nothing
rs.Open SQLSTR, MHVDB




Do While rs.EOF <> True
For j = 0 To 2
If UCase(xl.Cells(2, 17 + j)) = UCase(MonthName(rs!mnth, True)) Then
xl.Cells(7, 17 + j) = Format(rs!regland, "##0.00")
End If
Next
i = i + 1
rs.MoveNext
Loop
Set rs = Nothing
rs.Open "select sum(regland) regland from tblregistrationrpt where " _
& "substring(farmercode,10,1)<>'F' " _
& "and staffcode in(select staffcode from tblmhvstaff where moniter='1') ", MHVDB
Do While rs.EOF <> True
xl.Cells(7, 20) = Format(rs!regland, "##0.00")
rs.MoveNext
Loop

' advocate

i = 0
Set rs = Nothing
SQLSTR = ""
SQLSTR = "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & lastyear & "' " _
& " and month(regdate) in('" & lastmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14) not " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & preyear & "' " _
& " and month(regdate) in('" & premonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14) not " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & curryear & "' " _
& " and month(regdate) in('" & currmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14) not " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) "


rs.Open SQLSTR, MHVDB

Do While rs.EOF <> True
For j = 0 To 2
If Mid(UCase(xl.Cells(2, 17 + j)), 1, 3) = UCase(MonthName(rs!mnth, True)) Then
xl.Cells(13, 17 + j) = Format(rs!regland, "##0.00")
End If
Next
i = i + 1
rs.MoveNext
Loop

Set rs = Nothing
rs.Open "select sum(regland) regland from tblregistrationrpt where " _
& " substring(farmercode,10,1)='F' and staffcode " _
& " in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14) not in(select farmercode " _
& " from tblplanted) ", MHVDB
Do While rs.EOF <> True
xl.Cells(13, 20) = Format(rs!regland, "##0.00")
rs.MoveNext
Loop



i = 0
Set rs = Nothing
SQLSTR = ""
SQLSTR = "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & lastyear & "' " _
& " and month(regdate) in('" & lastmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14)  " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & preyear & "' " _
& " and month(regdate) in('" & premonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14)  " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)='F' and year(regdate)='" & curryear & "' " _
& " and month(regdate) in('" & currmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14)  " _
& " in(select farmercode from tblplanted) group by year(regdate),month(regdate) "


rs.Open SQLSTR, MHVDB




Do While rs.EOF <> True
For j = 0 To 2
If UCase(xl.Cells(2, 17 + j)) = UCase(MonthName(rs!mnth, True)) Then
xl.Cells(14, 17 + j) = Format(rs!regland, "##0.00")
End If
Next
i = i + 1
rs.MoveNext
Loop
Set rs = Nothing
rs.Open "select sum(regland) regland from tblregistrationrpt where " _
& "substring(farmercode,10,1)='F'  " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') and substring(farmercode,1,14)  in(select farmercode " _
& "from tblplanted)  ", MHVDB
Do While rs.EOF <> True
xl.Cells(14, 20) = Format(rs!regland, "##0.00")
rs.MoveNext
Loop

i = 0

SQLSTR = ""
SQLSTR = "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)<>'F' and year(regdate)='" & lastyear & "' " _
& " and month(regdate) in('" & lastmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)<>'F' and year(regdate)='" & preyear & "' " _
& " and month(regdate) in('" & premonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') group by year(regdate),month(regdate) " _
& " union " _
& "select year(regdate) yr,month(regdate) mnth,sum(regland) regland from " _
& " tblregistrationrpt where substring(farmercode,10,1)<>'F' and year(regdate)='" & curryear & "' " _
& " and month(regdate) in('" & currmonth & "') " _
& " and staffcode in(select staffcode from tblmhvstaff where advocate='1') group by year(regdate),month(regdate) "

Set rs = Nothing
rs.Open SQLSTR, MHVDB




Do While rs.EOF <> True
For j = 0 To 2
If UCase(xl.Cells(2, 17 + j)) = UCase(MonthName(rs!mnth, True)) Then
xl.Cells(15, 17 + j) = Format(rs!regland, "##0.00")
End If
Next
i = i + 1
rs.MoveNext
Loop
Set rs = Nothing
rs.Open "select sum(regland) regland from tblregistrationrpt where " _
& "substring(farmercode,10,1)<>'F' " _
& "and staffcode in(select staffcode from tblmhvstaff where advocate='1') ", MHVDB
Do While rs.EOF <> True
xl.Cells(15, 20) = Format(rs!regland, "##0.00")
rs.MoveNext
Loop


End Sub
Private Sub filltbl()
Dim SQLSTR As String
Dim marray
Dim i, colcnt As Integer
Dim rs As New ADODB.Recordset
SQLSTR = "SELECT farmerid ,'S' sharedtype,regdate,monitor shared, outreach shared1,  ''shared2,  ''shared3,  ''shared4, SUM( regland ) AS regland " _
      & "FROM  `tbllandreg` " _
& "Where Length(monitor) = 5 " _
& "AND LENGTH( outreach ) =5 and status not in('D','R') and year(regdate)='" & cboyear.BoundText & "' " _
& "GROUP BY farmerid,monitor, outreach,regdate " _
& "Union " _
& "SELECT farmerid ,'I' sharedtype,regdate, individual shared,  ''shared1,  ''shared2,  ''shared3,  ''sgared4, SUM( regland ) AS regland " _
& "FROM  `tbllandreg` " _
& "Where Length(individual) = 5  and status not in('D','R') and regdate>='2013-11-01' " _
& "GROUP BY farmerid,individual,regdate " _
& "Union " _
& "SELECT farmerid ,'S' sharedtype ,regdate,cgmonitor shared,  ''shared1,  ''shared2,  ''shared3,  ''sgared4, SUM( regland ) AS regland " _
& "FROM  `tbllandreg` " _
& "Where Length(cgmonitor) = 5   and status not in('D','R') and regdate>='2013-11-01' " _
& "GROUP BY farmerid, cgmonitor,regdate " _
& "Union " _
& " SELECT farmerid ,'S' sharedtype ,regdate,leadstaff shared,  `SUPPORT1` shared1,  `SUPPORT2` sahred2,  `SUPPORT3` shared3,  `SUPPORT4` shared4, SUM( regland ) AS regland " _
& " FROM  `tbllandreg` " _
& " Where Length(leadstaff) = 5  and status not in('D','R') and regdate>='2013-11-01' " _
& " GROUP BY farmerid ,leadstaff,support1,support2,support3,support4,regdate"
MHVDB.Execute "delete from  tblsharedtemp"
Set rs = Nothing
rs.Open SQLSTR, MHVDB

Do While rs.EOF <> True
colcnt = 0
    For i = 3 To 7
    If Len(Trim(rs.Fields(i).Value)) = 0 Then Exit For
    If Len(Trim(rs.Fields(i).Value)) = 5 Then
    colcnt = colcnt + 1
    
    End If
    Next

For i = 0 To colcnt - 1
MHVDB.Execute "insert into tblsharedtemp(staffcode,regland,farmercode,sharedtype,regdate)values( " _
& "'" & Trim(rs.Fields(i + 3).Value) & "','" & Trim(rs.Fields(8).Value) / colcnt & "','" & rs!farmerid & "','" & rs!sharedtype & "','" & Format(rs!regdate, "yyyy-MM-dd") & "'" _
& ")"
Next

rs.MoveNext
Loop
End Sub
Private Sub monitor(xl As Object)
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim tr As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim rsytd As New ADODB.Recordset
Dim actcnt As Integer
Dim trstr As String
Dim i, j As Integer
xl.Cells(2, 16) = Format(Now, "dd/MM/yyyy")
Set rs = Nothing
rs.Open "select staffcode from tblmhvstaff where moniter='1' order by msupervisor", MHVDB
i = 5
Do While rs.EOF <> True
FindsTAFF rs!staffcode
xl.Cells(i, 2) = rs!staffcode & "  " & sTAFF
sTAFF = ""
Set rs1 = Nothing
rs1.Open "select * from tblmhvstaff where staffcode='" & rs!staffcode & "'", MHVDB
If rs1.EOF <> True Then
FindsTAFF rs1!msupervisor
xl.Cells(i, 3) = rs1!msupervisor & "  " & sTAFF
xl.Cells(i, 4) = rs1!mteritory
End If

Set rs1 = Nothing
rs1.Open "select count(value) as cnt from dailyacthub9_activities where " _
& " _parent_auri in(select _uri from dailyacthub9_core where month(end)='" & Month(Now) & "' and staffbarcode='" & rs!staffcode & "') ", ODKDB
If rs1.EOF <> True Then
xl.Cells(i, 5) = rs1!cnt
End If

Set RS2 = Nothing
RS2.Open "select * from tblregistrationtargetdetail where staffcode='" & rs!staffcode & "'", MHVDB
If RS2.EOF <> True Then
xl.Cells(i, 6) = RS2!Target
End If

Set rs1 = Nothing
rs1.Open "select target,sum(pacre) pacre,sum(gacre) gacre,(sum(pacre)+sum(gacre)) tacre ," _
& " staffcode from tblregistrationrpt where month(regdate)='" & Month(Now) & "' and staffcode='" & rs!staffcode & "' " _
& " group by staffcode ", MHVDB
If rs1.EOF <> True Then


xl.Cells(i, 7) = rs1!pacre
xl.Cells(i, 8) = rs1!gacre
xl.Cells(i, 9) = rs1!tacre
End If
Set rs1 = Nothing
rs1.Open "select (sum(pacre)+sum(gacre)) ttacre " _
& "  from tblregistrationrpt where staffcode='" & rs!staffcode & "' " _
& " group by staffcode ", MHVDB
If rs1.EOF <> True Then
xl.Cells(i, 10) = rs1!ttacre
End If
i = i + 1
rs.MoveNext
Loop

End Sub
Private Sub ADVOCATE(xl As Object)
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim tr As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim rsytd As New ADODB.Recordset
Dim actcnt As Integer
Dim trstr As String
Dim i, j As Integer
xl.Cells(2, 16) = Format(Now, "dd/MM/yyyy")
Set rs = Nothing
rs.Open "select staffcode from tblmhvstaff where advocate='1'", MHVDB
i = 7
Do While rs.EOF <> True
FindsTAFF rs!staffcode
xl.Cells(i, 2) = sTAFF
xl.Cells(i, 3) = rs!staffcode

Set rs1 = Nothing
rs1.Open "select target,sum(pacre) pacre,sum(gacre) gacre,(sum(pacre)+sum(gacre)) tacre ," _
& " staffcode from tblregistrationrpt where month(regdate)='" & Month(Now) & "' " _
& " and staffcode='" & rs!staffcode & "' and substring(sharedtype,1,1)='I' " _
& " group by substring(sharedtype,1,1) ", MHVDB
If rs1.EOF <> True Then
xl.Cells(i, 4) = rs1!Target
xl.Cells(i, 5) = rs1!pacre
xl.Cells(i, 6) = rs1!gacre
End If

Set rs1 = Nothing
rs1.Open "select target,sum(pacre) pacre,sum(gacre) gacre,(sum(pacre)+sum(gacre)) tacre ," _
& " staffcode from tblregistrationrpt where month(regdate)='" & Month(Now) & "' " _
& " and staffcode='" & rs!staffcode & "' and substring(sharedtype,1,1)='S' " _
& " group by substring(sharedtype,1,1) ", MHVDB
If rs1.EOF <> True Then
xl.Cells(i, 4) = rs1!Target
xl.Cells(i, 7) = rs1!pacre
xl.Cells(i, 8) = rs1!gacre
End If



Set rs1 = Nothing
rs1.Open "select target,sum(pacre) pacre,sum(gacre) gacre,(sum(pacre)+sum(gacre)) tacre ," _
& " staffcode from tblregistrationrpt where  " _
& "  staffcode='" & rs!staffcode & "' and substring(sharedtype,1,1)='I' " _
& " group by substring(sharedtype,1,1) ", MHVDB
If rs1.EOF <> True Then
xl.Cells(i, 4) = rs1!Target
xl.Cells(i, 9) = rs1!pacre
xl.Cells(i, 10) = rs1!gacre
End If

Set rs1 = Nothing
rs1.Open "select target,sum(pacre) pacre,sum(gacre) gacre,(sum(pacre)+sum(gacre)) tacre ," _
& " staffcode from tblregistrationrpt where  " _
& "  staffcode='" & rs!staffcode & "' and substring(sharedtype,1,1)='S' " _
& " group by substring(sharedtype,1,1) ", MHVDB
If rs1.EOF <> True Then
xl.Cells(i, 4) = rs1!Target
xl.Cells(i, 11) = rs1!pacre
xl.Cells(i, 12) = rs1!gacre
End If



i = i + 1
rs.MoveNext
Loop
End Sub
Private Sub monitoring()
Dim SQLSTR As String
Dim i, j As Integer
Dim m As Integer
Dim col As Integer
Dim rs As New ADODB.Recordset
Dim mnthstr As String
    Dim xl As Excel.Application
    Dim var As Variant
    Set xl = CreateObject("excel.Application")
    
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    End If
    
    Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='6'", MHVDB
    If rs.EOF <> True Then
    getSheet 6, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsm"
    xl.Visible = True
    
    
    xl.Visible = False
    
    Select Case cbomnth.ListIndex
     Case 0
     
     Case 1
     
     Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
     
    End Select
    
   
    xl.Sheets("Farm Visits").Select
    farmvisit xl
    xl.Sheets("Daily Act").Select
    dailyact xl
    xl.Sheets("Monthly Activity Log").Select
    activity xl
    xl.Sheets("Farmers Not Visited").Select
    monitoringsmmary xl
    xl.Visible = True
    End Sub
Private Sub monitoringsmmary(xl As Object)

     
     Dim SQLSTR As String
     Dim i As Integer
Dim farmerstr As String
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblzerovisit where status<>'X'", MHVDB
i = 4
Do While rs.EOF <> True
xl.Cells(i, 2) = rs!farmername
xl.Cells(i, 3) = rs!farmercode
If Len(rs!staffcode) = 5 Then
xl.Cells(i, 4) = rs!staffcode & "  " & rs!staffname
Else
xl.Cells(i, 4) = rs!staffname
End If
i = i + 1
rs.MoveNext
Loop
     
        
        






















End Sub

Private Sub activity(xl As Object)
Dim i, j As Integer
Dim dt As Date
Dim mact As String
Dim mstaff As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim ra As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim SQLSTR As String

mygrid.Clear
For i = 1 To 14
mygrid.TextMatrix(0, i) = "activity" & i
Next
i = 1
j = 1
Set rsm = Nothing
SQLSTR = "SELECT staffcode from tblmhvstaff where moniter='1'"
rsm.Open SQLSTR, MHVDB
Do While rsm.EOF <> True
mygrid.TextMatrix(j, 0) = rsm!staffcode
For i = 1 To 14
Set rs1 = Nothing
mact = "activity" & i
rs1.Open "select count(*) cnt from dailyacthub9_activities where value='" & mact & "' and  _PARENT_AURI in(SELECT _URI FROM `dailyacthub9_core` WHERE staffbarcode='" & rsm!staffcode & "'  and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "')", ODKDB
mygrid.TextMatrix(j, i) = rs1!cnt

Next
j = j + 1
rsm.MoveNext
Loop



        
         
         


Dim rng As Range

For i = 1 To mygrid.Rows - 1

FindsTAFF mygrid.TextMatrix(i, 0)
xl.Cells(6 + i, 1) = mygrid.TextMatrix(i, 0) & "  " & sTAFF

If (Len(mygrid.TextMatrix(i, 0))) = 0 Then Exit Sub

For j = 1 To 14
If "activity" & xl.Cells(1, j + 1) = mygrid.TextMatrix(0, j) Then
Set ra = Nothing
ra.Open "select * from tbldailyactchoices where name='" & mygrid.TextMatrix(0, j) & "' ", ODKDB
If ra.EOF <> True Then


  Set rng = xl.Cells(6, j + 1)
    If rng.Comment Is Nothing Then rng.AddComment (ra!label)
    'rng.Comment.Text ra!label
    
End If
xl.Cells(6 + i, j + 1) = mygrid.TextMatrix(i, j)
End If
Next

Next



End Sub
Private Sub farmvisit(xl As Object)
Dim i, j As Integer
Dim dt As Date
Dim mstaff As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim SQLSTR As String
Dim dgtstr As String
Dim farmerstr As String
dt = txtfromdate.Value
mygrid.Clear
For i = 1 To 7
mygrid.TextMatrix(0, i) = WDayName(dt, 0)
dt = dt + 1
Next
i = 1
Set rs = Nothing
SQLSTR = "SELECT staffcode from tblmhvstaff where moniter='1'"
Set rs = Nothing
rs.Open SQLSTR, MHVDB

Do While rs.EOF <> True
mygrid.TextMatrix(i, 0) = rs!staffcode
i = i + 1
rs.MoveNext
Loop

For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 0)) = 0 Then Exit For
Set rs = Nothing
rs.Open "select count(*) cnt from tblfarmer where monitor='" & Trim(mygrid.TextMatrix(i, 0)) & "'", MHVDB
If rs.EOF <> True Then
mygrid.TextMatrix(i, 1) = rs!cnt
End If



dgtstr = ""
Set rsm = Nothing
rsm.Open "select distinct(substring(idfarmer,1,9)) as dgt from tblfarmer where monitor='" & Trim(mygrid.TextMatrix(i, 0)) & "'", MHVDB
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

'rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from phealthhub15_core where staffbarcode='" & mygrid.TextMatrix(i, 0) & "' and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from phealthhub15_core where substring(farmerbarcode,1,9) in " & dgtstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
If rs.EOF <> True Then
mygrid.TextMatrix(i, 2) = rs!cnt
End If

farmerstr = ""
Set rs1 = Nothing
rs1.Open "select distinct farmerbarcode from phealthhub15_core where substring(farmerbarcode,1,9) in " & dgtstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
Do While rs1.EOF <> True
farmerstr = farmerstr + "'" + Trim(rs1!farmerbarcode) + "',"
rs1.MoveNext
Loop


If Len(farmerstr) > 0 Then
farmerstr = "(" + Left(farmerstr, Len(farmerstr) - 1) + ")"
Else
farmerstr = "(" + "'" + A99 & "'" & ")"
End If


Set rs = Nothing
'rs.Open "select count(distinct farmerbarcode) fcnt from storagehub6_core where  staffbarcode='" & mygrid.TextMatrix(i, 0) & "' and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from storagehub6_core where substring(farmerbarcode,1,9) in " & dgtstr & " and farmerbarcode not in " & farmerstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB

If rs.EOF <> True Then

mygrid.TextMatrix(i, 3) = rs!cnt
End If

Set rs = Nothing
'rs.Open "select count(distinct farmerbarcode) fcnt from storagehub6_core where  staffbarcode='" & mygrid.TextMatrix(i, 0) & "' and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from storagehub6_core where substring(farmerbarcode,1,9) in " & dgtstr & " and  substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB

If rs.EOF <> True Then

mygrid.TextMatrix(i, 4) = rs!cnt
End If

Next

xl.Cells(2, 5) = "'" & Format(Now, "dd/MM/yyyy hh:mm:ss")


For i = 1 To mygrid.Rows - 1
If (Len(mygrid.TextMatrix(i, 0))) = 0 Then Exit Sub
FindsTAFF mygrid.TextMatrix(i, 0)
Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & mygrid.TextMatrix(i, 0) & "' and moniter='1'", MHVDB
If rs.EOF <> True Then
xl.Cells(4 + i, 2) = mygrid.TextMatrix(i, 0)
xl.Cells(4 + i, 3) = sTAFF
xl.Cells(4 + i, 4) = mygrid.TextMatrix(i, 1)
xl.Cells(4 + i, 5) = mygrid.TextMatrix(i, 2)
xl.Cells(4 + i, 6) = mygrid.TextMatrix(i, 4)
xl.Cells(4 + i, 7) = Val(mygrid.TextMatrix(i, 3)) + Val(mygrid.TextMatrix(i, 2))
End If
Next





End Sub
Private Sub dailyact(xl As Object)
Dim i, j As Integer
Dim dt As Date
Dim mstaff As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim SQLSTR As String
dt = Now - 47
mygrid.Clear



For i = 1 To 45
mygrid.TextMatrix(0, i) = Format(dt, "dd/MM/yyyy")
xl.Cells(8, 4 + i) = dt
dt = dt + 1
Next
i = 1
Set rsm = Nothing
'SQLSTR = "SELECT staffcode from tblmhvstaff where moniter='1'"


Set rs = Nothing
SQLSTR = "SELECT _URI,end,staffbarcode,year(end) year, month(end) month ,day(END) as day FROM `dailyacthub9_core` WHERE substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "' order by staffbarcode,END"
rs.Open SQLSTR, ODKDB

Do Until rs.EOF
mstaff = rs!staffbarcode

Do While mstaff = rs!staffbarcode
'activity4
'If mstaff = "S0331" Then
'MsgBox "jj"
'End If
Set rs1 = Nothing
rs1.Open "select count(*) cnt from dailyacthub9_activities where _PARENT_AURI='" & rs![_uri] & "'", ODKDB
If rs1.EOF <> True Then
mygrid.TextMatrix(i, 0) = rs!staffbarcode

For j = 1 To 45
If Trim(mygrid.TextMatrix(0, j)) = Format(Trim(Mid(rs!End, 1, 10)), "dd/MM/yyyy") Then
mygrid.TextMatrix(i, j) = 1
End If
Next
End If


rs.MoveNext
         If rs.EOF Then Exit Do
         Loop
         i = i + 1
         Loop
         
Set rs = Nothing
rs.Open "select staffcode from mhv.tblmhvstaff where moniter='1' and staffcode not in(SELECT staffbarcode FROM `dailyacthub9_core` WHERE substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "')", ODKDB
Do While rs.EOF <> True
mygrid.TextMatrix(i, 0) = rs!staffcode
i = i + 1
rs.MoveNext
Loop
         
         
         

'For i = 0 To 44
'
'
'Next


For i = 1 To mygrid.Rows - 1
'If mygrid.TextMatrix(i, 0) = "S0331" Then
'MsgBox "hf"
'End If
If (Len(mygrid.TextMatrix(i, 0))) = 0 Then Exit Sub
For j = 0 To 44

If Format(xl.Cells(8, 5 + j), "dd/MM/yyyy") = mygrid.TextMatrix(0, 1 + j) Then
FindsTAFF mygrid.TextMatrix(i, 0)
xl.Cells(8 + i, 2) = i
xl.Cells(8 + i, 3) = mygrid.TextMatrix(i, 0) & "  " & sTAFF
xl.Cells(8 + i, 5 + j) = mygrid.TextMatrix(i, 1 + j)

End If

Next

Next






End Sub
Private Sub acceptedforstagging()
Dim arrValues(1 To 12, 1 To 3)
Dim i As Integer
For i = 1 To 12
   mygridKATC.TextMatrix(0, i) = mygridKATC.TextMatrix(0, i)
   arrValues(i, 1) = mygridKATC.TextMatrix(0, i)
   arrValues(i, 2) = CLng(Val(Format(mygridKATC.TextMatrix(1, i), "######")))
   arrValues(i, 3) = CLng(Val(Format(mygridKATC.TextMatrix(2, i), "######")))
  
Next i


Frame2.Caption = "Accepted for Staging"
MSChart1.ChartData = arrValues
MSChart1.Column = 1
MSChart1.ColumnLabel = "Target"
MSChart1.Column = 2
MSChart1.ColumnLabel = "Actual"
Frame2.Visible = True
End Sub

Private Sub KATCdefects()
Dim arrValues(1 To 12, 1 To 3)
Dim i As Integer
For i = 1 To 12
   mygridKATC.TextMatrix(0, i) = mygridKATC.TextMatrix(0, i)
   arrValues(i, 1) = mygridKATC.TextMatrix(0, i)
   arrValues(i, 2) = mygridKATC.TextMatrix(3, i)
   arrValues(i, 3) = mygridKATC.TextMatrix(4, i)
  
Next i


Frame3.Caption = "KATC Defects"
MSChart2.ChartData = arrValues
MSChart2.Column = 1
MSChart2.ColumnLabel = "Defect %"
MSChart2.Column = 2
MSChart2.ColumnLabel = "Cost of Defects(Nu.)'000"
Frame3.Visible = True
End Sub


Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Command3_Click()
frmdashboardnursery.Width = 5685
End Sub

Private Sub ehkmailme_Click()

End Sub

Private Sub Command4_Click()
frmdashboardnursery.Width = 10395
End Sub

Private Sub Command5_Click()
Dim SQLSTR As String
Dim marray
Dim i, colcnt As Integer
Dim rs As New ADODB.Recordset
SQLSTR = "SELECT farmerid ,'S' sharedtype,regdate,monitor shared, outreach shared1,  ''shared2,  ''shared3,  ''shared4, SUM( regland ) AS regland " _
      & "FROM  `tbllandreg` " _
& "Where Length(monitor) = 5 " _
& "AND LENGTH( outreach ) =5 and status not in('D','R') and regdate>='2013-11-01' " _
& "GROUP BY farmerid,monitor, outreach,regdate " _
& "Union " _
& "SELECT farmerid ,'I' sharedtype,regdate, individual shared,  ''shared1,  ''shared2,  ''shared3,  ''sgared4, SUM( regland ) AS regland " _
& "FROM  `tbllandreg` " _
& "Where Length(individual) = 5  and status not in('D','R') and regdate>='2013-11-01' " _
& "GROUP BY farmerid,individual,regdate " _
& "Union " _
& "SELECT farmerid ,'S' sharedtype ,regdate,cgmonitor shared,  ''shared1,  ''shared2,  ''shared3,  ''sgared4, SUM( regland ) AS regland " _
& "FROM  `tbllandreg` " _
& "Where Length(cgmonitor) = 5   and status not in('D','R') and regdate>='2013-11-01' " _
& "GROUP BY farmerid, cgmonitor,regdate " _
& "Union " _
& " SELECT farmerid ,'S' sharedtype ,regdate,leadstaff shared,  `SUPPORT1` shared1,  `SUPPORT2` sahred2,  `SUPPORT3` shared3,  `SUPPORT4` shared4, SUM( regland ) AS regland " _
& " FROM  `tbllandreg` " _
& " Where Length(leadstaff) = 5  and status not in('D','R') and regdate>='2013-11-01' " _
& " GROUP BY farmerid ,leadstaff,support1,support2,support3,support4,regdate"
MHVDB.Execute "delete from  tblregistrationrpt"
Set rs = Nothing
rs.Open SQLSTR, MHVDB

Do While rs.EOF <> True
colcnt = 0
    For i = 3 To 7
    If Len(Trim(rs.Fields(i).Value)) = 0 Then Exit For
    If Len(Trim(rs.Fields(i).Value)) = 5 Then
    colcnt = colcnt + 1
    
    End If
    Next

For i = 0 To colcnt - 1
MHVDB.Execute "insert into tblregistrationrpt(staffcode,regland,farmercode,sharedtype,regdate)values( " _
& "'" & Trim(rs.Fields(i + 3).Value) & "','" & Trim(rs.Fields(8).Value) / colcnt & "','" & rs!farmerid & "','" & rs!sharedtype & "','" & Format(rs!regdate, "yyyy-MM-dd") & "'" _
& ")"
Next

rs.MoveNext
Loop

MHVDB.Execute "update tblregistrationrpt set landtype='GRF/SRF' where substring(farmercode,10,1) ='G'"
MHVDB.Execute "update tblregistrationrpt set landtype='Private' where substring(farmercode,10,1) ='F'"
MHVDB.Execute "update tblregistrationrpt set landtype='CF' where  substring(farmercode,10,1)='C'"


MHVDB.Execute "update tblregistrationrpt set sharedtype='Individual' where sharedtype ='I'"
MHVDB.Execute "update tblregistrationrpt set sharedtype='Shared' where sharedtype ='S'"


MHVDB.Execute "update tblregistrationrpt a ,tblfarmer b set farmercode=concat(farmercode,'  ',farmername) where farmercode=idfarmer"

End Sub

Private Sub Command6_Click()
Dim i, j As Integer
Dim dt As Date
Dim mstaff As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim SQLSTR As String
Dim dgtstr As String
Dim farmerstr As String
dt = txtfromdate.Value
mygrid.Clear
For i = 1 To 7
mygrid.TextMatrix(0, i) = WDayName(dt, 0)
dt = dt + 1
Next
i = 1
Set rs = Nothing
SQLSTR = "SELECT staffcode from tblmhvstaff where moniter='1'"
Set rs = Nothing
rs.Open SQLSTR, MHVDB

Do While rs.EOF <> True
mygrid.TextMatrix(i, 0) = rs!staffcode
i = i + 1
rs.MoveNext
Loop

For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 0)) = 0 Then Exit For
Set rs = Nothing
rs.Open "select count(*) cnt from tblfarmer where monitor='" & Trim(mygrid.TextMatrix(i, 0)) & "'", MHVDB
If rs.EOF <> True Then
mygrid.TextMatrix(i, 1) = rs!cnt
End If

dgtstr = ""
Set rsm = Nothing
rsm.Open "select distinct(substring(idfarmer,1,9)) as dgt from tblfarmer where monitor='" & Trim(mygrid.TextMatrix(i, 0)) & "'", MHVDB
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

'rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from phealthhub15_core where staffbarcode='" & mygrid.TextMatrix(i, 0) & "' and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from phealthhub15_core where substring(farmerbarcode,1,9) in " & dgtstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
If rs.EOF <> True Then
mygrid.TextMatrix(i, 2) = rs!cnt
End If

farmerstr = ""
Set rs1 = Nothing
rs1.Open "select distinct farmerbarcode from phealthhub15_core where substring(farmerbarcode,1,9) in " & dgtstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
Do While rs1.EOF <> True
farmerstr = farmerstr + "'" + Trim(rs1!farmerbarcode) + "',"
rs1.MoveNext
Loop


If Len(farmerstr) > 0 Then
farmerstr = "(" + Left(farmerstr, Len(farmerstr) - 1) + ")"
Else
farmerstr = "(" + "'" + A99 & "'" & ")"
End If


Set rs = Nothing
'rs.Open "select count(distinct farmerbarcode) fcnt from storagehub6_core where  staffbarcode='" & mygrid.TextMatrix(i, 0) & "' and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from storagehub6_core where substring(farmerbarcode,1,9) in " & dgtstr & " and farmerbarcode not in " & farmerstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB

If rs.EOF <> True Then

mygrid.TextMatrix(i, 3) = rs!cnt
End If


Set rs = Nothing
'rs.Open "select count(distinct farmerbarcode) fcnt from storagehub6_core where  staffbarcode='" & mygrid.TextMatrix(i, 0) & "' and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB
rs.Open "select count(distinct farmerbarcode) cnt,count(farmerbarcode) fcnt from storagehub6_core where substring(farmerbarcode,1,9) in " & dgtstr & " and substring(end,1,10)>='" & Format(Now - 47, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(Now - 2, "yyyy-MM-dd") & "'", ODKDB

If rs.EOF <> True Then

mygrid.TextMatrix(i, 4) = rs!cnt
End If

Next

'xl.Cells(2, 5) = "'" & Format(Now, "dd/MM/yyyy hh:mm:ss")

Dim tmp As Double
Dim rsfarm As New ADODB.Recordset

For i = 1 To mygrid.Rows - 1
If (Len(mygrid.TextMatrix(i, 0))) = 0 Then Exit Sub
FindsTAFF mygrid.TextMatrix(i, 0)
Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & mygrid.TextMatrix(i, 0) & "' and moniter='1'", MHVDB

If rs.EOF <> True Then
tmp = 0

If Val(mygrid.TextMatrix(i, 1)) = 0 Then
tmp = 0
Else
tmp = (Val(mygrid.TextMatrix(i, 2)) + Val(mygrid.TextMatrix(i, 3))) * 100 / Val(mygrid.TextMatrix(i, 1))

End If
Set rsfarm = Nothing
rsfarm.Open "select * from tblfarmvisit where staffbarcode='" & mygrid.TextMatrix(i, 0) & "'", ODKDB
If rsfarm.EOF <> True Then
ODKDB.Execute "update tblfarmvisit set staffname='" & sTAFF & "',nooffarmers='" & mygrid.TextMatrix(i, 1) & "', " _
& " fieldvisit='" & mygrid.TextMatrix(i, 2) & "', " _
& " storagevisit='" & mygrid.TextMatrix(i, 3) & "' , " _
& " fieldstorage='" & mygrid.TextMatrix(i, 4) & "'  " _
& ",nooffarmersvisted='" & Val(mygrid.TextMatrix(i, 2)) + Val(mygrid.TextMatrix(i, 3)) & "'," _
& "farmersnotvisited='" & Val(mygrid.TextMatrix(i, 1)) - Val(mygrid.TextMatrix(i, 2)) - Val(mygrid.TextMatrix(i, 3)) & "'," _
& " percentage='" & tmp & "'" _
& " where staffbarcode='" & mygrid.TextMatrix(i, 0) & "'"
Else
ODKDB.Execute "insert into tblfarmvisit(staffbarcode,staffname,nooffarmers, " _
& " fieldvisit,storagevisit,nooffarmersvisted,farmersnotvisited,percentage,fieldstorage)values( " _
& "'" & mygrid.TextMatrix(i, 0) & "','" & sTAFF & "' " _
& " ,'" & mygrid.TextMatrix(i, 1) & "','" & mygrid.TextMatrix(i, 2) & "' " _
& ",'" & mygrid.TextMatrix(i, 3) & "','" & Val(mygrid.TextMatrix(i, 2)) + Val(mygrid.TextMatrix(i, 3)) & "'" _
& ",'" & Val(mygrid.TextMatrix(i, 1)) - Val(mygrid.TextMatrix(i, 2)) - Val(mygrid.TextMatrix(i, 3)) & "'" _
& ",'" & tmp & "','" & Val(mygrid.TextMatrix(i, 4)) & "')"
End If
End If
Next


End Sub

Private Sub Command7_Click()
Dim SQLSTR As String
Dim farmerstr As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString

    GetTbl
    SQLSTR = ""
    SQLSTR = "insert into " & Mtblname & " (end,farmercode,fdcode,fs) select n.end,n.farmerbarcode,n.fdcode,'F' " _
            & " from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
            & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
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



     SQLSTR = "insert into " & Mtblname & " (end,farmercode,fs)  select n.end,n.farmerbarcode,'S' from " _
            & "storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
            & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
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
rs1.Open "select * from phealthhub15_core where farmerbarcode='" & rs!farmercode & "'", db
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
rs1.Open "select * from storagehub6_core where farmerbarcode='" & rs!farmercode & "'", db
If rs1.EOF <> True Then
'delete from zero visit
MHVDB.Execute "delete from tblzerovisit where farmercode='" & rs1!farmerbarcode & "'"
End If


rs.MoveNext
Loop
 
End Sub

Private Sub Form_Load()
frmdashboardnursery.Width = 5685


Set db = New ADODB.Connection
Dim tt As String
Dim rs As New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open CnnString
Set rs = Nothing
If rs.State = adStateOpen Then rs.Close

rs.Open "select distinct year(entrydate) Year from tblqmsboxdetail" _
& "  order by year(entrydate),month(entrydate)", db



Set cboyear.RowSource = rs
cboyear.ListField = "Year"
cboyear.BoundColumn = "Year"




Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
If UCase(MUSER) = "ADMIN" Then
rs.Open "select deptid,deptname from tbldept order by deptid", db
Else

 rs.Open "select deptid,deptname from tbldept where remarks like  " & "'%" & UserId & "%'" & "  order by deptid", db

End If
Set cbodept.RowSource = rs
cbodept.ListField = "deptname"
cbodept.BoundColumn = "deptid"

End Sub
Private Sub finance()

Dim SQLSTR As String
Dim i, j As Integer
Dim m As Integer
Dim col As Integer
Dim rs As New ADODB.Recordset
Dim mnthstr As String
    Dim xl As Excel.Application
    Dim var As Variant
    Set xl = CreateObject("excel.Application")
    
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    End If
    
     Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='2'", MHVDB
    If rs.EOF <> True Then
    getSheet 2, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Visible = True
    xl.Sheets("Cost Metrics").Select
    xl.Visible = False
    
    Select Case cbomnth.ListIndex
     Case 0
     
     Case 1
     
     Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
     
    End Select
    
   'tc purchase
    mygrid.Clear
    SQLSTR = "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC,sum(healthyplant) HealthyPlant," _
        & "sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged,sum(oversize) OverSize," _
        & "sum(deadplant) DeadPlant , sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived from tblqmsboxdetail " _
        & "where planttype in(1) and  year(entrydate)='" & cboyear.Text & "' and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') group by year(entrydate),month(entrydate) order by year(entrydate),month(entrydate)" _

For i = 1 To 12
mygrid.TextMatrix(0, i) = MonthName(i, True)
Next
j = 0
m = 1
For i = cbomnth.ListIndex - 1 To cbomnth.ListIndex + 1
 xl.Cells(6, 3 + j) = MonthName(cbomnth.ListIndex - m, True)
 j = j + 1
 m = m - 1
Next

Set rs = Nothing
rs.Open SQLSTR, MHVDB
Do While rs.EOF <> True
mygrid.TextMatrix(1, rs!Month) = rs!healthyplant
rs.MoveNext
Loop



For i = 1 To 12
If Val(mygrid.TextMatrix(1, i)) > 0 Then
For j = 0 To 2
If xl.Cells(6, 3 + j) = mygrid.TextMatrix(0, i) Then
xl.Cells(8, 3 + j) = mygrid.TextMatrix(1, i)
End If
Next
End If

Next
   'nut purchase
   mygrid.Clear
SQLSTR = "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC,sum(healthyplant) HealthyPlant," _
        & "sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged,sum(oversize) OverSize," _
        & "sum(deadplant) DeadPlant , sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived from tblqmsboxdetail " _
        & "where planttype in(3) and  year(entrydate)='" & cboyear.Text & "' and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') group by year(entrydate),month(entrydate) order by year(entrydate),month(entrydate)" _

Set rs = Nothing
rs.Open SQLSTR, MHVDB


For i = 1 To 12
mygrid.TextMatrix(0, i) = MonthName(i, True)
Next


Do While rs.EOF <> True
mygrid.TextMatrix(1, rs!Month) = rs!KATC
rs.MoveNext
Loop



For i = 1 To 12
If Val(mygrid.TextMatrix(1, i)) > 0 Then
For j = 0 To 2
If xl.Cells(6, 3 + j) = mygrid.TextMatrix(0, i) Then
xl.Cells(12, 3 + j) = mygrid.TextMatrix(1, i)
End If
Next
End If

Next


'distribution

mygrid.Clear
Set rs = Nothing
rs.Open "SELECT YEAR( entrydate ) YEAR, MONTH( entrydate ) " _
& " MONTH , SUM( credit - debit ) AS qty " _
& "FROM  `tblqmsplanttransaction` " _
& "Where transactiontype " _
& "IN ('4',  '5') and year(entrydate)='" & cboyear.Text & "' and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "')  " _
& "GROUP BY YEAR( entrydate ) , MONTH( entrydate ) " _
& "ORDER BY YEAR( entrydate ) , MONTH( entrydate )", MHVDB


For i = 1 To 12
mygrid.TextMatrix(0, i) = MonthName(i, True)
Next


Do While rs.EOF <> True
mygrid.TextMatrix(1, rs!Month) = rs!qty
rs.MoveNext
Loop



For i = 1 To 12
If Val(mygrid.TextMatrix(1, i)) > 0 Then
For j = 0 To 2
If xl.Cells(6, 3 + j) = mygrid.TextMatrix(0, i) Then
xl.Cells(24, 3 + j) = mygrid.TextMatrix(1, i)
End If
Next
End If

Next






' out reach

mygrid.Clear
Set rs = Nothing
rs.Open "SELECT YEAR( regdate ) year , MONTH( regdate ) month , SUM( regland ) regland " _
& "FROM  `tbllandreg` " _
& "WHERE STATUS NOT " _
& "IN ('D',  'R') and length(outreach)>'3' and   year(regdate)='" & cboyear.Text & "' and month(regdate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') " _
& "GROUP BY YEAR( regdate ) , MONTH( regdate ) " _
& "ORDER BY YEAR( regdate ) , MONTH( regdate )", MHVDB


For i = 1 To 12
mygrid.TextMatrix(0, i) = MonthName(i, True)
Next


Do While rs.EOF <> True
mygrid.TextMatrix(1, rs!Month) = rs!regland
rs.MoveNext
Loop



For i = 1 To 12
If Val(mygrid.TextMatrix(1, i)) > 0 Then
For j = 0 To 2
If xl.Cells(6, 3 + j) = mygrid.TextMatrix(0, i) Then
xl.Cells(61, 3 + j) = mygrid.TextMatrix(1, i)
End If
Next
End If

Next

' monitor


mygrid.Clear
Set rs = Nothing
rs.Open "SELECT YEAR( regdate ) year , MONTH( regdate ) month , SUM( regland ) regland " _
& "FROM  `tbllandreg` " _
& "WHERE STATUS NOT " _
& "IN ('D',  'R') and length(monitor)>'3' and   year(regdate)='" & cboyear.Text & "' and month(regdate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') " _
& "GROUP BY YEAR( regdate ) , MONTH( regdate ) " _
& "ORDER BY YEAR( regdate ) , MONTH( regdate )", MHVDB


For i = 1 To 12
mygrid.TextMatrix(0, i) = MonthName(i, True)
Next


Do While rs.EOF <> True
mygrid.TextMatrix(1, rs!Month) = rs!regland
rs.MoveNext
Loop



For i = 1 To 12
If Val(mygrid.TextMatrix(1, i)) > 0 Then
For j = 0 To 2
If xl.Cells(6, 3 + j) = mygrid.TextMatrix(0, i) Then
xl.Cells(66, 3 + j) = mygrid.TextMatrix(1, i)
End If
Next
End If

Next


' no of plantx
Dim mdt As Date
Dim dt As Integer
Dim intYear, intMonth, intDay As Integer
mygrid.Clear
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex - 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1





rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('S','N','H','T','C')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB


xl.Cells(19, 3) = rs!stock

Set rs = Nothing
intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('S','N','H','T','C')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(19, 4) = rs!stock
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex + 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('S','N','H','T','C')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(19, 5) = rs!stock


'area monitored
mygrid.Clear

Set rs = Nothing
rs.Open "select year(end) year, month(end) month, sum(regland) regland from tbllandreg a, odk_prodlocal.phealthhub15_core b  where farmerid=farmerbarcode and year(end)='" & cboyear.BoundText & "' and month(end) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') group by year(end),month(end) ", MHVDB
Do While rs.EOF <> True
mygrid.TextMatrix(1, rs!Month) = rs!regland
rs.MoveNext
Loop


xl.Cells(74, 3) = mygrid.TextMatrix(1, cbomnth.ListIndex - 1)
xl.Cells(74, 4) = mygrid.TextMatrix(1, cbomnth.ListIndex)
xl.Cells(74, 5) = mygrid.TextMatrix(1, cbomnth.ListIndex + 1)


    xl.Visible = True


Set xl = Nothing

Screen.MousePointer = vbDefault
End Sub
Private Sub getmnthstr(mycase As Integer)

End Sub
Private Sub nursery()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim mnthstr As String
Dim xl As Excel.Application
Dim var As Variant






    Set xl = CreateObject("excel.Application")
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    End If
    Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='1'", MHVDB
    If rs.EOF <> True Then
    getSheet 1, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Sheets("TC Plants").Select
    xl.Visible = False
    
    
    
    
    
    
    Select Case cbomnth.ListIndex
     Case 0
     
     Case 1
     
     Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
     
    End Select
    
   'Received from KATC
                        
                        receivedfromKATC xl
                        
                        
    ' nuts received
                         xl.Sheets("Nut Plants").Select
                        receivedfromKATCnuts xl
   'inventory tc
                         xl.Sheets("Inventory").Select
                        inventoryTC xl
   ' inventory nuts
                        inventoryNUTS xl
                        
             ' lmt sent to field and back to nursery
              xl.Sheets("Hard Plants").Select
                        lmtsenttofield xl
                        
                        
    'stagging house servival
                        'staginghouseservival xl
    ' sent to ngt
                        'senttoNGT xl

      'utilization xl
                        'utilization xl
       ' net & hoop survival
                        nethoop xl



  xl.Sheets("Summary").Select
                        summary xl



xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault

End Sub
Private Sub lmtsenttofield(xl As Object)
Dim mdt As Date
Dim dt As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim intYear, intMonth, intDay As Integer
mygrid.Clear


Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex - 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1





rs.Open "select year(entrydate) as year,month(entrydate) as Month ,sum(credit) as stock from tblqmsplanttransaction where status<>'C' " _
& " and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('T')) and transactiontype in(4) " _
& " and   year(entrydate)='" & cboyear.BoundText & "' " _
& "and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "', " _
& "'" & cbomnth.ListIndex + 1 & "') group by year(entrydate),month(entrydate) order by year(entrydate)," _
& " month(entrydate) ", MHVDB

 Do While rs.EOF <> True
             
                
                mygrid.TextMatrix(1, rs!Month) = rs!stock
                
                rs.MoveNext
            Loop
            
            
           
               
              
                xl.Cells(19, 12) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)))
                xl.Cells(19, 13) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex)))
                xl.Cells(19, 14) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)))
                
                
                ' lmt send to field
                mygrid.Clear
                Set rs = Nothing
rs.Open "select year(entrydate) as year,month(entrydate) as Month ,sum(credit) as stock from tblqmsplanttransaction where status<>'C' " _
& " and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype not in ('T')) and transactiontype in(4) " _
& " and   year(entrydate)='" & cboyear.BoundText & "' " _
& "and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "', " _
& "'" & cbomnth.ListIndex + 1 & "') group by year(entrydate),month(entrydate) order by year(entrydate)," _
& " month(entrydate) ", MHVDB

 Do While rs.EOF <> True
             
                
                mygrid.TextMatrix(1, rs!Month) = rs!stock
                
                rs.MoveNext
            Loop
            
            
           
               
              
                xl.Cells(17, 12) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)))
                xl.Cells(17, 13) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex)))
                xl.Cells(17, 14) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)))
                
       ' field to lmt
                mygrid.Clear
                 Set rs = Nothing
rs.Open "select year(entrydate) as year,month(entrydate) as Month ,sum(debit) as stock from tblqmsplanttransaction where status<>'C' " _
& " and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype not in ('T')) and transactiontype in(5) " _
& " and   year(entrydate)='" & cboyear.BoundText & "' " _
& "and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "', " _
& "'" & cbomnth.ListIndex + 1 & "') group by year(entrydate),month(entrydate) order by year(entrydate)," _
& " month(entrydate) ", MHVDB

 Do While rs.EOF <> True
             
                
                mygrid.TextMatrix(1, rs!Month) = rs!stock
                
                rs.MoveNext
            Loop
            
            
           
               
              
                xl.Cells(21, 12) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)))
                xl.Cells(21, 13) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex)))
                xl.Cells(21, 14) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)))
                
   ' field to ngt
                mygrid.Clear
                 Set rs = Nothing
rs.Open "select year(entrydate) as year,month(entrydate) as Month ,sum(debit) as stock from tblqmsplanttransaction where status<>'C' " _
& " and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype  in ('T')) and transactiontype in(5) " _
& " and   year(entrydate)='" & cboyear.BoundText & "' " _
& "and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "', " _
& "'" & cbomnth.ListIndex + 1 & "') group by year(entrydate),month(entrydate) order by year(entrydate)," _
& " month(entrydate) ", MHVDB

 Do While rs.EOF <> True
             
                
                mygrid.TextMatrix(1, rs!Month) = rs!stock
                
                rs.MoveNext
            Loop
            
            
           
               
              
                xl.Cells(23, 12) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex - 1)))
                xl.Cells(23, 13) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex)))
                xl.Cells(23, 14) = IIf(Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)) = 0, "", Val(mygrid.TextMatrix(1, cbomnth.ListIndex + 1)))
                
              Dim tt, tt1 As Integer
            Select Case cbomnth.ListIndex + 1
            Case 11
                    tt = 12
                    tt1 = 1
            Case 12
                    tt = 1
                    tt1 = 2
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
                    tt = cbomnth.ListIndex + 2
                    tt1 = cbomnth.ListIndex + 3
            
            End Select
                
          xl.Cells(22, 5) = MonthName(tt, True) & "(Mature)"
           xl.Cells(22, 7) = MonthName(tt1, True) & "(Mature)"
          
                
End Sub
Private Sub nethoop(xl As Object)
Dim mdt As Date
Dim dt As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim intYear, intMonth, intDay As Integer
mygrid.Clear


Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex - 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1





rs.Open "select varietyid,sum(debit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "'", MHVDB
xl.Cells(5, 12) = rs!stock




Set rs = Nothing
rs.Open "select varietyid,sum(debit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "'", MHVDB


Set rs1 = Nothing
rs1.Open "select varietyid,sum(credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(3) and varietyid in(select varietyid from tblqmsplantvariety) and plantbatch in(select plantbatch from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "')", MHVDB
xl.Cells(8, 12) = (rs!stock - IIf(IsNull(rs1!stock), 0, rs1!stock)) / rs!stock


Set rs = Nothing
intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "'", MHVDB

xl.Cells(5, 13) = rs!stock

Set rs = Nothing
rs.Open "select varietyid,sum(debit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(6) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "'", MHVDB


Set rs1 = Nothing
rs1.Open "select varietyid,sum(credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(3) and varietyid in(select varietyid from tblqmsplantvariety) and plantbatch in(select plantbatch from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "')", MHVDB
xl.Cells(8, 13) = (rs!stock - IIf(IsNull(rs1!stock), 0, rs1!stock)) / rs!stock

Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex + 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(6) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "'", MHVDB

xl.Cells(5, 14) = rs!stock


Set rs = Nothing
rs.Open "select varietyid,sum(debit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "'", MHVDB


Set rs1 = Nothing
rs1.Open "select varietyid,sum(credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and transactiontype in(3) and varietyid in(select varietyid from tblqmsplantvariety) and plantbatch in(select plantbatch from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility where housetype in ('N','H')) and transactiontype in(12) and varietyid in(select varietyid from tblqmsplantvariety) and month(entrydate)='" & Month(mdt) & "' and  year(entrydate)='" & Year(mdt) & "')", MHVDB
xl.Cells(8, 14) = (rs!stock - IIf(IsNull(rs1!stock), 0, rs1!stock)) / rs!stock

End Sub
Private Sub utilization(xl As Object)
Dim rs As New ADODB.Recordset
' no of plantx
Dim mdt As Date
Dim dt As Integer
Dim intYear, intMonth, intDay As Integer
mygrid.Clear

' S
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex - 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1





rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('S')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB


xl.Cells(87, 3) = rs!stock / 1000000

Set rs = Nothing
intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('S')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(87, 4) = rs!stock / 1000000
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex + 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('S')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(87, 5) = rs!stock / 1000000

'n,h
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex - 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1





rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB


xl.Cells(92, 3) = rs!stock / 1000000

Set rs = Nothing
intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(92, 4) = rs!stock / 1000000
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex + 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('N','H')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(92, 5) = rs!stock / 1000000


'T

Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex - 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1





rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('T')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB


xl.Cells(97, 3) = rs!stock / 1000000

Set rs = Nothing
intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('T')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(97, 4) = rs!stock / 1000000
Set rs = Nothing

intYear = CInt(cboyear.BoundText)
intMonth = CInt((cbomnth.ListIndex + 1)) + 1
intDay = 1
mdt = DateSerial(intYear, intMonth, intDay) - 1

rs.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and facilityid in (select facilityid from tblqmsfacility " _
& "where housetype in ('T')) and varietyid in(select varietyid from tblqmsplantvariety) and entrydate<='" & Format(mdt, "yyyy-MM-dd") & "'", MHVDB
xl.Cells(97, 5) = rs!stock / 1000000



End Sub
Private Sub senttoNGT(xl As Object)
Dim avgdays As Integer
Dim mdays As Integer
Dim mqty As Double
Dim batchno As Integer
Dim plantcount As Double
Dim CRTOT As Double
Dim mrow As Integer
totdr = 0
totcr = 0
Dim rs1 As New ADODB.Recordset
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsfacility where location='NGT'", MHVDB
If rs.EOF <> True Then
Do While rs.EOF <> True
fid = fid + "'" + rs!facilityid + "',"
rs.MoveNext
Loop
Else

End If
If Len(fid) > 0 Then
fid = "(" + Left(fid, Len(fid) - 1) + ")"
End If
    
   MHVDB.Execute "delete from tblqmstemp"
    SQLSTR = "select entrydate,plantbatch,debit from tblqmsplanttransaction where status='ON' and   year(entrydate)='" & cboyear.Text & "' and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') and  transactiontype='9' and debit>0 and facilityid in " & fid & "  order by plantbatch,entrydate"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
    
         Do Until rs.EOF
         avgdays = 0
         mdays = 0
         mqty = 0
         batchno = rs!plantBatch
         Do While batchno = rs!plantBatch
         Set rs1 = Nothing
         rs1.Open "select * from tblqmsplanttransaction where transactiontype='9' and plantbatch='" & rs!plantBatch & "' and entrydate='" & Format(rs!entrydate, "yyyy-MM-dd") & "' and credit>0 AND facilityid in " & fid & "", MHVDB
         If rs1.EOF <> True Then
         Else
         MHVDB.Execute "insert into tblqmstemp(shipmentfrom,billofladding)values " _
         & "('" & Format(rs!entrydate, "yyyy-MM-yy") & "','" & rs!debit & "')"
       
         End If
         rs.MoveNext
         If rs.EOF Then Exit Do
         Loop
         Loop
    
        End If
mygrid.Clear
SQLSTR = "SELECT year(shipmentfrom) year,month(shipmentfrom) month,sum(billofladding) qty  FROM tblqmstemp group by year(shipmentfrom) ,month(shipmentfrom) order by year(shipmentfrom) ,month(shipmentfrom)"
    Set rs = Nothing
    i = 1
    rs.Open SQLSTR, MHVDB
         If rs.EOF <> True Then
          Do While rs.EOF <> True
            mygrid.TextMatrix(1, i) = rs!qty
                     
          i = i + 1
          rs.MoveNext
          Loop
          
          
          End If
          
                   
                
               
                
                
               
               xl.Cells(83, 32) = Val(mygrid.TextMatrix(1, 1))
                xl.Cells(83, 33) = Val(mygrid.TextMatrix(1, 2))
                 xl.Cells(83, 34) = Val(mygrid.TextMatrix(1, 3))
    End Sub
Private Sub staginghouseservival(xl As Object)
Dim SQLSTR As String
Dim maxdate As Date
Dim mydt As Date
Dim shipmentno As Integer
Dim excel_app As Object
Dim excel_sheet As Object
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rspt As New ADODB.Recordset
totbillladding = 0
totreceivedin = 0
totoversize = 0
totundersize = 0
totweak = 0
toticedamaged = 0
totreceivedex = 0
totdead = 0
tothealthy10 = 0
tothealthyeye = 0
tothardenplants = 0
j = 0
'and plantbatch in(select plantbatch from tblqmsplanttransaction where status='ON' and transactiontype='6')
MHVDB.Execute "delete from tblqmstemp"
    SQLSTR = "select trnid,entrydate,plantbatch,sum(shipmentsize) shipmentsize,sum(healthyplant) healthyplant,sum(weakplant) weakplant," _
           & " sum(undersize) undersize,sum(icedamaged) icedamaged ,sum(oversize) oversize,sum(deadplant) deadplant from tblqmsboxdetail" _
           & " where trnid>=39 and  year(entrydate)='" & cboyear.Text & "' and month(entrydate) in('" & cbomnth.ListIndex - 1 & "','" & cbomnth.ListIndex & "','" & cbomnth.ListIndex + 1 & "') and plantbatch in(select plantbatch from tblqmsplanttransaction where status='ON' and transactiontype='6')" _
           & " group by trnid,entrydate,plantbatch" _
           & " order by trnid,plantbatch"
              

    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_sheet.Name = "Detail"
       Dim sl As Integer
    sl = 1
    i = 1
    'excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "From"
    excel_sheet.Cells(3, 2) = "To"
    excel_sheet.Cells(3, 3) = "Shipment No."
    excel_sheet.Cells(3, 4) = "Plant Batch"
    excel_sheet.Cells(3, 5) = "Facility"
    excel_sheet.Cells(3, 6) = "Bill of Lading"
    excel_sheet.Cells(3, 7) = "Plants Received (including weak, Over Size & Ice Damage plants)"
    excel_sheet.Cells(3, 8) = "Oversize"
    excel_sheet.Cells(3, 9) = "Under Size*"
    excel_sheet.Cells(3, 10) = "Weak received"
    excel_sheet.Cells(3, 11) = "Ice Damaged"
    excel_sheet.Cells(3, 12) = "Plants Received (Excluding Weak, Ice Damaged & Under Size)"
    excel_sheet.Cells(3, 13) = "No. Plants dead within 10 days"
    excel_sheet.Cells(3, 14) = "Healthy Plants (After 10 Days)"
    excel_sheet.Cells(3, 15) = "Eye Assessment"
    excel_sheet.Cells(3, 16) = "Healthy Plants After Eye "
    excel_sheet.Cells(3, 17) = "From "
    excel_sheet.Cells(3, 18) = "To"
    excel_sheet.Cells(3, 19) = "Harden Plants"
    excel_sheet.Cells(3, 20) = "% Survival on the basis of Initial Healthy Plants(Excluding Weak, Ice Damaged & Under Size)"
    excel_sheet.Cells(3, 21) = "% Survival on the basis of Healthy Plants after 10 days in Staging House"
    excel_sheet.Cells(3, 22) = "% Survival on the basis of total plants received (Including weak, ice damaged and undersize)"
    excel_sheet.Cells(3, 23) = "Remarks"
    i = 4
                  
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
     Do Until rs.EOF
     deadin10 = 0
     healthyin10 = 0
     healthyplantsaftereye = 0
     hardenplants = 0
     servival1 = 0
       shipmentno = rs!trnid
'                    Set rspt = Nothing
'                    rspt.Open "select max(entrydate)maxdt,sum(credit) as qty from tblqmsplanttransaction where" _
'                    & " transactiontype='6' and plantbatch='" & rs!plantBatch & "'", MHVDB
'                    If rspt.EOF <> True Then
'                       mydt = rspt!maxdt
'                    End If
'
       
       mydt = rs!entrydate
       
       
    
       Set rspt = Nothing
                    rspt.Open "select entrydate as mindt from tblqmsplantbatchhdr where" _
                    & " trnid='" & rs!trnid & "'", MHVDB
                    If rspt.EOF <> True Then
                        excel_sheet.Cells(i, 1) = "'" & Format(rspt!mindt, "dd/MM/yyyy")
                      
                    End If
                    
                    
                    Set rspt = Nothing
                    rspt.Open "select max(entrydate)maxdt from tblqmsplanttransaction where" _
                    & " transactiontype='2' and plantbatch='" & rs!plantBatch & "'", MHVDB
                    If rspt.EOF <> True Then
                        excel_sheet.Cells(i, 2) = "'" & Format(rspt!maxdt, "dd/MM/yyyy")
                        maxdate = IIf(IsNull(rspt!maxdt), "01/01/1999", rspt!maxdt)
                    End If
                    
                    
       excel_sheet.Cells(i, 3) = shipmentno
       excel_sheet.Cells(i, 3).Font.Bold = True
       Do While shipmentno = rs!trnid
                    ' transplant to bed
                    
                    'no of dead within 10 days
                    Set rspt = Nothing
                    rspt.Open "select sum(credit) as dead from tblqmsplanttransaction where" _
                    & " transactiontype='3' and plantbatch='" & rs!plantBatch & "'" _
                    & " and entrydate>='" & Format(maxdate, "yyyy-MM-dd") & "' and entrydate<='" & Format(maxdate + 10, "yyyy-MM-dd") & "'", MHVDB
                    If rspt.EOF <> True Then
                      excel_sheet.Cells(i, 13) = IIf(rspt!dead = 0, "", rspt!dead)
                      deadin10 = deadin10 + IIf(IsNull(rspt!dead), 0, rspt!dead)
                    End If
                    
             excel_sheet.Cells(i, 4) = rs!plantBatch
             'excel_sheet.Cells(i, 5) = rs!facilityid
             excel_sheet.Cells(i, 6) = rs!shipmentsize
             
             excel_sheet.Cells(i, 8) = IIf(rs!oversize = 0, "", rs!oversize)
             excel_sheet.Cells(i, 9) = IIf(rs!undersize = 0, "", rs!undersize)
             excel_sheet.Cells(i, 10) = IIf(rs!weakplant = 0, "", rs!weakplant)
             excel_sheet.Cells(i, 11) = IIf(rs!icedamaged = 0, "", rs!icedamaged)
             excel_sheet.Cells(i, 12) = rs!healthyplant '- rs!oversize - rs!undersize - rs!weakplant - rs!icedamaged
             excel_sheet.Cells(i, 7) = rs!healthyplant + rs!oversize + rs!undersize + rs!weakplant + rs!icedamaged
             excel_sheet.Cells(i, 14) = Val(excel_sheet.Cells(i, 7)) - rs!oversize - rs!undersize - rs!weakplant - rs!icedamaged - Val(excel_sheet.Cells(i, 13))
             ' eye assstment
              Set rspt = Nothing
             rspt.Open "select sum(credit) as qty from tblqmsplanttransaction where" _
                    & " verificationtype='3' and transactiontype='3' and plantbatch='" & rs!plantBatch & "'", MHVDB
                    If rspt.EOF <> True Then
                      If IIf(IsNull(rspt!qty), 0, rspt!qty) <> 0 Then
                       excel_sheet.Cells(i, 15) = 1 - ((IIf(IsNull(rspt!qty), 0, rspt!qty)) / Val(excel_sheet.Cells(i, 12)))
                       
                       excel_sheet.Cells(i, 15).NumberFormat = "0%"
                       Else
                       excel_sheet.Cells(i, 15) = "100%"
                       End If
                        ' excel_sheet.Cells(i, 15) = excel_sheet.Cells(i, 15) & "%"
                    End If
                    
             excel_sheet.Cells(i, 16) = IIf(Val(excel_sheet.Cells(i, 7)) * Val(excel_sheet.Cells(i, 15)) = 0, "", Val(excel_sheet.Cells(i, 7)) * Val(excel_sheet.Cells(i, 15)))
             excel_sheet.Cells(i, 16).NumberFormat = "0"
             healthyplantsaftereye = healthyplantsaftereye + Val(excel_sheet.Cells(i, 16))
             healthyin10 = healthyin10 + Val(excel_sheet.Cells(i, 14))
             
             ' transplant hard to bags
             Set rspt = Nothing
             rspt.Open "select min(entrydate) mindt, max(entrydate)maxdt,sum(credit) as qty from tblqmsplanttransaction where" _
                    & " transactiontype='6' and plantbatch='" & rs!plantBatch & "'", MHVDB
                    If rspt.EOF <> True Then
                        excel_sheet.Cells(i, 17) = "'" & Format(rspt!mindt, "dd/MM/yyyy")
                        excel_sheet.Cells(i, 18) = "'" & Format(rspt!maxdt, "dd/MM/yyyy")
                        excel_sheet.Cells(i, 19) = IIf((IIf(IsNull(rspt!qty), 0, rspt!qty)) = 0, "", Abs(IIf(IsNull(rspt!qty), 0, rspt!qty)))
                        hardenplants = hardenplants + Val(excel_sheet.Cells(i, 19))
                    End If
                    
                    
            excel_sheet.Cells(i, 20) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 12))
            excel_sheet.Cells(i, 20).NumberFormat = "0%"
            excel_sheet.Cells(i, 21) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 14))
             excel_sheet.Cells(i, 21).NumberFormat = "0%"
            excel_sheet.Cells(i, 22) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 7))
             excel_sheet.Cells(i, 22).NumberFormat = "0%"
             
             totbillladding = totbillladding + Val(excel_sheet.Cells(i, 6))
            totreceivedin = totreceivedin + Val(excel_sheet.Cells(i, 7))
            totoversize = totoversize + Val(excel_sheet.Cells(i, 8))
            totundersize = totundersize + Val(excel_sheet.Cells(i, 9))
            totweak = totweak + Val(excel_sheet.Cells(i, 10))
            toticedamaged = toticedamaged + Val(excel_sheet.Cells(i, 11))
            totreceivedex = totreceivedex + Val(excel_sheet.Cells(i, 12))
            totdead = totdead + Val(excel_sheet.Cells(i, 13))
            tothealthy10 = tothealthy10 + Val(excel_sheet.Cells(i, 14))
            tothealthyeye = tothealthyeye + Val(excel_sheet.Cells(i, 16))
            tothardenplants = tothardenplants + Val(excel_sheet.Cells(i, 19))
             
             
             
             
          i = i + 1
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       'i = i + 1
'        excel_sheet.Cells(i, 3) = "Shipment Total"
'        excel_sheet.Cells(i, 3).Font.Bold = True
       getshipmenttot shipmentno
       excel_sheet.Cells(i, 6) = shipmentsize
       excel_sheet.Cells(i, 7) = IIf(totreceived = 0, "", totreceived)
       excel_sheet.Cells(i, 8) = IIf(oversize = 0, "", oversize)
       excel_sheet.Cells(i, 9) = IIf(undersize = 0, "", undersize)
       excel_sheet.Cells(i, 10) = IIf(weakreceived = 0, "", weakreceived)
       excel_sheet.Cells(i, 11) = IIf(icedamaged = 0, "", icedamaged)
       excel_sheet.Cells(i, 12) = healthyplant
       excel_sheet.Cells(i, 13) = IIf(deadin10 = 0, "", deadin10)
       excel_sheet.Cells(i, 14) = IIf(healthyin10 = 0, "", healthyin10)
       
       
       Set rspt = Nothing
      
             rspt.Open "select sum(credit) as qty from tblqmsplanttransaction where" _
                    & " verificationtype='3' and transactiontype='3' and plantbatch in(select plantbatch from tblqmsplantbatchdetail where trnid='" & shipmentno & "')", MHVDB
                    If rspt.EOF <> True Then
                      If IIf(IsNull(rspt!qty), 0, rspt!qty) <> 0 Then
                       excel_sheet.Cells(i, 15) = 1 - ((IIf(IsNull(rspt!qty), 0, rspt!qty)) / Val(excel_sheet.Cells(i, 12)))
                       
                       excel_sheet.Cells(i, 15).NumberFormat = "0%"
                       Else
                       excel_sheet.Cells(i, 15) = "100%"
                       End If
                        ' excel_sheet.Cells(i, 15) = excel_sheet.Cells(i, 15) & "%"
                    End If
      
      
     
       

       
       
       
                excel_sheet.Cells(i, 16) = IIf(healthyplantsaftereye = 0, "", healthyplantsaftereye)
                excel_sheet.Cells(i, 16).NumberFormat = "0"
                excel_sheet.Cells(i, 19) = IIf(hardenplants = 0, "", hardenplants)
                excel_sheet.Cells(i, 20) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 12))
                excel_sheet.Cells(i, 20).NumberFormat = "0%"
                excel_sheet.Cells(i, 21) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 14))
                excel_sheet.Cells(i, 21).NumberFormat = "0%"
                excel_sheet.Cells(i, 22) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 7))
                excel_sheet.Cells(i, 22).NumberFormat = "0%"
       
       
       
       
       
       MHVDB.Execute "insert into tblqmstemp (shipmentfrom,shipmentno, " _
                   & "billofladding,plantreceivedin,oversize," _
                   & "undersize,weakplant,icedamaged,plantreceivedex,deadin10,healthyafter10," _
                   & "eyeass,healthyeyeass,hardenplants,survivalex,survival10,survivalin" _
                   & ")values " _
                   & "('" & Format(mydt, "yyyy-MM-dd") & "','" & shipmentno & "'," _
                   & "'" & shipmentsize & "','" & totreceived & "','" & oversize & "','" & undersize & "'," _
                   & "'" & weakreceived & "','" & icedamaged & "','" & healthyplant & "','" & deadin10 & "','" & healthyin10 & "', " _
                   & "'" & Val(excel_sheet.Cells(i, 15)) & "','" & healthyplantsaftereye & "', " _
                   & "'" & hardenplants & "','" & Val(excel_sheet.Cells(i, 20)) & "','" & Val(excel_sheet.Cells(i, 21)) & "','" & Val(excel_sheet.Cells(i, 22)) & "')"
        i = i + 1
    Loop
    


    
    
    
    End If











        excel_app.ActiveWorkbook.Close False
        excel_app.DisplayAlerts = False
        excel_app.Quit
      
    Screen.MousePointer = vbDefault
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault
   
    i = 1
    Set rs = Nothing
    mygrid.Clear
    rs.Open "SELECT YEAR( shipmentfrom ) YEAR, MONTH( shipmentfrom ) " _
          & "MONTH , sum(hardenplants)/sum(`plantreceivedex`) ex, " _
          & "sum(hardenplants)/sum(`healthyafter10`) ten, " _
          & "sum(hardenplants)/sum(`plantreceivedin`) inn " _
          & "FROM  `tblqmstemp` " _
          & "GROUP BY YEAR( shipmentfrom ) , MONTH( shipmentfrom ) " _
          & "ORDER BY YEAR( shipmentfrom ) , MONTH( shipmentfrom ) ", MHVDB
          
          
          For i = 1 To 12
          mygrid.TextMatrix(0, i) = MonthName(i, True)
          Next
          
          
          If rs.EOF <> True Then
          Do While rs.EOF <> True
            mygrid.TextMatrix(1, rs!Month) = IIf(IsNull(rs!ex), 0, rs!ex)
            mygrid.TextMatrix(2, rs!Month) = IIf(IsNull(rs!ten), 0, rs!ten)
            mygrid.TextMatrix(3, rs!Month) = IIf(IsNull(rs!inn), 0, rs!inn)
          
          i = i + 1
          rs.MoveNext
          Loop
          End If
          
                   
                
               
                
                
               
                 xl.Cells(14, 3) = mygrid.TextMatrix(1, cbomnth.ListIndex - 1)
                 xl.Cells(14, 4) = mygrid.TextMatrix(1, cbomnth.ListIndex)
                 xl.Cells(14, 5) = mygrid.TextMatrix(1, cbomnth.ListIndex + 1)
                 
                 xl.Cells(15, 3) = mygrid.TextMatrix(2, cbomnth.ListIndex - 1)
                 xl.Cells(15, 4) = mygrid.TextMatrix(2, cbomnth.ListIndex)
                 xl.Cells(15, 5) = mygrid.TextMatrix(2, cbomnth.ListIndex + 1)
                 
                 
'                 xl.Cells(16, 3) = mygrid.TextMatrix(3, 1)
'                 xl.Cells(16, 4) = mygrid.TextMatrix(3, 2)
'                 xl.Cells(16, 5) = mygrid.TextMatrix(3, 3)
              


End Sub
Private Sub summary(xl As Object)

Dim SQLSTR As String
Dim i As Integer
Dim col, row As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
mygrid.Clear



col = 1
row = 1

Set rs1 = Nothing


rs1.Open "select sum(debit-credit) as stock from tblqmsplanttransaction where " _
& " facilityid in(select facilityid from tblqmsfacility where housetype in('N','H','T')) and " _
& " varietyid=12 and status<>'C'", MHVDB
If rs1.EOF <> True Then
    xl.Cells(13, 7) = rs1!stock
Else

xl.Cells(13, 7) = ""
End If
Set rs1 = Nothing
rs1.Open "select sum(debit-credit) as stock from tblqmsplanttransaction where " _
& " facilityid in(select facilityid from tblqmsfacility where housetype in('S')) and " _
& " varietyid=12 and status<>'C'", MHVDB
If rs1.EOF <> True Then
    xl.Cells(14, 7) = rs1!stock
    Else
    xl.Cells(14, 7) = ""
    
    End If
End Sub

Private Sub inventoryTC(xl As Object)
Dim SQLSTR As String
Dim i As Integer
Dim col, row As Integer
Dim rs As New ADODB.Recordset
mygrid.Clear
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
i = 1
rs.Open "select * from tblqmsplantvariety where varietyid not in(5,6,12,13)  order by varietyid", MHVDB
'
Do While rs.EOF <> True
mygrid.TextMatrix(i, 0) = rs!varietyId
i = i + 1
rs.MoveNext
Loop
col = 1
row = 1
Dim r As Integer
Set rs = Nothing
rs.Open "select distinct housetype from tblqmsfacility where housetype in('S','N','H','T') ", MHVDB
Do While rs.EOF <> True
'col 1 for s,2 for net 3 for hoop
Set rs1 = Nothing
row = 1
mygrid.TextMatrix(0, col) = rs!housetype

rs1.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' and varietyid<>'12' and facilityid in(select facilityid from tblqmsfacility where housetype ='" & rs!housetype & "') group by varietyid order by varietyid ", MHVDB

Do While rs1.EOF <> True
For row = 1 To mygrid.Rows - 1
If mygrid.TextMatrix(row, 0) = rs1!varietyId Then

mygrid.TextMatrix(row, col) = rs1!stock
Else

End If

Next

row = row + 1
rs1.MoveNext

Loop

col = col + 1
rs.MoveNext

Loop



            
            
            
            col = 1
            
                       
                For i = 1 To mygrid.Rows - 1
                
                If (Len(mygrid.TextMatrix(i, 0))) = 0 Then Exit For
                col = 1
                FindqmsPlantVariety CInt(mygrid.TextMatrix(i, 0))
               ' xl.Cells(52 + i, 31 + j) = Mygrid.TextMatrix(i, 0)
               For j = 1 To 4
                
                Select Case mygrid.TextMatrix(0, col)
                
                Case "H"
                xl.Cells(5, 17 + j) = "Hoop House" 'Mygrid.TextMatrix(0, col)
                Case "N"
                xl.Cells(5, 17 + j) = "Net House" 'Mygrid.TextMatrix(0, col)
                Case "S"
                xl.Cells(5, 17 + j) = "Staging House" 'Mygrid.TextMatrix(0, col)
                Case "T"
                xl.Cells(5, 17 + j) = "NGT" 'Mygrid.TextMatrix(0, col)
                                 
                
                End Select
                
                
                FindqmsPlantVariety CInt(mygrid.TextMatrix(i, 0))
               xl.Cells(5 + i, 17) = qmsPlantVariety
                xl.Cells(5 + i, 17 + j) = mygrid.TextMatrix(i, j)
                col = col + 1
               Next
                
                Next
             
End Sub
Private Sub receivedfromKATC(xl As Object)
Dim SQLSTR As String
Dim i, j As Integer
Dim col, row As Integer
Dim rs As New ADODB.Recordset
Dim currmonth, premonth, lastmonth, curryear, preyear, lastyear As Integer
  Select Case cbomnth.ListIndex
     Case 0
        curryear = cboyear.BoundText
        preyear = cboyear.BoundText - 1
        lastyear = cboyear.BoundText - 1
        currmonth = 1
        premonth = 12
        lastmonth = 11
     
    Case 1
        curryear = cboyear.BoundText
        preyear = cboyear.BoundText
        lastyear = cboyear.BoundText - 1
        currmonth = 2
        premonth = 1
        lastmonth = 12
     
    Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
        curryear = cboyear.BoundText
        preyear = cboyear.BoundText
        lastyear = cboyear.BoundText
        currmonth = cbomnth.ListIndex + 1
        premonth = cbomnth.ListIndex
        lastmonth = cbomnth.ListIndex - 1
     End Select
    xl.Cells(2, 17) = MonthName(lastmonth, True)
    xl.Cells(2, 18) = MonthName(premonth, True)
    xl.Cells(2, 19) = MonthName(currmonth, True)


mygrid.Clear
   SQLSTR = "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC, " _
            & " sum(healthyplant) HealthyPlant, " _
            & "sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged, " _
            & " sum(oversize) OverSize, " _
            & " sum(deadplant) DeadPlant , " _
            & " sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived " _
            & " from tblqmsboxdetail " _
            & " where plantvariety<>'12' and year(entrydate)='" & lastyear & "' and " _
            & " month(entrydate) in('" & lastmonth & "') group by year(entrydate),month(entrydate) " _
            & " union " _
            & "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC, " _
            & " sum(healthyplant) HealthyPlant, " _
            & "sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged, " _
            & " sum(oversize) OverSize, " _
            & " sum(deadplant) DeadPlant , " _
            & " sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived " _
            & " from tblqmsboxdetail " _
            & " where plantvariety<>'12' and year(entrydate)='" & preyear & "' and " _
            & " month(entrydate) in('" & premonth & "') group by year(entrydate),month(entrydate) "
      
            

            
   SQLSTR = SQLSTR & " union " & "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC, " _
            & " sum(healthyplant) HealthyPlant, " _
            & "sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged, " _
            & " sum(oversize) OverSize, " _
            & " sum(deadplant) DeadPlant , " _
            & " sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived " _
            & " from tblqmsboxdetail " _
            & " where plantvariety<>'12' and year(entrydate)='" & curryear & "' and " _
            & " month(entrydate) in('" & currmonth & "') group by year(entrydate),month(entrydate) "
            
            
            
            
            
            
            
            
            
            '& order by year(entrydate),month(entrydate) " _

            
            
            
            
            
            
            
            Set rs = Nothing
            rs.Open SQLSTR, MHVDB
            
            For i = 1 To 12
             mygrid.TextMatrix(0, i) = MonthName(i, True)
            Next
            
            xl.Cells(3, 18) = MonthName(lastmonth, True)
            xl.Cells(3, 19) = MonthName(premonth, True)
            xl.Cells(3, 20) = MonthName(currmonth, True)
  
             Do While rs.EOF <> True
             
                
                mygrid.TextMatrix(1, rs!Month) = rs!KATC
                mygrid.TextMatrix(2, rs!Month) = rs!healthyplant
                rs.MoveNext
            Loop
            
            
           
               
              
                xl.Cells(4, 18) = IIf(Val(mygrid.TextMatrix(1, lastmonth)) = 0, "", Val(mygrid.TextMatrix(1, lastmonth)))
                xl.Cells(4, 19) = IIf(Val(mygrid.TextMatrix(1, premonth)) = 0, "", Val(mygrid.TextMatrix(1, premonth)))
                xl.Cells(4, 20) = IIf(Val(mygrid.TextMatrix(1, currmonth)) = 0, "", Val(mygrid.TextMatrix(1, currmonth)))
                
                
                
                xl.Cells(5, 18) = IIf(Val(mygrid.TextMatrix(2, lastmonth)) = 0, "", Val(mygrid.TextMatrix(2, lastmonth)))
                xl.Cells(5, 19) = IIf(Val(mygrid.TextMatrix(2, premonth)) = 0, "", Val(mygrid.TextMatrix(2, premonth)))
                xl.Cells(5, 20) = IIf(Val(mygrid.TextMatrix(2, currmonth)) = 0, "", Val(mygrid.TextMatrix(2, currmonth)))
                         
                xl.Cells(6, 18) = (Val(mygrid.TextMatrix(1, lastmonth)) - Val(mygrid.TextMatrix(2, lastmonth)))
                xl.Cells(6, 19) = (Val(mygrid.TextMatrix(1, premonth)) - Val(mygrid.TextMatrix(2, premonth)))
                xl.Cells(6, 20) = (Val(mygrid.TextMatrix(1, currmonth)) - Val(mygrid.TextMatrix(2, currmonth)))
                         
                If Val(mygrid.TextMatrix(1, lastmonth)) > 0 Then
                 
                 xl.Cells(7, 18) = (Val(mygrid.TextMatrix(1, lastmonth)) - Val(mygrid.TextMatrix(2, lastmonth))) / Val(mygrid.TextMatrix(1, lastmonth))
                
                Else
                xl.Cells(7, 18) = ""
                End If
                
                 If Val(mygrid.TextMatrix(1, premonth)) > 0 Then
                  
                xl.Cells(7, 19) = (Val(mygrid.TextMatrix(1, premonth)) - Val(mygrid.TextMatrix(2, premonth))) / Val(mygrid.TextMatrix(1, premonth))
                
                Else
                xl.Cells(7, 19) = ""
                End If
                
                 If Val(mygrid.TextMatrix(1, currmonth)) > 0 Then
                  
                  xl.Cells(7, 20) = (Val(mygrid.TextMatrix(1, currmonth)) - Val(mygrid.TextMatrix(2, currmonth))) / Val(mygrid.TextMatrix(1, currmonth))
                
                Else
                xl.Cells(7, 20) = ""
                End If
               
   
End Sub

Private Sub mygridKATC_Click()
mygridKATC.Editable = flexEDKbdMouse
End Sub

Private Sub mygridKATC_LeaveCell()
If Len(cboyear.Text) = 0 Then Exit Sub
'pulldata
'acceptedforstagging
'KATCdefects
End Sub

Private Sub Text1_DblClick()
frmdashboardnursery.Width = 14670
End Sub
Private Sub getshipmenttot(shipmentno As Integer)
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
shipmentsize = 0
healthyplant = 0
oversize = 0
undersize = 0
weakreceived = 0
icedamaged = 0
totreceived = 0
 SQLSTR = "select trnid,sum(shipmentsize) shipmentsize,sum(healthyplant) healthyplant,sum(weakplant) weakplant," _
           & " sum(undersize) undersize,sum(icedamaged) icedamaged ,sum(oversize) oversize,sum(deadplant) deadplant from tblqmsboxdetail" _
           & " where trnid='" & shipmentno & "'" _
           & " group by trnid"
 Set rs = Nothing
 rs.Open SQLSTR, MHVDB
 If rs.EOF <> True Then
    shipmentsize = rs!shipmentsize
    healthyplant = rs!healthyplant
    oversize = rs!oversize
    undersize = rs!undersize
    weakreceived = rs!weakplant
    icedamaged = rs!icedamaged
    totreceived = rs!healthyplant + rs!oversize + rs!undersize + rs!weakplant + rs!icedamaged
 End If

End Sub

Private Sub receivedfromKATCnuts(xl As Object)
Dim SQLSTR As String
Dim i, j As Integer
Dim col, row As Integer
Dim rs As New ADODB.Recordset

Dim currmonth, premonth, lastmonth, curryear, preyear, lastyear As Integer
  Select Case cbomnth.ListIndex
     Case 0
        curryear = cboyear.BoundText
        preyear = cboyear.BoundText - 1
        lastyear = cboyear.BoundText - 1
        currmonth = 1
        premonth = 12
        lastmonth = 11
     
    Case 1
        curryear = cboyear.BoundText
        preyear = cboyear.BoundText
        lastyear = cboyear.BoundText - 1
        currmonth = 2
        premonth = 1
        lastmonth = 12
     
    Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
        curryear = cboyear.BoundText
        preyear = cboyear.BoundText
        lastyear = cboyear.BoundText
        currmonth = cbomnth.ListIndex + 1
        premonth = cbomnth.ListIndex
        lastmonth = cbomnth.ListIndex - 1
     End Select




mygrid.Clear
   SQLSTR = "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC,sum(healthyplant) " _
   & " HealthyPlant, sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged, " _
   & " sum(oversize) OverSize,sum(deadplant) DeadPlant , " _
   & " sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived from tblqmsboxdetail " _
  & "where plantvariety='12' and year(entrydate)='" & lastyear & "' and month(entrydate) " _
  & " in('" & lastmonth & "') " _
  & " group by year(entrydate),month(entrydate) " _
  & " union " _
  & "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC,sum(healthyplant) " _
   & " HealthyPlant, sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged, " _
   & " sum(oversize) OverSize,sum(deadplant) DeadPlant , " _
   & " sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived from tblqmsboxdetail " _
  & "where plantvariety='12' and year(entrydate)='" & preyear & "' and month(entrydate) " _
  & " in('" & premonth & "') " _
  & " group by year(entrydate),month(entrydate) " _
  & " union " _
  & "select year(entrydate) Year,month(entrydate) Month,sum(shipmentsize)KATC,sum(healthyplant) " _
   & " HealthyPlant, sum(undersize) UnderSize,sum(weakplant) WeakPlant,sum(icedamaged) IceDamaged, " _
   & " sum(oversize) OverSize,sum(deadplant) DeadPlant , " _
   & " sum(healthyplant+undersize+weakplant+icedamaged+oversize+deadplant)TotalReceived from tblqmsboxdetail " _
  & "where plantvariety='12' and year(entrydate)='" & curryear & "' and month(entrydate) " _
  & " in('" & currmonth & "') " _
  & " group by year(entrydate),month(entrydate) "

  
  
  
  
  
  
  
            
            
            Set rs = Nothing
            rs.Open SQLSTR, MHVDB
            
             
             Do While rs.EOF <> True
             
                
                mygrid.TextMatrix(1, rs!Month) = rs!KATC
                mygrid.TextMatrix(2, rs!Month) = rs!healthyplant
                rs.MoveNext
            Loop
            
            
           
               
              
                xl.Cells(6, 12) = IIf(Val(mygrid.TextMatrix(1, lastmonth)) = 0, "", Val(mygrid.TextMatrix(1, lastmonth)))
                xl.Cells(6, 13) = IIf(Val(mygrid.TextMatrix(1, premonth)) = 0, "", Val(mygrid.TextMatrix(1, premonth)))
                xl.Cells(6, 14) = IIf(Val(mygrid.TextMatrix(1, currmonth)) = 0, "", Val(mygrid.TextMatrix(1, currmonth)))
                
          
   
End Sub
Private Sub inventoryNUTS(xl As Object)
Dim SQLSTR As String
Dim i As Integer
Dim col, row As Integer
Dim rs As New ADODB.Recordset
mygrid.Clear
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
i = 1
rs.Open "select * from tblqmsplantvariety where varietyid in(12)  order by varietyid", MHVDB
'
Do While rs.EOF <> True
mygrid.TextMatrix(i, 0) = rs!varietyId
i = i + 1
rs.MoveNext
Loop
col = 1
row = 1
Dim r As Integer
Set rs = Nothing
rs.Open "select distinct housetype from tblqmsfacility where housetype in('C','S','N','H','T') ", MHVDB
Do While rs.EOF <> True
'col 1 for s,2 for net 3 for hoop
Set rs1 = Nothing
row = 1
mygrid.TextMatrix(0, col) = rs!housetype



rs1.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where status<>'C' " _
& " and varietyid='12' and facilityid in(select facilityid from tblqmsfacility where " _
& " housetype ='" & rs!housetype & "') group by varietyid order by varietyid ", MHVDB


Do While rs1.EOF <> True
For row = 1 To mygrid.Rows - 1
If mygrid.TextMatrix(row, 0) = rs1!varietyId Then

mygrid.TextMatrix(row, col) = rs1!stock
Else

End If

Next

row = row + 1
rs1.MoveNext

Loop

col = col + 1
rs.MoveNext

Loop



            
            
            
            col = 1
            
                       
                For i = 1 To mygrid.Rows - 1
                
                If (Len(mygrid.TextMatrix(i, 0))) = 0 Then Exit For
                col = 1
                Findqmsnuttype mygrid.TextMatrix(i, 0)
               ' xl.Cells(52 + i, 31 + j) = Mygrid.TextMatrix(i, 0)
               For j = 1 To 5
                
                Select Case mygrid.TextMatrix(0, col)
                Case "C"
                xl.Cells(17, 17 + j) = "Cold House" 'Mygrid.TextMatrix(0, col)
                xl.Cells(18, 17 + j) = mygrid.TextMatrix(i, j)
                               
                Case "H"
                xl.Cells(17, 17 + j) = "Hoop House" 'Mygrid.TextMatrix(0, col)
                xl.Cells(20, 17 + j) = mygrid.TextMatrix(i, j)
                Case "N"
                xl.Cells(17, 17 + j) = "Net House" 'Mygrid.TextMatrix(0, col)
                xl.Cells(20, 17 + j) = mygrid.TextMatrix(i, j)
                Case "S"
                xl.Cells(17, 17 + j) = "Staging House" 'Mygrid.TextMatrix(0, col)
                 xl.Cells(19, 17 + j) = mygrid.TextMatrix(i, j)
                Case "T"
                xl.Cells(17, 17 + j) = "NGT" 'Mygrid.TextMatrix(0, col)
                xl.Cells(20, 17 + j) = mygrid.TextMatrix(i, j)
                
                End Select
                
                col = col + 1
              
               Next
                
                Next
             
End Sub
