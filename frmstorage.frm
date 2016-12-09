VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstorage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STORAGE REPORT"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9195
   Icon            =   "frmstorage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CHKMOREOPTION 
      Caption         =   "MORE OPTION"
      Height          =   195
      Left            =   3840
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "MORE OPTION"
      Height          =   4695
      Left            =   5640
      TabIndex        =   18
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton OPTLEAFPEST 
         Caption         =   "LEAF PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton OPTSTEMPEST 
         Caption         =   "STEM PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton optrootpest 
         Caption         =   "ROOT PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TXTVALUE 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   3720
         Width           =   1095
      End
      Begin VB.OptionButton optdead 
         Caption         =   "DEAD"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton optmoist 
         Caption         =   "MOISTURE"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERCENTAGE VALUE GREATER THEN"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   11
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
         ItemData        =   "frmstorage.frx":076A
         Left            =   1080
         List            =   "frmstorage.frx":076C
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   74842113
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   74842113
         CurrentDate     =   41362
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   945
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
      Left            =   2040
      Picture         =   "frmstorage.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
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
      Left            =   360
      Picture         =   "frmstorage.frx":1438
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame 
      Caption         =   "RECORD SELECTION"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
      Begin VB.OptionButton OPTALLVISIT 
         Caption         =   "ALL VISIT"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox TXTRECORDNO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   350
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OPTTOPN 
         Caption         =   "LAST VISIT"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "VISIT"
         Height          =   195
         Left            =   3435
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DATE SELECTION"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton OPTSEL 
         Caption         =   "SELECTIVE"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CheckBox CHKDGT 
      Caption         =   "SUMMARY"
      Height          =   195
      Left            =   3840
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   4800
   End
End
Attribute VB_Name = "frmstorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub storagedaily()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsch As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mchk = True
db.Open OdkCnnString
                      

GetTbl

SQLSTR = ""


         SQLSTR = "select _uri,start,tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,totaltrees,other,gmoisture,pmoisture,gmoisture+pmoisture as totaltally," _
         & "dtrees,ndtrees,wlogged,pdamage,adamage,monitorcomments from storagehub6_core "
         SQLSTR = SQLSTR & "where status<>'BAD' and substring(start,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by staffbarcode"
               
                
                
                
                
         
 
  db.Execute SQLSTR


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
    'excel_app.Visible = False
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("start date")
    excel_sheet.Cells(3, 3) = ProperCase("tdate")
    excel_sheet.Cells(3, 4) = ProperCase("end date")
    excel_sheet.Cells(3, 5) = ProperCase("STAFF CODE - NAME")
    excel_sheet.Cells(3, 6) = ProperCase("DZONGKHAG")
    excel_sheet.Cells(3, 7) = ProperCase("GEWOG")
    excel_sheet.Cells(3, 8) = ProperCase("TSHOWOG")
    excel_sheet.Cells(3, 9) = ProperCase("Farmer code - name")
    excel_sheet.Cells(3, 10) = ProperCase("storage condition")
    excel_sheet.Cells(3, 11) = ProperCase("storage problem")
    excel_sheet.Cells(3, 12) = ProperCase("action recommended")
    excel_sheet.Cells(3, 13) = ProperCase("Total Trees Distributed - Planted List")
    excel_sheet.Cells(3, 14) = ProperCase("Total Trees")
    excel_sheet.Cells(3, 15) = ProperCase("Good Moisture")
    excel_sheet.Cells(3, 16) = ProperCase("Poor Moisture")
    excel_sheet.Cells(3, 17) = ProperCase("Total Mositure Tally")
    excel_sheet.Cells(3, 18) = ProperCase("Dead Missing")
    'excel_sheet.Cells(3, 17) = ProperCase("Slow Growing")
    'excel_sheet.Cells(3, 18) = ProperCase("Dormant")
    'excel_sheet.Cells(3, 19) = ProperCase("Active Growing")
    'excel_sheet.Cells(3, 20) = ProperCase("Shock")
    excel_sheet.Cells(3, 19) = ProperCase("Nutrient Deficient")
    excel_sheet.Cells(3, 20) = ProperCase("Water Logg")
    excel_sheet.Cells(3, 21) = ProperCase("pest damage")
    'excel_sheet.Cells(3, 20) = ProperCase("Active Pest")
    'excel_sheet.Cells(3, 21) = ProperCase("Stem Pest")
    'excel_sheet.Cells(3, 22) = ProperCase("Root Pest")
    excel_sheet.Cells(3, 22) = ProperCase("Animal Damage")
    excel_sheet.Cells(3, 23) = ProperCase("comments")
   i = 4
  Set rs = Nothing
  SLNO = 1
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
mchk = True
excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!Start
excel_sheet.Cells(i, 3) = "'" & rs!tdate
excel_sheet.Cells(i, 4) = "'" & rs!End
FindsTAFF rs!staffbarcode
excel_sheet.Cells(i, 5) = rs!staffbarcode & " " & sTAFF

FindDZ Mid(rs!farmerbarcode, 1, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmerbarcode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3)
excel_sheet.Cells(i, 7) = Mid(rs!farmerbarcode, 4, 3) & " " & GEname
FindTs Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3), Mid(rs!farmerbarcode, 7, 3)
excel_sheet.Cells(i, 8) = Mid(rs!farmerbarcode, 7, 3) & " " & TsName

FindFA rs!farmerbarcode, "F"
excel_sheet.Cells(i, 9) = rs!farmerbarcode & " " & FAName




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
excel_sheet.Cells(i, 10) = actstring
End If
Else

excel_sheet.Cells(i, 10) = "" 'IIf(IsNull(RS!id), "", RS!id)
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
excel_sheet.Cells(i, 11) = actstring
End If
Else

excel_sheet.Cells(i, 11) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If


Set rs1 = Nothing
rs1.Open "select * from storagehub6_draction where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tblstoragechoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
If UCase(rsch!Name) = UCase("action7") Then
actstring = rs!OTHER
Else
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If
End If
rs1.MoveNext
Loop

If Len(actstring) > 0 Then
 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.Cells(i, 12) = actstring
End If
Else

excel_sheet.Cells(i, 12) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If

Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as nooftrees from tblplanted where farmercode='" & rs!farmerbarcode & "' group by farmercode ", MHVDB
If rs1.EOF <> True Then
excel_sheet.Cells(i, 13) = rs1!nooftrees

Else

excel_sheet.Cells(i, 13) = ""
End If

excel_sheet.Cells(i, 14) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!gmoisture), "", rs!gmoisture)
excel_sheet.Cells(i, 16) = IIf(IsNull(rs!pmoisture), "", rs!pmoisture)
excel_sheet.Cells(i, 17) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.Cells(i, 18) = IIf(IsNull(rs!dtrees), "", rs!dtrees)
'excel_sheet.Cells(i, 17) = "" 'IIf(IsNull(rs!slowgrowing), "", rs!slowgrowing)
'excel_sheet.Cells(i, 18) = "" 'IIf(IsNull(rs!dor), "", rs!dor)
'excel_sheet.Cells(i, 19) = "" 'IIf(IsNull(rs!activegrowing), "", rs!activegrowing)
'excel_sheet.Cells(i, 20) = "" 'IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.Cells(i, 19) = IIf(IsNull(rs!ndtrees), "", rs!ndtrees)
excel_sheet.Cells(i, 20) = IIf(IsNull(rs!wlogged), "", rs!wlogged)
excel_sheet.Cells(i, 21) = IIf(IsNull(rs!pdamage), "", rs!pdamage)
'excel_sheet.Cells(i, 20) = "" 'IIf(IsNull(rs!activepest), "", rs!activepest)
'excel_sheet.Cells(i, 21) = "" 'IIf(IsNull(rs!stempest), "", rs!stempest)
'excel_sheet.Cells(i, 22) = "" 'IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.Cells(i, 22) = IIf(IsNull(rs!adamage), "", rs!adamage)
excel_sheet.Cells(i, 23) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)




SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up




'   excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(i, 15)).Select
'    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
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


'With excel_sheet
'    .LeftMargin = Application.InchesToPoints(0.5)
'    .RightMargin = Application.InchesToPoints(0.75)
'    .TopMargin = Application.InchesToPoints(1.5)
'    .BottomMargin = Application.InchesToPoints(1)
'    .HeaderMargin = Application.InchesToPoints(0.5)
'    .FooterMargin = Application.InchesToPoints(0.5)
'End With


 
'MsgBox CountOfBreaks

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
db.Close
Exit Sub
err:
MsgBox err.Description
err.Clear
End Sub

Private Sub allstorage()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
mchk = True
db.Open OdkCnnString
                      

GetTbl

SQLSTR = ""


If OPTTOPN.Value = False Then
         SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
         & "waterlog,leafpest,animaldamage,monitorcomments) select start,tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,ndtrees,wlogged,pdamage,adamage,monitorcomments from storagehub6_core "
         
                If optall.Value = True Then
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD'"
                SQLSTR = SQLSTR & "where   status<>'BAD'"
                Else
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                 SQLSTR = SQLSTR & "where status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                End If
                
                
                
                
                
         
Else



      SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
         & "waterlog,leafpest,animaldamage,monitorcomments)  select end,end,end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,ndtrees,wlogged,pdamage,adamage,monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
         
            
         
         
         
         
End If
         
  
  db.Execute SQLSTR


'On Error Resume Next





SQLSTR = ""
If CHKMOREOPTION.Value = 0 Then

   SQLSTR = "select * from " & Mtblname & ""
 

Else
    If optmoist.Value = True Then
        SQLSTR = "select * from " & Mtblname & " where (poormoisture/totaltally)*100>'" & Val(txtvalue.Text) & "'"
    ElseIf optrootpest.Value = True Then
        SQLSTR = "select * from " & Mtblname & " where (rootpest/totaltrees)*100>'" & Val(txtvalue.Text) & "'"
    ElseIf OPTSTEMPEST.Value = True Then
        SQLSTR = "select * from " & Mtblname & " where (stempest/totaltrees)*100>'" & Val(txtvalue.Text) & "'"
    ElseIf OPTLEAFPEST.Value = True Then
        SQLSTR = "select * from " & Mtblname & " where (leafpest/totaltrees)*100>'" & Val(txtvalue.Text) & "'"
    Else
        SQLSTR = "select * from " & Mtblname & " where (deadmissing/totaltrees)*100>'" & Val(txtvalue.Text) & "'"
    End If

End If


SQLSTR = SQLSTR & " order by end,farmercode,fdcode"















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
    'excel_app.Visible = False
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
        excel_sheet.Cells(3, 2) = ProperCase("DATE" & "(END)")
        excel_sheet.Cells(3, 3) = ProperCase("STAFF CODE-NAME")
    excel_sheet.Cells(3, 4) = ProperCase("DZONGKHAG")
    excel_sheet.Cells(3, 5) = ProperCase("GEWOG")
    excel_sheet.Cells(3, 6) = ProperCase("TSHOWOG")
    excel_sheet.Cells(3, 7) = ProperCase("Farmer ID")
    excel_sheet.Cells(3, 8) = ProperCase("Total Distributed")
    excel_sheet.Cells(3, 9) = ProperCase("Field ID")
    excel_sheet.Cells(3, 10) = ProperCase("Total Trees Distributed - Planted List")
    excel_sheet.Cells(3, 11) = ProperCase("Total Trees")
    excel_sheet.Cells(3, 12) = ProperCase("Good Moisture")
    excel_sheet.Cells(3, 13) = ProperCase("Poor Moisture")
    excel_sheet.Cells(3, 14) = ProperCase("Total Mositure Tally")
    excel_sheet.Cells(3, 15) = ProperCase("Dead Missing")
    excel_sheet.Cells(3, 16) = ProperCase("Slow Growing")
    excel_sheet.Cells(3, 17) = ProperCase("Dormant")
    excel_sheet.Cells(3, 18) = ProperCase("Active Growing")
    excel_sheet.Cells(3, 19) = ProperCase("Shock")
    excel_sheet.Cells(3, 20) = ProperCase("Nutrient Deficient")
    excel_sheet.Cells(3, 21) = ProperCase("Water Logg")
    excel_sheet.Cells(3, 22) = ProperCase("pest damage")
    excel_sheet.Cells(3, 23) = ProperCase("Active Pest")
    excel_sheet.Cells(3, 24) = ProperCase("Stem Pest")
    excel_sheet.Cells(3, 25) = ProperCase("Root Pest")
    excel_sheet.Cells(3, 26) = ProperCase("Animal Damage")
    excel_sheet.Cells(3, 27) = ProperCase("comments")
   i = 4
  Set rs = Nothing
  SLNO = 1
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
mchk = True
excel_sheet.Cells(i, 1) = SLNO

excel_sheet.Cells(i, 2) = "'" & rs!End
FindsTAFF rs!id
excel_sheet.Cells(i, 3) = rs!id & " " & sTAFF

FindDZ Mid(rs!farmercode, 1, 3)
excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
excel_sheet.Cells(i, 5) = Mid(rs!farmercode, 4, 3) & " " & GEname
FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmercode, 7, 3) & " " & TsName

FindFA excel_sheet.Cells(i, 7), "F"
excel_sheet.Cells(i, 7) = rs!farmercode & " " & FAName




excel_sheet.Cells(i, 8) = ""
excel_sheet.Cells(i, 9) = ""
Set rs1 = Nothing
rs1.Open "select sum(nooftrees) as nooftrees from tblplanted where farmercode='" & rs!farmercode & "' group by farmercode ", MHVDB
If rs1.EOF <> True Then
excel_sheet.Cells(i, 10) = rs1!nooftrees

Else

excel_sheet.Cells(i, 10) = ""
End If
excel_sheet.Cells(i, 11) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.Cells(i, 12) = IIf(IsNull(rs!goodmoisture), "", rs!goodmoisture)
excel_sheet.Cells(i, 13) = IIf(IsNull(rs!poormoisture), "", rs!poormoisture)
excel_sheet.Cells(i, 14) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!tree_count_deadmissing), "", rs!tree_count_deadmissing)
excel_sheet.Cells(i, 16) = "" 'IIf(IsNull(rs!slowgrowing), "", rs!slowgrowing)
excel_sheet.Cells(i, 17) = "" 'IIf(IsNull(rs!dor), "", rs!dor)
excel_sheet.Cells(i, 18) = "" 'IIf(IsNull(rs!activegrowing), "", rs!activegrowing)
excel_sheet.Cells(i, 19) = "" 'IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.Cells(i, 20) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.Cells(i, 21) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.Cells(i, 22) = IIf(IsNull(rs!leafpest), "", rs!leafpest)
excel_sheet.Cells(i, 23) = "" 'IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.Cells(i, 24) = "" 'IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.Cells(i, 25) = "" 'IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.Cells(i, 26) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.Cells(i, 27) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)




SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up




'   excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(i, 15)).Select
'    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
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


'With excel_sheet
'    .LeftMargin = Application.InchesToPoints(0.5)
'    .RightMargin = Application.InchesToPoints(0.75)
'    .TopMargin = Application.InchesToPoints(1.5)
'    .BottomMargin = Application.InchesToPoints(1)
'    .HeaderMargin = Application.InchesToPoints(0.5)
'    .FooterMargin = Application.InchesToPoints(0.5)
'End With


 
'MsgBox CountOfBreaks

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
db.Close
Exit Sub
err:
MsgBox err.Description
err.Clear


End Sub


Private Sub CHKMOREOPTION_Click()
If CHKMOREOPTION.Value = 1 Then
frmstorage.Width = 9435
OPTTOPN.Enabled = False
Else
frmstorage.Width = 5595
OPTTOPN.Enabled = True
End If
optmoist.Value = True
txtvalue.Text = 30

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
'storagemonthlyvisit
allstorage
'storagedaily
End Sub
Private Sub storagemonthlyvisit() 'receipient_id As Integer, nextemaildate As Date, frequency As Integer)
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
db.Execute " insert into tempfarmernotinfield(end,farmercode,staffbarcode)" _
           & "select distinct '' as end, farmerbarcode,staffbarcode from storagehub6_core " _
           & "where farmerbarcode not in (select farmerbarcode from phealthhub15_core) group by farmerbarcode"
           
SQLSTR = " delete from tempfarmernotinfield  where farmercode  in(" _
& "select farmercode from mhv.tblplanted as a , mhv.tblfarmer as b  " _
& "where farmercode=idfarmer)"

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
      excel_sheet.Cells(jrow, 1) = ProperCase("FARMER")
      
     
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
    excel_sheet.Cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
    excel_sheet.Cells(jrow, 2) = ProperCase("LAST VISITED")
       
    K = 2
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(jrow, K) = ProperCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.Cells(jrow, jCol) = ProperCase("Total")
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
          i = rs!procmonth + 3 - Month(txtfrmdate)
          
          j = IIf(rs!jval = "", 0, rs!jval)
          jtot = jtot + j
         
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, 2) = "'" & rs!End
         
          
          'fdcnt = fdcnt + 1
          If rs!farmerbarcode = "D03G06T02F0006" Then
          MsgBox "qwf"
          End If
          excel_sheet.Cells(jrow, i) = IIf(Val(j) = 0, "", Val(j))
          

                        
                         
          
         pyear = rs!id
          rs.MoveNext
         
          If rs.EOF Then Exit Do
          'jrow = jrow + 1
       Loop
       
     
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
       If Val(jtot) = 0 Then
                            excel_sheet.Range(excel_sheet.Cells(jrow, 1), _
                             excel_sheet.Cells(jrow, 1)).Select
                             excel_app.Selection.Interior.ColorIndex = 15
                             End If
       'rs.MoveNext
       'jtot = 0
    Loop
   jtot = 0
    'excel_sheet.Cells(jrow + 1, 3) = fdcnt
    excel_sheet.Cells(jrow + 1, 1) = ProperCase("Total")
    For i = 3 To jCol - 1
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
SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from  storagehub6_core where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"


Set rs = Nothing
rs.Open SQLSTR, OdkCnnString
jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs!id
    'excel_sheet.Cells(2, 1) = "MONTHLY ACTIVITY OF MONITOR " & rs!id & " " & sTAFF
    excel_sheet.Cells(3, 1) = ProperCase("MONITOR")
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.Cells(3, jCol) = ProperCase("Total")
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
    excel_sheet.Cells(jrow + 1, 1) = ProperCase("Total")
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


Private Sub Form_Load()
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
frmstorage.Width = 5595
Mname = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")

End Sub

Private Sub OPTALL_Click()
Frame1.Enabled = False
End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
End Sub

Private Sub OPTTOPN_Click()
If OPTTOPN.Value = True Then
CHKMOREOPTION.Enabled = False
Else
CHKMOREOPTION.Enabled = True

End If
End Sub
