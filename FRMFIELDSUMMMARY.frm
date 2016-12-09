VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMFIELDSUMMMARY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIELD SUMMARY"
   ClientHeight    =   4125
   ClientLeft      =   7005
   ClientTop       =   2955
   ClientWidth     =   5310
   Icon            =   "FRMFIELDSUMMMARY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5310
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   1200
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
         ItemData        =   "FRMFIELDSUMMMARY.frx":076A
         Left            =   1080
         List            =   "FRMFIELDSUMMMARY.frx":076C
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78512129
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78512129
         CurrentDate     =   41362
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   900
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
      Left            =   2280
      Picture         =   "FRMFIELDSUMMMARY.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
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
      Left            =   600
      Picture         =   "FRMFIELDSUMMMARY.frx":1438
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "DATE SELECTION"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OPTSEL 
         Caption         =   "SELECTIVE"
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FRMFIELDSUMMMARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If OPTALL.Value = True Then
FIELDALL
Else
FIELDSEL
End If
End Sub
Private Sub FIELDSEL()
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
                       
If OPTALL.Value = True Then
Mindex = 51
End If

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
'If OPTALL.Value = True Then


          
'       SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select n.end,staffbarcode,dcode," _
         & "gcode,tcode,n.farmerbarcode,'0',n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture as totaltally," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'GROUP BY n.farmerbarcode, n.fdcode"
          
          
              
            SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' GROUP BY n.farmerbarcode, n.fdcode"
          
          
'
' Else
'
'
'
'
'      SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
'         & "goodmoisture,poormoisture,totaltally,deadmissing,slowgrowing,dor,activegrowing,shock,nutrient," _
'         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select n.end,id,dcode," _
'         & "gcode,tcode,n.farmerbarcode,treesreceived,n.fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
'         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
'         & "rootpest,animaldamage,monitorcomments from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
'         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
'         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit AND  SUBSTRING(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'" _
'         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
'
'
' End If
'
 
          
  db.Execute SQLSTR


SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead,sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(leafpest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""


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
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    excel_sheet.Cells(1, 1) = UCase("Total No. of hazelnut field")
    excel_sheet.Cells(2, 1) = UCase("Total No. of trees in the field")
    excel_sheet.Cells(3, 1) = UCase("Total acres")
    excel_sheet.Cells(4, 1) = UCase("Slow growing")
    excel_sheet.Cells(5, 1) = UCase("Dormant")
    excel_sheet.Cells(6, 1) = UCase("Dead ")
    excel_sheet.Cells(7, 1) = UCase("Active growing")
    excel_sheet.Cells(8, 1) = UCase("Shock")
    excel_sheet.Cells(9, 1) = UCase("Nutrient deficeint")
    excel_sheet.Cells(10, 1) = UCase("Waterlog")
    excel_sheet.Cells(11, 1) = UCase("Leafpest")
    excel_sheet.Cells(12, 1) = UCase("Active pest")
    excel_sheet.Cells(13, 1) = UCase("Stem pest")
    excel_sheet.Cells(14, 1) = UCase("Root pest")
    excel_sheet.Cells(15, 1) = UCase("Animal Damage")
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.Cells(1, 2) = rs!fieldcode
    excel_sheet.Cells(2, 2) = rs!totaltrees
    excel_sheet.Cells(3, 2) = rs![area]
    excel_sheet.Cells(4, 2) = rs!slowgrowing
    excel_sheet.Cells(5, 2) = rs!dor
    excel_sheet.Cells(6, 2) = rs!dead
    excel_sheet.Cells(7, 2) = rs!activegrowing
    excel_sheet.Cells(8, 2) = rs!shock
    excel_sheet.Cells(9, 2) = rs!nutrient
    excel_sheet.Cells(10, 2) = rs!waterlog
    excel_sheet.Cells(11, 2) = rs!leafpest
    excel_sheet.Cells(12, 2) = rs!activepest
    excel_sheet.Cells(13, 2) = rs!stempest
    excel_sheet.Cells(14, 2) = rs!rootpest
    excel_sheet.Cells(15, 2) = rs!animaldamage

'  Do While rs.EOF <> True
'
'
'
'i = i + 1
'rs.MoveNext
'   Loop



   'make up




'    excel_sheet.Cells(4, 2).Select
'    excel_app.ActiveWindow.FreezePanes = True
'    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    
     excel_sheet.Range("a1:b15").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "FIELDS SUMMARY"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("a").Select
 excel_app.Selection.ColumnWidth = 31



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

Private Sub FIELDALL()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
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
                    
If OPTALL.Value = True Then
Mindex = 51
End If

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
          
'         SQLSTR = " select n.end,n.dcode," _
'         & "n.gcode,n.tcode,n.fcode,farmerbarcode,treesreceived,n.fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
'         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
'         & "rootpest,animaldamage,area from phealthhub15_core n INNER JOIN (SELECT dcode,gcode,tcode,fcode,fdcode, MAX(END )" _
'         & "lastEdit FROM phealthhub15_core GROUP BY dcode,gcode,tcode,fcode,fdcode)x ON " _
'         & "n.dcode = x.dcode and n.gcode=x.gcode and n.tcode=x.tcode and n.fcode=x.fcode AND n.end = x.LastEdit AND n.farmerbarcode =''" _
'         & "AND STATUS <>  'BAD'GROUP BY n.dcode,n.gcode,n.tcode,n.fcode, n.fdcode"
'
'     rss.Open SQLSTR, db
'  Do While rss.EOF <> True
'
'  mdcode = "0000" & rss!dcode
'  mdcode = "D" & Right(mdcode, 2)
'  mgcode = "00000" & rss!gcode
'  mgcode = "G" & Right(mgcode, 2)
'  mtcode = "0000" & rss!tcode
'  mtcode = "T" & Right(mtcode, 2)
'  mfcode = "000000" & CStr(rss!fcode)
'  mfcode = "F" & Right(mfcode, 4)
'  mfcode = mdcode & mgcode & mtcode & mfcode
'
'  Set rsf = Nothing
'
'  rsf.Open "select * from " & Mtblname & " where farmercode='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
'  If rsf.EOF <> True Then
'
'  If rss!End > rsf!End Then
'  db.Execute "update " & Mtblname & " set end='" & Format(rss!End, "yyyy-MM-dd") & "' ," _
'            & "totaltrees='" & rss!totaltrees & "',area='" & rss!Area & "'," _
'            & "slowgrowing='" & rss!slowgrowing & "',dor='" & rss!dor & "'," _
'            & "deadmissing='" & rss!deadmissing & "',activegrowing='" & rss!activegrowing & "'," _
'            & "shock='" & rss!shock & "',nutrient='" & rss!nutrient & "'," _
'            & "waterlog='" & rss!waterlog & "',leafpest='" & rss!leafpest & "'," _
'            & "activepest='" & rss!activepest & "',stempest='" & rss!stempest & "'," _
'            & "rootpest='" & rss!rootpest & "',animaldamage='" & rss!animaldamage & "'" _
'            & "where farmercode='" & mfcode & "'  and fdcode='" & rss!FDCODE & "' "
'  End If
'
'
'  Else
'
'  db.Execute "insert into  " & Mtblname & "(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage)values('" & Format(rss!End, "yyyy-MM-dd") & "','" & rss!dcode & "','" & rss!gcode & "','" & rss!tcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "','" & rss!Area & "','" & rss!slowgrowing & "','" & rss!dor & "','" & rss!deadmissing & "','" & rss!activegrowing & "','" & rss!shock & "','" & rss!nutrient & "','" & rss!waterlog & "','" & rss!leafpest & "','" & rss!activepest & "','" & rss!stempest & "','" & rss!rootpest & "','" & rss!animaldamage & "') "
'
'  End If
'
'  rss.MoveNext
'  Loop
  

SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area, " _
& " sum(tree_count_slowgrowing) as slowgrowing,sum(tree_count_dor) as dor,sum(tree_count_deadmissing) as dead, " _
& " sum(tree_count_activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as " _
& " waterlog,sum(activepest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) " _
& " as rootpest,sum(animaldamage) as animaldamage from   " & Mtblname & ""


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
    'excel_app.Caption = "MHV"
    Dim sl As Integer
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    excel_sheet.Cells(1, 1) = ProperCase("Total No. of hazelnut field")
    excel_sheet.Cells(2, 1) = ProperCase("Total No. of trees in the field")
    excel_sheet.Cells(3, 1) = ProperCase("Total acres")
    excel_sheet.Cells(4, 1) = ProperCase("Slow growing")
    excel_sheet.Cells(5, 1) = ProperCase("Dormant")
    excel_sheet.Cells(6, 1) = ProperCase("Dead ")
    excel_sheet.Cells(7, 1) = ProperCase("Active growing")
    excel_sheet.Cells(8, 1) = ProperCase("Shock")
    excel_sheet.Cells(9, 1) = ProperCase("Nutrient deficeint")
    excel_sheet.Cells(10, 1) = ProperCase("Waterlog")
    excel_sheet.Cells(11, 1) = ProperCase("Leafpest")
    excel_sheet.Cells(12, 1) = ProperCase("Active pest")
    excel_sheet.Cells(13, 1) = ProperCase("Stem pest")
    excel_sheet.Cells(14, 1) = ProperCase("Root pest")
    excel_sheet.Cells(15, 1) = ProperCase("Animal Damage")
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.Cells(1, 2) = rs!fieldcode
    excel_sheet.Cells(2, 2) = rs!totaltrees
    excel_sheet.Cells(3, 2) = totregland
    excel_sheet.Cells(4, 2) = rs!slowgrowing
    excel_sheet.Cells(5, 2) = rs!dor
    excel_sheet.Cells(6, 2) = rs!dead
    excel_sheet.Cells(7, 2) = rs!activegrowing
    excel_sheet.Cells(8, 2) = rs!shock
    excel_sheet.Cells(9, 2) = rs!nutrient
    excel_sheet.Cells(10, 2) = rs!waterlog
    excel_sheet.Cells(11, 2) = rs!leafpest
    excel_sheet.Cells(12, 2) = rs!activepest
    excel_sheet.Cells(13, 2) = rs!stempest
    excel_sheet.Cells(14, 2) = rs!rootpest
    excel_sheet.Cells(15, 2) = rs!animaldamage


    With excel_sheet
    
     excel_sheet.Range("a1:b15").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "FIELDS SUMMARY"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("a").Select
 excel_app.Selection.ColumnWidth = 31


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


Private Sub Form_Load()
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
'populatedate "phealthhub15_core", 11
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

Private Sub OPTALL_Click()
Frame1.Enabled = False
End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
End Sub
