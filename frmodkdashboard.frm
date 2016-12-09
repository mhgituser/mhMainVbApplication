VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmodkdashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODK DASHBOARD"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   Icon            =   "frmodkdashboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Frame Frame2 
         Caption         =   "SHEET"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   3015
         Begin VB.OptionButton optmortality 
            Caption         =   "Mortality"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Field"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   -2040
            Width           =   735
         End
         Begin VB.OptionButton optfieldsnapshot 
            Caption         =   "Field Snapshot"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optstorage 
            Caption         =   "Storage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optfield 
            Caption         =   "Field"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
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
         Left            =   2520
         Picture         =   "frmodkdashboard.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3360
         Width           =   1215
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
         Left            =   1200
         Picture         =   "frmodkdashboard.frx":1B0C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox chkfarmer 
         Caption         =   "FARMER"
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkgrf 
         Caption         =   "GRF"
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkcf 
         Caption         =   "CF"
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "frmodkdashboard.frx":2276
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "frmodkdashboard.frx":228B
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
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
      Begin MSDataListLib.DataCombo CBOTSHOWOG 
         Bindings        =   "frmodkdashboard.frx":22A0
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   2880
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
      Begin MSDataListLib.DataCombo cbodept 
         Bindings        =   "frmodkdashboard.frx":22B5
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SHEET"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG"
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
         TabIndex        =   8
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GEWOG"
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
         Left            =   720
         TabIndex        =   7
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TSHOWOG"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3000
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmodkdashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sheetname As String
Dim sheetcase As Integer
Private Sub cboDzongkhag_Change()
cbogewog.Text = ""
CBOTSHOWOG.Text = ""
End Sub

Private Sub cboDzongkhag_LostFocus()
Dim rsGe As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsGe = Nothing

If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where DzongkhagId='" & cboDzongkhag.BoundText & "' order by dzongkhagid,gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"
End Sub

Private Sub cbogewog_Change()
CBOTSHOWOG.Text = ""
End Sub

Private Sub cbogewog_LostFocus()

Dim rsTs As New ADODB.Recordset


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsTs = Nothing

If rsTs.State = adStateOpen Then rsTs.Close
rsTs.Open "select concat(tshewogid , ' ', tshewogname) as tshewogname,tshewogid  from tbltshewog where dzongkhagid='" & cboDzongkhag.BoundText & "' and gewogid='" & cbogewog.BoundText & "' order by dzongkhagid,gewogid", db
Set CBOTSHOWOG.RowSource = rsTs
CBOTSHOWOG.ListField = "tshewogname"
CBOTSHOWOG.BoundColumn = "tshewogid"
End Sub

Private Sub Command1_Click()
On Error GoTo err



If chkfarmer.Value = 0 And chkcf.Value = 0 And chkgrf.Value = 0 Then
MsgBox "Select one of the farmers type."
Exit Sub
End If


Select Case sheetcase
Case 1
        dzfield

Case 2
        dzstorage
Case 3
        fieldsnapshot
Case 4
        mortality

End Select

Exit Sub
err:
    MsgBox err.Description
    
    
    
End Sub
Private Sub mortality()
 Dim xl As Excel.Application
    Dim rs As New ADODB.Recordset
    Dim SQLSTR As String
    Dim var As Variant
    Dim i, j As Integer
    Set xl = CreateObject("excel.Application")
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    End If
    
    FileCopy excelPath + "\template\" & LCase(sheetname) & ".xlsx", App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Sheets("Data").Select
    xl.Visible = False
    

    
   
   
   
   
       GetTbl
        
    
SQLSTR = ""

           
         SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,area) select n.end,staffbarcode," _
         & "region_gcode,region,n.farmerbarcode,0,n.fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,'0' from phealthhub15_core n INNER JOIN (SELECT farmerbarcode,fdcode, MAX(END )" _
         & "lastEdit FROM phealthhub15_core GROUP BY farmerbarcode,fdcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode, n.fdcode"
           
           
  ODKDB.Execute SQLSTR


Set rs = Nothing
rs.Open "select end, substring(farmercode,1,3) dz,substring(farmercode,4,3) ge,substring(farmercode,7,3) ts,farmercode,dcode,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & "   group by farmercode order by farmercode", ODKDB
i = 0
Do While rs.EOF <> True
FindDZ rs!dz
FindGE rs!dz, rs!ge
FindTs rs!dz, rs!ge, rs!ts
FindFA rs!farmercode, "F"
FindsTAFF rs!dcode


xl.Cells(7 + i, 2) = rs!dz & "  " & Dzname
xl.Cells(7 + i, 3) = rs!ge & "  " & GEname
xl.Cells(7 + i, 4) = rs!ts & "  " & TsName
xl.Cells(7 + i, 5) = rs!farmercode
xl.Cells(7 + i, 6) = FAName
xl.Cells(7 + i, 7) = rs!dcode & "  " & sTAFF
xl.Cells(7 + i, 9) = "'" & Format(rs!End, "dd/MM/yyyy")
xl.Cells(7 + i, 11) = rs!totaltrees
xl.Cells(7 + i, 12) = rs!deadmissing


i = i + 1
rs.MoveNext
Loop


'If Len(cboDzongkhag.Text) > 0 Then
'
'xl.Sheets("Field health (Ge)").Select
'xl.Cells(49, 4) = cboDzongkhag.Text
'Set rs = Nothing
'rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
'& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
'& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
'& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "'  group by dz,ge order by dz,ge", ODKDB
'i = 0
'Do While rs.EOF <> True
'FindGE rs!dz, rs!ge
'xl.Cells(52 + i, 3) = rs!ge
'xl.Cells(52 + i, 4) = GEname
'xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = rs!activegrowing
'xl.Cells(52 + i, 7) = rs!slowgrowing
'xl.Cells(52 + i, 8) = rs!dor
'xl.Cells(52 + i, 9) = rs!deadmissing
'xl.Cells(52 + i, 10) = rs!poormoisture
'xl.Cells(52 + i, 11) = rs!nutrient
'xl.Cells(52 + i, 12) = rs!waterlog
'xl.Cells(52 + i, 13) = rs!activepest
'xl.Cells(52 + i, 14) = rs!animaldamage
'i = i + 1
'rs.MoveNext
'Loop
'
'
'
'
'
'End If
'
'
'If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 Then
'
'xl.Sheets("Field health (Ts)").Select
'xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text
'Set rs = Nothing
'rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
'& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
'& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
'& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' group by dz,ge,ts order by dz,ge,ts", ODKDB
'i = 0
'Do While rs.EOF <> True
'FindTs rs!dz, rs!ge, rs!ts
'xl.Cells(52 + i, 3) = rs!ts
'xl.Cells(52 + i, 4) = TsName
'xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = rs!activegrowing
'xl.Cells(52 + i, 7) = rs!slowgrowing
'xl.Cells(52 + i, 8) = rs!dor
'xl.Cells(52 + i, 9) = rs!deadmissing
'xl.Cells(52 + i, 10) = rs!poormoisture
'xl.Cells(52 + i, 11) = rs!nutrient
'xl.Cells(52 + i, 12) = rs!waterlog
'xl.Cells(52 + i, 13) = rs!activepest
'xl.Cells(52 + i, 14) = rs!animaldamage
'i = i + 1
'rs.MoveNext
'Loop
'
'
'
'
'
'End If
'
'
'
'
'
'
'If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 And Len(CBOTSHOWOG.Text) > 0 Then
'
'xl.Sheets("Field health (Farmer)").Select
'xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text & "  " & CBOTSHOWOG.Text
'Set rs = Nothing
'rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees),farmercode, totaltrees,sum(tree_count_activegrowing) activegrowing," _
'& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
'& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
'& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' and substring(farmercode,7,3)='" & CBOTSHOWOG.BoundText & "' group by dz,ge,ts,farmercode order by dz,ge,ts,farmercode", ODKDB
'i = 0
'Do While rs.EOF <> True
'FindFA rs!farmercode, "F"
'xl.Cells(52 + i, 3) = rs!farmercode
'xl.Cells(52 + i, 4) = FAName
'xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = rs!activegrowing
'xl.Cells(52 + i, 7) = rs!slowgrowing
'xl.Cells(52 + i, 8) = rs!dor
'xl.Cells(52 + i, 9) = rs!deadmissing
'xl.Cells(52 + i, 10) = rs!poormoisture
'xl.Cells(52 + i, 11) = rs!nutrient
'xl.Cells(52 + i, 12) = rs!waterlog
'xl.Cells(52 + i, 13) = rs!activepest
'xl.Cells(52 + i, 14) = rs!animaldamage
'i = i + 1
'rs.MoveNext
'Loop
'
'
'
'
'
'End If






ODKDB.Execute "drop table " & Mtblname & ""
xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault
End Sub
Private Sub fieldsnapshot()
 Dim xl As Excel.Application
    Dim rs As New ADODB.Recordset
    Dim rst As New ADODB.Recordset
    Dim SQLSTR As String
    Dim var As Variant
    Dim i, j As Integer
    Set xl = CreateObject("excel.Application")
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    End If
    
  Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='7'", MHVDB
    If rs.EOF <> True Then
    getSheet 7, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Sheets("Field Snapshot (Dz)").Select
    xl.Visible = False
    

    
   
   
   
   
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
           
           
  ODKDB.Execute SQLSTR


Set rs = Nothing
rs.Open "select count(farmercode) cnt, mid(farmercode,1,3) dz,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & "   group by dz order by dz", ODKDB
i = 0
Do While rs.EOF <> True
FindDZ rs!dz
xl.Cells(49 + i, 2) = rs!dz & "  " & Dzname

Set rst = Nothing
rst.Open "select sum(regland) as regland from tbllandreg where substring(farmerid,1,3)='" & rs!dz & "' group by substring(farmerid,1,3) ", MHVDB
If rs.EOF <> True Then
xl.Cells(49 + i, 3) = IIf(IsNull(rst!regland), 0, rst!regland)
End If
xl.Cells(49 + i, 5) = rs!cnt



i = i + 1
rs.MoveNext
Loop


'If Len(cboDzongkhag.Text) > 0 Then
'
'xl.Sheets("Field health (Ge)").Select
'xl.Cells(49, 4) = cboDzongkhag.Text
'Set rs = Nothing
'rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
'& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
'& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
'& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "'  group by dz,ge order by dz,ge", ODKDB
'i = 0
'Do While rs.EOF <> True
'FindGE rs!dz, rs!ge
'xl.Cells(52 + i, 3) = rs!ge
'xl.Cells(52 + i, 4) = GEname
'xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = rs!activegrowing
'xl.Cells(52 + i, 7) = rs!slowgrowing
'xl.Cells(52 + i, 8) = rs!dor
'xl.Cells(52 + i, 9) = rs!deadmissing
'xl.Cells(52 + i, 10) = rs!poormoisture
'xl.Cells(52 + i, 11) = rs!nutrient
'xl.Cells(52 + i, 12) = rs!waterlog
'xl.Cells(52 + i, 13) = rs!activepest
'xl.Cells(52 + i, 14) = rs!animaldamage
'i = i + 1
'rs.MoveNext
'Loop
'
'
'
'
'
'End If
'
'
'If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 Then
'
'xl.Sheets("Field health (Ts)").Select
'xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text
'Set rs = Nothing
'rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
'& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
'& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
'& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' group by dz,ge,ts order by dz,ge,ts", ODKDB
'i = 0
'Do While rs.EOF <> True
'FindTs rs!dz, rs!ge, rs!ts
'xl.Cells(52 + i, 3) = rs!ts
'xl.Cells(52 + i, 4) = TsName
'xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = rs!activegrowing
'xl.Cells(52 + i, 7) = rs!slowgrowing
'xl.Cells(52 + i, 8) = rs!dor
'xl.Cells(52 + i, 9) = rs!deadmissing
'xl.Cells(52 + i, 10) = rs!poormoisture
'xl.Cells(52 + i, 11) = rs!nutrient
'xl.Cells(52 + i, 12) = rs!waterlog
'xl.Cells(52 + i, 13) = rs!activepest
'xl.Cells(52 + i, 14) = rs!animaldamage
'i = i + 1
'rs.MoveNext
'Loop
'
'
'
'
'
'End If
'
'
'
'
'
'
'If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 And Len(CBOTSHOWOG.Text) > 0 Then
'
'xl.Sheets("Field health (Farmer)").Select
'xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text & "  " & CBOTSHOWOG.Text
'Set rs = Nothing
'rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees),farmercode, totaltrees,sum(tree_count_activegrowing) activegrowing," _
'& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
'& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
'& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' and substring(farmercode,7,3)='" & CBOTSHOWOG.BoundText & "' group by dz,ge,ts,farmercode order by dz,ge,ts,farmercode", ODKDB
'i = 0
'Do While rs.EOF <> True
'FindFA rs!farmercode, "F"
'xl.Cells(52 + i, 3) = rs!farmercode
'xl.Cells(52 + i, 4) = FAName
'xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = rs!activegrowing
'xl.Cells(52 + i, 7) = rs!slowgrowing
'xl.Cells(52 + i, 8) = rs!dor
'xl.Cells(52 + i, 9) = rs!deadmissing
'xl.Cells(52 + i, 10) = rs!poormoisture
'xl.Cells(52 + i, 11) = rs!nutrient
'xl.Cells(52 + i, 12) = rs!waterlog
'xl.Cells(52 + i, 13) = rs!activepest
'xl.Cells(52 + i, 14) = rs!animaldamage
'i = i + 1
'rs.MoveNext
'Loop
'
'
'
'
'
'End If






ODKDB.Execute "drop table " & Mtblname & ""
xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault
End Sub
Private Sub dzstorage()
    Dim xl As Excel.Application
    Dim rs As New ADODB.Recordset
    Dim SQLSTR As String
    Dim var As Variant
    Dim i, j As Integer
    Set xl = CreateObject("excel.Application")
    If Dir$(App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx") <> vbNullString Then
    Kill App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    End If
    
    Set rs = Nothing
    rs.Open "select * from tbldashbordtrn where trnid='5'", MHVDB
    If rs.EOF <> True Then
    getSheet 5, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Sheets("Storage health (Dz)").Select
    xl.Visible = False
    

    
   
   
   
   
       GetTbl
        
    
SQLSTR = ""

           
         SQLSTR = "insert into " & Mtblname & " (start,tdate,end,id,dcode,gcode,tcode,farmercode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
         & "waterlog,leafpest,animaldamage,activepest,monitorcomments)  select start,tdate,end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,ndtrees,wlogged,pdamage,adamage,ddamage,monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
        & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
           
           
  ODKDB.Execute SQLSTR


Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz,sum(totaltrees) totaltrees," _
& "sum(tree_count_deadmissing) deadmissing,sum(ANIMALDAMAGE) ANIMALDAMAGE," _
& "sum(leafpest) pestdamage,sum(activepest) ddamage,sum(nutrient) nutrient,sum(waterlog) waterlog " _
& " from " & Mtblname & "   group by dz order by dz", ODKDB
i = 0
Do While rs.EOF <> True
FindDZ rs!dz
xl.Cells(52 + i, 3) = rs!dz
xl.Cells(52 + i, 4) = Dzname
xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = ""
xl.Cells(52 + i, 7) = rs!deadmissing
xl.Cells(52 + i, 8) = rs!animaldamage
xl.Cells(52 + i, 9) = rs!pestdamage
xl.Cells(52 + i, 10) = rs!ddamage
xl.Cells(52 + i, 11) = rs!waterlog
xl.Cells(52 + i, 12) = rs!nutrient
i = i + 1
rs.MoveNext
Loop


If Len(cboDzongkhag.Text) > 0 Then
xl.Sheets("Storage health (Ge)").Select
xl.Cells(49, 4) = cboDzongkhag.Text
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz,mid(farmercode,4,3) ge,sum(totaltrees) totaltrees," _
& "sum(tree_count_deadmissing) deadmissing,sum(ANIMALDAMAGE) ANIMALDAMAGE," _
& "sum(leafpest) pestdamage,sum(activepest) ddamage,sum(nutrient) nutrient,sum(waterlog) waterlog " _
& " from " & Mtblname & "  where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' group by dz,ge order by dz,ge", ODKDB
i = 0
Do While rs.EOF <> True
FindGE rs!dz, rs!ge
xl.Cells(52 + i, 3) = rs!ge
xl.Cells(52 + i, 4) = GEname
xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = ""
xl.Cells(52 + i, 7) = rs!deadmissing
xl.Cells(52 + i, 8) = rs!animaldamage
xl.Cells(52 + i, 9) = rs!pestdamage
xl.Cells(52 + i, 10) = rs!ddamage
xl.Cells(52 + i, 11) = rs!waterlog
xl.Cells(52 + i, 12) = rs!nutrient
i = i + 1
rs.MoveNext
Loop


End If



If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 Then
xl.Sheets("Storage health (Ts)").Select
xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees) totaltrees," _
& "sum(tree_count_deadmissing) deadmissing,sum(ANIMALDAMAGE) ANIMALDAMAGE," _
& "sum(leafpest) pestdamage,sum(activepest) ddamage,sum(nutrient) nutrient,sum(waterlog) waterlog " _
& " from " & Mtblname & "  where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "'  group by dz,ge,ts order by dz,ge,ts", ODKDB
i = 0
Do While rs.EOF <> True
FindTs rs!dz, rs!ge, rs!ts
xl.Cells(52 + i, 3) = rs!ts
xl.Cells(52 + i, 4) = TsName
xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = ""
xl.Cells(52 + i, 7) = rs!deadmissing
xl.Cells(52 + i, 8) = rs!animaldamage
xl.Cells(52 + i, 9) = rs!pestdamage
xl.Cells(52 + i, 10) = rs!ddamage
xl.Cells(52 + i, 11) = rs!waterlog
xl.Cells(52 + i, 12) = rs!nutrient
i = i + 1
rs.MoveNext
Loop


End If







If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 And Len(CBOTSHOWOG.Text) > 0 Then
xl.Sheets("Storage health (Farmers)").Select
xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text & "  " & CBOTSHOWOG.Text
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,farmercode,sum(totaltrees) totaltrees," _
& "sum(tree_count_deadmissing) deadmissing,sum(ANIMALDAMAGE) ANIMALDAMAGE," _
& "sum(leafpest) pestdamage,sum(activepest) ddamage,sum(nutrient) nutrient,sum(waterlog) waterlog " _
& " from " & Mtblname & "   where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' and substring(farmercode,7,3)='" & CBOTSHOWOG.BoundText & "' group by dz,ge,ts,farmercode order by dz,ge,ts,farmercode", ODKDB
i = 0
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
xl.Cells(52 + i, 3) = rs!farmercode
xl.Cells(52 + i, 4) = FAName
xl.Cells(52 + i, 5) = rs!totaltrees
'xl.Cells(52 + i, 6) = ""
xl.Cells(52 + i, 7) = rs!deadmissing
xl.Cells(52 + i, 8) = rs!animaldamage
xl.Cells(52 + i, 9) = rs!pestdamage
xl.Cells(52 + i, 10) = rs!ddamage
xl.Cells(52 + i, 11) = rs!waterlog
xl.Cells(52 + i, 12) = rs!nutrient
i = i + 1
rs.MoveNext
Loop

End If






ODKDB.Execute "drop table " & Mtblname & ""
xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault
        
        
End Sub
Private Sub dzfield()
    Dim xl As Excel.Application
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
    getSheet 4, rs!FileName
    End If
    FileCopy dashBoardName, App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    If Dir$(dashBoardName) <> vbNullString Then
        Kill dashBoardName
    End If
    
    
    xl.Workbooks.Open App.Path + "\" + Format(Now, "ddMMyyyy") + " " + cbodept.Text + ".xlsx"
    xl.Sheets("Field health (Dz)").Select
    xl.Visible = False
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
           
           
  ODKDB.Execute SQLSTR


Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & "   group by dz order by dz", ODKDB
i = 0
Do While rs.EOF <> True
FindDZ rs!dz
xl.Cells(52 + i, 3) = rs!dz
xl.Cells(52 + i, 4) = Dzname
xl.Cells(52 + i, 5) = rs!totaltrees
xl.Cells(52 + i, 6) = rs!activegrowing
xl.Cells(52 + i, 7) = rs!slowgrowing
xl.Cells(52 + i, 8) = rs!dor
xl.Cells(52 + i, 9) = rs!deadmissing
xl.Cells(52 + i, 10) = rs!poormoisture
xl.Cells(52 + i, 11) = rs!nutrient
xl.Cells(52 + i, 12) = rs!waterlog
xl.Cells(52 + i, 13) = rs!activepest
xl.Cells(52 + i, 14) = rs!animaldamage
i = i + 1
rs.MoveNext
Loop


If Len(cboDzongkhag.Text) > 0 Then

xl.Sheets("Field health (Ge)").Select
xl.Cells(49, 4) = cboDzongkhag.Text
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "'  group by dz,ge order by dz,ge", ODKDB
i = 0
Do While rs.EOF <> True
FindGE rs!dz, rs!ge
xl.Cells(52 + i, 3) = rs!ge
xl.Cells(52 + i, 4) = GEname
xl.Cells(52 + i, 5) = rs!totaltrees
xl.Cells(52 + i, 6) = rs!activegrowing
xl.Cells(52 + i, 7) = rs!slowgrowing
xl.Cells(52 + i, 8) = rs!dor
xl.Cells(52 + i, 9) = rs!deadmissing
xl.Cells(52 + i, 10) = rs!poormoisture
xl.Cells(52 + i, 11) = rs!nutrient
xl.Cells(52 + i, 12) = rs!waterlog
xl.Cells(52 + i, 13) = rs!activepest
xl.Cells(52 + i, 14) = rs!animaldamage
i = i + 1
rs.MoveNext
Loop





End If


If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 Then

xl.Sheets("Field health (Ts)").Select
xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees) totaltrees,sum(tree_count_activegrowing) activegrowing," _
& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' group by dz,ge,ts order by dz,ge,ts", ODKDB
i = 0
Do While rs.EOF <> True
FindTs rs!dz, rs!ge, rs!ts
xl.Cells(52 + i, 3) = rs!ts
xl.Cells(52 + i, 4) = TsName
xl.Cells(52 + i, 5) = rs!totaltrees
xl.Cells(52 + i, 6) = rs!activegrowing
xl.Cells(52 + i, 7) = rs!slowgrowing
xl.Cells(52 + i, 8) = rs!dor
xl.Cells(52 + i, 9) = rs!deadmissing
xl.Cells(52 + i, 10) = rs!poormoisture
xl.Cells(52 + i, 11) = rs!nutrient
xl.Cells(52 + i, 12) = rs!waterlog
xl.Cells(52 + i, 13) = rs!activepest
xl.Cells(52 + i, 14) = rs!animaldamage
i = i + 1
rs.MoveNext
Loop





End If






If Len(cboDzongkhag.Text) > 0 And Len(cbogewog.Text) > 0 And Len(CBOTSHOWOG.Text) > 0 Then

xl.Sheets("Field health (Farmer)").Select
xl.Cells(49, 4) = cboDzongkhag.Text & "  " & cbogewog.Text & "  " & CBOTSHOWOG.Text
Set rs = Nothing
rs.Open "select mid(farmercode,1,3) dz ,mid(farmercode,4,3) ge,mid(farmercode,7,3) ts,sum(totaltrees),farmercode, totaltrees,sum(tree_count_activegrowing) activegrowing," _
& "sum(tree_count_slowgrowing) slowgrowing,sum(tree_count_dor) dor,sum(tree_count_deadmissing) deadmissing," _
& "sum(poormoisture) poormoisture,sum(nutrient) nutrient,sum(waterlog) waterlog,sum(activepest)activepest," _
& "sum(ANIMALDAMAGE) ANIMALDAMAGE from " & Mtblname & " where substring(farmercode,1,3)='" & cboDzongkhag.BoundText & "' and substring(farmercode,4,3)='" & cbogewog.BoundText & "' and substring(farmercode,7,3)='" & CBOTSHOWOG.BoundText & "' group by dz,ge,ts,farmercode order by dz,ge,ts,farmercode", ODKDB
i = 0
Do While rs.EOF <> True
FindFA rs!farmercode, "F"
xl.Cells(52 + i, 3) = rs!farmercode
xl.Cells(52 + i, 4) = FAName
xl.Cells(52 + i, 5) = rs!totaltrees
xl.Cells(52 + i, 6) = rs!activegrowing
xl.Cells(52 + i, 7) = rs!slowgrowing
xl.Cells(52 + i, 8) = rs!dor
xl.Cells(52 + i, 9) = rs!deadmissing
xl.Cells(52 + i, 10) = rs!poormoisture
xl.Cells(52 + i, 11) = rs!nutrient
xl.Cells(52 + i, 12) = rs!waterlog
xl.Cells(52 + i, 13) = rs!activepest
xl.Cells(52 + i, 14) = rs!animaldamage
i = i + 1
rs.MoveNext
Loop





End If






ODKDB.Execute "drop table " & Mtblname & ""
xl.Visible = True
Set xl = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub ge()

End Sub
Private Sub ts()

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
Dim rsDz As New ADODB.Recordset

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cboDzongkhag.RowSource = rsDz
cboDzongkhag.ListField = "dzongkhagname"
cboDzongkhag.BoundColumn = "dzongkhagcode"
   
       
       
    Set rsDz = Nothing
If rsDz.State = adStateOpen Then rs.Close
If UCase(MUSER) = "ADMIN" Then
rsDz.Open "select deptid,deptname from tbldept where deptid in(6,7) order by deptid", db
Else

 rsDz.Open "select deptid,deptname from tbldept where deptid in(6,7) and remarks like  " & "'%" & UserId & "%'" & "  order by deptid", db

End If
Set cbodept.RowSource = rsDz
cbodept.ListField = "deptname"
cbodept.BoundColumn = "deptid"
       
sheetname = "field"
  sheetcase = 1
      



            

       




   
      
      
      
      
      
      
      
      
      
      
      
      

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub optfield_Click()
sheetname = "field"
sheetcase = 1
End Sub

Private Sub optfieldsnapshot_Click()
sheetname = "fieldsnapshot"
sheetcase = 3
End Sub

Private Sub optmortality_Click()
sheetname = "mortality"
sheetcase = 4
End Sub

Private Sub optstorage_Click()
sheetname = "storage"
sheetcase = 2
End Sub
