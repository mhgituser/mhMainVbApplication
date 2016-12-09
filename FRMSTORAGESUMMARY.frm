VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMSTORAGESUMMARY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STORAGE SUMMARY"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   Icon            =   "FRMSTORAGESUMMARY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "DATE SELECTION"
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton OPTSEL 
         Caption         =   "SELECTIVE"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   855
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
      Left            =   600
      Picture         =   "FRMSTORAGESUMMARY.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
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
      Left            =   2280
      Picture         =   "FRMSTORAGESUMMARY.frx":0ED4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1080
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
         ItemData        =   "FRMSTORAGESUMMARY.frx":1B9E
         Left            =   1080
         List            =   "FRMSTORAGESUMMARY.frx":1BA0
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81985537
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81985537
         CurrentDate     =   41362
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   1200
         Width           =   705
      End
   End
End
Attribute VB_Name = "FRMSTORAGESUMMARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub storagesumall()
Dim rsfr As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsf As New ADODB.Recordset
Dim frcount As Integer
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcount = 0
db.Open OdkCnnString
                      
GetTbl

 SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
         & "waterlog,leafpest,stempest,animaldamage,monitorcomments)  select end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,ndtrees,wlogged,pdamage,ddamage,adamage,monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'  GROUP BY n.farmerbarcode"
         
db.Execute SQLSTR
SQLSTR = ""
SQLSTR = "select count(farmercode) as frcount,'' as fieldcode,sum(totaltrees) as totaltrees,'' as area,sum(animaldamage) as adamage,sum(leafpest) as pdamage,sum(stempest) as ddamage,sum(tree_count_deadmissing) as dtrees,sum(waterlog) as wlogged,sum(nutrient) as ndtrees from   " & Mtblname & ""



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
    excel_sheet.Cells(1, 1) = UCase("Total No. of farmers in storaged")
    excel_sheet.Cells(2, 1) = UCase("Total No. of trees in the storage")
    excel_sheet.Cells(3, 1) = UCase("Total acres")
    excel_sheet.Cells(4, 1) = UCase("Animal Damage")
    excel_sheet.Cells(5, 1) = UCase("Pest Damage")
    excel_sheet.Cells(6, 1) = UCase("Disease Damge ")
    excel_sheet.Cells(7, 1) = UCase("Dead Trees")
    excel_sheet.Cells(8, 1) = UCase("Waterlogged")
    excel_sheet.Cells(9, 1) = UCase("Nutrient Deficient")
    
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.Cells(1, 2) = rs!frcount
    excel_sheet.Cells(2, 2) = rs!totaltrees
    excel_sheet.Cells(3, 2) = rs![Area]
    excel_sheet.Cells(4, 2) = rs!adamage
    excel_sheet.Cells(5, 2) = rs!pdamage
    excel_sheet.Cells(6, 2) = rs!ddamage
    excel_sheet.Cells(7, 2) = rs!dtrees
    excel_sheet.Cells(8, 2) = rs!wlogged
    excel_sheet.Cells(9, 2) = rs!ndtrees
  

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
    
     excel_sheet.Range("a1:b9").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "STORAGE SUMMARY"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("a9").Select
 excel_app.Selection.ColumnWidth = 35
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With


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

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If OPTALL.Value = True Then
storagesumall
Else
storagesumsel
End If
End Sub

Private Sub storagesumsel()
Dim rsfr As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsf As New ADODB.Recordset
Dim frcount As Integer
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcount = 0
db.Open OdkCnnString
                    
GetTbl

 SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,nutrient," _
         & "waterlog,leafpest,stempest,animaldamage,monitorcomments)  select end,staffbarcode,region_dcode," _
         & "region_gcode,region,n.farmerbarcode,totaltrees,gmoisture,pmoisture,gmoisture+pmoisture," _
         & "dtrees,ndtrees,wlogged,pdamage,ddamage,adamage,monitorcomments from storagehub6_core n INNER JOIN (SELECT farmerbarcode,MAX(END )" _
         & "lastEdit FROM storagehub6_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' GROUP BY n.farmerbarcode"
         
db.Execute SQLSTR
SQLSTR = ""
SQLSTR = "select count(farmercode) as frcount,'' as fieldcode,sum(totaltrees) as totaltrees,'' as area,sum(animaldamage) as adamage,sum(leafpest) as pdamage,sum(stempest) as ddamage,sum(tree_count_deadmissing) as dtrees,sum(waterlog) as wlogged,sum(nutrient) as ndtrees from   " & Mtblname & ""



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
    excel_sheet.Cells(1, 1) = UCase("Total No. of farmers in storaged")
    excel_sheet.Cells(2, 1) = UCase("Total No. of trees in the storage")
    excel_sheet.Cells(3, 1) = UCase("Total acres")
    excel_sheet.Cells(4, 1) = UCase("Animal Damage")
    excel_sheet.Cells(5, 1) = UCase("Pest Damage")
    excel_sheet.Cells(6, 1) = UCase("Disease Damge ")
    excel_sheet.Cells(7, 1) = UCase("Dead Trees")
    excel_sheet.Cells(8, 1) = UCase("Waterlogged")
    excel_sheet.Cells(9, 1) = UCase("Nutrient Deficient")
    
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.Cells(1, 2) = rs!frcount
    excel_sheet.Cells(2, 2) = rs!totaltrees
    excel_sheet.Cells(3, 2) = rs![Area]
    excel_sheet.Cells(4, 2) = rs!adamage
    excel_sheet.Cells(5, 2) = rs!pdamage
    excel_sheet.Cells(6, 2) = rs!ddamage
    excel_sheet.Cells(7, 2) = rs!dtrees
    excel_sheet.Cells(8, 2) = rs!wlogged
    excel_sheet.Cells(9, 2) = rs!ndtrees
  

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
    
     excel_sheet.Range("a1:b9").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "STORAGE SUMMARY"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("a9").Select
 excel_app.Selection.ColumnWidth = 35
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With


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


Private Sub OPTALL_Click()
Frame1.Enabled = False
End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
End Sub
