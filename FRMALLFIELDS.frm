VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMALLFIELDS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIELD REPORT"
   ClientHeight    =   4800
   ClientLeft      =   6615
   ClientTop       =   2745
   ClientWidth     =   9630
   Icon            =   "FRMALLFIELDS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9630
   Begin VB.CheckBox CHKSUMMARY 
      Caption         =   "SUMMARY"
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "SUMMARY"
      Height          =   1215
      Left            =   3240
      TabIndex        =   21
      Top             =   720
      Width           =   2175
      Begin VB.OptionButton OPTSTORAGEFIELD 
         Caption         =   "STORAGE FIELD"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton OPTSTORAGE 
         Caption         =   "STORAGE"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton OPTFIELD 
         Caption         =   "FIELD"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox CHKMOREOPTION 
      Caption         =   "MORE OPTION"
      Height          =   195
      Left            =   4080
      TabIndex        =   20
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "MORE OPTION"
      Height          =   3735
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton OPTLEAFPEST 
         Caption         =   "LEAF PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton OPTSTEMPEST 
         Caption         =   "STEM PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optrootpest 
         Caption         =   "ROOT PEST"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TXTVALUE 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Top             =   3120
         Width           =   1095
      End
      Begin VB.OptionButton optdead 
         Caption         =   "DEAD"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optpestdamage 
         Caption         =   "PEST DAMAGE"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optmoist 
         Caption         =   "MOISTURE"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "PERCENTAGE VALUE GREATER THEN"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5295
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
         ItemData        =   "FRMALLFIELDS.frx":076A
         Left            =   1080
         List            =   "FRMALLFIELDS.frx":076C
         TabIndex        =   9
         Top             =   360
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131268609
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131268609
         CurrentDate     =   41362
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   945
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
   End
   Begin VB.OptionButton OPTALL 
      Caption         =   "ALL"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton OPTSEL 
      Caption         =   "SELECTIVE"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame 
      Caption         =   "REPORT TYPE"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton OPTALLSTORAGE 
         Caption         =   "ALL STORAGE"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OPTALLFIELDS 
         Caption         =   "ALL FIELDS"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
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
      Left            =   2160
      Picture         =   "FRMALLFIELDS.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Left            =   480
      Picture         =   "FRMALLFIELDS.frx":1438
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   5760
      X2              =   5760
      Y1              =   0
      Y2              =   4800
   End
End
Attribute VB_Name = "FRMALLFIELDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mindex As Integer
Private Sub CBODATE_LostFocus()
If OPTALLFIELDS.Value = True Then
findindex "phealthhub15_core", 11
Else

findindex "storagehub6_core", 17

End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Check1_Click()
Frame3.Visible = True
CHKMOREOPTION.Enabled = False
Frame.Visible = False
End Sub

Private Sub CHKMOREOPTION_Click()
If CHKMOREOPTION.Value = 1 Then
FRMALLFIELDS.Width = 9645
CHKSUMMARY.Enabled = False
Else
FRMALLFIELDS.Width = 5640
CHKSUMMARY.Enabled = True
End If
optmoist.Value = True
TXTVALUE.Text = 30
End Sub

Private Sub CHKSUMMARY_Click()

If CHKSUMMARY.Value = 1 Then
Frame3.Visible = True
CHKMOREOPTION.Enabled = False
Frame.Visible = False
Else
Frame3.Visible = False
Frame.Visible = True
CHKMOREOPTION.Enabled = True
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()

If OPTSEL.Value = True And Len(CBODATE.Text) = 0 Then
MsgBox "Please Select The Date Type."
Exit Sub
End If
If CHKSUMMARY.Value = 0 Then

If OPTALLFIELDS.Value = True Then
allfields
ElseIf OPTALLSTORAGE.Value = True Then
allstorage
Else
MsgBox "INVALID SELECTION OF OPTION."
End If
Else


If OPTFIELD.Value = True Then
If OPTALL.Value = True Then
FIELDALL
Else
FIELDSEL
End If



ElseIf OPTSTORAGE.Value = True Then
If OPTALL.Value = True Then
storageallsum
Else
storageselsum
End If


Else
If OPTALL.Value = True Then
'STORAGEFIELDALL
Else
'STORAGEFIELDSEL
End If


End If






End If


End Sub
Private Sub storageselsum()
Dim SLNO As Integer
Dim frcount As Integer
Dim rsfr As New ADODB.Recordset
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
frcount = 0
db.Open OdkCnnString
                      
If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
Mindex = 51
End If

Dim SQLSTR As String
SQLSTR = ""
SLNO = 1

SQLSTR = "select '' as fieldcode,sum(totaltrees) as totaltrees,'' as area,sum(adamage) as adamage,sum(pdamage) as pdamage,sum(ddamage) as ddamage,sum(dtrees) as dtrees,sum(wlogged) as wlogged,sum(ndtrees) as ndtrees from   storagehub6_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
Set rsfr = Nothing
rsfr.Open "select count(*) as fieldcode from  storagehub6_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' group by fcode", db

Do While rsfr.EOF <> True

frcount = frcount + rsfr!fieldcode
rsfr.MoveNext
Loop
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
    excel_sheet.cells(1, 1) = UCase("Total No. of farmers in storaged")
    excel_sheet.cells(2, 1) = UCase("Total No. of trees in the storage")
    excel_sheet.cells(3, 1) = UCase("Total acres")
    excel_sheet.cells(4, 1) = UCase("Animal Damage")
    excel_sheet.cells(5, 1) = UCase("Pest Damage")
    excel_sheet.cells(6, 1) = UCase("Disease Damge ")
    excel_sheet.cells(7, 1) = UCase("Dead Trees")
    excel_sheet.cells(8, 1) = UCase("Waterlogged")
    excel_sheet.cells(9, 1) = UCase("Nutrient Deficient")
    
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.cells(1, 2) = frcount
    excel_sheet.cells(2, 2) = rs!totaltrees
    excel_sheet.cells(3, 2) = rs![area]
    excel_sheet.cells(4, 2) = rs!adamage
    excel_sheet.cells(5, 2) = rs!pdamage
    excel_sheet.cells(6, 2) = rs!ddamage
    excel_sheet.cells(7, 2) = rs!dtrees
    excel_sheet.cells(8, 2) = rs!wlogged
    excel_sheet.cells(9, 2) = rs!ndtrees
  

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
 excel_app.selection.columnWidth = 35
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
db.Close
Exit Sub
err:
MsgBox err.Description
err.Clear
End Sub
Private Sub storageallsum()
Dim rsfr As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim frcount As Integer
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
frcount = 0
db.Open OdkCnnString
                        
If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
Mindex = 51
End If

Dim SQLSTR As String
SQLSTR = ""
SLNO = 1







mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""


'totaltrees = totaltrees
'adamage = animaldamage
'pdamage = leafpest
'ddamage = stempest
'dtrees = deadmissing
'wlogged = waterlog
'ndtrees = nutrient

SQLSTR = ""
    db.Execute "delete from tbltemp"
SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,deadmissing,nutrient,waterlog,leafpest,stempest,ANIMALDAMAGE) SELECT max(end), dcode, gcode, tcode,fcode, scanlocation, totaltrees,dtrees,ndtrees,wlogged,pdamage,ddamage,adamage FROM storagehub6_core where scanlocation<>'' group  by scanlocation"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT  max(end) as end, dcode, gcode, tcode,fcode, scanlocation, totaltrees,dtrees,ndtrees,wlogged,pdamage,ddamage,adamage FROM storagehub6_core WHERE scanlocation ='' GROUP BY dcode, gcode, tcode, fcode", db
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
  
  rsF.Open "select * from tbltemp where farmercode='" & mfcode & "'", db
  If rsF.EOF <> True Then
    
  If rsF!end > rss!end Then
  db.Execute "update tbltemp set end='" & Format(rsF!end, "yyyy-MM-dd") & "' , totaltrees='" & rss!totaltrees & "',deadmissing='" & rss!dtrees & "',nutrient='" & rss!ndtrees & "',waterlog='" & rss!wlogged & "',leafpest='" & rss!pdamage & "',stempest='" & rss!ddamage & "',animaldamage='" & rss!adamage & "' where farmercode='" & mfcode & "'  "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,deadmissing,nutrient,waterlog,leafpest,stempest,animaldamage)values('" & Format(rss!end, "yyyy-MM-dd") & "','" & rss!dcode & "','" & rss!gcode & "','" & rss!tcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','" & rss!dtrees & "','" & rss!ndtrees & "','" & rss!wlogged & "','" & rss!pdamage & "','" & rss!ddamage & "','" & rss!adamage & "') "
  
  End If
  
  rss.MoveNext
  Loop





SQLSTR = "select count(farmercode) as frcount,'' as fieldcode,sum(totaltrees) as totaltrees,'' as area,sum(animaldamage) as adamage,sum(leafpest) as pdamage,sum(stempest) as ddamage,sum(deadmissing) as dtrees,sum(waterlog) as wlogged,sum(nutrient) as ndtrees from   tbltemp"
Set rsfr = Nothing


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
    excel_sheet.cells(1, 1) = UCase("Total No. of farmers in storaged")
    excel_sheet.cells(2, 1) = UCase("Total No. of trees in the storage")
    excel_sheet.cells(3, 1) = UCase("Total acres")
    excel_sheet.cells(4, 1) = UCase("Animal Damage")
    excel_sheet.cells(5, 1) = UCase("Pest Damage")
    excel_sheet.cells(6, 1) = UCase("Disease Damge ")
    excel_sheet.cells(7, 1) = UCase("Dead Trees")
    excel_sheet.cells(8, 1) = UCase("Waterlogged")
    excel_sheet.cells(9, 1) = UCase("Nutrient Deficient")
    
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.cells(1, 2) = rs!frcount
    excel_sheet.cells(2, 2) = rs!totaltrees
    excel_sheet.cells(3, 2) = rs![area]
    excel_sheet.cells(4, 2) = rs!adamage
    excel_sheet.cells(5, 2) = rs!pdamage
    excel_sheet.cells(6, 2) = rs!ddamage
    excel_sheet.cells(7, 2) = rs!dtrees
    excel_sheet.cells(8, 2) = rs!wlogged
    excel_sheet.cells(9, 2) = rs!ndtrees
  

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
 excel_app.selection.columnWidth = 35
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
db.Close
Exit Sub
err:
MsgBox err.Description
err.Clear
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
                        
If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
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
    db.Execute "delete from tbltemp"
SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,ANIMALDAMAGE) SELECT max(end), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage FROM phealthhub15_core where farmerbarcode<>'' group  by farmerbarcode,fdcode"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT max(END ) as end , dcode, gcode, tcode, fcode, farmerbarcode, (totaltrees),fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage FROM phealthhub15_core WHERE farmerbarcode ='' GROUP BY dcode, gcode, tcode, fcode,fdcode", db
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
  
  rsF.Open "select * from tbltemp where farmercode='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
  If rsF.EOF <> True Then
    
  If rsF!end > rss!end Then
  db.Execute "update tbltemp set end='" & Format(rsF!end, "yyyy-MM-dd") & "' , totaltrees='" & rss!totaltrees & "',area='" & rss!area & "',slowgrowing='" & rss!slowgrowing & "',dor='" & rss!dor & "',deadmissing='" & rss!deadmissing & "',activegrowing='" & rss!activegrowing & "',shock='" & rss!shock & "',nutrient='" & rss!nutrient & "',waterlog='" & rss!waterlog & "',leafpest='" & rss!leafpest & "',activepest='" & rss!activepest & "',stempest='" & rss!stempest & "',rootpest='" & rss!rootpest & "',animaldamage='" & rss!animaldamage & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage)values('" & Format(rss!end, "yyyy-MM-dd") & "','" & rss!dcode & "','" & rss!gcode & "','" & rss!tcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "','" & rss!area & "','" & rss!slowgrowing & "','" & rss!dor & "','" & rss!deadmissing & "','" & rss!activegrowing & "','" & rss!shock & "','" & rss!nutrient & "','" & rss!waterlog & "','" & rss!leafpest & "','" & rss!activepest & "','" & rss!stempest & "','" & rss!rootpest & "','" & rss!animaldamage & "') "
  
  End If
  
  rss.MoveNext
  Loop
  




SQLSTR = "select sum(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(slowgrowing) as slowgrowing,sum(dor) as dor,sum(deadmissing) as dead,sum(activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(leafpest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   tbltemp where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"


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
    excel_sheet.cells(1, 1) = UCase("Total No. of hazelnut field")
    excel_sheet.cells(2, 1) = UCase("Total No. of trees in the field")
    excel_sheet.cells(3, 1) = UCase("Total acres")
    excel_sheet.cells(4, 1) = UCase("Slow growing")
    excel_sheet.cells(5, 1) = UCase("Dormant")
    excel_sheet.cells(6, 1) = UCase("Dead ")
    excel_sheet.cells(7, 1) = UCase("Active growing")
    excel_sheet.cells(8, 1) = UCase("Shock")
    excel_sheet.cells(9, 1) = UCase("Nutrient deficeint")
    excel_sheet.cells(10, 1) = UCase("Waterlog")
    excel_sheet.cells(11, 1) = UCase("Leafpest")
    excel_sheet.cells(12, 1) = UCase("Active pest")
    excel_sheet.cells(13, 1) = UCase("Stem pest")
    excel_sheet.cells(14, 1) = UCase("Root pest")
    excel_sheet.cells(15, 1) = UCase("Animal Damage")
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.cells(1, 2) = rs!fieldcode
    excel_sheet.cells(2, 2) = rs!totaltrees
    excel_sheet.cells(3, 2) = rs![area]
    excel_sheet.cells(4, 2) = rs!slowgrowing
    excel_sheet.cells(5, 2) = rs!dor
    excel_sheet.cells(6, 2) = rs!dead
    excel_sheet.cells(7, 2) = rs!activegrowing
    excel_sheet.cells(8, 2) = rs!shock
    excel_sheet.cells(9, 2) = rs!nutrient
    excel_sheet.cells(10, 2) = rs!waterlog
    excel_sheet.cells(11, 2) = rs!leafpest
    excel_sheet.cells(12, 2) = rs!activepest
    excel_sheet.cells(13, 2) = rs!stempest
    excel_sheet.cells(14, 2) = rs!rootpest
    excel_sheet.cells(15, 2) = rs!animaldamage

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
 excel_app.selection.columnWidth = 31
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
db.Close
Exit Sub
err:
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
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                       
If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
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
    db.Execute "delete from tbltemp"
SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,ANIMALDAMAGE) SELECT max(end), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage FROM phealthhub15_core where farmerbarcode<>'' group  by farmerbarcode,fdcode"
  db.Execute SQLSTR
  Set rss = Nothing
  rss.Open "SELECT max(END ) as end , dcode, gcode, tcode, fcode, farmerbarcode, (totaltrees),fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage FROM phealthhub15_core WHERE farmerbarcode ='' GROUP BY dcode, gcode, tcode, fcode,fdcode", db
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
  
  rsF.Open "select * from tbltemp where farmercode='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
  If rsF.EOF <> True Then
    
  If rsF!end > rss!end Then
  db.Execute "update tbltemp set end='" & Format(rsF!end, "yyyy-MM-dd") & "' , totaltrees='" & rss!totaltrees & "',area='" & rss!area & "',slowgrowing='" & rss!slowgrowing & "',dor='" & rss!dor & "',deadmissing='" & rss!deadmissing & "',activegrowing='" & rss!activegrowing & "',shock='" & rss!shock & "',nutrient='" & rss!nutrient & "',waterlog='" & rss!waterlog & "',leafpest='" & rss!leafpest & "',activepest='" & rss!activepest & "',stempest='" & rss!stempest & "',rootpest='" & rss!rootpest & "',animaldamage='" & rss!animaldamage & "' where farmercode='" & mfcode & "' and fs='F' and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,area,slowgrowing,dor,deadmissing,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest,rootpest,animaldamage)values('" & Format(rss!end, "yyyy-MM-dd") & "','" & rss!dcode & "','" & rss!gcode & "','" & rss!tcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "','" & rss!area & "','" & rss!slowgrowing & "','" & rss!dor & "','" & rss!deadmissing & "','" & rss!activegrowing & "','" & rss!shock & "','" & rss!nutrient & "','" & rss!waterlog & "','" & rss!leafpest & "','" & rss!activepest & "','" & rss!stempest & "','" & rss!rootpest & "','" & rss!animaldamage & "') "
  
  End If
  
  rss.MoveNext
  Loop
  





















SQLSTR = "select count(fdcode) as fieldcode,sum(totaltrees) as totaltrees,sum(area) as area,sum(slowgrowing) as slowgrowing,sum(dor) as dor,sum(deadmissing) as dead,sum(activegrowing) as activegrowing,sum(shock) as shock,sum(nutrient) as nutrient,sum(waterlog) as waterlog,sum(leafpest) as leafpest,sum(activepest) as activepest,sum(stempest) as stempest,sum(rootpest) as rootpest,sum(animaldamage) as animaldamage from   tbltemp"


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
    excel_sheet.cells(1, 1) = UCase("Total No. of hazelnut field")
    excel_sheet.cells(2, 1) = UCase("Total No. of trees in the field")
    excel_sheet.cells(3, 1) = UCase("Total acres")
    excel_sheet.cells(4, 1) = UCase("Slow growing")
    excel_sheet.cells(5, 1) = UCase("Dormant")
    excel_sheet.cells(6, 1) = UCase("Dead ")
    excel_sheet.cells(7, 1) = UCase("Active growing")
    excel_sheet.cells(8, 1) = UCase("Shock")
    excel_sheet.cells(9, 1) = UCase("Nutrient deficeint")
    excel_sheet.cells(10, 1) = UCase("Waterlog")
    excel_sheet.cells(11, 1) = UCase("Leafpest")
    excel_sheet.cells(12, 1) = UCase("Active pest")
    excel_sheet.cells(13, 1) = UCase("Stem pest")
    excel_sheet.cells(14, 1) = UCase("Root pest")
    excel_sheet.cells(15, 1) = UCase("Animal Damage")
    
    
    
    
    
   
    
  ' i = 4
  Set rs = Nothing
rs.Open SQLSTR, db



    excel_sheet.cells(1, 2) = rs!fieldcode
    excel_sheet.cells(2, 2) = rs!totaltrees
    excel_sheet.cells(3, 2) = rs![area]
    excel_sheet.cells(4, 2) = rs!slowgrowing
    excel_sheet.cells(5, 2) = rs!dor
    excel_sheet.cells(6, 2) = rs!dead
    excel_sheet.cells(7, 2) = rs!activegrowing
    excel_sheet.cells(8, 2) = rs!shock
    excel_sheet.cells(9, 2) = rs!nutrient
    excel_sheet.cells(10, 2) = rs!waterlog
    excel_sheet.cells(11, 2) = rs!leafpest
    excel_sheet.cells(12, 2) = rs!activepest
    excel_sheet.cells(13, 2) = rs!stempest
    excel_sheet.cells(14, 2) = rs!rootpest
    excel_sheet.cells(15, 2) = rs!animaldamage

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
 excel_app.selection.columnWidth = 31
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
                       


























'If OPTALL.Value = True And OPTALLSTORAGE.Value = True Then
'Mindex = 48
'End If
'
'Dim SQLSTR As String
'SQLSTR = ""
'SLNO = 1
'If OPTALL.Value = True Then
'If CHKMOREOPTION.Value = 0 Then
'SQLSTR = "select * from storagehub6_core order by end "
'Else
'
'
'If optmoist.Value = True Then
'SQLSTR = "select * from storagehub6_core where (pmoisture/ttally)*100>'" & Val(txtValue.Text) & "' ORDER BY  end"
'
'ElseIf optpestdamage.Value = True Then
'SQLSTR = "select * from storagehub6_core where (pdamage/totaltrees)*100>'" & Val(txtValue.Text) & "' ORDER BY end"
'
'Else
'SQLSTR = "select * from storagehub6_core where (dtrees/totaltrees)*100>'" & Val(txtValue.Text) & "' ORDER BY end"
'End If
'
'End If
'
'Else
'If CHKMOREOPTION.Value = 0 Then
'SQLSTR = "select * from storagehub6_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' ORDER BY " & CBODATE.Text & ""
'
'Else
'
'If optmoist.Value = True Then
'SQLSTR = "select * from storagehub6_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (pmoisture/ttally)*100>'" & Val(txtValue.Text) & "' ORDER BY " & CBODATE.Text & ""
'ElseIf optpestdamage.Value = True Then
'SQLSTR = "select * from storagehub6_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (pdamage/totaltrees)*100>'" & Val(txtValue.Text) & "' ORDER BY " & CBODATE.Text & ""
'
'
'Else
'SQLSTR = "select * from storagehub6_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (dtrees/totaltrees)*100>'" & Val(txtValue.Text) & "' ORDER BY " & CBODATE.Text & ""
'End If
'
'
'End If
'End If


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
    'excel_app.Visible = False
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
    If OPTALL.Value = True And OPTALLSTORAGE.Value = True Then
    excel_sheet.cells(3, 2) = "DATE" & "(END)"
    Else
    excel_sheet.cells(3, 2) = "DATE" & "(" & CBODATE.Text & ")"
    End If
    excel_sheet.cells(3, 3) = "STAFF CODE-NAME"
    excel_sheet.cells(3, 4) = "D"
    excel_sheet.cells(3, 5) = "G"
    excel_sheet.cells(3, 6) = "T"
    excel_sheet.cells(3, 7) = UCase("Farmer ID")
    excel_sheet.cells(3, 8) = UCase("Total Distributed")
    excel_sheet.cells(3, 9) = UCase("Field ID")
    excel_sheet.cells(3, 10) = UCase("Total Trees Distributed - Planted List")
    excel_sheet.cells(3, 11) = UCase("Total Trees")
    excel_sheet.cells(3, 12) = UCase("Good Moisture")
    excel_sheet.cells(3, 13) = UCase("Poor Moisture")
    excel_sheet.cells(3, 14) = UCase("Total Mositure Tally")
    excel_sheet.cells(3, 15) = UCase("Dead Missing")
    excel_sheet.cells(3, 16) = UCase("Slow Growing")
    excel_sheet.cells(3, 17) = UCase("Dormant")
    excel_sheet.cells(3, 18) = UCase("Active Growing")
    excel_sheet.cells(3, 19) = UCase("Shock")
    excel_sheet.cells(3, 20) = UCase("Nutrient Deficient")
    excel_sheet.cells(3, 21) = UCase("Water Logg")
    excel_sheet.cells(3, 22) = UCase("pest damage")
    excel_sheet.cells(3, 23) = UCase("Active Pest")
    excel_sheet.cells(3, 24) = UCase("Stem Pest")
    excel_sheet.cells(3, 25) = UCase("Root Pest")
    excel_sheet.cells(3, 26) = UCase("Animal Damage")
    excel_sheet.cells(3, 27) = UCase("comments")
   i = 4
  Set rs = Nothing
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
mchk = True
excel_sheet.cells(i, 1) = SLNO

excel_sheet.cells(i, 2) = "'" & rs.Fields(Mindex)
excel_sheet.cells(i, 3) = rs!id
'If Len(rs!dcode) = 1 Then
'mdcode = "D0" & rs!dcode
'Else
'mdcode = "D" & rs!dcode
'End If
'
'
'If Len(rs!gcode) = 1 Then
'mgcode = "G0" & rs!gcode
'Else
'mgcode = "G" & rs!gcode
'End If
'
'If Len(rs!tcode) = 1 Then
'mtcode = "T0" & rs!tcode
'Else
'mtcode = "T" & rs!tcode
'End If
'
'If Len(rs!fcode) = 1 Then
'mfcode = mdcode & mgcode & mtcode & "F000" & rs!fcode
'ElseIf Len(rs!fcode) = 2 Then
'mfcode = mdcode & mgcode & mtcode & "F00" & rs!fcode
'ElseIf Len(rs!fcode) = 3 Then
'mfcode = mdcode & mgcode & mtcode & "F0" & rs!fcode
'Else
'mfcode = mdcode & mgcode & mtcode & "F" & rs!fcode
'End If
'
'If mdcode = "D00" Then
'excel_sheet.Cells(i, 1).Font.Color = vbRed
'excel_sheet.Cells(i, 2).Font.Color = vbRed
'excel_sheet.Cells(i, 3).Font.Color = vbRed
'excel_sheet.Cells(i, 4).Font.Color = vbRed
'excel_sheet.Cells(i, 5).Font.Color = vbRed
'excel_sheet.Cells(i, 6).Font.Color = vbRed
'excel_sheet.Cells(i, 7).Font.Color = vbRed
'excel_sheet.Cells(i, 1).Font.Bold = True
'excel_sheet.Cells(i, 2).Font.Bold = True
'excel_sheet.Cells(i, 3).Font.Bold = True
'excel_sheet.Cells(i, 4).Font.Bold = True
'excel_sheet.Cells(i, 5).Font.Bold = True
'excel_sheet.Cells(i, 6).Font.Bold = True
'excel_sheet.Cells(i, 7).Font.Bold = True
'excel_sheet.Cells(i, 4) = mdcode
'Else
'excel_sheet.Cells(i, 1).Font.Color = vbBlack
'excel_sheet.Cells(i, 2).Font.Color = vbBlack
'excel_sheet.Cells(i, 3).Font.Color = vbBlack
'excel_sheet.Cells(i, 4).Font.Color = vbBlack
'excel_sheet.Cells(i, 5).Font.Color = vbBlack
'excel_sheet.Cells(i, 6).Font.Color = vbBlack
'excel_sheet.Cells(i, 7).Font.Color = vbBlack
'excel_sheet.Cells(i, 1).Font.Bold = False
'excel_sheet.Cells(i, 2).Font.Bold = False
'excel_sheet.Cells(i, 3).Font.Bold = False
'excel_sheet.Cells(i, 4).Font.Bold = False
'excel_sheet.Cells(i, 5).Font.Bold = False
'excel_sheet.Cells(i, 6).Font.Bold = False
'excel_sheet.Cells(i, 7).Font.Bold = False
'excel_sheet.Cells(i, 4) = mdcode
'End If
'
'
'excel_sheet.Cells(i, 5) = mgcode
'excel_sheet.Cells(i, 6) = mtcode
'If Len(mfcode) <> 14 Then
'excel_sheet.Cells(i, 1).Font.Color = vbRed
'excel_sheet.Cells(i, 2).Font.Color = vbRed
'excel_sheet.Cells(i, 3).Font.Color = vbRed
'excel_sheet.Cells(i, 4).Font.Color = vbRed
'excel_sheet.Cells(i, 5).Font.Color = vbRed
'excel_sheet.Cells(i, 6).Font.Color = vbRed
'excel_sheet.Cells(i, 7).Font.Color = vbRed
'excel_sheet.Cells(i, 1).Font.Bold = True
'excel_sheet.Cells(i, 2).Font.Bold = True
'excel_sheet.Cells(i, 3).Font.Bold = True
'excel_sheet.Cells(i, 4).Font.Bold = True
'excel_sheet.Cells(i, 5).Font.Bold = True
'excel_sheet.Cells(i, 6).Font.Bold = True
'excel_sheet.Cells(i, 7).Font.Bold = True
'excel_sheet.Cells(i, 7) = mfcode
'ElseIf Len(mfcode) = 14 And mdcode <> "D00" Then
'excel_sheet.Cells(i, 1).Font.Color = vbBlack
'excel_sheet.Cells(i, 2).Font.Color = vbBlack
'excel_sheet.Cells(i, 3).Font.Color = vbBlack
'excel_sheet.Cells(i, 4).Font.Color = vbBlack
'excel_sheet.Cells(i, 5).Font.Color = vbBlack
'excel_sheet.Cells(i, 6).Font.Color = vbBlack
'excel_sheet.Cells(i, 7).Font.Color = vbBlack
'excel_sheet.Cells(i, 1).Font.Bold = False
'excel_sheet.Cells(i, 2).Font.Bold = False
'excel_sheet.Cells(i, 3).Font.Bold = False
'excel_sheet.Cells(i, 4).Font.Bold = False
'excel_sheet.Cells(i, 5).Font.Bold = False
'excel_sheet.Cells(i, 6).Font.Bold = False
'excel_sheet.Cells(i, 7).Font.Bold = False
'excel_sheet.Cells(i, 7) = mfcode
'End If


If rs!scanlocation = "" Then

If Len(rs!dcode) = 1 Then
mdcode = "D0" & rs!dcode
Else
mdcode = "D" & rs!dcode
End If


If Len(rs!gcode) = 1 Then
mgcode = "G0" & rs!gcode
Else
mgcode = "G" & rs!gcode
End If

If Len(rs!tcode) = 1 Then
mtcode = "T0" & rs!tcode
Else
mtcode = "T" & rs!tcode
End If

If Len(rs!fcode) = 1 Then
mfcode = mdcode & mgcode & mtcode & "F000" & rs!fcode
ElseIf Len(rs!fcode) = 2 Then
mfcode = mdcode & mgcode & mtcode & "F00" & rs!fcode
ElseIf Len(rs!fcode) = 3 Then
mfcode = mdcode & mgcode & mtcode & "F0" & rs!fcode
Else
mfcode = mdcode & mgcode & mtcode & "F" & rs!fcode
End If

If mdcode = "D00" Then
excel_sheet.cells(i, 1).Font.Color = vbRed
excel_sheet.cells(i, 2).Font.Color = vbRed
excel_sheet.cells(i, 3).Font.Color = vbRed
excel_sheet.cells(i, 4).Font.Color = vbRed
excel_sheet.cells(i, 5).Font.Color = vbRed
excel_sheet.cells(i, 6).Font.Color = vbRed
excel_sheet.cells(i, 7).Font.Color = vbRed
excel_sheet.cells(i, 1).Font.Bold = True
excel_sheet.cells(i, 2).Font.Bold = True
excel_sheet.cells(i, 3).Font.Bold = True
excel_sheet.cells(i, 4).Font.Bold = True
excel_sheet.cells(i, 5).Font.Bold = True
excel_sheet.cells(i, 6).Font.Bold = True
excel_sheet.cells(i, 7).Font.Bold = True
Else
excel_sheet.cells(i, 1).Font.Bold = False
excel_sheet.cells(i, 2).Font.Bold = False
excel_sheet.cells(i, 3).Font.Bold = False
excel_sheet.cells(i, 4).Font.Bold = False
excel_sheet.cells(i, 5).Font.Bold = False
excel_sheet.cells(i, 6).Font.Bold = False
excel_sheet.cells(i, 7).Font.Bold = False
excel_sheet.cells(i, 1).Font.Color = vbBlue
excel_sheet.cells(i, 2).Font.Color = vbBlue
excel_sheet.cells(i, 3).Font.Color = vbBlue
excel_sheet.cells(i, 4).Font.Color = vbBlue
excel_sheet.cells(i, 5).Font.Color = vbBlue
excel_sheet.cells(i, 6).Font.Color = vbBlue
excel_sheet.cells(i, 7).Font.Color = vbBlue
End If


excel_sheet.cells(i, 4) = mdcode
excel_sheet.cells(i, 5) = mgcode
excel_sheet.cells(i, 6) = mtcode
excel_sheet.cells(i, 7) = mfcode

Else
If Len(rs!scanlocation) <> 14 Then

excel_sheet.cells(i, 1).Font.Bold = True
excel_sheet.cells(i, 2).Font.Bold = True
excel_sheet.cells(i, 3).Font.Bold = True
excel_sheet.cells(i, 4).Font.Bold = True
excel_sheet.cells(i, 5).Font.Bold = True
excel_sheet.cells(i, 6).Font.Bold = True
excel_sheet.cells(i, 7).Font.Bold = True
excel_sheet.cells(i, 1).Font.Color = vbRed
excel_sheet.cells(i, 2).Font.Color = vbRed
excel_sheet.cells(i, 3).Font.Color = vbRed
excel_sheet.cells(i, 4).Font.Color = vbRed
excel_sheet.cells(i, 5).Font.Color = vbRed
excel_sheet.cells(i, 6).Font.Color = vbRed
excel_sheet.cells(i, 7).Font.Color = vbRed
excel_sheet.cells(i, 4) = Mid(rs!scanlocation, 1, 3)
excel_sheet.cells(i, 5) = Mid(rs!scanlocation, 4, 3)
excel_sheet.cells(i, 6) = Mid(rs!scanlocation, 7, 3)
excel_sheet.cells(i, 7) = IIf(IsNull(rs!scanlocation), "PLEASE CHECK FARMER CODE", rs!scanlocation)
Else





excel_sheet.cells(i, 1).Font.Bold = False
excel_sheet.cells(i, 2).Font.Bold = False
excel_sheet.cells(i, 3).Font.Bold = False
excel_sheet.cells(i, 4).Font.Bold = False
excel_sheet.cells(i, 5).Font.Bold = False
excel_sheet.cells(i, 6).Font.Bold = False
excel_sheet.cells(i, 7).Font.Bold = False
excel_sheet.cells(i, 1).Font.Color = vbBlack
excel_sheet.cells(i, 2).Font.Color = vbBlack
excel_sheet.cells(i, 3).Font.Color = vbBlack
excel_sheet.cells(i, 4).Font.Color = vbBlack
excel_sheet.cells(i, 5).Font.Color = vbBlack
excel_sheet.cells(i, 6).Font.Color = vbBlack
excel_sheet.cells(i, 7).Font.Color = vbBlack
excel_sheet.cells(i, 4) = Mid(rs!scanlocation, 1, 3)
excel_sheet.cells(i, 5) = Mid(rs!scanlocation, 4, 3)
excel_sheet.cells(i, 6) = Mid(rs!scanlocation, 7, 3)
excel_sheet.cells(i, 7) = IIf(IsNull(rs!scanlocation), "PLEASE CHECK FARMER CODE", rs!scanlocation)

End If

End If

FindFA excel_sheet.cells(i, 7), "F"
If chkred = True Then
                                
                               excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbRed
                                End If
                                
           chkred = False
           mchk = False

excel_sheet.cells(i, 8) = ""
excel_sheet.cells(i, 9) = ""
excel_sheet.cells(i, 10) = ""
excel_sheet.cells(i, 11) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.cells(i, 12) = IIf(IsNull(rs!gmoisture), "", rs!gmoisture)
excel_sheet.cells(i, 13) = IIf(IsNull(rs!pmoisture), "", rs!pmoisture)
excel_sheet.cells(i, 14) = IIf(IsNull(rs!ttally), "", rs!ttally)
excel_sheet.cells(i, 15) = IIf(IsNull(rs!dtrees), "", rs!dtrees)
excel_sheet.cells(i, 16) = "" 'IIf(IsNull(rs!slowgrowing), "", rs!slowgrowing)
excel_sheet.cells(i, 17) = "" 'IIf(IsNull(rs!dor), "", rs!dor)
excel_sheet.cells(i, 18) = "" 'IIf(IsNull(rs!activegrowing), "", rs!activegrowing)
excel_sheet.cells(i, 19) = "" 'IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.cells(i, 20) = IIf(IsNull(rs!ndtrees), "", rs!ndtrees)
excel_sheet.cells(i, 21) = IIf(IsNull(rs!wlogged), "", rs!wlogged)
excel_sheet.cells(i, 22) = IIf(IsNull(rs!pdamage), "", rs!pdamage)
excel_sheet.cells(i, 23) = "" 'IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.cells(i, 24) = "" 'IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.cells(i, 25) = "" 'IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.cells(i, 26) = IIf(IsNull(rs!adamage), "", rs!adamage)
excel_sheet.cells(i, 27) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)




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
                     
db.Execute "delete from tbltemp"
If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
Mindex = 51
End If

Dim SQLSTR As String
SQLSTR = ""
SLNO = 1



SQLSTR = ""

If OPTALL.Value = True Then

SQLSTR = "insert into tbltemp (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,deadmissing,slowgrowing,dor,activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select end,id,dcode," _
         & "gcode,tcode,farmerbarcode,treesreceived,fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core"
  Else
  
  SQLSTR = "insert into tbltemp (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,deadmissing,slowgrowing,dor,activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select " & CBODATE.Text & ",id,dcode," _
         & "gcode,tcode,farmerbarcode,treesreceived,fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core"
  
  
  End If
         
         
         
         
   SQLSTR = SQLSTR & "  where farmerbarcode<>''"
   
  db.Execute SQLSTR
  Set rss = Nothing
  SQLSTR = ""
  
  
  If OPTALL.Value = True Then
  SQLSTR = " select end as end,id,dcode," _
         & "gcode,tcode,fcode,farmerbarcode,treesreceived,fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core"
         Else
         
           SQLSTR = " select " & CBODATE.Text & " as end,id,dcode," _
         & "gcode,tcode,fcode,farmerbarcode,treesreceived,fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core"
         
         End If
         
         
    Set rss = Nothing
  rss.Open SQLSTR & "  WHERE farmerbarcode ='' ", db
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
'  If mfcode = "D04G02T04F0099" Then
'
'  MsgBox "sdfsdf"
'  End If
  Set rsF = Nothing
  
  rsF.Open "select * from tbltemp where farmercode='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
  If rsF.EOF <> True Then
    
  If rss!end > rsF!end Then
'  db.Execute "update tbltemp set end='" & Format(rss!End, "yyyy-MM-dd") & "' , totaltrees='" & rss!totaltrees & "',id='" & rss!id & "',treesreceived='" & rss!treesreceived & "'," _
'            & "goodmoisture='" & rss!goodmoisture & "',poormoisture='" & rss!poormoisture & "'," _
'            & "totaltally='" & rss!totaltally & "',deadmissing='" & rss!deadmissing & "',slowgrowing='" & rss!slowgrowing & "' ," _
'            & "dor='" & rss!dor & "',activegrowing='" & rss!activegrowing & "'," _
'            & "shock='" & rss!shock & "',nutrient='" & rss!nutrient & "',waterlog='" & rss!waterlog & "'," _
'            & "leafpest='" & rss!leafpest & "',activepest='" & rss!activepest & "',stempest='" & rss!stempest & "'," _
'            & "rootpest='" & rss!rootpest & "',animaldamage='" & rss!animaldamage & "'," _
'            & "monitorcomments='" & rss!monitorcomments & "' where farmercode='" & mfcode & "'  and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into tbltemp (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,deadmissing,slowgrowing,dor,activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) values( '" & Format(rss!end, "yyyy-MM-dd") & "','" & rss!id & "'," _
         & "'" & rss!dcode & "'," _
         & "'" & rss!gcode & "','" & rss!tcode & "','" & mfcode & "','" & rss!treesreceived & "','" & rss!FDCODE & "'," _
         & "'" & rss!totaltrees & "','" & rss!goodmoisture & "','" & rss!poormoisture & "','" & rss!totaltally & "'," _
         & "'" & rss!deadmissing & "','" & rss!slowgrowing & "','" & rss!dor & "','" & rss!activegrowing & "'," _
         & "'" & rss!shock & "','" & rss!nutrient & "','" & rss!waterlog & "','" & rss!leafpest & "','" & rss!activepest & "','" & rss!stempest & "'," _
         & "'" & rss!rootpest & "','" & rss!animaldamage & "','" & rss!monitorcomments & "')"
  
  End If
  
  rss.MoveNext
  Loop
         
         
         
         
         
         

' need to rectify the query later


If OPTALL.Value = True Then
If CHKMOREOPTION.Value = 0 Then
SQLSTR = "select * from tbltemp"
Else
If optmoist.Value = True Then

SQLSTR = "select * from tbltemp where (poormoisture/totaltally)*100>'" & Val(TXTVALUE.Text) & "' order by end"
ElseIf optrootpest.Value = True Then
SQLSTR = "select * from tbltemp where (rootpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' order by end"
ElseIf OPTSTEMPEST.Value = True Then
SQLSTR = "select * from tbltemp where (stempest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' order by end"
ElseIf OPTLEAFPEST.Value = True Then

SQLSTR = "select * from tbltemp where (leafpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' order by end"
Else
SQLSTR = "select * from tbltemp where (deadmissing/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' order by end"
End If
End If
Else
'"select * from dailyacthub9_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' "
If CHKMOREOPTION.Value = 0 Then
SQLSTR = "select * from tbltemp where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' ORDER BY end"
Else

If optmoist.Value = True Then

SQLSTR = "select * from tbltemp where end='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and end<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (poormoisture/totaltally)*100>'" & Val(TXTVALUE.Text) & "' ORDER BY end"
ElseIf optrootpest.Value = True Then
SQLSTR = "select * from tbltemp where end>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and end<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (rootpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' ORDER BY end"
ElseIf OPTSTEMPEST.Value = True Then
SQLSTR = "select * from tbltemp where end>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and end<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (stempest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' ORDER BY end"
ElseIf OPTLEAFPEST.Value = True Then
SQLSTR = "select * from tbltemp where end>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and end<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (leafpest/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' ORDER BY end"
Else

SQLSTR = "select * from tbltemp where end>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and end<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and (deadmissing/totaltrees)*100>'" & Val(TXTVALUE.Text) & "' ORDER BY end"
End If

End If
End If






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
    excel_sheet.cells(3, 1) = "SL.NO."
    
    If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
    excel_sheet.cells(3, 2) = "DATE" & "(END)"
    Else
    excel_sheet.cells(3, 2) = "DATE" & "(" & CBODATE.Text & ")"
    End If
     excel_sheet.cells(3, 3) = "STAFF CODE-NAME"
    excel_sheet.cells(3, 4) = "D"
    excel_sheet.cells(3, 5) = "G"
    excel_sheet.cells(3, 6) = "T"
    excel_sheet.cells(3, 7) = UCase("Farmer ID")
    excel_sheet.cells(3, 8) = UCase("Total Distributed")
    excel_sheet.cells(3, 9) = UCase("Field ID")
    excel_sheet.cells(3, 10) = UCase("Total Trees Distributed - Planted List")
    excel_sheet.cells(3, 11) = UCase("Total Trees")
    excel_sheet.cells(3, 12) = UCase("Good Moisture")
    excel_sheet.cells(3, 13) = UCase("Poor Moisture")
    excel_sheet.cells(3, 14) = UCase("Total Mositure Tally")
    excel_sheet.cells(3, 15) = UCase("Dead Missing")
    excel_sheet.cells(3, 16) = UCase("Slow Growing")
    excel_sheet.cells(3, 17) = UCase("Dormant")
    excel_sheet.cells(3, 18) = UCase("Active Growing")
    excel_sheet.cells(3, 19) = UCase("Shock")
    excel_sheet.cells(3, 20) = UCase("Nutrient Deficient")
    excel_sheet.cells(3, 21) = UCase("Water Logg")
    excel_sheet.cells(3, 22) = UCase("Leaf Pest")
    excel_sheet.cells(3, 23) = UCase("Active Pest")
    excel_sheet.cells(3, 24) = UCase("Stem Pest")
    excel_sheet.cells(3, 25) = UCase("Root Pest")
    excel_sheet.cells(3, 26) = UCase("Animal Damage")
    excel_sheet.cells(3, 27) = UCase("comments")
   i = 4
  Set rs = Nothing
  
  'SQLSTR = ""
  ' 'D01G09T04F0097','D01G09T04F0033','D01G05T03F0189','D01G01T03F0096','D01G05T03F0180','D01G04T01F0038','D02G06T04F0054','D02G02T03F0006','D02G06T03F0226','D03G11T02F004,'D03G02T05F0058','D03G01T04F0275','D03G11T01F0053','D04G08T02F0014','D04G02T04F0099','D04G05T02F0084','D05G04T03F0113','D05G04T02F0040','D05G01T02F0012','D06G01T03F0031'
  'SQLSTR = "select * from tbltemp where farmercode in('D01G09T04F0097','D01G09T04F0033','D01G05T03F0189','D01G01T03F0096','D01G05T03F0180','D01G04T01F0038','D02G06T04F0054','D02G02T03F0006','D02G06T03F0226','D03G11T02F0040','D03G02T05F0058','D03G01T04F0275','D03G11T01F0053','D04G08T02F0014','D04G02T04F0099','D04G05T02F0084','D05G04T03F0113','D05G04T02F0040','D05G01T02F0012','D06G01T03F0031') order by farmercode"
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = "'" & rs!end  'rs.Fields(Mindex)
excel_sheet.cells(i, 3) = rs!id
'If rs!farmerbarcode = "" Then
'
'If Len(rs!dcode) = 1 Then
'mdcode = "D0" & rs!dcode
'Else
'mdcode = "D" & rs!dcode
'End If
'
'
'If Len(rs!gcode) = 1 Then
'mgcode = "G0" & rs!gcode
'Else
'mgcode = "G" & rs!gcode
'End If
'
'If Len(rs!tcode) = 1 Then
'mtcode = "T0" & rs!tcode
'Else
'mtcode = "T" & rs!tcode
'End If
'
'If Len(rs!fcode) = 1 Then
'mfcode = mdcode & mgcode & mtcode & "F000" & rs!fcode
'ElseIf Len(rs!fcode) = 2 Then
'mfcode = mdcode & mgcode & mtcode & "F00" & rs!fcode
'ElseIf Len(rs!fcode) = 3 Then
'mfcode = mdcode & mgcode & mtcode & "F0" & rs!fcode
'Else
'mfcode = mdcode & mgcode & mtcode & "F" & rs!fcode
'End If
'
'If mdcode = "D00" Then
'excel_sheet.Cells(i, 1).Font.Color = vbRed
'excel_sheet.Cells(i, 2).Font.Color = vbRed
'excel_sheet.Cells(i, 3).Font.Color = vbRed
'excel_sheet.Cells(i, 4).Font.Color = vbRed
'excel_sheet.Cells(i, 5).Font.Color = vbRed
'excel_sheet.Cells(i, 6).Font.Color = vbRed
'excel_sheet.Cells(i, 7).Font.Color = vbRed
'excel_sheet.Cells(i, 1).Font.Bold = True
'excel_sheet.Cells(i, 2).Font.Bold = True
'excel_sheet.Cells(i, 3).Font.Bold = True
'excel_sheet.Cells(i, 4).Font.Bold = True
'excel_sheet.Cells(i, 5).Font.Bold = True
'excel_sheet.Cells(i, 6).Font.Bold = True
'excel_sheet.Cells(i, 7).Font.Bold = True
'Else
'excel_sheet.Cells(i, 1).Font.Bold = False
'excel_sheet.Cells(i, 2).Font.Bold = False
'excel_sheet.Cells(i, 3).Font.Bold = False
'excel_sheet.Cells(i, 4).Font.Bold = False
'excel_sheet.Cells(i, 5).Font.Bold = False
'excel_sheet.Cells(i, 6).Font.Bold = False
'excel_sheet.Cells(i, 7).Font.Bold = False
'excel_sheet.Cells(i, 1).Font.Color = vbBlue
'excel_sheet.Cells(i, 2).Font.Color = vbBlue
'excel_sheet.Cells(i, 3).Font.Color = vbBlue
'excel_sheet.Cells(i, 4).Font.Color = vbBlue
'excel_sheet.Cells(i, 5).Font.Color = vbBlue
'excel_sheet.Cells(i, 6).Font.Color = vbBlue
'excel_sheet.Cells(i, 7).Font.Color = vbBlue
'End If
'
'
'excel_sheet.Cells(i, 4) = mdcode
'excel_sheet.Cells(i, 5) = mgcode
'excel_sheet.Cells(i, 6) = mtcode
'excel_sheet.Cells(i, 7) = mfcode
'
'Else
'excel_sheet.Cells(i, 1).Font.Bold = False
'excel_sheet.Cells(i, 2).Font.Bold = False
'excel_sheet.Cells(i, 3).Font.Bold = False
'excel_sheet.Cells(i, 4).Font.Bold = False
'excel_sheet.Cells(i, 5).Font.Bold = False
'excel_sheet.Cells(i, 6).Font.Bold = False
'excel_sheet.Cells(i, 7).Font.Bold = False
'excel_sheet.Cells(i, 1).Font.Color = vbBlack
'excel_sheet.Cells(i, 2).Font.Color = vbBlack
'excel_sheet.Cells(i, 3).Font.Color = vbBlack
'excel_sheet.Cells(i, 4).Font.Color = vbBlack
'excel_sheet.Cells(i, 5).Font.Color = vbBlack
'excel_sheet.Cells(i, 6).Font.Color = vbBlack
'excel_sheet.Cells(i, 7).Font.Color = vbBlack
'excel_sheet.Cells(i, 4) = Mid(rs!farmerbarcode, 1, 3)
'excel_sheet.Cells(i, 5) = Mid(rs!farmerbarcode, 4, 3)
'excel_sheet.Cells(i, 6) = Mid(rs!farmerbarcode, 7, 3)
'excel_sheet.Cells(i, 7) = IIf(IsNull(rs!farmerbarcode), "PLEASE CHECK FARMER CODE", rs!farmerbarcode)
'End If
excel_sheet.cells(i, 4) = Mid(rs!farmercode, 1, 3)
excel_sheet.cells(i, 5) = Mid(rs!farmercode, 4, 3)
excel_sheet.cells(i, 6) = Mid(rs!farmercode, 7, 3)
excel_sheet.cells(i, 7) = IIf(IsNull(rs!farmercode), "", rs!farmercode)


excel_sheet.cells(i, 8) = IIf(IsNull(rs!treesreceived), "", rs!treesreceived)
excel_sheet.cells(i, 9) = IIf(IsNull(rs!FDCODE), "", rs!FDCODE)
excel_sheet.cells(i, 10) = ""
excel_sheet.cells(i, 11) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.cells(i, 12) = IIf(IsNull(rs!goodmoisture), "", rs!goodmoisture)
excel_sheet.cells(i, 13) = IIf(IsNull(rs!poormoisture), "", rs!poormoisture)
excel_sheet.cells(i, 14) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.cells(i, 15) = IIf(IsNull(rs!deadmissing), "", rs!deadmissing)
excel_sheet.cells(i, 16) = IIf(IsNull(rs!slowgrowing), "", rs!slowgrowing)
excel_sheet.cells(i, 17) = IIf(IsNull(rs!dor), "", rs!dor)
excel_sheet.cells(i, 18) = IIf(IsNull(rs!activegrowing), "", rs!activegrowing)
excel_sheet.cells(i, 19) = IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.cells(i, 20) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.cells(i, 21) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.cells(i, 22) = IIf(IsNull(rs!leafpest), "", rs!leafpest)
excel_sheet.cells(i, 23) = IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.cells(i, 24) = IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.cells(i, 25) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.cells(i, 26) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.cells(i, 27) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)


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
     excel_sheet.Range("A3:AA3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ALL FIELDS"
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
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear


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

CBODATE.AddItem rs.Fields(j).name
End If
Next

Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub findindex(tt As String, id As Integer)
Dim i, j, fcount As Integer

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                       
db.Open OdkCnnString
                        
Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & id & "' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)

Set rs = Nothing
rs.Open "SELECT * FROM " & tt & " where 1", CONNLOCAL
For j = 0 To fcount - 1
If rs.Fields(j).Type = 135 Then

If rs.Fields(j).name = CBODATE.Text Then
Mindex = j
Exit For
Else
Mindex = 2
End If


End If
Next

Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub Form_Load()
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
FRMALLFIELDS.Width = 5640
populatedate "phealthhub15_core", 11
End Sub

Private Sub Form_Unload(Cancel As Integer)
mchk = False
End Sub

Private Sub OPTALL_Click()
'Frame1.Enabled = False
'Frame3.Visible = False
'CHKMOREOPTION.Enabled = True
'Frame.Visible = True

If CHKSUMMARY.Value = 1 Then
Frame.Visible = False
Frame1.Enabled = False
Frame3.Visible = True
Else
Frame1.Enabled = False
CHKMOREOPTION.Enabled = True
Frame.Visible = True
Frame3.Visible = False
End If



End Sub

Private Sub OPTALLFIELDS_Click()
populatedate "phealthhub15_core", 11
End Sub

Private Sub OPTALLSTORAGE_Click()
populatedate "storagehub6_core", 17
End Sub

Private Sub optdead_Click()
TXTVALUE.Text = 20
End Sub

Private Sub OPTLEAFPEST_Click()
TXTVALUE.Text = 5
End Sub

Private Sub optmoist_Click()
TXTVALUE.Text = 30
End Sub

Private Sub optpestdamage_Click()
TXTVALUE.Text = 5
End Sub

Private Sub optrootpest_Click()
TXTVALUE.Text = 5
End Sub

Private Sub OPTSEL_Click()

If CHKSUMMARY.Value = 1 Then
Frame.Visible = False
Frame1.Enabled = True
Frame3.Visible = True
Else
Frame1.Enabled = True
CHKMOREOPTION.Enabled = True
Frame.Visible = True
Frame3.Visible = False
End If
End Sub

Private Sub OPTSUMMARY_Click()

End Sub

Private Sub OPTSTEMPEST_Click()
TXTVALUE.Text = 5
End Sub

Private Sub TXTVALUE_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
