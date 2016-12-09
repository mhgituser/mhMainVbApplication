VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMRPTLANDDETAILS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FARMER LAND DETAILS"
   ClientHeight    =   5205
   ClientLeft      =   945
   ClientTop       =   825
   ClientWidth     =   14100
   Icon            =   "FRMRPTLANDDETAILS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   14100
   Begin VB.CommandButton Command7 
      Caption         =   "odkerror"
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6360
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox LSTPR 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      ItemData        =   "FRMRPTLANDDETAILS.frx":0E42
      Left            =   7680
      List            =   "FRMRPTLANDDETAILS.frx":0E44
      Style           =   1  'Checkbox
      TabIndex        =   22
      Top             =   0
      Width           =   6015
   End
   Begin VB.CheckBox chkpririty 
      Caption         =   "PRIORITIZE  LIST"
      Height          =   195
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkdist 
      Caption         =   "DISTRIBUTION LIST"
      Height          =   195
      Left            =   3840
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton OPTWO 
      Caption         =   "WITHOUT ZERO ADDITIONAL LAND"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   840
      Width           =   3135
   End
   Begin VB.OptionButton OPTW 
      Caption         =   "WITH ZERO ADDITIONAL LAND"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   480
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "SELECTION TYPE"
      Height          =   1335
      Left            =   3720
      TabIndex        =   14
      Top             =   2280
      Width           =   3495
      Begin VB.OptionButton OPTWITHDate 
         Caption         =   "REGISTRATION BY DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton OPTWITHOUTDATE 
         Caption         =   "REGISTRATION WITHOUT DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   4335
      End
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "REPORT TYPE"
      Height          =   975
      Left            =   3720
      TabIndex        =   11
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton OPTSUMMARY 
         Caption         =   "SUMMARY"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OPTDETAIL 
         Caption         =   "DETAIL"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATE SELECTION"
      Height          =   1575
      Left            =   3720
      TabIndex        =   6
      Top             =   3600
      Width           =   3255
      Begin MSComCtl2.DTPicker TXTFROMDATE 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81526785
         CurrentDate     =   41208
      End
      Begin MSComCtl2.DTPicker TXTTODATE 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81526785
         CurrentDate     =   41208
      End
      Begin VB.Label Label2 
         Caption         =   "FROM DATE"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "TO DATE DATE"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
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
      Left            =   240
      Picture         =   "FRMRPTLANDDETAILS.frx":0E46
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
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
      Left            =   1920
      Picture         =   "FRMRPTLANDDETAILS.frx":15B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REVERSE SELECTION"
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
      Left            =   1920
      Picture         =   "FRMRPTLANDDETAILS.frx":227A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELECT ALL"
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
      Left            =   240
      Picture         =   "FRMRPTLANDDETAILS.frx":2B44
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
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
      Height          =   2535
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   7560
      X2              =   7560
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "DZONGKHAG SELECTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "FRMRPTLANDDETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DZstr As String
Private Sub chkpririty_Click()
DZstr = ""
If chkpririty.Value = 1 Then
FRMRPTLANDDETAILS.Width = 13980

For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       DZstr = DZstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
If Len(DZstr) > 0 Then
   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
 
Else
FRMRPTLANDDETAILS.Width = 7560
chkpririty.Value = 0
LSTPR.Clear
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTPR.Clear
rs.Open "select * from tbltshewog where dzongkhagid in " & DZstr & "order by dzongkhagid,gewogid,tshewogid", MHVDB, adOpenStatic
With rs
Do While Not .EOF
FindDZ rs!dzongkhagid
FindGE rs!dzongkhagid, rs!gewogid





   LSTPR.AddItem rs!dzongkhagid & " " & Trim(Dzname) + " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With


Else
FRMRPTLANDDETAILS.Width = 7560
End If




End Sub

Private Sub Command1_Click()
Dim i As Long
For i = 0 To DZLIST.ListCount - 1
    DZLIST.Selected(i) = True
Next
End Sub

Private Sub Command2_Click()
If DZLIST.ListCount > 0 Then
   For i = 0 To DZLIST.ListCount - 1
       If DZLIST.Selected(i) Then
          DZLIST.Selected(i) = False
       Else
          DZLIST.Selected(i) = True
       End If
   Next
End If
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
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode) SELECT max(END), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode FROM phealthhub15_core where farmerbarcode<>'' group  by farmerbarcode,fdcode"
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
  
  rsF.Open "select * from tbltemp where farmercode='" & mfcode & "' and fdcode='" & rss!FDCODE & "'", db
  If rsF.EOF <> True Then
    
  If rsF!End > rss!End Then
  db.Execute "update tbltemp set end='" & Format(rsF!End, "yyyy-MM-dd") & "' , totaltrees='" & rss!totaltrees & "' where farmercode='" & mfcode & "' and fs='F' and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode)values('" & Format(rss!End, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "') "
  
  End If
  
  rss.MoveNext
  Loop
  
  
  'storage
  SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode) SELECT max(END), dcode, gcode, tcode,fcode, scanlocation, (totaltrees),'S'  as fs,'' FROM storagehub6_core where scanlocation<>'' group  by scanlocation"
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
  
  rsF.Open "select * from tbltemp where farmercode='" & mfcode & "'", db
  If rsF.EOF <> True Then
    
  If rsF!End > rss!End Then
  db.Execute "update tbltemp set end='" & Format(rsF!End, "yyyy-MM-dd") & "' , totaltrees='" & rss!totaltrees & "' where farmercode='" & mfcode & "' and fs='S' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='S' "
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs)values('" & Format(rss!End, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','S') "
  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
     excel_sheet.Cells(3, 5) = "FARMER CODE"
    excel_sheet.Cells(3, 6) = "FAMER"
    excel_sheet.Cells(3, 7) = "REG. LAND (ACRE)"
    excel_sheet.Cells(3, 8) = "PLANTED(ACRE)"
    excel_sheet.Cells(3, 9) = "ACTUAL DISTRIBUTED"
    excel_sheet.Cells(3, 10) = "TREES(FIELD)"
    excel_sheet.Cells(3, 11) = "REES(STORAGE"
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
                            excel_sheet.Cells(i, 1) = sl
                            FindDZ Mid(rs!farmercode, 1, 3)
     excel_sheet.Cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
     FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
   excel_sheet.Cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
   FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
    excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
  FindFA rs!farmercode, "F"
  
  
    excel_sheet.Cells(i, 5) = rs!farmercode
  excel_sheet.Cells(i, 6) = FAName
  
  
  
  Set rsF = Nothing
  rsF.Open "SELECT farmerid,sum(regland) as rl  from tbllandreg   farmerid where farmerid='" & rs!farmercode & "'", MHVDB
  If rsF.EOF <> True Then
  excel_sheet.Cells(i, 7) = IIf(IsNull(rsF!rl), "", rsF!rl)
  
  End If
  
  
  
  
  Set rss = Nothing
  rss.Open "SELECT sum(acreplanted) as pl ,sum(nooftrees) as t FROM tblplanted where farmercode='" & rs!farmercode & "' GROUP BY farmercode", db

  If rss.EOF <> True Then
  
  excel_sheet.Cells(i, 8) = rss!PL
  excel_sheet.Cells(i, 9) = rss!T
  
  End If
  
  
 newchk = False
 fchk = False
  schk = False
  
  Set rss = Nothing
  rss.Open "SELECT max( end ) , dcode, gcode, tcode, fcode, farmercode, sum( totaltrees ) AS totaltrees, fs, fdcode FROM tbltemp where fs='F' and farmercode='" & rs!farmercode & "' GROUP BY farmercode", db

  If rss.EOF <> True Then
  
    excel_sheet.Cells(i, 10) = rss!totaltrees
    newchk = True
    
    
    
    Else
    
    excel_sheet.Cells(i, 10) = ""
'   excel_sheet.Range(excel_sheet.Cells(i, 1), _
'                             excel_sheet.Cells(i, 11)).Select
'                             excel_app.Selection.Font.Color = vbBlue

  fchk = True
    
  End If
  
 
   Set rss = Nothing
  rss.Open "SELECT max( end ) , dcode, gcode, tcode, fcode, farmercode, sum( totaltrees ) AS totaltrees, fs, fdcode FROM tbltemp where fs='S' and farmercode='" & rs!farmercode & "' GROUP BY farmercode", db

  If rss.EOF <> True Then
  
    excel_sheet.Cells(i, 11) = rss!totaltrees
    newchk = True
  
    
    Else
     excel_sheet.Cells(i, 11) = ""
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
  
     excel_sheet.Range(excel_sheet.Cells(i, 1), _
                             excel_sheet.Cells(i, 11)).Select
                             excel_app.Selection.Font.Color = vbRed
                  
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
                   excel_sheet.Cells(i, 11) = "Match=" & newmatched & "Not Match=" & newnotmatched
                            'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 11)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
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
Private Sub ldexcel()
Dim SQLSTR As String
Dim myphone As String
Dim TOTLAND As Double
Dim totadd As Double
TOTLAND = 0
totadd = 0
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Set rsadd = Nothing
j = 0

DZstr = ""
SQLSTR = ""


For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       DZstr = DZstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
If Len(DZstr) > 0 Then
   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
 
Else
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


If OPTDETAIL.Value = True Then


If optall.Value = True Then

SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID AND  SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & " ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

ElseIf OPTWITHOUTDATE.Value = True Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID AND REGDATE>='1900-01-01' AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & " ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Else
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID AND REGDATE>='" & Format(TXTFROMDATE.Value, "yyyy-MM-dd") & "' and regdate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & " ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

End If




Else



If optall.Value = True Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,phone1,phone2,SUM(REGLAND) as regland FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME,phone1,phone2 ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"


ElseIf OPTWITHOUTDATE.Value = True Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND-acreplanted)) AS aditionalland,sum(regland) as regland FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID and REGDATE>='1900-01-01' AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Else
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND)) AS aditionalland,sum(regland) as regland FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID and  REGDATE>='" & Format(TXTFROMDATE.Value, "yyyy-MM-dd") & "' and regdate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' AND  SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

End If




End If



                            
                        
                        
                        
                        

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
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
     excel_sheet.Cells(3, 5) = "FARMER CODE"
    excel_sheet.Cells(3, 6) = "FAMER"
    excel_sheet.Cells(3, 7) = "CONTACT NO."
    excel_sheet.Cells(3, 8) = "REG. LAND (ACRE)"
    excel_sheet.Cells(3, 9) = "ADITIONAL LAND (ACRE)"
    
      i = 4
                        
                        
                        
                        
                        
                            Set rs = Nothing
                            rs.Open SQLSTR, MHVDB
                            If rs.EOF <> True Then
                            Do While rs.EOF <> True
                            
                            excel_sheet.Cells(i, 1) = sl
                            FindDZ Mid(rs!idfarmer, 1, 3)
     excel_sheet.Cells(i, 2) = rs!dzcode & " " & Dzname
     FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
   excel_sheet.Cells(i, 3) = rs!GECODE & " " & GEname
   FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
    excel_sheet.Cells(i, 4) = rs!tscode & " " & TsName
  FindFA rs!idfarmer, "F"
  
  
    excel_sheet.Cells(i, 5) = rs!idfarmer
  excel_sheet.Cells(i, 6) = FAName
  myphone = IIf(IsNull(rs!phone1), "", rs!phone1) & "," & IIf(IsNull(rs!phone2), "", rs!phone2)
  If Len(myphone) > 0 Then
   myphone = Left(myphone, Len(myphone) - 1)
   End If
   
     excel_sheet.Cells(i, 7) = myphone
     
    excel_sheet.Cells(i, 8) = Format(IIf(IsNull(rs!regland), 0, rs!regland), "####0.00")
    
   
    
    Set rsadd = Nothing
    rsadd.Open "select sum(acreplanted)as acreplanted  from tblplanted where farmercode='" & rs!idfarmer & "'", MHVDB
    
    
     
  
    
    If rsadd.EOF <> True Then
    
    If OPTWO.Value = True Then
    If Format(IIf(IsNull(rs!regland), 0, rs!regland), "###0.000") - Format(IIf(IsNull(rsadd!acreplanted), 0, rsadd!acreplanted), "###0.000") <> 0 Then
    excel_sheet.Cells(i, 9) = Format(IIf(IsNull(rs!regland), 0, rs!regland) - IIf(IsNull(rsadd!acreplanted), 0, rsadd!acreplanted), "####0.00")
    
    totadd = totadd + Format(IIf(IsNull(rs!regland), 0, rs!regland) - IIf(IsNull(rsadd!acreplanted), 0, rsadd!acreplanted), "####0.00")
     TOTLAND = TOTLAND + IIf(IsNull(rs!regland), 0, rs!regland)
    i = i + 1
      sl = sl + 1
    Else
     'excel_sheet.Cells(i, 8) = ""
    End If
    Else
    
    
    excel_sheet.Cells(i, 9) = Format(IIf(IsNull(rs!regland), 0, rs!regland) - IIf(IsNull(rsadd!acreplanted), 0, rsadd!acreplanted), "####0.00")
    
    totadd = totadd + Format(IIf(IsNull(rs!regland), 0, rs!regland) - IIf(IsNull(rsadd!acreplanted), 0, rsadd!acreplanted), "####0.00")
    TOTLAND = TOTLAND + IIf(IsNull(rs!regland), 0, rs!regland)
    i = i + 1
      sl = sl + 1
    
    
    End If
    
    
    
     
    Else
   ' excel_sheet.Cells(i, 8) = ""
    End If
    
 
                            
                            rs.MoveNext
                            Loop
          excel_sheet.Cells(i, 6).Font.Bold = True
     excel_sheet.Cells(i, 6) = "TOTAL "
     If OPTWO.Value = True Then
           excel_sheet.Cells(i, 8).Font.Bold = True
     excel_sheet.Cells(i, 8) = Format(TOTLAND, "####0.00")
         excel_sheet.Cells(i, 9).Font.Bold = True
     excel_sheet.Cells(i, 9) = Format(totadd, "####0.00")
                   Else
                     excel_sheet.Cells(i, 8).Font.Bold = True
     excel_sheet.Cells(i, 8) = Format(TOTLAND, "####0.00")
         excel_sheet.Cells(i, 9).Font.Bold = True
     excel_sheet.Cells(i, 9) = Format(totadd, "####0.00")
                   End If
                            End If
                            
                            
                            'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "REGISTARTION FROM " & TXTFROMDATE.Value & "  TO " & TXTTODATE.Value
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

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If chkdist.Value = 0 Then
ldexcel
Else
DISTLIST
End If
End Sub
Private Sub DISTLIST()
'On Error Resume Next
Dim s As Integer
Dim SQLSTR As String
Dim totplant As Integer
Dim myphone As String
Dim TOTLAND As Double
Dim tcode As String
Dim totadd As Double
TOTLAND = 0
Dim MM
totadd = 0
totplant = 0
DZstr = ""
SQLSTR = ""
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Set rsadd = Nothing
mchk = True
j = 0
If chkpririty.Value = 0 Then
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       DZstr = DZstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
Else
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    MM = Split(LSTPR.List(i), "|", -1, 1)
       DZstr = DZstr & "'" & Mid(MM(0), 1, 3) & Mid(MM(1), 1, 3) & Mid(MM(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = LSTPR.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE LOCATION TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next

End If




If Len(DZstr) > 0 Then
   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
 
Else
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


If OPTDETAIL.Value = True Then


If optall.Value = True Then

SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND  SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & " ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

ElseIf OPTWITHOUTDATE.Value = True Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND REGDATE>='1900-01-01' AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & " ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Else
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND REGDATE>='" & Format(TXTFROMDATE.Value, "yyyy-MM-dd") & "' and regdate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & " ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

End If




Else



If optall.Value = True Then
If chkpririty.Value = 0 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,phone1,phone2,SUM(REGLAND) as regland,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME,phone1,phone2 ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Else

SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,phone1,phone2,SUM(REGLAND) as regland,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND SUBSTRING(IDFARMER,1,9)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME,phone1,phone2 ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

End If

ElseIf OPTWITHOUTDATE.Value = True Then
If chkpririty.Value = 0 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND-acreplanted)) AS aditionalland,sum(regland) as regland,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID and REGDATE>='1900-01-01' AND SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Else
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND-acreplanted)) AS aditionalland,sum(regland) as regland,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID and REGDATE>='1900-01-01' AND SUBSTRING(IDFARMER,1,9)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"


End If
Else
If chkpririty.Value = 0 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND-acreplanted)) AS aditionalland,sum(regland) as regland,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID and AND REGDATE>='" & Format(TXTFROMDATE.Value, "yyyy-MM-dd") & "' and regdate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' AND  SUBSTRING(IDFARMER,1,3)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Else
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND-acreplanted)) AS aditionalland,sum(regland) as regland,village FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID and AND REGDATE>='" & Format(TXTFROMDATE.Value, "yyyy-MM-dd") & "' and regdate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' AND  SUBSTRING(IDFARMER,1,9)IN  " & DZstr
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

End If
End If




End If



                            
                        
                        
                        
                        

Dim excel_app As Object
Dim excel_sheet As Object

Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_sheet = Nothing
    Set excel_app = Nothing
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
    'excel_app.Visible = False
    '
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "S/N"
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
     excel_sheet.Cells(3, 5) = "FARMER CODE"
    excel_sheet.Cells(3, 6) = "FAMER"
    excel_sheet.Cells(3, 7) = "CONTACT #"
    excel_sheet.Cells(3, 8) = "VILLAGE"
    excel_sheet.Cells(3, 9) = "LAND (ACRE)"
    excel_sheet.Cells(3, 10) = "TOTAL PLANT"
    excel_sheet.Cells(3, 11) = UCase("Crates #")
    excel_sheet.Cells(3, 12) = UCase("B (Crate)")
    excel_sheet.Cells(3, 13) = UCase("E(Crate)")
    excel_sheet.Cells(3, 14) = UCase("B (No)")
    excel_sheet.Cells(3, 15) = UCase("P/L(Nos)")
    excel_sheet.Cells(3, 16) = UCase("Crates")
    excel_sheet.Cells(3, 17) = UCase("SSP (Kg)")
    excel_sheet.Cells(3, 18) = UCase("MOP(Kg)")
    excel_sheet.Cells(3, 19) = UCase("Urea(Kg)")
    excel_sheet.Cells(3, 20) = UCase("Dolomite(Kg)")
    excel_sheet.Cells(3, 21) = UCase("Total (Kg)")
    excel_sheet.Cells(3, 22) = UCase("Amount (Nu)")
    excel_sheet.Cells(3, 23) = UCase("Kg")
    excel_sheet.Cells(3, 24) = UCase("Amount (Nu)")
    excel_sheet.Cells(3, 25) = UCase("Total Amount(Nu)")
    i = 4
    s = 4
                        
                            Set rs = Nothing
                            rs.Open SQLSTR, MHVDB
                            If rs.EOF <> True Then
                            
                            chkred = False
                            
                           ' excel_sheet.Cells(i, 9).Formula = "=H4*3.85*100"
                            tcode = rs!tscode
                            
                            Do While rs.EOF <> True
                            
                             'If tcode <> rs!TSCODE Then
                             If totplant > 3800 Or totplant > 4000 Then
                             Dim m As Integer
                             
                             
                              excel_sheet.Range(excel_sheet.Cells(i, 2), _
                             excel_sheet.Cells(i, 26)).Select
                             
                            excel_app.Selection.Font.Color = vbRed
                             excel_app.Selection.Font.Bold = True
                                'excel_app.Selection.Columns.AutoFit
                            excel_app.Selection.Interior.ColorIndex = 15
                            
                            
                            
                            
                            
                            
                             excel_app.Range(excel_sheet.Cells(s, 1), _
                             excel_app.Cells(i - 1, 1)).Select
                                                         
                            
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter ' xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            'Selection.Merge
                            
                            
                            '
                            
                            excel_sheet.Range(excel_sheet.Cells(s, 26), _
                             excel_sheet.Cells(i - 1, 27)).Select
                             
                            
                            
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
                            'Selection.Merge
                            
                            
                            excel_sheet.Cells(s, 1) = sl
                            
                            
                            
                            
                            'excel_app.Selection.Interior.Pattern = xlSolid
                             'excel_sheet.Cells(i, 9) = totplant
                            excel_sheet.Cells(i, 9).Formula = "=SUM(H" & s & ":H" & i - 1 & ")"
                            excel_sheet.Cells(i, 10).Formula = "=SUM(I" & s & ":I" & i - 1 & ")"
                            excel_sheet.Cells(i, 11).Formula = "=SUM(J" & s & ":J" & i - 1 & ")"
                            excel_sheet.Cells(i, 12).Formula = "=SUM(K" & s & ":K" & i - 1 & ")"
                            excel_sheet.Cells(i, 13).Formula = "=SUM(L" & s & ":L" & i - 1 & ")"
                            excel_sheet.Cells(i, 14).Formula = "" ' "=SUM(M" & s & ":M" & i - 1 & ")"
                            excel_sheet.Cells(i, 15).Formula = "" ' "=SUM(N" & s & ":N" & i - 1 & ")"
                            excel_sheet.Cells(i, 16).Formula = "" ' "=SUM(O" & s & ":O" & i - 1 & ")"
                            excel_sheet.Cells(i, 17).Formula = "=SUM(P" & s & ":P" & i - 1 & ")"
                            excel_sheet.Cells(i, 18).Formula = "=SUM(Q" & s & ":Q" & i - 1 & ")"
                            excel_sheet.Cells(i, 19).Formula = "=SUM(R" & s & ":R" & i - 1 & ")"
                            excel_sheet.Cells(i, 20).Formula = "=SUM(S" & s & ":S" & i - 1 & ")"
                            excel_sheet.Cells(i, 21).Formula = "=SUM(T" & s & ":T" & i - 1 & ")"
                            excel_sheet.Cells(i, 22).Formula = "=SUM(U" & s & ":U" & i - 1 & ")"
                            excel_sheet.Cells(i, 23).Formula = "=SUM(V" & s & ":V" & i - 1 & ")"
                            excel_sheet.Cells(i, 24).Formula = "=SUM(W" & s & ":W" & i - 1 & ")"
                            excel_sheet.Cells(i, 25).Formula = "=SUM(X" & s & ":X" & i - 1 & ")"
                            
                            
                            i = i + 1
                            m = i
                            s = m
                            sl = sl + 1
                            totplant = 0
                             End If
'                             excel_app.Selection.Font.Color = vbBlack
'                             excel_app.Selection.Font.Bold = False
                             
                             
                             
                             
                            'excel_sheet.Cells(i, 1) = sl
                            FindDZ Mid(rs!idfarmer, 1, 3)
                            excel_sheet.Cells(i, 2) = rs!dzcode & " " & Dzname
                            
                            
                            FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
                            
                            excel_sheet.Cells(i, 3) = rs!GECODE & " " & GEname
                            FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
                            excel_sheet.Cells(i, 4) = rs!tscode & " " & TsName
                            FindFA rs!idfarmer, "F"
                            
                         
                            
  
               If chkred = True Then
  excel_sheet.Range(excel_sheet.Cells(i, 2), _
                             excel_sheet.Cells(i, 25)).Select
                             excel_app.Selection.Font.Color = vbRed

  End If
  
  chkred = False
                            excel_sheet.Cells(i, 5) = rs!idfarmer
                            excel_sheet.Cells(i, 6) = FAName
                            myphone = IIf(IsNull(rs!phone1), "", rs!phone1) & "," & IIf(IsNull(rs!phone2), "", rs!phone2)
                            If Len(myphone) > 0 Then
                            myphone = Left(myphone, Len(myphone) - 1)
                            End If
                         
                            excel_sheet.Cells(i, 7) = myphone
                            excel_sheet.Cells(i, 8) = rs!VILLAGE
                            Set rsadd = Nothing
                            rsadd.Open "select sum(acreplanted) as rl from tblplanted where farmercode='" & rs!idfarmer & "' ", MHVDB

                           totadd = IIf(IsNull(rsadd!rl), 0, rsadd!rl)



                         totadd = Format(IIf(IsNull(rs!regland), 0, rs!regland) - totadd, "####0.00")
                         If totadd < 0 Then
                           excel_sheet.Range(excel_sheet.Cells(i, 2), _
                             excel_sheet.Cells(i, 25)).Select
                             excel_app.Selection.Font.Color = vbRed
                         
                         End If
                           If OPTWO.Value = True Then
                           
                           If totadd <> 0 Then
                           
                           
                           excel_sheet.Cells(i, 9) = Format(totadd, "####0.00") 'Format(IIf(IsNull(rs!REGLAND), 0, rs!REGLAND), "####0.00") '
                         
                         
                         
                            excel_sheet.Cells(i, 10).Formula = "=ROUND(" & "I" & i & " *420,0)"
                            totplant = totplant + excel_sheet.Cells(i, 9)
                            excel_sheet.Cells(i, 11) = "=ROUND(J" & i & "/35,0)"

                            excel_sheet.Cells(i, 12).Formula = "=ROUND(K" & i & "*0.6,0)"
                            excel_sheet.Cells(i, 13).Formula = "=K" & i & "-P" & i & "-L" & i
                            excel_sheet.Cells(i, 14).Formula = ""
                            excel_sheet.Cells(i, 15).Formula = ""
                            excel_sheet.Cells(i, 16).Formula = ""
                            excel_sheet.Cells(i, 17).Formula = "=ROUND(J" & i & "*0.02,0)"
                            excel_sheet.Cells(i, 18).Formula = "=ROUND(J" & i & "*0.005,0)"
                            excel_sheet.Cells(i, 19).Formula = "=ROUND(J" & i & "*0.0075,0)"
                            excel_sheet.Cells(i, 20).Formula = "=ROUND(J" & i & "*0.1,0)"
                            excel_sheet.Cells(i, 21).Formula = "=SUM(Q" & i & ":T" & i & ")"
                            excel_sheet.Cells(i, 22).Formula = "=ROUND((Q" & i & "*13.38)+(R" & i & " *24.6)+(S" & i & "*13.9)+(T" & i & "*3.92),0)"
                            excel_sheet.Cells(i, 23).Formula = "=ROUND(J" & i & "*0.25,0)"
                            excel_sheet.Cells(i, 24).Formula = "=ROUND(W" & i & "*3.92,0)"
                            excel_sheet.Cells(i, 25).Formula = "=ROUND(X" & i & "+V" & i & ",0)"
               
                             i = i + 1
                             Else
                             
                             End If
                             Else
                             excel_sheet.Cells(i, 9) = Format(totadd, "####0.00") 'Format(IIf(IsNull(rs!REGLAND), 0, rs!REGLAND), "####0.00") '
                             excel_sheet.Cells(i, 10).Formula = "=ROUND(" & "I" & i & " *3.85*100,0)"
                            totplant = totplant + excel_sheet.Cells(i, 9)
                            excel_sheet.Cells(i, 11) = "=ROUNDUP(I" & i & "/35,0)"

                            excel_sheet.Cells(i, 12).Formula = "=ROUNDUP(K" & i & "*0.6,0)"
                            excel_sheet.Cells(i, 13).Formula = "=K" & i & "-P" & i & "-L" & i
                            excel_sheet.Cells(i, 14).Formula = ""
                            excel_sheet.Cells(i, 15).Formula = ""
                            excel_sheet.Cells(i, 16).Formula = ""
                            excel_sheet.Cells(i, 17).Formula = "=ROUND(J" & i & "*0.02,0)"
                            excel_sheet.Cells(i, 18).Formula = "=ROUND(J" & i & "*0.005,0)"
                            excel_sheet.Cells(i, 19).Formula = "=ROUND(J" & i & "*0.0075,0)"
                            excel_sheet.Cells(i, 20).Formula = "=ROUNDUP(J" & i & "*0.1,0)"
                            excel_sheet.Cells(i, 21).Formula = "=SUM(J" & i & ":T" & i & ")"
                            excel_sheet.Cells(i, 22).Formula = "=ROUND((Q" & i & "*13.38)+(R" & i & " *24.6)+(S" & i & "*13.9)+(T" & i & "*3.79),0)"
                            excel_sheet.Cells(i, 23).Formula = "=ROUND(J" & i & "*0.25,0)"
                            excel_sheet.Cells(i, 24).Formula = "=ROUND(W" & i & "*3.79,0)"
                            excel_sheet.Cells(i, 25).Formula = "=ROUND(X" & i & "+V" & i & ",0)"
                             i = i + 1
                             
                             
                             End If
                             
                             
                             
                             
                             
                              tcode = rs!tscode
                                           rs.MoveNext
                                           Loop
                        
                            End If
                      
                                i = i + 1
                             
                             
                              excel_sheet.Range(excel_sheet.Cells(i, 2), _
                             excel_sheet.Cells(i, 25)).Select
                             
                            excel_app.Selection.Font.Color = vbRed
                             excel_app.Selection.Font.Bold = True
                                'excel_app.Selection.Columns.AutoFit
                            excel_app.Selection.Interior.ColorIndex = 15
                            
                            
                            
                            
                            
                            
                             excel_app.Range(excel_sheet.Cells(s, 1), _
                             excel_app.Cells(i - 1, 1)).Select
                                                         
                            
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter ' xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            'Selection.Merge
                            
                            
                            '
                            
                            excel_sheet.Range(excel_sheet.Cells(s, 26), _
                             excel_sheet.Cells(i - 1, 27)).Select
                                           
                            
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
                            'Selection.Merge
                            
                            
                            excel_sheet.Cells(s, 1) = sl
                            
                            
                            
                            
                            'excel_app.Selection.Interior.Pattern = xlSolid
                             'excel_sheet.Cells(i, 9) = totplant
                             excel_sheet.Cells(i, 9).Formula = "=SUM(I" & s & ":I" & i - 1 & ")"
                            excel_sheet.Cells(i, 10).Formula = "=SUM(J" & s & ":J" & i - 1 & ")"
                            excel_sheet.Cells(i, 11).Formula = "=SUM(K" & s & ":K" & i - 1 & ")"
                            excel_sheet.Cells(i, 12).Formula = "=SUM(L" & s & ":L" & i - 1 & ")"
                            excel_sheet.Cells(i, 13).Formula = "=SUM(M" & s & ":M" & i - 1 & ")"
                            excel_sheet.Cells(i, 14).Formula = "" ' "=SUM(M" & s & ":M" & i - 1 & ")"
                            excel_sheet.Cells(i, 15).Formula = "" ' "=SUM(N" & s & ":N" & i - 1 & ")"
                            excel_sheet.Cells(i, 16).Formula = "" ' "=SUM(O" & s & ":O" & i - 1 & ")"
                            excel_sheet.Cells(i, 17).Formula = "=SUM(Q" & s & ":Q" & i - 1 & ")"
                            excel_sheet.Cells(i, 18).Formula = "=SUM(R" & s & ":R" & i - 1 & ")"
                            excel_sheet.Cells(i, 19).Formula = "=SUM(S" & s & ":S" & i - 1 & ")"
                            excel_sheet.Cells(i, 20).Formula = "=SUM(T" & s & ":T" & i - 1 & ")"
                            excel_sheet.Cells(i, 21).Formula = "=SUM(U" & s & ":U" & i - 1 & ")"
                            excel_sheet.Cells(i, 22).Formula = "=SUM(V" & s & ":V" & i - 1 & ")"
                            excel_sheet.Cells(i, 23).Formula = "=SUM(W" & s & ":W" & i - 1 & ")"
                            excel_sheet.Cells(i, 34).Formula = "=SUM(X" & s & ":X" & i - 1 & ")"
                            excel_sheet.Cells(i, 25).Formula = "=SUM(Y" & s & ":Y" & i - 1 & ")"
                            
                            
                            i = i + 1
                            m = i
                            s = m
                            sl = sl + 1
                            totplant = 0





                            
                            'make up
                            excel_sheet.Range(excel_sheet.Cells(4, 8), _
                            excel_sheet.Cells(i, 8)).Select
                            excel_app.Selection.NumberFormat = "####0.00"
                            excel_sheet.Range(excel_sheet.Cells(1, 12), _
                            excel_sheet.Cells(1, 16)).Select
                             
                            
                            
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
                            
                            
                            excel_sheet.Cells(1, 12) = "Variety"
                            
                            
                            
                            
                            excel_sheet.Range(excel_sheet.Cells(1, 17), _
                             excel_sheet.Cells(1, 22)).Select
                             
                            
                            
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
                           
                            excel_sheet.Cells(1, 17) = "Pre - Mixed Fertilizer"
                            
                            
                            
                            excel_sheet.Range(excel_sheet.Cells(1, 23), _
                             excel_sheet.Cells(1, 24)).Select
                             
                            
                            
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
                            
                            
                            excel_sheet.Cells(1, 23) = "Dolomite"
                            
                            
            excel_sheet.Range(excel_sheet.Cells(2, 16), _
                             excel_sheet.Cells(2, 18)).Select
                             
                            
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
                           
                            
                            excel_sheet.Cells(2, 16) = "Mixed Variety"
                            
                            excel_sheet.Range(excel_sheet.Cells(2, 26), _
                             excel_sheet.Cells(3, 27)).Select
                             
                            
                            
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
                           
                            excel_sheet.Cells(2, 26) = "Schedule Date, Vehicle No & Team Captainy"
                            
                            
                            
                            
                            
                            

                            
                            
 excel_sheet.Range(excel_sheet.Cells(1, 1), _
                             excel_sheet.Cells(i, 27)).Select
'excel_sheet.Columns("A:A").Select
 excel_app.Selection.Font.Size = 10

 
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
                            
                            
                            
                            
                            
 





excel_sheet.Columns("A:A").Select
 excel_app.Selection.ColumnWidth = 3.57
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


    excel_sheet.Columns("b:d").Select
 excel_app.Selection.ColumnWidth = 14.86
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


  

excel_sheet.Columns("e:f").Select
 excel_app.Selection.ColumnWidth = 17
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


excel_sheet.Columns("g:Y").Select
 excel_app.Selection.ColumnWidth = 8
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

excel_sheet.Columns("Z:Z").Select
 excel_app.Selection.ColumnWidth = 7
With excel_app.Selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With







                            
                            
                            
     excel_sheet.Range(excel_sheet.Cells(1, 1), _
                             excel_sheet.Cells(i, 27)).Select
                            
                   excel_app.Selection.Font.Name = "Times New Roman"
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A1:Z3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "PLANT DISTRIBUTION LIST"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
'    Set excel_sheet = Nothing
'    Set excel_app = Nothing

   

excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

End Sub
Private Sub alignwidth()
    

End Sub

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
   SQLSTR = "insert into tbltemp(end ,dcode,gcode,tcode,fcode,farmercode,totaltrees,) SELECT max(END), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode FROM phealthhub15_core where farmerbarcode<>'' group  by farmerbarcode,fdcode"
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
    
  If rsF!var1 > rss!End Then
  db.Execute "update tbltemp set var1='" & Format(rsF!var1, "yyyy-MM-dd") & "' , var7='" & rss!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & rss!FDCODE & "' "
  End If
    
    
  Else
  'db.Execute "update tbltemp set var1='" & RSS!End & "' , var7='" & RSS!totaltrees & "' where var6='" & mfcode & "' and fs='F' and fdcode='" & RSS!fdcode & "' "
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs,fdcode)values('" & Format(rss!End, "yyyy-MM-dd") & "','99','99','99','99','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "') "
  
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
  db.Execute "insert into  tbltemp(var1,var2,var3,var4,var5,var6,var7,fs)values('" & Format(rss!End, "yyyy-MM-dd") & "','" & mdcode & "','" & mgcode & "','" & mtcode & "','99','" & mfcode & "','" & rss!totaltrees & "','S') "
  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "date last visited"
    excel_sheet.Cells(3, 3) = "farmercode"
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
                            excel_sheet.Cells(i, 1) = sl
                           
     excel_sheet.Cells(i, 2) = "'" & rs!var1
   
   excel_sheet.Cells(i, 3) = rs!var6
  
  
  
  
  
    i = i + 1
      sl = sl + 1
    
    
    
    
    
  
    
 
                            
                            rs.MoveNext
                            Loop
End If

      
                            
                            
                            'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 11)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
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
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,id,sname,fname) SELECT (END), dcode, gcode, tcode,fcode, farmerbarcode, (totaltrees),'F' as fs,fdcode,id,sname,fname FROM phealthhub15_core where farmerbarcode<>''"
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
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,fdcode,id,sname,fname)values('" & Format(rss!End, "yyyy-MM-dd") & "','" & mdcode & "','" & mgcode & "','" & mtcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','F','" & rss!FDCODE & "','" & rss!id & "','" & rss!sname & "','" & rss!fname & "') "

  
  rss.MoveNext
  Loop
  
  
  'storage
  SQLSTR = ""
   SQLSTR = "insert into tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,id,sname,fname) SELECT (END), dcode, gcode, tcode,fcode, scanlocation, (totaltrees),'S' as fs,id,sname,fname FROM storagehub6_core where scanlocation<>''"
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
  db.Execute "insert into  tbltemp(end,dcode,gcode,tcode,fcode,farmercode,totaltrees,fs,id,sname,fname)values('" & Format(rss!End, "yyyy-MM-dd") & "','" & mdcode & "','" & mgcode & "','" & mtcode & "','" & rss!fcode & "','" & mfcode & "','" & rss!totaltrees & "','S','" & rss!id & "','" & rss!sname & "','" & rss!fname & "') "
  
  
  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
     excel_sheet.Cells(3, 2) = "DATE VISITED"
    excel_sheet.Cells(3, 3) = "DZONGKHAG"
    excel_sheet.Cells(3, 4) = "GEWOG"
    excel_sheet.Cells(3, 5) = "TSHOWOG"
     excel_sheet.Cells(3, 6) = "FARMER CODE"
    excel_sheet.Cells(3, 7) = "FAMER"
    excel_sheet.Cells(3, 8) = "S. ID"
    excel_sheet.Cells(3, 9) = "S. NAME"
      i = 4
                        
                        
                        SQLSTR = ""
'SQLSTR = "SELECT farmercode,sum(acreplanted)as pl,sum(nooftrees) as t from tblplanted group by farmercode"
         '
               
              SQLSTR = "select * from tbltemp where substring(end,1,2)<>0 and farmercode not in(select farmercode from allfarmersexdropped where FarmerCode is not null) group by farmercode order by fname desc"
                            Set rs = Nothing
                            rs.Open SQLSTR, db
                            If rs.EOF <> True Then
                            Do While rs.EOF <> True
                            chkred = False
                            excel_sheet.Cells(i, 1) = sl
                  excel_sheet.Cells(i, 2) = "'" & rs!End
                excel_sheet.Cells(i, 3) = "'" & rs!dcode
                excel_sheet.Cells(i, 4) = "'" & rs!gcode
                excel_sheet.Cells(i, 5) = "'" & rs!tcode
                If sl = 126 Then
                
                MsgBox "asdlcjads"
                End If
                excel_sheet.Cells(i, 6) = rs!farmercode
                excel_sheet.Cells(i, 7) = rs!fname
    
   excel_sheet.Cells(i, 8) = rs!id
 excel_sheet.Cells(i, 9) = rs!sname
  
  
  
    i = i + 1
      sl = sl + 1
    
    
    
    
    
  
    
 
                            
                            rs.MoveNext
                            Loop
End If

      
                            
                            
                            'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 11)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
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

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
FRMRPTLANDDETAILS.Width = 7560
Set rs = Nothing

rs.Open "select DZONGKHAGCODE,DZONGKHAGNAME from tbldzongkhag Order by DZONGKHAGCODE", MHVDB, adOpenStatic
With rs
Do While Not .EOF
   DZLIST.AddItem Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
mchk = False
End Sub

Private Sub OPTALL_Click()
Frame1.Visible = False
End Sub

Private Sub OPTWITHDate_Click()
If OPTWITHDate.Value = True Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
End Sub

Private Sub OPTWITHOUTDATE_Click()
Frame1.Visible = False
End Sub
