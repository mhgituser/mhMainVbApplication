VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmfieldvisit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIELD VISIT REPORT"
   ClientHeight    =   5085
   ClientLeft      =   8475
   ClientTop       =   3045
   ClientWidth     =   9165
   Icon            =   "frmfieldvisit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9165
   Begin VB.ListBox lstm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   5640
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CheckBox CHKMOREOPTION 
      Caption         =   "SUMMARY"
      Height          =   195
      Left            =   3960
      TabIndex        =   20
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "MORE OPTION"
      Height          =   1695
      Left            =   5640
      TabIndex        =   17
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "-"
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
         Left            =   2880
         TabIndex        =   26
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TXTYR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1000
         Width           =   855
      End
      Begin VB.OptionButton OPTMONTHLY 
         Caption         =   "MONTHLY"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OPTYEARLY 
         Caption         =   "YEARLY"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "YEAR"
         Height          =   195
         Left            =   1800
         TabIndex        =   23
         Top             =   1000
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   2520
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
         ItemData        =   "frmfieldvisit.frx":076A
         Left            =   1080
         List            =   "frmfieldvisit.frx":076C
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83755009
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   83755009
         CurrentDate     =   41362
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
         Height          =   195
         Left            =   2760
         TabIndex        =   14
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
      Left            =   2280
      Picture         =   "frmfieldvisit.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
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
      Picture         =   "frmfieldvisit.frx":1438
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame 
      Caption         =   "RECORD SELECTION"
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
      Begin VB.OptionButton optnovisit 
         Caption         =   "0 VISIT"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox CHKMONITOR 
         Caption         =   "SELECT MONITOR"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton OPTTOPN 
         Caption         =   "LAST VISIT"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton OPTALLVISIT 
         Caption         =   "ALL VISIT"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TXTRECORDNO 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   350
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "VISIT"
         Height          =   195
         Left            =   3435
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DATE SELECTION"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton OPTSEL 
         Caption         =   "SELECTIVE"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   5040
   End
End
Attribute VB_Name = "frmfieldvisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMoniter As New ADODB.Recordset
Dim DZstr As String
Private Sub Check1_Click()

End Sub

Private Sub CHKMONITOR_Click()
If CHKMONITOR.Value = 1 Then

If CHKMOREOPTION.Value = 1 Then
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False
lstm.Height = 2985
lstm.Top = 2040
Frame2.Visible = True
lstm.Visible = True
populatelist
Else
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = True
lstm.Height = 4560
lstm.Top = 120
Frame2.Visible = False
lstm.Visible = True
populatelist
End If


Else

If CHKMOREOPTION.Value = 1 Then
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False
Frame2.Visible = True
lstm.Visible = False
Else
frmfieldvisit.Width = 5595
OPTTOPN.Enabled = True
TXTYR.Text = SysYear
End If


End If

If CHKMONITOR.Value = 0 And CHKMOREOPTION.Value = 0 Then
frmfieldvisit.Width = 5595
End If

End Sub

Private Sub CHKMOREOPTION_Click()
If CHKMOREOPTION.Value = 1 Then
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False

If CHKMONITOR.Value = 1 Then
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False
lstm.Height = 2985
lstm.Top = 2040
Frame2.Visible = True
lstm.Visible = True
populatelist
Else
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False

Frame2.Visible = True
lstm.Visible = False



End If


Else


OPTTOPN.Enabled = True

If CHKMONITOR.Value = 1 Then
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False
lstm.Height = 4560
lstm.Top = 120
Frame2.Visible = False
lstm.Visible = True
populatelist
Else
frmfieldvisit.Width = 9435
OPTTOPN.Enabled = False

Frame2.Visible = True
lstm.Visible = False



End If

End If

OPTMONTHLY.Value = True
Frame1.Visible = True
TXTYR.Text = SysYear
If CHKMONITOR.Value = 0 And CHKMOREOPTION.Value = 0 Then
frmfieldvisit.Width = 5595
End If
End Sub

Private Sub Command1_Click()
TXTYR.Text = Val(TXTYR + 1)
End Sub

Private Sub Command2_Click()
TXTYR.Text = Val(TXTYR - 1)
If Val(TXTYR.Text) < 2012 Then
TXTYR.Text = 2012
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
'If OPTSEL.Value = True And Len(CBODATE.Text) = 0 Then
'MsgBox "Please Select The Date Type."
'Exit Sub
'End If


myvisit
Exit Sub



If optnovisit.Value = True Then

novisit
Exit Sub

End If



DZstr = ""

If CHKMONITOR.Value = 1 Then
For i = 0 To lstm.ListCount - 1
    If lstm.Selected(i) Then
       DZstr = DZstr + "'" + Trim(Mid(lstm.List(i), InStr(1, lstm.List(i), "|") + 1)) + "',"
    End If
Next





If Len(DZstr) > 0 Then
   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
 
Else
   MsgBox "MONITOR NOT SELECTED !!!"
   Exit Sub
End If
End If








If OPTYEARLY.Value = False Then
allVISIT
Else
yearlyvisit
End If

End Sub
Private Sub myvisit()
Dim fdcnt As Integer
Dim excel_app As Object
Dim excel_sheet As Object
Dim Excel_WBook As Object
Dim exactrow As Integer
Dim Excel_Chart As Object
'On Error Resume Next
Dim jrow As Long
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
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


'GetTbl
'
'
'      SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
'         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
'         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select end,staffbarcode,region_dcode," _
'         & "region_gcode,region,farmerbarcode,0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
'         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
'         & "rootpest,animaldamage,monitorcomments from phealthhub15_core where   status<>'BAD' "
'
'
'                SQLSTR = SQLSTR & "where   status<>'BAD'"
'
'
'  db.Execute SQLSTR
SQLSTR = " select end,staffbarcode,farmerbarcode,fdcode from phealthhub15_core where   status<>'BAD' "


''
'SQLSTR = ""
'SQLSTR = "select  max(END) as end,id,farmercode,fdcode,count(fdcode) as fcnt from " & Mtblname & ""
'SQLSTR = SQLSTR & " group by id,farmercode,fdcode order by id,end,farmercode,fdcode"


'SQLSTR = "select  max(END) as end,id,farmercode,fdcode,count(fdcode) as fcnt from " & Mtblname & ""
'SQLSTR = SQLSTR & " group by id,farmercode,fdcode order by id,end,farmercode,fdcode"

'SQLSTR = "select max(END) as end,farmercode as id ,count(farmercode) as jval,year(end) as procyear,month(end) as procmonth from " & Mtblname & "  where  staffbarcode='" & RS1!staffbarcode & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),farmercode,fdcode order by farmercode,fdcode,year(end),month(end)"

fdcnt = 0
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
    excel_app.Visible = True
     jrow = 2
      excel_sheet.Cells(jrow, 1) = "FARMER"
      
     
    Set RS1 = Nothing
    RS1.Open "SELECT DISTINCT staffbarcode FROM dailyacthub9_core", db
    Do While RS1.EOF <> True
    SQLSTR = "select max(END) as end,farmerbarcode,concat(farmerbarcode,cast(fdcode as char))  as id ,fdcode,count(farmerbarcode) as jval,year(end) as procyear,month(end) as procmonth from phealthhub15_core  where  staffbarcode='" & RS1!staffbarcode & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),farmerbarcode,fdcode order by farmerbarcode,fdcode,year(end),month(end)"
    Set rs = Nothing
    rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    
    jCol = 5 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF RS1!staffbarcode
     
       jrow = jrow + 1
    excel_sheet.Cells(jrow, 1) = RS1!staffbarcode & " " & sTAFF
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
          j = rs!jval
          jtot = jtot + j
         
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, 2) = "'" & rs!End
          excel_sheet.Cells(jrow, 3) = CInt(rs!fdcode)
          
          'fdcnt = fdcnt + 1
          excel_sheet.Cells(jrow, i) = Val(j)
         pyear = rs!id
          rs.MoveNext
         
          If rs.EOF Then Exit Do
          'jrow = jrow + 1
       Loop
      fdcnt = fdcnt + 1
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
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
    RS1.MoveNext
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
    
  
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A1:o3").Font.Bold = True
    






    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
    
 




End Sub
Private Sub novisit()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsf As New ADODB.Recordset
Dim actstring As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
GetTblmhv
MHVDB.Execute "insert into " & Mtblname & "(id) select staffbarcode from odk_prodlocal.phealthhub15_core"
SLNO = 1

SQLSTR = ""
 
   
   
  'db.Execute SQLSTR

SQLSTR = "select * from tblfarmer where status='A' and length(monitor)='5' and monitor not in (select id from " & Mtblname & ")"






'On Error Resume Next

mchk = True


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
   
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "MONITOR ID. "
    excel_sheet.Cells(3, 3) = "MONITOR NAME "
'    If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
'    excel_sheet.Cells(3, 4) = "DATE" & "(END)"
'    Else
'    excel_sheet.Cells(3, 4) = "DATE" & "(" & CBODATE.Text & ")"
'    End If
    
    
        excel_sheet.Cells(3, 5) = "D"
    excel_sheet.Cells(3, 6) = "G"
    excel_sheet.Cells(3, 7) = "T"
    excel_sheet.Cells(3, 8) = UCase("Farmer ID")
    excel_sheet.Cells(3, 9) = UCase("Farmer name")
    excel_sheet.Cells(3, 10) = UCase("FIELD ID")
    excel_sheet.Cells(3, 11) = UCase("VISIT")
   i = 4
  Set rs = Nothing
  
rs.Open SQLSTR, MHVDB
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = rs!id '"'" & rs!End  'rs.Fields(Mindex)

FindsTAFF "S0" & rs!id
excel_sheet.Cells(i, 3) = sTAFF
excel_sheet.Cells(i, 4) = "'" & rs!End

excel_sheet.Cells(i, 5) = Mid(rs!idfarmer, 1, 3)
excel_sheet.Cells(i, 6) = Mid(rs!idfarmer, 4, 3)
excel_sheet.Cells(i, 7) = Mid(rs!idfarmer, 7, 3)
excel_sheet.Cells(i, 8) = IIf(IsNull(rs!idfarmer), "", rs!idfarmer)
FindFA rs!farmercode, "F"
excel_sheet.Cells(i, 9) = FAName
excel_sheet.Cells(i, 10) = rs!fdcode

excel_sheet.Cells(i, 11) = rs!FCNT

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
     excel_sheet.Range("A3:AA3").Font.Bold = True
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
MHVDB.Execute "drop table " & Mtblname & ""
db.Close
Exit Sub
err:
db.Execute "drop table " & Mtblname & ""
MsgBox err.Description
err.Clear
End Sub
Private Sub yearlyvisit()
Dim excel_app As Object
Dim excel_sheet As Object
Dim Excel_WBook As Object
Dim Excel_Chart As Object
'On Error Resume Next
Dim jrow As Long
Dim maxyear, minyear As Integer
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
Dim DT2 As Date
intYear = CInt(TXTYR.Text)
intMonth = 1
intDay = 1
 DT1 = DateSerial(intYear, intMonth, intDay)
txtfrmdate.Value = Format(DT1, "dd/MM/yyyy")
intYear = CInt(TXTYR.Text)
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
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core "
         
                If OPTALL.Value = True Then
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD'"
                SQLSTR = SQLSTR & "where   status<>'BAD'"
                Else
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                 SQLSTR = SQLSTR & "where status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                End If
   
  db.Execute SQLSTR
                  
Set rs = Nothing

rs.Open "select year(max(end)) as mm from " & Mtblname & " ", db
maxyear = rs!MM
Set rs = Nothing
 rs.Open "select year(min(end)) as mm from " & Mtblname & " ", db
minyear = rs!MM

If CHKMONITOR.Value = 1 And Len(DZstr) <> 0 Then
SQLSTR = "select id ,count(fdcode) as jval,year(end) as procyear from " & Mtblname & " where  id in " & DZstr & "  group by year(end),id order by id,year(end)"

Else
SQLSTR = "select id ,count(fdcode) as jval,year(end) as procyear from " & Mtblname & "  group by year(end),id order by id,year(end)"
End If




For i = 1 To 13
    mtot(i) = 0
Next
Set rs = Nothing
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
    excel_app.Visible = True
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    excel_sheet.Cells(3, 1) = "MONITOR"
    K = 1
    Dim yy As Integer
    yy = minyear
    For i = minyear To maxyear
        K = K + 1
        excel_sheet.Cells(3, K) = yy
        yy = yy + 1
    Next
    jCol = K + 1
    excel_sheet.Cells(3, jCol) = UCase("Total")
    
    
    jrow = 3
    Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       FindsTAFF rs!id
       excel_sheet.Cells(jrow, 1) = rs!id & " " & sTAFF
       jtot = 0
       Do While pyear = rs!id
          i = rs!ProcYear + 3 - maxyear
          j = rs!jval
          jtot = jtot + j
          mtot(i - 1) = mtot(i - 1) + j
          excel_sheet.Cells(jrow, i) = Val(j)
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
    Loop
    jtot = 0
    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
    For i = 2 To jCol - 1
        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
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
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing

     
db.Execute "drop table " & Mtblname & ""
     Screen.MousePointer = vbDefault
     End Sub
Private Sub visitM()
Dim excel_app As Object
Dim excel_sheet As Object
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
Dim DT2 As Date
intYear = CInt(TXTYR.Text)
intMonth = 1
intDay = 1
 DT1 = DateSerial(intYear, intMonth, intDay)
txtfrmdate.Value = Format(DT1, "dd/MM/yyyy")
intYear = CInt(TXTYR.Text)
intMonth = 12
intDay = 31
 DT2 = DateSerial(intYear, intMonth, intDay)
txttodate.Value = Format(DT2, "dd/MM/yyyy")



Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
               
db.Open OdkCnnString
                  
If CHKMONITOR.Value = 1 And Len(DZstr) <> 0 Then
SQLSTR = "select id ,count(fdcode) as jval,year(end) as procyear,month(end) as procmonth from " & Mtblname & " where  id in " & DZstr & " and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),id order by id,year(end),month(end)"

Else
SQLSTR = "select id ,count(fdcode) as jval,year(end) as procyear,month(end) as procmonth from " & Mtblname & " where end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),id order by id,year(end),month(end)"
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
    excel_app.Visible = True
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    excel_sheet.Cells(3, 1) = "MONITOR"
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.Cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & TXTYR.Text
    Next
    excel_sheet.Cells(3, jCol) = UCase("Total")
    
    
    jrow = 3
    Do Until rs.EOF
       jrow = jrow + 1
       pyear = rs!id
       FindsTAFF rs!id
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
       Loop
       excel_sheet.Cells(jrow, jCol) = Val(jtot)
    Loop
    jtot = 0
    excel_sheet.Cells(jrow + 1, 1) = UCase("Total")
    For i = 2 To jCol - 1
        excel_sheet.Cells(jrow + 1, i) = mtot(i - 1)
        jtot = jtot + mtot(i - 1)
    Next
    excel_sheet.Cells(jrow + 1, jCol) = Val(jtot)
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
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
    

End Sub
Private Sub allVISIT()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsf As New ADODB.Recordset
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

Dim SQLSTR As String
SQLSTR = ""
SLNO = 1



SQLSTR = ""

'If OPTALL.Value = True Then


 
      SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments) select end,staffbarcode,region_dcode," _
         & "region_gcode,region,farmerbarcode,0,fdcode,tree_count_totaltrees,qc_tally_goodmoisture,qc_tally_poormoisture,qc_tally_goodmoisture+qc_tally_poormoisture," _
         & "tree_count_deadmissing,tree_count_slowgrowing,tree_count_dor,tree_count_activegrowing,shock,nutrient,waterlog,activepest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments from phealthhub15_core "
         
                If OPTALL.Value = True Then
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD'"
                SQLSTR = SQLSTR & "where   status<>'BAD'"
                Else
                'SQLSTR = SQLSTR & "where farmerbarcode<>'' and  status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                 SQLSTR = SQLSTR & "where status<>'BAD' and substring(end,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(end,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'"
                End If
   
  db.Execute SQLSTR

SQLSTR = ""



If CHKMOREOPTION.Value = 0 Then
If CHKMONITOR.Value = 1 And Len(DZstr) <> 0 Then
    SQLSTR = "select max(END) as end,id,farmercode,fdcode,count(fdcode) as fcnt from " & Mtblname & " where id in " & DZstr
    Else
    SQLSTR = "select  max(END) as end,id,farmercode,fdcode,count(fdcode) as fcnt from " & Mtblname & ""
End If
Else

If OPTYEARLY.Value = True Then
'YEARLY

Exit Sub
Else
visitM
db.Execute "drop table " & Mtblname & ""
Exit Sub


End If
End If




SQLSTR = SQLSTR & " group by id,farmercode,fdcode order by id,end,farmercode,fdcode"




On Error Resume Next

mchk = True


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
   
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "MONITOR ID. "
    excel_sheet.Cells(3, 3) = "MONITOR NAME "
'    If OPTALL.Value = True And OPTALLFIELDS.Value = True Then
  excel_sheet.Cells(3, 4) = UCase("LAST VISITED")
'    Else
'    excel_sheet.Cells(3, 4) = "DATE" & "(" & CBODATE.Text & ")"
'    End If
    
    
        excel_sheet.Cells(3, 5) = "DZONGKHAG"
    excel_sheet.Cells(3, 6) = "GEWOG"
    excel_sheet.Cells(3, 7) = "TSHOWOG"
    excel_sheet.Cells(3, 8) = UCase("Farmer ID")
    excel_sheet.Cells(3, 9) = UCase("Farmer name")
    excel_sheet.Cells(3, 10) = UCase("FIELD ID")
    excel_sheet.Cells(3, 11) = UCase("VISIT")
     excel_sheet.Cells(3, 12) = UCase("LAST VISITED")
   i = 4
  Set rs = Nothing
  
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = rs!id '"'" & rs!End  'rs.Fields(Mindex)

FindsTAFF rs!id
excel_sheet.Cells(i, 3) = sTAFF
excel_sheet.Cells(i, 4) = "'" & rs!End

excel_sheet.Cells(i, 5) = Mid(rs!farmercode, 1, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmercode, 4, 3)
excel_sheet.Cells(i, 7) = Mid(rs!farmercode, 7, 3)
excel_sheet.Cells(i, 8) = IIf(IsNull(rs!farmercode), "", rs!farmercode)
FindFA rs!farmercode, "F"
excel_sheet.Cells(i, 9) = FAName
excel_sheet.Cells(i, 10) = rs!fdcode

excel_sheet.Cells(i, 11) = rs!FCNT

'Set rss = Nothing
'rss.Open "select max(end) as end from phealthhub15_core where id='" & rs!id & "' and farmerbarcode='" & rs!farmercode & "'", db
''If rss.EOF <> True Then
'
excel_sheet.Cells(i, 12) = "'" & rs!End
''End If




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
     excel_sheet.Range("A3:AA3").Font.Bold = True
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


Private Sub populatelist()
Dim rs As New ADODB.Recordset
lstm.Clear
Set rs = Nothing

rs.Open "select staffcode,staffname from tblmhvstaff where moniter='1' order by staffcode ", MHVDB, adOpenStatic
With rs
Do While Not .EOF
   lstm.AddItem Trim(!staffname) + " | " + !staffcode
   .MoveNext
Loop
End With
End Sub
Private Sub Form_Load()

txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
frmfieldvisit.Width = 5595


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


 TXTYR.Text = SysYear

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

Private Sub OPTMONTHLY_Click()
If OPTMONTHLY.Value = True Then
Frame1.Visible = False
Else
Frame1.Visible = True
End If
End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
End Sub

Private Sub OPTTOPN_Click()
If OPTTOPN.Value = True Then
CHKMOREOPTION.Enabled = False
Else
CHKMOREOPTION.Enabled = True

End If
End Sub

Private Sub OPTYEARLY_Click()
If OPTYEARLY.Value = True Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
End Sub
