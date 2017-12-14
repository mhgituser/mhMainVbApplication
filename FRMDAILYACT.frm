VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMDAILYACT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAILY ACTIVITY"
   ClientHeight    =   3840
   ClientLeft      =   8190
   ClientTop       =   2520
   ClientWidth     =   5730
   Icon            =   "FRMDAILYACT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5730
   Begin VB.CheckBox chkmonthlyact 
      Caption         =   "MONTHLY ACTIVITY"
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
      Left            =   4320
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATE SELECTION"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton OPTALL 
         Caption         =   "ALL"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OPTSEL 
         Caption         =   "SELECTIVE"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1695
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
      Left            =   2520
      Picture         =   "FRMDAILYACT.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
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
      Left            =   840
      Picture         =   "FRMDAILYACT.frx":1434
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   5295
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131268609
         CurrentDate     =   41362
      End
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
         ItemData        =   "FRMDAILYACT.frx":1B9E
         Left            =   1080
         List            =   "FRMDAILYACT.frx":1BA0
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131268609
         CurrentDate     =   41362
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   900
      End
   End
   Begin MSDataListLib.DataCombo cbostaffcode 
      Bindings        =   "FRMDAILYACT.frx":1BA2
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1320
      TabIndex        =   12
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "MONITOR"
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
      Left            =   360
      TabIndex        =   14
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "STAFF ID"
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
      Left            =   -1080
      TabIndex        =   13
      Top             =   960
      Width           =   840
   End
End
Attribute VB_Name = "FRMDAILYACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mindex As Integer
Private Sub CBODATE_LostFocus()
Dim i, j, fcount As Integer

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                       
db.Open OdkCnnString
                       
Set rs = Nothing
rs.Open "select * from tbltable where tblid='38' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)

Set rs = Nothing
rs.Open "SELECT * FROM dailyacthub9_core where 1", CONNLOCAL
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

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If chkmonthlyact.Value = 0 Then
DAILYRPT
Else

ACTIVITYMONTHLY

End If

End Sub
Private Sub ACTIVITYMONTHLY()
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
Dim mtot(1 To 13), jtot As Double
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
                  
'If Len(cbostaffcode.Text) <> 0 Then
'SQLSTR = "select value as id ,count(value) as jval,year(end) as procyear,month(end) as procmonthfrom dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and staffbarcode='" & cbostaffcode.BoundText & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),value order by convert(substring(value,9,2) ,unsigned integer),year(end),month(end)"
'Else
''SQLSTR = "select staffbarcode as id ,count(staffbarcode) as jval,year(end) as procyear,month(end) as procmonth from dailyacthub9_core  where  end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),staffbarcode order by staffbarcode,year(end),month(end)"
'End If
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
    Set rs1 = Nothing
    rs1.Open "SELECT DISTINCT staffbarcode FROM dailyacthub9_core", db
    Do While rs1.EOF <> True
    SQLSTR = "select value as id ,count(value) as jval,year(end) as procyear,month(end) as procmonth from dailyacthub9_activities as a ,dailyacthub9_core as b  where  _parent_auri=b._uri and staffbarcode='" & rs1!staffbarcode & "' and end between '" & Format(DT1, "yyyy-MM-dd") & "' and '" & Format(DT2, "yyyy-MM-dd") & "' group by year(end),month(end),value order by convert(substring(value,9,2) ,unsigned integer),year(end),month(end)"
    Set rs = Nothing
    rs.Open SQLSTR, db, adOpenStatic, adLockReadOnly
    
    jCol = 3 - Month(txtfrmdate) + Month(txttodate)
    FindsTAFF rs1!staffbarcode
      excel_sheet.cells(jrow, 1) = "ACTIVITY"
       jrow = jrow + 1
    excel_sheet.cells(jrow, 1) = rs1!staffbarcode & " " & sTAFF
   
  
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.cells(jrow, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.cells(jrow, jCol) = UCase("Total")
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
       'jtot = 0
    Loop
   jtot = 0
    excel_sheet.cells(jrow + 1, 1) = UCase("Total")
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
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
        .PageSetup.LeftFooter = "mhv"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
  
    '.PageSetup.LeftHeader = "MHV"
     'excel_sheet.Range("A1:Aa15").Font.Bold = True
    
   excel_sheet.name = "Detail"

    Excel_WBook.Sheets("sheet2").Activate
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
    'excel_sheet.Cells(2, 1) = "MONTHLY ACTIVITY OF MONITOR " & rs!id & " " & sTAFF
    excel_sheet.cells(3, 1) = "MONITOR"
    K = 1
    For i = Month(txtfrmdate) To Month(txttodate)
        K = K + 1
        excel_sheet.cells(3, K) = UCase(Left(Mname(i), 3)) & "'" & CInt(Year(txtfrmdate.Value))
    Next
    excel_sheet.cells(3, jCol) = UCase("Total")
    excel_sheet.cells(3, jCol + 1) = ("Detail")
    
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
    excel_sheet.cells(jrow + 1, 1) = UCase("Total")
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
    
  
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A1:o3").Font.Bold = True
    






    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
    
End Sub
Private Sub DAILYRPT()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rsch As New ADODB.Recordset
Dim actstring As String
Dim mstaff As String
Dim tt As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
mstaff = ""
db.Open OdkCnnString
                        
If OPTSEL.Value = True And Len(CBODATE.Text) = 0 Then
MsgBox "Please Select The Date Type."
Exit Sub
End If

If OPTALL.Value = True Then
Mindex = 34
End If
Dim SQLSTR As String
mchk = True
SQLSTR = ""
SLNO = 1
If OPTALL.Value = True Then
If Len(cbostaffcode.Text) = 0 Then
SQLSTR = "select * from dailyacthub9_core order by end"
Else
SQLSTR = "select * from dailyacthub9_core WHERE staffbarcode='" & cbostaffcode.BoundText & "' order by end"
End If
ElseIf OPTSEL.Value = True Then
If Len(cbostaffcode.Text) = 0 Then
SQLSTR = "select * from dailyacthub9_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' ORDER BY staffbarcode "
Else
SQLSTR = "select * from dailyacthub9_core where  staffbarcode='" & cbostaffcode.BoundText & "' and SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' ORDER BY " & CBODATE.Text & " "
End If
Else
MsgBox "INVALIDE SELECTION OF OPTION"
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
    ' excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
    If OPTALL.Value = True Then
    excel_sheet.cells(3, 2) = "DATE" & "(END)"
    Else
    excel_sheet.cells(3, 2) = "DATE" & "(" & CBODATE.Text & ")"
    End If
    excel_sheet.cells(3, 3) = "NAME"
    excel_sheet.cells(3, 4) = "SURVEYOR ID"
    excel_sheet.cells(3, 5) = UCase("Yesterday's Activity")
    excel_sheet.cells(3, 6) = UCase("No. of field visits")
    excel_sheet.cells(3, 7) = UCase("No. of field failed")
    excel_sheet.cells(3, 8) = UCase("Reason failed")
    excel_sheet.cells(3, 9) = UCase("No. of storage visits")
    excel_sheet.cells(3, 10) = UCase("No. of storage failed")
    excel_sheet.cells(3, 11) = UCase("Reason failed")
    excel_sheet.cells(3, 12) = UCase("Farmer registered")
    excel_sheet.cells(3, 13) = UCase("Acre registered")
    excel_sheet.cells(3, 14) = UCase("Travelling from")
    excel_sheet.cells(3, 15) = UCase("Travelling to")
    excel_sheet.cells(3, 16) = UCase("comments")
  
   i = 4
  Set rs = Nothing
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
 ' tt = "rs!" & CBODATE.Text
excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = "'" & rs.Fields(Mindex)
'excel_sheet.Cells(i, 3) = IIf(IsNull(RS!sname), "", RS!sname)
excel_sheet.cells(i, 4) = IIf(IsNull(rs!staffbarcode), "", rs!staffbarcode)
FindsTAFF excel_sheet.cells(i, 4)

excel_sheet.cells(i, 3) = sTAFF 'IIf(IsNull(RS!sname), "", RS!sname)
If chkred = True Then
 excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 15)).Select
                             excel_app.selection.Font.Color = vbRed
                             excel_sheet.cells(i, 3) = rs!sname
                             
End If
chkred = False

'activity

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

 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 5) = actstring
Else

excel_sheet.cells(i, 5) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
'activity ends here
excel_sheet.cells(i, 6) = IIf(IsNull(rs!field), "", rs!field)
excel_sheet.cells(i, 7) = IIf(IsNull(rs!nofailed), "", rs!nofailed)

'qc failed

Set rs1 = Nothing
rs1.Open "select * from dailyacthub9_qcfailed where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
'actstring = IIf(IsNull(rs1!Value), "", rs1!Value) & "," & actstring

Set rsch = Nothing
rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If

rs1.MoveNext
Loop

 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 8) = actstring
Else

excel_sheet.cells(i, 8) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
' qc failed ends here
excel_sheet.cells(i, 9) = IIf(IsNull(rs!storage), "", rs!storage)
excel_sheet.cells(i, 10) = IIf(IsNull(rs!nofailed1), "", rs!nofailed1)
' storage  failed
Set rs1 = Nothing
rs1.Open "select * from dailyacthub9_qcfailed1 where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If rs1.EOF <> True Then
Do While rs1.EOF <> True
'actstring = IIf(IsNull(rs1!Value), "", rs1!Value) & "," & actstring
Set rsch = Nothing
rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(rs1!Value), "", rs1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If


rs1.MoveNext
Loop

 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.cells(i, 11) = actstring
Else

excel_sheet.cells(i, 11) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
'storage failed ends here
excel_sheet.cells(i, 12) = IIf(IsNull(rs!registered), "", rs!registered)
excel_sheet.cells(i, 13) = IIf(IsNull(rs!privateland), "", rs!privateland)
excel_sheet.cells(i, 14) = IIf(IsNull(rs!travel1), "", rs!travel1)
excel_sheet.cells(i, 15) = IIf(IsNull(rs!travel2), "", rs!travel2)
excel_sheet.cells(i, 16) = IIf(IsNull(rs!Comments), "", rs!Comments)
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
     excel_sheet.Range("A3:p3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "DAILY ACTIVITY"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("A:p").Select
 excel_app.selection.columnWidth = 12
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
Dim mFilepath As String
' MsgBox ExecuteExcel4Macro("Get.Document(50)")
'mFilepath = "C:\" & Format(Now, "yyMMdd") & "_Daily_Activity"
'ActiveWorkbook.SaveAs FileName:="C:\" & Format(Now, "yyMMdd") & "DailyActivity", FileFormat:= _
'xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
', CreateBackup:=False
'
'Application.DisplayAlerts = False 'RESETS DISPLAY ALERTS
'mFilepath = "C:\" & Format(Now, "yyMMdd") & "DailyActivity.xls"
'excel_app.Close SaveChanges:=False
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

db.Close
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear


'Dim emailmessage As String
'Dim EMAILIDS As String
'
'EMAILIDS = "muktitcc@gmail.com"
'
'       ' EMAILIDS = "muktitcc@gmail.com,muktitcc@gmail.com,sikkagurung@gmail.com"
'
'
'     emailmessage = "ODK DATA UPLOAD SUCCESSFULLY DONE ON " & Format(Now, "dd/MM/yyyy hh:mm:ss")
'
'   ' emailmessage = " SONAM, THERE IS AN ERROR ON UPLOADING ODK DATA,PLEASE CHECK. REPORTED ON " & Format(Now, "dd/MM/yyyy hh:mm:ss ERROR:  ") & UCase(ERR.Description)
'
'
'
'
'
'
'Dim mCONNECTION As String
'Dim retVal          As String
'mCONNECTION = "smtp.tashicell.com"
'
' retVal = SendMail(EMAILIDS, "Daily Activity Report", "NAS@MHV.COM", _
'    emailmessage, mCONNECTION, 25, _
'    "habizabi", "habizabi", _
'    mFilepath, CBool(False))
'
'If retVal = "ok" Then
'Else
'MsgBox "Please Check Internet Connection " & retVal
'End If



End Sub

Private Sub Form_Load()
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
Dim i, j, fcount As Integer
Dim Srs As New ADODB.Recordset
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
rs.Open "select * from tbltable where tblid='38' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)
CBODATE.Clear
Set rs = Nothing
rs.Open "SELECT * FROM dailyacthub9_core where 1", CONNLOCAL
For j = 0 To fcount - 1
If rs.Fields(j).Type = 135 Then
CBODATE.AddItem rs.Fields(j).name
End If
Next




db.Close
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where moniter='1' order by STAFFCODE", db
Set cbostaffcode.RowSource = Srs
cbostaffcode.ListField = "STAFFNAME"
cbostaffcode.BoundColumn = "STAFFCODE"


Mname = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")


Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
mchk = False
End Sub

Private Sub OPTALL_Click()
Frame1.Enabled = False
End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
End Sub
