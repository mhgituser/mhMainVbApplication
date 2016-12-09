VERSION 5.00
Begin VB.Form frmemaillog 
   Caption         =   "Email Log"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   Icon            =   "frmemaillog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   2760
   End
End
Attribute VB_Name = "frmemaillog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DAILYRPT
End Sub

Private Sub DAILYRPT()
Dim SLNO As Integer
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rsch As New ADODB.Recordset
Dim actstring As String
Dim mstaff As String
Dim tt As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
mstaff = ""
db.Open OdkCnnString
                        



Dim SQLSTR As String
mchk = True
SQLSTR = ""
SLNO = 1

SQLSTR = "select * from dailyacthub9_core where  SUBSTRING( end ,1,10)>='" & Format(Now - 2, "yyyy-MM-dd") & "' and SUBSTRING( end ,1,10)<='" & Format(Now - 1, "yyyy-MM-dd") & "'  order by end "

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
    excel_app.Visible = False
    ' excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    
    excel_sheet.Cells(3, 2) = "DATE" & "(END)"
   
    excel_sheet.Cells(3, 3) = "NAME"
    excel_sheet.Cells(3, 4) = "SERVVYOR ID"
    excel_sheet.Cells(3, 5) = UCase("Yesterday's Activity")
    excel_sheet.Cells(3, 6) = UCase("No. of field visits")
    excel_sheet.Cells(3, 7) = UCase("No. of field failed")
    excel_sheet.Cells(3, 8) = UCase("Reason failed")
    excel_sheet.Cells(3, 9) = UCase("No. of storage visits")
    excel_sheet.Cells(3, 10) = UCase("No. of storage failed")
    excel_sheet.Cells(3, 11) = UCase("Reason failed")
    excel_sheet.Cells(3, 12) = UCase("Farmer registered")
    excel_sheet.Cells(3, 13) = UCase("Acre registered")
    excel_sheet.Cells(3, 14) = UCase("Travelling from")
    excel_sheet.Cells(3, 15) = UCase("Travelling to")
    excel_sheet.Cells(3, 16) = UCase("comments")
  
   i = 4
  Set rs = Nothing
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  chkred = False
 ' tt = "rs!" & CBODATE.Text
excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!End
'excel_sheet.Cells(i, 3) = IIf(IsNull(RS!sname), "", RS!sname)
excel_sheet.Cells(i, 4) = IIf(IsNull(rs!staffbarcode), "", rs!staffbarcode)
FindsTAFF excel_sheet.Cells(i, 4)

excel_sheet.Cells(i, 3) = sTAFF 'IIf(IsNull(RS!sname), "", RS!sname)
If chkred = True Then
 excel_sheet.Range(excel_sheet.Cells(i, 1), _
                             excel_sheet.Cells(i, 15)).Select
                             excel_app.Selection.Font.Color = vbRed
                             excel_sheet.Cells(i, 3) = rs!sname
                             
End If
chkred = False

'activity

Dim chcnt As Integer

Set RS1 = Nothing
RS1.Open "select * from dailyacthub9_activities where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If RS1.EOF <> True Then
Do While RS1.EOF <> True
Set rsch = Nothing
rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(RS1!Value), "", RS1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If

RS1.MoveNext
Loop

 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.Cells(i, 5) = actstring
Else

excel_sheet.Cells(i, 5) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
'activity ends here
excel_sheet.Cells(i, 6) = IIf(IsNull(rs!Field), "", rs!Field)
excel_sheet.Cells(i, 7) = IIf(IsNull(rs!nofailed), "", rs!nofailed)

'qc failed

Set RS1 = Nothing
RS1.Open "select * from dailyacthub9_qcfailed where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If RS1.EOF <> True Then
Do While RS1.EOF <> True
'actstring = IIf(IsNull(rs1!Value), "", rs1!Value) & "," & actstring

Set rsch = Nothing
rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(RS1!Value), "", RS1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If

RS1.MoveNext
Loop

 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.Cells(i, 8) = actstring
Else

excel_sheet.Cells(i, 8) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
' qc failed ends here
excel_sheet.Cells(i, 9) = IIf(IsNull(rs!Storage), "", rs!Storage)
excel_sheet.Cells(i, 10) = IIf(IsNull(rs!nofailed1), "", rs!nofailed1)
' storage  failed
Set RS1 = Nothing
RS1.Open "select * from dailyacthub9_qcfailed1 where _parent_auri='" & rs![_uri] & "' ", db
actstring = ""
If RS1.EOF <> True Then
Do While RS1.EOF <> True
'actstring = IIf(IsNull(rs1!Value), "", rs1!Value) & "," & actstring
Set rsch = Nothing
rsch.Open "select * from tbldailyactchoices where name='" & IIf(IsNull(RS1!Value), "", RS1!Value) & "' ", db
If rsch.EOF <> True Then
actstring = IIf(IsNull(rsch!label), "", rsch!label) & " # " & actstring
End If


RS1.MoveNext
Loop

 actstring = Left(actstring, Len(actstring) - 3)
excel_sheet.Cells(i, 11) = actstring
Else

excel_sheet.Cells(i, 11) = "" 'IIf(IsNull(RS!id), "", RS!id)
End If
'storage failed ends here
excel_sheet.Cells(i, 12) = IIf(IsNull(rs!registered), "", rs!registered)
excel_sheet.Cells(i, 13) = IIf(IsNull(rs!privateland), "", rs!privateland)
excel_sheet.Cells(i, 14) = IIf(IsNull(rs!travel1), "", rs!travel1)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!travel2), "", rs!travel2)
excel_sheet.Cells(i, 16) = IIf(IsNull(rs!Comments), "", rs!Comments)
SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop

   'make up


    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet

     excel_sheet.Range("A3:p3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "DAILY ACTIVITY"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With

excel_sheet.Columns("A:p").Select
 excel_app.Selection.ColumnWidth = 12
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
mFilepath = App.Path & "\reportlog\" & Format(Now, "yyMMdd") & " " & "DailyReport.xls"
        excel_app.ActiveWorkbook.SaveCopyAs FileName:=mFilepath
        excel_app.ActiveWorkbook.Close False
        excel_app.DisplayAlerts = False
        excel_app.Quit
     
Screen.MousePointer = vbDefault
Set excel_sheet = Nothing
Set excel_app = Nothing
db.Close

Dim emailmessage As String
Dim EMAILIDS As String

EMAILIDS = "muktitcc@gmail.com"
emailmessage = "daily report"
Dim mCONNECTION As String
Dim retVal          As String
mCONNECTION = "smtp.tashicell.com"
'mCONNECTION = "smtp1.btl.bt"
 retVal = SendMail(EMAILIDS, "Daily Activity Report", "noreply@mhv.com", _
          emailmessage, mCONNECTION, 25, _
          "habizabi", "habizabi", _
           mFilepath, CBool(False))

If retVal = "ok" Then
Else
MsgBox "Please Check Internet Connection " & retVal
End If



End Sub

