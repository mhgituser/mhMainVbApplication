VERSION 5.00
Begin VB.Form FRMFIELDERROR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR CHECK LIST"
   ClientHeight    =   4110
   ClientLeft      =   6345
   ClientTop       =   3045
   ClientWidth     =   6630
   Icon            =   "FRMFIELDERROR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6630
   Begin VB.OptionButton Option7 
      Caption         =   "MONITORS TSHOWOG CHECK LIST"
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
      Top             =   2760
      Width           =   6015
   End
   Begin VB.OptionButton Option6 
      Caption         =   "PLANTED LIST FARMERS WITHOUT MONITORS"
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
      TabIndex        =   7
      Top             =   2400
      Width           =   6015
   End
   Begin VB.OptionButton Option5 
      Caption         =   "FARMER WITH INVALID CODE (DGTF AND BARCODE NULL) IN ODK FIELD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   6015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "FARMER CODE IN REGISTRATION BUT NOT IN ODK FIELD"
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
      TabIndex        =   5
      Top             =   1080
      Width           =   6015
   End
   Begin VB.OptionButton Option4 
      Caption         =   "FARMER CODE IN PLANTED LIST  BUT NOT IN  REGISTRATION"
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
      TabIndex        =   4
      Top             =   360
      Width           =   6015
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
      Left            =   1560
      Picture         =   "FRMFIELDERROR.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
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
      Left            =   3240
      Picture         =   "FRMFIELDERROR.frx":0ED4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "MONITOR CODE IN ODK FIELD BUT NOT IN  REGISTRATION"
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
      Top             =   2040
      Width           =   6015
   End
   Begin VB.OptionButton Option2 
      Caption         =   "FARMER CODE IN ODK FIELD BUT NOT IN  REGISTRATION"
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
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
End
Attribute VB_Name = "FRMFIELDERROR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
'On Error GoTo err
Select Case RptOption
       Case "1"
       REGVSPLST
       Case "2"
       ODKVSREG
       Case "3"
       plantedvsodk
        Case "4"
       mm
       Case "5"
       INVALIDFCODEODK
       Case "6"
       minitornotassigned
       Case "7"
       monitortshowogchecklist
End Select
'Exit Sub
'err:
'MsgBox err.Description
End Sub
Private Sub monitortshowogchecklist()
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsm As New ADODB.Recordset

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
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "S/N"
    excel_sheet.Cells(3, 2) = "DGT Code"
    excel_sheet.Cells(3, 3) = "Monitor"
    i = 4





Set rs = Nothing
MHVDB.Execute "delete from tbltmpmontshowog"
MHVDB.Execute "insert into tbltmpmontshowog " _
& " SELECT DISTINCT monitor, SUBSTRING( idfarmer, 1, 9 )" _
& " From tblfarmer " _
& " WHERE LENGTH( monitor ) =5 order by monitor"


rs.Open "SELECT * from tbltmpmontshowog group by dgt having count(dgt)>1 ", MHVDB

If rs.EOF <> True Then

Do While rs.EOF <> True
excel_sheet.Cells(i, 1) = sl
Set rsm = Nothing
rsm.Open "select * from tbltmpmontshowog where dgt='" & rs!dgt & "'", MHVDB
Do While rsm.EOF <> True
FindsTAFF rsm!staffcode
excel_sheet.Cells(i, 2) = rsm!dgt
excel_sheet.Cells(i, 3) = rsm!staffcode & "  " & sTAFF
i = i + 1
rsm.MoveNext
Loop
  


i = i + 1
sl = sl + 1

rs.MoveNext
Loop











Else

'MsgBox "uuummmm"
End If


'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 7)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:G3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ERROR LISTING"
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
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear
End Sub
Private Sub minitornotassigned()
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset

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
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
  excel_sheet.Cells(3, 4) = "TSHOWOG"
  excel_sheet.Cells(3, 5) = "FARMER CODE"
  excel_sheet.Cells(3, 6) = "FARMER NAME"
  excel_sheet.Cells(3, 7) = "MONITOR"
    i = 4





Set rs = Nothing
rs.Open "SELECT distinct farmercode FROM tblplanted where farmercode in(select idfarmer from tblfarmer where status='A' and length(monitor)<>'5') and status<>'C' group by farmercode order by farmercode ", MHVDB

If rs.EOF <> True Then

Do While rs.EOF <> True

excel_sheet.Cells(i, 1) = sl
FindDZ Mid(rs!farmercode, 1, 3)
    excel_sheet.Cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
   FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
    excel_sheet.Cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
  excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
  
  excel_sheet.Cells(i, 5) = rs!farmercode
  FindFA rs!farmercode, "F"
  excel_sheet.Cells(i, 6) = rs!farmercode & " " & FAName
  
  
excel_sheet.Cells(i, 7) = ""

i = i + 1
sl = sl + 1

rs.MoveNext
Loop











Else

'MsgBox "uuummmm"
End If


'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 7)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:G3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ERROR LISTING"
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
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear

End Sub
Private Sub INVALIDFCODEODK()
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
                     
GetTbl

mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1




  
  
  
  
  
  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DATE "
    excel_sheet.Cells(3, 3) = "MONITOR ID"
     excel_sheet.Cells(3, 4) = UCase("MONITOR NAME")
     excel_sheet.Cells(3, 5) = "DZONGKHAG"
     
      excel_sheet.Cells(3, 6) = "GEWOG"
       excel_sheet.Cells(3, 7) = "TSHOWOG"
   
    excel_sheet.Cells(3, 8) = UCase("FARMER NAME")
    excel_sheet.Cells(3, 9) = UCase("FIELD/STORAGE")
    
   i = 4
  Set rs = Nothing
  SQLSTR = ""
  SQLSTR = "SELECT farmerbarcode,end,staffid,fname,'STORAGE' AS fs FROM tblstorageqc_core WHERE status<>'BAD' and  substring(farmerbarcode,10,5)='F0000' union SELECT farmerbarcode,end,staffid,fname,'FIELD' as fs FROM tblfieldqc_core WHERE status<>'BAD' and  substring(farmerbarcode,10,5)='F0000'"
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  FindDZ Mid(rs!farmerbarcode, 1, 3)
  FindGE Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3)
  FindTs Mid(rs!farmerbarcode, 1, 3), Mid(rs!farmerbarcode, 4, 3), Mid(rs!farmerbarcode, 7, 3)
  
  
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!End  'rs.Fields(Mindex)
excel_sheet.Cells(i, 3) = IIf(IsNull(rs!staffid), "", rs!staffid)
excel_sheet.Cells(i, 4) = IIf(IsNull(rs!staff_name), "", rs!staff_name)
excel_sheet.Cells(i, 5) = Mid(rs!farmerbarcode, 1, 3) & " " & Dzname
excel_sheet.Cells(i, 6) = Mid(rs!farmerbarcode, 4, 3) & " " & GEname
excel_sheet.Cells(i, 7) = Mid(rs!farmerbarcode, 7, 3) & " " & TsName
excel_sheet.Cells(i, 8) = IIf(IsNull(rs!fname), "", rs!fname)
excel_sheet.Cells(i, 9) = IIf(IsNull(rs!fs), "", rs!fs)
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
db.Close
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear
End Sub
Private Sub mm()
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
                       
mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1




  
  
  
  
  
  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DATE "
    excel_sheet.Cells(3, 3) = "MONITOR ID"
    excel_sheet.Cells(3, 4) = UCase("MONITOR NAME")
     excel_sheet.Cells(3, 5) = UCase("field/storage")
   i = 4
  Set rs = Nothing
  SQLSTR = ""
  SQLSTR = "SELECT end,staffid,'FIELD' as fs FROM tblfieldqc_core WHERE status<>'BAD' and staffid not in(select substring(staffcode,1,5) from MHV.tblmhvstaff) union SELECT end,staffid,'STORAGE' as fs FROM tblstorageqc_core WHERE status<>'BAD' and staffid not in(select substring(staffcode,1,5) from MHV.tblmhvstaff) "
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  
  
  
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!End  'rs.Fields(Mindex)
excel_sheet.Cells(i, 3) = IIf(IsNull(rs!staffid), "", rs!staffid)
excel_sheet.Cells(i, 4) = IIf(IsNull(rs!staff_name), "", rs!staff_name)
excel_sheet.Cells(i, 5) = IIf(IsNull(rs!fs), "", rs!fs)

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
db.Close
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear
End Sub
Private Sub plantedvsodk()
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
                       
GetTbl

mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1


 
     SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,fstype) select n.end,staffid,dcode," _
         & "gcode,tcode,n.farmerbarcode,'FIELD' from tblfieldqc_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM tblfieldqc_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"
     db.Execute SQLSTR
     SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,fstype) select n.end,staffid,dcode," _
         & "gcode,tcode,n.farmerbarcode,'STORAGE' from tblstorageqc_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM tblstorageqc_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD'GROUP BY n.farmerbarcode"
    
    
    
  db.Execute SQLSTR

  
  
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
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DATE "
    excel_sheet.Cells(3, 3) = "DZONGKHAG"
    excel_sheet.Cells(3, 4) = "GEWOG"
    excel_sheet.Cells(3, 5) = "TSHOWOG"
    excel_sheet.Cells(3, 6) = "FARMERCODE"
    excel_sheet.Cells(3, 7) = "FARMER NAME"
    excel_sheet.Cells(3, 8) = UCase("MONITOR")
'    excel_sheet.Cells(3, 9) = UCase("FIELD/STORAGE")
    mchk = True
    chkred = True
   i = 4
  Set rs = Nothing
  SQLSTR = ""
  'SQLSTR = "SELECT * FROM MHV.tblfarmer WHERE status='A' and idfarmer not in(select farmercode from " & Mtblname & ") group by idfarmer"
  SQLSTR = "SELECT distinct farmercode as idfarmer FROM MHV.tblplanted WHERE  farmercode not in(select farmercode from " & Mtblname & ") group by idfarmer"
  rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
'excel_sheet.Cells(i, 2) = "'" & rs!End  'rs.Fields(Mindex)
FindDZ Mid(rs!idfarmer, 1, 3)
excel_sheet.Cells(i, 3) = Mid(rs!idfarmer, 1, 3) & " " & Dzname
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
excel_sheet.Cells(i, 4) = Mid(rs!idfarmer, 4, 3) & " " & GEname
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
excel_sheet.Cells(i, 5) = Mid(rs!idfarmer, 7, 3) & " " & TsName
excel_sheet.Cells(i, 6) = IIf(IsNull(rs!idfarmer), "", rs!idfarmer)
FindFA rs!idfarmer, "F"
excel_sheet.Cells(i, 7) = FAName

'excel_sheet.Cells(i, 8) = rs!monitor

Set rss = Nothing
rss.Open "select  monitor from tblfarmer where idfarmer='" & rs!idfarmer & "'", MHVDB
If rss.EOF <> True Then
FindsTAFF rss!monitor
excel_sheet.Cells(i, 8) = rss!monitor & " " & sTAFF
Else
excel_sheet.Cells(i, 8) = ""
End If
'excel_sheet.Cells(i, 9) = rs!fstype
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
'err:
'db.Execute "drop table " & Mtblname & ""
'MsgBox err.Description
'err.Clear
End Sub
Private Sub ODKVSREG()
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
                       
GetTbl

mchk = True
Dim SQLSTR As String
SQLSTR = ""
SLNO = 1


    

    
          SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,fstype,id,fname) select n.end,dcode," _
         & "gcode,tcode,n.farmerbarcode,'FIELD',staffid,fname from tblfieldqc_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM tblfieldqc_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
       db.Execute SQLSTR
    
    
      SQLSTR = "insert into " & Mtblname & " (end,dcode,gcode,tcode,farmercode,fstype,id,fname) select n.end,dcode," _
         & "gcode,tcode,n.farmerbarcode,'STORAGE',staffid,fname from tblstorageqc_core n INNER JOIN (SELECT farmerbarcode, MAX(END )" _
         & "lastEdit FROM tblstorageqc_core GROUP BY farmerbarcode)x ON " _
         & "n.farmerbarcode = x.farmerbarcode AND n.end = x.LastEdit " _
         & "AND STATUS <>  'BAD' GROUP BY n.farmerbarcode"
       db.Execute SQLSTR


  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DATE "
    excel_sheet.Cells(3, 3) = "DZONGKHAG"
    excel_sheet.Cells(3, 4) = "GEWOG"
    excel_sheet.Cells(3, 5) = "TSHOWOG"
    excel_sheet.Cells(3, 6) = "FARMERCODE"
    excel_sheet.Cells(3, 7) = "COUNT"
    excel_sheet.Cells(3, 8) = UCase("MONITOR")
      excel_sheet.Cells(3, 9) = UCase("field/storage")
   i = 4
  Set rs = Nothing
  SQLSTR = ""
  SQLSTR = "SELECT fstype,end,id,fname,farmercode,count(farmercode) as cnt FROM " & Mtblname & " WHERE farmercode not in(select idfarmer from MHV.tblfarmer) group by farmercode"
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  
  
  
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!End  'rs.Fields(Mindex)
FindDZ Mid(rs!farmercode, 1, 3)
excel_sheet.Cells(i, 3) = Mid(rs!farmercode, 1, 3) & " " & Dzname
FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 4, 3) & " " & GEname
FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
excel_sheet.Cells(i, 5) = Mid(rs!farmercode, 7, 3) & " " & TsName

excel_sheet.Cells(i, 6) = IIf(IsNull(rs!farmercode), "", rs!farmercode) & " " & rs!fname
excel_sheet.Cells(i, 7) = IIf(IsNull(rs!cnt), "", rs!cnt)
FindsTAFF rs!id
excel_sheet.Cells(i, 8) = rs!id & " " & sTAFF

 excel_sheet.Cells(i, 9) = rs!fstype

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
Private Sub REGVSPLST()
Dim excel_app As Object
Dim excel_sheet As Object
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset

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
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
  excel_sheet.Cells(3, 4) = "TSHOWOG"
  excel_sheet.Cells(3, 5) = "FARMER CODE"
  excel_sheet.Cells(3, 6) = "FARMER NAME"
  excel_sheet.Cells(3, 7) = "MONITOR"
    i = 4





Set rs = Nothing
rs.Open "SELECT distinct * FROM tblplanted where farmercode not in(select idfarmer from tblfarmer)group by farmercode order by farmercode ", MHVDB

If rs.EOF <> True Then

Do While rs.EOF <> True

excel_sheet.Cells(i, 1) = sl
FindDZ Mid(rs!farmercode, 1, 3)
    excel_sheet.Cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
   FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
    excel_sheet.Cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
  excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
  
  excel_sheet.Cells(i, 5) = rs!farmercode
  FindFA rs!farmercode, "F"
  excel_sheet.Cells(i, 6) = rs!farmercode & " " & FAName
  
  Set rss = Nothing
rss.Open "select  monitor from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
If rss.EOF <> True Then
FindsTAFF rss!monitor
excel_sheet.Cells(i, 7) = rs!monitor & " " & sTAFF
Else
excel_sheet.Cells(i, 7) = ""
End If
i = i + 1
sl = sl + 1

rs.MoveNext
Loop











Else

'MsgBox "uuummmm"
End If


'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 7)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:G3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ERROR LISTING"
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
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear





End Sub
Private Sub Fnotinfield()
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
                        


Dim SQLSTR As String
SQLSTR = ""
SLNO = 1



SQLSTR = ""


    
       SQLSTR = "insert into " & Mtblname & " (end,id,dcode,gcode,tcode,farmercode,treesreceived,fdcode,totaltrees," _
         & "goodmoisture,poormoisture,totaltally,deadmissing,slowgrowing,dor,activegrowing,shock,nutrient," _
         & "waterlog,leafpest,activepest,stempest,rootpest,animaldamage,monitorcomments,fname) select end,id,dcode," _
         & "gcode,tcode,farmerbarcode,treesreceived,fdcode,totaltrees,goodmoisture,poormoisture,totaltally," _
         & "deadmissing,slowgrowing,dor,activegrowing,shock,nutrient,waterlog,leafpest,activepest,stempest," _
         & "rootpest,animaldamage,monitorcomments,fname from phealthhub15_core where  status<>'BAD' "
         
    
    
    
    
  db.Execute SQLSTR

  
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    
    
    excel_sheet.Cells(3, 2) = "DATE "
   
     excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "FARMERCODE"
    excel_sheet.Cells(3, 6) = "FARMER NAME"
    excel_sheet.Cells(3, 7) = UCase("Farmer ID")
    excel_sheet.Cells(3, 8) = UCase("Total Distributed")
    excel_sheet.Cells(3, 9) = UCase("Field ID")
    excel_sheet.Cells(3, 10) = UCase("Total Trees Distributed - Planted List")
    excel_sheet.Cells(3, 11) = UCase("Total Trees")
    excel_sheet.Cells(3, 12) = UCase("Good Moisture")
    excel_sheet.Cells(3, 13) = UCase("Poor Moisture")
    excel_sheet.Cells(3, 14) = UCase("Total Mositure Tally")
    excel_sheet.Cells(3, 15) = UCase("Dead Missing")
    excel_sheet.Cells(3, 16) = UCase("Slow Growing")
    excel_sheet.Cells(3, 17) = UCase("Dormant")
    excel_sheet.Cells(3, 18) = UCase("Active Growing")
    excel_sheet.Cells(3, 19) = UCase("Shock")
    excel_sheet.Cells(3, 20) = UCase("Nutrient Deficient")
    excel_sheet.Cells(3, 21) = UCase("Water Logg")
    excel_sheet.Cells(3, 22) = UCase("Leaf Pest")
    excel_sheet.Cells(3, 23) = UCase("Active Pest")
    excel_sheet.Cells(3, 24) = UCase("Stem Pest")
    excel_sheet.Cells(3, 25) = UCase("Root Pest")
    excel_sheet.Cells(3, 26) = UCase("Animal Damage")
    excel_sheet.Cells(3, 27) = UCase("comments")
   i = 4
  Set rs = Nothing
  
rs.Open SQLSTR, db
  Do While rs.EOF <> True
  

excel_sheet.Cells(i, 1) = SLNO
excel_sheet.Cells(i, 2) = "'" & rs!End  'rs.Fields(Mindex)
excel_sheet.Cells(i, 3) = rs!id

excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 1, 3)
excel_sheet.Cells(i, 5) = Mid(rs!farmercode, 4, 3)
excel_sheet.Cells(i, 6) = Mid(rs!farmercode, 7, 3)
excel_sheet.Cells(i, 7) = IIf(IsNull(rs!farmercode), "", rs!farmercode)


excel_sheet.Cells(i, 8) = IIf(IsNull(rs!treesreceived), "", rs!treesreceived)
excel_sheet.Cells(i, 9) = IIf(IsNull(rs!FDCODE), "", rs!FDCODE)
excel_sheet.Cells(i, 10) = ""
excel_sheet.Cells(i, 11) = IIf(IsNull(rs!totaltrees), 0, rs!totaltrees)
excel_sheet.Cells(i, 12) = IIf(IsNull(rs!goodmoisture), "", rs!goodmoisture)
excel_sheet.Cells(i, 13) = IIf(IsNull(rs!poormoisture), "", rs!poormoisture)
excel_sheet.Cells(i, 14) = IIf(IsNull(rs!totaltally), "", rs!totaltally)
excel_sheet.Cells(i, 15) = IIf(IsNull(rs!deadmissing), "", rs!deadmissing)
excel_sheet.Cells(i, 16) = IIf(IsNull(rs!slowgrowing), "", rs!slowgrowing)
excel_sheet.Cells(i, 17) = IIf(IsNull(rs!dor), "", rs!dor)
excel_sheet.Cells(i, 18) = IIf(IsNull(rs!activegrowing), "", rs!activegrowing)
excel_sheet.Cells(i, 19) = IIf(IsNull(rs!shock), "", rs!shock)
excel_sheet.Cells(i, 20) = IIf(IsNull(rs!nutrient), "", rs!nutrient)
excel_sheet.Cells(i, 21) = IIf(IsNull(rs!waterlog), "", rs!waterlog)
excel_sheet.Cells(i, 22) = IIf(IsNull(rs!leafpest), "", rs!leafpest)
excel_sheet.Cells(i, 23) = IIf(IsNull(rs!activepest), "", rs!activepest)
excel_sheet.Cells(i, 24) = IIf(IsNull(rs!stempest), "", rs!stempest)
excel_sheet.Cells(i, 25) = IIf(IsNull(rs!rootpest), "", rs!rootpest)
excel_sheet.Cells(i, 26) = IIf(IsNull(rs!animaldamage), "", rs!animaldamage)
excel_sheet.Cells(i, 27) = IIf(IsNull(rs!monitorcomments), "", rs!monitorcomments)


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



Private Sub Option1_Click()
RptOption = "3"
End Sub

Private Sub Option2_Click()
RptOption = "2"
End Sub

Private Sub Option3_Click()
RptOption = "4"
End Sub

Private Sub Option4_Click()
RptOption = "1"
End Sub

Private Sub Option5_Click()
RptOption = "5"
End Sub

Private Sub Option6_Click()
RptOption = "6"
End Sub

Private Sub Option7_Click()
RptOption = "7"
End Sub
