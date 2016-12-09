VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMLANDDETAILSEL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAND DETAIL SELECTION MODE"
   ClientHeight    =   2475
   ClientLeft      =   7575
   ClientTop       =   1680
   ClientWidth     =   6375
   Icon            =   "FRMLANDDETAILSEL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6375
   Begin VB.CommandButton Command1 
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
      Left            =   1920
      Picture         =   "FRMLANDDETAILSEL.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
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
      Left            =   3480
      Picture         =   "FRMLANDDETAILSEL.frx":15AC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo cboDzongkhag 
      Bindings        =   "FRMLANDDETAILSEL.frx":2276
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSDataListLib.DataCombo cbogewog 
      Bindings        =   "FRMLANDDETAILSEL.frx":228B
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSDataListLib.DataCombo cbotshowog 
      Bindings        =   "FRMLANDDETAILSEL.frx":22A0
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
      _ExtentX        =   8493
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
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   720
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "TSHOWOG ID"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1230
   End
End
Attribute VB_Name = "FRMLANDDETAILSEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDz As New ADODB.Recordset
Dim rsGe As New ADODB.Recordset
Dim rsTs As New ADODB.Recordset



Private Sub cboDzongkhag_LostFocus()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog WHERE DZONGKHAGID='" & cbodzongkhag.BoundText & "' order by dzongkhagid,gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"

End Sub

Private Sub cbogewog_LostFocus()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
If rsTs.State = adStateOpen Then rsTs.Close
rsTs.Open "select concat(tshewogid , ' ', tshewogname) as tshewogname,tshewogid  from tbltshewog WHERE DZONGKHAGID='" & cbodzongkhag.BoundText & "' AND GEWOGID='" & cbogewog.BoundText & "' order by dzongkhagid,gewogid,tshewogid", db
Set CBOTSHOWOG.RowSource = rsTs
CBOTSHOWOG.ListField = "tshewogname"
CBOTSHOWOG.BoundColumn = "tshewogid"
End Sub

Private Sub Command1_Click()
Select Case RptOption
            Case "LS"
            lsexcel1
            Case "CD"
            CD

End Select
End Sub
Private Sub CD()

Dim SQLSTR As String
Dim j As Integer
Dim MSTR As String

MSTR = cbodzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText

If Len(MSTR) = 9 Then
SQLSTR = "SELECT * FROM tblcontact WHERE SUBSTRING(contactid,1,9)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "ORDER BY roleid,dzongkhagid,gewogid,tshowogid,contactid"

Mge = False
ElseIf Len(MSTR) = 6 Then

SQLSTR = "SELECT * FROM tblcontact WHERE SUBSTRING(contactid,1,6)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "ORDER BY roleid,dzongkhagid,gewogid,tshowogid,contactid"


Mge = True
ElseIf Len(MSTR) = 3 Then

SQLSTR = "SELECT * FROM tblcontact WHERE SUBSTRING(contactid,1,3)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "ORDER BY roleid,dzongkhagid,gewogid,tshowogid,contactid"




ElseIf Len(MSTR) = 0 Then
SQLSTR = "SELECT * FROM tblcontact "
SQLSTR = SQLSTR & "  " & "ORDER BY roleid,dzongkhagid,gewogid,tshowogid,contactid"
Else
MsgBox "Invalid Selection."
Mge = False
Exit Sub
End If


Dim rs As New ADODB.Recordset
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
    excel_sheet.Cells(3, 2) = "ROLE"
    excel_sheet.Cells(3, 3) = "DZONGKHAG"
    excel_sheet.Cells(3, 4) = "GEWOG"
    excel_sheet.Cells(3, 5) = "TSHOWOG"
    excel_sheet.Cells(3, 6) = "CONTACT NAME"
    excel_sheet.Cells(3, 7) = "PHONE(WORK)"
    excel_sheet.Cells(3, 8) = "PHONE(RESIDENCE)"
    excel_sheet.Cells(3, 9) = "MOBILE"
    excel_sheet.Cells(3, 10) = "EMAIL"
    excel_sheet.Cells(3, 11) = "LOCATION DESCRIPTION"
    excel_sheet.Cells(3, 12) = "DEPARTMENT"
    excel_sheet.Cells(3, 13) = "RELATIVES"
    excel_sheet.Cells(3, 14) = "OTHER NOTES"
    
  
    i = 4
  Set rs = Nothing
  rs.Open SQLSTR, MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
   FindROLL rs!roleid
   excel_sheet.Cells(i, 2) = rs!roleid & " " & rOLEnAME
   
   
   FindDZ Mid(rs!CONTACTID, 1, 3)
 excel_sheet.Cells(i, 3) = Mid(rs!CONTACTID, 1, 3) & " " & Dzname

If (Mid(rs!CONTACTID, 4, 1)) = "G" Then
FindGE Mid(rs!CONTACTID, 1, 3), Mid(rs!CONTACTID, 4, 3)
 excel_sheet.Cells(i, 4) = Mid(rs!CONTACTID, 4, 3) & " " & GEname
End If

If (Mid(rs!CONTACTID, 4, 1)) = "G" And Mid(rs!CONTACTID, 7, 1) = "T" Then
FindTs Mid(rs!CONTACTID, 1, 3), Mid(rs!CONTACTID, 4, 3), Mid(rs!CONTACTID, 7, 3)
excel_sheet.Cells(i, 5) = Mid(rs!CONTACTID, 7, 3) & " " & TsName
End If

   
   
   excel_sheet.Cells(i, 6) = IIf(IsNull(rs!CONTACTID), "", rs!CONTACTID) & " " & IIf(IsNull(rs!firstname), "", rs!firstname) & " " & IIf(IsNull(rs!secondname), "", rs!secondname)
   excel_sheet.Cells(i, 7) = IIf(IsNull(rs!phonework), "", rs!phonework)
   excel_sheet.Cells(i, 8) = IIf(IsNull(rs!phonehome), "", rs!phonehome)
   excel_sheet.Cells(i, 9) = IIf(IsNull(rs!mobile), "", rs!mobile)
   excel_sheet.Cells(i, 10) = IIf(IsNull(rs!email), "", rs!email)
   excel_sheet.Cells(i, 11) = IIf(IsNull(rs!location), "", rs!location)
    excel_sheet.Cells(i, 12) = IIf(IsNull(rs!Dept), "", rs!Dept)
   excel_sheet.Cells(i, 13) = IIf(IsNull(rs!relatives), "", rs!relatives)
   excel_sheet.Cells(i, 14) = IIf(IsNull(rs!importaintnote), "", rs!importaintnote)
     
   sl = sl + 1
   i = i + 1
   rs.MoveNext
   Loop
   
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
     excel_sheet.Range("A3:N3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "CONTACT LIST"
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
Exit Sub
err:
MsgBox err.Description
err.Clear



End Sub
Private Sub lsexcel1()
Dim SQLSTR As String
Dim j As Integer
Dim MSTR As String

MSTR = cbodzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText

If Len(MSTR) = 9 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGAREA)+(REGLAND)) AS REGLAND,houseno,thramno,village FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID  and a.status='A' AND SUBSTRING(IDFARMER,1,9)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Mge = False
ElseIf Len(MSTR) = 6 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGAREA)+(REGLAND)) AS REGLAND,houseno,thramno,village FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID and a.status='A' AND SUBSTRING(IDFARMER,1,6)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Mge = True
ElseIf Len(MSTR) = 3 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGLAND)) AS REGLAND,houseno,thramno,village FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID and a.status='A' AND SUBSTRING(IDFARMER,1,3)='" & MSTR & "' AND REGLAND<>0"
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER "

ElseIf Len(MSTR) = 0 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGAREA)+(REGLAND)) AS REGLAND,houseno,thramno,village FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID and a.status='A' AND  REGLAND IS NULL"
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER "

Else
MsgBox "Invalid Selection."
Mge = False
Exit Sub
End If
chkred = True
mchk = True

Dim rs As New ADODB.Recordset
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
    excel_sheet.Cells(3, 2) = "DZONGKHAG CODE"
    excel_sheet.Cells(3, 3) = "DZONGKHAG NAME"
    excel_sheet.Cells(3, 4) = "GEWOG CODE"
    excel_sheet.Cells(3, 5) = "GEWOG NAME"
    excel_sheet.Cells(3, 6) = "FAREMER CODE"
    excel_sheet.Cells(3, 6) = "FARMER NAME"
    excel_sheet.Cells(3, 8) = "TOTAL LAND REGISTERED(ACRE)"
    excel_sheet.Cells(3, 9) = "HOUSE NO."
    excel_sheet.Cells(3, 10) = "THRAM NO."
    excel_sheet.Cells(3, 11) = "VILLAGE"
     excel_sheet.Cells(3, 12) = "TSHOWOG"
     
      i = 4
  Set rs = Nothing
  rs.Open SQLSTR, MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
     excel_sheet.Cells(i, 2) = rs!dzcode
     FindDZ Mid(rs!idfarmer, 1, 3)
   excel_sheet.Cells(i, 3) = Dzname
 excel_sheet.Cells(i, 4) = rs!GECODE
 FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
  excel_sheet.Cells(i, 5) = GEname
   excel_sheet.Cells(i, 6) = rs!idfarmer
   FindFA rs!idfarmer, "F"
     excel_sheet.Cells(i, 7) = FAName
    excel_sheet.Cells(i, 8) = IIf(IsNull(rs!regland), 0, rs!regland)
     excel_sheet.Cells(i, 9) = IIf(IsNull(rs!houseno), "", rs!houseno)
      excel_sheet.Cells(i, 10) = IIf(IsNull(rs!thramno), 0, rs!thramno)
       excel_sheet.Cells(i, 11) = IIf(IsNull(rs!VILLAGE), 0, rs!VILLAGE)
    FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
    excel_sheet.Cells(i, 12) = TsName
    
   i = i + 1
   sl = sl + 1
   rs.MoveNext
   Loop
   
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
     excel_sheet.Range("A3:f3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "LISTING"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
    excel_app.Visible = True
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault


Screen.MousePointer = vbDefault
Exit Sub
err:
MsgBox err.Description
err.Clear

End Sub

Private Sub Command3_Click()
Select Case RptOption
    Case "LS"
      lsacr
    

End Select


End Sub
Private Sub lsacr()
Dim SQLSTR As String
Dim j As Integer
Dim MSTR As String

MSTR = cbodzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText

'If Len(MSTR) = 9 Then
'
'Else
'MsgBox "INVALID SELECTION."
'Exit Sub
'End If
'J = 0
'For i = 0 To DZLIST.ListCount - 1
'    If DZLIST.Selected(i) Then
'       DZstr = DZstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
'       Mcat = DZLIST.List(i)
'       J = J + 1
'    End If
'    If RepName = "5" Then
'       If J > 1 Then
'          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
'          Exit Sub
'       End If
'    End If
'Next
'If Len(DZstr) > 0 Then
'   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
'   'SQLSTR = SQLSTR + " AND c.Itemcode in " + Acntstr
'Else
'   MsgBox "DZONGKHAG NOT SELECTED !!!"
'   Exit Sub
'End If
If Len(MSTR) = 9 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGAREA)+(REGLAND)) AS REGLAND FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID AND SUBSTRING(IDFARMER,1,9)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"
Mge = False
ElseIf Len(MSTR) = 6 Then
SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,SUM((REGAREA)+(REGLAND)) AS REGLAND FROM tblfarmer A,tbllandreg B WHERE A.IDFARMER=B.FARMERID AND SUBSTRING(IDFARMER,1,6)='" & MSTR & "'"
SQLSTR = SQLSTR & "  " & "GROUP BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER,FARMERNAME ORDER BY SUBSTRING(IDFARMER,1,3) ,SUBSTRING(IDFARMER,4,3) ,SUBSTRING(IDFARMER,8,3) ,IDFARMER"

Mge = True
Else
MsgBox "Invalid Selection."
Mge = False
Exit Sub

End If


                           

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Mge = False
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cbodzongkhag.RowSource = rsDz
cbodzongkhag.ListField = "dzongkhagname"
cbodzongkhag.BoundColumn = "dzongkhagcode"



End Sub
