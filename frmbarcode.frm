VERSION 5.00
Begin VB.Form frmbarcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QR CODE PRINTING"
   ClientHeight    =   5445
   ClientLeft      =   3435
   ClientTop       =   2070
   ClientWidth     =   14685
   Icon            =   "frmbarcode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   14685
   Begin VB.CheckBox CHKGT 
      Caption         =   "MONITOR WISE"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   4320
      Width           =   2655
   End
   Begin VB.ListBox LSTFARMER 
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
      ItemData        =   "frmbarcode.frx":0E42
      Left            =   9960
      List            =   "frmbarcode.frx":0E49
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   0
      Width           =   4575
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
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   600
      Width           =   3375
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
      Left            =   1680
      Picture         =   "frmbarcode.frx":0E58
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BAR ME!"
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
      Left            =   0
      Picture         =   "frmbarcode.frx":1B22
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox CHKIF 
      Caption         =   "INDIVIDUAL FARMER"
      Height          =   195
      Left            =   7680
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
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
      ItemData        =   "frmbarcode.frx":228C
      Left            =   3720
      List            =   "frmbarcode.frx":228E
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
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
      Left            =   10320
      TabIndex        =   9
      Top             =   5160
      Width           =   75
   End
   Begin VB.Label lblcnt 
      AutoSize        =   -1  'True
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
      Left            =   4080
      TabIndex        =   8
      Top             =   5160
      Width           =   75
   End
   Begin VB.Line Line2 
      BorderWidth     =   10
      X1              =   9840
      X2              =   9840
      Y1              =   0
      Y2              =   5400
   End
   Begin VB.Label Label1 
      Caption         =   "SUPERVISOR"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   5400
   End
End
Attribute VB_Name = "frmbarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dzstr As String

Private Sub CHKGT_Click()
Dzstr = ""

If CHKGT.Value = 1 Then
frmbarcode.Width = 9840

For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
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
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
frmbarcode.Width = 3570
CHKGT.Value = 0
CHKIF.Value = 0
LSTPR.Clear
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTPR.Clear
'rs.Open "select * from tbltshewog where dzongkhagid in " & DZstr & "order by dzongkhagid,gewogid,tshewogid", MHVDB, adOpenStatic
rs.Open "select * from tblmhvstaff where msupervisor in " & Dzstr & "order by staffcode", MHVDB, adOpenStatic

'With rs
'Do While Not .EOF
'FindDZ rs!dzongkhagid
'FindGE rs!dzongkhagid, rs!gewogid
'
'
'
'
'
'   LSTPR.AddItem rs!dzongkhagid & " " & Trim(Dzname) + " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
'   .MoveNext
'Loop
'End With



With rs
Do While Not .EOF

   LSTPR.AddItem rs!staffname + " |" + rs!staffcode '+ " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With







Else
frmbarcode.Width = 3570
End If

End Sub

Private Sub CHPIF_Click()

End Sub

Private Sub CHKIF_Click()
Dim mm

Dzstr = ""

If CHKIF.Value = 1 Then
frmbarcode.Width = 14775

For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    mm = Split(LSTPR.List(i), "|", -1, 1)
       'DZstr = DZstr & "'" & Mid(MM(0), 1, 3) & Mid(MM(1), 1, 3) & Mid(MM(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
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
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
frmbarcode.Width = 9840
CHKIF.Value = 0
LSTFARMER.Clear
   MsgBox "LOCATION NOT SELECTED !!!"
   Exit Sub
End If


Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTFARMER.Clear
'rs.Open "select * from tblfarmer where monitor in " & Dzstr & "order by idfarmer", MHVDB, adOpenStatic
rs.Open "select distinct substring(idfarmer,1,9) as dgt  from tblfarmer where monitor in " & Dzstr & "order by substring(idfarmer,1,9)", MHVDB, adOpenStatic
' continue from here after lunch
Do While rs.EOF <> True
'FindFA rs!idfarmer, "F"
FindDZ Mid(rs!dgt, 1, 3)
FindGE Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3)
FindTs Mid(rs!dgt, 1, 3), Mid(rs!dgt, 4, 3), Mid(rs!dgt, 7, 3)

LSTFARMER.AddItem Mid(rs!dgt, 1, 3) & " " & Trim(Dzname) + " |" + Mid(rs!dgt, 4, 3) & " " & Trim(GEname) + " |" + Mid(rs!dgt, 7, 3) & " " & Trim(TsName) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE

'LSTFARMER.AddItem Trim(rs!farmername) & " |" & rs!idfarmer


   'LSTFARMER.AddItem Trim(rs!farmername) & " |" & rs!idfarmer
   rs.MoveNext
Loop



Else
frmbarcode.Width = 9840
End If



End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
createhtml

End Sub


Private Sub createhtml()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim cnt As Integer
Dim i, j As Integer
Dim row, col As Integer
Dim html As String
Dim mm
Dim SQLSTR As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
Dzstr = ""
SQLSTR = ""
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                        
row = 1
col = 1






j = 0
If CHKGT.Value = 0 Then
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
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
ElseIf CHKGT.Value = 1 And CHKIF.Value = 0 Then
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    mm = Split(LSTPR.List(i), "|", -1, 1)
      ' DZstr = DZstr & "'" & Mid(MM(0), 1, 3) & Mid(MM(1), 1, 3) & Mid(MM(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       
       
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

Else



For i = 0 To LSTFARMER.ListCount - 1
    If LSTFARMER.Selected(i) Then
      mm = Split(LSTFARMER.List(i), "|", -1, 1)
       Dzstr = Dzstr & "'" & Mid(mm(0), 1, 3) & Mid(mm(1), 1, 3) & Mid(mm(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = LSTFARMER.List(i)




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




If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If



















If CHKGT.Value = 1 And CHKIF.Value = 1 Then
' if gt and if are checked


SQLSTR = "SELECT idfarmer,farmername FROM tblfarmer where length(monitor)='5' and substring(idfarmer,1,9) IN  " & Dzstr


ElseIf CHKGT.Value = 1 Then
' if only gt is check
SQLSTR = "SELECT idfarmer,farmername FROM tblfarmer where monitor IN  " & Dzstr


Else

' if no check
SQLSTR = "SELECT idfarmer,FARMERNAME FROM tblfarmer where monitor in (select staffcode from tblmhvstaff where msupervisor in " & Dzstr & ")"

End If




Set rs = Nothing
rs.Open SQLSTR, MHVDB



' Build KML Feature
Dim FileNum As Integer
    FileNum = FreeFile
HtmlFileName = App.Path & "\barcode.html"

    Open HtmlFileName For Output As #FileNum

Print #FileNum, "<html><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""><title>QR code sheet</title></head><body><table border=""0"" cellspacing=""20"">"

Do While rs.EOF <> True


If col = 1 Then
Print #FileNum, "<tr><td  align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!idfarmer & "&ecc=M&size=150x150""><small><br>" & rs!idfarmer & "<br><b>" & rs!farmername & "</b></small></td> <td></td><td></td>"
col = col + 1
ElseIf col = 2 Then
Print #FileNum, "<td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!idfarmer & "&ecc=M&size=150x150""><small><br>" & rs!idfarmer & "<br><b>" & rs!farmername & "</b></small></td><td></td><td></td>"
col = col + 1
ElseIf col = 3 Then
Print #FileNum, "<td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!idfarmer & "&ecc=M&size=150x150""><small><br>" & rs!idfarmer & "<br><b>" & rs!farmername & "</b></small></td><td></td><td></td>"
col = col + 1
Else
Print #FileNum, "<td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!idfarmer & "&ecc=M&size=150x150""><small><br>" & rs!idfarmer & "<br><b>" & rs!farmername & "</b></small></td><td></td><td></td></tr><tr></tr><tr></tr>"
col = 1
End If

rs.MoveNext
Loop

Print #FileNum, "</table></body></html>"
Close #FileNum
frmbarcodeweb.Show 1
End Sub


Private Sub DZLIST_ItemCheck(Item As Integer)
CHKGT.Value = 0
CHKIF.Value = 0
frmbarcode.Width = 3570
End Sub

Private Sub Form_Load()
frmbarcode.Width = 3570

Dim rs As New ADODB.Recordset
Set rs = Nothing




rs.Open "select distinct MSUPERVISOR as mm from tblmhvstaff where MSUPERVISOR<>'' Order by MSUPERVISOR", MHVDB, adOpenStatic

With rs
Do While Not .EOF
FindsTAFF !mm
   DZLIST.AddItem Trim(sTAFF) + " | " + !mm
   .MoveNext
Loop
End With



End Sub

Private Sub LSTFARMER_Click()

mycnt

End Sub

Private Sub mycnt()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim cnt As Integer
Dim i, j As Integer
Dim row, col As Integer
Dim html As String
Dim mm
Dim SQLSTR As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
Dzstr = ""
SQLSTR = ""
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                       






For i = 0 To LSTFARMER.ListCount - 1
    If LSTFARMER.Selected(i) Then
      mm = Split(LSTFARMER.List(i), "|", -1, 1)
       Dzstr = Dzstr & "'" & Mid(mm(0), 1, 3) & Mid(mm(1), 1, 3) & Mid(mm(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = LSTFARMER.List(i)




       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE LOCATION TO VIEW THIS REPORT."
          Exit Sub
       End If
       
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


SQLSTR = "SELECT count(*) as mn  FROM tblfarmer where status='A' and length(monitor)=5 and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr

Set rs = Nothing
rs.Open SQLSTR, MHVDB
cnt = IIf(IsNull(rs!mn), 0, rs!mn)
Label2.Caption = "Farmer selected : " & cnt
End Sub

Private Sub LSTPR_Click()
farmercnt
End Sub

Private Sub LSTPR_ItemCheck(Item As Integer)
CHKIF.Value = 0
frmbarcode.Width = 9840
End Sub
Private Sub farmercnt()
Dim mm

Dzstr = ""

For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    mm = Split(LSTPR.List(i), "|", -1, 1)
       'DZstr = DZstr & "'" & Mid(MM(0), 1, 3) & Mid(MM(1), 1, 3) & Mid(MM(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = LSTPR.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          'MsgBox "SELECT ATLEAST ONE LOCATION TO VIEW THIS REPORT."
          Exit Sub
       End If
       
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
lblcnt.Caption = ""
Exit Sub
End If


Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTFARMER.Clear
rs.Open "select count(*) cnt from tblfarmer where status='A' and monitor in " & Dzstr & "order by idfarmer", MHVDB, adOpenStatic
If rs.EOF <> True Then
lblcnt.Caption = "Total Farmer selected " & rs!cnt
End If


End Sub
