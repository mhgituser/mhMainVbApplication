VERSION 5.00
Begin VB.Form frmbarcodemhvstaff 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Advocate"
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
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Monitor"
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
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1215
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
      ItemData        =   "frmbarcodemhvstaff.frx":0000
      Left            =   3720
      List            =   "frmbarcodemhvstaff.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   120
      Width           =   6015
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
      Picture         =   "frmbarcodemhvstaff.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
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
      Left            =   1680
      Picture         =   "frmbarcodemhvstaff.frx":076E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   5400
   End
End
Attribute VB_Name = "frmbarcodemhvstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DZstr As String

Private Sub createhtml()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim cnt As Integer
Dim i, j As Integer
Dim row, col As Integer
Dim html As String
Dim MM
Dim SQLSTR As String
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
DZstr = ""
SQLSTR = ""
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                        
row = 1
col = 1






j = 0

For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    MM = Split(LSTPR.List(i), "|", -1, 1)
      ' DZstr = DZstr & "'" & Mid(MM(0), 1, 3) & Mid(MM(1), 1, 3) & Mid(MM(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       DZstr = DZstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
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




If Len(DZstr) > 0 Then
   DZstr = "(" + Left(DZstr, Len(DZstr) - 1) + ")"
 
Else
   MsgBox "NOT SELECTED !!!"
   Exit Sub
End If


















If Option1.Value = 1 Then
' if gt and if are checked
SQLSTR = "SELECT staffcode,staffname FROM tblmhvstaff where staffcode IN  " & DZstr


Else
' if only gt is check
SQLSTR = "SELECT staffcode,staffname FROM tblmhvstaff where staffcode IN  " & DZstr

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
Print #FileNum, "<tr><td  align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!staffcode & "&ecc=M&size=150x150""><small><br>" & rs!staffcode & "<br><b>" & rs!staffname & "</b></small></td> <td></td><td></td><td></td><td></td><td></td><td></td><td></td>"
col = col + 1
ElseIf col = 2 Then
Print #FileNum, "<td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!staffcode & "&ecc=M&size=150x150""><small><br>" & rs!staffcode & "<br><b>" & rs!staffname & "</b></small></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>"
col = col + 1
'ElseIf col = 3 Then
'Print #FileNum, "<td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!idfarmer & "&ecc=M&size=150x150""><small><br>" & rs!idfarmer & "<br><b>" & rs!farmername & "</b></small></td><td></td><td></td>"
'col = col + 1
Else
Print #FileNum, "<td align=""Center""><img src=""http://api.qrserver.com/v1/create-qr-code/?data=" & rs!staffcode & "&ecc=M&size=150x150""><small><br>" & rs!staffcode & "<br><b>" & rs!staffname & "</b></small></td><td></td><td></td></tr><tr></tr><tr></tr><tr></tr><tr></tr><tr></tr><tr></tr>"
col = 1
End If


rs.MoveNext
Loop

Print #FileNum, "</table></body></html>"
Close #FileNum
frmbarcodeweb.Show 1
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
createhtml
End Sub

Private Sub Option1_Click()

Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTPR.Clear
rs.Open "select * from tblmhvstaff where moniter='1' order by staffcode", MHVDB, adOpenStatic
With rs
Do While Not .EOF

   LSTPR.AddItem rs!staffname + " |" + rs!staffcode '+ " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With

End Sub

Private Sub Option2_Click()
Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTPR.Clear
rs.Open "select * from tblmhvstaff where advocate='1' order by staffcode", MHVDB, adOpenStatic
With rs
Do While Not .EOF

   LSTPR.AddItem rs!staffname + " |" + rs!staffcode '+ " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(rs!tshewogname) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With
End Sub
