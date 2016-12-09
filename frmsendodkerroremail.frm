VERSION 5.00
Begin VB.Form frmsendodkerroremail 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   2175
   ClientTop       =   1785
   ClientWidth     =   16785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   16785
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   10320
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   16335
   End
   Begin VB.CommandButton btnSendMail 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
End
Attribute VB_Name = "frmsendodkerroremail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
   Dim X As Integer
   Dim PadLength As Integer
Private Sub sendErrorMail()
emailMessageString = ""
Dim recordno As Integer
Dim header As String
Dim param As Integer
Dim rs As New ADODB.Recordset
Dim mCONNECTION As String
Dim retVal          As String
mCONNECTION = "smtp.tashicell.com"

Text1.Text = ""

Set rs = Nothing
rs.Open "select * from tblodkfollowuplog where emailstatus='ON' order by paraid", ODKDB
'header = Lpad("", "-", 196) & vbCrLf
'
'header = header & Rpad("S/N.", " ", 8) + Rpad("Date", " ", 20) + Rpad("Threshold Value", " ", 22) + Rpad("ODK Value", " ", 15) + Rpad("Report Name", " ", 25) + Rpad("Monitor", " ", 10) + Rpad("Farmer", " ", 20) + Lpad("Field Code", " ", 20) & vbCrLf
'header = header & Lpad("", "-", 196) & vbCrLf
'' header = "S/N:" & "Date:" & "Parameter of Concern:" & "Acceptable Threshold Value:" _
''            & "ODK Value:" & "Report Name:" & "Monitor Name:" & "Farmer Name:" & "Dzongkhag Name:" _
''            & "Gewog Name:" & "Tshowg Name:" & "Field Code:" & vbCrLf

    Do Until rs.EOF
       
       param = rs!paraid
       findParamDetails rs!paraid
     
      recordno = 0
       Do While param = rs!paraid
      FindsTAFF rs!staffcode
      FindFA rs!farmercode, "F"
       FindDZ Mid(rs!farmercode, 1, 3)
       FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
       FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
            recordno = recordno + 1
           
'            emailMessageString = emailMessageString & CStr(recordno) + " " + PadWithChar(IIf(IsNull(sTAFF), "", sTAFF), 26, " ", 2) + PadWithChar(Format(rs!odkValue, "########0.00"), 12, " ", 2) + PadWithChar(Format(acceptableThresholdValue, "########0.00"), 12, " ", 1) + PadWithChar(FAName, 12, " ", 2) & vbNewLine
            
            emailMessageString = emailMessageString & "S/N:" & PadWithChar(" ", 45, " ", 1) & recordno & vbCrLf
            emailMessageString = emailMessageString & "Date:" & PadWithChar(" ", 44, " ", 1) & Format(rs!odkStartDate, "dd/MM/yyyy") & vbCrLf
            emailMessageString = emailMessageString & "Parameter of Concern:" & PadWithChar(" ", 17, " ", 1) & paramName & vbCrLf
            emailMessageString = emailMessageString & "Acceptable Threshold Value:" & PadWithChar(" ", 7, " ", 1) & acceptableThresholdValue & vbCrLf
            emailMessageString = emailMessageString & "ODK Value:" & PadWithChar(" ", 34, " ", 1) & rs!odkValue & vbCrLf
            emailMessageString = emailMessageString & "Report Name:" & PadWithChar(" ", 31, " ", 1) & ReportName & vbCrLf
            emailMessageString = emailMessageString & "Monitor Name:" & PadWithChar(" ", 30, " ", 1) & rs!staffcode & " " & sTAFF & vbCrLf
            emailMessageString = emailMessageString & "Farmer Name:" & PadWithChar(" ", 31, " ", 1) & rs!farmercode & " " & FAName & vbCrLf
            emailMessageString = emailMessageString & "Dzongkhag Name:" & PadWithChar(" ", 23, " ", 1) & Mid(rs!farmercode, 1, 3) & " " & Dzname & vbCrLf
            emailMessageString = emailMessageString & "Gewog Name:" & PadWithChar(" ", 30, " ", 1) & Mid(rs!farmercode, 4, 3) & " " & GEname & vbCrLf
            emailMessageString = emailMessageString & "Tshowg Name:" & PadWithChar(" ", 29, " ", 1) & Mid(rs!farmercode, 7, 3) & " " & TsName & vbCrLf
            emailMessageString = emailMessageString & "Field Code:" & PadWithChar(" ", 35, " ", 1) & rs!fieldcode & vbCrLf
            emailMessageString = emailMessageString & vbCrLf & vbCrLf
            
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       'emailMessageString = header & emailMessageString & vbCrLf
      
      retVal = SendMail(emailId, "ODK Error On " & paramName, "ODKerror@mhv.com", _
          emailMessageString, mCONNECTION, 25, _
          "habizabi", "habizabi", _
           "", CBool(False))
'
        If retVal = "ok" Then
'        ODKDB.Execute "update tblodkfollowuplog set emailstatus='C' where emailstatus='ON' and paraid='" & param & "' "
        Else
        'ODKDB.Execute "update tblodkfollowuplog set emailstatus='C' where emailstatus='ON' and paraid='" & param & "' "
        End If
      Text1.Text = emailMessageString
     ' emailMessageString = ""
    Loop

End Sub

Private Sub btnSendMail_Click()
sendErrorMail
End Sub

 Function Lpad(MyValue As String, MyPadCharacter As String, MyPaddedLength As Integer)

      PadLength = MyPaddedLength - Len(MyValue)
      Dim PadString As String
      For X = 1 To PadLength
         PadString = PadString & MyPadCharacter
      Next
      Lpad = PadString + MyValue

   End Function
 Function Rpad(MyValue As String, MyPadCharacter As String, MyPaddedLength As Integer)

      PadLength = MyPaddedLength - Len(MyValue)
      Dim PadString As String
      For X = 1 To PadLength
         PadString = MyPadCharacter & PadString
      Next
      Rpad = MyValue + PadString

   End Function

