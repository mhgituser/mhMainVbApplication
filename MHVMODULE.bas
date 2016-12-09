Attribute VB_Name = "mhvmodule"
Public sTAFF As String
Public OdkCnnString As String
Public CnnString As String

Public Sub FindsTAFF(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
sTAFF = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
sTAFF = rs!staffname
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "INVALID STAFF CODE=" & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = err.Description
End Function

