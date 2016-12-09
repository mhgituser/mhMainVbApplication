VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RPTLANDDETAILS 
   Caption         =   "LAND DETAILS"
   ClientHeight    =   10950
   ClientLeft      =   255
   ClientTop       =   315
   ClientWidth     =   20250
   Icon            =   "RPTLANDDETAILS.dsx":0000
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "RPTLANDDETAILS.dsx":0E42
End
Attribute VB_Name = "RPTLANDDETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dzname, GEname, TsName As String
Dim TSTOT, GETOT, DZTOT, GRANDTOT As Double

Private Sub ActiveReport_ReportStart()
TSTOT = 0
GETOT = 0
DZTOT = 0
GRANDTOT = 0
End Sub

Private Sub Detail_Format()
txtREGLAND.Text = Format(txtREGLAND.Text, "#####0.00")
TSTOT = TSTOT + Format(Val(txtREGLAND.Text), "#####0.00")
GETOT = GETOT + Format(Val(txtREGLAND.Text), "#####0.00")
DZTOT = DZTOT + Format(Val(txtREGLAND.Text), "#####0.00")
GRANDTOT = GRANDTOT + Format(Val(txtREGLAND.Text), "#####0.00")
End Sub

Private Sub GroupFooter1_Format()
TXTDZTOT = Format(DZTOT, "#####0.00")
 DZTOT = 0
End Sub

Private Sub GroupFooter2_Format()
TXTGETOT = Format(GETOT, "#####0.00")
GETOT = 0
End Sub

Private Sub GroupFooter3_Format()
TXTTSTOT.Text = Format(TSTOT, "#####0.00")
TSTOT = 0
End Sub

Private Sub GroupHeader1_Format()
FindDZ Trim(txtDZCODE)
TXTDZ.Text = UCase(Dzname)
End Sub

Private Sub GroupHeader2_Format()
FindGE Trim(txtDZCODE), Trim(txtGECODE)
TXTGE.Text = UCase(GEname)
End Sub

Private Sub GroupHeader3_Format()
FindTs Trim(txtDZCODE), Trim(txtGECODE), Trim(txtTSCODE)
TXTTS.Text = UCase(TsName)
End Sub

Private Sub FindDZ(dd As String)
'On Error GoTo err
Dim RS As New ADODB.Recordset
Dzname = ""
Set RS = Nothing
RS.Open "select * from tbldzongkhag where dzongkhagcode='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If RS.EOF <> True Then
Dzname = RS!DZONGKHAGNAME
Else
MsgBox "Record Not Found."
End If
Exit Sub
'err:
'MsgBox err.Description
End Sub
Private Sub FindGE(dd As String, GG As String)
On Error GoTo err

Dim ERRSTR As String
ERRSTR = ""
ERRSTR = dd + "   " + GG
Dim RS As New ADODB.Recordset
GEname = ""
Set RS = Nothing
RS.Open "select * from TBLGEWOG where dzongkhagID='" & dd & "' AND GEWOGID='" & GG & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If RS.EOF <> True Then
GEname = RS!gewogname
Else
MsgBox ERRSTR & "   " & "Record Not Found."

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub FindTs(dd As String, GG As String, tt As String)
On Error GoTo err
Dim ERRSTR As String
ERRSTR = ""
ERRSTR = dd + "   " + GG + "   " + tt
Dim RS As New ADODB.Recordset
TsName = ""
Set RS = Nothing
RS.Open "select * from tbltshewog where dzongkhagID='" & dd & "' AND GEWOGID='" & GG & "' and tshewogid='" & tt & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If RS.EOF <> True Then
TsName = RS!tshewogname

Else
MsgBox ERRSTR & "   " & "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub



Private Sub ReportFooter_Format()
TXTGRANDTOT.Text = Format(GRANDTOT, "#####0.00")

End Sub
