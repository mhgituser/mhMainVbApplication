VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FRMODKDATACOLLECT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODK DATA COLLECT"
   ClientHeight    =   1845
   ClientLeft      =   6720
   ClientTop       =   4095
   ClientWidth     =   7800
   Icon            =   "FRMODKDATACOLLECT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   7800
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1350
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   3840
      Picture         =   "FRMODKDATACOLLECT.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COLLECT"
      Height          =   735
      Left            =   2640
      Picture         =   "FRMODKDATACOLLECT.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "FRMODKDATACOLLECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo ERR
Dim emailmessage As String
Dim EMAILIDS As String
Dim m, n As Integer
Dim Msg As String
Dim rstblcount As New ADODB.Recordset
Dim MYSQLSTR As String
Dim PRGBAR As Integer
Dim MTEMPVAR As String
Dim rsd As New ADODB.Recordset
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim RsRemote As New ADODB.Recordset
Dim RsLocal As New ADODB.Recordset
ProgressBar1.Visible = True

Dim connRemote As New ADODB.Connection
Dim CONNLOCAL As New ADODB.Connection

 MYSQLSTR = ""
 
Set rsd = Nothing

 
 Screen.MousePointer = vbHide
 DoEvents
 
  If connRemote.State = adStateOpen Then
    connRemote.Close
    End If
    
    
 connRemote.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                        & "SERVER=206.217.141.114;" _
                        & " DATABASE=odk_prod;" _
                        & "UID=odk_user;PWD=none; OPTION=3"
' connRemote.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
'                        & "SERVER=localhost;" _
'                        & " DATABASE=odk_prod;" _
'                        & "UID=admin;PWD=password; OPTION=3"
                        
                        
                        
 connRemote.Open

CONNLOCAL.Open OdkCnnString
                      
Set rsd = Nothing
rsd.Open "select * from tblmuksys", CONNLOCAL
'odk_prodLocal
Set rstblcount = Nothing
rstblcount.Open "SELECT COUNT(*) AS TBLCOUNT FROM tbltable WHERE STATUS='ON'", CONNLOCAL
PRGBAR = 100 / IIf(IsNull(rstblcount!TBLCOUNT), 1, rstblcount!TBLCOUNT)
m = IIf(IsNull(rstblcount!TBLCOUNT), 1, rstblcount!TBLCOUNT)
n = 0

Set rs = Nothing
rs.Open "SELECT * FROM tbltable WHERE STATUS='ON' order by tblid", CONNLOCAL



Do While rs.EOF <> True


        Set RsRemote = Nothing
        RsRemote.Open "SELECT * FROM " & rs!tblname & "  ", connRemote
        'where substring(_creation_date,1,10)>='" & Format(rsd!mydate, "yyyy-MM-dd") & "'
        If RsRemote.EOF <> True Then
        
                Do While RsRemote.EOF <> True
                        Set RsLocal = Nothing
                        RsLocal.Open "SELECT * FROM " & LCase(rs!tblname) & " WHERE  " & rs!Key & "='" & RsRemote.Fields(0) & "' ", CONNLOCAL  'IN(SELECT " & RS!Key & " FROM  " & RS!tblname & " )  ", CONNLOCAL
                        
                        If RsLocal.EOF <> True Then
                        
                        Else
                                For i = 0 To rs!fieldcount - 1
                                        If RsRemote.Fields(i).Type = 200 Then
                                                MTEMPVAR = ValidateString(IIf(IsNull(RsRemote.Fields(i)), "", RsRemote.Fields(i)))
                                        ElseIf RsRemote.Fields(i).Type = 135 Then
                                                MTEMPVAR = Format(RsRemote.Fields(i), "yyyy-MM-dd hh:mm:ss")
                                        Else
                                                MTEMPVAR = IIf(IsNull(RsRemote.Fields(i)), "", RsRemote.Fields(i))
                                        End If
                                        
                                        MYSQLSTR = MYSQLSTR + "'" + Trim(Mid(MTEMPVAR, InStr(1, MTEMPVAR, "|") + 1)) + "',"
                                Next
                                MYSQLSTR = MYSQLSTR + "'" + " " + "'" & ","
                                MYSQLSTR = "(" + Mid(MYSQLSTR, 1, Len(MYSQLSTR) - 1) + ")"
                                CONNLOCAL.Execute "INSERT INTO " & LCase(rs!tblname) & "  VALUES " + MYSQLSTR
                                MYSQLSTR = ""
                        End If
                
                RsRemote.MoveNext
                Loop
        Else
        
                'MsgBox "no records at source"
        
        End If
        
        
        n = n + 1
        Msg = "Transferring Table : " + rs!tblname & ".    " & "Transferring " & " " & (n) & "/" & m & "  Tables."
        sb.SimpleText = Msg
        
        If ProgressBar1.Value + PRGBAR >= 100 Then
                ProgressBar1.Value = 100
        Else
                ProgressBar1.Value = ProgressBar1.Value + PRGBAR
        End If

rs.MoveNext
Loop
CONNLOCAL.Execute "update tblmuksys set mydate='" & Format(rsd!mydate + 1, "yyyy-MM-dd") & "'"
ProgressBar1.Visible = False
 Screen.MousePointer = vbDefault
 Command1.Enabled = False




'ERR:

ProgressBar1.Visible = False
 Screen.MousePointer = vbDefault
 Command1.Enabled = False
EMAILIDS = "muktitcc@gmail.com,swangchuk@mountainhazelnuts.com"

       ' EMAILIDS = "muktitcc@gmail.com,muktitcc@gmail.com,sikkagurung@gmail.com"

    If Len(err.Description) = 0 Then
     emailmessage = "ODK DATA DOWNLOAD SUCCESSFULLY DONE ON " & Format(Now, "dd/MM/yyyy hh:mm:ss")
     LogRemarks = "ODK DATA DOWNLOAD SUCCESSFULLY DONE ON " & Format(Now, "dd/MM/yyyy hh:mm:ss")
     updateodklog "no uri", Now, MUSER, LogRemarks, ""
    Else
    emailmessage = " SONAM, THERE IS AN ERROR ON DOWNLOADING ODK DATA,PLEASE CHECK. REPORTED ON " & Format(Now, "dd/MM/yyyy hh:mm:ss ERROR:  ") & UCase(err.Description)
     LogRemarks = "ERROR IN ODK DATA DOWNLOAD  " & Format(Now, "dd/MM/yyyy hh:mm:ss") & " REPORTED ON " & Format(Now, "dd/MM/yyyy hh:mm:ss ERROR:  ") & UCase(err.Description)
     updateodklog "no uri", Now, MUSER, LogRemarks, ""
    End If
    LogRemarks = ""




Dim mCONNECTION As String
Dim retVal          As String
mCONNECTION = "smtp.tashicell.com"
  updateField
  updateStorage
  updateDailyact
  updatesiatribution
 retVal = SendMail(EMAILIDS, "STATUS ON ODK DATA TRANSFER.", "NAS@MHV.COM", _
    emailmessage, mCONNECTION, 25, _
    "habizabi", "habizabi", "", CBool(False))
  
If retVal = "ok" Then
MsgBox "done"
Else
MsgBox "Please Check Internet Connection " & retVal
End If




End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim emailmessage As String
Dim EMAILIDS As String

EMAILIDS = "muktitcc@gmail.com"

       ' EMAILIDS = "muktitcc@gmail.com,muktitcc@gmail.com,sikkagurung@gmail.com"

    
     emailmessage = "ODK DATA UPLOAD SUCCESSFULLY DONE ON " & Format(Now, "dd/MM/yyyy hh:mm:ss")
  
   ' emailmessage = " SONAM, THERE IS AN ERROR ON UPLOADING ODK DATA,PLEASE CHECK. REPORTED ON " & Format(Now, "dd/MM/yyyy hh:mm:ss ERROR:  ") & UCase(ERR.Description)
  





Dim mCONNECTION As String
Dim retVal          As String
mCONNECTION = "smtp.tashicell.com"
  
 retVal = SendMail(EMAILIDS, "STATUS ON ODK DATA TRANSFER.", "NAS@MHV.COM", _
    emailmessage, mCONNECTION, 25, _
    "habizabi", "habizabi", _
    "", CBool(False))
  
If retVal = "ok" Then
Else
MsgBox "Please Check Internet Connection " & retVal
End If


End Sub

Private Sub Form_Load()
 Command1.Enabled = True
End Sub
