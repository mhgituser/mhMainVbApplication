VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdatadownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DOWNLOAD NGT DATA"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10560
   Icon            =   "frmdatadownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      Picture         =   "frmdatadownload.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      Begin VB.TextBox txtgap2mys 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtgap4dateamys 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtgap2acc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtgap4dateacc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkgap2 
         Caption         =   "MS-GAP-02 NGT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox chkgap1 
         Caption         =   "MS-GAP-01 NGT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkgap4 
         Caption         =   "MS-GAP-04 Nursery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker txtfromdategap4 
         Height          =   375
         Left            =   6840
         TabIndex        =   15
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78446593
         CurrentDate     =   41519
      End
      Begin MSComCtl2.DTPicker txttodategap4 
         Height          =   375
         Left            =   8760
         TabIndex        =   16
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78446593
         CurrentDate     =   41519
      End
      Begin MSComCtl2.DTPicker txtfromdategap1 
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   78446593
         CurrentDate     =   41519
      End
      Begin MSComCtl2.DTPicker txttodategap1 
         Height          =   375
         Left            =   8760
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   78446593
         CurrentDate     =   41519
      End
      Begin MSComCtl2.DTPicker txtfromdategap2 
         Height          =   375
         Left            =   6840
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78446593
         CurrentDate     =   41519
      End
      Begin MSComCtl2.DTPicker txttodategap2 
         Height          =   375
         Left            =   8760
         TabIndex        =   20
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   78446593
         CurrentDate     =   41519
      End
      Begin VB.Label Label5 
         Caption         =   "Last Date in Mysql"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Last Date in Access"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Data Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3480
      Picture         =   "frmdatadownload.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "frmdatadownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ngtconn As New ADODB.Connection
Dim myerr As Boolean

Private Sub Command1_Click()
If chkgap4.Value = 1 Then
gap4
ElseIf chkgap1.Value = 1 Then


ElseIf chkgap2.Value = 1 Then
gap2
Else

MsgBox "No Table Selected for updation."
myerr = True
End If


If myerr = True Then
Command1.Enabled = True
Else
Command1.Enabled = False
MsgBox "successfully updated."
ngtconn.Close
Kill "C:\NGT.accdb"
End If



End Sub
Private Sub gap4()
On Error GoTo err
myerr = False
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Dim rsmysql As New ADODB.Recordset
Dim rsngt As New ADODB.Recordset
Dim ngtplantbatch As Integer
    Dim NewtrnId, credit, debit As Long
    Dim entrydate As Date
    Dim plantBatch, varietyId, planttype, verificationType, transactiontype As Integer
    Dim staffid, status, location, facilityid As String
    
MHVDB.BeginTrans

MHVDB.Execute "delete from tblqmsplanttransaction where entrydate>='" & Format(txtfromdategap4.Value, "yyyy-MM-dd") & "' and entrydate<='" & Format(txttodategap4.Value, "yyyy-MM-dd") & "' and facilityid in(select facilityid from tblqmsfacility where housetype='T')"

Set rs = Nothing
SQLSTR = "select * from [MS-GAP-04 Nursery] where date>=datevalue('" & txtfromdategap4.Value & "') and date<=datevalue('" & txttodategap4.Value & "')and FID in(9,10,11,12,34,35,36,37,38)"

rs.Open SQLSTR, ngtconn, dbOpenForwardOnly

Do While rs.EOF <> True

                entrydate = Format(rs!Date, "yyyy-MM-dd")
                credit = IIf(IsNull(rs!credit), 0, rs!credit)
                debit = IIf(IsNull(rs!debit), 0, rs!debit)
                transactiontype = Trim(rs![Transitional Type])
                verificationType = Trim(rs![Verification Type])
                
                Set rsngt = Nothing
                rsngt.Open "select * from [MS-GAP-11] where ID=" & rs!PBID & "", ngtconn
                If rs.EOF <> True Then
                plantBatch = rsngt![Plant Batch ID]
                Else
                MsgBox "The Plant Batch no " & rs!PBID & " does not match,hence you cannot continue to upload."
                MHVDB.RollbackTrans
                myerr = True
                Exit Sub
                End If
                
                
                Set rsngt = Nothing
                rsngt.Open "select * from [MS-GAP-13] where ID=" & rs![Staff ID] & "", ngtconn
                If rs.EOF <> True Then
                staffid = rsngt![Staff ID]
                Else
                MsgBox " No such Staff ID " & rs![Staff ID] & " found,hence you cannot continue to upload."
                MHVDB.RollbackTrans
                myerr = True
                End If
                
                Set rsngt = Nothing
                rsngt.Open "select * from [MS-GAP-14] where ID=" & rs![fid] & "", ngtconn
                If rs.EOF <> True Then
                facilityid = rsngt![Facility ID]
                Else
                MsgBox " No such facility  ID " & rs![fid] & " found,hence you cannot continue to upload."
                MHVDB.RollbackTrans
                myerr = True
                End If
                
                
                
                Set rsmysql = Nothing
                rsmysql.Open "select * from tblqmsplantbatchdetail where plantbatch='" & plantBatch & "'", MHVDB
                If rsmysql.EOF <> True Then
                varietyId = rsmysql!plantvariety
                planttype = rsmysql!planttype
                
                Else
                MsgBox "The Plant Batch no " & plantBatch & " does not match,hence you cannot continue to upload."
                MHVDB.RollbackTrans
                myerr = True
                Exit Sub
                End If
                
                
                
                Set rsmysql = Nothing
                rsmysql.Open "select max(trnid) as max from tblqmsplanttransaction", MHVDB
                If rsmysql.EOF <> True Then
                NewtrnId = rsmysql!max + 1
                End If
                
                SQLSTR = "insert into tblqmsplanttransaction(trnid,entrydate, " _
                & "facilityid,plantbatch,varietyid,verificationtype,transactiontype,credit,debit,staffid,status,location) " _
                & "values(" _
                & "'" & NewtrnId & "'," _
                & "'" & Format(entrydate, "yyyy-MM-dd") & "'," _
                & "'" & facilityid & "'," _
                & "'" & plantBatch & "'," _
                & "'" & varietyId & "'," _
                & "'" & verificationType & "'," _
                & "'" & transactiontype & "'," _
                & "'" & credit & "'," _
                & "'" & debit & "'," _
                & "'" & staffid & "'," _
                & "'ON'," _
                & "'NGT'" _
                & ")"
                
                MHVDB.Execute SQLSTR
                
rs.MoveNext
Loop

MHVDB.CommitTrans
Exit Sub
err:
MsgBox err.Description & " on date " & Format(entrydate, "yyyy-MM-dd")
    MHVDB.RollbackTrans
myerr = True
End Sub
Private Sub gap2()
On Error GoTo err
myerr = False
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Dim rsmysql As New ADODB.Recordset
Dim rsngt As New ADODB.Recordset
Dim chemicalid, applicationmethod As Integer
    Dim NewtrnId As Long
    Dim entrydate, endtime, starttime As Date
    Dim facilityid, reason, staffid, location As String
    Dim area, chemicalqty, totvolume As Double
    
    
    
  
MHVDB.BeginTrans

MHVDB.Execute "delete from tblqmsspray where entrydate>='" & Format(txtfromdategap2.Value, "yyyy-MM-dd") & "' and entrydate<='" & Format(txttodategap2.Value, "yyyy-MM-dd") & "' and location='NGT'"

Set rs = Nothing
SQLSTR = "select * from [MS-GAP-02_NGT] where date1>=datevalue('" & txtfromdategap2.Value & "') and date1<=datevalue('" & txttodategap2.Value & "')and [Facility ID] in(9,10,11,12,34,35,36,37,38)"

rs.Open SQLSTR, ngtconn, dbOpenForwardOnly

Do While rs.EOF <> True

                entrydate = Format(rs!Date1, "yyyy-MM-dd")
              
                
                endtime = Format(rs![Time Finish], "hh:mm")
                starttime = Format(rs![Time Start], "hh:mm")
                area = rs![m2 Sprayed]
                chemicalid = rs![Spray Chemical Number]
                chemicalqty = rs![kg or litres of Chemical]
                applicationmethod = rs![Method of Application]
                totvolume = rs![Total Volume Applied kg or litres]
                reason = rs![Reason for applying]
                
                
                Set rsngt = Nothing
                rsngt.Open "select * from [MS-GAP-13] where ID=" & rs![Staff ID] & "", ngtconn
                If rs.EOF <> True Then
                staffid = rsngt![Staff ID]
                Else
                MsgBox " No such Staff ID " & rs![Staff ID] & " found,hence you cannot continue to upload."
                MHVDB.RollbackTrans
                myerr = True
                End If
                
                Set rsngt = Nothing
                rsngt.Open "select * from [MS-GAP-14] where ID=" & rs![Facility ID] & "", ngtconn
                If rs.EOF <> True Then
                facilityid = rsngt![Facility ID]
                Else
                MsgBox " No such facility  ID " & rs![fid] & " found,hence you cannot continue to upload."
                MHVDB.RollbackTrans
                myerr = True
                End If
                
                
                
                          
                
                
                Set rsmysql = Nothing
                rsmysql.Open "select max(trnid) as max from tblqmsspray", MHVDB
                If rsmysql.EOF <> True Then
                NewtrnId = rsmysql!max + 1
                End If
                
                SQLSTR = "insert into tblqmsspray(trnid,entrydate, " _
                & "facilityid,endtime,starttime,area,chemicalQty,chemicalId,applicationMethod,totalVol,reason,staffid,status,location) " _
                & "values(" _
                & "'" & NewtrnId & "'," _
                & "'" & Format(entrydate, "yyyy-MM-dd") & "'," _
                & "'" & facilityid & "'," _
                & "'" & endtime & "'," _
                & "'" & starttime & "'," _
                & "'" & area & "'," _
                & "'" & chemicalqty & "'," _
                & "'" & chemicalid & "'," _
                & "'" & applicationmethod & "'," _
                & "'" & totvolume & "'," _
                & "'" & reason & "'," _
                & "'" & staffid & "'," _
                & "'ON'," _
                & "'NGT'" _
                & ")"
                
                MHVDB.Execute SQLSTR
                
rs.MoveNext
Loop

MHVDB.CommitTrans
Exit Sub
err:
MsgBox err.Description
    MHVDB.RollbackTrans
myerr = True
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
myerr = False
Dim ngtfile As String
Dim rs As New ADODB.Recordset
Set ngtconn = New ADODB.Connection
ngtconn.CursorLocation = adUseClient
mdbfile = "C:\NGT.accdb"
ngtconn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\NGT.accdb;" _
            & "Persist Security Info=False;"
            
            
            Set rs = Nothing
rs.Open "select max(date) as maxdate  from [MS-GAP-04 Nursery]", ngtconn
If rs.EOF <> True Then
txtgap4dateacc.Text = Format(rs!maxdate, "dd/MM/yyyy")
txtfromdategap4.Value = Format(rs!maxdate + 1, "dd/MM/yyyy")
txttodategap4.Value = Format(rs!maxdate + 1, "dd/MM/yyyy")
End If


Set rs = Nothing
rs.Open "select max(entrydate) as maxdate  from tblqmsplanttransaction where facilityid in (select facilityid from tblqmsfacility where housetype='T') ", MHVDB
If rs.EOF <> True Then
txtgap4dateamys.Text = Format(rs!maxdate, "dd/MM/yyyy")

End If

'MS-GAP-02_NGT
Set rs = Nothing
rs.Open "select max(date1) as maxdate  from [MS-GAP-02_NGT]", ngtconn
If rs.EOF <> True Then
txtgap2acc.Text = Format(rs!maxdate, "dd/MM/yyyy")
txtfromdategap2.Value = Format(rs!maxdate + 1, "dd/MM/yyyy")
txttodategap2.Value = Format(rs!maxdate + 1, "dd/MM/yyyy")
End If
Set rs = Nothing
rs.Open "select max(entrydate) as maxdate  from tblqmsspray where location='NGT'", MHVDB
If rs.EOF <> True Then
txtgap2mys.Text = Format(rs!maxdate, "dd/MM/yyyy")
Else
txtgap2mys.Text = ""
End If







Command1.Enabled = True
Exit Sub
err:
    MsgBox "Access File Error, copy and past the NGT updated file in location C:\  and rename it to NGT."
    Command1.Enabled = False

End Sub
