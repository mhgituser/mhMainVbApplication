VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmgoogleview 
   Caption         =   "GOOGLE EARTH"
   ClientHeight    =   3570
   ClientLeft      =   3045
   ClientTop       =   1815
   ClientWidth     =   6075
   Icon            =   "frmgoogleview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6075
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
      Left            =   3000
      Picture         =   "frmgoogleview.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   5295
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81002497
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81002497
         CurrentDate     =   41362
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.OptionButton OPTSEL 
      Caption         =   "SELECTIVE"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton OPTALL 
      Caption         =   "ALL"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox CHKALLFIELD 
      Caption         =   "ALL"
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW"
      Height          =   735
      Left            =   1440
      Picture         =   "frmgoogleview.frx":1434
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo CBOMONITOR 
      Bindings        =   "frmgoogleview.frx":2076
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo CBOFARMER 
      Bindings        =   "frmgoogleview.frx":208B
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1200
      TabIndex        =   12
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo CBOFDCODE 
      Bindings        =   "frmgoogleview.frx":20A0
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1200
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "FARMER"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MONITOR"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   765
   End
End
Attribute VB_Name = "frmgoogleview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsmonitor As New ADODB.Recordset
Dim retVal As Boolean

Private Sub CBOFARMER_LostFocus()
Dim rs As New ADODB.Recordset


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                      
Set rsmonitor = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct fdcode   from phealthhub15_core where farmerbarcode='" & CBOFARMER.BoundText & "' and start between '" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and '" & Format(txttodate.Value, "yyyy-MM-dd") & "'", db
Set CBOFDCODE.RowSource = rs
CBOFDCODE.ListField = "fdcode"
CBOFDCODE.BoundColumn = "fdcode"
End Sub

Private Sub CBOMONITOR_LostFocus()
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""

db.Open OdkCnnString
                        
Set rsmonitor = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct farmerbarcode   from phealthhub15_core where staffbarcode='" & CBOMONITOR.BoundText & "' and start between '" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and '" & Format(txttodate.Value, "yyyy-MM-dd") & "'  AND GPS_COORDINATES_LAT<>0 AND GPS_COORDINATES_LNG<>0", db
Set CBOFARMER.RowSource = rs
CBOFARMER.ListField = "farmerbarcode"
CBOFARMER.BoundColumn = "farmerbarcode"
End Sub

Private Sub CHKALLFIELD_Click()
If CHKALLFIELD.Value = 1 Then
CBOFDCODE.Enabled = False
Else

CBOFDCODE.Enabled = True
End If
End Sub

Private Sub Command1_Click()



createkml



If retVal = True Then
Navigate App.Path & "\tempfile.kml"
Else
MsgBox "There Is No Visit Information."
End If
End Sub
Private Sub createkml()
Dim rs As New ADODB.Recordset
Dim rschk As New ADODB.Recordset

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                        
If OPTALL.Value = True Then
Exit Sub

Else
If CHKALLFIELD.Value = 0 Then
If Len(CBOFARMER.Text) <> 0 Then
rs.Open "select * from phealthhub15_core where substring(start,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and staffbarcode='" & CBOMONITOR.BoundText & "' AND GPS_COORDINATES_LAT<>0 AND GPS_COORDINATES_LNG<>0 and farmerbarcode='" & CBOFARMER.BoundText & "'", db
Else
rs.Open "select * from phealthhub15_core where substring(start,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and staffbarcode='" & CBOMONITOR.BoundText & "' AND GPS_COORDINATES_LAT<>0 AND GPS_COORDINATES_LNG<>0 ", db
End If
Else
rs.Open "select * from phealthhub15_core where substring(start,1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and substring(start,1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and staffbarcode='" & CBOMONITOR.BoundText & "' AND GPS_COORDINATES_LAT<>0 AND GPS_COORDINATES_LNG<>0 and farmerbarcode='" & CBOFARMER.BoundText & "' ", db
End If
End If





' Build KML Feature
Dim FileNum As Integer
    FileNum = FreeFile
KmlFileName = App.Path & "\tempfile.kml"

    Open KmlFileName For Output As #FileNum

Print #FileNum, "<kml xmlns=""http://earth.google.com/kml/2.0"">"
Print #FileNum, "<Document>"
Print #FileNum, "<name>" & CBOMONITOR.Text & "</name>"
'Print #FileNum, "<Placemark>"
'
'
'Print #FileNum, " <name>mukti</name>"
'Print #FileNum, "<description><![CDATA["
'
'Print #FileNum, " ]]>"
'Print #FileNum, " </description>"
'Print #FileNum, " <Point>"
'Print #FileNum, "    <coordinates>91.5920023400, 27.1888975300, 0</coordinates>"
'Print #FileNum, "  </Point>"
'Print #FileNum, "</Placemark>"
If rs.EOF <> True Then
mchk = True
retVal = True
Do While rs.EOF <> True
Print #FileNum, " <Placemark>"
Print #FileNum, "   <name> " & rs!farmerbarcode & "(" & rs!FDCODE & ")" & " </name>"
Print #FileNum, "  <description><![CDATA["
     Print #FileNum, "    DATE VISITED: " & Format(rs!Start, "yyyy-MM-dd")
     FindsTAFF rs!staffbarcode
     Print #FileNum, "    MONITOR: " & rs!staffbarcode & " " & sTAFF
     FindFA rs!farmerbarcode, "F"
     Print #FileNum, "    FARMER: " & FAName
Print #FileNum, "    ]]>"
Print #FileNum, " </description>"
Print #FileNum, " <Point>"
Print #FileNum, "   <coordinates>" & rs!GPS_COORDINATES_LNG & "," & rs!GPS_COORDINATES_LAT & " </coordinates>"



Print #FileNum, " </Point>"
Print #FileNum, " <TimeStamp>"
Print #FileNum, "<when>" & Format(rs!Start, "yyyy-MM-ddThh:mm:ssZ") & "</when>"
Print #FileNum, " </TimeStamp>"
Set rschk = Nothing
rschk.Open "select * from tblfarmer where idfarmer='" & rs!farmerbarcode & "'", MHVDB
If rschk.EOF <> True Then
Print #FileNum, " <Style id=""yellow"">"
 Print #FileNum, "  <IconStyle>"
  Print #FileNum, "   <Icon>"
 Print #FileNum, "      <href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png </href>"
   Print #FileNum, "  </Icon>"
  Print #FileNum, " </IconStyle>"
Print #FileNum, " </Style>"

Else
Print #FileNum, " <Style id=""red"">"
 Print #FileNum, "  <IconStyle>"
  Print #FileNum, "   <Icon>"
 Print #FileNum, "      <href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png </href>"
   Print #FileNum, "  </Icon>"
  Print #FileNum, " </IconStyle>"
Print #FileNum, " </Style>"

End If

Print #FileNum, " </Placemark>"
rs.MoveNext
Loop
Else
retVal = False
End If
mchk = False
'Print #FileNum, "<Placemark>"
'Print #FileNum, " <name>nirvana</name>"
'Print #FileNum, " <description>This is location 3</description>"
'Print #FileNum, "  <Point>"
'Print #FileNum, "   <coordinates>91.5920650800,27.1890155200</coordinates>"
'Print #FileNum, " </Point>"
'Print #FileNum, " </Placemark>"


Print #FileNum, "</Document>"
Print #FileNum, "</kml>"

Close #FileNum
    
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Set db1 = New ADODB.Connection
db1.CursorLocation = adUseClient
db1.Open CnnString
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim mdcode, mgcode, mtcode, mfcode As String
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
db.Open OdkCnnString
                       
Set rsmonitor = Nothing

If rsmonitor.State = adStateOpen Then rsmonitor.Close
rsmonitor.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff where moniter='1' order by staffcode", db1
Set CBOMONITOR.RowSource = rsmonitor
CBOMONITOR.ListField = "staffname"
CBOMONITOR.BoundColumn = "staffcode"


txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")


End Sub

Private Sub OPTALL_Click()


Frame1.Enabled = False




End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
End Sub
