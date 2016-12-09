VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMSHIPMENTBYBOX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TC SHIPMENT REPORT BY BOX"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3690
   Icon            =   "FRMSHIPMENTBYBOX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkall 
      Caption         =   "Show All"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Report"
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
      Left            =   720
      Picture         =   "FRMSHIPMENTBYBOX.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   2040
      Picture         =   "FRMSHIPMENTBYBOX.frx":11CC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cbotrnid 
      Bindings        =   "FRMSHIPMENTBYBOX.frx":1E96
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Shipment No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1395
   End
End
Attribute VB_Name = "FRMSHIPMENTBYBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkall_Click()
If chkall.Value = 1 Then
cbotrnid.Text = ""
cbotrnid.Enabled = False
Else
cbotrnid.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
SHIPMENTBOXDETAIL
End Sub

Private Sub Form_Load()
On Error GoTo err
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnid  as description,trnid  from tblqmsplantbatchhdr order by trnid", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "trnid"





Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub SHIPMENTBOXDETAIL()
Dim mshipmentno As Integer
Dim avgdays As Integer
Dim mdays As Integer
Dim batchno As Integer
Dim plantcount As Double
Dim CRTOT As Double
Dim totshipmentsize, tothealthy, totweak, totundersize, totoversize, toticedamage, totdead, totavgsize, totavgroot As Double
Dim received, totreceived As Double
Dim reccount As Integer
Dim mrow As Integer
reccount = 0
totshipmentsize = 0
tothealthy = 0
totweak = 0
totundersize = 0
totoversize = 0
toticedamage = 0
totdead = 0
totavgsize = 0
totavgroot = 0
totreceived = 0
received = 0
Dim rs1 As New ADODB.Recordset
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Set rs = Nothing

rs.Open "select * from tblqmsfacility where location='NGT'", MHVDB
If rs.EOF <> True Then
Do While rs.EOF <> True
fid = fid + "'" + rs!facilityid + "',"
rs.MoveNext
Loop
Else


End If
If Len(fid) > 0 Then
fid = "(" + Left(fid, Len(fid) - 1) + ")"
End If
    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_sheet.Name = "Detail"
    sl = 1
    i = 1
    excel_app.Visible = False
     'excel_sheet.Cells(2, 1) = "Shipment No. " & cbotrnid.Text
     'excel_sheet.Cells(2, 1).Font.Bold = True
     
    excel_sheet.Cells(3, 1) = ProperCase("Date")
    excel_sheet.Cells(3, 2) = ProperCase("Box No.")
    excel_sheet.Cells(3, 3) = ProperCase("Plant Batch")
    excel_sheet.Cells(3, 4) = ProperCase("TC")
    excel_sheet.Cells(3, 5) = ProperCase("Variety")
    excel_sheet.Cells(3, 6) = ProperCase("B/L Number Plants")
    excel_sheet.Cells(3, 7) = ProperCase("Healthy Plants")
    excel_sheet.Cells(3, 8) = ProperCase("Weak Plants")
    excel_sheet.Cells(3, 9) = ProperCase("Under Size")
    excel_sheet.Cells(3, 10) = ProperCase("Over Size")
    excel_sheet.Cells(3, 11) = ProperCase("Ice Damaged")
    excel_sheet.Cells(3, 12) = ProperCase("Dead Plants")
     excel_sheet.Cells(3, 13) = ProperCase("total received")
    excel_sheet.Cells(3, 14) = ProperCase("Avg. Size(cm)")
    excel_sheet.Cells(3, 15) = ProperCase("Avg. Root(cm)")
    excel_sheet.Cells(3, 16) = ProperCase("Date Box Planted")
    excel_sheet.Cells(3, 17) = ProperCase("Comments")
  
    i = 4
  
   If chkall.Value = 1 Then
    SQLSTR = "select * from tblqmsboxdetail where status='ON' and location='" & Mlocation & "' order by trnid, convert(boxno,unsigned integer)"
    Else
    If Len(cbotrnid.Text) = 0 Then
    MsgBox "Select Shipment No."
    Exit Sub
    End If
    SQLSTR = "select * from tblqmsboxdetail where status='ON' and location='" & Mlocation & "' and  trnid='" & cbotrnid.BoundText & "' order by convert(boxno,unsigned integer)"
    End If
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do Until rs.EOF
    mshipmentno = rs!trnid
    i = i + 1
     excel_sheet.Cells(i, 1) = "Shipment No. " & mshipmentno
     excel_sheet.Cells(i, 1).Font.Bold = True
     i = i + 1
     
     reccount = 0
     totshipmentsize = 0
tothealthy = 0
totweak = 0
totundersize = 0
totoversize = 0
toticedamage = 0
totdead = 0
totavgsize = 0
totavgroot = 0
totreceived = 0
     
     
    Do While mshipmentno = rs!trnid
    
    received = rs!healthyplant + rs!weakplant + rs!undersize + rs!oversize + rs!icedamaged + rs!deadplant
    excel_sheet.Cells(i, 1) = "'" & Format(rs!entrydate, "dd/MM/yyyy")
    excel_sheet.Cells(i, 2) = rs!boxno
    excel_sheet.Cells(i, 3) = rs!plantBatch
    FindqmsPlanttype rs!planttype
    excel_sheet.Cells(i, 4) = qmsPlantType
    FindqmsPlantVariety rs!plantvariety
    excel_sheet.Cells(i, 5) = qmsPlantVariety
    excel_sheet.Cells(i, 6) = Format(rs!shipmentsize, "##,##,##")
    excel_sheet.Cells(i, 7) = Format(rs!healthyplant, "##,##,##")
    excel_sheet.Cells(i, 8) = Format(rs!weakplant, "##,##,##")
    excel_sheet.Cells(i, 9) = Format(rs!undersize, "##,##,##")
    excel_sheet.Cells(i, 10) = Format(rs!oversize, "##,##,##")
    excel_sheet.Cells(i, 11) = Format(rs!icedamaged, "##,##,##")
    excel_sheet.Cells(i, 12) = Format(rs!deadplant, "##,##,##")
    excel_sheet.Cells(i, 13) = Format(received, "##,##,##")
    excel_sheet.Cells(i, 14) = rs!avgsize
    excel_sheet.Cells(i, 14).NumberFormat = "##0.00"
    excel_sheet.Cells(i, 15) = rs!avgroot
    excel_sheet.Cells(i, 15).NumberFormat = "##0.00"
    excel_sheet.Cells(i, 16) = "'" & Format(rs!dateplantedbox, "dd/MM/yyyy")
    excel_sheet.Cells(i, 17) = rscomments
    
    
reccount = reccount + 1
totshipmentsize = totshipmentsize + rs!shipmentsize
tothealthy = tothealthy + rs!healthyplant
totweak = totweak + rs!weakplant
totundersize = totundersize + rs!undersize
totoversize = totoversize + rs!oversize
toticedamage = toticedamage + rs!icedamaged
totdead = totdead + rs!deadplant
totavgsize = totavgsize + rs!avgsize
totavgroot = totavgroot + rs!avgroot
totreceived = totreceived + received
    
 received = 0
 i = i + 1
 'sl = sl + 1
 rs.MoveNext
 If rs.EOF Then Exit Do
    Loop
     
     excel_sheet.Cells(i, 5) = "TOTAL"
     excel_sheet.Cells(i, 5).Font.Bold = True
     excel_sheet.Cells(i, 6) = IIf(totshipmentsize = 0, "", Format(totshipmentsize, "##,##,##"))
     excel_sheet.Cells(i, 6).Font.Bold = True
     excel_sheet.Cells(i, 7) = IIf(tothealthy = 0, "", Format(tothealthy, "##,##,##"))
     excel_sheet.Cells(i, 7).Font.Bold = True
     excel_sheet.Cells(i, 8) = IIf(totweak = 0, "", Format(totweak, "##,##,##"))
     excel_sheet.Cells(i, 8).Font.Bold = True
     excel_sheet.Cells(i, 9) = IIf(totundersize = 0, "", Format(totundersize, "##,##,##"))
     excel_sheet.Cells(i, 9).Font.Bold = True
     excel_sheet.Cells(i, 10) = IIf(totoversize = 0, "", Format(totoversize, "##,##,##"))
     excel_sheet.Cells(i, 10).Font.Bold = True
     excel_sheet.Cells(i, 11) = IIf(toticedamage = 0, "", Format(toticedamage, "##,##,##"))
     excel_sheet.Cells(i, 11).Font.Bold = True
     excel_sheet.Cells(i, 12) = IIf(totdead = 0, "", Format(totdead, "##,##,##"))
     excel_sheet.Cells(i, 12).Font.Bold = True
     
     excel_sheet.Cells(i, 13) = IIf(totreceived = 0, "", Format(totreceived, "##,##,##"))
     excel_sheet.Cells(i, 13).Font.Bold = True
     
     excel_sheet.Cells(i, 14) = IIf(Round(totavgsize / reccount, 2) = 0, "", Round(totavgsize / reccount, 2))
     excel_sheet.Cells(i, 14).NumberFormat = "##0.00"
     excel_sheet.Cells(i, 14).Font.Bold = True
     excel_sheet.Cells(i, 15) = IIf(Round(totavgroot / reccount, 2) = 0, "", Round(totavgroot / reccount, 2))
     excel_sheet.Cells(i, 15).NumberFormat = "##0.00"
     excel_sheet.Cells(i, 15).Font.Bold = True
     
     
     Loop
     
     
    
    End If
  
  
    
    
    'make up
'    excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(i, 6)).Select
'    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:p3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "SHIPMENT BOX DETAIL"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    
    
   'excel_sheet.Columns("A:p").Select
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
   excel_sheet.Cells(3, 16)).Select
 excel_app.Selection.ColumnWidth = 10
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
    
    
    excel_app.Visible = True
    
    
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault

End Sub


