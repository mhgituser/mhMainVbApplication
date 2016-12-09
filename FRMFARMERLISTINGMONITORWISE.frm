VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMFARMERLISTINGMONITORWISE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MONITOR WISE FARMER LISTING"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkall 
      Caption         =   "All"
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
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW"
      Height          =   735
      Left            =   960
      Picture         =   "FRMFARMERLISTINGMONITORWISE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1455
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
      Left            =   2520
      Picture         =   "FRMFARMERLISTINGMONITORWISE.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo CBOMONITOR 
      Bindings        =   "FRMFARMERLISTINGMONITORWISE.frx":1054
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   960
      TabIndex        =   2
      Top             =   360
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "MONITOR"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   765
   End
End
Attribute VB_Name = "FRMFARMERLISTINGMONITORWISE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsmonitor As New ADODB.Recordset

Private Sub chkall_Click()
If chkall.Value = 1 Then
cbomonitor.Enabled = False
Else
cbomonitor.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Len(cbomonitor.Text) > 0 And chkall.Value = 1 Then
MsgBox "Invalid Selection of monitors."
Exit Sub
End If
Dim rsm As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
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
    excel_sheet.Cells(3, 1) = "Sl. No."
    'excel_sheet.Cells(3, 2) = ProperCase("Dzongkhag code")
    'excel_sheet.Cells(3, 3) = ProperCase("Dzongkhag name")
    'excel_sheet.Cells(3, 4) = ProperCase("Gewog code")
    'excel_sheet.Cells(3, 5) = ProperCase("gewog name")
    'excel_sheet.Cells(3, 6) = ProperCase("tshowog code")
    'excel_sheet.Cells(3, 7) = ProperCase("tshowog name")
    excel_sheet.Cells(3, 2) = ProperCase("farmer code")
    excel_sheet.Cells(3, 3) = ProperCase("farmer name")
   ' excel_sheet.Cells(3, 10) = ProperCase("reg. area")
    'excel_sheet.Cells(3, 11) = ProperCase("Trees 2011")
   ' excel_sheet.Cells(3, 12) = ProperCase("trees 2012")
    'excel_sheet.Cells(3, 13) = ProperCase("total trees")
    
    i = 4
    Dim tempmonitor As String
    tempmonitor = ""
  Set rs = Nothing
  If chkall.Value = 1 Then
  rs.Open "SELECT idfarmer ,monitor FROM tblfarmer  where status not in('D','R') and length(monitor)=5 ORDER BY monitor,idfarmer", MHVDB
  
  Else
  rs.Open "SELECT idfarmer ,monitor FROM tblfarmer  where where status not in('D','R') and length(monitor)=5 and monitor='" & cbomonitor.BoundText & "' ORDER BY farmercode", MHVDB
  End If
   Do Until rs.EOF
  tempmonitor = rs!monitor
  FindsTAFF rs!monitor
  excel_sheet.Cells(i - 2, 1) = "Farmers under " & rs!monitor & " " & sTAFF
  excel_sheet.Cells(i - 2, 1).Font.Bold = True
  Do While tempmonitor = rs!monitor
   
   excel_sheet.Cells(i, 1) = sl
'   excel_sheet.Cells(i, 2) = Mid(rs!idfarme, 1, 3)
'   FindDZ Mid(rs!farmercode, 1, 3)
'   excel_sheet.Cells(i, 3) = Dzname
'   excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 4, 3)
'   FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
'   excel_sheet.Cells(i, 5) = GEname
'    excel_sheet.Cells(i, 6) = Mid(rs!farmercode, 7, 3)
'    FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
'    excel_sheet.Cells(i, 7) = TsName
    excel_sheet.Cells(i, 2) = rs!idfarmer
    FindFA rs!idfarmer, "F"
    excel_sheet.Cells(i, 3) = FAName

' Set rs1 = Nothing
'    rs1.Open "select sum(regland) as regland from tbllandreg where farmerid='" & rs!farmercode & "'", MHVDB
'    If rs1.EOF <> True Then
'     excel_sheet.Cells(i, 10) = IIf(IsNull(rs1!regland), 0, rs1!regland)
'    Else
'     excel_sheet.Cells(i, 10) = ""
'    End If
'
' Set rs1 = Nothing
'    rs1.Open "select nooftrees from tblplanted where farmercode='" & rs!farmercode & "' and year='2011'", MHVDB
'    If rs1.EOF <> True Then
'     excel_sheet.Cells(i, 11) = IIf(IsNull(rs1!nooftrees), 0, rs1!nooftrees)
'    Else
'     excel_sheet.Cells(i, 11) = ""
'    End If
'
' Set rs1 = Nothing
'    rs1.Open "select nooftrees from tblplanted where farmercode='" & rs!farmercode & "' and year='2012'", MHVDB
'    If rs1.EOF <> True Then
'     excel_sheet.Cells(i, 12) = IIf(IsNull(rs1!nooftrees), 0, rs1!nooftrees)
'    Else
'     excel_sheet.Cells(i, 12) = ""
'    End If
'
'   excel_sheet.Cells(i, 13) = Val(excel_sheet.Cells(i, 11)) + Val(excel_sheet.Cells(i, 12))
'
    
    
    
 
 
   i = i + 1
   sl = sl + 1
   rs.MoveNext
   If rs.EOF Then Exit Do
   Loop
   
   sl = 1
   i = i + 3
   Loop
   
    'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 13)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:m3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "FARMER LISTING"
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

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
                       
Set rsmonitor = Nothing

If rsmonitor.State = adStateOpen Then rsmonitor.Close
rsmonitor.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff where moniter='1' order by staffcode", db
Set cbomonitor.RowSource = rsmonitor
cbomonitor.ListField = "staffname"
cbomonitor.BoundColumn = "staffcode"

End Sub
