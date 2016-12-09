VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmchemicalchecklistprn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chemical Check List Print"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6210
   Icon            =   "frmchemicalchecklistprn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Check List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Picture         =   "frmchemicalchecklistprn.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin MSDataListLib.DataCombo cbotrnid 
      Bindings        =   "frmchemicalchecklistprn.frx":11CC
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
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
      Caption         =   "Schedule Id"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
End
Attribute VB_Name = "frmchemicalchecklistprn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MCOL As Long
Private Sub Command1_Click()
printchecklist
End Sub

Private Sub Form_Load()
Dim RSTR As New ADODB.Recordset
maxDistNo = 0
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(cast(trnid as char) ,' ',distributionname,' ',cast(year as char),' ',cast(mnth as char)) as dname,trnid  from tblplantdistributionheader where status='ON' and planneddist='Y' order by trnid desc", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "dname"
cbotrnid.BoundColumn = "trnid"
End Sub

Private Sub printchecklist()
'On Error Resume Next
Dim s As Integer
Dim m, n As Integer
Dim SQLSTR As String
Dim totplant As Integer
Dim myphone As String
Dim TOTLAND As Double
Dim tcode As String
Dim totadd As Double
Dim lastrow As Long
TOTLAND = 0
Dim mm
totadd = 0
totplant = 0
Dzstr = ""
SQLSTR = ""
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

mchk = True
j = 0
                    
   Dim tdist As Integer
                        
                        
                        

Dim excel_app As Object
Dim excel_sheet As Object

Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_sheet = Nothing
    Set excel_app = Nothing
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
    i = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = False
    '
    excel_sheet.Cells(2, 2) = "Fertilizer Check List"
    
    excel_sheet.Cells(3, 1) = "S/N"
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "FARMER"
    excel_sheet.Cells(3, 6) = "LAND (ACRE)"
    'premix
    excel_sheet.Cells(3, 7) = ("Total (Kg)")
   ' rs!totalkg1
    ' dolomite
    'excel_sheet.Cells(3, 23) = UCase("Kg")
    excel_sheet.Name = "Chemical Checklist"
    i = 4
    s = 4
    SQLSTR = "select * from tblplantdistributiondetail where subtotindicator<>'T' and trnid='" & cbotrnid.BoundText & "' order by sno"
    
                        tdist = 0
                            Set rs = Nothing
                            rs.Open SQLSTR, MHVDB
                            mchk = True
                            Do While rs.EOF <> True
                            'excel_sheet.Cells(i, 1) = rs!distno '"D/N"
                            FindDZ Mid(rs!farmercode, 1, 3)
                            FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
                            FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
                             FindFA rs!farmercode, "F"
                             excel_sheet.Cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
                             excel_sheet.Cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
                             excel_sheet.Cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
                             excel_sheet.Cells(i, 5) = rs!farmercode & "  " & FAName
                             excel_sheet.Cells(i, 6) = rs!area '"LAND (ACRE)"
                             excel_sheet.Cells(i, 7) = rs!totalkg1
                             ' dolomite stores in column 70-- temp store
                             excel_sheet.Cells(i, 70) = rs!kg
                             
                             
                             
                             
                             
                             
                             
                             'tdist = 0
                             If rs!subtotindicator = "" Then
                             tdist = rs!distno
                             End If
                             
                            If rs!subtotindicator = "S" Then
                            
                             excel_sheet.Range(excel_sheet.Cells(i, 2), _
                             excel_sheet.Cells(i, 26)).Select
                             excel_app.Selection.Interior.ColorIndex = 15
                             
                                                       
                             excel_sheet.Range(excel_sheet.Cells(s, 1), _
                             excel_sheet.Cells(i - 1, 1)).Select
                             
                                excel_sheet.Cells(s, 1) = tdist
                            With excel_app.Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            'Selection.Merge
                            
                            
                             excel_sheet.Cells(i, 1) = ""
                            
                            
    
                            
                            
                                                   
                             
                             
                             
                            
                            s = i + 1
                            End If
                            
                            i = i + 1
                            
                            rs.MoveNext
                            Loop
                                
                         
                          createpremix excel_sheet
                            createdolomite excel_sheet
                            pagebreak excel_sheet
                            'make up
'                            excel_sheet.Range(excel_sheet.Cells(4, 8), _
'                            excel_sheet.Cells(i, 8)).Select
'                            excel_app.Selection.NumberFormat = "####0.00"
'                            excel_sheet.Range(excel_sheet.Cells(1, 12), _
'                            excel_sheet.Cells(1, 16)).Select
'
'
'
'                            With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'
'                            excel_sheet.Cells(1, 12) = "Variety"
'
'
'
'
'                            excel_sheet.Range(excel_sheet.Cells(1, 17), _
'                             excel_sheet.Cells(1, 22)).Select
'
'
'
'                            With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(1, 17) = "Pre - Mixed Fertilizer"
'
'
'
'                            excel_sheet.Range(excel_sheet.Cells(1, 23), _
'                             excel_sheet.Cells(1, 24)).Select
'
'
'
'                            With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'
'                            excel_sheet.Cells(1, 23) = "Dolomite"
'
'
'            excel_sheet.Range(excel_sheet.Cells(2, 16), _
'                             excel_sheet.Cells(2, 18)).Select
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'
'                            excel_sheet.Cells(2, 16) = "Mixed Variety"
'
'                            excel_sheet.Range(excel_sheet.Cells(2, 26), _
'                             excel_sheet.Cells(3, 27)).Select
'
'
'
'                           With excel_app.Selection
'                                .HorizontalAlignment = xlCenter
'                                .VerticalAlignment = xlCenter 'xlBottom
'                                .WrapText = False
'                                .Orientation = 0
'                                .AddIndent = False
'                                .IndentLevel = 0
'                                .ShrinkToFit = False
'                                .ReadingOrder = xlContext
'                                .MergeCells = True
'                            End With
'
'                            excel_sheet.Cells(2, 26) = "Schedule Date, Vehicle No & Team Captainy"
'
'
'
'
'
'
'
'
'
' excel_sheet.Range(excel_sheet.Cells(1, 1), _
'                             excel_sheet.Cells(i, 27)).Select
''excel_sheet.Columns("A:A").Select
' excel_app.Selection.Font.Size = 10
'
'
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'
'
'
'
'
'
'
'
'
'
'
'excel_sheet.Columns("A:A").Select
' excel_app.Selection.ColumnWidth = 3.57
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'
'
'    excel_sheet.Columns("b:d").Select
' excel_app.Selection.ColumnWidth = 14.86
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'
'
'
'
'excel_sheet.Columns("e:f").Select
' excel_app.Selection.ColumnWidth = 17
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'
'
'excel_sheet.Columns("g:Y").Select
' excel_app.Selection.ColumnWidth = 8
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'
'excel_sheet.Columns("Z:Z").Select
' excel_app.Selection.ColumnWidth = 7
'With excel_app.Selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
'.WrapText = True
'.Orientation = 0
'End With
'
'
'
'
'
'
'
'
'
'
'     excel_sheet.Range(excel_sheet.Cells(1, 1), _
'                             excel_sheet.Cells(i, 27)).Select
'
'                   excel_app.Selection.Font.Name = "Times New Roman"
'   excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(i, 6)).Select
'    excel_app.Selection.Columns.AutoFit
'   ' Freeze the header row so it doesn't scroll.
'    excel_sheet.Cells(4, 2).Select
'    excel_app.ActiveWindow.FreezePanes = True
'    excel_sheet.Cells(1, 1).Select
'    With excel_sheet
'    '.PageSetup.LeftHeader = "MHV"
'     excel_sheet.Range("A1:Z3").Font.Bold = True
'    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
'    .PageSetup.CenterFooter = "PLANT DISTRIBUTION LIST"
'        .PageSetup.LeftFooter = "MHV"
'        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
'        .PageSetup.PrintGridlines = True
'    End With
'
    
  

   

excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

End Sub
Private Sub createpremix(excel_sheet As Object)
Dim lastrow, lastcolumn As Long
Dim i  As Integer
Dim tcol As Long
Dim j As Integer
Dim col As Integer
Dim bagcnt As Integer
bagcnt = 0
 MCOL = 0
lastrow = excel_sheet.UsedRange.Rows.Count
lastcolumn = excel_sheet.UsedRange.Columns.Count
excel_sheet.Cells(3, 8) = "28 Kg Bag"
excel_sheet.Cells(3, 9) = "Kg Left"
For i = 4 To lastrow
excel_sheet.Cells(i, 9) = excel_sheet.Cells(i, 7) Mod 28
excel_sheet.Cells(i, 8) = (excel_sheet.Cells(i, 7) - excel_sheet.Cells(i, 9)) / 28
    If Len(Trim(excel_sheet.Cells(i, 5).Value)) > 0 Then
       col = 10
      
        For j = 1 To Val(excel_sheet.Cells(i, 8).Value)
            bagcnt = bagcnt + 1
            excel_sheet.Cells(i, col).Value = "28"
                       
            excel_sheet.Cells(3, col).Value = "Bag" & bagcnt
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Nursery"
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Monitor"
           col = col + 1
            
            
           
            
            
        Next
            If Val(excel_sheet.Cells(i, 9).Value) > 0 Then
             bagcnt = bagcnt + 1
             excel_sheet.Cells(i, col).Value = Val(excel_sheet.Cells(i, 9).Value)
                       
            excel_sheet.Cells(3, col).Value = "Bag" & bagcnt
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Nursery"
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Monitor"
           col = col + 1
            
            
               
            End If
    End If
    tcol = col
    bagcnt = 0
    If tcol > MCOL Then
    MCOL = tcol
    End If
Next
End Sub
Private Sub createdolomite(excel_sheet As Object)
'Dim i  As Integer
'Dim j As Integer
'Dim col As Integer
'Dim bagcnt As Integer
'bagcnt = 0
'
'For i = 4 To 2027
'    If Len(Sheets("August Distribution").Cells(i, 6).Value) > 0 Then
'        col = 82
'
'        For j = 1 To Val(Sheets("August Distribution").Cells(i, 80).Value)
'            bagcnt = bagcnt + 1
'            Sheets("August Distribution").Cells(i, col).Value = "48"
'
'            Sheets("August Distribution").Cells(3, col).Value = "Bag" & bagcnt
'            col = col + 1
'            Sheets("August Distribution").Cells(3, col).Value = "Nursery"
'            col = col + 1
'            Sheets("August Distribution").Cells(3, col).Value = "Monitor"
'           col = col + 1
'
'
'
'
'
'        Next
'            If Val(Sheets("August Distribution").Cells(i, 81).Value) > 0 Then
'                Sheets("August Distribution").Cells(i, col).Value = Val(Sheets("August Distribution").Cells(i, 81).Value)
'            End If
'    End If
'    bagcnt = 0
'Next

Dim lastrow, lastcolumn As Long
Dim i  As Integer
Dim j As Integer
Dim col As Integer
Dim bagcnt As Integer
bagcnt = 0

lastrow = excel_sheet.UsedRange.Rows.Count
lastcolumn = excel_sheet.UsedRange.Columns.Count
excel_sheet.Cells(3, MCOL + 1) = "Kg"
excel_sheet.Cells(3, MCOL + 1) = "48 Kg Bag"
excel_sheet.Cells(3, MCOL + 2) = "Kg Left"
For i = 4 To lastrow
excel_sheet.Cells(i, MCOL) = excel_sheet.Cells(i, 70)
excel_sheet.Cells(i, MCOL + 2) = excel_sheet.Cells(i, MCOL) Mod 48
excel_sheet.Cells(i, MCOL + 1) = (excel_sheet.Cells(i, MCOL) - excel_sheet.Cells(i, MCOL + 2)) / 48
    If Len(Trim(excel_sheet.Cells(i, 5).Value)) > 0 Then
       col = MCOL + 3
        
        For j = 1 To Val(excel_sheet.Cells(i, MCOL + 1).Value)
            bagcnt = bagcnt + 1
            excel_sheet.Cells(i, col).Value = "48"
                       
            excel_sheet.Cells(3, col).Value = "Bag" & bagcnt
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Nursery"
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Monitor"
           col = col + 1
            
            
           
            
            
        Next
            If Val(excel_sheet.Cells(i, MCOL + 2).Value) > 0 Then
             bagcnt = bagcnt + 1
             excel_sheet.Cells(i, col).Value = Val(excel_sheet.Cells(i, MCOL + 2).Value)
            excel_sheet.Cells(3, col).Value = "Bag" & bagcnt
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Nursery"
            col = col + 1
            excel_sheet.Cells(3, col).Value = "Monitor"
           col = col + 1
            
            
               
            End If
    End If
    bagcnt = 0
Next


excel_sheet.Columns(70).Delete




End Sub


Private Sub pagebreak(excel_sheet As Object)
Dim lastrow, lastcolumn As Long
Dim i  As Integer
Dim j As Integer
Dim col As Integer
Dim bagcnt As Integer
bagcnt = 0

lastrow = excel_sheet.UsedRange.Rows.Count
lastcolumn = excel_sheet.UsedRange.Columns.Count

'Set excel_sheet = xlApp.ActiveSheet
For i = 4 To lastrow
    If Len(Trim(excel_sheet.Cells(i, 5).Value)) = 0 Then
     excel_sheet.HPageBreaks.Add Before:=excel_sheet.Cells(i + 1, 3)
     
     End If
    Next
End Sub
