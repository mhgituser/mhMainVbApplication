VERSION 5.00
Begin VB.Form frmhardreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HARD REPORT"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2835
   Icon            =   "frmhardreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox DZLIST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   0
      Width           =   2415
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
      Left            =   240
      Picture         =   "frmhardreport.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
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
      Left            =   1440
      Picture         =   "frmhardreport.frx":11CC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmhardreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim xlApp As Excel.Application
'Dim xlWbook As Excel.Workbook
'Dim xlWSheet As Excel.Worksheet
'Dim xlptCache As Excel.PivotCache
'Dim xlptTable As Excel.PivotTable
'
'
'Private Sub cbofacilityid_Click(area As Integer)
'
'End Sub
'
'Private Sub cbolocation_GotFocus()
'cbofacilityid.Text = ""
'End Sub
'
'Private Sub cbolocation_LostFocus()
'If Len(cbolocation.Text) = 0 Then Exit Sub
'Dim rsF As New ADODB.Recordset
'Set db = New ADODB.Connection
'db.CursorLocation = adUseClient
'db.Open CnnString
'Set rsF = Nothing
'If rsF.State = adStateOpen Then rsF.Close
'rsF.Open "select concat(facilityid,'  ',description) as description,facilityid   from tblqmsfacility where location='" & cbolocation.Text & "' order by facilityid", db
'Set cbofacilityid.RowSource = rsF
'cbofacilityid.ListField = "description"
'cbofacilityid.BoundColumn = "facilityid"
'End Sub
'
'Private Sub Command1_Click()
'Unload Me
'End Sub
'
'Private Sub Command2_Click()
'Command2.Enabled = False
'Dim rs As New ADODB.Recordset
'Dim rshr As New ADODB.Recordset
'Dim rs1 As New ADODB.Recordset
'Dim Dzstr As String
'Dim SQLSTR As String
'Dim excel_app As Object
'Dim excel_sheet As Object
'Dim myfacility As String
'Dim plantBatch As Integer
'Dim i As Integer
'
'Dzstr = ""
'SQLSTR = ""
'
'
'For i = 0 To DZLIST.ListCount - 1
'    If DZLIST.Selected(i) Then
'       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
'       Mcat = DZLIST.List(i)
'       j = j + 1
'    End If
'    If RepName = "5" Then
'       If j > 1 Then
'          MsgBox "SELECT ATLEAST ONE FACILITY TYPE."
'          Exit Sub
'       End If
'    End If
'Next
'If Len(Dzstr) > 0 Then
'   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
'
'Else
'   MsgBox "FACILITY NOT SELECTED !!!"
'   Exit Sub
'End If
'
'
'
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    Set excel_app = CreateObject("Excel.Application")
'    Set Excel_WBook = excel_app.Workbooks.Add
'    If Val(excel_app.Application.Version) >= 8 Then
'        Set excel_sheet = excel_app.ActiveSheet
'    Else
'        Set excel_sheet = excel_app
'    End If
'    'excel_app.Caption = "MHV"
'    Dim sl As Integer
'    sl = 1
'    'excel_app.DisplayFullScreen = True
'    excel_app.Visible = False
'    excel_sheet.cells(3, 1) = "Facility"
'    excel_sheet.cells(3, 2) = "Opening Date"
'    excel_sheet.cells(3, 3) = "Plant Batch"
'    excel_sheet.cells(3, 4) = "Variety"
'
'
'    excel_sheet.cells(3, 5) = ProperCase("Sent to temporary storage") 'ok
'    excel_sheet.cells(3, 6) = ProperCase("Dead Plant Removal") 'ok
'    excel_sheet.cells(3, 7) = ProperCase("Sent To Field")       'ok
'    excel_sheet.cells(3, 8) = ProperCase("Back To nursery")     'ok
'    excel_sheet.cells(3, 9) = ProperCase("Transplant Hard to Bags")    'ok
'    excel_sheet.cells(3, 10) = ProperCase("Debit by PV")        'ok
'    excel_sheet.cells(3, 11) = ProperCase("Hard Moved")     'ok
'    excel_sheet.cells(3, 12) = ProperCase("TC Plantation to Beds") 'ok
'    excel_sheet.cells(3, 13) = ProperCase("Transplant Hard to Beds") 'ok
'    excel_sheet.cells(3, 14) = ProperCase("Nut Moved") 'ok
'    excel_sheet.cells(3, 15) = ProperCase("Transplant Nut to Bags")   'ok
'    excel_sheet.cells(3, 16) = ProperCase("Nut Plantation to Beds")    'ok
'
'     excel_sheet.cells(3, 17) = ProperCase("Sent to Vip/Others")
'
'    excel_sheet.cells(3, 18) = ProperCase("current stock")
'
'    i = 4
'  Set rs = Nothing
'
'    SQLSTR = "select facilityid,plantbatch,varietyid,sum(debit) as debit,sum(credit) as credit  from tblqmsplanttransaction where status<>'C' and facilityid  in (select facilityid from tblqmsfacility where housetype in" & Dzstr & ") group by facilityid,plantbatch order by facilityid,plantbatch"
'    rs.Open SQLSTR, MHVDB
'
'
'
'     Do Until rs.EOF
'
'     myfacility = rs!facilityid
'
'     Do While myfacility = rs!facilityid
'
'
'
'   findQmsfacility rs!facilityid
'   excel_sheet.cells(i, 1) = qmsFacility
'
'   'opening date
'
'   Set rs1 = Nothing
'   rs1.Open "select min(entrydate) entrydate from tblqmsplanttransaction where status<>'C' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "' and debit>0", MHVDB
'   If rs1.EOF <> True Then
'   excel_sheet.cells(i, 2) = "'" & Format(rs1!entrydate, "dd/MM/yyyy")
'   Else
'   excel_sheet.cells(i, 2) = ""
'   End If
'   findQmsBatchDetail rs!plantBatch
'   excel_sheet.cells(i, 3) = rs!plantBatch
'   excel_sheet.cells(i, 4) = qmsplantbatch3
'
'
'
''   Set rshr = Nothing
''   rshr.Open "select transactiontype,sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and  plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'  group by transactiontype", MHVDB
''
''
''   Do While rshr.EOF <> True
''
''
''
''
''
''   If rshr!transactiontype = 2 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 5) = rshr!qty
''            excel_sheet.Cells(i, 5).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 5) = ""
''            End If
''   ElseIf rshr!transactiontype = 3 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 6) = rshr!qty
''            excel_sheet.Cells(i, 6).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 6) = ""
''            End If
''   ElseIf rshr!transactiontype = 4 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 7) = rshr!qty
''            excel_sheet.Cells(i, 7).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 7) = ""
''            End If
''   ElseIf rshr!transactiontype = 5 Then
''           If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 8) = rshr!qty
''            excel_sheet.Cells(i, 8).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 8) = ""
''            End If
''   ElseIf rshr!transactiontype = 6 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 9) = rshr!qty
''            excel_sheet.Cells(i, 9).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 9) = ""
''            End If
''   ElseIf rshr!transactiontype = 7 Then
''           If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 10) = rshr!qty
''            excel_sheet.Cells(i, 10).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 10) = ""
''            End If
''   ElseIf rshr!transactiontype = 8 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 11) = rshr!qty
''            excel_sheet.Cells(i, 11).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 11) = ""
''            End If
''   ElseIf rshr!transactiontype = 9 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 12) = rshr!qty
''            excel_sheet.Cells(i, 12).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 12) = ""
''            End If
''   ElseIf rshr!transactiontype = 10 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 13) = rshr!qty
''            excel_sheet.Cells(i, 13).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 13) = ""
''            End If
''   ElseIf rshr!transactiontype = 11 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 14) = rshr!qty
''            excel_sheet.Cells(i, 14).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 14) = ""
''            End If
''   ElseIf rshr!transactiontype = 12 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 15) = rshr!qty
''            excel_sheet.Cells(i, 15).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 15) = ""
''            End If
''   ElseIf rshr!transactiontype = 13 Then
''            If rshr!qty <> 0 Then
''            excel_sheet.Cells(i, 16) = rshr!qty
''            excel_sheet.Cells(i, 16).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
''            Else
''            excel_sheet.Cells(i, 16) = ""
''            End If
''   Else
''
''   End If
''
''   rshr.MoveNext
''   Loop
'
'
'   '14 tc plantation to bed
'
'    Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='14' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 5) = rshr!qty
'     excel_sheet.cells(i, 5).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 5) = ""
'   End If
'
'    ' transactiontype 3, dead plant removal
'   Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='3' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 6) = rshr!qty
'     excel_sheet.cells(i, 6).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 6) = ""
'   End If
'
'    ' transactiontype 4, sent to field
'   Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='4' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 7) = rshr!qty
'   excel_sheet.cells(i, 7).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 7) = ""
'   End If
'
'
'   ' transactiontype 5, back to nursary
'   Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='5' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 8) = rshr!qty
'     excel_sheet.cells(i, 8).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 8) = ""
'   End If
'
'   ' transactiontype 6, transplant to bag
'    Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction  where status<>'C' and transactiontype='6' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 9) = rshr!qty
'     excel_sheet.cells(i, 9).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 9) = ""
'   End If
'
'    '7 debot by pv
'
'    Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='8' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 10) = rshr!qty
'     excel_sheet.cells(i, 10).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 10) = ""
'   End If
'
'    ' hard move
'   Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='9' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 11) = rshr!qty
'     excel_sheet.cells(i, 11).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 11) = ""
'   End If
'
'    ' tc plantation to bed
'   Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='2' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 12) = rshr!qty
'     excel_sheet.cells(i, 12).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 12) = ""
'   End If
'
'     ' 10 plantation hard to beds
'     Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='10' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 13) = rshr!qty
'     excel_sheet.cells(i, 13).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 13) = ""
'   End If
'
'   '11 nut moved
'   Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='11' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 14) = rshr!qty
'     excel_sheet.cells(i, 14).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 14) = ""
'   End If
'
'
'
'   ' 12 transplant nut to bags
'    Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='12' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 15) = rshr!qty
'     excel_sheet.cells(i, 15).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 15) = ""
'   End If
'   ' 13 nut plantation to bed
'
'     Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='13' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 16) = rshr!qty
'     excel_sheet.cells(i, 16).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 16) = ""
'   End If
'
'    ' sent to vip/others
'    Set rshr = Nothing
'   rshr.Open "select sum(debit-credit) as qty from tblqmsplanttransaction where status<>'C' and transactiontype='15' and plantbatch='" & rs!plantBatch & "' and facilityid='" & rs!facilityid & "'", MHVDB
'   If rshr.EOF <> True Then
'   excel_sheet.cells(i, 17) = rshr!qty
'   excel_sheet.cells(i, 17).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   Else
'   excel_sheet.cells(i, 17) = ""
'   End If
'
'
'
'   excel_sheet.cells(i, 18) = Val(excel_sheet.cells(i, 5)) + Val(excel_sheet.cells(i, 6)) + Val(excel_sheet.cells(i, 7)) + Val(excel_sheet.cells(i, 8)) + Val(excel_sheet.cells(i, 9)) + Val(excel_sheet.cells(i, 10)) + Val(excel_sheet.cells(i, 11)) + Val(excel_sheet.cells(i, 12)) + Val(excel_sheet.cells(i, 13)) + Val(excel_sheet.cells(i, 14)) + Val(excel_sheet.cells(i, 15)) + Val(excel_sheet.cells(i, 16)) + Val(excel_sheet.cells(i, 17))
'   excel_sheet.cells(i, 18).NumberFormat = """""#,##0_);[BLACK]\(#,##)"
'   If Val(excel_sheet.cells(i, 18)) <> 0 Then
'    i = i + 1
'   End If
'
'
'    rs.MoveNext
'    If rs.EOF Then Exit Do
'       Loop
'        Loop
'
'
'    'make up
'   excel_sheet.Range(excel_sheet.cells(3, 1), _
'    excel_sheet.cells(i, 18)).Select
'    excel_app.selection.Columns.AutoFit
'   ' Freeze the header row so it doesn't scroll.
'    excel_sheet.cells(4, 2).Select
'    excel_app.ActiveWindow.FreezePanes = True
'    excel_sheet.cells(1, 1).Select
'    With excel_sheet
'    '.PageSetup.LeftHeader = "MHV"
'     excel_sheet.Range("A3:q3").Font.Bold = True
'    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
'    .PageSetup.CenterFooter = "HARD REPORT"
'        .PageSetup.LeftFooter = "MHV"
'        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
'        .PageSetup.PrintGridlines = True
'    End With
'
'
'    '
'    ''excel_sheet.Cells(1, 1) = "MHV"
'    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
'
'
'    Screen.MousePointer = vbDefault
'
'
'excel_app.Visible = True
'Set excel_sheet = Nothing
'Set excel_app = Nothing
'Screen.MousePointer = vbDefault
'
'Exit Sub
'err:
'MsgBox err.Description
'err.Clear
'End Sub
'
'Private Sub Form_Load()
'Command2.Enabled = True
'Dim rs As New ADODB.Recordset
'
'Set rs = Nothing
'
'rs.Open "select distinct housetype from tblqmsfacility  Order by housetype", MHVDB, adOpenStatic
'With rs
'Do While Not .EOF
'
'Select Case rs!housetype
'            Case "C"
'            DZLIST.AddItem "Cold House" + " | " + Trim(!housetype)
'            Case "N"
'            DZLIST.AddItem "Net House" + " | " + Trim(!housetype)
'            Case "H"
'            DZLIST.AddItem "Hoop House" + " | " + Trim(!housetype)
'            Case "T"
'            DZLIST.AddItem "Terrace" + " | " + Trim(!housetype)
'            Case "S"
'            DZLIST.AddItem "Staging House" + " | " + Trim(!housetype)
'End Select
'
'   .MoveNext
'Loop
'End With
'End Sub
