VERSION 5.00
Begin VB.Form FRMPLANTHISTORY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLNAT HISTORY"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3735
   Icon            =   "FRMPLANTHISTORY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Query Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton Option1 
         Caption         =   "Plant Sent to Field(LMT)"
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Avg. Time Spent(LMT) sent to ngt"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Avg. Time spent(LMT)"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Plant Sent to Field (NGT)"
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
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
      Left            =   1800
      Picture         =   "FRMPLANTHISTORY.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
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
      Left            =   480
      Picture         =   "FRMPLANTHISTORY.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "FRMPLANTHISTORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reportcase As Integer
Dim totdr As Double

Dim totcr As Double

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Select Case reportcase
            Case 1
            plantsenttofieldLMT
            Case 2
            plantsenttofieldNGT
            Case 3
            avgtimespentLMT
            Case 4
            avgtimespentNGT

End Select


End Sub
Private Sub avgtimespentLMT()
Dim avgdays As Integer
Dim mdays As Integer
Dim mqty As Double
Dim batchno As Integer
Dim plantcount As Double
Dim CRTOT As Double
Dim mrow As Integer
totdr = 0
totcr = 0
Dim rs1 As New ADODB.Recordset
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Set rs = Nothing

rs.Open "select * from tblqmsfacility where housetype in('H','N')", MHVDB
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
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("Batch No.")
    excel_sheet.Cells(3, 3) = ProperCase("received Date")
    excel_sheet.Cells(3, 4) = ProperCase("sent to field date")
    excel_sheet.Cells(3, 5) = ProperCase("Plants in batch")
    excel_sheet.Cells(3, 6) = ProperCase("plants in shipment")
    excel_sheet.Cells(3, 7) = ProperCase("Qty. sent")
    excel_sheet.Cells(3, 8) = ProperCase("currrent stock")
    excel_sheet.Cells(3, 9) = ProperCase("no. of days in (LMT)")
    i = 4
  
   
    SQLSTR = "select entrydate,plantbatch,credit from tblqmsplanttransaction where status='ON' and  transactiontype='4' and facilityid in " & fid & "  order by plantbatch,entrydate"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
    
       Do Until rs.EOF
       avgdays = 0
       mdays = 0
       mqty = 0
        batchno = rs!plantbatch
        excel_sheet.Cells(i, 1) = sl
        findQmsBatchDetail rs!plantbatch
        excel_sheet.Cells(i, 2) = qmsBatchdetail1
        findQmsBatchReceivedDate rs!plantbatch, "LMT"
        excel_sheet.Cells(i, 2).Font.Bold = True
        excel_sheet.Cells(i, 3) = "'" & Format(qmsBatchReceivedDate, "dd/MM/yyyy")
        excel_sheet.Cells(i, 5) = qmsbatchTotal
       excel_sheet.Cells(i, 6) = qmsshiptotal
       Set rs1 = Nothing
      rs1.Open "select sum(debit-credit) as bal from tblqmsplanttransaction where plantbatch='" & rs!plantbatch & "'", MHVDB
      
       Do While batchno = rs!plantbatch
         excel_sheet.Cells(i, 4) = "'" & rs!entrydate
         excel_sheet.Cells(i, 7) = rs!credit
         excel_sheet.Cells(i, 9) = DateDiff("d", Format(qmsBatchReceivedDate, "yyyy-MM-dd"), Format(rs!entrydate, "yyyy-MM-dd"))
         mdays = mdays + DateDiff("d", Format(qmsBatchReceivedDate, "yyyy-MM-dd"), Format(rs!entrydate, "yyyy-MM-dd"))
         avgdays = avgdays + 1
         mqty = mqty + rs!credit
         
        i = i + 1
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       
      i = i + 1
      sl = sl + 1
'      Set rs1 = Nothing
'      rs1.Open "select sum(debit-credit) as bal from tblqmsplanttransaction where plantbatch='" & rs!plantbatch & "'", MHVDB
'
      If rs1.EOF <> True Then
      excel_sheet.Cells(i - 1, 8) = IIf(rs1!bal <= 0, "", rs1!bal)
      excel_sheet.Cells(i - 1, 8).Font.Bold = True
      Else
       excel_sheet.Cells(i - 1, 8) = ""
      End If
     
      excel_sheet.Cells(i - 1, 7) = mqty
      excel_sheet.Cells(i - 1, 7).Font.Bold = True
      excel_sheet.Cells(i - 1, 9) = Round(((mdays / avgdays) * qmsbatchTotal) / qmsshiptotal, 0)
      excel_sheet.Cells(i - 1, 9).Font.Bold = True
 
    Loop
    
    
    
    End If
  
'     excel_sheet.Cells(i, 3) = "TOTAL"
'        excel_sheet.Cells(i, 3).Font.Bold = True
'    excel_sheet.Cells(i, 4) = IIf(totcr = 0, "", totcr)
'    excel_sheet.Cells(i, 4).Font.Bold = True
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = qmsReportName
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    
    
    Excel_WBook.Sheets("sheet2").Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.Name = "Summary"
    
    
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("Batch No.")
    excel_sheet.Cells(3, 3) = ProperCase("NO. OF PLANTS SENT TO FIELD(LMT)")
      excel_sheet.Cells(3, 4) = ProperCase("Detail")
    i = 4
  
   CRTOT = 0
    SQLSTR = "select entrydate,plantbatch,credit from tblqmsplanttransaction where status='ON' and  transactiontype='4' and facilityid in " & fid & "  order by plantbatch,entrydate"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
    
       Do Until rs.EOF
       avgdays = 0
       mdays = 0
       plantcount = 0
        batchno = rs!plantbatch
        excel_sheet.Cells(i, 1) = sl
        findQmsBatchDetail rs!plantbatch
        excel_sheet.Cells(i, 2) = qmsBatchdetail1
        findQmsBatchReceivedDate rs!plantbatch, "LMT"
   
       Do While batchno = rs!plantbatch

         mdays = mdays + DateDiff("d", Format(qmsBatchReceivedDate, "yyyy-MM-dd"), Format(rs!entrydate, "yyyy-MM-dd"))
         avgdays = avgdays + 1
         plantcount = plantcount + rs!credit
         CRTOT = CRTOT + rs!credit
   
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
  
  excel_sheet.Cells(i, 3) = plantcount
  
       excel_sheet.Cells(i, 4).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT( D3&" & Chr(34) & "!B:B" & Chr(34) & "),MATCH(" & "B" & i & ",INDIRECT( D3 &" & Chr(34) & "!B:B" & Chr(34) & "),0)))," & Chr(34) & Round(((mdays / avgdays) * qmsbatchTotal) / qmsshiptotal, 0) & Chr(34) & ")"
       
      i = i + 1
      sl = sl + 1

 
    Loop
    
    
    
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(CRTOT = 0, "", CRTOT)
    excel_sheet.Cells(i, 3).Font.Bold = True
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = qmsReportName
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    
    
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    
    
    
    
    
    
    
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault

End Sub
Private Sub avgtimespentNGT()
Dim avgdays As Integer
Dim mdays As Integer
Dim mqty As Double
Dim batchno As Integer
Dim plantcount As Double
Dim CRTOT As Double
Dim mrow As Integer
totdr = 0
totcr = 0
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
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("Batch No.")
    excel_sheet.Cells(3, 3) = ProperCase("received Date")
    excel_sheet.Cells(3, 4) = ProperCase("sent to field date")
    excel_sheet.Cells(3, 5) = ProperCase("Plants in batch")
    excel_sheet.Cells(3, 6) = ProperCase("plants in shipment")
    excel_sheet.Cells(3, 7) = ProperCase("Qty. sent")
    excel_sheet.Cells(3, 8) = ProperCase("currrent stock")
    excel_sheet.Cells(3, 9) = ProperCase("no. of days in (LMT)")
    i = 4
  
   
    SQLSTR = "select entrydate,plantbatch,debit from tblqmsplanttransaction where status='ON' and  transactiontype='9' and debit>0 and facilityid in " & fid & "  order by plantbatch,entrydate"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
    
       Do Until rs.EOF
       avgdays = 0
       mdays = 0
       mqty = 0
        batchno = rs!plantbatch
        excel_sheet.Cells(i, 1) = sl
        findQmsBatchDetail rs!plantbatch
        excel_sheet.Cells(i, 2) = qmsBatchdetail1
        findQmsBatchReceivedDate rs!plantbatch, "LMT"
        excel_sheet.Cells(i, 2).Font.Bold = True
        excel_sheet.Cells(i, 3) = "'" & Format(qmsBatchReceivedDate, "dd/MM/yyyy")
        excel_sheet.Cells(i, 5) = qmsbatchTotal
       excel_sheet.Cells(i, 6) = qmsshiptotal
      
       Do While batchno = rs!plantbatch
       If rs!plantbatch = 142 Then
       
       MsgBox "AD"
       End If
         Set rs1 = Nothing
         rs1.Open "select * from tblqmsplanttransaction where transactiontype='9' and plantbatch='" & rs!plantbatch & "' and entrydate='" & Format(rs!entrydate, "yyyy-MM-dd") & "' and credit>0 AND facilityid in " & fid & "", MHVDB
      If rs1.EOF <> True Then
      Else
         excel_sheet.Cells(i, 4) = "'" & rs!entrydate
         excel_sheet.Cells(i, 7) = rs!debit
         excel_sheet.Cells(i, 9) = DateDiff("d", Format(qmsBatchReceivedDate, "yyyy-MM-dd"), Format(rs!entrydate, "yyyy-MM-dd"))
         mdays = mdays + DateDiff("d", Format(qmsBatchReceivedDate, "yyyy-MM-dd"), Format(rs!entrydate, "yyyy-MM-dd"))
         avgdays = avgdays + 1
         mqty = mqty + rs!debit
         
        i = i + 1
        End If
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
       
      i = i + 1
      sl = sl + 1
'      Set rs1 = Nothing
'      rs1.Open "select sum(debit-credit) as bal from tblqmsplanttransaction where plantbatch='" & rs!plantbatch & "'", MHVDB
'
'      If rs1.EOF <> True Then
'      excel_sheet.Cells(i - 1, 8) = IIf(rs1!bal <= 0, "", rs1!bal)
'      excel_sheet.Cells(i - 1, 8).Font.Bold = True
'      Else
'       excel_sheet.Cells(i - 1, 8) = ""
'      End If
     
      excel_sheet.Cells(i - 1, 7) = mqty
      excel_sheet.Cells(i - 1, 7).Font.Bold = True
      excel_sheet.Cells(i - 1, 9) = Round(((mdays / avgdays) * qmsbatchTotal) / qmsshiptotal, 0)
      excel_sheet.Cells(i - 1, 9).Font.Bold = True
 
    Loop
    
    
    
    End If
  
'     excel_sheet.Cells(i, 3) = "TOTAL"
'        excel_sheet.Cells(i, 3).Font.Bold = True
'    excel_sheet.Cells(i, 4) = IIf(totcr = 0, "", totcr)
'    excel_sheet.Cells(i, 4).Font.Bold = True
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = qmsReportName
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    
    
    Excel_WBook.Sheets("sheet2").Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.Name = "Summary"
    
    
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("Batch No.")
    excel_sheet.Cells(3, 3) = ProperCase("NO. OF PLANTS SENT TO FIELD(LMT)")
      excel_sheet.Cells(3, 4) = ProperCase("Detail")
    i = 4
  
   CRTOT = 0
    SQLSTR = "select entrydate,plantbatch,debit from tblqmsplanttransaction where status='ON' and  transactiontype='9' and facilityid in " & fid & "  order by plantbatch,entrydate"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
    
       Do Until rs.EOF
       avgdays = 0
       mdays = 0
       plantcount = 0
        batchno = rs!plantbatch
        excel_sheet.Cells(i, 1) = sl
        findQmsBatchDetail rs!plantbatch
        excel_sheet.Cells(i, 2) = qmsBatchdetail1
        findQmsBatchReceivedDate rs!plantbatch, "LMT"
   
       Do While batchno = rs!plantbatch
       Set rs1 = Nothing
  rs1.Open "select * from tblqmsplanttransaction where transactiontype='9' and plantbatch='" & rs!plantbatch & "' and entrydate='" & Format(rs!entrydate, "yyyy-MM-dd") & "' and credit>0 and   facilityid in " & fid & "", MHVDB
      If rs1.EOF <> True Then
      Else
         mdays = mdays + DateDiff("d", Format(qmsBatchReceivedDate, "yyyy-MM-dd"), Format(rs!entrydate, "yyyy-MM-dd"))
         avgdays = avgdays + 1
         plantcount = plantcount + rs!debit
         CRTOT = CRTOT + rs!debit
   End If
          rs.MoveNext
          If rs.EOF Then Exit Do
       Loop
  
  excel_sheet.Cells(i, 3) = plantcount
  
       excel_sheet.Cells(i, 4).Formula = "=HYPERLINK(" & Chr(34) & "#" & Chr(34) & "&CELL(" & Chr(34) & "address" & Chr(34) & ",INDEX(INDIRECT( D3&" & Chr(34) & "!B:B" & Chr(34) & "),MATCH(" & "B" & i & ",INDIRECT( D3 &" & Chr(34) & "!B:B" & Chr(34) & "),0)))," & Chr(34) & Round(((mdays / avgdays) * qmsbatchTotal) / qmsshiptotal, 0) & Chr(34) & ")"
       
      i = i + 1
      sl = sl + 1

 
    Loop
    
    
    
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(CRTOT = 0, "", CRTOT)
    excel_sheet.Cells(i, 3).Font.Bold = True
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = qmsReportName
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    
    
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    
    
    
    
    
    
    
    
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub plantsenttofieldLMT()

totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Set rs = Nothing

rs.Open "select * from tblqmsfacility where housetype in('H','N')", MHVDB
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
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("Batch No.")
    excel_sheet.Cells(3, 3) = ProperCase("Facility id")
    excel_sheet.Cells(3, 4) = ProperCase("#s of plant sent")

    i = 4
  
   
    SQLSTR = "select plantbatch,facilityid,sum(debit) as debit,sum(credit) as credit from tblqmsplanttransaction where status='ON' and  transactiontype='4' and facilityid in " & fid & " group by plantbatch,facilityid order by plantbatch,facilityid"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    findQmsBatchDetail rs!plantbatch
    excel_sheet.Cells(i, 2) = qmsBatchdetail1
    findQmsfacility rs!facilityid
     excel_sheet.Cells(i, 3) = qmsFacility
    'excel_sheet.Cells(i, 4) = IIf(rs!debit = 0, "", rs!debit)
    excel_sheet.Cells(i, 4) = IIf(rs!credit = 0, "", rs!credit)
    totdr = totdr + rs!debit
    totcr = totcr + rs!credit
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 3) = "TOTAL"
        excel_sheet.Cells(i, 3).Font.Bold = True
    excel_sheet.Cells(i, 4) = IIf(totcr = 0, "", totcr)
    excel_sheet.Cells(i, 4).Font.Bold = True
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = qmsReportName
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub plantsenttofieldNGT()

totdr = 0
totcr = 0
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
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("Batch No.")
    excel_sheet.Cells(3, 3) = ProperCase("Facility id")
    excel_sheet.Cells(3, 4) = ProperCase("#s plants sent")
   
    i = 4
  
   
    SQLSTR = "select plantbatch,facilityid,sum(debit) as debit,sum(credit) as credit from tblqmsplanttransaction where status='ON' and  transactiontype='4' and facilityid in " & fid & " group by plantbatch,facilityid order by plantbatch,facilityid"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    findQmsBatchDetail rs!plantbatch
    excel_sheet.Cells(i, 2) = qmsBatchdetail1
    findQmsfacility rs!facilityid
     excel_sheet.Cells(i, 3) = qmsFacility
    'excel_sheet.Cells(i, 4) = IIf(rs!debit = 0, "", rs!debit)
    excel_sheet.Cells(i, 4) = IIf(rs!credit = 0, "", rs!credit)
    totdr = totdr + rs!debit
    totcr = totcr + rs!credit
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 3) = "TOTAL"
        excel_sheet.Cells(i, 3).Font.Bold = True
'    excel_sheet.Cells(i, 4) = IIf(totdr = 0, "", totdr)
'    excel_sheet.Cells(i, 4).Font.Bold = True
    excel_sheet.Cells(i, 4) = IIf(totcr = 0, "", totcr)
    excel_sheet.Cells(i, 4).Font.Bold = True
    
    
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:i3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = qmsReportName
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    Screen.MousePointer = vbDefault
    'excel_app.Visible = True
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub Option1_Click()
reportcase = 1
End Sub

Private Sub Option12_Click()
reportcase = 4
End Sub

Private Sub Option13_Click()
reportcase = 3
End Sub

Private Sub Option14_Click()
reportcase = 2
End Sub
