VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmqueries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Q U E R I E S . . . "
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10680
   Icon            =   "frmqueries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SELECT ALL"
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
      Left            =   3480
      Picture         =   "frmqueries.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "REVERSE "
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
      Left            =   5040
      Picture         =   "frmqueries.frx":170C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4080
      Width           =   1575
   End
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
      Height          =   4785
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   20
      Top             =   0
      Width           =   3375
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
      Left            =   8040
      Picture         =   "frmqueries.frx":1FD6
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
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
      Left            =   9360
      Picture         =   "frmqueries.frx":2360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4080
      Width           =   1215
   End
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
      Height          =   2655
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin VB.OptionButton Option14 
         Caption         =   "MR_02_Hardlmt_Week"
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
         TabIndex        =   17
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option13 
         Caption         =   "MR_03_HardNGT_Week"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton Option12 
         Caption         =   "MR_07_DeadTC_Week"
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
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
      End
      Begin VB.OptionButton Option10 
         Caption         =   "MR_08_DeadHardNGT_Week"
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
         TabIndex        =   14
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton Option9 
         Caption         =   "MR_09_DeadHardLMT_Week"
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
         TabIndex        =   13
         Top             =   2160
         Width           =   2895
      End
      Begin VB.OptionButton Option8 
         Caption         =   "MR-17_TC_Received"
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
         Left            =   3600
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton Option7 
         Caption         =   "MR_18_HardBags_LMT_Received"
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
         Left            =   3600
         TabIndex        =   11
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton Option6 
         Caption         =   "MR_19_HardNGT_Received"
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
         Left            =   3600
         TabIndex        =   10
         Top             =   1080
         Width           =   3015
      End
      Begin VB.OptionButton Option5 
         Caption         =   "MR_22_CurBal_TC"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   1440
         Width           =   2895
      End
      Begin VB.OptionButton Option4 
         Caption         =   "MR_23_CurBal_Hard _LMT"
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
         Left            =   3600
         TabIndex        =   8
         Top             =   1800
         Width           =   2655
      End
      Begin VB.OptionButton Option3 
         Caption         =   "MR_24_CurBal_NGT"
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
         Left            =   3600
         TabIndex        =   7
         Top             =   2160
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MR_1 TC_ week"
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
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
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
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81330177
         CurrentDate     =   41479
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81330177
         CurrentDate     =   41479
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         TabIndex        =   4
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmqueries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totdr As Double
Dim totcr As Double
Dim Dzstr As String
Private Sub populatefaciltiy()
Dzstr = ""



For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
      End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
   MsgBox "FACILITY NOT SELECTED !!!"
   Exit Sub
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Select Case querytype
    Case 1
           MR qmsQueryPara
    Case 2
           MRTCRCV qmsQueryPara
     Case 3
           MRTCRCV1
     Case 4
           MR1 qmsQueryPara
     Case 8
           currbalTC
     Case 9
           currbalhardlmt
     Case 10
           currbalhardNGT
           
 End Select
End Sub
Private Sub currbalhardNGT()
totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object

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
    excel_sheet.Cells(3, 2) = ProperCase("Variety")
    excel_sheet.Cells(3, 3) = ProperCase("Current Balance")
    
    i = 4
  populatefaciltiy
    SQLSTR = "select varietyid,sum(debit-credit) as debit from tblqmsplanttransaction where status='ON' and entrydate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' and facilityid in " & Dzstr & " group by varietyid having debit>0"
   
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    FindqmsPlantVariety rs!varietyId
    excel_sheet.Cells(i, 2) = qmsPlantVariety
    
    excel_sheet.Cells(i, 3) = rs!debit
    totdr = totdr + rs!debit
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(totdr = 0, "", totdr)
    excel_sheet.Cells(i, 3).Font.Bold = True
 
    fid = ""
    qmshousetype = ""
    
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
    Screen.MousePointer = vbDefaul
End Sub

Private Sub currbalhardlmt()
totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
Set rs = Nothing
totdr = 0

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
    excel_sheet.Cells(3, 2) = ProperCase("Variety")
    excel_sheet.Cells(3, 3) = ProperCase("Current Balance")
    
    i = 4
  populatefaciltiy
    SQLSTR = "select varietyid,sum(debit-credit) as debit from tblqmsplanttransaction where status='ON' and entrydate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' and facilityid in " & Dzstr & " group by varietyid having debit>0"
   
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    FindqmsPlantVariety rs!varietyId
    excel_sheet.Cells(i, 2) = qmsPlantVariety
'    If Trim(rs!varietyid) = "12" Then
'    Set rs1 = Nothing
'    rs1.Open "select sum(debit-credit) as debit from tblqmsplanttransaction where facilityid in('H1','H3','H8','H9','H12') and varietyid='12'", MHVDB
'    excel_sheet.Cells(i, 3) = rs1!debit
'    Else
   excel_sheet.Cells(i, 3) = rs!debit
'    End If
'    If Trim(rs!varietyid) <> "12" Then
    totdr = totdr + rs!debit
'    Else
'    totdr = totdr + rs1!debit
'    End If
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(totdr = 0, "", totdr)
    excel_sheet.Cells(i, 3).Font.Bold = True
 
    fid = ""
    qmshousetype = ""
    
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
    Screen.MousePointer = vbDefaul
End Sub

Private Sub currbalTC()
totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
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
    excel_sheet.Cells(3, 2) = ProperCase("Variety")
    excel_sheet.Cells(3, 3) = ProperCase("Current Balance")
    
    i = 4
    populatefaciltiy
    SQLSTR = "select varietyid,sum(debit-credit) as debit from tblqmsplanttransaction where status='ON' and entrydate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' and facilityid in " & Dzstr & " group by varietyid  having debit>0"
   
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    FindqmsPlantVariety rs!varietyId
    excel_sheet.Cells(i, 2) = qmsPlantVariety
    
    excel_sheet.Cells(i, 3) = rs!debit
    totdr = totdr + rs!debit
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(totdr = 0, "", totdr)
    excel_sheet.Cells(i, 3).Font.Bold = True
 
    fid = ""
    qmshousetype = ""
    
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
    Screen.MousePointer = vbDefaul
End Sub
Private Sub MRTCRCV(querypara As String)
totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
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
    excel_sheet.Cells(3, 2) = ProperCase("Variety")
    excel_sheet.Cells(3, 3) = ProperCase("Plant received")
    
    i = 4
    populatefaciltiy
    SQLSTR = "select varietyid,sum(debit) as debit from tblqmsplanttransaction where status='ON' and transactiontype='" & querypara & "'  and facilityid in " & Dzstr & " group by varietyid"
    
   
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    FindqmsPlantVariety rs!varietyId
    excel_sheet.Cells(i, 2) = qmsPlantVariety
    
    excel_sheet.Cells(i, 3) = rs!debit
    totdr = totdr + rs!debit
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(totdr = 0, "", totdr)
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
    Screen.MousePointer = vbDefaul
End Sub
Private Sub MRTCRCV1()
totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object

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
    excel_sheet.Cells(3, 2) = ProperCase("Variety")
    excel_sheet.Cells(3, 3) = ProperCase("Current Balance")
    
    i = 4
    populatefaciltiy
  
    SQLSTR = "select varietyid,sum(debit-credit) as debit from tblqmsplanttransaction where status='ON' and facilityid in " & Dzstr & " group by varietyid"
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    FindqmsPlantVariety rs!varietyId
    excel_sheet.Cells(i, 2) = qmsPlantVariety
    
    excel_sheet.Cells(i, 3) = rs!debit
    totdr = totdr + rs!debit
 
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 2) = "TOTAL"
        excel_sheet.Cells(i, 2).Font.Bold = True
    excel_sheet.Cells(i, 3) = IIf(totdr = 0, "", totdr)
    excel_sheet.Cells(i, 3).Font.Bold = True
 
    fid = ""
    qmshousetype = ""
    
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
    Screen.MousePointer = vbDefaul
End Sub

Private Sub Command3_Click()
If DZLIST.ListCount > 0 Then
   For i = 0 To DZLIST.ListCount - 1
       If DZLIST.Selected(i) Then
          DZLIST.Selected(i) = False
       Else
          DZLIST.Selected(i) = True
       End If
   Next
End If
End Sub

Private Sub Command4_Click()
Dim i As Long
For i = 0 To DZLIST.ListCount - 1
    DZLIST.Selected(i) = True
Next
End Sub

Private Sub Form_Load()
On Error GoTo err
operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString



Set rsF = Nothing
'
'If rsF.State = adStateOpen Then rsF.Close
'rsF.Open "select locationid,locationname  from tblqmslocation ", db
'Set cbolocation.RowSource = rsF
'cbolocation.ListField = "locationname"
'cbolocation.BoundColumn = "locationid"





Set rsF = Nothing

rsF.Open "select * from tblqmsfacility where status='ON' Order by facilityid", MHVDB, adOpenStatic
With rsF
Do While Not .EOF
   DZLIST.AddItem Trim(!Description) + " | " + !facilityid
   .MoveNext
Loop
End With



Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Option1_Click()
querytype = 1
qmsReportName = "MR_1_TC_week"
qmsQueryPara = 2
Frame4.Visible = True
End Sub
Private Sub MR(querypara As String)

totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
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
    excel_sheet.Cells(3, 2) = ProperCase("date.")
    excel_sheet.Cells(3, 3) = ProperCase("facility")
    excel_sheet.Cells(3, 4) = ProperCase("plant batch")
    excel_sheet.Cells(3, 5) = ProperCase("debit")
    excel_sheet.Cells(3, 6) = ProperCase("credit")
    i = 4
    populatefaciltiy
    
    SQLSTR = "select entrydate,facilityid,plantbatch,varietyid,debit,credit from tblqmsplanttransaction where transactiontype='" & querypara & "' and entryDate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entryDate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' and facilityid in " & Dzstr
    
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    excel_sheet.Cells(i, 2) = "'" & rs!entrydate
    findQmsfacility rs!facilityid
    excel_sheet.Cells(i, 3) = rs!facilityid & "  " & qmsFacility
    findQmsBatchDetail rs!plantBatch
    excel_sheet.Cells(i, 4) = qmsBatchdetail1
    excel_sheet.Cells(i, 5) = IIf(rs!debit = 0, "", rs!debit)
    excel_sheet.Cells(i, 6) = IIf(rs!credit = 0, "", rs!credit)
    totdr = totdr + rs!debit
    totcr = totcr + rs!credit
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 4) = "TOTAL"
        excel_sheet.Cells(i, 4).Font.Bold = True
    excel_sheet.Cells(i, 5) = IIf(totdr = 0, "", totdr)
    excel_sheet.Cells(i, 5).Font.Bold = True
    excel_sheet.Cells(i, 6) = IIf(totcr = 0, "", totcr)
    excel_sheet.Cells(i, 6).Font.Bold = True
    
    
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
Private Sub MR1(querypara As String)

totdr = 0
totcr = 0
Dim fid As String
Dim i, sl As Integer
Dim rs As New ADODB.Recordset
Dim excel_app As Object
Dim excel_sheet As Object
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
    excel_sheet.Cells(3, 2) = ProperCase("date.")
    excel_sheet.Cells(3, 3) = ProperCase("facility")
    excel_sheet.Cells(3, 4) = ProperCase("plant batch")
    excel_sheet.Cells(3, 5) = ProperCase("debit")
    excel_sheet.Cells(3, 6) = ProperCase("credit")
    i = 4
    populatefaciltiy
   
    SQLSTR = "select entrydate,facilityid,plantbatch,varietyid,debit,credit from tblqmsplanttransaction where  transactiontype='" & querypara & "' and entryDate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entryDate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' and facilityid in " & Dzstr
    SQLSTR = SQLSTR & "union select entrydate,facilityid,plantbatch,'' varietyid,0 as debit,noofdead from tblqmsdeadremoval where  entryDate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entryDate<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' and facilityid in " & Dzstr
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    Do While rs.EOF <> True
    excel_sheet.Cells(i, 1) = sl
    excel_sheet.Cells(i, 2) = "'" & rs!entrydate
    findQmsfacility rs!facilityid
    excel_sheet.Cells(i, 3) = rs!facilityid & "  " & qmsFacility
    findQmsBatchDetail rs!plantBatch
    excel_sheet.Cells(i, 4) = qmsBatchdetail1
    excel_sheet.Cells(i, 5) = IIf(IIf(IsNull(rs!debit), 0, rs!debit) = 0, "", rs!debit)
    excel_sheet.Cells(i, 6) = IIf(rs!credit = 0, "", rs!credit)
    totdr = totdr + rs!debit
    totcr = totcr + rs!credit
    sl = sl + 1
    i = i + 1
    rs.MoveNext
    Loop
    End If
  
     excel_sheet.Cells(i, 4) = "TOTAL"
        excel_sheet.Cells(i, 4).Font.Bold = True
    excel_sheet.Cells(i, 5) = IIf(totdr = 0, "", totdr)
    excel_sheet.Cells(i, 5).Font.Bold = True
    excel_sheet.Cells(i, 6) = IIf(totcr = 0, "", totcr)
    excel_sheet.Cells(i, 6).Font.Bold = True
    
    
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


Private Sub Option10_Click()
querytype = 1
qmsReportName = "MR_08_DeadHardNGT_Week"
qmsQueryPara = "3"
Frame4.Visible = True
End Sub

Private Sub Option12_Click()
querytype = 4
qmsReportName = "MR_07_DeadTC_Week"
qmsQueryPara = "3"
Frame4.Visible = True
qmshousetype = "S"
End Sub

Private Sub Option13_Click()
querytype = 1
qmsReportName = "MR_02_Hardlmt_Week"
qmsQueryPara = "9"
Frame4.Visible = True
End Sub

Private Sub Option14_Click()
querytype = 1
qmsReportName = "MR_02_Hardlmt_Week"
qmsQueryPara = "6"
Frame4.Visible = True
End Sub

Private Sub Option3_Click()
querytype = 10
End Sub

Private Sub Option4_Click()
querytype = 9

End Sub

Private Sub Option5_Click()
querytype = 8
End Sub

Private Sub Option6_Click()
querytype = 2
qmsReportName = "MR_19_HardNGT_Received"
qmsQueryPara = "9"
'Frame4.Visible = False

End Sub

Private Sub Option7_Click()
querytype = 2
qmsReportName = "MR_18_HardBags_LMT_Received"
qmsQueryPara = "6"
'Frame4.Visible = False
End Sub

Private Sub Option8_Click()
querytype = 2
qmsReportName = "MR-17_TC_Received"
qmsQueryPara = "2"
Frame4.Visible = False
End Sub

Private Sub Option9_Click()
querytype = 4
qmsReportName = "MR_09_DeadHardLMT_Week"
qmsQueryPara = "3"
Frame4.Visible = True
qmshousetype = "M"
End Sub
