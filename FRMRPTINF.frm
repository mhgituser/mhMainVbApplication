VERSION 5.00
Begin VB.Form FRMRPTINF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFLUENTIAL REPORT"
   ClientHeight    =   2535
   ClientLeft      =   7965
   ClientTop       =   2070
   ClientWidth     =   3030
   Icon            =   "FRMRPTINF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3030
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW"
      Height          =   735
      Left            =   360
      Picture         =   "FRMRPTINF.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   1560
      Picture         =   "FRMRPTINF.frx":15AC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton OPTFARMERLISTING 
         Caption         =   "FARMER"
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
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ABSENTEE"
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
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "ALL"
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
         Left            =   600
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   45
   End
End
Attribute VB_Name = "FRMRPTINF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
Set rs = Nothing
If FATYPEINF = "F" Then
rs.Open "SELECT * FROM tblinfluential WHERE FATYPE='F' ORDER BY FARMERID", MHVDB


ElseIf FATYPEINF = "A" Then
rs.Open "SELECT * FROM tblinfluential WHERE FATYPE='A' ORDER BY FARMERID", MHVDB

ElseIf FATYPEINF = "O" Then
rs.Open "SELECT * FROM tblinfluential  ORDER BY FATYPE, FARMERID", MHVDB


Else
MsgBox "INVALID TYPE SELECTION, PLEASE SELECT THE APPROPRIATE OPTION FROM THE MENU."
Exit Sub
End If


'Dim RS As New ADODB.Recordset
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
    excel_sheet.Cells(3, 1) = "SL.NO."
    If FATYPEINF = "F" Then
    excel_sheet.Cells(3, 2) = "FARMER ID"
    excel_sheet.Cells(3, 3) = "FARMER NAME"
    ElseIf FATYPEINF = "A" Then
       excel_sheet.Cells(3, 2) = "ABSENTEE ID"
    excel_sheet.Cells(3, 3) = "ABSENTEE NAME"
    Else
       excel_sheet.Cells(3, 2) = "FARMER/ABSENTEE ID"
    excel_sheet.Cells(3, 3) = "FARMER/ABSENTEE NAME"
    
    End If
    
    excel_sheet.Cells(3, 4) = "JOB TITLE"
    excel_sheet.Cells(3, 5) = "DEPARTMENT"
    excel_sheet.Cells(3, 6) = "IMPORTAINT RELATIVES"
    i = 4
  'Set RS = Nothing
 
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
     excel_sheet.Cells(i, 2) = rs!FARMERID
     FindFA rs!FARMERID, rs!FATYPE
   excel_sheet.Cells(i, 3) = FAName
   
 excel_sheet.Cells(i, 4) = rs!JOBTITLE
 excel_sheet.Cells(i, 5) = rs!dept
 excel_sheet.Cells(i, 6) = rs!RELATION
   i = i + 1
   sl = sl + 1
   rs.MoveNext
   Loop
   
    'make up
   excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:f3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    
    If FATYPEINF = "F" Then
    .PageSetup.CenterFooter = " INFLUENTIAL(FARMER)"
    ElseIf FATYPEINF = "A" Then
     .PageSetup.CenterFooter = " INFLUENTIAL(ABSENTEE)"
    Else
    .PageSetup.CenterFooter = " INFLUENTIAL(FARMER AND ABSENTEE)"
    End If
    
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
ERR:
MsgBox ERR.Description
ERR.Clear





End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
FATYPEINF = ""
End Sub

Private Sub OPTFARMERLISTING_Click()
FATYPEINF = "F"
End Sub

Private Sub Option1_Click()
FATYPEINF = "A"
End Sub

Private Sub Option2_Click()
FATYPEINF = "O"
End Sub
