VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMFARMERLISTING 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FARMER LISTING"
   ClientHeight    =   7245
   ClientLeft      =   6975
   ClientTop       =   1665
   ClientWidth     =   8130
   Icon            =   "FRMFARMERLISTING.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8130
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
      Height          =   5175
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   7575
      Begin VB.OptionButton Option14 
         Caption         =   "Formhub choice"
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
         Left            =   3480
         TabIndex        =   23
         Top             =   3120
         Width           =   3735
      End
      Begin VB.OptionButton Option13 
         Caption         =   "PLANTED LIST"
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
         TabIndex        =   22
         Top             =   2760
         Width           =   2295
      End
      Begin VB.OptionButton Option12 
         Caption         =   "MONITOR LISTING(Form Hub)"
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
         Left            =   3480
         TabIndex        =   21
         Top             =   2760
         Width           =   3015
      End
      Begin VB.OptionButton Option11 
         Caption         =   "TSHOWOG LISTING(Form Hub)"
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
         Left            =   3480
         TabIndex        =   20
         Top             =   2400
         Width           =   3015
      End
      Begin VB.OptionButton Option10 
         Caption         =   "GRF LISTING"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   2040
         Width           =   2175
      End
      Begin VB.OptionButton Option9 
         Caption         =   "FARMER LISTING(DROPOUT AND REJECTED)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton Option8 
         Caption         =   "FARMER LISTING(NEW)"
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
         TabIndex        =   17
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton Option7 
         Caption         =   "FARMER WITH ZERO REGISTERED LAND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3480
         TabIndex        =   16
         Top             =   840
         Width           =   3975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CG  LISTING"
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
         Left            =   3480
         TabIndex        =   15
         Top             =   600
         Width           =   2295
      End
      Begin VB.Frame Frame6 
         Caption         =   "SORT BY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   4800
         TabIndex        =   12
         Top             =   4440
         Visible         =   0   'False
         Width           =   2295
         Begin VB.OptionButton OPTBYID 
            Caption         =   " ID"
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
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton OPTNAME 
            Caption         =   "NAME"
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
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.OptionButton Option5 
         Caption         =   "INDIVIDUAL FARMER REGISTRATION"
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
         Left            =   3480
         TabIndex        =   10
         Top             =   1320
         Width           =   3855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "TSHOWOG LISTING"
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
         Left            =   3480
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Caption         =   "GEWOG LISTING"
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
         TabIndex        =   8
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DZONGKHAG LISTING"
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
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ABSENTEE LISTING"
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
         TabIndex        =   6
         Top             =   1800
         Width           =   2295
      End
      Begin VB.OptionButton OPTFARMERLISTING 
         Caption         =   "FARMER LISTING"
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
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "FRMFARMERLISTING.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   600
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
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
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   5160
      Picture         =   "FRMFARMERLISTING.frx":0E57
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   3720
      Picture         =   "FRMFARMERLISTING.frx":11E1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW"
      Height          =   735
      Left            =   2520
      Picture         =   "FRMFARMERLISTING.frx":1EAB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   45
   End
End
Attribute VB_Name = "FRMFARMERLISTING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Dzname As String
'Dim GEname As String
Dim rsfr As New ADODB.Recordset
Private Sub Command1_Click()
On Error GoTo err
Select Case RptOption
       Case "FL"
       FL
       Case "AB"
       
       Case "DZ"
       dl
       Case "GE"
       gl
       Case "TS"
       tl
       Case "CG"
       CG
       Case "Z"
       ZE
       Case "IFR"
       PRINTFINFO
       Case "FLN"
       FLN
       Case "FDO"
       FDO
       Case "GRF"
       grf
        Case "FH"
       fh
       Case "ML"
       ml
       Case "PL"
       PL
       Case "FHC"
       fhc
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub PL()
Dim rs As New ADODB.Recordset
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
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "FARMER ID"
    excel_sheet.Cells(3, 3) = "FARMER NAME"
    excel_sheet.Cells(3, 4) = "CID#"
    excel_sheet.Cells(3, 5) = "VILLAGE"
    excel_sheet.Cells(3, 6) = "LOCATION NAME"
    excel_sheet.Cells(3, 7) = "REG. LAND"
    i = 4
  Set rs = Nothing
  rs.Open "SELECT distinct idfarmer,farmername,cidno,village,locationname,sum(regland) as regland FROM  tblfarmer as a, tbllandreg as b  where farmerid=idfarmer and idfarmer in(select farmercode from tblplanted) group by idfarmer ORDER BY IDFARMER", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
     excel_sheet.Cells(i, 2) = rs!idfarmer
   excel_sheet.Cells(i, 3) = rs!farmername
 excel_sheet.Cells(i, 4) = rs!cidno
 excel_sheet.Cells(i, 5) = rs!VILLAGE
 excel_sheet.Cells(i, 6) = rs!LocationName
 excel_sheet.Cells(i, 7) = rs!regland
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
    .PageSetup.CenterFooter = "FARMER LISTING"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
     excel_app.Visible = False
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
Private Sub ml()

Dim rs As New ADODB.Recordset
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
    Dim dzcode As String
    Dim gcode As String
    gcode = ""
    dzcode = ""
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    'excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "S Code"
    excel_sheet.Cells(3, 3) = "S Name"
    excel_sheet.Cells(3, 4) = "Status"
    i = 4
  Set rs = Nothing
  rs.Open "SELECT * FROM tblmhvstaff where moniter='1'", MHVDB
  
  
   Do While rs.EOF <> True
   

'   FindsTAFF rs!staffcode
   excel_sheet.Cells(i, 2) = rs!staffcode
   excel_sheet.Cells(i, 3) = rs!staffname
   excel_sheet.Cells(i, 4) = ""
   i = i + 1
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
    .PageSetup.CenterFooter = "TSHOWOG LISTING"
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
'err:
'MsgBox err.Description
''err.Clear
End Sub

Private Sub fh()
Dim rs As New ADODB.Recordset
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
    Dim dzcode As String
    Dim gcode As String
    gcode = ""
    dzcode = ""
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    'excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "Dzongkhag code"
    excel_sheet.Cells(3, 3) = "Gewog code"
    excel_sheet.Cells(3, 4) = "Tshowog code"
    i = 4
  Set rs = Nothing
    rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  ' rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  ' rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  ' rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  
   Do While rs.EOF <> True
   

   FindDZ rs!dzongkhagid
   excel_sheet.Cells(i, 2) = rs!dzongkhagid & "   " & Dzname
   FindGE rs!dzongkhagid, rs!gewogid
   excel_sheet.Cells(i, 3) = rs!gewogid & " " & GEname 'rs!tshewogid
   excel_sheet.Cells(i, 4) = rs!tshewogid & " " & rs!tshewogname
   i = i + 1
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
    .PageSetup.CenterFooter = "TSHOWOG LISTING"
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
Private Sub grf()
Dim rs As New ADODB.Recordset
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
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "DGT"
    excel_sheet.Cells(3, 6) = "FARMER ID"
    excel_sheet.Cells(3, 7) = "FARMER NAME"
    excel_sheet.Cells(3, 8) = "CID#"
    excel_sheet.Cells(3, 9) = "VILLAGE"
    excel_sheet.Cells(3, 10) = "LOCATION NAME"
    excel_sheet.Cells(3, 11) = "REG. LAND"
    i = 4
  Set rs = Nothing
  rs.Open "SELECT idfarmer,farmername,cidno,village,locationname,sum(regland) as regland FROM tblfarmer a ,tbllandreg b   where substring(idfarmer,10,1)='G' and idfarmer=farmerid and a.status='A' and idfarmer  not in(select farmercode from tblplanted) group by idfarmer ORDER BY IDFARMER", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
   FindDZ Mid(rs!idfarmer, 1, 3)
   excel_sheet.Cells(i, 2) = Mid(rs!idfarmer, 1, 3) & " " & Dzname
   FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
   excel_sheet.Cells(i, 3) = Mid(rs!idfarmer, 4, 3) & " " & GEname
   FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
   excel_sheet.Cells(i, 4) = Mid(rs!idfarmer, 7, 3) & " " & TsName
   excel_sheet.Cells(i, 5) = Mid(rs!idfarmer, 1, 9)
   excel_sheet.Cells(i, 6) = rs!idfarmer
   excel_sheet.Cells(i, 7) = rs!farmername
   excel_sheet.Cells(i, 8) = rs!cidno
   excel_sheet.Cells(i, 9) = rs!VILLAGE
   excel_sheet.Cells(i, 10) = rs!LocationName
   excel_sheet.Cells(i, 11) = rs!regland
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
    .PageSetup.CenterFooter = "FARMER LISTING"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
    excel_app.Visible = True
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
Private Sub FDO()
Dim rs As New ADODB.Recordset
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
      excel_sheet.Cells(3, 2) = "DZONGKHAG"
        excel_sheet.Cells(3, 3) = "GEWOG"
          excel_sheet.Cells(3, 4) = "TSHOWOG"
            excel_sheet.Cells(3, 5) = "DGT"
    
    excel_sheet.Cells(3, 6) = "FARMER ID"
    excel_sheet.Cells(3, 7) = "FARMER NAME"
    excel_sheet.Cells(3, 8) = "CID#"
    excel_sheet.Cells(3, 9) = "VILLAGE"
    excel_sheet.Cells(3, 10) = "LOCATION NAME"
    excel_sheet.Cells(3, 11) = "REG. LAND"
    excel_sheet.Cells(3, 12) = "STATUS"
    i = 4
  Set rs = Nothing
  rs.Open "SELECT idfarmer,farmername,cidno,village,locationname,regland,a.status FROM tblfarmer a ,tbllandreg b   where idfarmer=farmerid and a.status in('D','R')  ORDER BY IDFARMER", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
   FindDZ Mid(rs!idfarmer, 1, 3)
   excel_sheet.Cells(i, 2) = Mid(rs!idfarmer, 1, 3) & " " & Dzname
   FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
   excel_sheet.Cells(i, 3) = Mid(rs!idfarmer, 4, 3) & " " & GEname
   FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
   excel_sheet.Cells(i, 4) = Mid(rs!idfarmer, 7, 3) & " " & TsName
   excel_sheet.Cells(i, 5) = Mid(rs!idfarmer, 1, 9)
   excel_sheet.Cells(i, 6) = rs!idfarmer
   excel_sheet.Cells(i, 7) = rs!farmername
   excel_sheet.Cells(i, 8) = rs!cidno
   excel_sheet.Cells(i, 9) = rs!VILLAGE
   excel_sheet.Cells(i, 10) = rs!LocationName
   excel_sheet.Cells(i, 11) = rs!regland
   excel_sheet.Cells(i, 12) = rs!status
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

Private Sub FLN()
Dim rs As New ADODB.Recordset
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
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "DGT"
    excel_sheet.Cells(3, 6) = "FARMER ID"
    excel_sheet.Cells(3, 7) = "FARMER NAME"
    excel_sheet.Cells(3, 8) = "CID#"
    excel_sheet.Cells(3, 9) = "VILLAGE"
    excel_sheet.Cells(3, 10) = "LOCATION NAME"
    excel_sheet.Cells(3, 11) = "REG. LAND"
    i = 4
  Set rs = Nothing
  rs.Open "SELECT idfarmer,farmername,cidno,village,locationname,sum(regland) as regland FROM tblfarmer a ,tbllandreg b   where idfarmer=farmerid and a.status='A' and idfarmer  not in(select farmercode from tblplanted) group by idfarmer ORDER BY IDFARMER", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
   FindDZ Mid(rs!idfarmer, 1, 3)
   excel_sheet.Cells(i, 2) = Mid(rs!idfarmer, 1, 3) & " " & Dzname
   FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
   excel_sheet.Cells(i, 3) = Mid(rs!idfarmer, 4, 3) & " " & GEname
   FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
   excel_sheet.Cells(i, 4) = Mid(rs!idfarmer, 7, 3) & " " & TsName
   excel_sheet.Cells(i, 5) = Mid(rs!idfarmer, 1, 9)
   excel_sheet.Cells(i, 6) = rs!idfarmer
   excel_sheet.Cells(i, 7) = rs!farmername
   excel_sheet.Cells(i, 8) = rs!cidno
   excel_sheet.Cells(i, 9) = rs!VILLAGE
   excel_sheet.Cells(i, 10) = rs!LocationName
   excel_sheet.Cells(i, 11) = rs!regland
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
    .PageSetup.CenterFooter = "FARMER LISTING"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
    excel_app.Visible = True
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
Private Sub ZE()
Dim rs As New ADODB.Recordset
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
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOGE"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "FARMER CODE"
     excel_sheet.Cells(3, 6) = "FARMER NAME"
     excel_sheet.Cells(3, 7) = "REG. LAND"
  
    i = 4
  Set rs = Nothing
  rs.Open "SELECT farmerid,sum(regland) as regland FROM tbllandreg WHERE regland is null and status='A' group by farmerid order by farmerid ", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
   FindDZ Mid(rs!FARMERID, 1, 3)
   excel_sheet.Cells(i, 2) = Mid(rs!FARMERID, 1, 3) & " " & Dzname
   FindGE Mid(rs!FARMERID, 1, 3), Mid(rs!FARMERID, 4, 3)
   excel_sheet.Cells(i, 3) = Mid(rs!FARMERID, 4, 3) & " " & GEname
   FindTs Mid(rs!FARMERID, 1, 3), Mid(rs!FARMERID, 4, 3), Mid(rs!FARMERID, 7, 3)
   excel_sheet.Cells(i, 4) = Mid(rs!FARMERID, 7, 3) & " " & TsName
    excel_sheet.Cells(i, 5) = rs!FARMERID
    FindFA rs!FARMERID, "F"
     excel_sheet.Cells(i, 6) = FAName
      excel_sheet.Cells(i, 7) = IIf(IsNull(rs!regland), "", rs!regland)
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
     excel_sheet.Range("A3:g3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "CG LISTING"
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
'ERR:
'MsgBox ERR.Description

End Sub
Private Sub CG()
Dim rs As New ADODB.Recordset
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
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOGE"
    excel_sheet.Cells(3, 4) = "TSHOWOG"
    excel_sheet.Cells(3, 5) = "CG CODE"
     excel_sheet.Cells(3, 6) = "CG NAME"
     excel_sheet.Cells(3, 7) = "NUMBER"
  
    i = 4
  Set rs = Nothing
  rs.Open "SELECT distinct idfarmer, farmername, phone1 FROM `tblfarmer` WHERE Isfarmercg =1 order by idfarmer", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
   FindDZ Mid(rs!idfarmer, 1, 3)
   excel_sheet.Cells(i, 2) = Mid(rs!idfarmer, 1, 3) & " " & Dzname
   FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
   excel_sheet.Cells(i, 3) = Mid(rs!idfarmer, 4, 3) & " " & GEname
   FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
   excel_sheet.Cells(i, 4) = Mid(rs!idfarmer, 7, 3) & " " & TsName
    excel_sheet.Cells(i, 5) = rs!idfarmer
     excel_sheet.Cells(i, 6) = rs!farmername
      excel_sheet.Cells(i, 7) = rs!phone1
   
   
   

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
    .PageSetup.CenterFooter = "CG LISTING"
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
'ERR:
'MsgBox ERR.Description

End Sub
Private Sub PRINTFINFO()

'On Error Resume Next
Dim excel_app As Object
Dim excel_sheet As Object
Dim row As Long
Dim statement As String
Dim i, j, K As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Screen.MousePointer = vbHourglass

FileCopy App.Path + "\FARMERINFO.XLS", App.Path + "\" + cbofarmerid.Text + Format(Now, "ddMMyyyy") + ".XLS"
Set excel_app = CreateObject("Excel.Application")
excel_app.Workbooks.Open FileName:=App.Path + "\" + cbofarmerid.Text + Format(Now, "ddMMyyyy") + ".XLS"
If Val(excel_app.Application.Version) >= 8 Then
   Set excel_sheet = excel_app.ActiveSheet
Else
   Set excel_sheet = excel_app
End If
excel_app.Visible = True
Set rs = Nothing
rs.Open "SELECT * FROM tblfarmer WHERE IDFARMER='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindDZ Mid(rs!idfarmer, 1, 3)
excel_sheet.Cells(5, 2) = Mid(rs!idfarmer, 1, 3) & " " & Dzname 'cboDzongkhag.Text
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
excel_sheet.Cells(5, 6) = Mid(rs!idfarmer, 1, 3) & Mid(rs!idfarmer, 4, 3) & " " & GEname
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
excel_sheet.Cells(6, 6) = Mid(rs!idfarmer, 1, 3) & Mid(rs!idfarmer, 4, 3) & Mid(rs!idfarmer, 7, 3) & " " & TsName
excel_sheet.Cells(7, 2) = rs!idfarmer & " " & rs!farmername
'excel_sheet.Cells(7, 6) = CBOCARETAKER.Text
excel_sheet.Cells(9, 2) = "'" & rs!cidno
If rs!sex = 0 Then
excel_sheet.Cells(9, 6) = "MALE" '
Else
excel_sheet.Cells(9, 6) = "FEMALE"
End If

excel_sheet.Cells(10, 2) = rs!houseno
excel_sheet.Cells(11, 2) = rs!VILLAGE
excel_sheet.Cells(11, 6) = rs!LocationName

excel_sheet.Cells(12, 2) = rs!phone1
excel_sheet.Cells(12, 6) = rs!phone2
If ISfarmercg = 0 Then
excel_sheet.Cells(13, 2) = "NO"
Else
excel_sheet.Cells(13, 2) = "YES"
End If
If rs!ISCARETAKER = 0 Then
excel_sheet.Cells(13, 6) = "NO"
Else
excel_sheet.Cells(13, 6) = "YES"
End If
If rs!ISCONTRACTSIGNED = 1 Then
excel_sheet.Cells(14, 2) = "YES"
If Format(rs!CONTRACTDATE, "dd/MM/yyyy") = "01/01/1900" Then
excel_sheet.Cells(14, 6) = ""
Else
excel_sheet.Cells(14, 6) = "'" & Format(rs!CONTRACTDATE, "dd/MM/yyyy") & "  " & "(DD/MM/YYYY)"
End If
Else
excel_sheet.Cells(14, 2) = "NO"
excel_sheet.Cells(14, 6) = ""

End If
excel_sheet.Cells(15, 2) = "'" & Format(IIf(IsNull(rs!TOTALAREA), 0, rs!TOTALAREA), "#####0.00")
Set rs1 = Nothing
rs1.Open "SELECT SUM(REGLAND)AS REGLAND FROM tbllandreg GROUP BY FARMERID", MHVDB
excel_sheet.Cells(15, 6) = "'" & Format(IIf(IsNull(rs!REGAREA), 0, rs!REGAREA) + IIf(IsNull(rs1!regland), 0, rs1!regland), "#####0.00")
excel_sheet.Cells(18, 1) = rs!remarks
Else
MsgBox "Record Not Found."
End If


Set rs = Nothing


'With excel_app.ActiveSheet.Pictures.Insert(App.Path + "\image\" + cboabsenteeid.BoundText & ".jpg")
'    With .ShapeRange
'        .LockAspectRatio = msoTrue
'        .Width = 60
'        .Height = 60
'    End With
'    .Left = excel_app.ActiveSheet.Cells(1, 8).Left
'    .Top = excel_app.ActiveSheet.Cells(1, 8).Top
'    .Placement = 1
'    .PrintObject = True
'End With



Screen.MousePointer = vbDefault
End Sub
Private Sub FL()
Dim rs As New ADODB.Recordset
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
    excel_sheet.Cells(3, 1) = ProperCase("SL.NO.")
    excel_sheet.Cells(3, 2) = ProperCase("dzongkhag code")
    excel_sheet.Cells(3, 3) = ProperCase("dzongkhag name")
    excel_sheet.Cells(3, 4) = ProperCase("gewog code")
    excel_sheet.Cells(3, 5) = ProperCase("gewog name")
    excel_sheet.Cells(3, 6) = ProperCase("tshowog code")
    excel_sheet.Cells(3, 7) = ProperCase("tshowog name")
    excel_sheet.Cells(3, 8) = ProperCase("FARMER ID")
    excel_sheet.Cells(3, 9) = ProperCase("FARMER NAME")
    excel_sheet.Cells(3, 10) = ProperCase("CID#")
    excel_sheet.Cells(3, 11) = ProperCase("contact No.")
    excel_sheet.Cells(3, 12) = ProperCase("VILLAGE")
    excel_sheet.Cells(3, 13) = ProperCase("LOCATION NAME")
    i = 4
  Set rs = Nothing
  rs.Open "SELECT * FROM tblfarmer where status='A' ORDER BY IDFARMER", MHVDB
   Do While rs.EOF <> True
   FindDZ Mid(rs!idfarmer, 1, 3)
   FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
   FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
   
        excel_sheet.Cells(i, 1) = sl
        excel_sheet.Cells(i, 2) = Mid(rs!idfarmer, 1, 3)
        excel_sheet.Cells(i, 3) = Dzname
        excel_sheet.Cells(i, 4) = Mid(rs!idfarmer, 4, 3)
        excel_sheet.Cells(i, 5) = GEname
        excel_sheet.Cells(i, 6) = Mid(rs!idfarmer, 7, 3)
        excel_sheet.Cells(i, 7) = TsName
        excel_sheet.Cells(i, 8) = rs!idfarmer
        excel_sheet.Cells(i, 9) = rs!farmername
        excel_sheet.Cells(i, 10) = rs!cidno
        excel_sheet.Cells(i, 11) = rs!phone1
        excel_sheet.Cells(i, 12) = rs!VILLAGE
        excel_sheet.Cells(i, 13) = rs!LocationName
 
 
 
 
 
 
   i = i + 1
   sl = sl + 1
   rs.MoveNext
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
     excel_sheet.Range("A3:n3").Font.Bold = True
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
Private Sub dl()
Dim rs As New ADODB.Recordset
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
    excel_sheet.Cells(3, 2) = "DZONGKHAG ID"
    excel_sheet.Cells(3, 3) = "DZONGKHAG NAME"
    excel_sheet.Cells(3, 4) = "REMARKS"
  
    i = 4
  Set rs = Nothing
  rs.Open "SELECT * FROM tbldzongkhag ORDER BY dzongkhagcode", MHVDB
   Do While rs.EOF <> True
   excel_sheet.Cells(i, 1) = sl
     excel_sheet.Cells(i, 2) = rs!DZONGKHAGCODE
   excel_sheet.Cells(i, 3) = rs!DZONGKHAGNAME
 excel_sheet.Cells(i, 4) = rs!remarks

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
    .PageSetup.CenterFooter = "DZONGKHAG LISTING"
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
Private Sub gl()
Dim rs As New ADODB.Recordset
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
    Dim dzcode As String
    dzcode = ""
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG"
    excel_sheet.Cells(3, 3) = "GEWOG ID"
    excel_sheet.Cells(3, 4) = "GEWOG NAME"
  excel_sheet.Cells(3, 5) = "REMARKS"
    i = 4
  Set rs = Nothing
  rs.Open "SELECT * FROM tblgewog ORDER BY DZONGKHAGID,gewogid", MHVDB
  
  
   Do While rs.EOF <> True
   
   If dzcode <> rs!dzongkhagid Then
   FindDZ rs!dzongkhagid
   
     excel_sheet.Cells(i, 2) = rs!dzongkhagid & "   " & UCase(Dzname)
     excel_sheet.Cells(i, 2).Font.Bold = True
       i = i + 1
          excel_sheet.Cells(i, 1) = sl
          sl = sl + 1
       excel_sheet.Cells(i, 3) = rs!gewogid
        excel_sheet.Cells(i, 4) = rs!gewogname
            excel_sheet.Cells(i, 5) = rs!remarks
     Else
     excel_sheet.Cells(i, 1) = sl
   excel_sheet.Cells(i, 3) = rs!gewogid
 excel_sheet.Cells(i, 4) = rs!gewogname
excel_sheet.Cells(i, 5) = rs!remarks
sl = sl + 1
End If
dzcode = rs!dzongkhagid
   i = i + 1
   
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
    excel_sheet.Range("A3:e3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "GEWOG LISTING"
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
Private Sub tl()
Dim rs As New ADODB.Recordset
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
    Dim dzcode As String
    Dim gcode As String
    gcode = ""
    dzcode = ""
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "DZONGKHAG/GEWOG"
    excel_sheet.Cells(3, 3) = "TSHOWOG ID"
    excel_sheet.Cells(3, 4) = "TSHOWOG NAME"
     excel_sheet.Cells(3, 5) = "REMARKS"

    i = 4
  Set rs = Nothing
  rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  
  
   Do While rs.EOF <> True
   
   If dzcode <> rs!dzongkhagid Then
   FindDZ rs!dzongkhagid
   
     excel_sheet.Cells(i, 2) = rs!dzongkhagid & "   " & UCase(Dzname)
     excel_sheet.Cells(i, 2).Font.Bold = True
       i = i + 1
   End If
   
   If gcode <> rs!gewogid Then
   FindGE rs!dzongkhagid, rs!gewogid
   
     excel_sheet.Cells(i, 2) = "     " & rs!gewogid & "   " & UCase(GEname)
     excel_sheet.Cells(i, 2).Font.Bold = True
       i = i + 1
   End If
     excel_sheet.Cells(i, 1) = sl
   excel_sheet.Cells(i, 3) = rs!tshewogid
 excel_sheet.Cells(i, 4) = rs!tshewogname
excel_sheet.Cells(i, 5) = rs!remarks
sl = sl + 1

dzcode = rs!dzongkhagid
gcode = rs!gewogid
   i = i + 1
   
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
    .PageSetup.CenterFooter = "TSHOWOG LISTING"
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

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Label1.Caption = "IF THE PAGE IS OUT OF RANGE SET THE PAGE TO LANDSCAPE...."
End Sub

Private Sub Command4_Click()



End Sub

Private Sub Form_Load()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing


If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Exit Sub
err:
MsgBox err.Description

cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Label1_Click()
Label1.Caption = ""
End Sub

Private Sub OPTBYID_Click()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing


cbofarmerid.Text = ""


If OPTNAME.Value = True Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(farmername  , ' ',idfarmer) as farmername,idfarmer  from tblfarmer order by FARMERNAME", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Else
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
End If
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub OPTFARMERLISTING_Click()
RptOption = ""
RptOption = "FL"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option1_Click()
RptOption = ""
RptOption = "AB"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option10_Click()
RptOption = ""
RptOption = "GRF"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option11_Click()
RptOption = ""
RptOption = "FH"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option12_Click()
RptOption = ""
RptOption = "ML"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option13_Click()
RptOption = ""
RptOption = "PL"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option14_Click()
RptOption = ""
RptOption = "FHC"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option2_Click()
RptOption = ""
RptOption = "DZ"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option3_Click()
RptOption = ""
RptOption = "GE"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option4_Click()
RptOption = ""
RptOption = "TS"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub
Private Sub FindDZ(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dzname = ""
Set rs = Nothing
rs.Open "select * from tbldzongkhag where dzongkhagcode='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Dzname = rs!DZONGKHAGNAME
Else
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
'Private Sub FindGE(dd As String, GG As String)
'On Error GoTo err
'Dim RS As New ADODB.Recordset
'GEname = ""
'Set RS = Nothing
'RS.Open "select * from tblgewog where dzongkhagID='" & dd & "' AND GEWOGID='" & GG & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
'If RS.EOF <> True Then
'GEname = RS!gewogname
'Else
'MsgBox "Record Not Found."
'End If
'Exit Sub
'err:
'MsgBox err.Description
'End Sub

Private Sub Option5_Click()
RptOption = ""
RptOption = "IFR"
cbofarmerid.Visible = True
Frame6.Visible = True
End Sub

Private Sub Option6_Click()
RptOption = ""
RptOption = "CG"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option7_Click()
RptOption = ""
RptOption = "Z"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option8_Click()
RptOption = ""
RptOption = "FLN"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub Option9_Click()
RptOption = ""
RptOption = "FDO"
cbofarmerid.Visible = False
Frame6.Visible = False
End Sub

Private Sub OPTNAME_Click()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing

cbofarmerid.Text = ""



If OPTNAME.Value = True Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(farmername  , ' ',idfarmer) as farmername,idfarmer  from tblfarmer order by FARMERNAME", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Else
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub fhc()
Dim rs As New ADODB.Recordset
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
    Dim dzcode As String
    Dim gcode As String
    gcode = ""
    dzcode = ""
    sl = 1
    'excel_app.DisplayFullScreen = True
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "SL.NO."
    excel_sheet.Cells(3, 2) = "name"
    excel_sheet.Cells(3, 3) = "label"
    excel_sheet.Cells(3, 4) = "Tshowog code"
    i = 4
  Set rs = Nothing
   ' rs.Open "SELECT DzongkhagCode,Dzongkhagname FROM tbldzongkhag ORDER BY DzongkhagCode", MHVDB
   'rs.Open "SELECT * FROM tblgewog ORDER BY dzongkhagid,gewogid", MHVDB
  rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  ' rs.Open "SELECT * FROM tbltshewog ORDER BY DZONGKHAGID,gewogid,tshewogid", MHVDB
  
   Do While rs.EOF <> True
   

   FindDZ rs!dzongkhagid
   excel_sheet.Cells(i, 2) = rs!dzongkhagid & rs!gewogid & rs!tshewogid
   FindTs rs!dzongkhagid, rs!gewogid, rs!tshewogid
   excel_sheet.Cells(i, 3) = rs!tshewogid & "   " & TsName
   excel_sheet.Cells(i, 4) = rs!dzongkhagid & rs!gewogid
   i = i + 1
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
    .PageSetup.CenterFooter = "TSHOWOG LISTING"
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

