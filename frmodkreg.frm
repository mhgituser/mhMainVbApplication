VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmodkreg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODK REGISTRATION"
   ClientHeight    =   5565
   ClientLeft      =   9105
   ClientTop       =   1845
   ClientWidth     =   5415
   Icon            =   "frmodkreg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5415
   Begin VB.Frame Frame2 
      Caption         =   "REPORT TYPE"
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   5175
      Begin VB.OptionButton OPTADVNEW 
         Caption         =   "ADVOCATE-MONITOR(NEW)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   2895
      End
      Begin VB.OptionButton OPTADVOLD 
         Caption         =   "ADVOCATE-MONITOR(OLD)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   3255
      End
      Begin VB.OptionButton OPTREGDET 
         Caption         =   "REGISTRATION(DETAIL)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton OPTADV 
         Caption         =   "ADVOCATE-MONITOR(ALL)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   2895
      End
      Begin VB.OptionButton OPTREGSUM 
         Caption         =   "REGISTRATION(SUMARY)"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.OptionButton OPTALL 
      Caption         =   "ALL"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton OPTSEL 
      Caption         =   "SELECTIVE"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   5295
      Begin VB.ComboBox CBODATE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmodkreg.frx":0E42
         Left            =   1080
         List            =   "frmodkreg.frx":0E44
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   110297089
         CurrentDate     =   41362
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   110297089
         CurrentDate     =   41362
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DATE TYPE"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "FROM DATE"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TO DATE"
         Height          =   195
         Left            =   2760
         TabIndex        =   6
         Top             =   1200
         Width           =   705
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SHOW"
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
      Left            =   600
      Picture         =   "frmodkreg.frx":0E46
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1575
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
      Left            =   2280
      Picture         =   "frmodkreg.frx":15B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
   End
End
Attribute VB_Name = "frmodkreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mindex As Integer
Private Sub Option1_Click()



End Sub
Private Sub REGSUM()
Dim SLNO As Integer
Dim totreg, totadd As Double
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim actstring As String
Dim tt As String
totreg = 0
totadd = 0
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                    
If OPTSEL.Value = True And Len(CBODATE.Text) = 0 Then
MsgBox "Please Select The Date Type."
Exit Sub
End If


Dim SQLSTR As String

SQLSTR = ""
SLNO = 1
If optall.Value = True Then
SQLSTR = "select staffbarcode, sum(regarea) as regarea,sum(AADDITIONAL_ACRE) as addland FROM farmer_registration4_core group by staffbarcode"
ElseIf OPTSEL.Value = True Then
SQLSTR = "select staffbarcode, sum(regarea) as regarea,sum(AADDITIONAL_ACRE) as addland FROM farmer_registration4_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' group by staffbarcode  "
Else
MsgBox "INVALIDE SELECTION OF OPTION"
End If


On Error Resume Next





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
    excel_app.Visible = False
    ' excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
   
    excel_sheet.cells(3, 2) = "STAFF CODE"
    
    excel_sheet.cells(3, 3) = "STAFF NAME"
   
    excel_sheet.cells(3, 4) = "ACRE REGISTERED"
 
   excel_sheet.cells(3, 5) = "ADDITIONAL LAND"
   i = 4
  Set rs = Nothing
rs.Open SQLSTR, db
  Do While rs.EOF <> True
 ' tt = "rs!" & CBODATE.Text
excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = rs!staffbarcode
FindsTAFF excel_sheet.cells(i, 2)
excel_sheet.cells(i, 3) = sTAFF
If IIf(IsNull(rs!REGAREA), 0, rs!REGAREA) <> 0 Then
excel_sheet.cells(i, 4) = IIf(IsNull(rs!REGAREA), 0, rs!REGAREA)
Else
excel_sheet.cells(i, 4) = ""
End If
If IIf(IsNull(rs!addland), 0, rs!addland) <> 0 Then
excel_sheet.cells(i, 5) = IIf(IsNull(rs!addland), 0, rs!addland)
Else
excel_sheet.cells(i, 5) = ""
End If
totreg = totreg + IIf(IsNull(rs!REGAREA), 0, rs!REGAREA)
totadd = totadd + IIf(IsNull(rs!addland), 0, rs!addland)

SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop
   excel_sheet.cells(i, 3).Font.Bold = True
   excel_sheet.cells(i, 4).Font.Bold = True
   excel_sheet.cells(i, 5).Font.Bold = True
    excel_sheet.cells(i, 3) = "TOTAL"
excel_sheet.cells(i, 4) = totreg
excel_sheet.cells(i, 5) = totadd

    
    
   'make up


'xlTmp.ActiveSheet.Columns("A:B").NumberFormat = "000000"

   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 15)).Select
    excel_app.selection.Columns("d:e").NumberFormat = "####0.00"
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:e3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ODK REGISTRATION (SUMMARY)"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
  



' excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(3, 15)).Select
excel_sheet.Columns("A").Select
 excel_app.selection.columnWidth = 7
 excel_sheet.Columns("B").Select
 excel_app.selection.columnWidth = 11
 
  excel_sheet.Columns("C").Select
 excel_app.selection.columnWidth = 20
 
 
  excel_sheet.Columns("D:E").Select
 excel_app.selection.columnWidth = 17

 
 
 
With excel_app.selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
db.Close
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear

End Sub

Private Sub ADVOCATE()
'On Error Resume Next
Dim mgender As String
Dim s As Integer
Dim SQLSTR As String
Dim totplant As Integer
Dim myphone As String
Dim TOTLAND As Double
Dim mregland As Double
Dim mdname, mgname, mtname As String
Dim totadd As Double
Dim i, j As Integer
Dim SLNO As Integer
Dim dcode, gcode, tcode, fcode As String
Dim rs As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Set rsadd = Nothing
Dim mfname As String
mchk = True

mfname = ""
mdname = 0
mgname = 0
mtname = 0

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                       
db.Open OdkCnnString
                        


dcode = ""
gcode = ""
tcode = ""
fcode = ""
mregland = 0


If optall.Value = True Then
If OPTADV.Value = True Then
SQLSTR = "select * FROM farmer_registration3_core order by regdate"
ElseIf OPTADVOLD.Value = True Then
SQLSTR = "select * FROM farmer_registration3_core where newold='no' order by regdate"
Else

SQLSTR = "select * FROM farmer_registration3_core where newold='yes' order by regdate"

End If

ElseIf OPTSEL.Value = True Then
If OPTADV.Value = True Then
SQLSTR = "select * FROM farmer_registration3_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' ORDER BY " & CBODATE.Text & "  "
ElseIf OPTADVOLD.Value = True Then
SQLSTR = "select * FROM farmer_registration3_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and newold='no' ORDER BY " & CBODATE.Text & "  "

Else

SQLSTR = "select * FROM farmer_registration3_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and newold='yes' ORDER BY " & CBODATE.Text & "  "

End If
Else
MsgBox "INVALIDE SELECTION OF OPTION"
End If



                            
                        
                        
                        
                        

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
    'excel_app.Visible = True
    excel_sheet.cells(3, 1) = "Sl. No"
    excel_sheet.cells(3, 2) = "Date Reg"
    excel_sheet.cells(3, 3) = "Code"
    excel_sheet.cells(3, 4) = " Name"
     excel_sheet.cells(3, 5) = "Code"
    excel_sheet.cells(3, 6) = "Name"
    excel_sheet.cells(3, 7) = "Code"
    excel_sheet.cells(3, 8) = "Name"
    excel_sheet.cells(3, 9) = "Farmer Code"
    excel_sheet.cells(3, 10) = "Name of interested Grower"
    excel_sheet.cells(3, 11) = "Gender of Interested Grower(f/m)"
    excel_sheet.cells(3, 12) = "Citizenship ID"
    excel_sheet.cells(3, 13) = "Mobile No"
    excel_sheet.cells(3, 14) = "House#"
    excel_sheet.cells(3, 15) = "Is The Farmer New or Old."
    excel_sheet.cells(3, 16) = "Serial #"
    excel_sheet.cells(3, 17) = "Village"
    excel_sheet.cells(3, 18) = "Local Location name"
    excel_sheet.cells(3, 19) = "Total Area Of Fallow Land (Acres)"
    excel_sheet.cells(3, 20) = "Thram #"
    excel_sheet.cells(3, 21) = "Gender Of Thram Holder"
    excel_sheet.cells(3, 22) = "Name of Thram Holder"
    excel_sheet.cells(3, 23) = "Relationship with Thram Holder"
    excel_sheet.cells(3, 24) = "Land Use Type"
    excel_sheet.cells(3, 25) = "Leased Land Type"
    excel_sheet.cells(3, 26) = "Area Registered (Acres)"
    excel_sheet.cells(3, 27) = "ID of Person Registering either MHV ID or CG Full ID"
    excel_sheet.cells(3, 28) = "Name of Person Registering the Farmer"
    i = 4
                 
                            Set rs = Nothing
                            rs.Open SQLSTR, db
                            If rs.EOF <> True Then
                            
                            mregland = 0
                            
        
                            
                            Do While rs.EOF <> True
                            mfname = ""
                            dcode = ""
                            gcode = ""
                            tcode = ""
                            chkred = False
                            If rs!newold = "no" Then
                             mregland = rs!regarea1
                             mgender = rs!gender1
'                             If RS!fcode = 225 Then
'                             MsgBox "asd"
'                             End If
                            mdname = rs!dname1
                            mgname = rs!gname1
                            mtname = rs!tname1
                             mfname = rs!fname1
                            If Len(rs!dcode1) = 1 Then
                            dcode = "D0" & rs!dcode1
                            Else
                            dcode = "D" & rs!dcode1
                            End If
                            '
                            
                            If Len(rs!gcode1) = 1 Then
                            gcode = "G0" & rs!gcode1
                            Else
                            gcode = "G" & rs!gcode1
                            End If
                           ' FindGE dcode, gcode
                            
                            If Len(rs!tcode1) = 1 Then
                            tcode = "T0" & rs!tcode1
                            Else
                            tcode = "T" & rs!tcode1
                            End If
                           ' FindTs dcode, gcode, tcode
                           
                           
                          
                            If Len(rs!fcode) = 1 Then
                            fcode = dcode & gcode & tcode & "F000" & rs!fcode
                            ElseIf Len(rs!fcode) = 2 Then
                            fcode = dcode & gcode & tcode & "F00" & rs!fcode
                            ElseIf Len(rs!fcode) = 3 Then
                            fcode = dcode & gcode & tcode & "F0" & rs!fcode
                            Else
                            fcode = dcode & gcode & tcode & "F" & rs!fcode
                            End If
                            
                            'FindFA fcode
                            
                           Else
                        
                        
                         mregland = rs!REGAREA
                        mgender = rs!gender
                        mfname = rs!fname
                           mdname = rs!dname
                            mgname = rs!gname
                            mtname = rs!tname
                        If Len(rs!dcode) = 1 Then
                            dcode = "D0" & rs!dcode
                            Else
                            dcode = "D" & rs!dcode
                            End If
                            'FindDZ (dcode)
                            
                            If Len(rs!gcode) = 1 Then
                            gcode = "G0" & rs!gcode
                            Else
                            gcode = "G" & rs!gcode
                            End If
                           ' FindGE dcode, gcode
                            
                            If Len(rs!tcode) = 1 Then
                            tcode = "T0" & rs!tcode
                            Else
                            tcode = "T" & rs!tcode
                            End If
                           ' FindTs dcode, gcode, tcode
                           
                           
                           
                            fcode = ""
                       
                        
                        
                        
                        End If
                            
                            
                          
                            
                            
                            
                                  
                                excel_sheet.cells(i, 1) = sl
                                excel_sheet.cells(i, 2) = "'" & rs!regdate
                                
                                
                                If Mid(dcode, 2, 3) = 0 Then
                                 excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbRed
                                End If
                                
                                excel_sheet.cells(i, 3) = dcode
                                FindDZ excel_sheet.cells(i, 3)
                                If Len(Dzname) <> 0 Then
                                excel_sheet.cells(i, 4) = Dzname
                                Else
                                excel_sheet.cells(i, 4) = mdname
                                
                                End If
                                
                                
                                
                                
                                If Mid(gcode, 2, 3) = 0 Then
                                 excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbRed
                                End If
                                excel_sheet.cells(i, 5) = gcode
                                
                                
                                FindGE excel_sheet.cells(i, 3), excel_sheet.cells(i, 5)
                                If Len(GEname) <> 0 Then
                                excel_sheet.cells(i, 6) = GEname
                                Else
                                excel_sheet.cells(i, 6) = mgname
                                
                                End If
                                If Mid(tcode, 2, 3) = 0 Then
                                 excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbRed
                                End If
                                
                                excel_sheet.cells(i, 7) = tcode
                                FindTs excel_sheet.cells(i, 3), excel_sheet.cells(i, 5), excel_sheet.cells(i, 7)
                                If Len(TsName) <> 0 Then
                                excel_sheet.cells(i, 8) = TsName
                                Else
                                excel_sheet.cells(i, 8) = mtname
                                End If
                                 If Len(fcode) <> 14 And Not fcode = "" Then
                                 excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbRed
                                End If
                                
                                excel_sheet.cells(i, 9) = fcode
                                If rs!newold = "no" Then
                                FindFA excel_sheet.cells(i, 9), "F"
'                                If fcode = "D04G08T02F0000" Then
'                                MsgBox "sdfs"
'                                End If
                                If chkred = True And Mid(fcode, 11, 14) <> 0 And Len(fcode) = 14 Then
                                
                               excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbBlue
                                End If
                                
                                 If chkred = True And Mid(fcode, 11, 14) = 0 And Len(fcode) = 14 And Mid(dcode, 2, 3) <> 0 And Mid(gcode, 2, 3) <> 0 And Mid(tcode, 2, 3) <> 0 Then
                                
                            excel_sheet.Range(excel_sheet.cells(i, 1), _
                             excel_sheet.cells(i, 28)).Select
                             excel_app.selection.Font.Color = vbRed
                                End If
                                
                                
                                
                                chkred = False
                                If Len(FAName) <> 0 Then
                                excel_sheet.cells(i, 10) = FAName
                                Else
                                
                                excel_sheet.cells(i, 10) = mfname
                                End If
                                Else
                                 excel_sheet.cells(i, 10) = mfname
                                End If
                                
                                
                                
                                excel_sheet.cells(i, 11) = rs!gender ' "Gender of Interested Grower(f/m)"
                                excel_sheet.cells(i, 12) = rs!fid ' "Citizenship ID"
                                excel_sheet.cells(i, 13) = rs!mnumber '"Mobile No"
                                excel_sheet.cells(i, 14) = rs!house '"House#"
                                If rs!newold = "no" Then
                                excel_sheet.cells(i, 15) = "OLD"
                                Else
                        
                                excel_sheet.cells(i, 15) = "NEW"
                                End If
                                excel_sheet.cells(i, 16) = sl
                                excel_sheet.cells(i, 17) = rs!VILLAGE '"Village"
                                excel_sheet.cells(i, 18) = rs!lname '"Local Location name"
                                excel_sheet.cells(i, 19) = rs!farea '"Total Area Of Fallow Land (Acres)"
                                excel_sheet.cells(i, 20) = rs!thram ' "Thram #"
                                excel_sheet.cells(i, 21) = rs!gender1 '"Gender Of Thram Holder"
                                excel_sheet.cells(i, 22) = rs!name '"Name of Thram Holder"
                                excel_sheet.cells(i, 23) = "" 'RS!Name '"Relationship with Thram Holder"
                                excel_sheet.cells(i, 24) = rs!LANDTYPE ' "Land Use Type"
'                                excel_sheet.Cells(i, 25) = "Leased Land Type"
                                excel_sheet.cells(i, 26) = mregland '"Area Registered (Acres)"
                                excel_sheet.cells(i, 27) = rs!staffcode '"ID of Person Registering either MHV ID or CG Full ID"
                                FindsTAFF "S0" & excel_sheet.cells(i, 27)
                                excel_sheet.cells(i, 28) = sTAFF
                             
                             
                             
                            
                             i = i + 1
                             sl = sl + 1
                                           rs.MoveNext
                                           Loop
                        
                            End If
                      





                            
                            'make up
                            excel_sheet.Range(excel_sheet.cells(4, 8), _
                            excel_sheet.cells(i, 8)).Select
                            ' excel_app.Selection.Columns("H:H").NumberFormat = "####0.00"
                            excel_app.selection.NumberFormat = "####0.00"
                   excel_sheet.Range(excel_sheet.cells(2, 3), _
                             excel_sheet.cells(2, 4)).Select



                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 90
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
'
'
                            excel_sheet.cells(2, 3) = "Dzongkhag"
                            
                            
                            excel_sheet.Range(excel_sheet.cells(2, 5), _
                             excel_sheet.cells(2, 6)).Select



                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 90
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
'
'
                            excel_sheet.cells(2, 5) = "Gewog"
                            
                              excel_sheet.Range(excel_sheet.cells(2, 7), _
                             excel_sheet.cells(2, 8)).Select
                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 90
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            excel_sheet.cells(2, 7) = "Tshowog"
                            
 excel_sheet.Range(excel_sheet.cells(1, 1), _
                             excel_sheet.cells(i, 28)).Select
'excel_sheet.Columns("A:A").Select
 excel_app.selection.Font.Size = 10

 
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
                            
                         
                            
                            
     excel_sheet.Range(excel_sheet.cells(1, 1), _
                             excel_sheet.cells(i, 26)).Select
                            
                   excel_app.selection.Font.name = "Times New Roman"
   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 6)).Select
    excel_app.selection.Columns.AutoFit
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A1:ab3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ADVOCATE-MONITOR FORM"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With
    
    
    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
'    Set excel_sheet = Nothing
'    Set excel_app = Nothing

   

excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault

End Sub

Private Sub CBODATE_LostFocus()
Dim i, j, fcount As Integer

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                      
db.Open OdkCnnString
                        
Set rs = Nothing
rs.Open "select * from tbltable where tblid='8' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)

Set rs = Nothing
rs.Open "SELECT * FROM farmer_registration3_core where 1", CONNLOCAL
For j = 0 To fcount - 1
If rs.Fields(j).Type = 135 Then

If rs.Fields(j).name = CBODATE.Text Then
Mindex = j
Exit For
Else
Mindex = 2
End If


End If
Next

Exit Sub
err:
MsgBox err.Description
End Sub



Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If OPTSEL.Value = True And Len(CBODATE.Text) = 0 Then
MsgBox "Please Select The Date Type."
Exit Sub
End If
If OPTREGSUM.Value = True Then
REGSUM
ElseIf OPTREGDET.Value = True Then
REGDET
Else
ADVOCATE
End If
End Sub
Private Sub REGDET()
Dim SLNO As Integer
Dim totreg, totadd As Double
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim actstring As String
Dim tt As String
Dim dcode, gcode, tcode, fcode As String
dcode = ""
gcode = ""
tcode = ""
fcode = ""
totreg = 0
totadd = 0
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                       
If OPTSEL.Value = True And Len(CBODATE.Text) = 0 Then
MsgBox "Please Select The Date Type."
Exit Sub
End If


Dim SQLSTR As String

SQLSTR = ""
SLNO = 1
If optall.Value = True Then
SQLSTR = "select * FROM farmer_registration3_core order by regdate"
ElseIf OPTSEL.Value = True Then
SQLSTR = "select * FROM farmer_registration3_core where SUBSTRING(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and SUBSTRING(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' ORDER BY " & CBODATE.Text & "  "
Else
MsgBox "INVALIDE SELECTION OF OPTION"
End If


On Error Resume Next


If optall.Value = True Then
Mindex = 30
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
    excel_app.Visible = False
    ' excel_app.Visible = True
    excel_sheet.cells(3, 1) = "SL.NO."
   If optall.Value = True Then
    excel_sheet.cells(3, 2) = "DATE" & "(REGDATE)"
    Else
    excel_sheet.cells(3, 2) = "DATE" & "(" & CBODATE.Text & ")"
    End If
    excel_sheet.cells(3, 3) = "STAFF CODE"
    
    excel_sheet.cells(3, 4) = "STAFF NAME"
   
    excel_sheet.cells(3, 5) = "ACRE REGISTERED"
 
   excel_sheet.cells(3, 6) = "ADDITIONAL LAND"
   excel_sheet.cells(3, 7) = "REMARKS"
   i = 4
  Set rs = Nothing
rs.Open SQLSTR, db
  Do While rs.EOF <> True
 ' tt = "rs!" & CBODATE.Text
excel_sheet.cells(i, 1) = SLNO
excel_sheet.cells(i, 2) = "'" & rs.Fields(Mindex)
excel_sheet.cells(i, 3) = "S0" & rs!staffcode
FindsTAFF excel_sheet.cells(i, 3)
excel_sheet.cells(i, 4) = sTAFF
If IIf(IsNull(rs!REGAREA), 0, rs!REGAREA) <> 0 Then
excel_sheet.cells(i, 5) = IIf(IsNull(rs!REGAREA), 0, rs!REGAREA)
Else
excel_sheet.cells(i, 5) = ""
End If
If IIf(IsNull(rs!regarea2), 0, rs!regarea2) <> 0 Then
excel_sheet.cells(i, 6) = IIf(IsNull(rs!regarea2), 0, rs!regarea2)
Else
excel_sheet.cells(i, 6) = ""
End If




totreg = totreg + IIf(IsNull(rs!REGAREA), 0, rs!REGAREA)
totadd = totadd + IIf(IsNull(rs!regarea2), 0, rs!regarea2)
If rs!newold = "no" Then
 excel_sheet.cells(i, 7) = "OLD"
ElseIf rs!newold = "yes" Then
 excel_sheet.cells(i, 7) = "NEW"
Else
excel_sheet.cells(i, 7) = ""
End If
SLNO = SLNO + 1
i = i + 1
rs.MoveNext
   Loop
   excel_sheet.cells(i, 4).Font.Bold = True
   excel_sheet.cells(i, 5).Font.Bold = True
   excel_sheet.cells(i, 6).Font.Bold = True
    excel_sheet.cells(i, 4) = "TOTAL"
excel_sheet.cells(i, 5) = totreg
excel_sheet.cells(i, 6) = totadd

    
    
   'make up


'xlTmp.ActiveSheet.Columns("A:B").NumberFormat = "000000"

   excel_sheet.Range(excel_sheet.cells(3, 1), _
    excel_sheet.cells(i, 15)).Select
    excel_app.selection.Columns("E:F").NumberFormat = "####0.00"
   ' Freeze the header row so it doesn't scroll.
    excel_sheet.cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.cells(1, 1).Select
    With excel_sheet
    '.PageSetup.LeftHeader = "MHV"
     excel_sheet.Range("A3:G3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "ODK REGISTRATION (DETAIL)"
        .PageSetup.LeftFooter = "MHV"
        .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
        .PageSetup.PrintGridlines = True
    End With


    '
    ''excel_sheet.Cells(1, 1) = "MHV"
    'excel_sheet.Cells(2, 1) = IIf(dcbunits.BoundText <> "100", dcbunits, " - All Depots") + "(Figures In " + IIf(RepName = "3", "Nu.", IIf(RepName = "8", "Pcs", IIf(RepName = "11", "C/s", JUnit))) + ")" '*
  



' excel_sheet.Range(excel_sheet.Cells(3, 1), _
'    excel_sheet.Cells(3, 15)).Select
excel_sheet.Columns("A").Select
 excel_app.selection.columnWidth = 7
 excel_sheet.Columns("B").Select
 excel_app.selection.columnWidth = 15
 
  excel_sheet.Columns("C").Select
 excel_app.selection.columnWidth = 11
 
 
  excel_sheet.Columns("D").Select
 excel_app.selection.columnWidth = 20
  excel_sheet.Columns("E:F").Select
 excel_app.selection.columnWidth = 17
  excel_sheet.Columns("G").Select
 excel_app.selection.columnWidth = 9
 
 
 
 
 
 
 
 
 
 
With excel_app.selection
'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


With excel_sheet
    .PageSetup.Orientation = xlLandscape
    '.PrintOut
End With




Dim PB As Integer
With excel_sheet.PageSetup
        ' MsgBox CInt(ExecuteExcel4Macro("Get.Document(50)"))
         PB = CInt(ExecuteExcel4Macro("Get.Document(50)"))
          .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = PB

End With

' MsgBox ExecuteExcel4Macro("Get.Document(50)")
excel_app.Visible = True
Set excel_sheet = Nothing
Set excel_app = Nothing
Screen.MousePointer = vbDefault
db.Close
'Exit Sub
'ERR:
'MsgBox ERR.Description
'ERR.Clear
End Sub
Private Sub Form_Load()
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
Dim i, j, fcount As Integer
Operation = ""
Mindex = 0
'Mygrid.Visible = False
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                      
db.Open OdkCnnString
                        
Set rs = Nothing
rs.Open "select * from tbltable where tblid='8' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)
CBODATE.Clear
Set rs = Nothing
rs.Open "SELECT * FROM farmer_registration3_core where 1", CONNLOCAL
For j = 0 To fcount - 1
If rs.Fields(j).Type = 135 Then
CBODATE.AddItem rs.Fields(j).name
End If
Next

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
mchk = False
End Sub

Private Sub OPTALL_Click()
Frame1.Enabled = False
End Sub

Private Sub OPTREG_Click()

End Sub

Private Sub OPTSEL_Click()
Frame1.Enabled = True
End Sub
