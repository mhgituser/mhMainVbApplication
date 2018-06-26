VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMDISTLISTPRINT 
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtmonth 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   7080
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cbomnth 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FRMDISTLISTPRINT.frx":0000
         Left            =   6840
         List            =   "FRMDISTLISTPRINT.frx":0028
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboyear 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FRMDISTLISTPRINT.frx":0068
         Left            =   4920
         List            =   "FRMDISTLISTPRINT.frx":007B
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMDISTLISTPRINT.frx":009D
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Trn. Id"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Distribution Description"
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
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1995
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
      Left            =   2400
      Picture         =   "FRMDISTLISTPRINT.frx":00B2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
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
      Left            =   3960
      Picture         =   "FRMDISTLISTPRINT.frx":081C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "FRMDISTLISTPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbotrnid_LostFocus()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblplantdistributionheader where trnid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
cbomnth.Text = MonthName(rs!mnth)
cboyear.Text = rs!Year
txtdesc.Text = rs!distributionname
txtmonth.Text = ""
txtmonth.Text = rs!mnth
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
DISTLIST
End Sub
Private Sub DISTLIST()
'On Error Resume Next
Dim s As Integer
Dim SQLSTR As String
Dim totplant As Integer
Dim myphone As String
Dim TOTLAND As Double
Dim tcode As String
Dim totadd As Double
Dim msubtot, malltot As Double
TOTLAND = 0
Dim mm
totadd = 0
totplant = 0
Dzstr = ""
SQLSTR = ""
msubtot = 0
malltot = 0
Dim subsidizedamt As Double
Dim subsidizedamtsubtotal As Double
Dim subsidizedamtalltotal As Double
Dim totwatercan, tothosepipe, totagronet, toturea As Double
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset


Dim mtotalplant As Double
Dim macre As Double
Dim mcrate As Double
Dim anet, anett As Double


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
    excel_app.Visible = True
    excel_sheet.cells(3, 1) = "S/N"
    excel_sheet.cells(3, 2) = "DZONGKHAG"
    excel_sheet.cells(3, 3) = "GEWOG"
    excel_sheet.cells(3, 4) = "TSHOWOG"
     excel_sheet.cells(3, 5) = "FARMER CODE"
    excel_sheet.cells(3, 6) = "FARMER"
    excel_sheet.cells(3, 7) = "CONTACT #"
    excel_sheet.cells(3, 8) = "VILLAGE"
    excel_sheet.cells(3, 9) = "LAND (ACRE)"
    excel_sheet.cells(3, 10) = "TOTAL PLANT"
    excel_sheet.cells(3, 11) = UCase("Crates #")
    excel_sheet.cells(3, 12) = UCase("B (Crate)")
    excel_sheet.cells(3, 13) = UCase("E(Crate)")
    excel_sheet.cells(3, 14) = UCase("P (No)")
    excel_sheet.cells(3, 15) = UCase("P1(Nos)")
    excel_sheet.cells(3, 16) = UCase("N")
    excel_sheet.cells(3, 17) = UCase("SSP (Kg)")
    excel_sheet.cells(3, 18) = UCase("MOP(Kg)")
    excel_sheet.cells(3, 19) = UCase("Urea(Kg)")
    excel_sheet.cells(3, 20) = UCase("Dolomite(Kg)")
    excel_sheet.cells(3, 21) = UCase("Total (Kg)")
    excel_sheet.cells(3, 22) = UCase("Amount (Nu)")
    excel_sheet.cells(3, 23) = UCase("Kg")
    excel_sheet.cells(3, 24) = UCase("Amount (Nu)")
    excel_sheet.cells(3, 25) = UCase("Total Amount(Nu)")
    excel_sheet.cells(3, 26) = UCase("Schedule Date, Vehicle No & Team Captainy")
    excel_sheet.cells(3, 27) = UCase("Farmer Type")
    excel_sheet.cells(3, 28) = UCase("Monitor")
    excel_sheet.cells(3, 29) = UCase("Amount to be collected (30%) ")
    
    excel_sheet.cells(2, 31) = UCase("Incentive Materials")
    excel_sheet.cells(3, 30) = UCase("Water Can")
    excel_sheet.cells(3, 31) = UCase("Hose Pipe")
    excel_sheet.cells(3, 32) = UCase("Agro Net")
    excel_sheet.cells(3, 33) = UCase("Urea")
    excel_sheet.cells(3, 34) = UCase("Note")
    excel_sheet.cells(3, 35) = UCase("Prodcution")
    excel_sheet.cells(3, 36) = UCase("Pollinizer")
   
    
    i = 4
    s = 4
    SQLSTR = "select * from tblplantdistributiondetail where trnid='" & cbotrnid.BoundText & "' and mnth='" & Val(txtmonth.Text) & "' and year='" & cboyear.Text & "' order by sno"
    
                        tdist = 0
                        subsidizedamt = 0
                            Set rs = Nothing
                            rs.Open SQLSTR, MHVDB
                            
                            
                            Do While rs.EOF <> True
                            
                            'excel_sheet.Cells(i, 1) = rs!distno '"D/N"
                            FindDZ Mid(rs!farmercode, 1, 3)
                            FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
                            FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
                                                      
                             excel_sheet.cells(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
                             excel_sheet.cells(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
                             excel_sheet.cells(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
                              excel_sheet.cells(i, 5) = rs!farmercode ' "FARMER CODE"
                              
                              Set rs1 = Nothing
                              rs1.Open "select * from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
                              If rs1.EOF <> True Then
                             excel_sheet.cells(i, 6) = rs1!farmername
                             
                                 myphone = IIf(IsNull(rs1!phone1), "", rs1!phone1) & "," & IIf(IsNull(rs1!phone2), "", rs1!phone2)
                                 If Len(myphone) > 0 Then
                                 myphone = Left(myphone, Len(myphone) - 1)
                                 End If
                                 
                             excel_sheet.cells(i, 7) = myphone
                             excel_sheet.cells(i, 8) = rs1!VILLAGE
                             End If
                           
                             
                             excel_sheet.cells(i, 9) = rs!area '"LAND (ACRE)"
                             excel_sheet.cells(i, 10) = rs!totalplant '"TOTAL PLANT"
                             excel_sheet.cells(i, 11) = rs!crateno 'UCase("Crates #")
                             excel_sheet.cells(i, 12) = rs!bcrate ' UCase("B (Crate)")
                             excel_sheet.cells(i, 13) = rs!ecrate 'UCase("E(Crate)")
                             excel_sheet.cells(i, 14) = rs!bno 'UCase("P (No)")
                             excel_sheet.cells(i, 15) = rs!plno 'UCase("P1(Nos)")
                             excel_sheet.cells(i, 16) = rs!crate 'UCase("N")
                             excel_sheet.cells(i, 17) = Format(rs!ssp, "####.##") 'UCase("SSP (Kg)")
                             excel_sheet.cells(i, 18) = Format(rs!mop, "####.##") 'UCase("MOP(Kg)")
                             excel_sheet.cells(i, 19) = Format(rs!urea, "####.##") 'UCase("Urea(Kg)")
                             excel_sheet.cells(i, 20) = Format(rs!dolomite, "####.##") 'UCase("Dolomite(Kg)")
                             excel_sheet.cells(i, 21) = rs!totalkg1 'UCase("Total (Kg)")
                             excel_sheet.cells(i, 22) = rs!amountnu1 ' UCase("Amount (Nu)")
                             excel_sheet.cells(i, 23) = rs!kg 'UCase("Kg")
                             excel_sheet.cells(i, 24) = rs!amountnu2 'UCase("Amount (Nu)")
                             farmertype rs!farmercode
                             excel_sheet.cells(i, 27) = mFarmerType 'UCase("Amount (Nu)")
                               FindMONITOR rs!farmercode
                             excel_sheet.cells(i, 28) = myMonitor 'UCase("Amount (Nu)")
                             
                             'subsidizedamt = Round(rs!totalamount * 0.3, 0)
                             excel_sheet.cells(i, 25) = 5 * Round(rs!totalamount / 5, 0)
                             excel_sheet.cells(i, 29) = 5 * Round((0.3 * rs!totalamount) / 5, 0)
                             
                             mtotalplant = rs!totalplant
                             macre = rs!area
                             mcrate = rs!crateno
                             
                             
                             If mtotalplant >= 175 Then
                                excel_sheet.cells(i, 30) = 1 ' water cane
                                excel_sheet.cells(i, 31) = 1 'hose pipe
                             Else
                                excel_sheet.cells(i, 30) = 0
                                excel_sheet.cells(i, 31) = 0
                            End If
                            
'' added by kinzang: begin
                    
                             
                             If mtotalplant >= 175 And mcrate >= 5 Then
                                If macre >= 0.5 Then
                                Dim te As Integer
                                Dim aa As Integer
                                                                  
                                   excel_sheet.cells(i, 32) = CInt(Round((mtotalplant / 350) + 0.0000001))  'agronet
                                Else
                                   excel_sheet.cells(i, 32) = CInt(Round((mcrate / 5) + 0.0000001)) 'agronet
                                   
                                End If
                             Else
                                  
                                    excel_sheet.cells(i, 32) = 0 'agronet
                                   
                             End If
                            
                            
                            
                            If mtotalplant >= 175 And mcrate >= 5 Then
                                If macre >= 0.5 Then
                                   excel_sheet.cells(i, 33) = Val(excel_sheet.cells(i, 32)) * 200
                                Else
                                   excel_sheet.cells(i, 33) = Val(excel_sheet.cells(i, 32)) * 200
                                End If
                             Else
                                    excel_sheet.cells(i, 33) = 0
                            End If
                            excel_sheet.cells(i, 34) = rs!note
                            excel_sheet.cells(i, 35) = rs!production
                            excel_sheet.cells(i, 36) = rs!pollinizer
 '' end
                             
                             
                          
'                             If Right(rs!totalamount, 1) = 1 Then
'                             excel_sheet.Cells(i, 25) = rs!totalamount - 1
'                             ElseIf Right(rs!totalamount, 1) = 2 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount - 2
'                             ElseIf Right(rs!totalamount, 1) = 3 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount + 2
'                             ElseIf Right(rs!totalamount, 1) = 4 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount + 1
'                             ElseIf Right(rs!totalamount, 1) = 6 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount - 1
'                             ElseIf Right(rs!totalamount, 1) = 7 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount - 2
'
'                             ElseIf Right(rs!totalamount, 1) = 8 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount + 2
'                             ElseIf Right(rs!totalamount, 1) = 9 Then
'                              excel_sheet.Cells(i, 25) = rs!totalamount + 1
'                             Else
'                             excel_sheet.Cells(i, 25) = rs!totalamount
'                             End If
'
'                             If Right(subsidizedamt, 1) = 1 Then
'                             excel_sheet.Cells(i, 29) = subsidizedamt - 1
'                             ElseIf Right(subsidizedamt, 1) = 2 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt - 2
'                             ElseIf Right(subsidizedamt, 1) = 3 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt + 2
'                             ElseIf Right(subsidizedamt, 1) = 4 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt + 1
'                             ElseIf Right(subsidizedamt, 1) = 6 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt - 1
'                             ElseIf Right(subsidizedamt, 1) = 7 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt - 2
'
'                             ElseIf Right(subsidizedamt, 1) = 8 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt + 2
'                             ElseIf Right(subsidizedamt, 1) = 9 Then
'                              excel_sheet.Cells(i, 29) = subsidizedamt + 1
'                             Else
'                             excel_sheet.Cells(i, 29) = subsidizedamt
'                             End If
                             
                             'tdist = 0
                             If rs!subtotindicator = "" Then
                             msubtot = msubtot + excel_sheet.cells(i, 25)
                             malltot = malltot + excel_sheet.cells(i, 25)
                             subsidizedamtsubtotal = subsidizedamtsubtotal + excel_sheet.cells(i, 29)
                             'subsidizedamtalltotal = subsidizedamtalltotal + excel_sheet.cells(i, 29)
                             
                             totwatercan = totwatercan + excel_sheet.cells(i, 30)
                             tothosepipe = tothosepipe + excel_sheet.cells(i, 31)
                             totagronet = totagronet + excel_sheet.cells(i, 32)
                             toturea = toturea + excel_sheet.cells(i, 33)
                             
                             tdist = rs!distno
                             End If
                             
                            If rs!subtotindicator = "S" Then
                            excel_sheet.cells(i, 25) = msubtot
                            excel_sheet.cells(i, 29) = subsidizedamtsubtotal
                            excel_sheet.cells(i, 30) = totwatercan
                            excel_sheet.cells(i, 31) = tothosepipe
                            excel_sheet.cells(i, 32) = totagronet
                            excel_sheet.cells(i, 33) = toturea
                            
                            msubtot = 0
                            subsidizedamtsubtotal = 0
                            subsidizedamt = 0
                            'totwatercan = 0
                           ' tothosepipe = 0
                            'totagronet = 0
                           ' toturea = 0
                            
                            
                             excel_sheet.Range(excel_sheet.cells(i, 2), _
                             excel_sheet.cells(i, 26)).Select
                             excel_app.selection.Interior.ColorIndex = 15
                             
                                                       
                             excel_sheet.Range(excel_sheet.cells(s, 1), _
                             excel_sheet.cells(i - 1, 1)).Select
                             
                                excel_sheet.cells(s, 1) = tdist
                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            'Selection.Merge
                            
                            
                             excel_sheet.cells(i, 1) = ""
                            
                            
                                                      
                             excel_sheet.Range(excel_sheet.cells(s, 26), _
                             excel_sheet.cells(i - 1, 26)).Select
                             
                            
                            
                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            'Selection.Merge
                                                      
                            'excel_sheet.Cells(s, 1) = rs!Schedule
                                                  
                                                   
                             
                             
                             
                            
                            s = i + 1
                            End If
                            
                            If rs!subtotindicator = "T" Then
                            excel_sheet.cells(i, 25) = malltot
                            excel_sheet.cells(i, 29) = subsidizedamtalltotal
                            excel_sheet.cells(i, 30) = totwatercan
                            excel_sheet.cells(i, 31) = tothosepipe
                            excel_sheet.cells(i, 32) = totagronet
                            excel_sheet.cells(i, 33) = toturea
                            End If
                            
                            
                            
                            i = i + 1
                            
                            rs.MoveNext
                            Loop
                                
                             
                            
                            
                            'make up
                            excel_sheet.Range(excel_sheet.cells(4, 8), _
                            excel_sheet.cells(i, 8)).Select
                            excel_app.selection.NumberFormat = "####0.00"
                            excel_sheet.Range(excel_sheet.cells(1, 12), _
                            excel_sheet.cells(1, 16)).Select
                             
                            
                            
                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            
                            
                            excel_sheet.cells(1, 12) = "Variety"
                            
                            
                            
                            
                            excel_sheet.Range(excel_sheet.cells(1, 17), _
                            excel_sheet.cells(1, 22)).Select
                             
                            
                            
                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.cells(1, 17) = "Pre - Mixed Fertilizer"
                            
                            
                            
                            excel_sheet.Range(excel_sheet.cells(1, 23), _
                             excel_sheet.cells(1, 24)).Select
                             
                            
                            
                            With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                            
                            
                            excel_sheet.cells(1, 23) = "Dolomite"
                            
                            
            excel_sheet.Range(excel_sheet.cells(2, 16), _
                             excel_sheet.cells(2, 18)).Select
                             
                            
                           With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            
                            excel_sheet.cells(2, 16) = "Mixed Variety"
                            
                            excel_sheet.Range(excel_sheet.cells(2, 26), _
                             excel_sheet.cells(3, 26)).Select
                             
                            
                            
                           With excel_app.selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter 'xlBottom
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .shrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                           
                            excel_sheet.cells(2, 26) = "Schedule Date, Vehicle No & Team Captainy"
                            
                            
                            
                            
                            
                            

                            
                            
 excel_sheet.Range(excel_sheet.cells(1, 1), _
                             excel_sheet.cells(i, 26)).Select
'excel_sheet.Columns("A:A").Select
 excel_app.selection.Font.Size = 10

 
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With
                            
                            
                            
                            
                            
 





excel_sheet.Columns("A:A").Select
 excel_app.selection.columnWidth = 3.57
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


    excel_sheet.Columns("b:d").Select
 excel_app.selection.columnWidth = 14.86
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


  

excel_sheet.Columns("e:f").Select
 excel_app.selection.columnWidth = 17
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With


excel_sheet.Columns("g:Y").Select
 excel_app.selection.columnWidth = 8
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With

excel_sheet.Columns("Z:Z").Select
 excel_app.selection.columnWidth = 7
With excel_app.selection
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
.WrapText = True
.Orientation = 0
End With







                            
                            
                            
     excel_sheet.Range(excel_sheet.cells(1, 1), _
                             excel_sheet.cells(i, 27)).Select
                            
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
     excel_sheet.Range("A1:Z3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "PLANT DISTRIBUTION LIST"
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


Private Sub Form_Load()
Dim RSTR As New ADODB.Recordset

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
