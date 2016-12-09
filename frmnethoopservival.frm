VERSION 5.00
Begin VB.Form frmnethoopservival 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servival"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2850
   Icon            =   "frmnethoopservival.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   1320
      Picture         =   "frmnethoopservival.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Graph"
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
      Left            =   120
      Picture         =   "frmnethoopservival.frx":1B0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
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
      Height          =   1410
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmnethoopservival"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shipmentsize, healthyplant, oversize, undersize, weakreceived, icedamaged, totreceived, deadin10 As Long
Dim healthyin10, healthyplantsaftereye, hardenplants, servival1 As Long
Dim totbillladding, totreceivedin, totoversize, totundersize, totweak, toticedamaged, totreceivedex, totdead, tothealthy10, tothealthyeye, tothardenplants As Double

Dim Dzstr As String
Private Sub Command1_Click()
Dzstr = ""
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE FACILITY TYPE."
          Exit Sub
       End If
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
   MsgBox "FACILITY NOT SELECTED !!!"
   Exit Sub
End If

If InStr(1, Dzstr, "S") > 0 Then
NU_TC_Survival
End If

If InStr(1, Dzstr, "N") > 0 Then
NetHoop_Survival "N"
End If

If InStr(1, Dzstr, "H") > 0 Then
NetHoop_Survival "H"
End If

If InStr(1, Dzstr, "T") > 0 Then
NetHoop_Survival "T"
End If




End Sub
Private Sub NetHoop_Survival(housetype As String)
Dim mplantbatch As String
Dim mhouse As String
Dim SQLSTR As String
Dim maxdate As Date
Dim mindate As Date
Dim excel_app As Object
Dim excel_sheet As Object
Dim i, j As Integer
Dim rs As New ADODB.Recordset
Dim rsm As New ADODB.Recordset
Dim rss As New ADODB.Recordset

j = 0
mplantbatch = ""
mhouse = ""
'    SQLSTR = "select plantbatch" _
'           & "  from tblqmsboxdetail" _
'           & " where trnid>=39 " _
'           & " order by plantbatch"
'

    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_sheet.name = "Detail"
       Dim sl As Integer
    sl = 1
    i = 1
    excel_app.Visible = False
    excel_sheet.Cells(3, 1) = "FID"
    excel_sheet.Cells(3, 2) = "From"
    excel_sheet.Cells(3, 3) = "To"
    excel_sheet.Cells(3, 4) = "FID"
    excel_sheet.Cells(3, 5) = "PBID"
    excel_sheet.Cells(3, 6) = "Total Plants Received"
    excel_sheet.Cells(3, 7) = "Dead Plants"
    excel_sheet.Cells(3, 8) = "Debit by PV"
    excel_sheet.Cells(3, 9) = "% Survival at LMT"
    i = 4
                  
'    Set rs = Nothing
'    rs.Open SQLSTR, MHVDB
     housetype = "'" & housetype & "'"
'    If rs.EOF <> True Then
'
'
'     Do While rs.EOF <> True
      
    Set rsm = Nothing
    
   

    rsm.Open "select distinct plantbatch, facilityid from tblqmsplanttransaction where  transactiontype='6' and debit>0 and  facilityid in(select facilityid from tblqmsfacility where housetype in(" & housetype & ")) order by plantbatch,facilityid", MHVDB
    Do While rsm.EOF <> True
                ' from house
                Set rss = Nothing
                
                rss.Open "select distinct facilityid from tblqmsplanttransaction where transactiontype='6' and credit>0 and plantbatch='" & rsm!plantBatch & "'", MHVDB
                mhouse = ""
                Do While rss.EOF <> True
                mhouse = rss!facilityid & "," & mhouse
                rss.MoveNext
                Loop
                If Len(mhouse) > 0 Then
                mhouse = Left(mhouse, Len(mhouse) - 1)
                End If
                excel_sheet.Cells(i, 1) = mhouse
                ' from date to date
                Set rss = Nothing
                rss.Open "select min(entrydate) minentrydate,max(entrydate) maxentrydate from tblqmsplanttransaction where transactiontype='6' and debit>0 and plantbatch='" & rsm!plantBatch & "' and facilityid ='" & rsm!facilityid & "'", MHVDB
                If rsm.EOF <> True Then
                mindate = Format(rss!minentrydate, "dd/MM/yyyy")
                maxdate = Format(rss!maxentrydate, "dd/MM/yyyy")
                End If
                excel_sheet.Cells(i, 2) = "'" & mindate
                excel_sheet.Cells(i, 3) = "'" & maxdate
                ' to house
                excel_sheet.Cells(i, 4) = rsm!facilityid
                excel_sheet.Cells(i, 5) = rsm!plantBatch
                ' plant received
                Set rss = Nothing
                rss.Open "select sum(debit) as qty from tblqmsplanttransaction where  transactiontype='6' and debit>0 and  plantbatch='" & rsm!plantBatch & "' and facilityid ='" & rsm!facilityid & "'", MHVDB
                excel_sheet.Cells(i, 6) = rss!qty
                'dead plant
                Set rss = Nothing
                rss.Open "select sum(credit) as qty from tblqmsplanttransaction where  transactiontype='3' and  plantbatch='" & rsm!plantBatch & "' and facilityid='" & rsm!facilityid & "'", MHVDB
                excel_sheet.Cells(i, 7) = rss!qty
                'debit by pv
                Set rss = Nothing
                rss.Open "select sum(debit) as qty from tblqmsplanttransaction where  transactiontype='8' and  plantbatch='" & rsm!plantBatch & "' and facilityid='" & rsm!facilityid & "'", MHVDB
                excel_sheet.Cells(i, 8) = rss!qty
               
                excel_sheet.Cells(i, 9) = (Val(excel_sheet.Cells(i, 6)) + Val(excel_sheet.Cells(i, 8)) - Val(excel_sheet.Cells(i, 7))) / (Val(excel_sheet.Cells(i, 6)) + Val(excel_sheet.Cells(i, 8)))
                excel_sheet.Cells(i, 9).NumberFormat = "0%"
    
    i = i + 1
    rsm.MoveNext
    Loop
    
    
    
'    rs.MoveNext
'
'    Loop
       
      

    
    
'End If












                       
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:w3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "NU TC Survival"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(3, 23)).Select
    excel_app.Selection.ColumnWidth = 10
    
    With excel_app.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    End With
    
    
    excel_sheet.Range(excel_sheet.Cells(2, 1), _
                             excel_sheet.Cells(2, 2)).Select
                                           
                            
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
                            
                            
                            excel_sheet.Cells(2, 1) = "Date TC planted in Bed"
    
    
    
    
Excel_WBook.Sheets("sheet2").Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
'excel_sheet.Name = "Summary"





excel_app.Visible = True

    
    Screen.MousePointer = vbDefault
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


Set rsF = Nothing


rsF.Open "select distinct housetype from tblqmsfacility where housetype in('H','N','S','T')  Order by housetype", MHVDB, adOpenStatic
With rsF
Do While Not .EOF

Select Case rsF!housetype
            Case "C"
            DZLIST.AddItem "Cold House" + " | " + Trim(!housetype)
            Case "N"
            DZLIST.AddItem "Net House" + " | " + Trim(!housetype)
            Case "H"
            DZLIST.AddItem "Hoop House" + " | " + Trim(!housetype)
            Case "T"
            DZLIST.AddItem "Terrace" + " | " + Trim(!housetype)
            Case "S"
            DZLIST.AddItem "Staging House" + " | " + Trim(!housetype)
End Select

   .MoveNext
Loop
End With



'For i = 0 To DZLIST.ListCount - 1
'DZLIST.Selected(i) = True
'
'Next


Exit Sub
err:
MsgBox err.Description
End Sub



Private Sub NU_TC_Survival()
Dim SQLSTR As String
Dim maxdate As Date
Dim shipmentno As Integer
Dim excel_app As Object
Dim excel_sheet As Object
Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rspt As New ADODB.Recordset
Dim rsPb As New ADODB.Recordset
totbillladding = 0
totreceivedin = 0
totoversize = 0
totundersize = 0
totweak = 0
toticedamaged = 0
totreceivedex = 0
totdead = 0
tothealthy10 = 0
tothealthyeye = 0
tothardenplants = 0
j = 0
'and plantbatch in(select plantbatch from tblqmsplanttransaction where status='ON' and transactiontype='6')
    SQLSTR = "select trnid,entrydate,plantbatch,sum(shipmentsize) shipmentsize,sum(healthyplant) healthyplant,sum(weakplant) weakplant," _
           & " sum(undersize) undersize,sum(icedamaged) icedamaged ,sum(oversize) oversize,sum(deadplant) deadplant from tblqmsboxdetail" _
           & " where trnid>=39 " _
           & " group by trnid,entrydate,plantbatch" _
           & " order by trnid,plantbatch"
              

    Screen.MousePointer = vbHourglass
    DoEvents
    Set excel_app = CreateObject("Excel.Application")
    Set Excel_WBook = excel_app.Workbooks.Add
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If
    excel_sheet.name = "Detail"
       Dim sl As Integer
    sl = 1
    i = 1
    excel_app.Visible = True
    excel_sheet.Cells(3, 1) = "From"
    excel_sheet.Cells(3, 2) = "To"
    excel_sheet.Cells(3, 3) = "Shipment No."
    excel_sheet.Cells(3, 4) = "Plant Batch"
    excel_sheet.Cells(3, 5) = "Facility"
    excel_sheet.Cells(3, 6) = "Bill of Lading"
    excel_sheet.Cells(3, 7) = "Plants Received (including weak, Over Size & Ice Damage plants)"
    excel_sheet.Cells(3, 8) = "Oversize"
    excel_sheet.Cells(3, 9) = "Under Size*"
    excel_sheet.Cells(3, 10) = "Weak received"
    excel_sheet.Cells(3, 11) = "Ice Damaged"
    excel_sheet.Cells(3, 12) = "Plants Received (Excluding Weak, Ice Damaged & Under Size)"
    excel_sheet.Cells(3, 13) = "No. Plants dead within 10 days"
    excel_sheet.Cells(3, 14) = "Healthy Plants (After 10 Days)"
    excel_sheet.Cells(3, 15) = "Eye Assessment"
    excel_sheet.Cells(3, 16) = "Healthy Plants After Eye "
    excel_sheet.Cells(3, 17) = "From "
    excel_sheet.Cells(3, 18) = "To"
    excel_sheet.Cells(3, 19) = "Harden Plants"
    excel_sheet.Cells(3, 20) = "% Survival on the basis of Initial Healthy Plants(Excluding Weak, Ice Damaged & Under Size)"
    excel_sheet.Cells(3, 21) = "% Survival on the basis of Healthy Plants after 10 days in Staging House"
    excel_sheet.Cells(3, 22) = "% Survival on the basis of total plants received (Including weak, ice damaged and undersize)"
    excel_sheet.Cells(3, 23) = "Remarks"
    i = 4
                  
    Set rs = Nothing
    rs.Open SQLSTR, MHVDB
    If rs.EOF <> True Then
    
     Do Until rs.EOF
     deadin10 = 0
     healthyin10 = 0
     healthyplantsaftereye = 0
     hardenplants = 0
     servival1 = 0
     shipmentno = rs!trnid
     Set rspt = Nothing
     rspt.Open "select entrydate as mindt from tblqmsplantbatchhdr where" _
     & " trnid='" & rs!trnid & "'", MHVDB
     If rspt.EOF <> True Then
     excel_sheet.Cells(i, 1) = "'" & Format(rspt!mindt, "dd/MM/yyyy")
     End If
     Set rspt = Nothing
     rspt.Open "select max(entrydate)maxdt from tblqmsplanttransaction where" _
     & " transactiontype='2' and plantbatch='" & rs!plantBatch & "'", MHVDB
     If rspt.EOF <> True Then
     excel_sheet.Cells(i, 2) = "'" & Format(rspt!maxdt, "dd/MM/yyyy")
     maxdate = IIf(IsNull(rspt!maxdt), "01/01/1999", rspt!maxdt)
     End If
     excel_sheet.Cells(i, 3) = shipmentno
     excel_sheet.Cells(i, 3).Font.Bold = True
     Do While shipmentno = rs!trnid
     ' transplant to bed
     'no of dead within 10 days
     Set rspt = Nothing
     rspt.Open "select sum(credit) as dead from tblqmsplanttransaction where" _
     & " transactiontype='3' and plantbatch='" & rs!plantBatch & "'" _
     & " and entrydate>='" & Format(maxdate, "yyyy-MM-dd") & "' and entrydate<='" & Format(maxdate + 10, "yyyy-MM-dd") & "'", MHVDB
     If rspt.EOF <> True Then
     excel_sheet.Cells(i, 13) = IIf(rspt!dead = 0, "", rspt!dead)
     deadin10 = deadin10 + IIf(IsNull(rspt!dead), 0, rspt!dead)
     End If
     excel_sheet.Cells(i, 4) = rs!plantBatch
     'excel_sheet.Cells(i, 5) = rs!facilityid
     excel_sheet.Cells(i, 6) = rs!shipmentsize
     excel_sheet.Cells(i, 8) = IIf(rs!oversize = 0, "", rs!oversize)
     excel_sheet.Cells(i, 9) = IIf(rs!undersize = 0, "", rs!undersize)
     excel_sheet.Cells(i, 10) = IIf(rs!weakplant = 0, "", rs!weakplant)
     excel_sheet.Cells(i, 11) = IIf(rs!icedamaged = 0, "", rs!icedamaged)
     excel_sheet.Cells(i, 12) = rs!healthyplant '- rs!oversize - rs!undersize - rs!weakplant - rs!icedamaged
     excel_sheet.Cells(i, 7) = rs!healthyplant + rs!oversize + rs!undersize + rs!weakplant + rs!icedamaged
     excel_sheet.Cells(i, 14) = Val(excel_sheet.Cells(i, 7)) - rs!oversize - rs!undersize - rs!weakplant - rs!icedamaged - Val(excel_sheet.Cells(i, 13))
     ' eye assstment
     Set rspt = Nothing
     rspt.Open "select sum(credit) as qty from tblqmsplanttransaction where" _
     & " verificationtype='3' and transactiontype='3' and plantbatch='" & rs!plantBatch & "'", MHVDB
     If rspt.EOF <> True Then
     If IIf(IsNull(rspt!qty), 0, rspt!qty) <> 0 Then
     excel_sheet.Cells(i, 15) = 1 - ((IIf(IsNull(rspt!qty), 0, rspt!qty)) / Val(excel_sheet.Cells(i, 12)))
     excel_sheet.Cells(i, 15).NumberFormat = "0%"
     Else
     excel_sheet.Cells(i, 15) = "100%"
     End If
     ' excel_sheet.Cells(i, 15) = excel_sheet.Cells(i, 15) & "%"
     End If
     excel_sheet.Cells(i, 16) = IIf(Val(excel_sheet.Cells(i, 7)) * Val(excel_sheet.Cells(i, 15)) = 0, "", Val(excel_sheet.Cells(i, 7)) * Val(excel_sheet.Cells(i, 15)))
     excel_sheet.Cells(i, 16).NumberFormat = "0"
     healthyplantsaftereye = healthyplantsaftereye + Val(excel_sheet.Cells(i, 16))
     healthyin10 = healthyin10 + Val(excel_sheet.Cells(i, 14))
     ' transplant hard to bags
     Set rspt = Nothing
     rspt.Open "select min(entrydate) mindt, max(entrydate)maxdt,sum(credit) as qty from tblqmsplanttransaction where" _
     & " transactiontype='6' and plantbatch='" & rs!plantBatch & "' ", MHVDB
     If rspt.EOF <> True Then
     excel_sheet.Cells(i, 17) = "'" & Format(rspt!mindt, "dd/MM/yyyy")
     excel_sheet.Cells(i, 18) = "'" & Format(rspt!maxdt, "dd/MM/yyyy")
     excel_sheet.Cells(i, 19) = IIf((IIf(IsNull(rspt!qty), 0, rspt!qty)) = 0, "", Abs(IIf(IsNull(rspt!qty), 0, rspt!qty)))
     hardenplants = hardenplants + Val(excel_sheet.Cells(i, 19))
     End If
     excel_sheet.Cells(i, 20) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 12))
     excel_sheet.Cells(i, 20).NumberFormat = "0%"
    excel_sheet.Cells(i, 21) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 14))
    excel_sheet.Cells(i, 21).NumberFormat = "0%"
    excel_sheet.Cells(i, 22) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 7))
    excel_sheet.Cells(i, 22).NumberFormat = "0%"
    
    Set rsPb = Nothing
    rsPb.Open "select * from tblqmsplanttransaction  where plantbatch='" & rs!plantBatch & "' and status='ON' and transactiontype='6'", MHVDB
    If rsPb.EOF <> True Then
    totbillladding = totbillladding + Val(excel_sheet.Cells(i, 6))
    totreceivedin = totreceivedin + Val(excel_sheet.Cells(i, 7))
    totoversize = totoversize + Val(excel_sheet.Cells(i, 8))
    totundersize = totundersize + Val(excel_sheet.Cells(i, 9))
    totweak = totweak + Val(excel_sheet.Cells(i, 10))
    toticedamaged = toticedamaged + Val(excel_sheet.Cells(i, 11))
    totreceivedex = totreceivedex + Val(excel_sheet.Cells(i, 12))
    totdead = totdead + Val(excel_sheet.Cells(i, 13))
    tothealthy10 = tothealthy10 + Val(excel_sheet.Cells(i, 14))
    tothealthyeye = tothealthyeye + Val(excel_sheet.Cells(i, 16))
    tothardenplants = tothardenplants + Val(excel_sheet.Cells(i, 19))
    Else
    'MsgBox "Mukti"
    End If
    
    
    
    i = i + 1
    rs.MoveNext
    If rs.EOF Then Exit Do
    Loop
       'i = i + 1
'        excel_sheet.Cells(i, 3) = "Shipment Total"
'        excel_sheet.Cells(i, 3).Font.Bold = True
        getshipmenttot shipmentno
        excel_sheet.Cells(i, 6) = shipmentsize
        excel_sheet.Cells(i, 7) = IIf(totreceived = 0, "", totreceived)
        excel_sheet.Cells(i, 8) = IIf(oversize = 0, "", oversize)
        excel_sheet.Cells(i, 9) = IIf(undersize = 0, "", undersize)
        excel_sheet.Cells(i, 10) = IIf(weakreceived = 0, "", weakreceived)
        excel_sheet.Cells(i, 11) = IIf(icedamaged = 0, "", icedamaged)
        excel_sheet.Cells(i, 12) = healthyplant
        excel_sheet.Cells(i, 13) = IIf(deadin10 = 0, "", deadin10)
        excel_sheet.Cells(i, 14) = IIf(healthyin10 = 0, "", healthyin10)
       
       
       Set rspt = Nothing
       rspt.Open "select sum(credit) as qty from tblqmsplanttransaction where" _
               & " verificationtype='3' and transactiontype='3' and plantbatch in(select plantbatch from tblqmsplantbatchdetail where trnid='" & shipmentno & "')", MHVDB
       If rspt.EOF <> True Then
       If IIf(IsNull(rspt!qty), 0, rspt!qty) <> 0 Then
        excel_sheet.Cells(i, 15) = 1 - ((IIf(IsNull(rspt!qty), 0, rspt!qty)) / Val(excel_sheet.Cells(i, 12)))
        excel_sheet.Cells(i, 15).NumberFormat = "0%"
        Else
        excel_sheet.Cells(i, 15) = "100%"
        End If
        ' excel_sheet.Cells(i, 15) = excel_sheet.Cells(i, 15) & "%"
        End If
      
      
        excel_sheet.Cells(i, 16) = IIf(healthyplantsaftereye = 0, "", healthyplantsaftereye)
        excel_sheet.Cells(i, 16).NumberFormat = "0"
        excel_sheet.Cells(i, 19) = IIf(hardenplants = 0, "", hardenplants)
        excel_sheet.Cells(i, 20) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 12))
        excel_sheet.Cells(i, 20).NumberFormat = "0%"
        excel_sheet.Cells(i, 21) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 14))
        excel_sheet.Cells(i, 21).NumberFormat = "0%"
        excel_sheet.Cells(i, 22) = (Val(excel_sheet.Cells(i, 19))) / Val(excel_sheet.Cells(i, 7))
        excel_sheet.Cells(i, 22).NumberFormat = "0%"
       
       

       
       
       
       excel_sheet.Range(excel_sheet.Cells(i, 6), _
       excel_sheet.Cells(i, 22)).Select
       
       excel_app.Selection.Interior.ColorIndex = 15
       excel_app.Selection.Font.Color = vbBlue
       excel_app.Selection.Font.Bold = True
       
        i = i + 1
    Loop
    
excel_sheet.Cells(1, 6) = totbillladding
excel_sheet.Cells(1, 7) = totreceivedin
excel_sheet.Cells(1, 8) = totoversize
excel_sheet.Cells(1, 9) = totundersize
excel_sheet.Cells(1, 10) = totweak
excel_sheet.Cells(1, 11) = toticedamaged
excel_sheet.Cells(1, 12) = totreceivedex
excel_sheet.Cells(1, 13) = totdead
excel_sheet.Cells(1, 14) = tothealthy10
excel_sheet.Cells(1, 16) = tothealthyeye
excel_sheet.Cells(1, 16).NumberFormat = "0"
excel_sheet.Cells(1, 19) = tothardenplants
excel_sheet.Cells(1, 20) = (Val(excel_sheet.Cells(1, 19))) / Val(excel_sheet.Cells(1, 12))
excel_sheet.Cells(1, 20).NumberFormat = "0%"
excel_sheet.Cells(1, 21) = (Val(excel_sheet.Cells(1, 19))) / Val(excel_sheet.Cells(1, 14))
excel_sheet.Cells(1, 21).NumberFormat = "0%"
excel_sheet.Cells(1, 22) = (Val(excel_sheet.Cells(1, 19))) / Val(excel_sheet.Cells(1, 7))
excel_sheet.Cells(1, 22).NumberFormat = "0%"
excel_sheet.Range(excel_sheet.Cells(1, 6), _
excel_sheet.Cells(1, 22)).Select
excel_app.Selection.Interior.ColorIndex = 15
excel_app.Selection.Font.Color = vbBlue
excel_app.Selection.Font.Bold = True
    
End If












                       
    'make up
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(i, 6)).Select
    excel_app.Selection.Columns.AutoFit
    excel_sheet.Cells(4, 2).Select
    excel_app.ActiveWindow.FreezePanes = True
    excel_sheet.Cells(1, 1).Select
    With excel_sheet
     excel_sheet.Range("A3:w3").Font.Bold = True
    .PageSetup.CenterHeader = "Mountain Hazelnut  Venture Private Limited"
    .PageSetup.CenterFooter = "NU TC Survival"
    .PageSetup.LeftFooter = "MHV"
    .PageSetup.RightFooter = "Print On " + Format(Date, "dd/mm/yyyy")
    .PageSetup.PrintGridlines = True
    End With
    excel_sheet.Range(excel_sheet.Cells(3, 1), _
    excel_sheet.Cells(3, 23)).Select
    excel_app.Selection.ColumnWidth = 10
    
    With excel_app.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    .Orientation = 0
    End With
    
    
    excel_sheet.Range(excel_sheet.Cells(2, 1), _
                             excel_sheet.Cells(2, 2)).Select
                                           
                            
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
                            
                            
                            excel_sheet.Cells(2, 1) = "Date TC planted in Bed"
    
    
    
    
Excel_WBook.Sheets("sheet2").Activate
If Val(excel_app.Application.Version) >= 8 Then
    Set excel_sheet = excel_app.ActiveSheet
Else
    Set excel_sheet = excel_app
End If
excel_sheet.name = "Summary"





excel_app.Visible = True

    
    Screen.MousePointer = vbDefault
    Set excel_sheet = Nothing
    Set excel_app = Nothing
    Screen.MousePointer = vbDefault


End Sub
Private Sub getshipmenttot(shipmentno As Integer)
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
shipmentsize = 0
healthyplant = 0
oversize = 0
undersize = 0
weakreceived = 0
icedamaged = 0
totreceived = 0
 SQLSTR = "select trnid,sum(shipmentsize) shipmentsize,sum(healthyplant) healthyplant,sum(weakplant) weakplant," _
           & " sum(undersize) undersize,sum(icedamaged) icedamaged ,sum(oversize) oversize,sum(deadplant) deadplant from tblqmsboxdetail" _
           & " where trnid='" & shipmentno & "'" _
           & " group by trnid"
 Set rs = Nothing
 rs.Open SQLSTR, MHVDB
 If rs.EOF <> True Then
    shipmentsize = rs!shipmentsize
    healthyplant = rs!healthyplant
    oversize = rs!oversize
    undersize = rs!undersize
    weakreceived = rs!weakplant
    icedamaged = rs!icedamaged
    totreceived = rs!healthyplant + rs!oversize + rs!undersize + rs!weakplant + rs!icedamaged
 End If

End Sub

