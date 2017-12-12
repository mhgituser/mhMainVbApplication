Attribute VB_Name = "MHVMODULE"
Public UserName
Public mFarmerType
Public myMonitor
Public mDescription As String
Public emailAddress As String
Public validateMsg As String
Public DeptName As String
Public mshorttime As String
Public paramName As String
Public monitorFromFarmer As String
Public excelPath As String
Public dashBoardName As String
Public Mnuttype As String
Public ispercentage As Boolean
Public tempStorageName As String
Public mtempStorageName As String
Public qmsplantbatch3 As String
Public mPlanttype As String
Public mPlantVariety As String
Public paramfieldname As String
Public paramFieldStorage As String
Public paramDesc As String
Public pestparamname As String
Public pestFarmer As String
Public emailId As String
Public paramValue As Double
Public emailMessageString As String
Public acceptableThresholdValue As String
Public errReportName As String
Public firstMondayOfTheMonth As Date
Public qmsShimentNo As Integer
Public fieldIndex As Integer
Public Mlocation As String
Public odkTableName As String
Public ReportName As String
Public qmsbatchTotal As Double
Public qmsshiptotal As Double
Public qmsBatchReceivedDate As Date
Public qmshousetype As String
Public querytype As Integer
Public qmsQueryPara As String
Public locationFromFid As String
Public qmsReportName As String
Public qmsShade As String
Public qmsTime As String
Public qmsCloud As String
Public qmsWind As String
Public qmsApplicationMethod As String
Public boxOperation As String
Public qmsTransactionType As String
Public qmsVerificationType As String
Public qmsFacility As String
Public qmsBatchdetail1 As String
Public qmsBatchdetail As String
Public qmsStatus As String
Public qmsChemical As String
Public qmsChemicalTradeName As String
Public qmsPlantVariety As String
Public qmsPlantType As String
Public UserId As String
Public Mmodule As String
Public Mserver As String
Public gEmailMessage As String
Public ProcYear As Integer
Public Mtblname, Mtblname1 As String
Public CnnString As String
Public CnnsecString As String
Public OdkCnnString As String
Public connRemote As String
Public MhwebCnnString As String
Public MhvhrCnnString As String
Public MHVDB As New ADODB.Connection
Public ODKDB As New ADODB.Connection
Public MHWEBDB As New ADODB.Connection
Public MHVHRDB As New ADODB.Connection
Private pnaConn As New ADODB.Connection
Public MHVsecDB As New ADODB.Connection
Public RptOption As String
Public Mcaretaker As String
Public ISCARETAKER As Boolean
Public MODULETYPE As String
Public mFARID As String
Public mbypass As Boolean
Public MypicPath As String
Public Operation As String
Public LogRemarks As String
Public ANAME As String
Public Mge As Boolean
Public Mname  '= Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
Public MAINMODULEID As String
Public MAINMODULENAME As String
Public mchk As Boolean
Public chkred As Boolean
Public KmlFileNameas As String
Public MUSER As String

Public Const pbStringFormat = "^Sl.No.|^Plant Variety|^Batch No.|^Plant Type|^B/L Shipment Size|^Healthy Plant|^Weak Plant|^Under Size|^Over Size|^Ice Damaged|^TC source|"

Public Dzname, GEname, TsName, sTAFF, FATYPEINF, FAName, rOLEnAME, Mstatus, ActName As String
Public SysYear As Integer
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As _
Long
Private Const SW_SHOW = 1

 
    Const nodeCount As Integer = 9          'number of nodes - 1
    Type DijkEdge
        weight As Integer                   'distance from vertices that it is connected to
        destination As Integer              'name of vertice that it is connected to
    End Type
 
    Type Vertex
        connections(nodeCount) As DijkEdge  'hold information above for each connection
        numConnect As Integer               'number of connections - 1
        distance As Integer                 'distance from all other vertices
        isDead As Boolean                   'distance calculated
        name As Integer                     'name of vertice
    End Type

'Option Explicit



Public Function ValidateString(strInput As String) As String
    Dim strInvalidChars As String
    Dim i As Long
    strInvalidChars = "'"
    For i = 1 To Len(strInvalidChars)
        strInput = Replace$(strInput, Mid$(strInvalidChars, i, 1), "")
    Next
    ValidateString = strInput
End Function
Public Function ValidateLocationString(strInput As String) As String
    Dim strInvalidChars As String
    Dim i As Long
    strInvalidChars = "_"
    For i = 1 To Len(strInvalidChars)
        strInput = Replace$(strInput, Mid$(strInvalidChars, i, 1), " ")
    Next
    ValidateLocationString = strInput
    
End Function
'Public Function GetTbl()
'Dim db As New ADODB.Connection
'Set db = New ADODB.Connection
'db.CursorLocation = adUseClient
'db.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=ODKLOCAL;Initial Catalog=odk_prodLocal" ' local connection
'
'Mtblname = ""
'Dim chrNum As Integer
'    Dim rndInt, rndCase As Long
'    Dim str As String
'    Dim allowDupe As Boolean
'
'    allowDupe = False 'Switch to "True" to allow duplicate letters
'
'    For chrNum = 1 To 10 'Number of letters to generate: 10
'        rndCase = numRandomBetween(1, 10) 'Generates a number between 1 and 10
'            If rndCase > 5 Then
'                rndInt = numRandomBetween(65, 90) 'Generates random upper-case ascii character integer number
'            Else
'                rndInt = numRandomBetween(97, 122) 'Generates random lower-case ascii character integer number
'            End If
'
'            If allowDupe Then
'                str = (str) & Chr$(rndInt) 'Adds the generated character to the final string
'            Else
'                If InStr(str, Chr$(rndInt)) <= 0 Then str = (str) & Chr$(rndInt)
'            End If
'    Next
'
'    Mtblname = str
'
''    If Len(Mtblname) > 4 Then
''    Mtblname = Left(Mtblname, 4)
''    End If
'    Dim CrtStr As String
'
'    CrtStr = "CREATE table " & Mtblname & " (`_URI` varchar(80) NOT NULL," _
'                    & "`end` date NOT NULL," _
'                    & "`dcode` int(11) NOT NULL," _
'                    & "`gcode` int(11) NOT NULL," _
'                    & "`tcode` int(11) NOT NULL,`fcode` int(11) NOT NULL," _
'                    & "`farmercode` varchar(20) NOT NULL,`totaltrees` int(11) NOT NULL, " _
'                    & "`fs` varchar(1) NOT NULL,  " _
'                    & "`fdcode` int(11) NOT NULL, `id` int(11) NOT NULL,  " _
'                    & "`sname` varchar(50) NOT NULL,`fname` varchar(50) NOT NULL, " _
'                    & " `area` decimal(10,0) NOT NULL,`slowgrowing` int(11) NOT NULL,  " _
'                    & "`dor` int(11) NOT NULL,`deadmissing` int(11) NOT NULL," _
'                    & " `activegrowing` int(11) NOT NULL, `shock` int(11) NOT NULL, " _
'                    & "`nutrient` int(11) NOT NULL,`waterlog` int(11) NOT NULL, " _
'                    & "`leafpest` int(11) NOT NULL,`activepest` int(11) NOT NULL," _
'                    & " `stempest` int(11) NOT NULL, `rootpest` int(11) NOT NULL, " _
'                    & "`ANIMALDAMAGE` int(11) NOT NULL,`treesreceived` int(11) NOT NULL, " _
'                    & " `goodmoisture` int(11) NOT NULL, `poormoisture` int(11) NOT NULL," _
'                    & " `totaltally` int(11) NOT NULL,`monitorcomments` varchar(500) NOT NULL," _
'                    & " `btree1` int(11) NOT NULL,`etree1` int(11) NOT NULL,`ptree1` int(11) NOT NULL)ENGINE=MyISAM DEFAULT CHARSET=utf8"
'
'
'    db.Execute CrtStr
'
'
'End Function
Function GetTblmhv()
Dim RandomString As String
    Dim i As Integer
    Dim db As New ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                       
RandomString = ""
Mtblname = ""
    Do While i < 9
        Randomize
        Select Case IIf(i = 0, Int(1 * Rnd + 1), Int(2 * Rnd))
            Case 0: RandomString = RandomString & Chr(Int(10 * Rnd + 48))
            Case 1: RandomString = RandomString & Chr(Int(26 * Rnd + 65))
        End Select
        i = i + 1
    Loop
    
    Mtblname = RandomString
    'db.Execute "drop table " & Mtblname & ""
    Dim CrtStr As String
    
    CrtStr = "CREATE table " & Mtblname & " (`_URI` varchar(80) NOT NULL," _
                    & "`start` datetime NOT NULL," _
                    & "`tdate` datetime NOT NULL," _
                    & "`end` datetime NOT NULL," _
                    & "`dcode` varchar(100) NOT NULL," _
                    & "`gcode` varchar(100) NOT NULL," _
                    & "`tcode` varchar(100) NOT NULL,`fcode` int(11) NOT NULL," _
                     & "`fstype` varchar(100) NOT NULL," _
                    & "`farmercode` varchar(20) NOT NULL,`totaltrees` int(11) NOT NULL, " _
                    & "`fs` varchar(1) NOT NULL,  " _
                    & "`fdcode` int(11) NOT NULL, `id` varchar(50) NOT NULL,  " _
                    & "`sname` varchar(50) NOT NULL,`fname` varchar(50) NOT NULL, " _
                    & " `area` decimal(10,0) NOT NULL,`tree_count_slowgrowing` int(11) NOT NULL,  " _
                    & "`tree_count_dor` int(11) NOT NULL,`tree_count_deadmissing` int(11) NOT NULL," _
                    & " `tree_count_activegrowing` int(11) NOT NULL, `shock` int(11) NOT NULL, " _
                    & "`nutrient` int(11) NOT NULL,`waterlog` int(11) NOT NULL, " _
                    & "`leafpest` int(11) NOT NULL,`activepest` int(11) NOT NULL," _
                    & " `stempest` int(11) NOT NULL, `rootpest` int(11) NOT NULL, " _
                    & "`ANIMALDAMAGE` int(11) NOT NULL,`treesreceived` int(11) NOT NULL, " _
                    & " `goodmoisture` int(11) NOT NULL, `poormoisture` int(11) NOT NULL," _
                    & " `totaltally` int(11) NOT NULL,`monitorcomments` varchar(500) NOT NULL," _
                    & " `btree1` int(11) NOT NULL,`etree1` int(11) NOT NULL,`ptree1` int(11) NOT NULL)ENGINE=MyISAM DEFAULT CHARSET=utf8"

    
    db.Execute CrtStr
End Function

Function randomKey() As String
Dim RandomString As String
    Dim i As Integer
    Dim db As New ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                       
RandomString = ""
randomKey = ""
i = 0
    Do While i <= 10
        Randomize
        Select Case IIf(i = 0, Int(1 * Rnd + 1), Int(2 * Rnd))
            Case 0: RandomString = RandomString & Chr(Int(10 * Rnd + 48))
            Case 1: RandomString = RandomString & Chr(Int(26 * Rnd + 65))
        End Select
        i = i + 1
        
    Loop
    
    randomKey = LCase(RandomString)
     
End Function


Function GetTbl()
Dim RandomString As String
    Dim i As Integer
    Dim db As New ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                       
RandomString = ""
Mtblname = ""
    Do While i < 9
        Randomize
        Select Case IIf(i = 0, Int(1 * Rnd + 1), Int(2 * Rnd))
            Case 0: RandomString = RandomString & Chr(Int(10 * Rnd + 48))
            Case 1: RandomString = RandomString & Chr(Int(26 * Rnd + 65))
        End Select
        i = i + 1
    Loop
    
    Mtblname = RandomString
    'db.Execute "drop table " & Mtblname & ""
    Dim CrtStr As String
    
    CrtStr = "CREATE table " & Mtblname & " (`_URI` varchar(80) NOT NULL, `odktable` int(11) NOT NULL," _
                    & "`start` datetime NOT NULL," _
                    & "`tdate` datetime NOT NULL," _
                    & "`end` datetime NOT NULL," _
                    & "`dcode` varchar(100) NOT NULL," _
                    & "`gcode` varchar(100) NOT NULL," _
                    & "`tcode` varchar(100) NOT NULL,`fcode` int(11) NOT NULL," _
                     & "`fstype` varchar(100) NOT NULL," _
                    & "`farmercode` varchar(20) NOT NULL,`totaltrees` int(11) NOT NULL, " _
                    & "`fs` varchar(1) NOT NULL,  " _
                    & "`fdcode` int(11) NOT NULL, `id` varchar(50) NOT NULL,  " _
                    & "`sname` varchar(50) NOT NULL,`fname` varchar(50) NOT NULL, " _
                    & " `area` decimal(10,0) NOT NULL,`tree_count_slowgrowing` int(11) NOT NULL,  " _
                    & "`tree_count_dor` int(11) NOT NULL,`tree_count_deadmissing` int(11) NOT NULL," _
                    & " `tree_count_activegrowing` int(11) NOT NULL, `shock` int(11) NOT NULL, " _
                    & "`nutrient` int(11) NOT NULL,`waterlog` int(11) NOT NULL, " _
                    & "`leafpest` int(11) NOT NULL,`activepest` int(11) NOT NULL," _
                    & " `stempest` int(11) NOT NULL, `rootpest` int(11) NOT NULL, " _
                    & "`ANIMALDAMAGE` int(11) NOT NULL,`treesreceived` int(11) NOT NULL, " _
                    & " `goodmoisture` int(11) NOT NULL, `poormoisture` int(11) NOT NULL," _
                    & " `totaltally` int(11) NOT NULL,`monitorcomments` varchar(500) NOT NULL," _
                    & " `gpslat` decimal(38,10) NOT NULL,`gpslng` decimal(38,10) NOT NULL," _
                    & " `btree1` int(11) NOT NULL,`etree1` int(11) NOT NULL,`ptree1` int(11) NOT NULL)ENGINE=MyISAM DEFAULT CHARSET=utf8"

    
    db.Execute CrtStr
    
End Function
Function GetTbl1()
Dim RandomString As String
    Dim i As Integer
    Dim db As New ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
     i = 0
RandomString = ""
Mtblname1 = ""
    Do While i < 9
        Randomize
        Select Case IIf(i = 0, Int(1 * Rnd + 1), Int(2 * Rnd))
            Case 0: RandomString = RandomString & Chr(Int(10 * Rnd + 48))
            Case 1: RandomString = RandomString & Chr(Int(26 * Rnd + 65))
        End Select
        i = i + 1
    Loop
    
    Mtblname1 = RandomString
    
    Dim CrtStr As String
    
   CrtStr = "CREATE table " & Mtblname1 & " (`_URI` varchar(80) NOT NULL," _
                    & "`start` datetime NOT NULL," _
                    & "`tdate` datetime NOT NULL," _
                    & "`end` datetime NOT NULL," _
                    & "`dcode` varchar(100) NOT NULL," _
                    & "`gcode` varchar(100) NOT NULL," _
                    & "`tcode` varchar(100) NOT NULL,`fcode` int(11) NOT NULL," _
                     & "`fstype` varchar(100) NOT NULL," _
                    & "`farmercode` varchar(20) NOT NULL,`totaltrees` int(11) NOT NULL, " _
                    & "`fs` varchar(1) NOT NULL,  " _
                    & "`fdcode` int(11) NOT NULL, `id` varchar(50) NOT NULL,  " _
                    & "`sname` varchar(50) NOT NULL,`fname` varchar(50) NOT NULL, " _
                    & " `area` decimal(10,0) NOT NULL,`tree_count_slowgrowing` int(11) NOT NULL,  " _
                    & "`tree_count_dor` int(11) NOT NULL,`tree_count_deadmissing` int(11) NOT NULL," _
                    & " `tree_count_activegrowing` int(11) NOT NULL, `shock` int(11) NOT NULL, " _
                    & "`nutrient` int(11) NOT NULL,`waterlog` int(11) NOT NULL, " _
                    & "`leafpest` int(11) NOT NULL,`activepest` int(11) NOT NULL," _
                    & " `stempest` int(11) NOT NULL, `rootpest` int(11) NOT NULL, " _
                    & "`ANIMALDAMAGE` int(11) NOT NULL,`treesreceived` int(11) NOT NULL, " _
                    & " `goodmoisture` int(11) NOT NULL, `poormoisture` int(11) NOT NULL," _
                    & " `totaltally` int(11) NOT NULL,`monitorcomments` varchar(500) NOT NULL," _
                    & "`goodmoisture` int(11) NOT NULL," _
                    & " `btree1` int(11) NOT NULL,`etree1` int(11) NOT NULL,`ptree1` int(11) NOT NULL)ENGINE=MyISAM DEFAULT CHARSET=utf8"

    
    db.Execute CrtStr
    
End Function




Private Function numRandomBetween(lowerNum As Integer, higherNum As Integer) As Long
 
    numRandomBetween = Int((higherNum - lowerNum + 1) * Rnd + lowerNum)
 
End Function
Public Sub Navigate(ByVal NavTo As String)
    If NavTo = "" Then Exit Sub
    Dim hBrowse As Long
    hBrowse = ShellExecute(0&, "open", NavTo, "", "", SW_SHOW)
End Sub
Public Function updatemhvlog(date_of_modification As Date, modified_by As String, desc As String, remarks As String)
On Error GoTo err
MHVDB.Execute "insert into tbltrnlog(date_of_modification,modified_by,description,remarks)" _
              & " values('" & Format(date_of_modification, "yyyy-MM-dd") & "','" & modified_by & "','" & desc & "','" & remarks & "')"
Exit Function
err:
        MsgBox err.Description
End Function
Public Function updateodklog(uri As String, date_of_modification As Date, modified_by As String, desc As String, remarks As String)
On Error GoTo err
    Dim db As New ADODB.Connection
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                       
db.Execute "insert into tblodkmodificationlog(_uri,date_of_modification,modified_by,description,table_name)" _
              & " values('" & uri & "','" & Format(date_of_modification, "yyyy-MM-dd") & "','" & modified_by & "','" & desc & "','" & UCase(remarks) & "')"
Exit Function
err:
        MsgBox err.Description
End Function
'Public Function hizbiz(ByVal Rep As String) As String
'
'       hizbiz = ""
'       X = Asc(Mid(Rep, 1, 1))
'       For i = 1 To Len(Rep)
'           j = Chr(Asc(Mid(Rep, i, 1)) + X)
'           If Asc(j) = 34 Or Asc(j) = 39 Then j = "|"
'           G = Chr((Asc(Mid(Rep, i, 1)) * (i + X)) Mod 255)
'           If Asc(G) = 34 Or Asc(G) = 39 Or Asc(G) = 0 Then G = "|"
'           hizbiz = hizbiz + j + G
'       Next
'End Function
Public Function Check93(ByVal Xstr As String, ByVal Jlen As Long, ByVal JMOD) As String
Dim i, JSum, j As Long
Dim Jstr As String
Xstr = Trim(Xstr)
If Jlen > 0 Then
   If Len(Xstr) < Jlen Then Xstr = UCase(PadWithChar(Xstr, Jlen, "0", 1))
End If
JSum = 0
j = 0
For i = Len(Xstr) To 1 Step -1
    j = j + 1
    Jstr = Mid(Xstr, i, 1)
    If Jstr >= "0" And Jstr <= "9" Then
        JSum = JSum + Val(Jstr) * j
    Else
        JSum = JSum + (Asc(Jstr) - Asc("A") + 10) * j
    End If
Next
JSum = JSum Mod JMOD
If JSum < 10 Then
   Check93 = Xstr + Trim(str(JSum))
Else
   Check93 = Xstr + Chr(55 + JSum)
End If
End Function
Public Function BreakePos(ByVal Xstr As String, UpToPos As Long) As Long
Dim JI As Integer
BreakePos = UpToPos
If Len(Trim(Xstr)) <= UpToPos Then
   BreakePos = Len(Trim(Xstr))
Else
   For JI = UpToPos To 1 Step -1
       If Mid(Xstr, JI, 1) = " " Then 'Or Mid(Xstr, JI, 1) = "," Or Mid(Xstr, JI, 1) = "." Or Mid(Xstr, JI, 1) = "-" Then
          BreakePos = JI
          Exit For
       End If
   Next
End If

End Function
'Public Function GraphPop(ByVal Rep As String)
'Dim DB
'Dim Sales As New ADODB.Recordset
'Set DB = New Connection
'DB.CursorLocation = adUseClient
'DB.Open CnnString
'frmChart.MSC.Title = "Monthly Sales Trend"
'frmChart.MSC.ColumnCount = 1 '.RecordCount
'Sales.Open "select procmonth,sum(actual) as sales from budget where procyear='" & SysYear & "' and actual>0 group by procmonth order by procmonth", DB, adOpenStatic
'With Sales
'.MoveFirst
'frmChart.MSC.RowCount = .RecordCount
''MSC.row = 1
'frmChart.ChartGrd.Cols = .RecordCount + 1
'frmChart.ChartGrd.ColWidth(0) = 2200
'For I = 1 To .RecordCount
'    frmChart.ChartGrd.TextMatrix(0, I) = Mname(I)
'    frmChart.ChartGrd.ColWidth(I) = 1100
'Next
'I = 1
'frmChart.MSC.Column = 1
'Do While Not .EOF
'   'frmChart.MSC.column = i
'   frmChart.MSC.row = I
'   'frmChart.MSC.ColumnLabel = !guesttype
'   frmChart.MSC.RowLabel = Mname(!procmonth)
'   frmChart.MSC.Data = Round(!Sales / 100000, 2)
'   frmChart.ChartGrd.TextMatrix(1, I) = Round(!Sales / 100000, 2)
'   .MoveNext
'   I = I + 1
'Loop
'End With
'DB.Close
'frmChart.Show 1
'End Function
Public Sub openDb()
Dim FPATH As String
On Error GoTo err
MHVDB.Open CnnString
'MHVsecDB.Open ""
Exit Sub
err:
MsgBox err.Description
err.Clear
End Sub
'Public Function BranchGraph()
'Dim DB
'Dim Sales As New ADODB.Recordset
'Set DB = New Connection
'DB.CursorLocation = adUseClient
'DB.Open CnnString
'frmChart.MSC.Title = "Branchwise Monthly Sales Trend"
'frmChart.MSC.ColumnCount = 3 '.RecordCount
'Sales.Open "select name,procmonth,sum(actual) as sales from budget as a,unitmaster as b where a.unitcode=b.unitcode and procyear='" & SysYear & "'  group by name,procmonth order by name,procmonth", DB, adOpenStatic
'With Sales
'.MoveFirst
'frmChart.MSC.RowCount = 3
''MSC.row = 1
'frmChart.ChartGrd.Cols = 4 ' .RecordCount + 1
'frmChart.ChartGrd.Rows = 4
'frmChart.ChartGrd.ColWidth(0) = 2200
'For I = 1 To 3
'    frmChart.ChartGrd.TextMatrix(0, I) = Mname(I)
'    frmChart.ChartGrd.ColWidth(I) = 1100
'
'Next
'I = 1
''frmChart.MSC.Column = 3
'Do While Not .EOF
'   frmChart.MSC.row = I
'   unit = !Name
'   frmChart.MSC.RowLabel = !Name
'   frmChart.ChartGrd.TextMatrix(I, 0) = !Name
'   K = 1
'   Do While unit = !Name
'
'      frmChart.MSC.Column = K
'      frmChart.MSC.ColumnLabel = Mname(K)
'      frmChart.MSC.Data = Round(!Sales / 100000, 2)
'      frmChart.ChartGrd.TextMatrix(I, K) = Round(!Sales / 100000, 2)
'      .MoveNext
'      K = K + 1
'      If .EOF Then Exit Do
'   Loop
'   I = I + 1
'   If I > 3 Then Exit Do
'Loop
'End With
'DB.Close
'frmChart.Show 1
'End Function
Public Function WDayName(jd As Date, jmode As Integer) As String
Select Case Weekday(jd)
       Case 1
       WDayName = "Sun"
       Case 2
       WDayName = "Mon"
       Case 3
       WDayName = "Tues"
       Case 4
       WDayName = "Wednes"
       Case 5
       WDayName = "Thurs"
       Case 6
       WDayName = "Fri"
       Case 7
       WDayName = "Satur"
End Select
If jmode = 1 Then
   WDayName = WDayName + "day"
Else
   WDayName = Left(WDayName, 3)
End If
End Function

Public Function PadWithChar(ByVal Jstr As String, ByVal Jlen As Integer, ByVal jpad As String, ByVal PadMode As Integer) As String
       ' PadMode - 0 for left align,1 for right,2 for center
       Dim i As Integer
       If Len(Trim(Jstr)) >= Jlen Then
          PadWithChar = Mid(Trim(Jstr), 1, Jlen)
          Exit Function
       End If
       Select Case PadMode
              Case 0
              PadWithChar = Trim(Jstr) + String(Jlen - Len(Trim(Jstr)), jpad)
              Case 1
              PadWithChar = String(Jlen - Len(Trim(Jstr)), jpad) + Trim(Jstr)
              Case 2
              PadWithChar = String(Int((Jlen - Len(Trim(Jstr))) / 2), jpad) + Trim(Jstr) + String(Jlen - Len(Trim(Jstr)) - Int((Jlen - Len(Trim(Jstr))) / 2), jpad)
              
       End Select
End Function
Public Function FigToWord(ByVal fig As Double) As String
Dim nu, ch, i As Long
Dim Jstr As String
nu = Int(fig)
ch = fig * 100 - nu * 100
FigToWord = ""
Jstr = Trim(str(nu))
i = Len(Jstr)
Do While i > 0
   Select Case i
          Case 1 To 2
          FigToWord = FigToWord + " " + SpellNum(Jstr)
          Exit Do
          Case 3
          FigToWord = FigToWord + " " + SpellNum(Left(Jstr, 1)) + " hundred"
          Jstr = Trim(Val(Mid(Jstr, 2)))
          Case 4 To 5
          FigToWord = FigToWord + " " + SpellNum(Left(Jstr, i - 3)) + " thousand"
          Jstr = Trim(Val(Mid(Jstr, i - 2)))
          Case 6 To 7
          FigToWord = FigToWord + " " + SpellNum(Left(Jstr, i - 5)) + " lakh"
          Jstr = Trim(Val(Mid(Jstr, i - 4)))
          Case 8 To 9
          FigToWord = FigToWord + " " + SpellNum(Left(Jstr, i - 7)) + " crore"
          Jstr = Trim(Val(Mid(Jstr, i - 6)))
   End Select
   i = Len(Jstr)
Loop
If ch > 0 Then
   FigToWord = FigToWord + " and Ch. " + SpellNum(ch)
End If
FigToWord = Trim(FigToWord)
FigToWord = "Nu. " + UCase(Left(FigToWord, 1)) + Mid(FigToWord, 2) + " Only"

End Function
Public Function SpellNum(ByVal Num As Integer) As String
Dim Nm, tens
SpellNum = ""
If Num = 0 Then Exit Function
Nm = Array("", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "forteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
tens = Array("", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
If Num > 19 Then
   SpellNum = tens(Int(Num / 10) - 1) + IIf(Int(Num Mod 10) > 0, Nm(Num Mod 10), "")
Else
   SpellNum = Nm(Num)
End If

End Function
Public Function ProperCase(strText As String) As String 'Capitalize all first characters of the input string 'and set all the others to lowercase
    On Error GoTo Err_ProperCase
   
    Dim i As Integer, intLen As Integer, strTemp As String, strFinal As String
    Dim isSpace As Boolean
    strTemp = LCase(Trim(strText))
    intLen = Len(Trim(strText))
    strFinal = String(intLen, " ")
    isSpace = True
    For i = 1 To intLen
        If Mid(strTemp, i, 1) = Chr(32) Then
            strFinal = Mid(strFinal, 1, i - 1) & Chr(32) & Mid(strFinal, i + 1)
            isSpace = True
        ElseIf isSpace Then
            strFinal = Mid(strFinal, 1, i - 1) & UCase(Mid(strTemp, i, 1)) & Mid(strFinal, i + 1)
            isSpace = False
        Else
            strFinal = Mid(strFinal, 1, i - 1) & Mid(strTemp, i, 1) & Mid(strFinal, i + 1)
        End If
    Next i
    ProperCase = strFinal
   
Exit_ProperCase:
    Exit Function
   
Err_ProperCase:
    MsgBox err.Description & vbCrLf & strText, vbInformation, "ProperCase function"
End Function
'Public Sub NewHash()
'    If CryptCreateHash(m_hProvider, CALG_MD5, 0&, 0&, m_hHash) = 0 Then
'        err.Raise vbObjectError Or &HC332&, _
'                  "MD5Hash", _
'                  "Failed to create CryptoAPI Hash object, system error " _
'                & CStr(err.LastDllError)
'    End If
'End Sub
'Public Sub HashBlock(ByRef Block() As Byte)
'    If CryptHashData(m_hHash, _
'                     Block(LBound(Block)), _
'                     UBound(Block) - LBound(Block) + 1, _
'                     0&) = 0 Then
'        err.Raise vbObjectError Or &HC312&, _
'                  "MD5Hash", _
'                  "Failed to hash data block, system error " _
'                & CStr(err.LastDllError)
'    End If
'End Sub
'Public Function HashBytes(ByRef Block() As Byte) As String
'    NewHash
'    HashBlock Block
'    HashBytes = HashValue()
'End Function
Public Sub FindDZ(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dzname = ""
Set rs = Nothing
rs.Open "select * from tbldzongkhag where dzongkhagcode='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Dzname = rs!DZONGKHAGNAME
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub

MsgBox "INVALID DZONGKHAG CODE " & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub getLocationFromFid(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
locationFromFid = ""
Set rs = Nothing
rs.Open "select * from tblqmsfacility where facilityId='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
locationFromFid = rs!location
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub

MsgBox "Invalid location  " & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Public Sub findMonitorFromFarmer(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
monitorFromFarmer = ""
Set rs = Nothing
rs.Open "select concat(staffcode,' ',staffname) mm from tblmhvstaff where staffcode in(select distinct MONITOR from tblfarmer where IDFARMER='" & dd & "')", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
monitorFromFarmer = rs!mm
'chkred = False
Else
monitorFromFarmer = ""
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindTempStorage(dd As Integer)
On Error GoTo err
Dim rs As New ADODB.Recordset
tempStorageName = ""
mtempStorageName = ""
Set rs = Nothing
rs.Open "select * from tblqmstemporarystorage where storageid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
tempStorageName = rs!dgt
mtempStorageName = rs!storagename
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub

'MsgBox "INVALID DZONGKHAG CODE " & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub


Public Sub FindDepartment(dd As Integer)
On Error GoTo err
Dim rs As New ADODB.Recordset
DeptName = ""
Set rs = Nothing
rs.Open "select * from tbldepartment where deptid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
DeptName = rs!DeptName
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub

MsgBox "INVALID DEPARTMENT CODE " & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Public Sub FindqmsStatus(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsStatus = ""
Set rs = Nothing
rs.Open "select * from tblqmsstatus where statusid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsStatus = rs!status
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsChemical(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsChemical = ""
Set rs = Nothing
rs.Open "select * from tblqmsactiveingredients where ingredientId='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsChemical = Right("0000" & rs!ingredientid, 3) & " " & rs!chemicalname & " " & rs!chemicalformula
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsChemicalTradeName(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsChemicalTradeName = ""
Set rs = Nothing
rs.Open "select * from tblqmschemicalhdr where chemicalid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsChemicalTradeName = rs!chemicalid & " " & rs!tradename
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub

Public Sub FindqmsPlantVariety(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsPlantVariety = ""
Set rs = Nothing
rs.Open "select * from tblqmsplantvariety where varietyid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsPlantVariety = rs!Description
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub Findqmsnuttype(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Mnuttype = ""
Set rs = Nothing
rs.Open "select * from tblqmsnuttype where typeid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Mnuttype = rs!TypeName
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub Findqmstransactiontype(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsTransactionType = ""
Set rs = Nothing
rs.Open "select * from tblqmstransitiontype where transitionid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsTransactionType = Trim(rs!Description)
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsTime(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsTime = ""
Set rs = Nothing
rs.Open "select * from tblqmsshorttime where id='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsTime = Trim(Format(rs!fulltime, "HH:mm:ss"))
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsCloud(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsCloud = ""
Set rs = Nothing
rs.Open "select * from tblqmscloudcover where id='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsCloud = Trim(rs!Description)
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsShade(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsShade = ""
Set rs = Nothing
rs.Open "select * from tblqmsshade where id='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsShade = Trim(rs!Description)
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsWind(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsWind = ""
Set rs = Nothing
rs.Open "select * from tblqmsWind where id='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsWind = Trim(rs!Description)
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub Findqmsverificationtype(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsVerificationType = ""
Set rs = Nothing
rs.Open "select * from tblqmsverificationtype where verificationid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsVerificationType = Trim(rs!Description)
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindqmsPlanttype(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
qmsPlantType = ""
Set rs = Nothing
rs.Open "select * from tblqmsplanttype where planttypeid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsPlantType = rs!Description
'chkred = False
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub Findstatus(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Mstatus = ""
Set rs = Nothing
rs.Open "select * from tblstatus where statusid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Mstatus = rs!desc
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub

MsgBox "INVALID STATUS CODE " & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindGE(dd As String, GG As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
GEname = ""
Set rs = Nothing
rs.Open "select * from tblgewog where dzongkhagID='" & dd & "' AND GEWOGID='" & GG & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
GEname = rs!gewogname
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "INVALID CODE." & "DZONGKHAG=" & dd & " AND GEWOG=" & GG
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findActivity(aa As String)
On Error GoTo err

Dim rs As New ADODB.Recordset
ActName = ""
Set rs = Nothing
rs.Open "SELECT * FROM tbldailyactchoices WHERE name='" & aa & "'", ODKDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
ActName = rs!label
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "INVALID ACTIVITY." & "ACTIVITY=" & aa
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findQmsfacility(aa As String)
On Error GoTo err

Dim rs As New ADODB.Recordset
qmsFacility = ""
Set rs = Nothing
rs.Open "SELECT * FROM tblqmsfacility WHERE facilityid='" & aa & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsFacility = rs!Description

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findQmsApplicationMethod(aa As String)
On Error GoTo err

Dim rs As New ADODB.Recordset
qmsApplicationMethod = ""
Set rs = Nothing
rs.Open "SELECT * FROM tblqmsapplicationmethod WHERE methodid='" & aa & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsApplicationMethod = rs!Description

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findreportname(aa As Integer)
On Error GoTo err

Dim rs As New ADODB.Recordset
ReportName = ""
Set rs = Nothing
rs.Open "SELECT * FROM tblemaillog WHERE id='" & aa & "'", ODKDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
ReportName = rs!ReportName

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findtablename(aa As Integer)
On Error GoTo err

Dim rs As New ADODB.Recordset
odkTableName = ""
Set rs = Nothing
rs.Open "SELECT * FROM tbltable WHERE tblid='" & aa & "'", ODKDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
odkTableName = rs!tblname

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findQmsBatchDetail(aa As Integer)
On Error GoTo err

Dim rs As New ADODB.Recordset
qmsBatchdetail = ""
qmsBatchdetail1 = ""
qmsplantbatch3 = ""
qmsShimentNo = 0

mPlanttype = ""
mPlantVariety = ""

Set rs = Nothing
rs.Open "SELECT * FROM tblqmsplantbatchdetail WHERE plantbatch='" & aa & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
FindqmsPlantVariety rs!plantvariety
FindqmsPlanttype rs!planttype
qmsBatchdetail = rs!plantBatch & "  " & Trim(LTrim(RTrim(qmsPlantType)))
qmsBatchdetail1 = rs!plantBatch & "  " & qmsPlantVariety & "  " & Trim(LTrim(RTrim(qmsPlantType)))
qmsplantbatch3 = qmsPlantVariety '& "  " & Trim(LTrim(RTrim(qmsPlantType)))
mPlanttype = rs!planttype
mPlantVariety = rs!plantvariety

qmsShimentNo = rs!trnid
Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub findQmsBatchReceivedDate(aa As Integer, location As String)
On Error GoTo err
Dim shipmentno As Integer
Dim rs As New ADODB.Recordset
'qmsBatchReceivedDate = ""
qmsbatchTotal = 0
qmsshiptotal = 0

Set rs = Nothing
If location = "LMT" Then
rs.Open "SELECT * FROM tblqmsplantbatchdetail  WHERE plantbatch='" & aa & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsBatchReceivedDate = rs!entrydate
shipmentno = rs!trnid
Else

End If
Else
rs.Open "SELECT MIN( entrydate ) AS entrydate, plantbatch, SUM( debit ) as debit FROM" _
& "`tblqmsplanttransaction` WHERE transactiontype in(9,11) AND SUBSTRING( facilityid, 1, 1 ) =  'T'" _
& " AND debit >0 and plantbatch='" & aa & "' GROUP BY plantbatch order by plantbatch", MHVDB
If rs.EOF <> True Then
qmsBatchReceivedDate = rs!entrydate
findQmsBatchDetail rs!plantBatch
shipmentno = qmsShimentNo
Else

End If


End If

Set rs = Nothing
rs.Open "SELECT sum(shipmentsize) as shipmentsize FROM tblqmsplantbatchdetail  WHERE plantbatch='" & aa & "' and trnid='" & shipmentno & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsbatchTotal = rs!shipmentsize

Else

End If

Set rs = Nothing
rs.Open "SELECT sum(shipmentsize) as shipmentsize FROM tblqmsplantbatchdetail  WHERE  trnid='" & shipmentno & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
qmsshiptotal = rs!shipmentsize
Else

End If

Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindTs(dd As String, GG As String, tt As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
TsName = ""
Set rs = Nothing
rs.Open "select * from tbltshewog where dzongkhagID='" & dd & "' AND GEWOGID='" & GG & "' and tshewogid='" & tt & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
TsName = rs!tshewogname
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "INVALID CODE." & "DZONGKHAG=" & dd & " ,GEWOG=" & GG & " AND TSHOWOG=" & tt
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Public Sub farmertype(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
mFarmerType = ""
Set rs = Nothing
rs.Open "select * from tblplanted where FarmerCode='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
mFarmerType = "Old"
'chkred = False
Else
mFarmerType = ""
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Public Sub FindMONITOR(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
myMonitor = ""
Set rs = Nothing
rs.Open "select distinct concat(staffcode,' ',staffname) st from tblfarmer a join tblmhvstaff b where MONITOR=staffcode and  substring(IDFARMER,1,9)='" & Mid(dd, 1, 9) & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
myMonitor = rs!st
'chkred = False
Else
myMonitor = ""
End If
Exit Sub
err:
MsgBox err.Description
End Sub


Public Sub FindDept(dd As Integer)
On Error GoTo err
Dim rs As New ADODB.Recordset
DeptName = ""
Set rs = Nothing
rs.Open "select * from tbldept where deptid='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
DeptName = rs!DeptName
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "Invalid department."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub getSheet(sheetid As Integer, FileName As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim mystream As ADODB.Stream

dashBoardName = ""
Set rs = Nothing
rs.Open "select * from tbldashbordtrn where trnid='" & sheetid & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
mystream.Write rs!mfile
mystream.SaveToFile "c:\\" & FileName, adSaveCreateOverWrite
mystream.Close
dashBoardName = "c:\\" & FileName
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "Invalid Sheet."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindsTAFF(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
sTAFF = ""
emailAddress = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
sTAFF = rs!staffname
emailAddress = rs!email
'chkred = False
Else
'chkred = True
'If mchk = True Then Exit Sub
'MsgBox "INVALID STAFF CODE=" & dd
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindROLL(dd As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
rOLEnAME = ""

Set rs = Nothing
rs.Open "select * from tblrole where roleid='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
rOLEnAME = rs!roledescription
Else
If mchk = True Then Exit Sub
MsgBox "Record Not Found."
End If
Exit Sub
err:
MsgBox err.Description
End Sub
Public Sub FindFA(ff As String, MTYPE As String)
On Error GoTo err
FAName = ""
Dim rs As New ADODB.Recordset
Set rs = Nothing

If MTYPE = "F" Then
rs.Open "select * from tblfarmer where idfarmer='" & ff & "'", MHVDB
If rs.EOF <> True Then
FAName = rs!farmername
'chkred = False
Else
chkred = True
If mchk = True Then Exit Sub
MsgBox "Record Not Found." & "FARMER CODE=" & ff
End If

ElseIf MTYPE = "A" Then

rs.Open "select * from TBLABSENTEE where ABSENTEEID='" & ff & "'", MHVDB
If rs.EOF <> True Then
FAName = rs!ABSENTEENAME
Else
If mchk = True Then Exit Sub
MsgBox "Record Not Found." & "ABSENTEE CODE=" & ff
End If


Else

End If



Exit Sub
err:
MsgBox err.Description
End Sub
Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = 0
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = err.Description
End Function
Public Function Decipher(ByVal Rep As String) As String
       Decipher = Trim(Right(Rep, 15))
End Function
Public Sub InsertRow(msf As MSFlexGrid, row As Integer)
 
    Dim intRow As Integer
    Dim intCol As Integer
    
    ' Move the rows down
    For intRow = msf.rows - 1 To row Step -1
        For intCol = 0 To msf.cols - 1
            msf.TextMatrix(intRow, intCol) = msf.TextMatrix(intRow - 1, intCol)
        Next
    Next
    
    ' Blank the row
    For intCol = 0 To msf.cols - 1
        msf.TextMatrix(row, intCol) = ""
    Next
    
End Sub
Public Sub updateField()

Dim SQLSTR As String
Dim frcode As String
Dim j As Integer
Dim rs As New ADODB.Recordset

Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                    

db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp SELECT _URI, region_dcode, region_gcode, region,fcode FROM phealthhub15_core where farmerbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = UCase(Mid(rss!dcode, 1, 3))
  mgcode = UCase(Mid(rss!gcode, 1, 3))
  mtcode = UCase(Mid(rss!tcode, 1, 3))
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  Set rsF = Nothing
    db.Execute "update phealthhub15_core set farmerbarcode='" & mfcode & "' where region_dcode='" & rss!dcode & "' and region_gcode='" & rss!gcode & "' and region='" & rss!tcode & "' and fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"
  frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
LogRemarks = "table phealthhub15_core updated successfully.farmerbarcode updated(" & frcode & ")"
updateodklog "no uri", Now, MUSER, LogRemarks, "phealthhub15_core"

frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM phealthhub15_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update phealthhub15_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
  LogRemarks = "table phealthhub15_core updated successfully.staffbarcode updated(" & frcode & ")"
  updateodklog "no uri", Now, MUSER, LogRemarks, "phealthhub15_core"







End Sub
Public Sub updateStorage()

Dim SQLSTR As String

Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
Dim frcode As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                       

db.Execute "delete from mtemp"


SQLSTR = ""
   SQLSTR = "insert into mtemp SELECT _URI, region_dcode, region_gcode, region,fcode FROM storagehub6_core where farmerbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = UCase(Mid(rss!dcode, 1, 3))
  mgcode = UCase(Mid(rss!gcode, 1, 3))
  mtcode = UCase(Mid(rss!tcode, 1, 3))
  mfcode = "000000" & CStr(rss!fcode)
  mfcode = "F" & Right(mfcode, 4)
  mfcode = mdcode & mgcode & mtcode & mfcode
  Set rsF = Nothing
  
  db.Execute "update storagehub6_core set farmerbarcode='" & mfcode & "' where region_dcode='" & rss!dcode & "' and region_gcode='" & rss!gcode & "' and region='" & rss!tcode & "' and fcode='" & rss!fcode & "' and  _URI='" & rss![_uri] & "'"

frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
LogRemarks = "table storagehub6_core updated successfully.farmerbarcode updated(" & frcode & ")"
updateodklog "no uri", Now, MUSER, LogRemarks, "storagehub6_core"


frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM storagehub6_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update storagehub6_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
  LogRemarks = "table storagehub6_core updated successfully.staffbarcode updated(" & frcode & ")"
  updateodklog "no uri", Now, MUSER, LogRemarks, "storagehub6_core"


                        
                        
End Sub
Public Sub updateDailyact()

Dim SQLSTR As String

Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
Dim frcode As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                       

db.Execute "delete from mtemp"



frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM dailyacthub9_core where staffbarcode=''"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update dailyacthub9_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
  LogRemarks = "table dailyacthub9_core updated successfully.staffbarcode updated(" & frcode & ")"
  updateodklog "no uri", Now, MUSER, LogRemarks, "dailyacthub9_core"


                        
                        
End Sub
Public Sub updatesiatribution()

Dim SQLSTR As String

Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
Dim frcode As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                       

db.Execute "delete from mtemp"



frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM distribution4_core where staffbarcode='' or staffbarcode is null"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update distribution4_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
  LogRemarks = "table distribution4_core updated successfully.staffbarcode updated(" & frcode & ")"
  updateodklog "no uri", Now, MUSER, LogRemarks, "distribution4_core"


                        
                        
End Sub
Public Sub firstMonday(nextdate As Date)

Dim intYear As Integer
Dim intMonth As Integer
Dim intDay As Integer
Dim i As Integer
Dim mydt As Date
i = 1
mydt = Format(Now, "dd/MM/yyyy")
intYear = CInt(Year(mydt))
intMonth = Month(mydt) + 1
intDay = i
For i = 1 To 7
If Mid(UCase(Format(DateSerial(intYear, intMonth, i), "dddd, mmm d yyyy")), 1, 3) = "MON" Then
firstMondayOfTheMonth = DateSerial(intYear, intMonth, i)
Exit For
End If


Next

End Sub
Public Sub updateemaillog(excel_app As Object, receipient_id As Integer, nextemaildate As Date, frequency As Integer)
Dim mFilepath As String
Dim emailmessage As String
Dim EMAILIDS As String
Dim newDate As Date
Dim rs As New ADODB.Recordset
Dim i As Integer

Set rs = Nothing
rs.Open "select * from tblemaillog where id='" & receipient_id & "'", ODKDB

EMAILIDS = rs!receipients
emailmessage = rs!ReportName
mFilepath = App.Path & "\reportlog\" & Format(Now, "yyMMdd") & " " & rs!ReportName & ".xls"
        excel_app.ActiveWorkbook.SaveCopyAs FileName:=mFilepath
        excel_app.ActiveWorkbook.Close False
        excel_app.DisplayAlerts = False
        excel_app.Quit
     
Screen.MousePointer = vbDefault

Dim mCONNECTION As String
Dim retVal          As String
mCONNECTION = "smtp.tashicell.com"
'mCONNECTION = "smtp1.btl.bt"
 retVal = SendMail(EMAILIDS, rs!Subject, "noreply@mhv.com", _
          emailmessage, mCONNECTION, 25, _
          "habizabi", "habizabi", _
           mFilepath, CBool(False))


If frequency = 30 Then
firstMonday Format(Now, "yyyy-MM-dd")
newDate = firstMondayOfTheMonth
Else
newDate = Now + frequency

End If

If retVal = "ok" Then
ODKDB.Execute "update tblemaillogtrn set status='1' where receipient_id='" & receipient_id & "' and status='0'"
ODKDB.Execute "insert  tblemaillogtrn(receipient_id,lastemailsentdate,nextemaildate,status) " _
& "values('" & receipient_id & "','" & Format(nextemaildate, "yyyy-MM-dd") & "','" & Format(newDate, "yyyy-MM-dd") & "','0')"

'frmemaillogNew.Label1.Visible = False
'frmemaillogNew.txterror.Visible = False
'frmemaillogNew.txterror.Text = ""
Else
'MsgBox "Please Check Internet Connection " & retVal
'frmemaillogNew.Label1.Visible = True
'frmemaillogNew.txterror.Visible = True
'frmemaillogNew.txterror.Text = retVal
ODKDB.Execute "update tblemaillogtrn set status='1' where receipient_id='" & receipient_id & "' and status='0'"
ODKDB.Execute "insert  tblemaillogtrn(receipient_id,lastemailsentdate,nextemaildate,status) " _
& "values('" & receipient_id & "','" & Format(nextemaildate, "yyyy-MM-dd") & "','" & Format(newDate, "yyyy-MM-dd") & "','0')"

End If

End Sub
Public Sub findfieldindes(tblid As Integer, tblname As String, fieldname As String)
fieldIndex = 0
Dim fcount, j As Integer
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & tblid & "'", ODKDB

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount) + 1

Set rs = Nothing
rs.Open "SELECT * FROM " & LCase(tblname) & " ", ODKDB
For j = 0 To fcount - 1


If UCase(rs.Fields(j).name) = UCase(fieldname) Then
fieldIndex = j
Exit For
End If
Next
End Sub

Public Sub findParamDetails(paraId As Integer)
 paramName = ""
 emailId = ""
 acceptableThresholdValue = ""
 paramValue = 0
 paramDesc = ""
 paramFieldStorage = ""
 ispercentage = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblodkalarmparameter where status='ON' and paraid='" & paraId & "'", ODKDB
If rs.EOF <> True Then
paramName = rs!paraname
acceptableThresholdValue = rs!Value
emailId = rs!receipents
findreportname rs!REPORT
paramValue = rs!Value
paramDesc = rs!Description
paramFieldStorage = rs!fstype
If rs!percentage = 1 Then
ispercentage = True
Else
ispercentage = False
End If
End If

End Sub

Public Sub updatefollowuplog(paraId As Integer, uri As String, tblname As String, fieldname As String, entrydate As Date, odkStartDate As Date, odkValue As Double, staffcode As String, farmercode As String, fieldcode As Integer)
Dim MaxId As Integer
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "Select max(trnid) as max from tblodkfollowuplog", ODKDB
If rs.EOF <> True Then
MaxId = IIf(IsNull(rs!max), 1, rs!max + 1)
End If
odkValue = Format(odkValue, "####0.00")
Set rs = Nothing
rs.Open "select * from tblodkfollowuplog where URI='" & uri & "' and tblname='" & tblname & "' and fieldname='" & fieldname & "' and odkvalue='" & odkValue & "'", ODKDB
If rs.EOF <> True Then

Else

ODKDB.Execute "insert into tblodkfollowuplog(trnid,uri,tblname,fieldname,paraid," _
            & "entrydate,odkstartdate,odkvalue,actiontaken," _
            & "recommendation,requireNextFollowUp,emailstatus,followupstatus,staffcode,farmercode,fieldcode)" _
            & " values(" _
            & "'" & MaxId & "'," _
            & "'" & uri & "'," _
            & "'" & tblname & "'," _
            & "'" & fieldname & "'," _
            & "'" & paraId & "'," _
            & "'" & Format(entrydate, "yyyy-MM-dd") & "'," _
            & "'" & Format(odkStartDate, "yyyy-MM-dd") & "'," _
            & "'" & odkValue & "'," _
            & "''," _
            & "''," _
            & "''," _
            & "'ON'," _
            & "'ON'," _
            & "'" & staffcode & "'," _
            & "'" & farmercode & "'," _
            & "'" & fieldcode & "'" _
            & ")"
    End If
End Sub
Public Sub updateregistration()

Dim SQLSTR As String

Dim j As Integer
Dim rs As New ADODB.Recordset
Dim rss As New ADODB.Recordset
Dim rsF As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Dim db As New ADODB.Connection
Set rsadd = Nothing
Dim frcode As String
j = 0
Dim mdcode, mgcode, mtcode, mfcode As String
Dim mdcode1, mgcode1, mtcode1, mfcode1 As Integer
mdcode = ""
mgcode = ""
mtcode = ""
mfcode = ""
frcode = ""
mdcode1 = 0
mgcode1 = 0
mtcode1 = 0
mfcode1 = 0
Dim tempstr As String
SQLSTR = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient

db.Open OdkCnnString
                       

db.Execute "delete from mtemp"



frcode = ""
db.Execute "delete from mtemp"
SQLSTR = ""
   SQLSTR = "insert into mtemp(_URI,dcode) SELECT _URI, staffid FROM farmer_registration4_core where staffbarcode='' or staffbarcode is null"
  db.Execute SQLSTR
  Set rss = Nothing
  
 rss.Open "select * from mtemp", db
  Do While rss.EOF <> True
   
  mdcode = "S0" & rss!dcode

  mfcode = mdcode
  Set rsF = Nothing
  
  db.Execute "update farmer_registration4_core set staffbarcode='" & mfcode & "' where staffid='" & rss!dcode & "'  and  _URI='" & rss![_uri] & "'"
 frcode = frcode & mfcode & ","
  rss.MoveNext
  Loop
  LogRemarks = "table farmer_registration4_core updated successfully.staffbarcode updated(" & frcode & ")"
  updateodklog "no uri", Now, MUSER, LogRemarks, "farmer_registration4_core"


                        
                        
End Sub
Public Function validatecombo(tblname As String, id As String, keyfield As String) As Boolean
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from " & tblname & " where " & id & " = " & "'" & keyfield & "'", MHVDB
If rs.EOF <> True Then
validatecombo = True
Else
validatecombo = False
End If

End Function

Public Function finddescription(tblname As String, id As String, keyfield As String, returnfield As String) As String
Dim rs As New ADODB.Recordset
Set rs = Nothing
mDescription = ""
rs.Open "select " & keyfield & "," & returnfield & " from " & tblname & " where " & keyfield & " = " & "'" & id & "'", MHVDB
If rs.EOF <> True Then
mDescription = rs.Fields(0) & "  " & rs.Fields(1)
Else
mDescription = ""
End If
finddescription = mDescription
End Function

Public Function IsValidEmail(email As String) As Boolean
Dim myAt As Integer
Dim myAtLastPos As Integer
Dim myDot As Integer
Dim myDotDot As Integer
Dim myDotAt As Integer
Dim myAtDot As Integer
Dim mySpace As Integer
IsValidEmail = True
mySpace = InStr(1, email, " ", vbTextCompare)
myAtLastPos = InStrRev(email, "@", , vbTextCompare)
myAt = InStr(1, email, "@", vbTextCompare)
myAtDot = InStr(1, email, "@.", vbTextCompare)
myDotAt = InStr(1, email, ".@", vbTextCompare)
myDot = InStr(myAt + 2, email, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
If myAtDot > 0 Or myDotAt > 0 Or myAtLastPos <> myAt Or mySpace > 0 Or myAt = 0 Or myDot = 0 Or myDotDot > 0 Or Right(email, 1) = "." Then IsValidEmail = False
End Function

Public Sub dijkstra_shortest_Path()
        Const infinity As Integer = 15000   'number that is larger than max distance
        Dim i As Integer                    'loop counter
        Dim j As Integer                    'loop counter
        Dim sourceP As Integer              'point to determine distance to from all nodes
        Dim inputData As String             'temp variable to ensure good data enterred
        Dim graph(nodeCount) As Vertex      'all inforamtion for each point (see Vertex declaration above)
        Dim nextP As Integer                'closest point that is not dead
        Dim min As Integer                  'distance of closest point not dead
        Dim outString As String             'string to display the output
        Dim goodSource As Boolean
 
        'user enters source point data and ensured that it is correct
        Do
            goodSource = True
            inputData = (InputBox("What is the source point: ", "Source Point between: 0 & " & nodeCount))
            If IsNumeric(inputData) Then
                sourceP = CInt(inputData)
                If sourceP > nodeCount Or sourceP < 0 Then
                    MsgBox "Source point must be between 0 & " & nodeCount & "."
                    goodSource = False
                End If
            Else
                MsgBox "Source point must be numeric and be between 0 & " & nodeCount & "."
                goodSource = False
            End If
        Loop While Not goodSource
        'get data so we can analyze the distances
        Call populateGraph(graph)
 
        'set default values to not dead and distances to infinity (unless distance is to itself)
        For i = 0 To nodeCount
            If graph(i).name = sourceP Then
                graph(i).distance = 0
                graph(i).isDead = False
            Else:
                graph(i).distance = infinity
                graph(i).isDead = False
            End If
        Next i
 
        For i = 0 To nodeCount
            min = infinity + 1
            'determine closest point that is not dead
            For j = 0 To nodeCount
                If Not graph(j).isDead And graph(j).distance < min Then
                    nextP = j
                    min = graph(j).distance
                End If
            Next j
            'calculate distances from the closest point & to all of its connections
            For j = 0 To graph(nextP).numConnect
                If graph(graph(nextP).connections(j).destination).distance > graph(nextP).distance + graph(nextP).connections(j).weight Then
                    graph(graph(nextP).connections(j).destination).distance = graph(nextP).distance + graph(nextP).connections(j).weight
                End If
            Next j
            'kill the value we just looked at so we can get the next point
            graph(nextP).isDead = True
        Next i
 
        'display the distance from the source point to all other points
        outString = ""
        For i = 0 To nodeCount
            outString = outString & "The distance between nodes " & sourceP & " and " & i & " is " & graph(i).distance & vbCrLf
        Next i
        MsgBox outString
    End Sub
    
    Private Sub populateGraph(vertexMatrix() As Vertex)
        'get data into graph matrix to determine distance from all points
        Dim i As Integer
        Dim j As Integer
 
        '0 connections
        vertexMatrix(0).name = 0
        vertexMatrix(0).numConnect = 3
        vertexMatrix(0).connections(0).destination = 1
        vertexMatrix(0).connections(1).destination = 2
        vertexMatrix(0).connections(2).destination = 6
        vertexMatrix(0).connections(3).destination = 7
        vertexMatrix(0).connections(0).weight = 10
        vertexMatrix(0).connections(1).weight = 15
        vertexMatrix(0).connections(2).weight = 30
        vertexMatrix(0).connections(3).weight = 50
 
        '1 connections
        vertexMatrix(1).name = 1
        vertexMatrix(1).numConnect = 3
        vertexMatrix(1).connections(0).destination = 0
        vertexMatrix(1).connections(1).destination = 3
        vertexMatrix(1).connections(2).destination = 4
        vertexMatrix(1).connections(3).destination = 9
        vertexMatrix(1).connections(0).weight = 10
        vertexMatrix(1).connections(1).weight = 16
        vertexMatrix(1).connections(2).weight = 5
        vertexMatrix(1).connections(3).weight = 40
 
        '2 connections
        vertexMatrix(2).name = 2
        vertexMatrix(2).numConnect = 3
        vertexMatrix(2).connections(0).destination = 0
        vertexMatrix(2).connections(1).destination = 7
        vertexMatrix(2).connections(2).destination = 8
        vertexMatrix(2).connections(3).destination = 9
        vertexMatrix(2).connections(0).weight = 15
        vertexMatrix(2).connections(1).weight = 33
        vertexMatrix(2).connections(2).weight = 18
        vertexMatrix(2).connections(3).weight = 60
 
        '3 connections
        vertexMatrix(3).name = 3
        vertexMatrix(3).numConnect = 1
        vertexMatrix(3).connections(0).destination = 1
        vertexMatrix(3).connections(1).destination = 4
        vertexMatrix(3).connections(0).weight = 16
        vertexMatrix(3).connections(1).weight = 22
 
        '4 connections
        vertexMatrix(4).name = 4
        vertexMatrix(4).numConnect = 2
        vertexMatrix(4).connections(0).destination = 1
        vertexMatrix(4).connections(1).destination = 3
        vertexMatrix(4).connections(2).destination = 5
        vertexMatrix(4).connections(0).weight = 5
        vertexMatrix(4).connections(1).weight = 22
        vertexMatrix(4).connections(2).weight = 30
 
        '5 connections
        vertexMatrix(5).name = 5
        vertexMatrix(5).numConnect = 0
        vertexMatrix(5).connections(0).destination = 4
        vertexMatrix(5).connections(0).weight = 30
 
        '6 connections
        vertexMatrix(6).name = 6
        vertexMatrix(6).numConnect = 1
        vertexMatrix(6).connections(0).destination = 0
        vertexMatrix(6).connections(1).destination = 7
        vertexMatrix(6).connections(0).weight = 30
        vertexMatrix(6).connections(1).weight = 40
 
        '7 connections
        vertexMatrix(7).name = 7
        vertexMatrix(7).numConnect = 3
        vertexMatrix(7).connections(0).destination = 0
        vertexMatrix(7).connections(1).destination = 2
        vertexMatrix(7).connections(2).destination = 8
        vertexMatrix(7).connections(3).destination = 6
        vertexMatrix(7).connections(0).weight = 50
        vertexMatrix(7).connections(1).weight = 33
        vertexMatrix(7).connections(2).weight = 3
        vertexMatrix(7).connections(3).weight = 40
 
        '8 connections
       vertexMatrix(8).name = 8
       vertexMatrix(8).numConnect = 1
       vertexMatrix(8).connections(0).destination = 2
       vertexMatrix(8).connections(1).destination = 7
       vertexMatrix(8).connections(0).weight = 18
       vertexMatrix(8).connections(1).weight = 3
 
        '9 connections
       vertexMatrix(9).name = 9
       vertexMatrix(9).numConnect = 1
       vertexMatrix(9).connections(0).destination = 1
       vertexMatrix(9).connections(1).destination = 2
       vertexMatrix(9).connections(0).weight = 40
       vertexMatrix(9).connections(1).weight = 60
    End Sub
    
    


