VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmodkcollect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ODK RECORD MODIFICATION"
   ClientHeight    =   9030
   ClientLeft      =   2190
   ClientTop       =   825
   ClientWidth     =   17640
   Icon            =   "frmodkcollect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   17640
   Begin VB.TextBox txtmod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   14520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox CBOSTATUS 
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
      ItemData        =   "frmodkcollect.frx":0E42
      Left            =   13440
      List            =   "frmodkcollect.frx":0E44
      TabIndex        =   13
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
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
      Left            =   16320
      Picture         =   "frmodkcollect.frx":0E46
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Enabled         =   0   'False
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
      Left            =   15120
      Picture         =   "frmodkcollect.frx":1B10
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   1215
   End
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
      ItemData        =   "frmodkcollect.frx":22BA
      Left            =   13440
      List            =   "frmodkcollect.frx":22BC
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker TXTFRMDATE 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   81854465
      CurrentDate     =   41313
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOAD"
      Enabled         =   0   'False
      Height          =   735
      Left            =   16080
      Picture         =   "frmodkcollect.frx":22BE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   17415
      _cx             =   30718
      _cy             =   13150
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmodkcollect.frx":2700
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSDataListLib.DataCombo CBOTBL 
      Bindings        =   "frmodkcollect.frx":2729
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker TXTTODATE 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   81854465
      CurrentDate     =   41313
   End
   Begin MSDataListLib.DataCombo CBOMONITOR 
      Bindings        =   "frmodkcollect.frx":273E
      Height          =   315
      Left            =   7320
      TabIndex        =   17
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "MONITOR"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
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
      Left            =   2520
      TabIndex        =   15
      Top             =   8520
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL RECORD"
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
      Left            =   840
      TabIndex        =   14
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "STATUS"
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
      Left            =   11880
      TabIndex        =   12
      Top             =   600
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DATE CRITERIA"
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
      Left            =   11880
      TabIndex        =   8
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TABLE SELECTION"
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
      Left            =   5520
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TO DATE"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FROM DATE"
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
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmodkcollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CBODATE_GotFocus()

Command1.Enabled = True

End Sub

Private Sub CBOTBL_LostFocus()
Dim i, j, fcount As Integer
Operation = ""
Mygrid.Visible = False
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection

db.Open OdkCnnString
                        
If Len(CBOTBL.Text) = 0 Then Exit Sub


Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & CBOTBL.BoundText & "' ", db

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount)
CBODATE.Clear
Set rs = Nothing
rs.Open "SELECT * FROM " & LCase(CBOTBL.Text) & " where 1", db
For j = 0 To fcount - 1
If rs.Fields(j).Type = 135 Then
CBODATE.AddItem rs.Fields(j).Name
End If
Next

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Command1_Click()
Mygrid.Visible = True

If txttodate.Value < txtfrmdate.Value Then
MsgBox "INVALID DATE SELECTION."
Mygrid.Clear
Mygrid.Visible = False
Exit Sub
End If

FillGrid
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
On Error GoTo err
Dim maxlog As Integer
Dim CONNLOCAL As New ADODB.Connection
Dim i, j, fcount As Integer

CONNLOCAL.Open OdkCnnString
                      
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & CBOTBL.BoundText & "'", CONNLOCAL
fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount) + 1
CONNLOCAL.BeginTrans
If Len(CBOMONITOR.Text) <> 0 Then
CONNLOCAL.Execute "delete FROM " & LCase(CBOTBL.Text) & " WHERE staffbarcode='" & CBOMONITOR.BoundText & "' and substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' "
Else
CONNLOCAL.Execute "delete FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' "
End If
Set rs = Nothing
rs.Open "select * from " & LCase(CBOTBL.Text) & " where 1", CONNLOCAL
For j = 1 To Mygrid.Rows - 1

                For i = 0 To fcount - 1
                               
                               
                                        If rs.Fields(i).Type = 200 Then
                                                MTEMPVAR = ValidateString(Mygrid.TextMatrix(j, i))
                                        ElseIf rs.Fields(i).Type = 135 Then
                                                MTEMPVAR = Format(Mygrid.TextMatrix(j, i), "yyyy-MM-dd hh:mm:ss")
                                        Else
                                                MTEMPVAR = Mygrid.TextMatrix(j, i)
                                        End If
                                        
                                        MYSQLSTR = MYSQLSTR + "'" + Trim(Mid(MTEMPVAR, InStr(1, MTEMPVAR, "|") + 1)) + "',"
                Next
                                
                                MYSQLSTR = "(" + Mid(MYSQLSTR, 1, Len(MYSQLSTR) - 1) + ")"
                                CONNLOCAL.Execute "INSERT INTO " & LCase(CBOTBL.Text) & "  VALUES " + MYSQLSTR
                                MYSQLSTR = ""
                                
                                
                                If Len(Mygrid.TextMatrix(j, Mygrid.Cols - 1)) <> 0 Then
                                CONNLOCAL.Execute "insert into tblodkmodificationlog (_URI,date_of_modification,modified_by,description,table_name)values('" & Mygrid.TextMatrix(j, 0) & "','" & Format(Now, "yyyy-MM-dd") & "','" & MUSER & "','" & Mygrid.TextMatrix(j, Mygrid.Cols - 1) & "','" & CBOTBL.Text & "')"
                                End If
                                
                                
                                
Next





txtmod.Text = ""
CONNLOCAL.CommitTrans
MsgBox "RECORD SUCCESSFULLY SAVED."
Exit Sub
err:
CONNLOCAL.RollbackTrans
  MsgBox err.Description

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
Label7.Caption = ""
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")
Operation = ""
Mygrid.Visible = False
Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString
                        
'odk_prodLocal
Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select *  from tbltable where status='ON' order by tblid", db
Set CBOTBL.RowSource = rs
CBOTBL.ListField = "TBLNAME"
CBOTBL.BoundColumn = "TBLID"


Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from MHV.tblmhvstaff where moniter='1' order by staffcode", db
Set CBOMONITOR.RowSource = rs
CBOMONITOR.ListField = "staffname"
CBOMONITOR.BoundColumn = "staffcode"




 'Mygrid.Cell(flexcpForeColor, i, i, i, 20) = vbRed
cbostatus.AddItem "ALL"
cbostatus.AddItem "GOOD"
cbostatus.AddItem "BAD"
cbostatus.Text = "GOOD"
txtmod.Text = ""
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub FillGrid()
On Error GoTo err
Dim isStaffBarCOde As Boolean
Dim RSCNT As New ADODB.Recordset
Label7.Caption = ""
Dim SQLSTR As String
Mygrid.Clear
isStaffBarCOde = False
Dim CONNLOCAL As New ADODB.Connection
Dim i, j, fcount As Integer
CONNLOCAL.Open OdkCnnString
                       
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & CBOTBL.BoundText & "'", CONNLOCAL

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount) + 1
Mygrid.Cols = 1
Set rs = Nothing
rs.Open "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "'", CONNLOCAL
For j = 0 To fcount - 1
Mygrid.Cols = Mygrid.Cols + 1
Mygrid.TextMatrix(0, j) = rs.Fields(j).Name
If UCase(rs.Fields(j).Name) = "STAFFBARCODE" Then
isStaffBarCOde = True
End If
Next
If isStaffBarCOde = True Then
CBOMONITOR.Visible = True
Else
CBOMONITOR.Visible = False
End If
Mygrid.Rows = 1
Set rs = Nothing
Set RSCNT = Nothing
If cbostatus.Text = "ALL" Then
If Len(CBOMONITOR.Text) <> 0 And isStaffBarCOde = True Then
SQLSTR = "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and staffbarcode='" & CBOMONITOR.BoundText & "' order by " & CBODATE.Text & " "
RSCNT.Open "SELECT COUNT(*) AS CNT FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and staffbarcode='" & CBOMONITOR.BoundText & "' order by " & CBODATE.Text & " ", CONNLOCAL

Else
SQLSTR = "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by " & CBODATE.Text & " "
RSCNT.Open "SELECT COUNT(*) AS CNT FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by " & CBODATE.Text & " ", CONNLOCAL
End If




ElseIf cbostatus.Text = "GOOD" Then
If Len(CBOMONITOR.Text) <> 0 And isStaffBarCOde = True Then
SQLSTR = "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' AND TRIM(STATUS)<>TRIM('BAD') and staffbarcode='" & CBOMONITOR.BoundText & "' order by " & CBODATE.Text & " "
RSCNT.Open "SELECT COUNT(*) AS CNT FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' AND TRIM(STATUS)<>TRIM('BAD') and staffbarcode='" & CBOMONITOR.BoundText & "' order by " & CBODATE.Text & " ", CONNLOCAL

Else
SQLSTR = "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' AND TRIM(STATUS)<>TRIM('BAD') order by " & CBODATE.Text & " "
RSCNT.Open "SELECT COUNT(*) AS CNT FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' AND TRIM(STATUS)<>TRIM('BAD') order by " & CBODATE.Text & " ", CONNLOCAL


End If
ElseIf cbostatus.Text = "BAD" Then

SQLSTR = "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' AND TRIM(STATUS)=TRIM('BAD') order by " & CBODATE.Text & " "
RSCNT.Open "SELECT COUNT(*) AS CNT FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' AND TRIM(STATUS)=TRIM('BAD') order by " & CBODATE.Text & " ", CONNLOCAL

Else
MsgBox "INVALID SELECTION OF STATUS."
Exit Sub
End If



'SQLSTR = "SELECT * FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(TXTFRMDATE.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' order by " & CBODATE.Text & " "

rs.Open SQLSTR, CONNLOCAL
Label7.Caption = IIf(IsNull(RSCNT!cnt), 0, RSCNT!cnt)
i = 1
If rs.EOF <> True Then
Mygrid.Visible = True
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1

For j = 0 To fcount - 1

Mygrid.TextMatrix(i, j) = IIf(IsNull(rs.Fields(j).Value), "", rs.Fields(j).Value)
Mygrid.ColWidth(j) = 2000 + Len(IIf(IsNull(rs.Fields(j).Value), 0, rs.Fields(j).Value))
Next


rs.MoveNext
i = i + 1
Loop

Command2.Enabled = True

Else
Mygrid.Visible = False
MsgBox "No Record Found."
Command2.Enabled = False
End If
Exit Sub
err:
MsgBox "PROBLEM WITH DATE SELECTION, PLEASE SELECT AGAIN."
End Sub

Private Sub Mygrid_Click()
If Mygrid.col <> 0 Then
Mygrid.Editable = flexEDKbdMouse
Else

Mygrid.Editable = flexEDNone
End If

If Mygrid.col = Mygrid.Cols - 2 Then
Mygrid.ComboList = "    |" & "BAD"
Else

Mygrid.ComboList = ""
End If

End Sub

Private Sub Mygrid_DblClick()
If Mygrid.col <> 0 Then
Mygrid.ColWidth(Mygrid.col) = 0
End If
End Sub

Private Sub Mygrid_LeaveCell()
'MsgBox "row=" & Mygrid.row & " col=" & Mygrid.Col
Dim m As Integer
Dim modstr As String
modstr = ""
On Error GoTo err
Command2.Enabled = True
Dim CONNLOCAL As New ADODB.Connection
Dim i, j, fcount As Integer
CONNLOCAL.Open OdkCnnString
                        
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tbltable where tblid='" & CBOTBL.BoundText & "'", CONNLOCAL
fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount) + 1
CONNLOCAL.BeginTrans
'CONNLOCAL.Execute "delete FROM " & LCase(CBOTBL.Text) & " WHERE substring(" & CBODATE.Text & ",1,10)>='" & Format(TXTFRMDATE.Value, "yyyy-MM-dd") & "' AND substring(" & CBODATE.Text & ",1,10)<='" & Format(TXTTODATE.Value, "yyyy-MM-dd") & "' order by " & CBODATE.Text & "  "
Set rs = Nothing
rs.Open "select * from " & LCase(CBOTBL.Text) & " where _URI='" & Mygrid.TextMatrix(Mygrid.row, 0) & "'", CONNLOCAL
If rs.EOF <> True Then


If rs.Fields(Mygrid.col) <> Mygrid.TextMatrix(Mygrid.row, Mygrid.col) Then
modstr = ""

For m = 1 To Mygrid.Cols - 2
If rs.Fields(m) <> Mygrid.TextMatrix(Mygrid.row, m) Then
modstr = modstr & " " & "Value of Field  " & Mygrid.TextMatrix(0, m) & "  changes from   " & rs.Fields(m) & "  to  " & Mygrid.TextMatrix(Mygrid.row, m) & "  of table  " & CBOTBL.Text & ",  "
End If
Next

If Len(modstr) > 0 Then
   modstr = Left(modstr, Len(modstr) - 3)
   End If
Mygrid.TextMatrix(Mygrid.row, Mygrid.Cols - 1) = modstr
End If





If rs.Fields(Mygrid.col) = Mygrid.TextMatrix(Mygrid.row, Mygrid.col) Then
modstr = ""

For m = 1 To Mygrid.Cols - 2
If rs.Fields(m) <> Mygrid.TextMatrix(Mygrid.row, m) Then
modstr = modstr & " " & "Value of Field  " & Mygrid.TextMatrix(0, m) & " changes from   " & rs.Fields(m) & "  to  " & Mygrid.TextMatrix(Mygrid.row, m) & "  of table  " & CBOTBL.Text & ",  "
End If
Next

If Len(modstr) > 0 Then
   modstr = Left(modstr, Len(modstr) - 3)
   End If
Mygrid.TextMatrix(Mygrid.row, Mygrid.Cols - 1) = modstr
End If




End If


txtmod.Text = modstr








Exit Sub
err:
CONNLOCAL.RollbackTrans
  MsgBox err.Description

End Sub

Private Sub Mygrid_ValidateEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
Command2.Enabled = False
End Sub
