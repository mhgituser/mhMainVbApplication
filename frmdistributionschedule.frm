VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmdistributionschedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DISTRIBUTION SCHEDULE"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   825
   ClientWidth     =   20400
   Icon            =   "frmdistributionschedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   20400
   Begin VB.TextBox txtmaxdistno 
      Height          =   615
      Left            =   1200
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chknick 
      Caption         =   "Check1"
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   16920
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   16800
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Modify/View Existing List"
      Height          =   615
      Left            =   6360
      Picture         =   "frmdistributionschedule.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add New List"
      Height          =   615
      Left            =   6360
      Picture         =   "frmdistributionschedule.frx":11CC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   8520
      TabIndex        =   13
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton Command10 
         Caption         =   "Finalize Schedule"
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
         Left            =   9600
         Picture         =   "frmdistributionschedule.frx":1556
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkunplanned 
         Caption         =   "Unplanned Distribution"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   37
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   3600
         Picture         =   "frmdistributionschedule.frx":1CC0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   4455
      End
      Begin VB.ComboBox cboyear 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmdistributionschedule.frx":246A
         Left            =   4920
         List            =   "frmdistributionschedule.frx":2483
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cbomnth 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmdistributionschedule.frx":24B1
         Left            =   6840
         List            =   "frmdistributionschedule.frx":24D9
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdload 
         Caption         =   "Load"
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
         Left            =   8280
         Picture         =   "frmdistributionschedule.frx":2519
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "frmdistributionschedule.frx":2DE3
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
         TabIndex        =   23
         Top             =   960
         Width           =   1995
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
         TabIndex        =   19
         Top             =   360
         Width           =   540
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
         TabIndex        =   18
         Top             =   360
         Width           =   405
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
         TabIndex        =   17
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.TextBox txtdno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtindecator 
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chksel 
      Caption         =   "SHOW DZONGKHAG  SELECTION PANEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   9
      Top             =   9000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10080
      Picture         =   "frmdistributionschedule.frx":2DF8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
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
      Left            =   8400
      Picture         =   "frmdistributionschedule.frx":3AC2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   9855
      Begin VB.CheckBox chkrefill 
         Caption         =   "Refillin"
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
         Left            =   1920
         TabIndex        =   40
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CheckBox chkPolinizer 
         Caption         =   "Polinizer"
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
         Left            =   1920
         TabIndex        =   39
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Select All"
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
         Left            =   8520
         TabIndex        =   36
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtcratecnt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtfcode 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4680
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chkcf 
         Caption         =   "CF"
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
         Left            =   1920
         TabIndex        =   32
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CheckBox chkgrf 
         Caption         =   "GRF"
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
         Left            =   1920
         TabIndex        =   31
         Top             =   4080
         Width           =   1455
      End
      Begin VB.ListBox LSTPR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         ItemData        =   "frmdistributionschedule.frx":426C
         Left            =   3600
         List            =   "frmdistributionschedule.frx":426E
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   600
         Width           =   6135
      End
      Begin VB.CheckBox chkpriority 
         Caption         =   "Priority List"
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
         Left            =   1920
         TabIndex        =   28
         Top             =   3840
         Width           =   1455
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
         Height          =   3210
         Left            =   0
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   600
         Width           =   3375
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
         Left            =   1680
         Picture         =   "frmdistributionschedule.frx":4270
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   120
         Picture         =   "frmdistributionschedule.frx":4F3A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderWidth     =   10
         X1              =   3480
         X2              =   3480
         Y1              =   120
         Y2              =   4560
      End
      Begin VB.Label Label4 
         Caption         =   "DZONGKHAG SELECTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Mygrid 
      Height          =   7260
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   12806
      _Version        =   393216
      Rows            =   5
      Cols            =   33
      RowHeightMin    =   400
      ForeColorFixed  =   -2147483635
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   3
      Appearance      =   0
      FormatString    =   $"frmdistributionschedule.frx":56A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmdistributionschedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngResult As Long
Dim ButtonDown As Boolean
Dim RowToMove As Integer
Dim DestRow As Integer
Dim CurrRow, ErrCTR As Long
Dim subtot, alltot, subtotplant, alltotplant As Double
Dim subtotcrateno, alltotcrateno, subtotbcrate, alltotbcrate As Double
Dim subtotecrate, alltotecrate, subtotbno, alltotbno, subtotplno, alltotplno, subtotcrate, alltotcrate As Double
Dim subtotssp, alltotssp, subtotmop, alltotmop, subtoturea, alltoturea, subtotdolomite, alltotdolomite As Double
Dim subtotkg1, alltotkg1, subtotamtnu1, alltotamtnu1 As Double
Dim subtotkg, alltotkg As Double
Dim subtotamtnu2, alltotamtnu2 As Double
Dim subtottotamtnu, alltottotamtnu As Double
Dim MaxTrnId As Integer
Dim i As Integer
Dim etype As Double
Dim btype As Double
Dim eplusb As Double
Dim mycase As Integer
Dim myfamercount As Integer
Dim ValidRow As Boolean
Dim maxDistNo As Integer
Dim sourceDno, DestDno As Integer



Private Sub Check1_Click()
mygrid.Visible = False
End Sub

Private Sub cbotrnid_LostFocus()
Dim rs As New ADODB.Recordset
cbotrnid.Enabled = False
Set rs = Nothing
rs.Open "select * from tblplantdistributionheader where trnid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then

txtdesc.Text = rs!distributionname

If Operation = "OPEN" Then


End If



End If


End Sub

Private Sub chkcf_Click()
If chkcf.Value = 1 Then
chkgrf.Value = 0
End If
End Sub

Private Sub chkgrf_Click()
If chkgrf.Value = 1 Then
chkcf.Value = 0
End If
End Sub

Private Sub chkPolinizer_Click()
If chkPolinizer.Value = 0 Then
chkgrf.Value = 0
chkcf.Value = 0
chkpriority.Value = 0
End If
Dzstr = ""
If chkPolinizer.Value = 1 Then
Frame1.Width = 9855
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
Frame1.Width = 3495
chkpriority.Value = 0
LSTPR.Clear
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTPR.Clear

If chkunplanned.Value = 0 Then
txtfcode.Visible = False
txtcratecnt.Visible = False
'rs.Open "select substring(dgt,1,3) dzongkhagid,substring(dgt,4,3) gewogid,substring(dgt,7,3) tshewogid  from  dgt where dgt in(select substring(farmercode,1,9) from tblplanted where year='2013') order by id ", MHVDB, adOpenStatic
rs.Open "select * from tbltshewog where dzongkhagid in " & Dzstr & "order by dzongkhagid,gewogid,tshewogid ", MHVDB, adOpenStatic

Else
txtfcode.Visible = True
txtcratecnt.Visible = True
rs.Open "select a.* from tblfarmer a,tblpolinizer b where a.IDFARMER=b.farmercode and substring(idfarmer,1,3) in " & Dzstr & "order by idfarmer ", MHVDB
End If
With rs
Do While Not .EOF
If chkunplanned.Value = 0 Then
FindDZ rs!dzongkhagid
FindGE rs!dzongkhagid, rs!gewogid
FindTs rs!dzongkhagid, rs!gewogid, rs!tshewogid
LSTPR.AddItem rs!dzongkhagid & " " & Trim(Dzname) + " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(TsName) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
Else
LSTPR.AddItem Trim(rs!farmername) + " | " + rs!idfarmer

End If
   .MoveNext
Loop
End With


Else
Frame1.Width = 3495
End If


End Sub

Private Sub chkpriority_Click()
If chkpriority.Value = 0 Then
chkgrf.Value = 0
chkcf.Value = 0
chkPolinizer.Value = 0
End If
Dzstr = ""
If chkpriority.Value = 1 Then
Frame1.Width = 9855
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
Frame1.Width = 3495
chkpriority.Value = 0
LSTPR.Clear
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If


Dim rs As New ADODB.Recordset

Set rs = Nothing
LSTPR.Clear

If chkunplanned.Value = 0 Then
txtfcode.Visible = False
txtcratecnt.Visible = False
'rs.Open "select substring(dgt,1,3) dzongkhagid,substring(dgt,4,3) gewogid,substring(dgt,7,3) tshewogid  from  dgt where dgt in(select substring(farmercode,1,9) from tblplanted where year='2013') order by id ", MHVDB, adOpenStatic
rs.Open "select * from tbltshewog where dzongkhagid in " & Dzstr & "order by dzongkhagid,gewogid,tshewogid ", MHVDB, adOpenStatic

Else
txtfcode.Visible = True
txtcratecnt.Visible = True
rs.Open "select * from tblfarmer where status='A' and substring(idfarmer,1,3) in " & Dzstr & "order by idfarmer ", MHVDB
End If
With rs
Do While Not .EOF
If chkunplanned.Value = 0 Then
FindDZ rs!dzongkhagid
FindGE rs!dzongkhagid, rs!gewogid
FindTs rs!dzongkhagid, rs!gewogid, rs!tshewogid
LSTPR.AddItem rs!dzongkhagid & " " & Trim(Dzname) + " |" + rs!gewogid & " " & Trim(GEname) + " |" + rs!tshewogid & " " & Trim(TsName) 'Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
Else
LSTPR.AddItem Trim(rs!farmername) + " | " + rs!idfarmer

End If
   .MoveNext
Loop
End With


Else
Frame1.Width = 3495
End If







End Sub

Private Sub chksel_Click()
If chksel.Value = 1 Then
Frame1.Visible = True
mygrid.Visible = False
cmdsave.Enabled = False
frmdistributionschedule.WindowState = 0
Else
Frame1.Visible = False
mygrid.Visible = True
cmdsave.Enabled = True
End If
End Sub

Private Sub cmdload_Click()

mygrid.rows = 5
If Operation = "ADD" Then
Frame1.Visible = True
txtindecator.Text = ""
Frame1.Visible = False
chksel.Value = 0
loadgrid
mygrid.Visible = True
cmdsave.Enabled = True

addgrid

ElseIf Operation = "OPEN" Then
txtindecator.Text = "S"
Frame1.Visible = False
cbotrnid.Enabled = False
mygrid.Visible = True

loadgridfromdb

cmdsave.Enabled = True

Else
End If

End Sub
Private Sub loadgridfromdb()
'On Error Resume Next
Dim s As Integer
Dim SQLSTR As String
mygrid.Clear

SQLSTR = ""
Dim i, j As Integer
Dim rs As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Set rsadd = Nothing
mchk = True
j = 0


mygrid.Clear
mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^Refillin |^N(Crt.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^|^|^|^|^ Production |^ pollenizer|^refillB|^refillE"

SQLSTR = "SELECT * from tblplantdistributionheader where trnid='" & cbotrnid.BoundText & "' and mnth='" & cbomnth.ListIndex + 1 & "' and year='" & cboyear.Text & "'"
rs.Open SQLSTR, MHVDB
If rs.EOF <> True Then
txtdesc.Text = rs!distributionname
loaddetail cbotrnid.BoundText, cbomnth.ListIndex + 1, cboyear.Text

Else
MsgBox "No Record Found."
End If
                            
                            
                            
                            
                            
                    
                            

                            
 

End Sub
Private Sub loaddetail(trnid As Integer, mnth As Integer, yr As Integer)

'On Error Resume Next
Dim s As Integer
Dim SQLSTR As String
SQLSTR = ""
Dim i, j As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rsadd = Nothing
mchk = True
j = 0
i = 1

SQLSTR = "SELECT * from tblplantdistributiondetail where trnid='" & trnid & "' and mnth='" & mnth & "' and year='" & yr & "' and status<>'C' order by sno"
rs.Open SQLSTR, MHVDB
txtdno.Text = Val(mygrid.TextMatrix(1, 1))
If rs.EOF <> True Then
 Do While rs.EOF <> True
                            If i >= 5 Then
                            mygrid.rows = mygrid.rows + 1
                            End If
                            mygrid.TextMatrix(i, 0) = rs!sno
                            mygrid.TextMatrix(i, 1) = IIf(rs!distno <> 0, rs!distno, "")
                            If mygrid.TextMatrix(i, 28) <> "S" Then
                            FindDZ Mid(rs!farmercode, 1, 3)
                            FindGE Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3)
                            FindTs Mid(rs!farmercode, 1, 3), Mid(rs!farmercode, 4, 3), Mid(rs!farmercode, 7, 3)
                            FindFA rs!farmercode, "F"
                           
                            mygrid.TextMatrix(i, 2) = Mid(rs!farmercode, 1, 3) & " " & Dzname
                            mygrid.TextMatrix(i, 3) = Mid(rs!farmercode, 4, 3) & " " & GEname
                            mygrid.TextMatrix(i, 4) = Mid(rs!farmercode, 7, 3) & " " & TsName
                            mygrid.TextMatrix(i, 5) = rs!farmercode
                            mygrid.TextMatrix(i, 6) = FAName
                            
                            Set rs1 = Nothing
                            rs1.Open "select * from tblfarmer where idfarmer='" & rs!farmercode & "'", MHVDB
                            If rs1.EOF <> True Then
                            mygrid.TextMatrix(i, 7) = IIf(IsNull(rs1!phone1), "", rs1!phone1)
                            mygrid.TextMatrix(i, 8) = rs1!VILLAGE
                            End If
                            End If
                            
                            mygrid.TextMatrix(i, 9) = Format(IIf(IsNull(rs!area), 0#, rs!area), "####0.00")
                            mygrid.TextMatrix(i, 10) = rs!totalplant
                            mygrid.TextMatrix(i, 11) = rs!crateno
                            
                            
                            
                            mygrid.TextMatrix(i, 12) = rs!bcrate
                            mygrid.TextMatrix(i, 13) = rs!ecrate
                            mygrid.TextMatrix(i, 14) = rs!bno
                            mygrid.TextMatrix(i, 15) = rs!plno
                            mygrid.TextMatrix(i, 16) = rs!crate
                            mygrid.TextMatrix(i, 17) = IIf(IsNull(rs!ssp), "", rs!ssp)
                            mygrid.TextMatrix(i, 18) = IIf(IsNull(rs!mop), "", rs!mop)
                            mygrid.TextMatrix(i, 19) = IIf(IsNull(rs!urea), "", rs!urea)
                            mygrid.TextMatrix(i, 20) = IIf(IsNull(rs!dolomite), "", rs!dolomite)
                            mygrid.TextMatrix(i, 21) = IIf(IsNull(rs!totalkg1), "", rs!totalkg1)
                            mygrid.TextMatrix(i, 22) = IIf(IsNull(rs!amountnu1), "", rs!amountnu1)
                            mygrid.TextMatrix(i, 23) = IIf(IsNull(rs!kg), "", rs!kg)
                            mygrid.TextMatrix(i, 24) = IIf(IsNull(rs!amountnu2), "", rs!amountnu2)
                            mygrid.TextMatrix(i, 25) = IIf(IsNull(rs!totalamount), "", rs!totalamount)
                            mygrid.TextMatrix(i, 26) = IIf(IsNull(rs!Schedule), "", rs!Schedule)
                            mygrid.TextMatrix(i, 27) = rs!serialmatch
                            mygrid.TextMatrix(i, 28) = rs!subtotindicator
                            mygrid.TextMatrix(i, 29) = rs!newold
                            mygrid.TextMatrix(i, 30) = rs!oldonly
                            mygrid.TextMatrix(i, 31) = rs!ferttranno
                            mygrid.TextMatrix(i, 32) = rs!refilltrnno
                            mygrid.TextMatrix(i, 33) = rs!production
                            mygrid.TextMatrix(i, 34) = rs!pollinizer
                            mygrid.TextMatrix(i, 35) = rs!refillB
                            mygrid.TextMatrix(i, 36) = rs!refillE
                            If mygrid.TextMatrix(i, 28) = "S" Then
                            formatsubtot1 (i)
                            End If
                         
                            i = i + 1
                            rs.MoveNext
                            Loop

Else
MsgBox "No Record Found."
End If

txtdno.Text = Val(mygrid.TextMatrix(1, 1))

   i = i - 1
 formatsubtot1 (i)
 

mygrid.MergeCells = flexMergeFree
mygrid.MergeCol(1) = True
mygrid.MergeCells = flexMergeFree
mygrid.MergeCol(26) = True
                            
                            
                            
                           
                      
End Sub

Private Sub cmdsave_Click()
On Error GoTo err
Dim disttype As String
If chkunplanned.Value = 1 Then
disttype = "N"
Else
disttype = "Y"
End If
If Len(cboyear.Text) = 0 Or Len(cbomnth.Text) = 0 Or Len(txtdesc.Text) = 0 Then
MsgBox "Fill The approprite fields like Year,Month and Description"

Exit Sub
End If

Dim SQLSTR As String
SQLSTR = ""
MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute "insert into tblplantdistributionheader(trnid,mnth,entrydate,distributionname,year,status,planneddist)" _
             & "values('" & cbotrnid.Text & "','" & cbomnth.ListIndex + 1 & "','" & Format(Now, "yyyy-MM-dd") & "'," _
             & "'" & txtdesc.Text & "','" & cboyear.Text & "','ON','" & disttype & "')"
             
             
'MHVDB.Execute "delete from tblplantdistributiondetail where tranid='" & cbotrnid.Text & "'," _
              & " and mnth='" & yy & "' and year='" & cboyear.Text & "'"
 For i = 1 To mygrid.rows - 1
 If Len(mygrid.TextMatrix(i, 0)) = 0 Then Exit For
 findMonitorFromFarmer mygrid.TextMatrix(i, 5)
SQLSTR = "insert into tblplantdistributiondetail(trnid,year,mnth,sno,distno," _
             & "farmercode,area,totalplant,crateno,bcrate,ecrate,bno,plno,crate,ssp," _
             & "mop,urea,dolomite,totalkg1,amountnu1,kg,amountnu2,totalamount," _
             & "schedule,serialmatch,subtotindicator,newold,oldonly,ferttranno," _
             & " refilltrnno,fname,dz,ge,ts,staffid,bcratefactor,ecratefactor,bnofactor,plnofactor,cratefactor,production,pollinizer,refillB,refillE) " _
             & " values('" & cbotrnid.Text & "','" & cboyear.Text & "','" & cbomnth.ListIndex + 1 & "'," _
             & "'" & mygrid.TextMatrix(i, 0) & "','" & mygrid.TextMatrix(i, 1) & "', " _
             & "'" & mygrid.TextMatrix(i, 5) & "','" & mygrid.TextMatrix(i, 9) & "'," _
             & "'" & mygrid.TextMatrix(i, 10) & "','" & mygrid.TextMatrix(i, 11) & "'," _
             & "'" & mygrid.TextMatrix(i, 12) & "','" & mygrid.TextMatrix(i, 13) & "'," _
            & "'" & Val(mygrid.TextMatrix(i, 14)) & "','" & Val(mygrid.TextMatrix(i, 15)) & "', " _
             & "'" & Val(mygrid.TextMatrix(i, 16)) & "','" & mygrid.TextMatrix(i, 17) & "'," _
             & "'" & mygrid.TextMatrix(i, 18) & "','" & mygrid.TextMatrix(i, 19) & "'," _
             & "'" & mygrid.TextMatrix(i, 20) & "','" & mygrid.TextMatrix(i, 21) & "'," _
             & "'" & mygrid.TextMatrix(i, 22) & "','" & mygrid.TextMatrix(i, 23) & "', " _
             & "'" & mygrid.TextMatrix(i, 24) & "','" & mygrid.TextMatrix(i, 25) & "'," _
             & "'" & mygrid.TextMatrix(i, 26) & "','" & mygrid.TextMatrix(i, 27) & "'," _
             & "'" & mygrid.TextMatrix(i, 28) & "','" & mygrid.TextMatrix(i, 29) & "','" & mygrid.TextMatrix(i, 30) & "','" & mygrid.TextMatrix(i, 31) & "','" & mygrid.TextMatrix(i, 32) & "', " _
             & "'" & mygrid.TextMatrix(i, 5) & " " & mygrid.TextMatrix(i, 6) & "','" & mygrid.TextMatrix(i, 2) & "','" & mygrid.TextMatrix(i, 3) & "','" & mygrid.TextMatrix(i, 4) & "','" & monitorFromFarmer & "','35','35','35','35','1','" & mygrid.TextMatrix(i, 33) & "','" & mygrid.TextMatrix(i, 34) & "','" & mygrid.TextMatrix(i, 35) & "','" & mygrid.TextMatrix(i, 36) & "')"
             
             MHVDB.Execute SQLSTR
             
Next

ElseIf Operation = "OPEN" Then

'MHVDB.Execute "update tblplantdistributionheader set trnid,mnth,entrydate,distributionname,year,status)" _
'             & "values('" & cbotrnid.Text & "','" & cbomnth.Index + 1 & "','" & Format(Now, "yyyy-MM-dd") & "'," _
'             & "'" & txtdesc.Text & "','" & cboyear.Text & "','ON')"
             
             
MHVDB.Execute "delete from tblplantdistributiondetail where trnid='" & cbotrnid.Text & "'" _
              & " and mnth='" & cbomnth.ListIndex + 1 & "' and year='" & cboyear.Text & "'"
 For i = 1 To mygrid.rows - 1
 If Len(mygrid.TextMatrix(i, 0)) = 0 Then Exit For
 findMonitorFromFarmer mygrid.TextMatrix(i, 5)
SQLSTR = "insert into tblplantdistributiondetail(trnid,year,mnth,sno,distno," _
             & "farmercode,area,totalplant,crateno,bcrate,ecrate,bno,plno,crate,ssp," _
             & "mop,urea,dolomite,totalkg1,amountnu1,kg,amountnu2,totalamount," _
             & "schedule,serialmatch,subtotindicator,newold,ferttranno,refilltrnno,oldonly,fname,dz,ge,ts,staffid,production,pollinizer,refillB,refillE) " _
             & " values('" & cbotrnid.Text & "','" & cboyear.Text & "','" & cbomnth.ListIndex + 1 & "'," _
             & "'" & mygrid.TextMatrix(i, 0) & "','" & mygrid.TextMatrix(i, 1) & "', " _
             & "'" & mygrid.TextMatrix(i, 5) & "','" & mygrid.TextMatrix(i, 9) & "'," _
             & "'" & mygrid.TextMatrix(i, 10) & "','" & mygrid.TextMatrix(i, 11) & "'," _
             & "'" & mygrid.TextMatrix(i, 12) & "','" & mygrid.TextMatrix(i, 13) & "'," _
              & "'" & Val(mygrid.TextMatrix(i, 14)) & "','" & Val(mygrid.TextMatrix(i, 15)) & "', " _
             & "'" & Val(mygrid.TextMatrix(i, 16)) & "','" & mygrid.TextMatrix(i, 17) & "'," _
             & "'" & mygrid.TextMatrix(i, 18) & "','" & mygrid.TextMatrix(i, 19) & "'," _
             & "'" & mygrid.TextMatrix(i, 20) & "','" & mygrid.TextMatrix(i, 21) & "'," _
             & "'" & mygrid.TextMatrix(i, 22) & "','" & mygrid.TextMatrix(i, 23) & "', " _
             & "'" & mygrid.TextMatrix(i, 24) & "','" & mygrid.TextMatrix(i, 25) & "'," _
             & "'" & mygrid.TextMatrix(i, 26) & "','" & mygrid.TextMatrix(i, 27) & "'," _
             & "'" & mygrid.TextMatrix(i, 28) & "','" & mygrid.TextMatrix(i, 29) & "'," _
             & "'" & mygrid.TextMatrix(i, 31) & "','" & mygrid.TextMatrix(i, 32) & "','" & mygrid.TextMatrix(i, 30) & "'," _
             & "'" & mygrid.TextMatrix(i, 5) & " " & mygrid.TextMatrix(i, 6) & "','" & mygrid.TextMatrix(i, 2) & "','" & mygrid.TextMatrix(i, 3) & "','" & mygrid.TextMatrix(i, 4) & "','" & monitorFromFarmer & "','" & mygrid.TextMatrix(i, 33) & "','" & mygrid.TextMatrix(i, 34) & "','" & mygrid.TextMatrix(i, 35) & "','" & mygrid.TextMatrix(i, 36) & "')"
             
             MHVDB.Execute SQLSTR
Next

Else
MsgBox "Invalid Selection of Criteria."
   MHVDB.RollbackTrans
   Exit Sub
End If
MsgBox "Record Saved Successfully."
cmdsave.Enabled = False
MHVDB.CommitTrans
Exit Sub
err:
    MHVDB.RollbackTrans
    MsgBox err.Description
    
MHVDB.Execute "update refillinheader set status='OK' where id='" & cbotrnid.Text & "'"
End Sub

Private Sub Command1_Click()
addgrid

End Sub

Private Sub Command10_Click()
If Len(cbotrnid.Text) = 0 Then

MsgBox "Select the schedule to finalize!"
Exit Sub
End If

If MsgBox("Are you sure , you want to finalize the schedule? Finalizing stops further editing!", vbYesNo) = vbYes Then


End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim rs As New ADODB.Recordset
Frame1.Visible = True
Frame1.Width = 3495
mygrid.Visible = False
Operation = "ADD"
cmdsave.Enabled = True
cbotrnid.Enabled = False
Set rs = Nothing
rs.Open "select max(trnid) as maxid from tblplantdistributionheader", MHVDB
cbotrnid.Text = IIf(IsNull(rs!MaxId), 0, rs!MaxId) + 1
txtdno.Text = ""




'cmdsave.Enabled = True
End Sub

Private Sub Command4_Click()
Frame1.Visible = False
End Sub

Private Sub loadgrid()
'On Error Resume Next
Dim polycont As Long
Dim mydgt As String
Dim cnt As Integer
mydgt = ""
Dim morderstr As String
Dim muk As Integer
muk = 0
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim s As Integer
Dim SQLSTR As String
Dim krecord As New ADODB.Recordset


Set rs1 = Nothing
rs1.Open "select max(distno) as maxid from tblplantdistributiondetail", MHVDB
maxDistNo = IIf(IsNull(rs1!MaxId), 0, rs1!MaxId) + 1
txtdno.Text = maxDistNo
Set rs1 = Nothing

morderstr = ""
mygrid.Clear
mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^Refillin |^P1 (Nos.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^|^ Production |^ pollenizer|^refillB|^refillE|"
etype = 0
ptype = 0
SQLSTR = ""
Dim i, j As Integer
Dim rs As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Set rsadd = Nothing
mchk = True
j = 0
Dzstr = ""
morderstr = ""
If chkpriority.Value = 0 Then
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
ElseIf chkpriority.Value = 1 And chkunplanned.Value = 0 Then
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    mm = Split(LSTPR.List(i), "|", -1, 1)
       Dzstr = Dzstr & "'" & Mid(mm(0), 1, 3) & Mid(mm(1), 1, 3) & Mid(mm(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = LSTPR.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE LOCATION TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
Else

Dzstr = ""
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
End If



If Len(Dzstr) > 0 Then
morderstr = Left(Dzstr, Len(Dzstr) - 1)
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"

Else
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If

'Dzstr = ""
'For i = 0 To LSTPR.ListCount - 1
'    If LSTPR.Selected(i) Then
'       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
'       Mcat = DZLIST.List(i)
'       j = j + 1
'    End If
'    If RepName = "5" Then
'       If j > 1 Then
'          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
'          Exit Sub
'       End If
'    End If
'Next
'If Len(Dzstr) > 0 Then
'   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
'
'End If


mygrid.Clear
'mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^P    |^P1 (Nos.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^"
mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^Refillin|^N (Crt.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^|^|^|^|^Production|^pollenizer|^refillB|^refillE"




If chkpriority.Value = 0 And chkPolinizer.Value = 0 And chkrefill.Value = 0 Then
    If chkgrf.Value = 0 And chkcf.Value = 0 Then
            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and plantedstatus='N' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9),IDFARMER"

        
       
        
        ElseIf chkgrf.Value = 1 Then
            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='G' and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9),IDFARMER"
    Else
    SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='C' and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9),IDFARMER"
    
    End If
    Else
        myfamercount = 0
        If chkgrf.Value = 0 And chkcf.Value = 0 And chkPolinizer.Value = 0 Then
        SQLSTR = ""
        If chkunplanned = 1 Then
         SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND (IDFARMER)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "


       MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
       SQLSTR = ""
       'SQLSTR = "SELECT SUBSTRING(farmercode,1,3) AS DZCODE,SUBSTRING(farmercode,4,3) AS GECODE,SUBSTRING(farmercode,7,3) AS TSCODE,farmercode,FARMERNAME,0 AS REGLAND,village,phone1 FROM tblfarmer A,refillin B WHERE A.status='A' and A.IDFARMER=B.farmercode and substring(farmercode,10,1)='F' and SUBSTRING(farmercode,1,9)IN  " & Dzstr & " and farmercode not in(select idfarmer from tbldistpreparetion)"
        '    SQLSTR = SQLSTR & "  " & "group by farmercode order by  FIELD(SUBSTRING(farmercode,1,9), " & morderstr & ") "
       'MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
      '            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
'            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
  SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1)" _
           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1 from " _
& "refillin a,tblfarmer b where idfarmer=farmercode and a.status='ON' and farmercode not in(select idfarmer from tbldistpreparetion)"
MHVDB.Execute SQLSTR
  SQLSTR = "SELECT * from tbldistpreparetion order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "

ElseIf chkrefill.Value = 1 Then
      MHVDB.Execute "delete from  tbldistpreparetion"
'      SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1,refillplant,totfillindemand)" _
'           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1,totplants,totfillin from " _
'           & "mhv.vrefillin a,mhv.tblfarmer b where idfarmer=farmercode and SUBSTRING(IDFARMER,1,3) IN  " & Dzstr & " and a.status='ON' and a.isfinalized='Yes'"
'           MHVDB.Execute SQLSTR
      SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1,refillplant,totfillindemand,totpollinizerdemand,totproductiondemand)" _
           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1,totplants,totfillin,P1,B from " _
           & "mhv.vrefillin a,mhv.tblfarmer b where idfarmer=farmercode and SUBSTRING(IDFARMER,1,3) IN  " & Dzstr & " and a.status='ON' and a.isfinalized='Yes'"
           MHVDB.Execute SQLSTR
      SQLSTR = "SELECT * from tbldistpreparetion order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "

Else

        SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(REGLAND) AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status not in('D','R','C')and B.status not in('D','R','C') and plantedstatus='N'  and A.IDFARMER=B.FARMERID and substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
        SQLSTR = SQLSTR & "  " & "group by idfarmer "
SQLSTR = SQLSTR & " union  SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE, " _
& " SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(acre) AS REGLAND,village,phone1 FROM " _
& " tblfarmer A,tbllandregdetail B WHERE A.status not in('D','R','C') and plantedstatus='N'  " _
& " and A.IDFARMER=B.farmercode  and substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr & " group by idfarmer order by " _
& " FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ")"

MHVDB.Execute "delete from  tbldistpreparetion"

       MHVDB.Execute "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1) " & SQLSTR
       SQLSTR = ""
       'SQLSTR = "SELECT SUBSTRING(farmercode,1,3) AS DZCODE,SUBSTRING(farmercode,4,3) AS GECODE,SUBSTRING(farmercode,7,3) AS TSCODE,farmercode,FARMERNAME,0 AS REGLAND,village,phone1 FROM tblfarmer A,refillin B WHERE A.status='A' and A.IDFARMER=B.farmercode and substring(farmercode,10,1)='F' and SUBSTRING(farmercode,1,9)IN  " & Dzstr & " and farmercode not in(select idfarmer from tbldistpreparetion)"
        '    SQLSTR = SQLSTR & "  " & "group by farmercode order by  FIELD(SUBSTRING(farmercode,1,9), " & morderstr & ") "
       'MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
      '            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
'            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "


  SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1)" _
           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1 from " _
& "vrefillin a,tblfarmer b where idfarmer=farmercode and a.status='ON' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr & " group by idfarmer"

MHVDB.Execute SQLSTR

 SQLSTR = "SELECT * from tbldistpreparetion group by idfarmer order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "
 'SQLSTR = "SELECT * from tbldistpreparetion where IDFARMER in ('D10G01T01F0021') group by idfarmer "

End If
        
        

        
        
        ElseIf chkgrf.Value = 1 Then
        
            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1,sum(plantqty)polinizercrate  FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='G' and  SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
        ElseIf chkPolinizer.Value = 1 Then
        
    'added by kinzang ## begin
    '    MHVDB.Execute "insert into mhv.tblpolinizer(farmercode,plantqty,myear,pvariety,originalplantqty,flag) select a.farmercode,noofcrate,mindistyear,3,noofcrate,1 from mhv.tblpolinizerjune a LEFT JOIN mhv.tblpollenizeranalysis b ON a.farmercode=b.farmercode"
    '## end
    ' add flag condition for code below
    'and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr * extracting with flag =0 is enough
    
        SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,0 AS REGLAND,village,phone1,sum(plantqty) polinizercrate FROM tblfarmer A,tblpolinizer B WHERE A.IDFARMER=B.farmercode and B.flag=0 "
        SQLSTR = SQLSTR & "  " & "group by idfarmer "
MHVDB.Execute "delete from  tbldistpreparetion"
     MHVDB.Execute "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1,polinizercrate) " & SQLSTR
       SQLSTR = ""
       'SQLSTR = "SELECT SUBSTRING(farmercode,1,3) AS DZCODE,SUBSTRING(farmercode,4,3) AS GECODE,SUBSTRING(farmercode,7,3) AS TSCODE,farmercode,FARMERNAME,0 AS REGLAND,village,phone1 FROM tblfarmer A,refillin B WHERE A.status='A' and A.IDFARMER=B.farmercode and substring(farmercode,10,1)='F' and SUBSTRING(farmercode,1,9)IN  " & Dzstr & " and farmercode not in(select idfarmer from tbldistpreparetion)"
        '    SQLSTR = SQLSTR & "  " & "group by farmercode order by  FIELD(SUBSTRING(farmercode,1,9), " & morderstr & ") "
       'MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
      '            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
'            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "

  SQLSTR = "SELECT * from tbldistpreparetion where order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "
  
    
    Else
      SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='C' and  SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
    
    
    End If

End If
Dim mrnd As Integer
Dim tmod As Integer
mrnd = 0

If chkrefill.Value <> 1 And chkpriority.Value = 0 Then
MHVDB.Execute "delete from mhv.tbldistpreparetion where substring(idfarmer,1,14) in( select substring(FARMER,1,14) from " _
& " (select n.FARMER,n.START,n.PLANTFUTURE,round(datediff(CURDATE(),n.START),0) recordage,abs(round(datediff(n.START,END),0)) daydiff" _
& " from odk_prodlocal.tblfuturefarmer_core n INNER JOIN (SELECT FARMER,MAX(START) lastEntry FROM odk_prodlocal.tblfuturefarmer_core" _
& " where PLANTFUTURE IN('no') GROUP BY FARMER)x ON n.FARMER = x.FARMER AND n.START = x.lastEntry AND STATUS <> 'BAD' GROUP BY n.FARMER)" _
& "mm )"
End If
                            rs.Open SQLSTR, MHVDB
                            
                            i = 1
                            cnt = maxDistNo
                            polycont = 0
                            Do Until rs.EOF
                            'If polycont > 77500 Then Exit Do (this was for 2013 extra plant distribution, as decided by mgmt)
                           
                                 mydgt = Mid(rs!idfarmer, 1, 9)
                                 Do While mydgt = Mid(rs!idfarmer, 1, 9)
                                 
                                         If i >= 5 Then
                                         mygrid.rows = mygrid.rows + 1
                                         End If
                                         mygrid.TextMatrix(i, 0) = i
                                         mygrid.TextMatrix(i, 1) = cnt
                                         FindDZ rs!dzcode
                                         FindGE rs!dzcode, rs!GECODE
                                         FindTs rs!dzcode, rs!GECODE, rs!tscode
                                         mygrid.TextMatrix(i, 2) = rs!dzcode & " " & Dzname
                                         mygrid.TextMatrix(i, 3) = rs!GECODE & " " & GEname
                                         mygrid.TextMatrix(i, 4) = rs!tscode & " " & TsName
                                         mygrid.TextMatrix(i, 5) = rs!idfarmer
                                         mygrid.TextMatrix(i, 6) = rs!farmername
                                         mygrid.TextMatrix(i, 7) = IIf(IsNull(rs!phone1), "", rs!phone1)
                                         mygrid.TextMatrix(i, 8) = rs!VILLAGE
                                         mygrid.TextMatrix(i, 9) = Format(IIf(IsNull(rs!regland), 0#, rs!regland), "####0.00")
                                         
                                                      If chkPolinizer.Value = 1 Then
                                                       mygrid.TextMatrix(i, 10) = Round(rs!polinizercrate * 35, 0)
                                                       mygrid.TextMatrix(i, 11) = Round(rs!polinizercrate, 2) '(Val(Mygrid.TextMatrix(i, 10)) - (Val(Mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
                                                      ' mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, mrnd)
                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       
                                                       'End If
                                                       End If
                                                    If chkrefill.Value = 1 Then
                                                       mygrid.TextMatrix(i, 10) = rs!totfillindemand
                                                       'Mygrid.TextMatrix(i, 12) = rs!totfillindemand
                                                        mygrid.TextMatrix(i, 12) = rs!totproductiondemand
                                                        mygrid.TextMatrix(i, 14) = rs!totpollinizerdemand
                                                        mygrid.TextMatrix(i, 31) = mygrid.TextMatrix(i, 12)
                                                    End If
                                         Set rs1 = Nothing
                                        rs1.Open "select group_concat(a.id) refilltrnno,sum(b) as b,sum(e) as e,sum(p1) as p1, sum(n) as n,refillB,refillE from refillin a, refillinheader b where a.headerid=b.id and farmercode='" & rs!idfarmer & "' and  isfinalized='Yes' and status='ON'   group by farmercode", MHVDB
                                         'rs1.Open "select sum(regland)*(420*.1*.5) as b,sum(regland)*(420*.1*.5) as e,sum(regland)*(420*.06) as p1, sum(regland)*(420*.1) as n from tbllandreg where farmerid='" & rs!idfarmer & "' group by farmerid ", MHVDB
                                         
                                         
                                         If rs1.EOF <> True And chkrefill.Value = 0 Then
                                                btype = Round(rs1!b, mrnd)
                                                etype = Round(rs1!e, mrnd)
                                                ptype = Round(rs1!p1, mrnd)
                                                Set RS2 = Nothing
                                                RS2.Open "select * from tbldistformula where status='ON'", MHVDB
                                                If RS2.EOF <> True Then
                                                mygrid.TextMatrix(i, 10) = Round(rs1!p1 + rs1!n + btype + etype, mrnd)
                                                       'If etype > 21 Then
'                                                       mygrid.TextMatrix(i, 13) = etype
'                                                       mygrid.TextMatrix(i, 10) = Round(btype + Val(mygrid.TextMatrix(i, 13)) + rs1!p1 + rs1!n, 2)
'                                                       mygrid.TextMatrix(i, 11) = Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, 2)
'                                                       mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, 2)
'                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, 2)
'                                                        If etype > 35 Then
'                                                                tmod = etype Mod 35
'                                                                If tmod > 17 Then
'                                                                etype = etype + 35 - tmod
'                                                                btype = btype - tmod
'                                                                Else
'                                                                etype = etype - tmod
'                                                                btype = btype - tmod
'
'                                                                End If
'                                                        Else
'                                                        tmod = 0
'                                                        etype = 35
'                                                        btype = 35
'
'                                                        End If

                                                       'mygrid.TextMatrix(i, 13) = etype
                                                       'mygrid.TextMatrix(i, 14) = ptype
                                                    '   mygrid.TextMatrix(i, 10) = Round(btype + rs1!p1 + rs1!n, mrnd)
                                                    '   mygrid.TextMatrix(i, 11) = (Val(mygrid.TextMatrix(i, 10)) - (Val(mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                        
                                                        'mygrid.TextMatrix(i, 10) = Round(((Val(mygrid.TextMatrix(i, 9)) * RS2!totalplant)), 0) + rs1!p1 + rs1!b
                                                        mygrid.TextMatrix(i, 10) = Round(((Val(mygrid.TextMatrix(i, 9)) * RS2!totalplant)), 0)
                                                        mygrid.TextMatrix(i, 11) = mygrid.TextMatrix(i, 10) 'Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))), 0) ' Val(mygrid.TextMatrix(i, 11)) + Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))) / rs!crateno, 0)
                                                        modval = mygrid.TextMatrix(i, 11)
                                                        mmod = modval Mod RS2!crateno
                                                        If (mmod > 17) Then
                                                        mygrid.TextMatrix(i, 11) = ((modval - mmod) / RS2!crateno) + 1
                                                        Else
                                                        mygrid.TextMatrix(i, 11) = ((modval - mmod) / RS2!crateno)
                                                        End If
                                                        mygrid.TextMatrix(i, 12) = Round((Val(mygrid.TextMatrix(i, 11)) * RS2!crateno * RS2!bcrate), 0) 'Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate) / rs!crateno, 0)
                                                        modval = mygrid.TextMatrix(i, 12)
                                                        mmod = modval Mod RS2!crateno
                                                        If (Val(mygrid.TextMatrix(i, 9)) <> 0) Then
                                                        If (mmod > 17) Then
                                                        
                                                        mygrid.TextMatrix(i, 12) = ((modval - mmod) / RS2!crateno) + 1
                                                        Else
                                                        mygrid.TextMatrix(i, 12) = ((modval - mmod) / RS2!crateno)
                                                        End If
                                                        Else
                                                        mygrid.TextMatrix(i, 11) = 0
                                                        mygrid.TextMatrix(i, 12) = 0
                                                        mygrid.TextMatrix(i, 13) = 0
                                                        End If
                                                       
                                                        
'                                                        If Mygrid.TextMatrix(i, 9) > 0 Then
'                                                        Else
'                                                        Mygrid.TextMatrix(i, 15) = Val(Mygrid.TextMatrix(i, 10)) - Round((Val(Mygrid.TextMatrix(i, 12)) * RS2!crateno))
'                                                        modval = Mygrid.TextMatrix(i, 15)
'                                                        mmod = modval Mod RS2!crateno
'                                                        If (mmod > 17) Then
'                                                        Mygrid.TextMatrix(i, 15) = ((modval - mmod) / RS2!crateno) + 1
'                                                        Else
'                                                        Mygrid.TextMatrix(i, 15) = ((modval - mmod) / RS2!crateno)
'                                                        End If
'                                                        End If
                                                       
                                                
                                                       ' mygrid.TextMatrix(i, 13) = Round((Val(mygrid.TextMatrix(i, 11)) * RS2!crateno * RS2!bcrate), 0) 'Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate) / rs!crateno, 0)
                                                      
                                                        
                                                     '  mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
                                                      ' mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, mrnd)
                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       
                                                       'End If
            
'                                                        mygrid.TextMatrix(i, 13) = etype
'                                                      ' mygrid.TextMatrix(i, 10) = rs!polinizercrate
'                                                       mygrid.TextMatrix(i, 11) = (Val(mygrid.TextMatrix(i, 10)) - (Val(mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
'                                                       mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
'                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
'                                                      ' mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, mrnd)
'                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       
                                                       'End If
            
            
         
                                                       
                                                       
                                                       
                                                End If
                                                

                                                mygrid.TextMatrix(i, 14) = rs1!b + rs1!p1
                                                mygrid.TextMatrix(i, 15) = Round(rs1!p1, mrnd) '
                                                mygrid.TextMatrix(i, 16) = Round(rs1!n, 0) 'Round(rs1!n / RS2!crateno, 0)
                                                polycont = polycont + Round(rs1!p1, mrnd) + Round(rs1!n, mrnd)
                                                mygrid.TextMatrix(i, 29) = "O"
                                                mygrid.TextMatrix(i, 30) = mygrid.TextMatrix(i, 16)
                                                mygrid.TextMatrix(i, 31) = RS2!fid
                                                mygrid.TextMatrix(i, 32) = rs1!refilltrnno
                                                mygrid.TextMatrix(i, 33) = Round(rs1!b, mrnd)
                                                mygrid.TextMatrix(i, 34) = Round(rs1!p1, mrnd)
                                                mygrid.TextMatrix(i, 35) = Round(rs1!refillB, mrnd)
                                                mygrid.TextMatrix(i, 36) = Round(rs1!refillE, mrnd)
                                         End If
                                         i = i + 1
                                         rs.MoveNext
                                         
                                         If rs.EOF Then Exit Do
                                 Loop
                                 cnt = cnt + 1
                                 mygrid.rows = mygrid.rows + 1
                                 mygrid.TextMatrix(i, 28) = "S"
                                 mygrid.TextMatrix(i, 0) = i
                                 i = i + 1
                            Loop
                            mygrid.rows = mygrid.rows + 1
                            mygrid.TextMatrix(i, 28) = "T"
                            mygrid.TextMatrix(i, 0) = i
              
End Sub

Private Sub addgrid()
Dim rowtot As Double
'Mygrid.TextMatrix(i, 8)
Dim tempcalc As Double
Dim schk As Integer
Dim myloop As Integer
Dim compareme As Double
Dim mmod As Integer
Dim modval, kk As Double
Dim rs As New ADODB.Recordset
Set rs = Nothing



rs.Open "select * from tbldistformula where status='ON'", MHVDB
Dim tt As Integer
tt = Val(txtdno.Text)
schk = 0
rowtot = 0
initvariables
tt = Val(txtdno.Text)

If txtindecator.Text = "" Then
         myloop = mygrid.rows - 1
    Else
        myloop = mygrid.rows - 2
End If

If Val(txtdno.Text) = 0 Then
MsgBox "Cannot Proceed, Invalid Distribution No. Noticed! Contact IT Dept. For REcitfication"
Exit Sub
End If


For i = 1 To myloop
rowtot = 0

'If mygrid.TextMatrix(i, 5) = "D09G01T01F0063" Then
'MsgBox "fuck"
'End If

If Len(mygrid.TextMatrix(i, 0)) = 0 Then Exit For
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = tt
If Operation = "ADD" Then
'mygrid.TextMatrix(i, 10) = Val(mygrid.TextMatrix(i, 10)) + Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant)), 0) + Val(mygrid.TextMatrix(i, 14))
mygrid.TextMatrix(i, 10) = Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant)), 0) + Val(mygrid.TextMatrix(i, 14))
mygrid.TextMatrix(i, 15) = Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant) * rs!ncrate), 0)
If chkPolinizer.Value = 0 Then
mygrid.TextMatrix(i, 11) = mygrid.TextMatrix(i, 10) - Val(mygrid.TextMatrix(i, 14)) 'Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))), 0) ' Val(mygrid.TextMatrix(i, 11)) + Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))) / rs!crateno, 0)
modval = mygrid.TextMatrix(i, 11)
mmod = modval Mod rs!crateno
If (mmod > 17) Then
mygrid.TextMatrix(i, 11) = ((modval - mmod) / rs!crateno) + 1
Else
mygrid.TextMatrix(i, 11) = ((modval - mmod) / rs!crateno)
End If



End If

' polinizer is for thimphu
   'Round(((Val(Mygrid.TextMatrix(i, 15)))), 0) / rs!crateno
 mygrid.TextMatrix(i, 15) = Round(Val(mygrid.TextMatrix(i, 15)) / rs!crateno, 0)
 If chkPolinizer.Value = 0 And chkrefill.Value = 0 Then
If mygrid.TextMatrix(i, 29) <> "O" Then
If Operation = "ADD" Then

mygrid.TextMatrix(i, 12) = Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate), 0) 'Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate) / rs!crateno, 0)
modval = mygrid.TextMatrix(i, 12)
mmod = modval Mod rs!crateno
If (mmod > 17) Then
mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno) + 1
Else
mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno)
End If




Else
mygrid.TextMatrix(i, 12) = Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)   'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!crateno * rs!bcrate) / rs!crateno, 0)
End If
End If
Else
If Val(mygrid.TextMatrix(i, 9)) = 0 Then
mygrid.TextMatrix(i, 12) = Round(((Val(mygrid.TextMatrix(i, 12)))) / rs!crateno, 0) 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)

Else
mygrid.TextMatrix(i, 12) = mygrid.TextMatrix(i, 12) + Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant) * rs!bcrate), 0) 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)

modval = mygrid.TextMatrix(i, 12)
mmod = modval Mod rs!crateno
If (mmod > 17) Then
mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno) + 1
Else
mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno)
End If

End If


End If


If Mid(mygrid.TextMatrix(i, 5), 10, 1) <> "G" Or Mid(mygrid.TextMatrix(i, 5), 10, 1) <> "C" Then
        If mygrid.TextMatrix(i, 29) <> "O" Then
            If Operation = "ADD" Then
           If chkPolinizer.Value = 0 Then
            mygrid.TextMatrix(i, 13) = mygrid.TextMatrix(i, 11) - mygrid.TextMatrix(i, 12) - mygrid.TextMatrix(i, 15)  ' Round((mygrid.TextMatrix(i, 11) * rs!ecrate), 0)  'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)Round((mygrid.TextMatrix(i, 11) * rs!crateno * rs!ecrate) / rs!crateno, 0) 'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)
            Else
            mygrid.TextMatrix(i, 14) = Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 12)) - Val(mygrid.TextMatrix(i, 15))  ' Round((mygrid.TextMatrix(i, 11) * rs!ecrate), 0)  'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)Round((mygrid.TextMatrix(i, 11) * rs!crateno * rs!ecrate) / rs!crateno, 0) 'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)
            End If
            
            Else
            
                If chkPolinizer.Value = 0 Then
           mygrid.TextMatrix(i, 13) = Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!crateno * rs!ecrate) / rs!crateno, 0) 'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)
            Else
            mygrid.TextMatrix(i, 14) = Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!crateno * rs!ecrate) / rs!crateno, 0) 'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)
            End If
            End If
        Else
If Val(mygrid.TextMatrix(i, 9)) = 0 Then
 If chkPolinizer.Value = 0 Then
'Mygrid.TextMatrix(i, 13) = Round(((Val(Mygrid.TextMatrix(i, 13)))), 0) / rs!crateno 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)

mygrid.TextMatrix(i, 15) = mygrid.TextMatrix(i, 11) - mygrid.TextMatrix(i, 12)
mygrid.TextMatrix(i, 13) = mygrid.TextMatrix(i, 11) - mygrid.TextMatrix(i, 12) - mygrid.TextMatrix(i, 15)
 Else
 mygrid.TextMatrix(i, 14) = Round(((Val(mygrid.TextMatrix(i, 14)))), 0) / rs!crateno 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)
 End If

Else
mygrid.TextMatrix(i, 13) = mygrid.TextMatrix(i, 11) - mygrid.TextMatrix(i, 12) - mygrid.TextMatrix(i, 15) 'mygrid.TextMatrix(i, 13) + Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant) * rs!ecrate), 0) 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)

End If
        End If

End If


End If

If chkrefill.Value = 1 Then
mygrid.TextMatrix(i, 11) = 0
mygrid.TextMatrix(i, 12) = mygrid.TextMatrix(i, 31)
mygrid.TextMatrix(i, 13) = 0
End If
      
        
        If Mid(mygrid.TextMatrix(i, 5), 10, 1) = "G" Or Mid(mygrid.TextMatrix(i, 5), 10, 1) = "C" Then
'                mygrid.TextMatrix(i, 15) = Val(mygrid.TextMatrix(i, 10)) * 0.06
'                mygrid.TextMatrix(i, 16) = Val(mygrid.TextMatrix(i, 10)) * 0.06
'                mygrid.TextMatrix(i, 10) = Val(mygrid.TextMatrix(i, 10)) + Val(mygrid.TextMatrix(i, 15)) + Val(mygrid.TextMatrix(i, 16))
        End If
'        If mygrid.TextMatrix(i, 29) = "O" Then
'                tempcalc = Val(mygrid.TextMatrix(i, 10)) - (Val(mygrid.TextMatrix(i, 10)) Mod 35) + 35
'                mygrid.TextMatrix(i, 17) = Round((tempcalc * rs!ssp), 2)
'                mygrid.TextMatrix(i, 18) = Round((tempcalc * rs!mop), 2)
'                mygrid.TextMatrix(i, 19) = Round((tempcalc * rs!urea), 2)
'                mygrid.TextMatrix(i, 20) = Round((tempcalc * rs!dolomite), 2)
'                mygrid.TextMatrix(i, 21) = Round(Val(mygrid.TextMatrix(i, 17)) + Val(mygrid.TextMatrix(i, 18)) + Val(mygrid.TextMatrix(i, 19)) + Val(mygrid.TextMatrix(i, 20)), 0)
'                mygrid.TextMatrix(i, 22) = Round(Val(mygrid.TextMatrix(i, 17) * rs!sspperkg) + Val(mygrid.TextMatrix(i, 18) * rs!mopperkg) + Val(mygrid.TextMatrix(i, 19) * rs!ureaperkg) + Val(mygrid.TextMatrix(i, 20) * rs!dolomiteperkg), 0)
'                mygrid.TextMatrix(i, 23) = Round((tempcalc * rs!kg), 0)
'                mygrid.TextMatrix(i, 24) = Round((mygrid.TextMatrix(i, 23) * rs!amountnu), 0)
        'Else
        If Operation = "ADD" Then
                
If chkrefill.Value = 1 Then
mygrid.TextMatrix(i, 17) = Round((((Val(mygrid.TextMatrix(i, 10))) + Val(mygrid.TextMatrix(i, 16))) * rs!ssp), 2)
                mygrid.TextMatrix(i, 18) = Round((((Val(mygrid.TextMatrix(i, 10))) + Val(mygrid.TextMatrix(i, 16))) * rs!mop), 2)
                mygrid.TextMatrix(i, 19) = Round((((Val(mygrid.TextMatrix(i, 10))) + Val(mygrid.TextMatrix(i, 16))) * rs!urea), 2)
                mygrid.TextMatrix(i, 20) = Round((((Val(mygrid.TextMatrix(i, 10))) + Val(mygrid.TextMatrix(i, 16))) * rs!dolomite), 2)
                mygrid.TextMatrix(i, 21) = Val(mygrid.TextMatrix(i, 17)) + Val(mygrid.TextMatrix(i, 18)) + Val(mygrid.TextMatrix(i, 19)) + Val(mygrid.TextMatrix(i, 20))
                mygrid.TextMatrix(i, 22) = Round(Val(mygrid.TextMatrix(i, 17) * rs!sspperkg) + Val(mygrid.TextMatrix(i, 18) * rs!mopperkg) + Val(mygrid.TextMatrix(i, 19) * rs!ureaperkg) + Val(mygrid.TextMatrix(i, 20) * rs!dolomiteperkg))
                mygrid.TextMatrix(i, 23) = Round((((Val(mygrid.TextMatrix(i, 10))) + Val(mygrid.TextMatrix(i, 16))) * rs!kg) + 0.00000001, 0)
                mygrid.TextMatrix(i, 24) = Round((mygrid.TextMatrix(i, 23) * rs!amountnu), 0)
             
Else
                
'                mygrid.TextMatrix(i, 17) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!ssp), 2)
'                mygrid.TextMatrix(i, 18) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!mop), 2)
'                mygrid.TextMatrix(i, 19) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!urea), 2)
'                mygrid.TextMatrix(i, 20) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!dolomite), 2)
                mygrid.TextMatrix(i, 17) = Round((Val(mygrid.TextMatrix(i, 10)) * rs!ssp), 2)
                mygrid.TextMatrix(i, 18) = Round((Val(mygrid.TextMatrix(i, 10)) * rs!mop), 2)
                mygrid.TextMatrix(i, 19) = Round((Val(mygrid.TextMatrix(i, 10)) * rs!urea), 2)
                mygrid.TextMatrix(i, 20) = Round((Val(mygrid.TextMatrix(i, 10)) * rs!dolomite), 2)
                
                mygrid.TextMatrix(i, 21) = Val(mygrid.TextMatrix(i, 17)) + Val(mygrid.TextMatrix(i, 18)) + Val(mygrid.TextMatrix(i, 19)) + Val(mygrid.TextMatrix(i, 20))
                mygrid.TextMatrix(i, 22) = Round(Val(mygrid.TextMatrix(i, 17) * rs!sspperkg) + Val(mygrid.TextMatrix(i, 18) * rs!mopperkg) + Val(mygrid.TextMatrix(i, 19) * rs!ureaperkg) + Val(mygrid.TextMatrix(i, 20) * rs!dolomiteperkg), 0)
                mygrid.TextMatrix(i, 23) = Round((Val(mygrid.TextMatrix(i, 10)) * rs!kg) + 0.00000001, 0)
               ' mygrid.TextMatrix(i, 23) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!kg) + 0.00000001, 0)
                mygrid.TextMatrix(i, 24) = Round((mygrid.TextMatrix(i, 23) * rs!amountnu), 0)
             End If
                If Val(mygrid.TextMatrix(i, 23)) < 0 Then
                mygrid.TextMatrix(i, 23) = 0
                mygrid.TextMatrix(i, 24) = 0
                End If
                
                mygrid.TextMatrix(i, 25) = Val(mygrid.TextMatrix(i, 22)) + Val(mygrid.TextMatrix(i, 24))
                mygrid.TextMatrix(i, 27) = tt
       End If
       ' End If
        
        
        If mygrid.TextMatrix(i, 29) <> "O" Then
                rowtot = rowtot + Val(mygrid.TextMatrix(i, 10))
        Else
                rowtot = rowtot + Val(mygrid.TextMatrix(i, 10))
        End If
        If mygrid.TextMatrix(i, 28) = "" Then
                subtot = subtot + Val(mygrid.TextMatrix(i, 9))
                alltot = alltot + Val(mygrid.TextMatrix(i, 9))
                
                subtotplant = subtotplant + Val(mygrid.TextMatrix(i, 10))
                alltotplant = alltotplant + Val(mygrid.TextMatrix(i, 10))
                
                subtotcrateno = subtotcrateno + Val(mygrid.TextMatrix(i, 11))
                alltotcrateno = alltotcrateno + Val(mygrid.TextMatrix(i, 11))
                    
                subtotbcrate = subtotbcrate + Val(mygrid.TextMatrix(i, 12))
                alltotbcrate = alltotbcrate + Val(mygrid.TextMatrix(i, 12))
                  
                subtotecrate = subtotecrate + Val(mygrid.TextMatrix(i, 13))
                alltotecrate = alltotecrate + Val(mygrid.TextMatrix(i, 13))
                    
                subtotbno = subtotbno + Val(mygrid.TextMatrix(i, 14))
                alltotbno = alltotbno + Val(mygrid.TextMatrix(i, 14))
                        
                subtotplno = subtotplno + Val(mygrid.TextMatrix(i, 15))
                alltotplno = alltotplno + Val(mygrid.TextMatrix(i, 15))
                    
                subtotcrate = subtotcrate + Val(mygrid.TextMatrix(i, 16))
                alltotcrate = alltotcrate + Val(mygrid.TextMatrix(i, 16))
                
                    
                subtotssp = subtotssp + Val(mygrid.TextMatrix(i, 17))
                alltotssp = alltotssp + Val(mygrid.TextMatrix(i, 17))
                    
                subtotmop = subtotmop + Val(mygrid.TextMatrix(i, 18))
                alltotmop = alltotmop + Val(mygrid.TextMatrix(i, 18))
                    
                subtoturea = subtoturea + Val(mygrid.TextMatrix(i, 19))
                alltoturea = alltoturea + Val(mygrid.TextMatrix(i, 19))
                    
                subtotdolomite = subtotdolomite + Val(mygrid.TextMatrix(i, 20))
                alltotdolomite = alltotdolomite + Val(mygrid.TextMatrix(i, 20))
                    
                subtotkg1 = subtotkg1 + Val(mygrid.TextMatrix(i, 21))
                alltotkg1 = alltotkg1 + Val(mygrid.TextMatrix(i, 21))
                    
                subtotamtnu1 = subtotamtnu1 + Val(mygrid.TextMatrix(i, 22))
                alltotamtnu1 = alltotamtnu1 + Val(mygrid.TextMatrix(i, 22))
                      
                subtotkg = subtotkg + Val(mygrid.TextMatrix(i, 23))
                alltotkg = alltotkg + Val(mygrid.TextMatrix(i, 23))
                    
                subtotamtnu2 = subtotamtnu2 + Val(mygrid.TextMatrix(i, 24))
                alltotamtnu2 = alltotamtnu2 + Val(mygrid.TextMatrix(i, 24))
                    
                subtottotamtnu = subtottotamtnu + Val(mygrid.TextMatrix(i, 25))
                alltottotamtnu = alltottotamtnu + Val(mygrid.TextMatrix(i, 25))
           
        End If
        If mygrid.TextMatrix(i, 28) = "S" Then
                mygrid.TextMatrix(i, 1) = ""
                formatsubtot (i)
                initvariablessub
                tt = tt + 1
        End If
Next

If txtindecator.Text = "S" Then
If mygrid.TextMatrix(i, 0) = mygrid.TextMatrix(i - 1, 0) Then
mygrid.TextMatrix(i, 0) = i
End If
End If

If txtindecator.Text = "" Then
If mygrid.TextMatrix(i - 1, 28) = "T" Then
If mygrid.TextMatrix(i - 1, 0) = mygrid.TextMatrix(i - 2, 0) Then
mygrid.TextMatrix(i, 0) = i
End If
formatalltot (i - 1)
End If
Else
formatalltot (i)
End If

mygrid.MergeCells = flexMergeFree
mygrid.MergeCol(1) = True
mygrid.MergeCells = flexMergeFree
mygrid.MergeCol(26) = True
txtindecator.Text = "S"
End Sub
Private Sub formatsubtot(i As Integer)
Dim refilcrate As Integer
                mygrid.TextMatrix(i, 9) = subtot
                mygrid.col = 9
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 10) = subtotplant
                mygrid.col = 10
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 12) = subtotbcrate
                mygrid.col = 12
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True



                mygrid.TextMatrix(i, 13) = subtotecrate
                mygrid.col = 13
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                mygrid.TextMatrix(i, 14) = subtotbno
                mygrid.col = 14
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                
                modval = mygrid.TextMatrix(i, 14)
                mmod = modval Mod 35
                If (mmod > 1) Then
                refilcrate = Int(subtotbno / 35) + 1
                Else
                refilcrate = Int(subtotbno / 35)
                End If
                
                
If (chkrefill.Value = 1) Then
                mygrid.TextMatrix(i, 11) = Int(subtotplant / 35) + 1
                mygrid.col = 11
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
Else
                mygrid.TextMatrix(i, 11) = subtotcrateno + refilcrate
                mygrid.col = 11
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
End If
                
                mygrid.TextMatrix(i, 15) = subtotplno
                 mygrid.col = 15
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                mygrid.TextMatrix(i, 16) = subtotcrate
                 mygrid.col = 16
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 17) = Round(subtotssp, 0)
                mygrid.col = 17
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 18) = Round(subtotmop, 0)
                mygrid.col = 18
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 19) = Round(subtoturea, 0)
                mygrid.col = 19
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 20) = Round(subtotdolomite, 0)
                mygrid.col = 20
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                  mygrid.TextMatrix(i, 21) = subtotkg1
                mygrid.col = 21
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                mygrid.TextMatrix(i, 22) = subtotamtnu1
                mygrid.col = 22
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 23) = subtotkg
                mygrid.col = 23
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 24) = subtotamtnu2
                mygrid.col = 24
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 25) = subtottotamtnu
                mygrid.col = 25
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                  
                mygrid.col = 26
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
End Sub
Private Sub formatalltot1(i As Integer)

     
                mygrid.col = 9
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


               
                mygrid.col = 10
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                mygrid.col = 11
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


             
                mygrid.col = 12
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True



                
                mygrid.col = 13
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
                
                mygrid.col = 14
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
           
                 mygrid.col = 15
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
             
                 mygrid.col = 16
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


           
                mygrid.col = 17
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

              
                mygrid.col = 18
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                
                mygrid.col = 19
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                 
                mygrid.col = 20
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


                 
                mygrid.col = 21
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

               
                mygrid.col = 22
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

              
                mygrid.col = 23
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                mygrid.col = 24
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


              
                mygrid.col = 25
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
                  
'                Mygrid.col = 26
'                Mygrid.row = i
'                Mygrid.CellBackColor = vbGreen
'                Mygrid.CellFontBold = True
End Sub
Private Sub formatalltot(i As Integer)

                mygrid.TextMatrix(i, 9) = alltot
                mygrid.col = 9
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 10) = alltotplant
                mygrid.col = 10
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                mygrid.TextMatrix(i, 11) = alltotcrateno
                mygrid.col = 11
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 12) = alltotbcrate
                mygrid.col = 12
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True



                mygrid.TextMatrix(i, 13) = alltotecrate
                mygrid.col = 13
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
                mygrid.TextMatrix(i, 14) = alltotbno
                mygrid.col = 14
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
                mygrid.TextMatrix(i, 15) = alltotplno
                 mygrid.col = 15
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
                mygrid.TextMatrix(i, 16) = alltotcrate
                 mygrid.col = 16
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


                 mygrid.TextMatrix(i, 17) = Round(alltotssp, 0)
                mygrid.col = 17
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 18) = Round(alltotmop, 0)
                mygrid.col = 18
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 19) = Round(alltoturea, 0)
                mygrid.col = 19
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 20) = Round(alltotdolomite, 0)
                mygrid.col = 20
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


                  mygrid.TextMatrix(i, 21) = alltotkg1
                mygrid.col = 21
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                mygrid.TextMatrix(i, 22) = alltotamtnu1
                mygrid.col = 22
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 23) = alltotkg
                mygrid.col = 23
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True

                 mygrid.TextMatrix(i, 24) = alltotamtnu2
                mygrid.col = 24
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True


                mygrid.TextMatrix(i, 25) = alltottotamtnu
                mygrid.col = 25
                mygrid.row = i
                mygrid.CellBackColor = vbGreen
                mygrid.CellFontBold = True
                
                  
'                Mygrid.col = 26
'                Mygrid.row = i
'                Mygrid.CellBackColor = vbGreen
'                Mygrid.CellFontBold = True
End Sub
Private Sub formatsubtot1(i As Integer)
    
                mygrid.col = 9
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


             
                mygrid.col = 10
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

  
                mygrid.col = 11
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


              
                mygrid.col = 12
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True



                
                mygrid.col = 13
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                 mygrid.col = 14
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                 mygrid.col = 15
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                 mygrid.col = 16
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                mygrid.col = 17
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                mygrid.col = 18
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                mygrid.col = 19
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                
                mygrid.col = 20
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


                 
                mygrid.col = 21
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                
                mygrid.col = 22
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                
                mygrid.col = 23
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True

                
                mygrid.col = 24
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True


              
                mygrid.col = 25
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
                
                 mygrid.col = 26
                mygrid.row = i
                mygrid.CellBackColor = vbRed
                mygrid.CellFontBold = True
End Sub

Private Sub initvariables()
                 subtot = 0
                 alltot = 0
                 subtotplant = 0
                 alltotplant = 0
                 subtotcrateno = 0
                 alltotcrateno = 0
                 subtotbcrate = 0
                 alltotbcrate = 0
                 subtotecrate = 0
                 alltotecrate = 0
                 subtotbno = 0
                 alltotbno = 0
                 subtotplno = 0
                 alltotplno = 0
                 subtotcrate = 0
                 alltotcrate = 0
                 subtotssp = 0
                 alltotssp = 0
                 subtotmop = 0
                 alltotmop = 0
                 subtoturea = 0
                 alltoturea = 0
                 subtotdolomite = 0
                 alltotdolomite = 0
                 subtotkg1 = 0
                 alltotkg1 = 0
                 subtotamtnu1 = 0
                 alltotamtnu1 = 0
                 subtotkg = 0
                 alltotkg = 0
                 subtotamtnu2 = 0
                 alltotamtnu2 = 0
                 subtottotamtnu = 0
                 alltottotamtnu = 0
End Sub
Private Sub initvariablessub()
                 subtot = 0
                 
                 subtotplant = 0
                 subtotcrateno = 0
                 
                 subtotbcrate = 0
                 
                 subtotecrate = 0
                
                 subtotbno = 0
                 
                 subtotplno = 0
                
                 subtotcrate = 0
                 
                 subtotssp = 0
                
                 subtotmop = 0
                 
                 subtoturea = 0
                 
                 subtotdolomite = 0
                 
                 subtotkg1 = 0
                
                 subtotamtnu1 = 0
                 
                 subtotkg = 0
                
                 subtotamtnu2 = 0
                 
                 subtottotamtnu = 0
                
End Sub

Private Sub removeblanks()
Dim i As Integer
For i = 1 To mygrid.rows - 1
'If Len(Mygrid.TextMatrix(i, 1)) = 0 Then Exit For
If mygrid.row > 0 And Len(mygrid.TextMatrix(i, 1)) = 0 Then
      mygrid.RemoveItem mygrid.row
      mygrid.AddItem ""
   Else
      Beep
      Beep
   End If
Next
End Sub

Private Sub Command6_Click()
Frame1.Visible = False
Operation = "OPEN"
cmdload.Enabled = True
cbotrnid.Enabled = True
cmdsave.Enabled = True
txtdno.Text = ""
End Sub

Private Sub Command8_Click()
Dim RSTR As New ADODB.Recordset

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(cast(trnid as char) ,' ',distributionname) as dname,trnid  from tblplantdistributionheader order by trnid", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "dname"
cbotrnid.BoundColumn = "trnid"

End Sub

Private Sub Command9_Click()
'addgrid


Dim i As Long
For i = 0 To LSTPR.ListCount - 1
    LSTPR.Selected(i) = True
Next
End Sub




Private Sub Form_Load()
Dim RSTR As New ADODB.Recordset
maxDistNo = 0
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(cast(trnid as char) ,' ',distributionname,' ',cast(year as char),' ',cast(mnth as char)) as dname,trnid  from tblplantdistributionheader where status='ON' order by trnid desc", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "dname"
cbotrnid.BoundColumn = "trnid"

mygrid.Visible = False
Dim rs As New ADODB.Recordset

Set rs = Nothing

rs.Open "select DZONGKHAGCODE,DZONGKHAGNAME from tbldzongkhag Order by DZONGKHAGCODE", MHVDB, adOpenStatic
With rs
Do While Not .EOF
   DZLIST.AddItem Trim(!DZONGKHAGNAME) + " | " + !DZONGKHAGCODE
   .MoveNext
Loop
End With
txtdno.Text = ""

End Sub











Private Sub mygrid_Click()
'If Mygrid.col = 26 And Mygrid.TextMatrix(Mygrid.row - 1, 28) = "S" And Mygrid.TextMatrix(Mygrid.row, 28) <> "T" Then
'
'      Text1.Top = Mygrid.Top + Mygrid.CellTop
'      Text1.Left = Mygrid.Left + Mygrid.CellLeft
'      Text1.Width = Mygrid.CellWidth
'      Text1.Height = Mygrid.CellHeight
'      Text1 = Mygrid.Text
'      Text1.Visible = True
'      Text1.SetFocus
'      Text1.SelLength = 1
'
'End If
'
'Dim mrow, MCOL As Integer
''txtselected.Visible = False
''ItemGrd.ColWidth(3) = 750
''If Not ValidRow And CurrRow <> ItemGrd.row Then
''   ItemGrd.row = CurrRow
''   Exit Sub
''End If
'mrow = Mygrid.row
'MCOL = Mygrid.col
'If mrow = 0 Then Exit Sub
'If mrow > 1 And Len(Mygrid.TextMatrix(mrow - 1, 4)) = 0 Then
'   Beep
'   Exit Sub
'End If
'Mygrid.TextMatrix(CurrRow, 0) = CurrRow
'CurrRow = mrow
'Mygrid.TextMatrix(CurrRow, 0) = Chr(174)
'
'Select Case MCOL
'
'
'       Case 1
'        txtdno.Left = Mygrid.Left + Mygrid.CellLeft
'        txtdno.Width = Mygrid.CellWidth
'        txtdno.Height = Mygrid.CellHeight
'       If Len(Mygrid.TextMatrix(mrow, 1)) > 0 Then
'            txtdno.Top = Mygrid.Top + Mygrid.CellTop
'            txtdno = Mygrid.Text
'            txtdno.Visible = True
'            txtdno.SetFocus
'       End If
'
'
'    End Select



End Sub

Private Sub mygrid_DblClick()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim rsland As New ADODB.Recordset
Dim mdgt As String
Dim myacre As Double
Dim i As Integer
If Operation = "ADD" Then Exit Sub
If mygrid.col = 2 And mygrid.row <> mygrid.rows - 1 And mygrid.TextMatrix(mygrid.row, 28) <> "S" Then
If MsgBox("Do you want add new row for sub total", vbQuestion + vbYesNo) = vbYes Then
'Mygrid.Rows = Mygrid.Rows + 1
mygrid.AddItem "", mygrid.row
'InsertRow Mygrid, Mygrid.row
mygrid.TextMatrix(mygrid.row, 0) = mygrid.row
mygrid.TextMatrix(mygrid.row, 28) = "S"
addgrid
Else


End If
'InsertRow Mygrid, Mygrid.row
End If
'
If mygrid.col = 9 And mygrid.row <> mygrid.rows - 1 And mygrid.TextMatrix(mygrid.row, 28) <> "S" And mygrid.TextMatrix(mygrid.row, 9) > 0 And mygrid.TextMatrix(mygrid.row, 31) > 0 Then
mdgt = ""
If MsgBox("Do you want update the land", vbQuestion + vbYesNo) = vbYes Then
' add and fetch things


'----------------
                                        i = mygrid.row
                                        mdgt = mygrid.TextMatrix(i, 5)
                                        
                                        
    SQLSTR = " SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE, " _
& " SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(acre) AS REGLAND,village,phone1 FROM " _
& " tblfarmer A,tbllandregdetail B WHERE A.status not in('D','R','C') and plantedstatus='N'  " _
& " and A.IDFARMER=B.farmercode  and substring(idfarmer,10,1)='F' and IDFARMER='" & mdgt & "' group by idfarmer "
              
'        SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(REGLAND) AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status not in('D','R','C')and B.status not in('D','R','C') and plantedstatus='N'  and A.IDFARMER=B.FARMERID and IDFARMER ='" & mdgt & "'"
'        SQLSTR = SQLSTR & "  " & "group by idfarmer "
'SQLSTR = SQLSTR & " union  SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE, " _
'& " SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(acre) AS REGLAND,village,phone1 FROM " _
'& " tblfarmer A,tbllandregdetail B WHERE A.status not in('D','R','C') and plantedstatus='N'  " _
'& " and A.IDFARMER=B.farmercode  and  IDFARMER='" & mdgt & "' group by idfarmer "

SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE, " _
& " SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(acre) AS REGLAND,village,phone1 FROM " _
& " tblfarmer A,tbllandregdetail B WHERE A.status not in('D','R','C') and plantedstatus='N'  " _
& " and A.IDFARMER=B.farmercode  and  IDFARMER='" & mdgt & "' group by idfarmer"
                                        
                                 Set rsland = Nothing
                                        
                                 rsland.Open SQLSTR, MHVDB
                                 If rsland.EOF <> True Then
                                 myacre = rsland!regland
                                 Else
                                 myacre = 0
                                 Exit Sub
                                 End If
                                 
                                        mygrid.TextMatrix(i, 9) = Round(myacre, 2)
                                        Set rs1 = Nothing
                                        myStr = "select ifnull(sum(b),0) as b,ifnull(sum(e),0) as e,ifnull(sum(p1),0) as p1, ifnull(sum(n),0) as n from refillin where id in(  '" & mygrid.TextMatrix(i, 32) & "' )"
                                        rs1.Open myStr, MHVDB
                                                                       
                                         If rs1.EOF <> True Then
                                                btype = Round(rs1!b, mrnd)
                                                etype = Round(rs1!e, mrnd)
                                                Set RS2 = Nothing
                                                RS2.Open "select * from tbldistformula where fid='" & mygrid.TextMatrix(i, 31) & "'", MHVDB
                                                If RS2.EOF <> True Then
'                                                mygrid.TextMatrix(i, 10) = Round(rs1!p1 + rs1!n + btype + etype, mrnd)

'
'                                                       mygrid.TextMatrix(i, 13) = etype
'                                                       mygrid.TextMatrix(i, 10) = Round(btype + Val(mygrid.TextMatrix(i, 13)) + rs1!p1 + rs1!n, mrnd)
'                                                       mygrid.TextMatrix(i, 11) = (Val(mygrid.TextMatrix(i, 10)) - (Val(mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
'                                                       mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
'                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
                                                        mygrid.TextMatrix(i, 14) = rs1!p1 + rs1!b
                                                        mygrid.TextMatrix(i, 10) = Round(((Val(mygrid.TextMatrix(i, 9)) * RS2!totalplant)), 0) + Val(mygrid.TextMatrix(i, 14))
                                                        mygrid.TextMatrix(i, 11) = mygrid.TextMatrix(i, 10) 'Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))), 0) ' Val(mygrid.TextMatrix(i, 11)) + Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))) / rs!crateno, 0)
                                                        modval = mygrid.TextMatrix(i, 11)
                                                        mmod = modval Mod RS2!crateno
                                                        If (mmod > 17) Then
                                                        mygrid.TextMatrix(i, 11) = ((modval - mmod) / RS2!crateno) + 1
                                                        Else
                                                        mygrid.TextMatrix(i, 11) = ((modval - mmod) / RS2!crateno)
                                                        End If
                                                        mygrid.TextMatrix(i, 12) = Round((Val(mygrid.TextMatrix(i, 11)) * RS2!crateno * RS2!bcrate), 0) 'Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate) / rs!crateno, 0)
                                                        modval = mygrid.TextMatrix(i, 12)
                                                        mmod = modval Mod RS2!crateno
                                                        If (Val(mygrid.TextMatrix(i, 9)) <> 0) Then
                                                        If (mmod > 17) Then
                                                        
                                                        mygrid.TextMatrix(i, 12) = ((modval - mmod) / RS2!crateno) + 1
                                                        Else
                                                        mygrid.TextMatrix(i, 12) = ((modval - mmod) / RS2!crateno)
                                                        End If
                                                        Else
                                                        mygrid.TextMatrix(i, 11) = 0
                                                        mygrid.TextMatrix(i, 12) = 0
                                                        mygrid.TextMatrix(i, 13) = 0
                                                        End If
                                                        
                                                                                         
                                                End If
                                                mygrid.TextMatrix(i, 15) = Round(rs1!p1, mrnd) '
                                                mygrid.TextMatrix(i, 16) = Round(rs1!n, 0) 'Round(rs1!n / RS2!crateno, 0)
                                                polycont = polycont + Round(rs1!p1, mrnd) + Round(rs1!n, mrnd)
                                                mygrid.TextMatrix(i, 29) = "O"
                                                
                                             
                                                 
                                         End If
                                         
                                         
'                                                Set rs = Nothing
'                                                rs.Open "select * from tbldistformula where fid='" & mygrid.TextMatrix(i, 31) & "'", MHVDB
'                                                If rs.EOF <> True Then ' if 1
'                                                    mygrid.TextMatrix(i, 10) = Val(mygrid.TextMatrix(i, 10)) + Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant)), 0)
'                                                    mygrid.TextMatrix(i, 15) = Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant) * rs!ncrate), 0)
'                                                    mygrid.TextMatrix(i, 11) = mygrid.TextMatrix(i, 10) 'Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))), 0) ' Val(mygrid.TextMatrix(i, 11)) + Round((Val(mygrid.TextMatrix(i, 10)) - Val(mygrid.TextMatrix(i, 15)) - Val(mygrid.TextMatrix(i, 16))) / rs!crateno, 0)
'                                                    modval = mygrid.TextMatrix(i, 11)
'                                                    mmod = modval Mod rs!crateno
'                                                    If (mmod > 17) Then
'                                                        mygrid.TextMatrix(i, 11) = ((modval - mmod) / rs!crateno) + 1
'                                                    Else
'                                                        mygrid.TextMatrix(i, 11) = ((modval - mmod) / rs!crateno)
'                                                    End If
'                                                    mygrid.TextMatrix(i, 15) = Round(Val(mygrid.TextMatrix(i, 15)) / rs!crateno, 0)
'
'                                                    If mygrid.TextMatrix(i, 29) <> "O" Then
'
'                                                    mygrid.TextMatrix(i, 12) = Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate), 0) 'Round((Val(mygrid.TextMatrix(i, 11)) * rs!crateno * rs!bcrate) / rs!crateno, 0)
'modval = mygrid.TextMatrix(i, 12)
'mmod = modval Mod rs!crateno
'If (mmod > 17) Then
'mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno) + 1
'Else
'mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno)
'End If
'Else
'If Val(mygrid.TextMatrix(i, 9)) = 0 Then
'mygrid.TextMatrix(i, 12) = Round(((Val(mygrid.TextMatrix(i, 12)))) / rs!crateno, 0) 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)
'Else
'mygrid.TextMatrix(i, 12) = mygrid.TextMatrix(i, 12) + Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant) * rs!bcrate), 0) 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)
'
'modval = mygrid.TextMatrix(i, 12)
'mmod = modval Mod rs!crateno
'If (mmod > 17) Then
'mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno) + 1
'Else
'mygrid.TextMatrix(i, 12) = ((modval - mmod) / rs!crateno)
'End If
'
'End If
'
'
'End If

'-------
If Mid(mygrid.TextMatrix(i, 5), 10, 1) <> "G" Or Mid(mygrid.TextMatrix(i, 5), 10, 1) <> "C" Then
        If mygrid.TextMatrix(i, 29) <> "O" Then
            mygrid.TextMatrix(i, 13) = mygrid.TextMatrix(i, 11) - mygrid.TextMatrix(i, 12) - mygrid.TextMatrix(i, 15)  ' Round((mygrid.TextMatrix(i, 11) * rs!ecrate), 0)  'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)Round((mygrid.TextMatrix(i, 11) * rs!crateno * rs!ecrate) / rs!crateno, 0) 'Val(mygrid.TextMatrix(i, 11)) * rs!crateno - rs!crateno - Val(mygrid.TextMatrix(i, 12) * rs!crateno)
      Else
If Val(mygrid.TextMatrix(i, 9)) = 0 Then
mygrid.TextMatrix(i, 13) = Round(((Val(mygrid.TextMatrix(i, 13)))), 0) / RS2!crateno 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)
Else
mygrid.TextMatrix(i, 13) = mygrid.TextMatrix(i, 11) - mygrid.TextMatrix(i, 12) - mygrid.TextMatrix(i, 15) 'mygrid.TextMatrix(i, 13) + Round(((Val(mygrid.TextMatrix(i, 9)) * rs!totalplant) * rs!ecrate), 0) 'Round(((Val(mygrid.TextMatrix(i, 11)) - Val(mygrid.TextMatrix(i, 30))) * rs!bcrate), 0)

End If
        End If

End If





      
        
        If Mid(mygrid.TextMatrix(i, 5), 10, 1) = "G" Or Mid(mygrid.TextMatrix(i, 5), 10, 1) = "C" Then

       End If


'=====


                                                        mygrid.TextMatrix(i, 17) = Round((Val(mygrid.TextMatrix(i, 10)) * RS2!ssp), 2)
                                                        mygrid.TextMatrix(i, 18) = Round((Val(mygrid.TextMatrix(i, 10)) * RS2!mop), 2)
                                                        mygrid.TextMatrix(i, 19) = Round((Val(mygrid.TextMatrix(i, 10)) * RS2!urea), 2)
                                                        mygrid.TextMatrix(i, 20) = Round((Val(mygrid.TextMatrix(i, 10)) * RS2!dolomite), 2)
                                                       
'                                                       mygrid.TextMatrix(i, 17) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!ssp), 2)
'                                                       mygrid.TextMatrix(i, 18) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!mop), 2)
'                                                       mygrid.TextMatrix(i, 19) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!urea), 2)
'                                                       mygrid.TextMatrix(i, 20) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!dolomite), 2)
                                                       mygrid.TextMatrix(i, 21) = Round(Val(mygrid.TextMatrix(i, 17)) + Val(mygrid.TextMatrix(i, 18)) + Val(mygrid.TextMatrix(i, 19)) + Val(mygrid.TextMatrix(i, 20)), 0)
                                                       mygrid.TextMatrix(i, 22) = Round(Val(mygrid.TextMatrix(i, 17) * RS2!sspperkg) + Val(mygrid.TextMatrix(i, 18) * RS2!mopperkg) + Val(mygrid.TextMatrix(i, 19) * RS2!ureaperkg) + Val(mygrid.TextMatrix(i, 20) * RS2!dolomiteperkg), 0)
'                                                       mygrid.TextMatrix(i, 23) = Round((((Val(mygrid.TextMatrix(i, 11)) * 35) + Val(mygrid.TextMatrix(i, 16))) * rs!kg), 0)
                                                       mygrid.TextMatrix(i, 23) = Round((Val(mygrid.TextMatrix(i, 10)) * RS2!kg) + 0.00000001, 0)
                                                       mygrid.TextMatrix(i, 24) = Round((mygrid.TextMatrix(i, 23) * RS2!amountnu), 0)
                                                       If Val(mygrid.TextMatrix(i, 23)) < 0 Then
                                                       mygrid.TextMatrix(i, 23) = 0
                                                       mygrid.TextMatrix(i, 24) = 0
                                                       End If
                                                       mygrid.TextMatrix(i, 25) = Val(mygrid.TextMatrix(i, 22)) + Val(mygrid.TextMatrix(i, 24))
                                                    
                                                End If 'end 1
                                         
                                         
                                         

'-------------------



addgrid
Else


End If
'InsertRow Mygrid, Mygrid.row
'End If




' temporary below code once down with below, uncoment above code

'Dim i As Integer
'Dim olddn As Integer
'Dim myinput As String
'If Mygrid.col = 1 And Len(Mygrid.TextMatrix(Mygrid.row, 1)) > 0 Then
'myinput = InputBox("Enter The New Distribution No.")
'            If Not IsNumeric(myinput) Then
'            MsgBox "Invalid number,Double Click again to enable the input box."
'            Else
'            olddn = Mygrid.TextMatrix(Mygrid.row, 1)
''            Mygrid.TextMatrix(Mygrid.row, 1) = CInt(myinput)
''            Mygrid.TextMatrix(Mygrid.row, 27) = CInt(myinput)
'
'            i = 0
'           For i = Mygrid.row To Mygrid.Rows - 1
'           If Mygrid.TextMatrix(i, 28) = "S" Then Exit For
'           If Val(Mygrid.TextMatrix(i, 1)) = olddn Then
'           Mygrid.TextMatrix(i, 1) = CInt(myinput)
'            Mygrid.TextMatrix(i, 27) = CInt(myinput)
'
'           End If
'
'
'           Next
'           If Mygrid.TextMatrix(i, 28) = "S" Then
'          Mygrid.TextMatrix(i, 27) = CInt(myinput)
'           End If
'
'            End If
'            Else
'
'End If
'




End Sub

Private Sub Mygrid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 2 Then

   If mygrid.row > 0 And mygrid.TextMatrix(mygrid.row, 28) <> "T" Then 'And Len(Mygrid.TextMatrix(cURRrOW, 1)) > 0 Then
   If MsgBox("Do you want to delete this row", vbQuestion + vbYesNo) = vbYes Then
      mygrid.RemoveItem mygrid.row
      addgrid
      End If
   Else
      Beep
      Beep
   End If
End If
End Sub

Private Sub mygrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'turn off the highlight feature
    mygrid.HighLight = flexHighlightNever
    mygrid.FocusRect = flexFocusHeavy
    'get the desired row to move
    RowToMove = mygrid.MouseRow
    'this lets us know we are clicking
    ButtonDown = True
    Label1.Caption = "Preparing to Move Row # " & RowToMove
End Sub


Private Sub mygrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Operation = "ADD" Then Exit Sub

    If ButtonDown = False Then
        'we haven't clicked yet, so just advise
        '     the row we are on
        Label1.Caption = "Click Mouse button to Move Row # " & mygrid.MouseRow
        Exit Sub
    End If
    'we have clicked, so advise of the start
    '     and current row


    If mygrid.MouseRow <> RowToMove Then
        Label1.Caption = "Release Mouse button to Move Row # " & RowToMove & " to " & mygrid.MouseRow
    End If
End Sub


Private Sub mygrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Operation = "ADD" Then Exit Sub
    
    Dim lRet As Long
    Dim RowClip$
    Dim MoveClip$


    With mygrid
        DestRow = .MouseRow
        'check if we are still in the same row a
        '     s we clicked
        If DestRow = RowToMove Then Exit Sub
        'this is just a confirmation, you don't
        '     really need this but it shows you it wor
        '     ked
        lRet = MsgBox("Do you want to move Row # " & RowToMove & " to " & DestRow, vbQuestion + vbYesNo, "Move Row?")
        sourceDno = mygrid.TextMatrix(RowToMove, 1)
        DestDno = mygrid.TextMatrix(DestRow, 1)

        If lRet = vbYes And mygrid.TextMatrix(mygrid.row, 28) <> "S" And mygrid.TextMatrix(mygrid.row, 28) <> "T" Then
            .Redraw = False
            'select the whole row for the cell click
            '     ed
            .row = RowToMove
            .col = 0
            .RowSel = RowToMove
            .ColSel = .cols - 1
            'copy the whole row's data to a string
            RowClip$ = .clip
            'delete the moved row
            .RemoveItem RowToMove
            'put the moved data to the new row
            .AddItem RowClip$, DestRow
            .Redraw = True
            'Mygrid.TextMatrix(Mygrid.row, 1) = ""
            
'        If sourceDno <> DestDno Then
'         Mygrid.TextMatrix(DestRow, 1) = DestDno
'Mygrid.MergeCells = flexMergeFree
'Mygrid.MergeCol(1) = True
'Mygrid.MergeCells = flexMergeFree
'Mygrid.MergeCol(26) = True
'        End If
            addgrid
        End If
    End With
    'release the variable that says we have
    '     the button down
    ButtonDown = False
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
On Error Resume Next
Dim s, T, MYROW As Integer
Text1.Visible = False
MYROW = mygrid.row
For s = 1 To mygrid.rows - 1
If mygrid.TextMatrix(mygrid.row, 1) <> mygrid.TextMatrix(MYROW, 27) Then Exit Sub
mygrid.TextMatrix(MYROW, 26) = Text1.Text
MYROW = MYROW + 1
Next

mygrid.MergeCells = flexMergeFree
mygrid.MergeCol(26) = True

End Sub
Private Sub VALIDATESCHEDULE()
On Error Resume Next
Dim s, T, MYROW As Integer
Text1.Visible = False
MYROW = mygrid.row
For s = 1 To mygrid.rows - 1
If mygrid.TextMatrix(mygrid.row, 1) <> mygrid.TextMatrix(MYROW, 27) Then Exit Sub
mygrid.TextMatrix(MYROW, 26) = Text1.Text
MYROW = MYROW + 1
Next

mygrid.MergeCells = flexMergeFree
mygrid.MergeCol(26) = True
End Sub


Private Sub txtdno_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(mygrid.TextMatrix(CurrRow, 0)) > 0 Then
If Not IsNumeric(txtdno) Then
   Beep
   MsgBox "Enter a valid No."
   ValidRow = False
   Cancel = True
   Exit Sub

Else
  
   mygrid.TextMatrix(CurrRow, 1) = Val(txtdno.Text)
   mygrid.TextMatrix(CurrRow, 27) = Val(txtdno.Text)
   ValidRow = True
   
End If
End If
If txtdno.Visible = True Then txtdno.Visible = False
End Sub

Private Sub txtfcode_DblClick()
Dim i As Integer
For i = 0 To LSTPR.ListCount - 1
If txtfcode.Text = Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) Then
LSTPR.Selected(i) = True
End If
Next
txtfcode.Text = ""
cratecnt
End Sub
Private Sub cratecnt()
Dim i As Integer
mcratecnt = 0
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
      mcratecnt = mcratecnt + 1
         End If
    Next
    txtcratecnt.Text = mcratecnt
End Sub

Private Sub txtfcode_KeyPress(KeyAscii As Integer)
If InStr(1, "DGTFC0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


Private Sub loadgridold()
'On Error Resume Next
Dim polycont As Long
Dim mydgt As String
Dim cnt As Integer
mydgt = ""
Dim morderstr As String
Dim muk As Integer
muk = 0
Dim rs1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim s As Integer
Dim SQLSTR As String

Set rs1 = Nothing
rs1.Open "select max(distno) as maxid from tblplantdistributiondetail", MHVDB
maxDistNo = IIf(IsNull(rs1!MaxId), 0, rs1!MaxId) + 1
txtdno.Text = maxDistNo
Set rs1 = Nothing

morderstr = ""
mygrid.Clear
mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^P    |^P1 (Nos.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^"
etype = 0
ptype = 0
SQLSTR = ""
Dim i, j As Integer
Dim rs As New ADODB.Recordset
Dim rsadd As New ADODB.Recordset
Set rsadd = Nothing
mchk = True
j = 0
Dzstr = ""
morderstr = ""
If chkpriority.Value = 0 Then
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
ElseIf chkpriority.Value = 1 And chkunplanned.Value = 0 Then
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
    mm = Split(LSTPR.List(i), "|", -1, 1)
       Dzstr = Dzstr & "'" & Mid(mm(0), 1, 3) & Mid(mm(1), 1, 3) & Mid(mm(2), 1, 3) & "',"  ' + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = LSTPR.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE LOCATION TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
Else

Dzstr = ""
For i = 0 To LSTPR.ListCount - 1
    If LSTPR.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
          Exit Sub
       End If
    End If
Next
End If



If Len(Dzstr) > 0 Then
morderstr = Left(Dzstr, Len(Dzstr) - 1)
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"

Else
   MsgBox "DZONGKHAG NOT SELECTED !!!"
   Exit Sub
End If

'Dzstr = ""
'For i = 0 To LSTPR.ListCount - 1
'    If LSTPR.Selected(i) Then
'       Dzstr = Dzstr + "'" + Trim(Mid(LSTPR.List(i), InStr(1, LSTPR.List(i), "|") + 1)) + "',"
'       Mcat = DZLIST.List(i)
'       j = j + 1
'    End If
'    If RepName = "5" Then
'       If j > 1 Then
'          MsgBox "SELECT ATLEAST ONE DZONGKHAG TO VIEW THIS REPORT."
'          Exit Sub
'       End If
'    End If
'Next
'If Len(Dzstr) > 0 Then
'   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
'
'End If


mygrid.Clear
'mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^P    |^P1 (Nos.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^"
mygrid.FormatString = "S/N|^D\N|Dzongkhag|^Gewog       |^Tshowog    |^Farmer Code               |^Farmer Name |^ Contact# |^Village|^Land(Acre)|^Total Plants|^Crates #|^B(Crate)|^E (Crate)|^P    |^N (Crt.)|^N (Nos.) |^SSP(Kg.)|^MOP (Kg.)|^Urea(Kg.)|^Dolomite (Kg.)|^Total(Kg.)|^Amount(Nu.)|^Kg.|^Amount(Nu.)|^Total Amount(Nu.)|^Schedule Date,Vehicle & Team Captency|^ |^|^|^|^|^"




If chkpriority.Value = 0 And chkPolinizer.Value = 0 And chkrefill.Value = 0 Then
    If chkgrf.Value = 0 And chkcf.Value = 0 Then
            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and plantedstatus='N' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9),IDFARMER"

        
       
        
        ElseIf chkgrf.Value = 1 Then
            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='G' and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9),IDFARMER"
    Else
    SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='C' and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9),IDFARMER"
    
    End If
    Else
        myfamercount = 0
        If chkgrf.Value = 0 And chkcf.Value = 0 And chkPolinizer.Value = 0 Then
        SQLSTR = ""
        If chkunplanned = 1 Then
         SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND (IDFARMER)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "


       MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
       SQLSTR = ""
       'SQLSTR = "SELECT SUBSTRING(farmercode,1,3) AS DZCODE,SUBSTRING(farmercode,4,3) AS GECODE,SUBSTRING(farmercode,7,3) AS TSCODE,farmercode,FARMERNAME,0 AS REGLAND,village,phone1 FROM tblfarmer A,refillin B WHERE A.status='A' and A.IDFARMER=B.farmercode and substring(farmercode,10,1)='F' and SUBSTRING(farmercode,1,9)IN  " & Dzstr & " and farmercode not in(select idfarmer from tbldistpreparetion)"
        '    SQLSTR = SQLSTR & "  " & "group by farmercode order by  FIELD(SUBSTRING(farmercode,1,9), " & morderstr & ") "
       'MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
      '            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
'            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
  SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1)" _
           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1 from " _
& "refillin a,tblfarmer b where idfarmer=farmercode and a.status='ON' and farmercode not in(select idfarmer from tbldistpreparetion)"
MHVDB.Execute SQLSTR
  SQLSTR = "SELECT * from tbldistpreparetion order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "

ElseIf chkrefill.Value = 1 Then
      MHVDB.Execute "delete from  tbldistpreparetion"
      SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1,refillplant,totfillindemand)" _
           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1,totplants,totfillin from " _
           & "mhv.vrefillin a,mhv.tblfarmer b where idfarmer=farmercode and SUBSTRING(IDFARMER,1,3) IN  " & Dzstr & " and a.status='ON' and a.isfinalized='Yes'"
           MHVDB.Execute SQLSTR
      SQLSTR = "SELECT * from tbldistpreparetion order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "

Else

        SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(REGLAND) AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status not in('D','R','C')and B.status not in('D','R','C') and plantedstatus='N'  and A.IDFARMER=B.FARMERID and substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
        SQLSTR = SQLSTR & "  " & "group by idfarmer "
SQLSTR = SQLSTR & " union  SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE, " _
& " SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,sum(acre) AS REGLAND,village,phone1 FROM " _
& " tblfarmer A,tbllandregdetail B WHERE A.status not in('D','R','C') and plantedstatus='N'  " _
& " and A.IDFARMER=B.farmercode  and substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr & " group by idfarmer order by " _
& " FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ")"

MHVDB.Execute "delete from  tbldistpreparetion"

       MHVDB.Execute "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1) " & SQLSTR
       SQLSTR = ""
       'SQLSTR = "SELECT SUBSTRING(farmercode,1,3) AS DZCODE,SUBSTRING(farmercode,4,3) AS GECODE,SUBSTRING(farmercode,7,3) AS TSCODE,farmercode,FARMERNAME,0 AS REGLAND,village,phone1 FROM tblfarmer A,refillin B WHERE A.status='A' and A.IDFARMER=B.farmercode and substring(farmercode,10,1)='F' and SUBSTRING(farmercode,1,9)IN  " & Dzstr & " and farmercode not in(select idfarmer from tbldistpreparetion)"
        '    SQLSTR = SQLSTR & "  " & "group by farmercode order by  FIELD(SUBSTRING(farmercode,1,9), " & morderstr & ") "
       'MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
      '            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
'            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
  SQLSTR = "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1)" _
           & " select substring(farmercode,1,3) dzcode,substring(farmercode,4,3) gecode,substring(farmercode,7,3) tscode,farmercode,farmername,0 regland,village,phone1 from " _
& "vrefillin a,tblfarmer b where idfarmer=farmercode and a.status='ON' and farmercode not in(select idfarmer from tbldistpreparetion)"
MHVDB.Execute SQLSTR
  SQLSTR = "SELECT * from tbldistpreparetion order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "
  

End If
        
        

        
        
        ElseIf chkgrf.Value = 1 Then
        
            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1,sum(plantqty)polinizercrate  FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='G' and  SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
        ElseIf chkPolinizer.Value = 1 Then
        
    'added by kinzang ## begin
    '    MHVDB.Execute "insert into mhv.tblpolinizer(farmercode,plantqty,myear,pvariety,originalplantqty,flag) select a.farmercode,noofcrate,mindistyear,3,noofcrate,1 from mhv.tblpolinizerjune a LEFT JOIN mhv.tblpollenizeranalysis b ON a.farmercode=b.farmercode"
    '## end
    ' add flag condition for code below
    'and SUBSTRING(IDFARMER,1,3)IN  " & Dzstr * extracting with flag =0 is enough
    
        SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,0 AS REGLAND,village,phone1,sum(plantqty) polinizercrate FROM tblfarmer A,tblpolinizer B WHERE A.IDFARMER=B.farmercode and B.flag=0 "
        SQLSTR = SQLSTR & "  " & "group by idfarmer "
MHVDB.Execute "delete from  tbldistpreparetion"
     MHVDB.Execute "insert into tbldistpreparetion(dzcode,gecode,tscode,idfarmer,farmername,regland,village,phone1,polinizercrate) " & SQLSTR
       SQLSTR = ""
       'SQLSTR = "SELECT SUBSTRING(farmercode,1,3) AS DZCODE,SUBSTRING(farmercode,4,3) AS GECODE,SUBSTRING(farmercode,7,3) AS TSCODE,farmercode,FARMERNAME,0 AS REGLAND,village,phone1 FROM tblfarmer A,refillin B WHERE A.status='A' and A.IDFARMER=B.farmercode and substring(farmercode,10,1)='F' and SUBSTRING(farmercode,1,9)IN  " & Dzstr & " and farmercode not in(select idfarmer from tbldistpreparetion)"
        '    SQLSTR = SQLSTR & "  " & "group by farmercode order by  FIELD(SUBSTRING(farmercode,1,9), " & morderstr & ") "
       'MHVDB.Execute "insert into tbldistpreparetion " & SQLSTR
      '            SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='F' and SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
'            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "

  SQLSTR = "SELECT * from tbldistpreparetion order by  FIELD(SUBSTRING(IDFARMER,1,9), " & morderstr & ") "
    
    Else
      SQLSTR = "SELECT SUBSTRING(IDFARMER,1,3) AS DZCODE,SUBSTRING(IDFARMER,4,3) AS GECODE,SUBSTRING(IDFARMER,7,3) AS TSCODE,IDFARMER,FARMERNAME,REGLAND AS REGLAND,village,phone1 FROM tblfarmer A,tbllandreg B WHERE A.status='A' and A.IDFARMER=B.FARMERID AND substring(idfarmer,10,1)='C' and  SUBSTRING(IDFARMER,1,9)IN  " & Dzstr
            SQLSTR = SQLSTR & "  " & "group by idfarmer ORDER BY SUBSTRING(IDFARMER,1,9) ,IDFARMER "
    
    
    End If

End If
Dim mrnd As Integer
Dim tmod As Integer
mrnd = 0

MHVDB.Execute "delete from mhv.tbldistpreparetion where substring(idfarmer,1,14) in( select substring(FARMER,1,14) from " _
& " (select n.FARMER,n.START,n.PLANTFUTURE,round(datediff(CURDATE(),n.START),0) recordage,abs(round(datediff(n.START,END),0)) daydiff" _
& " from odk_prodlocal.tblfuturefarmer_core n INNER JOIN (SELECT FARMER,MAX(START) lastEntry FROM odk_prodlocal.tblfuturefarmer_core" _
& " where PLANTFUTURE IN('no') GROUP BY FARMER)x ON n.FARMER = x.FARMER AND n.START = x.lastEntry AND STATUS <> 'BAD' GROUP BY n.FARMER)" _
& "mm )"

                            rs.Open SQLSTR, MHVDB
                            
                            i = 1
                            cnt = maxDistNo
                            polycont = 0
                            Do Until rs.EOF
                            'If polycont > 77500 Then Exit Do (this was for 2013 extra plant distribution, as decided by mgmt)
                           
                                 mydgt = Mid(rs!idfarmer, 1, 9)
                                 Do While mydgt = Mid(rs!idfarmer, 1, 9)
                                 
                                         If i >= 5 Then
                                         mygrid.rows = mygrid.rows + 1
                                         End If
                                         mygrid.TextMatrix(i, 0) = i
                                         mygrid.TextMatrix(i, 1) = cnt
                                         FindDZ rs!dzcode
                                         FindGE rs!dzcode, rs!GECODE
                                         FindTs rs!dzcode, rs!GECODE, rs!tscode
                                         mygrid.TextMatrix(i, 2) = rs!dzcode & " " & Dzname
                                         mygrid.TextMatrix(i, 3) = rs!GECODE & " " & GEname
                                         mygrid.TextMatrix(i, 4) = rs!tscode & " " & TsName
                                         mygrid.TextMatrix(i, 5) = rs!idfarmer
                                         mygrid.TextMatrix(i, 6) = rs!farmername
                                         mygrid.TextMatrix(i, 7) = IIf(IsNull(rs!phone1), "", rs!phone1)
                                         mygrid.TextMatrix(i, 8) = rs!VILLAGE
                                         mygrid.TextMatrix(i, 9) = Format(IIf(IsNull(rs!regland), 0#, rs!regland), "####0.00")
                                         
                                                                                         If chkPolinizer.Value = 1 Then
                                                                                                       mygrid.TextMatrix(i, 10) = Round(rs!polinizercrate * 35, 0)
                                                       mygrid.TextMatrix(i, 11) = Round(rs!polinizercrate, 2) '(Val(Mygrid.TextMatrix(i, 10)) - (Val(Mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
                                                      ' mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, mrnd)
                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       
                                                       'End If
                                                       End If
                                                    If chkrefill.Value = 1 Then
                                                       mygrid.TextMatrix(i, 10) = rs!totfillindemand
                                                       mygrid.TextMatrix(i, 12) = rs!totfillindemand
                                                    End If
                                         Set rs1 = Nothing
                                        rs1.Open "select group_concat(a.id) refilltrnno,sum(b) as b,sum(e) as e,sum(p1) as p1, sum(n) as n from refillin a, refillinheader b where a.headerid=b.id and farmercode='" & rs!idfarmer & "' and  isfinalized='Yes' and status='ON'   group by farmercode", MHVDB
                                         'rs1.Open "select sum(regland)*(420*.1*.5) as b,sum(regland)*(420*.1*.5) as e,sum(regland)*(420*.06) as p1, sum(regland)*(420*.1) as n from tbllandreg where farmerid='" & rs!idfarmer & "' group by farmerid ", MHVDB
                                         
                                         
                                         If rs1.EOF <> True And chkrefill.Value = 0 Then
                                                btype = Round(rs1!b, mrnd)
                                                etype = Round(rs1!e, mrnd)
                                                Set RS2 = Nothing
                                                RS2.Open "select * from tbldistformula where status='ON'", MHVDB
                                                If RS2.EOF <> True Then
                                                mygrid.TextMatrix(i, 10) = Round(rs1!p1 + rs1!n + btype + etype, mrnd)
                                                       'If etype > 21 Then
'                                                       mygrid.TextMatrix(i, 13) = etype
'                                                       mygrid.TextMatrix(i, 10) = Round(btype + Val(mygrid.TextMatrix(i, 13)) + rs1!p1 + rs1!n, 2)
'                                                       mygrid.TextMatrix(i, 11) = Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, 2)
'                                                       mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, 2)
'                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, 2)
'                                                        If etype > 35 Then
'                                                                tmod = etype Mod 35
'                                                                If tmod > 17 Then
'                                                                etype = etype + 35 - tmod
'                                                                btype = btype - tmod
'                                                                Else
'                                                                etype = etype - tmod
'                                                                btype = btype - tmod
'
'                                                                End If
'                                                        Else
'                                                        tmod = 0
'                                                        etype = 35
'                                                        btype = 35
'
'                                                        End If

                                                       mygrid.TextMatrix(i, 13) = etype
                                                       mygrid.TextMatrix(i, 10) = Round(btype + Val(mygrid.TextMatrix(i, 13)) + rs1!p1 + rs1!n, mrnd)
                                                       mygrid.TextMatrix(i, 11) = (Val(mygrid.TextMatrix(i, 10)) - (Val(mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
                                                      ' mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, mrnd)
                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       
                                                       'End If
            
                                                        mygrid.TextMatrix(i, 13) = etype
                                                       mygrid.TextMatrix(i, 10) = rs!polinizercrate
                                                       mygrid.TextMatrix(i, 11) = (Val(mygrid.TextMatrix(i, 10)) - (Val(mygrid.TextMatrix(i, 10)) Mod 35)) / RS2!crateno '- rs1!p1 - rs1!n 'Round(btype + Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       mygrid.TextMatrix(i, 12) = Round(btype, mrnd)
                                                       mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)), mrnd)
                                                      ' mygrid.TextMatrix(i, 12) = Round(btype / RS2!crateno, mrnd)
                                                      ' mygrid.TextMatrix(i, 13) = Round(Val(mygrid.TextMatrix(i, 13)) / RS2!crateno, mrnd)
                                                       
                                                       'End If
            
            
         
                                                       
                                                       
                                                       
                                                End If
                                                

                                                
                                                mygrid.TextMatrix(i, 15) = Round(rs1!p1, mrnd) '
                                                mygrid.TextMatrix(i, 16) = Round(rs1!n, 0) 'Round(rs1!n / RS2!crateno, 0)
                                                polycont = polycont + Round(rs1!p1, mrnd) + Round(rs1!n, mrnd)
                                                mygrid.TextMatrix(i, 29) = "O"
                                                 mygrid.TextMatrix(i, 30) = mygrid.TextMatrix(i, 16)
                                                 mygrid.TextMatrix(i, 31) = RS2!fid
                                                  mygrid.TextMatrix(i, 32) = rs1!refilltrnno
                                                 
                                         End If
                                         i = i + 1
                                         rs.MoveNext
                                         
                                         If rs.EOF Then Exit Do
                                 Loop
                                 cnt = cnt + 1
                                 mygrid.rows = mygrid.rows + 1
                                 mygrid.TextMatrix(i, 28) = "S"
                                 mygrid.TextMatrix(i, 0) = i
                                 i = i + 1
                            Loop
                            mygrid.rows = mygrid.rows + 1
                            mygrid.TextMatrix(i, 28) = "T"
                            mygrid.TextMatrix(i, 0) = i
              
End Sub

