VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmdistributionsummary 
   Caption         =   "Distribution SUmmary"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13995
   Icon            =   "frmdistributionsummary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   6000
      Picture         =   "frmdistributionsummary.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   3960
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdistributionsummary.frx":15AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdistributionsummary.frx":1946
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdistributionsummary.frx":1CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdistributionsummary.frx":29BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdistributionsummary.frx":2E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdistributionsummary.frx":35C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13995
      _ExtentX        =   24686
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ADD"
            Key             =   "ADD"
            Object.ToolTipText     =   "ADDS NEW RECORD"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "New"
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Open"
                  Text            =   "Open"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OPEN"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SAVE"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "DELETE"
            Key             =   "DELETE"
            Object.ToolTipText     =   "DELETE THE RECORD"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            Key             =   "EXIT"
            Object.ToolTipText     =   "EXIT FROM THE FORM"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
   Begin MSDataListLib.DataCombo cbotrnid 
      Bindings        =   "frmdistributionsummary.frx":3960
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Top             =   960
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
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   13815
      _cx             =   24368
      _cy             =   11245
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmdistributionsummary.frx":3975
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.Label label 
      AutoSize        =   -1  'True
      Caption         =   "Distribution Schedule Schedule No."
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
      TabIndex        =   2
      Top             =   1080
      Width           =   3045
   End
End
Attribute VB_Name = "frmdistributionsummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loadcombo(Operation As String)
Dim RSTR As New ADODB.Recordset

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
If Operation = "ADD" Then
TB.Buttons(3).Enabled = True
RSTR.Open "select concat(cast(trnid as char) ,' ',distributionname,' ',cast(year as char),' ',cast(mnth as char)) as dname,trnid  from tblplantdistributionheader where status='ON' and planneddist='Y' and trnid not in(select scheduleid from tbldistributionsummary) order by trnid", db
Else
TB.Buttons(3).Enabled = False
RSTR.Open "select distinct scheduleid as trnid,concat(cast(b.trnid as char) ,' ',distributionname,' ',cast(b.year as char),' ',cast(mnth as char)) as dname from tbldistributionsummary as a,tblplantdistributionheader as b where b.trnid=scheduleid and a.year=b.year order by scheduleid", db
End If

Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "dname"
cbotrnid.BoundColumn = "trnid"
End Sub

Private Sub FillGrid()
Dim SQLSTR As String
Dim rs As New ADODB.Recordset
Set rs = Nothing
SQLSTR = ""

        
        
        
mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^D.No|^Year|^Dzongkhag|^Gewog|^Tshewog|^Loading Date|^Travelling Date|^DCM Detail|^Delivery Status|^"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 750
mygrid.ColWidth(2) = 1875
mygrid.ColWidth(3) = 1620
mygrid.ColWidth(4) = 1950
mygrid.ColWidth(5) = 1305
mygrid.ColWidth(6) = 1455
mygrid.ColWidth(7) = 2145
mygrid.ColWidth(8) = 1995
mygrid.ColWidth(9) = 525


If Operation = "ADD" Then
SQLSTR = "SELECT distno,year," _
        & " substring(farmercode,1,3) dzongkhag," _
        & " substring(farmercode,4,3) gewog," _
        & " substring(farmercode,7,3) tshowog from tblplantdistributiondetail where " _
        & " trnid='" & cbotrnid.BoundText & "' and length(farmercode)>0 and status<>'C' group by distno," _
        & " substring(farmercode,1,3)," _
        & " substring(farmercode,4,3) ," _
        & " substring(farmercode,7,3)  order by distno , substring(farmercode,1,3)," _
        & " substring(farmercode,4,3) ," _
        & " substring(farmercode,7,3) "
        
 Else
 
 SQLSTR = "SELECT distno,year," _
        & "  dzongkhag," _
        & "  gewog," _
        & "  tshowog from tbldistributionsummary where " _
        & " scheduleid='" & cbotrnid.BoundText & "' group by distno," _
        & " dzongkhag," _
        & " gewog ," _
        & " tshowog  order by distno , dzongkhag," _
        & " gewog ," _
        & " tshowog "
 
 End If
        
        
        
rs.Open SQLSTR, MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1

FindDZ rs!dzongkhag
FindGE rs!dzongkhag, rs!gewog
FindTs rs!dzongkhag, rs!gewog, rs!tshowog

mygrid.TextMatrix(i, 0) = rs!distno
mygrid.TextMatrix(i, 1) = rs!Year
mygrid.TextMatrix(i, 2) = rs!dzongkhag & "  " & Dzname
mygrid.TextMatrix(i, 3) = rs!gewog & "  " & GEname
mygrid.TextMatrix(i, 4) = rs!tshowog & "  " & TsName
rs.MoveNext
i = i + 1
Loop


mygrid.ColAlignment(2) = flexAlignLeftCenter
mygrid.ColAlignment(3) = flexAlignLeftCenter
mygrid.ColAlignment(4) = flexAlignLeftCenter
mygrid.MergeCol(0) = True
mygrid.MergeCells = 1


mygrid.MergeCol(1) = True
mygrid.MergeCells = 1
mygrid.MergeCol(2) = True
mygrid.MergeCells = 1
mygrid.MergeCol(3) = True
mygrid.MergeCells = 1

'Mygrid.MergeCol(3) = True
'Mygrid.MergeCells = 3

End Sub

Private Sub Command1_Click()
If Len(cbotrnid.Text) = 0 Then Exit Sub
cbotrnid.Enabled = False
FillGrid
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       CLEARCONTROLL
        Operation = "ADD"
     cbotrnid.Enabled = True
     loadcombo "ADD"
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbotrnid.Enabled = True
       loadcombo "OPEN"
        
       Case "SAVE"
       MNU_SAVE
        TB.Buttons(3).Enabled = False
         
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
cbotrnid.Text = ""
mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^D.No|^Year|^Dzongkhag|^Gewog|^Tshewog|^Loading Date|^Travelling Date|^DCM Detail|^Delivery Status|^"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 750
mygrid.ColWidth(2) = 1875
mygrid.ColWidth(3) = 1620
mygrid.ColWidth(4) = 1950
mygrid.ColWidth(5) = 1305
mygrid.ColWidth(6) = 1455
mygrid.ColWidth(7) = 2145
mygrid.ColWidth(8) = 1995
mygrid.ColWidth(9) = 525
End Sub
Private Sub MNU_SAVE()
On Error GoTo err
Dim rs As New ADODB.Record
Dim i As Integer
MHVDB.BeginTrans
For i = 1 To mygrid.Rows - 1
If Len(Trim(mygrid.TextMatrix(i, 0))) = 0 Then Exit For

MHVDB.Execute "insert into tbldistributionsummary(uid,scheduleid,distno,year,dzongkhag,gewog, " _
& " tshowog,updatedon,deliverystatus,ddzongkhag,dgewog,dtshowog) " _
& "values('" & ((Trim(mygrid.TextMatrix(i, 0))) & (Trim(mygrid.TextMatrix(i, 1)))) & "','" & cbotrnid.BoundText & "','" & Trim(mygrid.TextMatrix(i, 0)) & "', " _
& " '" & Trim(mygrid.TextMatrix(i, 1)) & "','" & Mid(Trim(mygrid.TextMatrix(i, 2)), 1, 3) & "', " _
& "'" & Mid(Trim(mygrid.TextMatrix(i, 3)), 1, 3) & "','" & Mid(Trim(mygrid.TextMatrix(i, 4)), 1, 3) & "'," _
& "'" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','Pending', " _
& " '" & Trim(mygrid.TextMatrix(i, 2)) & "','" & Trim(mygrid.TextMatrix(i, 3)) & "','" & Trim(mygrid.TextMatrix(i, 4)) & "')"

Next
MHVDB.CommitTrans
Exit Sub
err:
    MHVDB.RollbackTrans
    MsgBox err.Description


End Sub
