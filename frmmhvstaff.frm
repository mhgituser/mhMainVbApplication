VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmmhvstaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "S T A F F    M A S T E R"
   ClientHeight    =   8115
   ClientLeft      =   7185
   ClientTop       =   1680
   ClientWidth     =   11355
   Icon            =   "frmmhvstaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11355
   Begin VB.Frame Frame3 
      Caption         =   "STAFF APPLICALE FOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   4080
      Width           =   5775
      Begin VB.CheckBox Check4 
         Caption         =   "VEHICLE BOOKING"
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
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "PASSENGER BOOKING"
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
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   5760
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtteritory 
         Appearance      =   0  'Flat
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
         Left            =   1440
         TabIndex        =   28
         Top             =   840
         Width           =   3735
      End
      Begin MSDataListLib.DataCombo cbomspervisor 
         Bindings        =   "frmmhvstaff.frx":0E42
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "TERRRITORY"
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
         TabIndex        =   27
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SUPERVISOR"
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
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "STAFF TYPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   3120
      Width           =   5775
      Begin VB.CheckBox chknursery 
         Caption         =   "NURSERY"
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
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkoutreach 
         Caption         =   "OUTREACH"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkadvocate 
         Caption         =   "ADVOCATE"
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
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkmoniter 
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
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame itemcode 
      Caption         =   "STAFF INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11295
      Begin VB.TextBox txtcontact 
         Appearance      =   0  'Flat
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
         Left            =   7080
         TabIndex        =   23
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
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
         Left            =   7080
         TabIndex        =   21
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtstaffname 
         Appearance      =   0  'Flat
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
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
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
         Left            =   2520
         TabIndex        =   2
         Top             =   1800
         Width           =   8655
      End
      Begin MSDataListLib.DataCombo cbostaffcode 
         Bindings        =   "frmmhvstaff.frx":0E57
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cbodept 
         Height          =   315
         Left            =   2520
         TabIndex        =   19
         ToolTipText     =   "Maximum 3 Charactar"
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "REMARKS"
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
         TabIndex        =   22
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL"
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
         Left            =   5880
         TabIndex        =   20
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DEPT"
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
         TabIndex        =   18
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "STAFF ID"
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
         TabIndex        =   7
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "STAFF NAME"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CONTACT #"
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
         Left            =   5880
         TabIndex        =   5
         Top             =   1440
         Width           =   1065
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   11175
      _cx             =   19711
      _cy             =   4895
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
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12632256
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmmhvstaff.frx":0E6C
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
   Begin MSComctlLib.ImageList IMG 
      Left            =   11400
      Top             =   720
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
            Picture         =   "frmmhvstaff.frx":0F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmhvstaff.frx":12E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmhvstaff.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmhvstaff.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmhvstaff.frx":27AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmhvstaff.frx":2F64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
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
   Begin VB.Label Label4 
      Caption         =   "DATA SCREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
End
Attribute VB_Name = "frmmhvstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Srs As New ADODB.Recordset

Private Sub cbostaffcode_LostFocus()
On Error GoTo err

   cbostaffcode.BackColor = vbWhite
   cbostaffcode.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblmhvstaff where staffcode='" & cbostaffcode.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
   txtstaffname.Text = IIf(IsNull(rs!staffname), "", rs!staffname)
   
   txtremarks.Text = IIf(IsNull(rs!remarks), "", rs!remarks)
   chkmoniter.Value = IIf(IsNull(rs!moniter), 0, rs!moniter)
    chkadvocate.Value = IIf(IsNull(rs!ADVOCATE), 0, rs!ADVOCATE)
    chkoutreach.Value = IIf(IsNull(rs!outreach), 0, rs!outreach)
    txtcontact.Text = IIf(IsNull(rs!contact), "", rs!contact)
    txtemail.Text = IIf(IsNull(rs!email), "", rs!email)
    txtteritory.Text = rs!mteritory
    
    If rs!Dept > 0 Then
    FindDepartment rs!Dept
    cbodept.Text = rs!Dept & " " & DeptName
    Else
    cbodept.Text = ""
    End If
   If chkmoniter.Value = 1 Then
   FindsTAFF rs!msupervisor
   Frame1.Visible = True
   cbomspervisor.Text = rs!msupervisor & " " & sTAFF
   Else
   Frame1.Visible = False

   End If
   
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub

Private Sub chkmoniter_Click()
If chkmoniter.Value = 1 Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff order by STAFFCODE", db
Set cbostaffcode.RowSource = Srs
cbostaffcode.ListField = "STAFFNAME"
cbostaffcode.BoundColumn = "STAFFCODE"

Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff  order by STAFFCODE", db
Set cbomspervisor.RowSource = Srs
cbomspervisor.ListField = "STAFFNAME"
cbomspervisor.BoundColumn = "STAFFCODE"
Set Srs = Nothing
If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(cast(deptid as char) , ' ', deptname) as deptname,deptid  from tbldepartment order by deptid", db
Set cbodept.RowSource = Srs
cbodept.ListField = "deptname"
cbodept.BoundColumn = "deptid"


FillGrid
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       cbostaffcode.Enabled = False
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       Dim id As String
       Dim rs As New ADODB.Recordset
       Set rs = Nothing
       rs.Open "SELECT MAX(SUBSTRING(staffcode,2,4))+1 AS MaxID from tblmhvstaff", MHVDB, adOpenForwardOnly, adLockOptimistic
       If rs.EOF <> True Then
       
id = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
If Len(id) = 1 Then
id = "000" & id
ElseIf Len(id) = 2 Then
id = "00" & id
ElseIf Len(id) = 3 Then
id = "0" & id
Else

End If
        cbostaffcode.Text = "S" & id
        
       Else
       cbostaffcode.Text = "S0001"
       End If
       
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbostaffcode.Enabled = True
      TB.Buttons(3).Enabled = True
       
       Case "SAVE"
       MNU_SAVE
        TB.Buttons(3).Enabled = False
        FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub MNU_SAVE()
On Error GoTo err

If chkmoniter.Value = 1 And Len(cbomspervisor.Text) = 0 Then
MsgBox "Select The Corresponding Supervisor for this monitor"
Exit Sub
End If

If Len(Trim(cbodept.Text)) = 0 Then
MsgBox "Select The Department!"
Exit Sub
End If



MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute "INSERT INTO tblmhvstaff (STAFFCODE,STAFFNAME,REMARKS,MONITER, " _
& "ADVOCATE,OUTREACH,msupervisor,nursery,dept,email,contact,mteritory) " _
& "VALUEs('" & cbostaffcode.Text & "','" & txtstaffname.Text & "', " _
& "'" & txtremarks.Text & "','" & chkmoniter.Value & "','" & chkadvocate.Value & "', " _
& "'" & chkoutreach.Value & "','" & cbomspervisor.BoundText & "', " _
& "'" & chknursery.Value & "','" & cbodept.BoundText & "','" & txtemail & "','" & txtcontact.Text & "','" & txtteritory.Text & "')"

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblmhvstaff set STAFFNAME='" & txtstaffname.Text & "', " _
& "remarks='" & txtremarks.Text & "',nursery='" & chknursery.Value & "', " _
& "MONITER='" & chkmoniter.Value & "',ADVOCATE='" & chkadvocate.Value & "',mteritory='" & txtteritory.Text & "', " _
& "OUTREACH='" & chkoutreach.Value & "',msupervisor='" & cbomspervisor.BoundText & "', " _
& "dept='" & cbodept.BoundText & "', email='" & txtemail.Text & "',contact='" & txtcontact.Text & "' where STAFFCODE='" & cbostaffcode.BoundText & "'"
Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='6'"
MHVDB.CommitTrans
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans




End Sub
Private Sub CLEARCONTROLL()

txtstaffname.Text = ""
txtremarks.Text = ""
cbomspervisor.Text = ""
txtemail.Text = ""
txtcontact.Text = ""
cbodept.Text = ""
chkmoniter.Value = 0
chkadvocate.Value = 0
chkoutreach.Value = 0
txtteritory.Text = ""
End Sub
Private Sub FillGrid()
On Error GoTo err

Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^SL.NO.|^STAFF ID|^STAFF NAME|^MONITER|^ADVOCATE|^OUTREACH|^REMARKS"
Mygrid.ColWidth(0) = 750
Mygrid.ColWidth(1) = 1155
Mygrid.ColWidth(2) = 2985
Mygrid.ColWidth(3) = 960
Mygrid.ColWidth(4) = 1155
Mygrid.ColWidth(5) = 1125
Mygrid.ColWidth(6) = 2895

rs.Open "select * from tblmhvstaff order by STAFFCODE", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i

Mygrid.TextMatrix(i, 1) = IIf(IsNull(rs!staffcode), "", rs!staffcode)
Mygrid.TextMatrix(i, 2) = IIf(IsNull(rs!staffname), "", rs!staffname)

If IIf(IsNull(rs!moniter), 0, rs!moniter) = 1 Then
Mygrid.TextMatrix(i, 3) = "YES"
Else
Mygrid.TextMatrix(i, 3) = "NO"
End If

If IIf(IsNull(rs!ADVOCATE), 0, rs!ADVOCATE) = 1 Then
Mygrid.TextMatrix(i, 4) = "YES"
Else
Mygrid.TextMatrix(i, 4) = "NO"
End If

If IIf(IsNull(rs!outreach), 0, rs!outreach) = 1 Then
Mygrid.TextMatrix(i, 5) = "YES"
Else
Mygrid.TextMatrix(i, 5) = "NO"
End If

Mygrid.TextMatrix(i, 6) = IIf(IsNull(rs!remarks), "", rs!remarks)
rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub txtdzname_Change()

End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
'If InStr(1, "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,6}$", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
'        KeyAscii = 0
'    End If
End Sub

Private Sub txtemail_Validate(Cancel As Boolean)
If IsValidEmail(txtemail.Text) = False Then
MsgBox "Invalid Email."
Cancel = True
End If
End Sub

Private Sub txtstaffname_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtstaffname.SelStart + 1
    Dim sText As String
    sText = Left$(txtstaffname.Text, iPos)
    If iPos = 1 Then GoTo Upit
    KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
    If iPos > 1 And _
     (InStr(iPos - 1, sText, " ") > 0 Or _
      InStr(iPos - 1, sText, "-") > 0 Or _
      InStr(iPos - 1, sText, ".") > 0 Or _
      InStr(iPos - 1, sText, "'") > 0) _
      Then GoTo Upit
    If iPos > 2 Then _
      If InStr(iPos - 2, sText, "Mc") > 0 _
        Then GoTo Upit
        
   End If
  Exit Sub
Upit:
  KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub
