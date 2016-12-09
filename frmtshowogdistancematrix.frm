VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmtshowogdistancematrix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tshowog Distance Matrix Entry"
   ClientHeight    =   7470
   ClientLeft      =   4860
   ClientTop       =   930
   ClientWidth     =   9855
   Icon            =   "frmtshowogdistancematrix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9855
   Begin VB.CommandButton Command1 
      Caption         =   "Update LatLng"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   720
      Width           =   2535
   End
   Begin VSFlex7Ctl.VSFlexGrid mygrid 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   9615
      _cx             =   16960
      _cy             =   8705
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Rows            =   200
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmtshowogdistancematrix.frx":0E42
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
   Begin VB.Frame framedestination 
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
      Begin MSDataListLib.DataCombo cbodzongkhagdest 
         Bindings        =   "frmtshowogdistancematrix.frx":0F23
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
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
      Begin MSDataListLib.DataCombo cbogewogdest 
         Bindings        =   "frmtshowogdistancematrix.frx":0F38
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gewog"
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
         TabIndex        =   11
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dzongkhag"
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
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framesource 
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
      Begin MSDataListLib.DataCombo cbodzongkhag 
         Bindings        =   "frmtshowogdistancematrix.frx":0F4D
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "frmtshowogdistancematrix.frx":0F62
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   ""
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gewog"
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
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dzongkhag"
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
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   5280
      Top             =   600
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
            Picture         =   "frmtshowogdistancematrix.frx":0F77
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtshowogdistancematrix.frx":1311
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtshowogdistancematrix.frx":16AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtshowogdistancematrix.frx":2385
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtshowogdistancematrix.frx":27D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtshowogdistancematrix.frx":2F91
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
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
      Bindings        =   "frmtshowogdistancematrix.frx":332B
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   ""
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
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   0
      Picture         =   "frmtshowogdistancematrix.frx":3340
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   0
      Picture         =   "frmtshowogdistancematrix.frx":36CA
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBtnUp 
      Height          =   240
      Left            =   0
      Picture         =   "frmtshowogdistancematrix.frx":3A54
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgBtnDn 
      Height          =   240
      Left            =   0
      Picture         =   "frmtshowogdistancematrix.frx":3DDE
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmtshowogdistancematrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstring As String

Private Sub cbodzongkhag_LostFocus()
If validatecombo("tbldzongkhag", "dzongkhagcode", cbodzongkhag.BoundText) = False Then
cbodzongkhag.SetFocus
MsgBox "Invalid Dzongkhag Code!"
Exit Sub
End If
cbodzongkhag.Enabled = False
cbogewog.Text = ""
'CBOTSHOWOG.Text = ""
fillgewogcombo
End Sub

Private Sub cbodzongkhagdest_LostFocus()
If validatecombo("tbldzongkhag", "dzongkhagcode", cbodzongkhagdest.BoundText) = False Then
cbodzongkhagdest.SetFocus
MsgBox "Invalid Dzongkhag Code!"
Exit Sub
End If
cbodzongkhagdest.Enabled = False
cbogewogdest.Text = ""
fillgewogcombodest
End Sub

Private Sub cbogewog_LostFocus()
If Len(cbodzongkhag.Text) = 0 Then Exit Sub
If validatecombo("tblgewog", "concat(DzongkhagId,GewogId)", cbodzongkhag.BoundText & cbogewog.BoundText) = False Then
cbogewog.SetFocus
MsgBox "Invalid Gewog Code!"
Exit Sub
End If
cbogewog.Enabled = False
'CBOTSHOWOG.Text = ""
filltshowogcombo
End Sub

Private Sub cbogewogdest_LostFocus()
If Len(cbodzongkhagdest.Text) = 0 Then Exit Sub
If validatecombo("tblgewog", "concat(DzongkhagId,GewogId)", cbodzongkhagdest.BoundText & cbogewogdest.BoundText) = False Then
cbogewogdest.SetFocus
MsgBox "Invalid Gewog Code!"
Exit Sub
End If
cbogewogdest.Enabled = False
End Sub

Private Sub CBOTSHOWOG_LostFocus()
If Len(cbodzongkhag.Text) = 0 Then Exit Sub
If validatecombo("tbltshewog", "concat(DzongkhagId,GewogId,TshewogId)", cbodzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText) = False Then
CBOTSHOWOG.SetFocus
MsgBox "Invalid Tshowog Code!"
Exit Sub
End If
CBOTSHOWOG.Enabled = False
End Sub

Private Sub cbotrnid_LostFocus()
cbodzongkhag.Enabled = False
cbogewog.Enabled = False
cbodzongkhagdest.Enabled = False
cbogewogdest.Enabled = False
FillGrid

End Sub

Private Sub Command1_Click()
updatelatlng
End Sub

Private Sub Form_Load()
On Error GoTo err
Dim rsDz As New ADODB.Recordset
Dim RSTR As New ADODB.Recordset
Operation = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing
If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cbodzongkhag.RowSource = rsDz
cbodzongkhag.ListField = "dzongkhagname"
cbodzongkhag.BoundColumn = "dzongkhagcode"

Set rsDz = Nothing
If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cbodzongkhagdest.RowSource = rsDz
cbodzongkhagdest.ListField = "dzongkhagname"
cbodzongkhagdest.BoundColumn = "dzongkhagcode"




Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select distinct headerid   from odk_prodlocal.tbltshowogdistancematrix order by trnid", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "headerid"
cbotrnid.BoundColumn = "headerid"



Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub fillgewogcombo()
Dim rsGe As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cbodzongkhag.BoundText & "' order by gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"
End Sub

Private Sub fillgewogcombodest()
Dim rsGe As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cbodzongkhagdest.BoundText & "' order by gewogid", db
Set cbogewogdest.RowSource = rsGe
cbogewogdest.ListField = "gewogname"
cbogewogdest.BoundColumn = "gewogid"
End Sub



Private Sub filltshowogcombo()
'Dim rsTs As New ADODB.Recordset
'Set db = New ADODB.Connection
'db.CursorLocation = adUseClient
'db.Open CnnString
'If rsTs.State = adStateOpen Then rsTs.Close
'rsTs.Open "select concat(tshewogid , ' ', tshewogname) as tshewogname,tshewogid  from tbltshewog where dzongkhagid='" & cbodzongkhag.BoundText & "' and gewogid='" & cbogewog.BoundText & "' order by dzongkhagid,gewogid", db
'Set CBOTSHOWOG.RowSource = rsTs
'CBOTSHOWOG.ListField = "tshewogname"
'CBOTSHOWOG.BoundColumn = "tshewogid"
End Sub

Private Sub FillGridCombo(col As Integer)
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        mn1 = "          |"
        StrComboList = ""
        formstr
            Set RstTemp = Nothing
            
            If col = 2 Then
            RstTemp.Open ("select  concat(DzongkhagId,GewogId,tshewogid)tshewogid ,tshewogname  from tbltshewog where " _
            & " dzongkhagid='" & cbodzongkhagdest.BoundText & "' and " _
            & " gewogid='" & cbogewogdest.BoundText & "'"), MHVDB
ElseIf col = 1 Then
     RstTemp.Open ("select concat(DzongkhagId,GewogId,tshewogid)tshewogid ,tshewogname  from tbltshewog where " _
            & " dzongkhagid='" & cbodzongkhag.BoundText & "' and " _
            & " gewogid='" & cbogewog.BoundText & "'"), MHVDB

Else
'nothing
End If
            
            
            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("tshewogid").Value & " " & RstTemp("tshewogname").Value, StrComboList & "|" & RstTemp("tshewogid").Value) & " " & RstTemp("tshewogname").Value

                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList


End Sub


Private Sub FillGridComboRoadType()
        Dim RstTemp As New ADODB.Recordset
        Dim NmcItemData As String
        Dim StrComboList As String
        Dim mn1 As String
        mn1 = "          |"
        StrComboList = ""
        
            Set RstTemp = Nothing
            RstTemp.Open ("select typeid,typedesc  from tblroadtype order by typeid"), MHVDB

            If Not RstTemp.BOF Then
               Do While Not RstTemp.EOF
                    NmcItemData = ""
                    StrComboList = IIf(StrComboList = "", RstTemp("typeid").Value & " " & RstTemp("typedesc").Value, StrComboList & "|" & RstTemp("typeid").Value) & " " & RstTemp("typedesc").Value

                    RstTemp.MoveNext
               Loop
            End If
            RstTemp.Close
       mygrid.ComboList = StrComboList


End Sub

Private Sub mygrid_Click()
If Len(cbogewogdest.Text) = 0 Then Exit Sub
If Len(cbogewog.Text) = 0 Then Exit Sub

putslno

mygrid.Editable = flexEDNone
If (mygrid.col = 1 Or mygrid.col = 2) And Len(mygrid.TextMatrix(mygrid.row - 1, 4)) > 0 Then
mygrid.Editable = flexEDKbdMouse

FillGridCombo mygrid.col

Exit Sub
Else
mygrid.ComboList = ""
mygrid.Editable = flexEDNone
End If

If mygrid.col = 4 And Val(mygrid.TextMatrix(mygrid.row, 3)) > 0 Then
mygrid.Editable = flexEDKbdMouse
FillGridComboRoadType
Exit Sub
Else

mygrid.Editable = flexEDNone
End If


If (mygrid.col = 3 Or mygrid.col = 5) And Len(mygrid.TextMatrix(mygrid.row, 1)) > 0 Then

mygrid.Editable = flexEDKbdMouse
Exit Sub
Else

mygrid.Editable = flexEDNone
End If
End Sub

Private Sub formstr()
Dim i As Integer
mstring = ""
For i = 1 To mygrid.Rows - 1
    If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit For
    mstring = mstring + "'" + Trim(Mid(mygrid.TextMatrix(i, 1), 1, 3)) + "',"
         

    
Next

If Len(mstring) > 0 Then

   mstring = "(" + Left(mstring, Len(mstring) - 1) + ")"

Else
     mstring = "('" + "XXX" + "'" + ")"
End If
End Sub

Private Sub addbtn()
    mygrid.Editable = flexEDKbdMouse
    mygrid.AllowUserResizing = flexResizeNone

   
    Dim i%
 If mygrid.col = 1 And Len(mygrid.TextMatrix(mygrid.row, 1)) > 0 Then
        mygrid.Cell(flexcpPicture, mygrid.row, 5) = imgBtnUp
        mygrid.Cell(flexcpPictureAlignment, mygrid.row, 5) = flexAlignRightCenter
   End If
End Sub

Private Sub mygrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 And Trim(mygrid.TextMatrix(mygrid.row, 6)) = "" And Len(Trim(mygrid.TextMatrix(mygrid.row, 1))) > 0 Then
If MsgBox("Do you want to delete this row?", vbYesNo) = vbYes Then
mygrid.RemoveItem (mygrid.row)
End If
End If

End Sub


Private Sub clearcontrol()



mygrid.Clear
mygrid.FormatString = "^Sl.No.|^From|^To|^Distance|Road Type|Note|frmdb|"
mygrid.ColWidth(0) = 585
mygrid.ColWidth(1) = 2220
mygrid.ColWidth(2) = 2220
mygrid.ColWidth(3) = 720
mygrid.ColWidth(4) = 1965
mygrid.ColWidth(5) = 1290
mygrid.ColWidth(6) = 1
mygrid.ColWidth(7) = 255

cbodzongkhag.Text = ""
cbogewog.Text = ""

cbodzongkhagdest.Text = ""
cbogewogdest.Text = ""



End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "ADD"
    Dim rs As New ADODB.Recordset
    Set rs = Nothing
    rs.Open "select (max(headerid)+1) as maxid from tbltshowogdistancematrix", ODKDB
    cbotrnid.Text = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
    clearcontrol
    cbodzongkhag.Enabled = True
    cbodzongkhagdest.Enabled = True
    cbogewog.Enabled = True
    cbogewogdest.Enabled = True
    cbotrnid.Enabled = False
    TB.Buttons(3).Enabled = True
    
    Case "OPEN"
    TB.Buttons(3).Enabled = True
    clearcontrol
    cbotrnid.Enabled = True
    Case "SAVE"
    MNU_SAVE
       
    Case "EXIT"
    Unload Me
End Select
End Sub


Private Sub MNU_SAVE()
Dim i As Integer

ODKDB.Execute "delete from tbltshowogdistancematrix where headerid='" & cbotrnid.BoundText & "'"
For i = 1 To mygrid.Rows - 1
If Len((mygrid.TextMatrix(i, 1))) = 0 Or Len((mygrid.TextMatrix(i, 2))) = 0 Or Val((mygrid.TextMatrix(i, 3))) = 0 Or Len((mygrid.TextMatrix(i, 4))) = 0 Then Exit For
ODKDB.Execute "insert into tbltshowogdistancematrix (headerid,slno,frmedge,toedge, " _
& " distance,roadtype,note) values('" & cbotrnid.BoundText & "','" & Val(mygrid.TextMatrix(i, 0)) & "', " _
& " '" & Mid(mygrid.TextMatrix(i, 1), 1, 9) & "','" & Mid(mygrid.TextMatrix(i, 2), 1, 9) & "', " _
& " '" & Val(mygrid.TextMatrix(i, 3)) & "','" & Mid(mygrid.TextMatrix(i, 4), 1, 3) & "', " _
& " '" & mygrid.TextMatrix(i, 5) & "' )"

Next

End Sub

Private Sub putslno()
Dim i As Integer
For i = 1 To mygrid.Rows - 1
If Len(mygrid.TextMatrix(i, 1)) = 0 Then Exit Sub
mygrid.TextMatrix(i, 0) = i

Next

End Sub



Private Sub FillGrid()
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
rs.Open "select * from tbltshowogdistancematrix where headerid='" & cbotrnid.BoundText & "'", ODKDB
i = 1
Do While rs.EOF <> True
cbodzongkhag.Text = finddescription("tbldzongkhag", Mid(rs!fromedge, 1, 3), "DzongkhagCode", "DzongkhagName")
cbodzongkhagdest.Text = finddescription("tbldzongkhag", Mid(rs!toedge, 1, 3), "DzongkhagCode", "DzongkhagName")

cbogewog.Text = finddescription("tblgewog", Mid(rs!fromedge, 1, 6), "concat(DzongkhagId,GewogId)", "GewogName")
cbogewogdest.Text = finddescription("tblgewog", Mid(rs!toedge, 1, 6), "concat(DzongkhagId,GewogId)", "GewogName")


mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = finddescription("tbltshewog", rs!fromedge, "concat(DzongkhagId,GewogId,TshewogId)", "TshewogName")

mygrid.TextMatrix(i, 2) = finddescription("tbltshewog", rs!toedge, "concat(DzongkhagId,GewogId,TshewogId)", "TshewogName")
mygrid.TextMatrix(i, 3) = rs!distance
mygrid.TextMatrix(i, 4) = finddescription("tblroadtype", rs!roadtype, "typeid", "typedesc")
mygrid.TextMatrix(i, 5) = rs!note
mygrid.TextMatrix(i, 6) = "M"

mygrid.ColAlignment(i) = flexAlignLeftTop
mygrid.ColAlignment(2) = flexAlignLeftTop

i = i + 1
rs.MoveNext
Loop





End Sub
Private Sub updatelatlng()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Set rs = Nothing
rs.Open "select frmedge from odk_prodlocal.tbltshowogdistancematrix group by frmedge", ODKDB
Do While rs.EOF <> True

Set rs1 = Nothing
rs1.Open "select concat(substring(REGION_DCODE,1,3),substring(REGION_GCODE,1,3),substring(REGION,1,3))" _
& " as dgtstr," _
& " REGION_DCODE,REGION_GCODE,REGION,AVG(GPS_COORDINATES_LAT) GPS_COORDINATES_LAT, " _
& " AVG(GPS_COORDINATES_LNG) GPS_COORDINATES_LNG " _
& " from odk_prodlocal.tblfieldlastvisitrpt where GPS_COORDINATES_LAT>0 and GPS_COORDINATES_LNG>0 " _
& " and concat(substring(REGION_DCODE,1,3),substring(REGION_GCODE,1,3),substring(REGION,1,3))='" & rs!frmedge & "' " _
& " group by concat(substring(REGION_DCODE,1,3),substring(REGION_GCODE,1,3),substring(REGION,1,3))", ODKDB
If rs1.EOF <> True Then
ODKDB.Execute "update tbltshowogdistancematrix set frmlat='" & rs1!GPS_COORDINATES_LAT & "' ,frmlng='" & rs1!GPS_COORDINATES_LNG & "' where frmedge='" & rs!frmedge & "'"
End If
rs.MoveNext
Loop


Set rs = Nothing
rs.Open "select toedge from odk_prodlocal.tbltshowogdistancematrix group by toedge", ODKDB
Do While rs.EOF <> True

Set rs1 = Nothing
rs1.Open "select concat(substring(REGION_DCODE,1,3),substring(REGION_GCODE,1,3),substring(REGION,1,3))" _
& " as dgtstr," _
& " REGION_DCODE,REGION_GCODE,REGION,AVG(GPS_COORDINATES_LAT) GPS_COORDINATES_LAT, " _
& " AVG(GPS_COORDINATES_LNG) GPS_COORDINATES_LNG " _
& " from odk_prodlocal.tblfieldlastvisitrpt where GPS_COORDINATES_LAT>0 and GPS_COORDINATES_LNG>0 " _
& " and concat(substring(REGION_DCODE,1,3),substring(REGION_GCODE,1,3),substring(REGION,1,3))='" & rs!toedge & "' " _
& " group by concat(substring(REGION_DCODE,1,3),substring(REGION_GCODE,1,3),substring(REGION,1,3))", ODKDB
If rs1.EOF <> True Then
ODKDB.Execute "update tbltshowogdistancematrix set tolat='" & rs1!GPS_COORDINATES_LAT & "' ,tolng='" & rs1!GPS_COORDINATES_LNG & "' where toedge='" & rs!toedge & "'"
End If




rs.MoveNext
Loop




End Sub
