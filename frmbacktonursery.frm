VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmbacktonursery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BACK TO NURSERY"
   ClientHeight    =   9990
   ClientLeft      =   2895
   ClientTop       =   585
   ClientWidth     =   15315
   Icon            =   "frmbacktonursery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   15315
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   7455
      Begin VB.TextBox txtyr 
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
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "frmbacktonursery.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1440
         TabIndex        =   21
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
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   6000
         TabIndex        =   22
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   103088129
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbostaff 
         Bindings        =   "frmbacktonursery.frx":0E57
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Distribution No."
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
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5400
         TabIndex        =   25
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Back To Nursery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   7455
      Begin MSDataListLib.DataCombo cbofacility 
         Bindings        =   "frmbacktonursery.frx":0E6C
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2280
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.TextBox txtnoofcrates 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtqty 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   3540
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6244
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "       |^ Batch No.|^Variety |^Facility |^No. Of Crates|^Crate #                    |Qty.         |"
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
      Begin MSDataListLib.DataCombo cbovariety 
         Bindings        =   "frmbacktonursery.frx":0E81
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1440
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
      End
      Begin VB.Line Line3 
         X1              =   5280
         X2              =   8880
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Label Label2 
         Caption         =   "Remarks :"
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
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Crate Release"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   7560
      TabIndex        =   6
      Top             =   720
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "Lock All"
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
         Left            =   4080
         Picture         =   "frmbacktonursery.frx":0E96
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   8400
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Release All"
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
         Left            =   5880
         Picture         =   "frmbacktonursery.frx":1220
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   8400
         Width           =   1695
      End
      Begin VB.ListBox DZLIST 
         Columns         =   8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         ItemData        =   "frmbacktonursery.frx":15AA
         Left            =   120
         List            =   "frmbacktonursery.frx":15AC
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Send To Field  Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   7215
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6975
         _cx             =   12303
         _cy             =   3625
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
         BackColorAlternate=   12648447
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmbacktonursery.frx":15AE
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
   End
   Begin VB.TextBox txtlockedcrates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   9600
      Width           =   1095
   End
   Begin VB.TextBox txtsendtofieldqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   8640
      Width           =   1095
   End
   Begin VB.TextBox txtbacktonurseryqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtactualdistributed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   9120
      Width           =   1095
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   5280
      Top             =   0
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
            Picture         =   "frmbacktonursery.frx":1699
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbacktonursery.frx":1A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbacktonursery.frx":1DCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbacktonursery.frx":2AA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbacktonursery.frx":2EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmbacktonursery.frx":36B3
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
      Width           =   15315
      _ExtentX        =   27014
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Locked Crates"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   9720
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sent To Field"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   8760
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Back To Nursery"
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
      Left            =   4560
      TabIndex        =   10
      Top             =   5880
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Actual Distributed"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   9240
      Width           =   1530
   End
End
Attribute VB_Name = "frmbacktonursery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ValidRow As Boolean
Dim totbacktonursary, totsendtofield As Long
Dim Dzstr As String
Dim CurrRow As Long

Private Sub cbofacility_LostFocus()
Dim Issue, Recv As Double
Dim rs As New ADODB.Recordset
ItemGrd.TextMatrix(CurrRow, 3) = cbofacility.BoundText
cbofacility.Visible = False
ItemGrd.ColWidth(3) = 750
End Sub

Private Sub cbotrnid_LostFocus()
If Len(cbotrnid.Text) = 0 Then Exit Sub
cbotrnid.Enabled = False
TB.buttons(3).Enabled = True
txtyr.Text = Mid(cbotrnid.Text, Len(cbotrnid.Text) - 4, 5)
txtyr.Text = Trim(txtyr.Text)
cbotrnid.Text = cbotrnid.BoundText
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsbacktonurseryhdr where distributionno='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindsTAFF rs!staffid
cbostaff.Text = rs!staffid & "  " & sTAFF
txtentrydate.Value = Format(rs!entrydate, "dd/MM/yyyy")
txtsendtofieldqty.Text = rs!sendtofieldqty
txtbacktonurseryqty.Text = rs!backtonursaryqty
txtactualdistributed.Text = rs!actualdistributed
txtlockedcrates.Text = rs!lockedcrate

End If


fillbacktonursery
fillsendtofield
loadcrate
If DZLIST.ListCount <> 0 Then
Frame2.Enabled = True
frmbacktonursery.Width = 15405
End If
End Sub
Private Sub loadcrate()
Dim cnt As Integer
Dim rs As New ADODB.Recordset
Dim i, j As Integer
Dim dd As Variant
Set rs = Nothing

DZLIST.Clear
rs.Open "select distinct crateno from tblqmssendtofielddetail where distributionno='" & cbotrnid.BoundText & "' and cratestatus='ON' and year='" & Val(txtyr.Text) & "' Order by  crateno", MHVDB
With rs
Do While Not .EOF
   DZLIST.AddItem Trim(!crateno)
    .MoveNext
Loop
End With




For i = 0 To DZLIST.ListCount - 1
DZLIST.Selected(i) = True
Next

txtlockedcrates.Text = DZLIST.ListCount



End Sub
Private Sub fillsendtofield()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
j = 0
i = 1
totdsheet = 0
mygrid.Clear
mygrid.FormatString = "^S/N|^Batch No.|^Facility|^No. of Crates|^Crate #|^Qty.|"
mygrid.ColWidth(0) = 585
mygrid.ColWidth(1) = 975
mygrid.ColWidth(2) = 750
mygrid.ColWidth(3) = 795
mygrid.ColWidth(4) = 1275
mygrid.ColWidth(5) = 1410
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 375

If Len(cbotrnid.Text) = 0 Then Exit Sub

totsendtofield = 0
mygrid.rows = 1
Set rs = Nothing
rs.Open "select * from tblqmsplanttransaction where transactiontype='4' and distributionno='" & cbotrnid.BoundText & "' and status='ON'", MHVDB

Do While rs.EOF <> True
mygrid.rows = mygrid.rows + 1
findQmsBatchDetail rs!plantBatch
mygrid.TextMatrix(i, 0) = i
mygrid.TextMatrix(i, 1) = rs!plantBatch
mygrid.TextMatrix(i, 2) = qmsplantbatch3
mygrid.TextMatrix(i, 3) = rs!facilityid
mygrid.TextMatrix(i, 4) = rs!cratecount
mygrid.TextMatrix(i, 5) = rs!crateno
mygrid.TextMatrix(i, 6) = rs!credit
totsendtofield = totsendtofield + rs!credit
i = i + 1
rs.MoveNext

Loop




txtsendtofieldqty.Text = totsendtofield
txtactualdistributed.Text = Val(txtsendtofieldqty.Text) - Val(txtbacktonurseryqty.Text)
End Sub

Private Sub fillbacktonursery()
Dim rs As New ADODB.Recordset
Dim i As Integer
i = 1
ItemGrd.Clear
ItemGrd.FormatString = "       |^ Batch No.|^Variety |^Facility |^No. Of Crates|^ Crate #                 |Qty.      |"
Set rs = Nothing
If Operation = "OPEN" Then
rs.Open "select * from tblqmsplanttransaction where transactiontype='5' and distributionno='" & cbotrnid.BoundText & "' and status='ON'", MHVDB
If rs.EOF <> True Then

Else
Set rs = Nothing
rs.Open "select * from tblqmsplanttransaction where  distributionno='" & cbotrnid.BoundText & "' and status='ON'", MHVDB
End If
Else
rs.Open "select * from tblqmsplanttransaction where  distributionno='" & cbotrnid.BoundText & "' and status='ON'", MHVDB
End If
Do While rs.EOF <> True
ItemGrd.rows = ItemGrd.rows + 1
findQmsBatchDetail rs!plantBatch
ItemGrd.TextMatrix(i, 1) = rs!plantBatch
ItemGrd.TextMatrix(i, 2) = qmsplantbatch3
ItemGrd.TextMatrix(i, 3) = rs!facilityid
ItemGrd.TextMatrix(i, 4) = rs!cratecount
ItemGrd.TextMatrix(i, 5) = rs!crateno
ItemGrd.TextMatrix(i, 6) = IIf(rs!debit = 0, "", rs!debit)

i = i + 1
rs.MoveNext

Loop

End Sub

Private Sub Command1_Click()
For i = 0 To DZLIST.ListCount - 1
DZLIST.Selected(i) = True
Next
txtlockedcrates.Text = DZLIST.ListCount
End Sub

Private Sub Command2_Click()
For i = 0 To DZLIST.ListCount - 1
DZLIST.Selected(i) = False

Next
txtlockedcrates.Text = ""
End Sub

Private Sub DZLIST_Click()
Dim lockcount As Integer
lockcount = 0
For i = 0 To DZLIST.ListCount - 1
If DZLIST.Selected(i) = True Then
lockcount = lockcount + 1
End If
Next
txtlockedcrates.Text = IIf(lockcount = 0, "", lockcount)
End Sub

Private Sub Form_Load()
Dim cnt As Integer
Dim rs As New ADODB.Recordset
Dim i, j As Integer
Dim dd As Variant
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                   
db.Open CnnString
frmbacktonursery.Width = 7650
Set rs = Nothing
i = 1

Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct distributionno,concat(cast(distributionno as char) , '  ', cast(year as char)) dist  from tblqmssendtofieldhdr  where status<>'C' order by distributionno desc", db
Set cbotrnid.RowSource = rs
cbotrnid.ListField = "dist"
cbotrnid.BoundColumn = "distributionno"




Set rs = Nothing
If rs.State = adStateOpen Then rsF.Close
If rs.State = adStateOpen Then Srs.Close
rs.Open "select concat(STAFFCODE,'  ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE nursery='1'  order by STAFFCODE", db
Set cbostaff.RowSource = rs
cbostaff.ListField = "STAFFNAME"
cbostaff.BoundColumn = "STAFFCODE"

'ItemGrd.col = 3
'cbofacility.Left = ItemGrd.Left + ItemGrd.CellLeft
'cbofacility.Width = 2000 'ItemGrd.CellWidth
'cbofacility.Height = ItemGrd.CellHeight


End Sub
Private Sub CLEARCONTROLL()
ItemGrd.Clear

ItemGrd.FormatString = "       |^ Batch No.|^Variety |^Facility |^No. Of Crates|^ Crate #                 |Qty.      |"
ItemGrd.rows = 1
txtentrydate.Value = Format(Now, "dd/MM/yyyy")
cbotrnid.Text = ""
txtbacktonurseryqty.Text = ""
cbostaff.Text = ""
txtsendtofieldqty.Text = ""
txtactualdistributed.Text = ""
txtlockedcrates.Text = ""
txtyr.Text = ""
End Sub
Private Sub populatedno(Operation As String)

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
Dim CONNLOCAL As New ADODB.Connection
CONNLOCAL.Open OdkCnnString
                   
db.Open CnnString

Set rs = Nothing
If Operation = "ADD" Then
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct distributionno,concat(cast(distributionno as char) , '  ', cast(distyear as char)) dist  from tblqmsplanttransaction  where status<>'C' and distributionno not in(select distributionno from tblqmsbacktonurseryhdr where status='ON') and distributionno>0 order by distributionno", db


Set cbotrnid.RowSource = rs
cbotrnid.ListField = "dist"
cbotrnid.BoundColumn = "distributionno"
ElseIf Operation = "OPEN" Then
If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct distributionno,concat(cast(distributionno as char) , '  ', cast(distyear as char)) dist  from tblqmsplanttransaction  where status<>'C' and distributionno  in(select distributionno from tblqmsbacktonurseryhdr where status='ON') order by distributionno", db
Set cbotrnid.RowSource = rs
cbotrnid.ListField = "dist"
cbotrnid.BoundColumn = "distributionno"
Else
MsgBox "Wrong Operation Selected."
End If
End Sub

Private Sub ItemGrd_DblClick()
Dim mrow, MCOL As Integer
'txtselected.Visible = False
'ItemGrd.ColWidth(3) = 750
'If Not ValidRow And CurrRow <> ItemGrd.row Then
'   ItemGrd.row = CurrRow
'   Exit Sub
'End If
If Len(cbotrnid.Text) = 0 Then Exit Sub
mrow = ItemGrd.row
MCOL = ItemGrd.col
If mrow = 0 Then Exit Sub
If mrow > 1 And Len(ItemGrd.TextMatrix(mrow - 1, 4)) = 0 Then
   Beep
   Exit Sub
End If
ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = mrow
ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)

Select Case MCOL
              Case 3
       If Trim(Len(ItemGrd.TextMatrix(mrow, 2))) > 0 Then
            ItemGrd.ColWidth(3) = 2000
            cbofacility.Top = ItemGrd.Top + ItemGrd.CellTop
            cbofacility = ItemGrd.Text
            cbofacility.Visible = True
            cbofacility.SetFocus
            End If
            
            Dim rs As New ADODB.Recordset
            Set db = New ADODB.Connection
            db.CursorLocation = adUseClient
            Dim CONNLOCAL As New ADODB.Connection
            CONNLOCAL.Open OdkCnnString
                         
            db.Open CnnString
            
            Set rs = Nothing
            
            If rs.State = adStateOpen Then rs.Close
            rs.Open "select distinct facilityid,description  from tblqmsfacility where infotype='NUR' order by facilityid", db
            Set cbofacility.RowSource = rs
            cbofacility.ListField = "description"
            cbofacility.BoundColumn = "facilityid"
      
       Case 6
        txtqty.Left = ItemGrd.Left + ItemGrd.CellLeft
        txtqty.Width = ItemGrd.CellWidth
        txtqty.Height = ItemGrd.CellHeight
       If Len(ItemGrd.TextMatrix(mrow, 1)) > 0 Then
            txtqty.Top = ItemGrd.Top + ItemGrd.CellTop
            txtqty = ItemGrd.Text
            txtqty.Visible = True
            txtqty.SetFocus
       End If
    
            
    End Select
End Sub
Private Sub getsum()
Dim i As Integer
totbacktonursary = 0
'totcrate = 0
For i = 1 To ItemGrd.rows - 1
If Len(ItemGrd.TextMatrix(i, 1)) = 0 Then Exit For
totbacktonursary = totbacktonursary + Val(ItemGrd.TextMatrix(i, 6))
'totcrate = totcrate + Val(ItemGrd.TextMatrix(i, 4))
Next

txtbacktonurseryqty.Text = totbacktonursary
'txttocrate.Text = totcrate
'txtshortexcees.Text = Val(txtdsheetqty.Text) - Val(txtsendtofield.Text)
txtactualdistributed.Text = Val(txtsendtofieldqty.Text) - Val(txtbacktonurseryqty.Text)


End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
       Operation = "ADD"
       populatedno "ADD"
       CLEARCONTROLL
       cbotrnid.Enabled = True
       TB.buttons(3).Enabled = False
       Case "OPEN"
       Operation = "OPEN"
       populatedno "OPEN"
       CLEARCONTROLL
       cbotrnid.Enabled = True
       TB.buttons(3).Enabled = False
       
       Case "SAVE"
       MNU_SAVE
        TB.buttons(3).Enabled = False
        'FillGrid
       
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
Dim rs As New ADODB.Recordset
Dim dd As Variant
Dim mm As Variant
Dim bb As Variant
Dim cnt As Integer
Dim mMaxId As Long
Dim crateStr As String
Dim i As Integer
If Len(cbotrnid.Text) = 0 Then
MsgBox "Distribution No. cannot be empty."
Exit Sub
End If

'If Val(txtbacktonurseryqty.Text) <= 0 Then
'MsgBox "Invalid back to nursery quantity"
'Exit Sub
'End If

If Val(txtyr.Text) <= 0 Then
MsgBox "Invalid distribution no."
Exit Sub
End If



    Set rs = Nothing
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmsplanttransaction", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        mMaxId = IIf(IsNull(rs!MaxId), "1", rs!MaxId)
        Else
        mMaxId = rs!MaxId
        End If

MHVDB.BeginTrans

 crateStr = ""
' For i = 1 To ItemGrd.Rows - 1
' If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
' mm = Split(Trim(ItemGrd.TextMatrix(i, 5)), ",", -1, vbTextCompare)
'cnt = Len(Trim(ItemGrd.TextMatrix(i, 5))) - Len(Replace(Trim(ItemGrd.TextMatrix(i, 5)), ",", ""))
'For j = 0 To cnt
'crateStr = Trim(ItemGrd.TextMatrix(i, 1)) & "|" & mm(j) & "," & crateStr
'Next
'
' Next
' crateStr = Left(crateStr, Len(crateStr) - 1)
 
If Operation = "ADD" Then
MHVDB.Execute "insert into tblqmsbacktonurseryhdr(distributionno,entrydate,staffid" _
& " ,backtonursaryqty,sendtofieldqty,actualdistributed,vehicleno,drivername,status,lockedcrate,year) values(" _
& "'" & cbotrnid.BoundText & "'," _
& "'" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
& "'" & cbostaff.BoundText & "'," _
& "'" & Val(txtbacktonurseryqty.Text) & "'," _
& "'" & Val(txtsendtofieldqty.Text) & "'," _
& "'" & Val(txtactualdistributed.Text) & "'," _
& "''," _
& "''," _
& "'ON'," _
& "'" & Val(txtlockedcrates.Text) & "','" & Val(txtyr.Text) & "')"

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsbacktonurseryhdr set " _
& "entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
& "staffid='" & cbostaff.BoundText & "'," _
& "backtonursaryqty='" & Val(txtbacktonurseryqty.Text) & "'," _
& "sendtofieldqty='" & Val(txtsendtofieldqty.Text) & "'," _
& "actualdistributed='" & Val(txtactualdistributed.Text) & "'," _
& "vehicleno=''," _
& "drivername=''," _
& "status='ON'," _
& "lockedcrate='" & Val(txtlockedcrates.Text) & "' where distributionno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "'"
Else
MsgBox "No Operation Selected."
MHVDB.RollbackTrans
Exit Sub
End If


' release crate here
Dzstr = ""
For i = 0 To DZLIST.ListCount - 1
    If Not DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim((DZLIST.List(i))) + "',"
    End If
    
Next

If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
Dzstr = "(" + "'" + A99 & "'" & ")"
'   MsgBox "No Crates are released !!!"
'   frmbacktonursery.Width = 7650
'   MHVDB.RollbackTrans
'   Exit Sub
End If

MHVDB.Execute "update tblqmssendtofielddetail set cratestatus='C' where distributionno='" & cbotrnid.BoundText & "' and year='" & Val(txtyr.Text) & "' " _
& " and crateno in" & Dzstr

MHVDB.Execute "update tblqmscrate set locked='0' where crateno in" & Dzstr












If Val(txtbacktonurseryqty.Text) > 0 Then
 
MHVDB.Execute "delete from tblqmsplanttransaction where distributionno='" & cbotrnid.BoundText & "' and distyear='" & Val(txtyr.Text) & "' " _
& "and verificationtype='2' and transactiontype='5'"


For i = 1 To ItemGrd.rows - 1
If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
findQmsBatchDetail Trim(ItemGrd.TextMatrix(i, 1))
getLocationFromFid Trim(ItemGrd.TextMatrix(i, 3))
MHVDB.Execute "INSERT INTO tblqmsplanttransaction (trnid,entrydate,plantbatch,varietyid," _
            & "facilityid,verificationtype,transactiontype,staffid,debit,credit,status,location,distributionno,crateno,cratecount,distyear) " _
            & "VALUEs('" & mMaxId & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & Trim(ItemGrd.TextMatrix(i, 1)) & "', " _
            & "'" & mPlantVariety & "','" & Trim(ItemGrd.TextMatrix(i, 3)) & "'," _
            & "'2','5', " _
            & "'" & cbostaff.BoundText & "','" & Val(ItemGrd.TextMatrix(i, 6)) & "','0','ON','" & locationFromFid & "','" & cbotrnid.BoundText & "','" & Trim(ItemGrd.TextMatrix(i, 5)) & "','" & Trim(ItemGrd.TextMatrix(i, 4)) & "','" & Val(txtyr.Text) & "')"
            
            
'If mPlantVariety = 12 Then
'
'MHVDB.Execute "INSERT INTO tblqmsnutdetail (trnid,entrydate,plantbatch,varietyid," _
'            & "facilityid,nuttype,transactiontype,nutdebit,nutcredit,status,location) " _
'            & "VALUEs('" & mMaxId & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & Trim(ItemGrd.TextMatrix(i, 1)) & "', " _
'            & "'" & mPlantVariety & "','" & Trim(ItemGrd.TextMatrix(i, 3)) & "'," _
'            & "'GN','5', " _
'            & "'" & Val(ItemGrd.TextMatrix(i, 6)) & "','0','ON','" & Mlocation & "')"
'
'
'End If

mMaxId = mMaxId + 1
 Next
 






End If
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='4'"
MHVDB.CommitTrans
frmbacktonursery.Width = 7650
Exit Sub
err:
    MHVDB.RollbackTrans
MsgBox err.Description



frmbacktonursery.Width = 7650
End Sub

Private Sub txtQty_validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not IsNumeric(txtqty) Then
   Beep
   MsgBox "Enter a valid No."
   ValidRow = False
   Cancel = True
   Exit Sub
'ElseIf Val(txtqty.Text) > 35 Then
' Beep
'   MsgBox "Enter a valid No."
'   ValidRow = False
'   Cancel = True
'   Exit Sub
Else
  
   ItemGrd.TextMatrix(CurrRow, 6) = Val(txtqty.Text)
   ValidRow = True
   
End If
End If
txtqty.Visible = False
getsum
End Sub
