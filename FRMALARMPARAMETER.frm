VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMALARMPARAMETER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THRESHOLD PARAMETER SETTINGS...."
   ClientHeight    =   8295
   ClientLeft      =   3525
   ClientTop       =   1260
   ClientWidth     =   12000
   Icon            =   "FRMALARMPARAMETER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12000
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7560
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
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
         Height          =   2205
         Left            =   0
         TabIndex        =   27
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Formula Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7560
      TabIndex        =   24
      Top             =   3360
      Width           =   3255
      Begin VB.TextBox txtformula 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "Double Click in the formula box to view the  field reference."
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Value Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   21
      Top             =   3360
      Width           =   2055
      Begin VB.OptionButton optpercentage 
         Caption         =   "Percentage"
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
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optno 
         Caption         =   "Flat Number"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5520
      TabIndex        =   15
      Top             =   4080
      Width           =   2055
      Begin VB.CheckBox chkisfarmercode 
         Caption         =   "Is Farmer Code"
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
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkisfieldcode 
         Caption         =   "Is Field Code"
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
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkisstaffcode 
         Caption         =   "Is Staff Code"
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
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      Begin MSComCtl2.DTPicker TXTAPPLICABLEFROM 
         Height          =   375
         Left            =   1920
         TabIndex        =   34
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   80281601
         CurrentDate     =   41495
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ComboBox CBOFIELDSTORAGE 
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
         ItemData        =   "FRMALARMPARAMETER.frx":27A2
         Left            =   4080
         List            =   "FRMALARMPARAMETER.frx":27AC
         TabIndex        =   29
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox TXTRECEIPIENTS 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtvalue 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
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
         ItemData        =   "FRMALARMPARAMETER.frx":27C0
         Left            =   1920
         List            =   "FRMALARMPARAMETER.frx":27C2
         TabIndex        =   2
         Top             =   1200
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo CBOTBL 
         Bindings        =   "FRMALARMPARAMETER.frx":27C4
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cboparaid 
         Bindings        =   "FRMALARMPARAMETER.frx":27D9
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
      Begin MSDataListLib.DataCombo cbostatus 
         Bindings        =   "FRMALARMPARAMETER.frx":27EE
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSDataListLib.DataCombo CBOREPORT 
         Bindings        =   "FRMALARMPARAMETER.frx":2803
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "APPLICABLE FROM"
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
         TabIndex        =   33
         Top             =   4320
         Width           =   1710
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FD./ST."
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
         Left            =   3360
         TabIndex        =   32
         Top             =   3960
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
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
         TabIndex        =   30
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FD/ST"
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
         Left            =   3360
         TabIndex        =   28
         Top             =   3480
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "RECEIPIENTS"
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
         TabIndex        =   20
         Top             =   3360
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "REPORT NAME"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "MAX. ACCEPTABLE THRESHOLD VALUE"
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
         Top             =   2760
         Width           =   3570
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PARAMETER ID"
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
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TABLE"
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
         TabIndex        =   5
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PARAMETER NAME"
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
         Top             =   1320
         Width           =   1740
      End
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
            Picture         =   "FRMALARMPARAMETER.frx":2818
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMALARMPARAMETER.frx":2BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMALARMPARAMETER.frx":2F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMALARMPARAMETER.frx":3C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMALARMPARAMETER.frx":4078
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMALARMPARAMETER.frx":4832
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
      Width           =   12000
      _ExtentX        =   21167
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
   Begin VSFlex7Ctl.VSFlexGrid Mygrid 
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   5400
      Width           =   11895
      _cx             =   20981
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FRMALARMPARAMETER.frx":4BCC
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
Attribute VB_Name = "FRMALARMPARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBOFIELDSTORAGE_LostFocus()
If CBOFIELDSTORAGE.Text = "Storage" Then
chkisfieldcode.Enabled = False
chkisfieldcode.Value = 0
chkisfarmercode.Value = 1
chkisstaffcode.Value = 1
Else
chkisfieldcode.Enabled = True
chkisfieldcode.Value = 1
chkisfarmercode.Value = 1
chkisstaffcode.Value = 1
End If

End Sub

Private Sub cboparaid_LostFocus()
 On Error GoTo err
   
   cboparaid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblodkalarmparameter where paraid='" & cboparaid.BoundText & "'", ODKDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
  findtablename rs!odktable
  findreportname rs!REPORT
  FindqmsStatus rs!Status
  CBOTBL.Text = odkTableName
  CBODATE.Text = rs!paraname
  CBOREPORT.Text = ReportName
  txtvalue.Text = rs!Value
  cbostatus.Text = qmsStatus
  chkisstaffcode.Value = rs!isstaffcode
  chkisfarmercode.Value = rs!isfarmercode
  chkisfieldcode.Value = rs!isfieldcode
  TXTRECEIPIENTS.Text = rs!receipents
  optno.Value = rs!flatno
  optpercentage.Value = rs!percentage
  txtformula.Text = rs!Formula
  CBOFIELDSTORAGE.Text = rs!fstype
  txtdesc.Text = rs!Description
  TXTAPPLICABLEFROM.Value = IIf(IsNull(rs!applicablefrom), "01/01/1999", Format(rs!applicablefrom, "dd/MM/yyyy"))
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub

Private Sub CBOTBL_GotFocus()
Frame5.Visible = False
End Sub

Private Sub CBOTBL_LostFocus()
Dim i, j, fcount As Integer
'Operation = ""
'Mygrid.Visible = False
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
CBODATE.AddItem rs.Fields(j).Name
Next

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub chkisfarmercode_Click()
If chkisfarmercode.Value = 1 Then
chkisstaffcode.Value = 1
Else
chkisfieldcode.Value = 0
End If
End Sub

Private Sub chkisfieldcode_Click()
If chkisfieldcode.Value = 1 Then
chkisstaffcode.Value = 1
chkisfarmercode.Value = 1
End If
End Sub

Private Sub chkisstaffcode_Click()
If chkisstaffcode.Value = 0 Then
chkisfarmercode.Value = 0
chkisfieldcode.Value = 0
End If
End Sub

Private Sub DZLIST_DblClick()
 Dim iPos As Long
'txtformula.SelText = DZLIST.Selected(i)
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
      txtformula.SelText = DZLIST.List(i)
     
'With txtformula
'  If .SelLength = 0 Then
'    iPos = .SelStart
'  Else
'    iPos = .SelStart + .SelLength
'  End If
'  'Debug.Print "The current cursor position in " & .Name & " is: " & iPos & " :-)"
'  txtformula.Text = DZLIST.List(1)
'End With
    End If
    
Next
End Sub

Private Sub Form_Load()
On Error GoTo err

Operation = ""

Dim rs As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open OdkCnnString

Set db1 = New ADODB.Connection
db1.CursorLocation = adUseClient
db1.Open CnnString
                     

Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select *  from tbltable where status='ON' order by tblid", db
Set CBOTBL.RowSource = rs
CBOTBL.ListField = "TBLNAME"
CBOTBL.BoundColumn = "TBLID"



Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open "select ID,REPORTNAME  from tblemaillog where status='ON' order by ID", db
Set CBOREPORT.RowSource = rs
CBOREPORT.ListField = "reportname"
CBOREPORT.BoundColumn = "id"

Set rs = Nothing
If rs.State = adStateOpen Then rs.Close
rs.Open " select concat(cast(paraid as char),'  ',paraname,'  ',fstype,'  ',cast(value as char)) as description,paraid from tblodkalarmparameter where status='ON'", db
Set cboparaid.RowSource = rs
cboparaid.ListField = "description"
cboparaid.BoundColumn = "paraid"

Set rs = Nothing

If rs.State = adStateOpen Then rs.Close
rs.Open "select statusid,status  from tblqmsstatus order by status", db1
Set cbostatus.RowSource = rs
cbostatus.ListField = "status"
cbostatus.BoundColumn = "statusid"

FillGrid

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cboparaid.Enabled = False
        TB.Buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(paraid)+1 AS MaxID from tblodkalarmparameter", ODKDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cboparaid.Text = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
        Else
        cboparaid.Text = rs!MaxId
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cboparaid.Text = ""
        cboparaid.Enabled = True
        TB.Buttons(3).Enabled = True
             
       Case "SAVE"
        MNU_SAVE
     
        FillGrid
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CLEARCONTROLL()
    CBOTBL.Text = ""
   CBODATE.Text = ""
     txtvalue.Text = ""
     cbostatus.Text = ""
     CBOREPORT.Text = ""
     TXTRECEIPIENTS.Text = ""
     chkisstaffcode.Value = 1
     chkisfarmercode.Value = 0
     chkisfieldcode.Value = 0
     txtformula.Text = ""
     optno.Value = False
     optpercentage.Value = False
     CBOFIELDSTORAGE.Text = ""
    txtdesc.Text = ""
End Sub
Private Sub FillGrid()

On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^Sl.No.|^PARAMETER ID|^TABLE NAME|^PARAMETER NAME|^REPORT|^VALUE|^STATUS|^"
mygrid.ColWidth(0) = 750
mygrid.ColWidth(1) = 1470
mygrid.ColWidth(2) = 3105
mygrid.ColWidth(3) = 1875
mygrid.ColWidth(4) = 2805
mygrid.ColWidth(5) = 720
mygrid.ColWidth(6) = 960
mygrid.ColWidth(7) = 135


rs.Open "select * from tblodkalarmparameter order by paraid", ODKDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i

mygrid.TextMatrix(i, 1) = rs!paraid
findtablename rs!odktable
mygrid.TextMatrix(i, 2) = odkTableName
mygrid.TextMatrix(i, 3) = UCase(rs!paraname)
findreportname rs!REPORT
mygrid.TextMatrix(i, 4) = ReportName
mygrid.TextMatrix(i, 5) = rs!Value
FindqmsStatus rs!Status
mygrid.TextMatrix(i, 6) = qmsStatus
rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub MNU_SAVE()
Dim mNoType As Integer
Dim mPercentageType As Integer
Dim rs As New ADODB.Recordset
On Error GoTo err
If Len(cboparaid.Text) = 0 Then
MsgBox "Select Parameter Id."
Exit Sub
End If
If optno.Value = True Then
mNoType = 1
Else
mNoType = 0
End If
If Len(txtformula.Text) = 0 Then
MsgBox "Formula is Must.For Falt Number,Type only the parameter name."
Exit Sub
End If
If Len(txtdesc.Text) = 0 Then
MsgBox "Please specify param description."
Exit Sub
End If

If optpercentage.Value = True Then
mPercentageType = 1
Else
mPercentageType = 0
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
ODKDB.Execute "INSERT INTO tblodkalarmparameter (paraid,paraname,description,applicablefrom,odktable," _
            & "report,value,status,isstaffcode,isfarmercode,isfieldcode,receipents,flatno,percentage,formula,fstype) " _
            & "VALUEs('" & cboparaid.Text & "','" & CBODATE.Text & "','" & txtdesc.Text & "','" & Format(TXTAPPLICABLEFROM.Value, "yyyy-MM-dd") & "','" & CBOTBL.BoundText & "', " _
            & "'" & CBOREPORT.BoundText & "','" & Val(txtvalue.Text) & "','" & cbostatus.BoundText & "'," _
            & "'" & chkisstaffcode.Value & "','" & chkisfarmercode.Value & "'," _
            & "'" & chkisfieldcode.Value & "','" & TXTRECEIPIENTS.Text & "','" & mNoType & "','" & mPercentageType & "','" & txtformula.Text & "','" & CBOFIELDSTORAGE.Text & "')"
 
 
LogRemarks = "Inserted new record" & cboparaid.BoundText & "," & " into table tblodkalarmparameter "
updateodklog "NA", Format(Now, "yyyy-MM-dd"), MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then

ODKDB.Execute "update tblodkalarmparameter set paraname='" & CBODATE.Text & "',description='" & txtdesc.Text & "'" _
            & ",odktable='" & CBOTBL.BoundText & "',report='" & CBOREPORT.BoundText & "' " _
            & ",value='" & Val(txtvalue.Text) & "',status='" & cbostatus.BoundText & "',applicablefrom='" & Format(TXTAPPLICABLEFROM.Value, "yyyy-MM-dd") & "', " _
            & "isstaffcode='" & chkisstaffcode.Value & "',isfarmercode='" & chkisfarmercode.Value & "'," _
            & "isfieldcode='" & chkisfieldcode.Value & "',receipents='" & TXTRECEIPIENTS.Text & "'," _
            & "flatno='" & mNoType & "',percentage='" & mPercentageType & "',formula='" & txtformula.Text & "',fstype='" & CBOFIELDSTORAGE.Text & "'" _
            & " where paraid='" & cboparaid.BoundText & "' "

LogRemarks = "Updated  record" & cboparaid.BoundText & "," & " on table tblodkalarmparameter "
updateodklog "NA", Format(Now, "yyyy-MM-dd"), MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
TB.Buttons(3).Enabled = False
Exit Sub

err:
MsgBox err.Description
TB.Buttons(3).Enabled = False
MHVDB.RollbackTrans


End Sub


Private Sub txtformula_DblClick()
If Len(CBOTBL.Text) = 0 Then Exit Sub
If Frame5.Visible = True Then
Frame5.Visible = False
Else
Frame5.Visible = True
Frame5.Caption = UCase(CBOTBL.Text)
fillfield
End If
End Sub

Private Sub fillfield()
Dim fcount, j As Integer
Dim rs As New ADODB.Recordset
Dim Fstring As String
Set rs = Nothing
Fstring = ""
DZLIST.Clear
rs.Open "select * from tbltable where tblid='" & CBOTBL.BoundText & "'", ODKDB

fcount = IIf(IsNull(rs!fieldcount), 0, rs!fieldcount) + 1

Set rs = Nothing
rs.Open "SELECT * FROM " & LCase(CBOTBL.Text) & "", ODKDB
For j = 0 To fcount - 1
'Fstring = rs.Fields(j).Name & "," & Fstring
DZLIST.AddItem Trim(rs.Fields(j).Name)
Next



End Sub

Private Sub txtformula_KeyPress(KeyAscii As Integer)
If InStr(1, "/*-+sum(){}[]0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
