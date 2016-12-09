VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMTEMPRATURERECORDING 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T E M P R E T U R E   R E C O R D I N G  . . . "
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16020
   Icon            =   "FRMTEMPRATURERECORDING.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Sort By"
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
      Left            =   13080
      TabIndex        =   28
      Top             =   2160
      Width           =   2895
      Begin VB.CommandButton Command2 
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
         Left            =   2040
         Picture         =   "FRMTEMPRATURERECORDING.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81657857
         CurrentDate     =   41479
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81657857
         CurrentDate     =   41479
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         TabIndex        =   32
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   15975
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   15855
         _cx             =   27966
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12640511
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
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRMTEMPRATURERECORDING.frx":11CC
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
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      Begin VB.TextBox txtcomments 
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
         Left            =   2400
         TabIndex        =   26
         Top             =   2040
         Width           =   7455
      End
      Begin VB.ComboBox cboupordown 
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
         ItemData        =   "FRMTEMPRATURERECORDING.frx":1378
         Left            =   5400
         List            =   "FRMTEMPRATURERECORDING.frx":1382
         TabIndex        =   25
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txttherm2rh 
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
         Left            =   2400
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txttherm2degree 
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
         Left            =   8400
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txttherm1rh 
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
         Left            =   5400
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txttherm1degree 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81657857
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMTEMPRATURERECORDING.frx":1390
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSDataListLib.DataCombo cboshorttime 
         Bindings        =   "FRMTEMPRATURERECORDING.frx":13A5
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   7080
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSDataListLib.DataCombo cbofacilityid 
         Bindings        =   "FRMTEMPRATURERECORDING.frx":13BA
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2400
         TabIndex        =   11
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
      Begin MSDataListLib.DataCombo cbostaffid 
         Bindings        =   "FRMTEMPRATURERECORDING.frx":13CF
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   7080
         TabIndex        =   20
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSDataListLib.DataCombo cbocloudcover 
         Bindings        =   "FRMTEMPRATURERECORDING.frx":13E4
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   8400
         TabIndex        =   27
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
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
         Left            =   6480
         TabIndex        =   18
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Plastic Up or Down"
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
         Left            =   3360
         TabIndex        =   17
         Top             =   1680
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Shade Closed %"
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Therm 2  RH"
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
         TabIndex        =   15
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Therm 2 Degree C"
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
         Left            =   6480
         TabIndex        =   14
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Therm 1 RH"
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
         Left            =   3360
         TabIndex        =   13
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Therm 1 Degree C"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Facility"
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
         TabIndex        =   10
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transaction  Id"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1560
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
         Left            =   4440
         TabIndex        =   4
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Time"
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
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   540
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
            Picture         =   "FRMTEMPRATURERECORDING.frx":13F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTEMPRATURERECORDING.frx":1793
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTEMPRATURERECORDING.frx":1B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTEMPRATURERECORDING.frx":2807
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTEMPRATURERECORDING.frx":2C59
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTEMPRATURERECORDING.frx":3413
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   16020
      _ExtentX        =   28258
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
End
Attribute VB_Name = "FRMTEMPRATURERECORDING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbotrnid_LostFocus()
On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmstempreturerecording where trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
    txtentrydate.Value = Format(rs!entrydate, "yyyy-MM-dd")
    FindqmsTime rs!shortTime
    cboshorttime.Text = qmsTime
     FindqmsShade rs!shadeclosed
    cbocloudcover.Text = rs!shadeclosed & "  " & qmsShade
    FindsTAFF rs!staffid
    cbostaffid.Text = rs!staffid & "  " & sTAFF
    txtcomments.Text = rs!Comments
    cboupordown.Text = rs!plasticposition
    txttherm1degree.Text = rs!therm1degree
    txttherm2degree.Text = rs!therm2degree
    txttherm1rh.Text = rs!therm1rh
    txttherm2rh.Text = rs!therm2rh
    
   
   Else
   MsgBox "Record Not Found."
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
End Sub

Private Sub Command2_Click()
FillGrid txtfrmdate.Value, txttodate.Value
End Sub

Private Sub Form_Load()
On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnid  from tblqmstempreturerecording order by trnid", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "trnid"
cbotrnid.BoundColumn = "trnid"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(facilityId , '  ', description) as description,facilityId  from tblqmsfacility order by facilityId", db
Set cbofacilityid.RowSource = rsF
cbofacilityid.ListField = "description"
cbofacilityid.BoundColumn = "facilityId"



Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select  DATE_FORMAT(fulltime, '%H:%i:%s') as description,shorttime  from tblqmsshorttime order by shorttime", db
Set cboshorttime.RowSource = rsF
cboshorttime.ListField = "description"
cboshorttime.BoundColumn = "shorttime"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select  description,id  from tblqmsshade order by id", db
Set cbocloudcover.RowSource = rsF
cbocloudcover.ListField = "description"
cbocloudcover.BoundColumn = "id"



Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
If rsF.State = adStateOpen Then Srs.Close
rsF.Open "select concat(STAFFCODE , '  ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE nursery='1'  order by STAFFCODE", db
Set cbostaffid.RowSource = rsF
cbostaffid.ListField = "STAFFNAME"
cbostaffid.BoundColumn = "STAFFCODE"

'FillGrid
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
  txttodate.Value = Format(Now, "dd/MM/yyyy")

Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err
Select Case Button.Key

       Case "ADD"
            
        cbotrnid.Enabled = False
        TB.Buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmstempreturerecording", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbotrnid.Text = IIf(IsNull(rs!MaxId), 1, rs!MaxId)
        Else
        cbotrnid.Text = rs!MaxId
        End If
       Case "OPEN"
        Operation = "OPEN"
        CLEARCONTROLL
        cbotrnid.Enabled = True
        TB.Buttons(3).Enabled = True
             
       Case "SAVE"
        MNU_SAVE
      
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
Dim rs As New ADODB.Recordset
On Error GoTo err
If Len(cbotrnid.Text) = 0 Then
MsgBox "Transaction No. is Must."
Exit Sub
End If



MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmstempreturerecording (trnid,entrydate,shorttime,facilityid,therm1degree," _
            & "therm1rh,therm2degree,therm2rh,shadeclosed,plasticposition,comments,staffid,status,location)" _
            & "values(" _
            & "'" & cbotrnid.BoundText & "'," _
            & "'" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
            & "'" & cboshorttime.BoundText & "'," _
            & "'" & cbofacilityid.BoundText & "'," _
            & "'" & Val(txttherm1degree.Text) & "'," _
            & "'" & Val(txttherm1rh.Text) & "'," _
            & "'" & Val(txttherm2degree.Text) & "'," _
            & "'" & Val(txttherm2rh.Text) & "'," _
            & "'" & cbocloudcover.BoundText & "'," _
            & "'" & cboupordown.Text & "'," _
            & "'" & txtcomments.Text & "'," _
            & "'" & cbostaffid.BoundText & "'," _
            & "'ON'," _
            & "'" & Mlocation & "'" _
            & ")"
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & Mlocation & ","
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmstempreturerecording set " _
            & "entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "'," _
            & "shorttime='" & cboshorttime.BoundText & "'," _
            & "facilityid='" & cbofacilityid.BoundText & "'," _
            & "therm1degree='" & Val(txttherm1degree.Text) & "'," _
            & "therm1rh='" & Val(txttherm1rh.Text) & "'," _
            & "therm2degree='" & Val(txttherm2degree.Text) & "'," _
            & "therm2rh='" & Val(txttherm2rh.Text) & "'," _
            & "therm2rh='" & Val(txttherm2rh.Text) & "'," _
            & "plasticposition='" & cboupordown.Text & "'," _
            & "comments='" & txtcomments.Text & "'," _
            & "staffid='" & cbostaffid.BoundText & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and location='" & Mlocation & "'"
            

LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If
 TB.Buttons(3).Enabled = False
MHVDB.CommitTrans
FillGrid txtfrmdate.Value, txttodate.Value
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans


End Sub
Private Sub FillGrid(frmdate As Date, todate As Date)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid.Clear
mygrid.Rows = 1
mygrid.FormatString = "^Sl.No.|^Trn. Id|^Date|^Time|^Facility|^Therm1 Deg. C|^Therm1 RH|^Therm2 Deg. C|^Therm2 RH|^Plastic Position|^Shade Closed %|^Staff|^Comments|^"
mygrid.ColWidth(0) = 645
mygrid.ColWidth(1) = 975
mygrid.ColWidth(2) = 1200
mygrid.ColWidth(3) = 915
mygrid.ColWidth(4) = 1110
mygrid.ColWidth(5) = 1200
mygrid.ColWidth(6) = 1275
mygrid.ColWidth(7) = 1455
mygrid.ColWidth(8) = 1395
mygrid.ColWidth(9) = 1395
mygrid.ColWidth(10) = 1440
mygrid.ColWidth(11) = 1605
mygrid.ColWidth(12) = 960
mygrid.ColWidth(13) = 135
rs.Open "select * from tblqmstempreturerecording where entrydate>='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by trnid", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid.Rows = mygrid.Rows + 1
mygrid.TextMatrix(i, 0) = i

mygrid.TextMatrix(i, 1) = rs!trnid
mygrid.TextMatrix(i, 2) = Format(rs!entrydate, "dd/MM/yyyy")
FindqmsTime rs!shortTime
mygrid.TextMatrix(i, 3) = qmsTime
mygrid.TextMatrix(i, 4) = rs!facilityid
mygrid.TextMatrix(i, 5) = rs!therm1degree
mygrid.TextMatrix(i, 6) = rs!therm1rh
mygrid.TextMatrix(i, 7) = rs!therm2degree
mygrid.TextMatrix(i, 8) = rs!therm2rh
mygrid.TextMatrix(i, 9) = rs!plasticposition
mygrid.TextMatrix(i, 10) = rs!shadeclosed
mygrid.TextMatrix(i, 11) = rs!staffid
mygrid.TextMatrix(i, 12) = rs!Comments


rs.MoveNext
i = i + 1
Loop

rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub CLEARCONTROLL()
  txtentrydate.Value = Format(Now, "dd/MM/yyyy")
  txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
  txttodate.Value = Format(Now, "dd/MM/yyyy")
  cboshorttime.Text = ""
  txttherm1degree.Text = ""
  txttherm1rh.Text = ""
  txttherm2degree.Text = ""
  txttherm2rh.Text = ""
  cbofacilityid.Text = ""
  cbocloudcover.Text = ""
  cboupordown.Text = ""
  cbostaffid.Text = ""
  txtcomments.Text = ""
End Sub


