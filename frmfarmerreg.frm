VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmfarmerreg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F A R M E R    R E G I S T E R A T I O N"
   ClientHeight    =   8430
   ClientLeft      =   3015
   ClientTop       =   1710
   ClientWidth     =   15345
   Icon            =   "frmfarmerreg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15345
   Begin VB.Frame Frame7 
      Caption         =   "Search Farmer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   5040
      TabIndex        =   65
      Top             =   3960
      Visible         =   0   'False
      Width           =   9495
      Begin VB.Frame Frame8 
         Caption         =   "Search By."
         Height          =   735
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   9135
         Begin VB.OptionButton optphone 
            Caption         =   "Phone No."
            Height          =   375
            Left            =   7680
            TabIndex        =   76
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opthno 
            Caption         =   "H.No."
            Height          =   375
            Left            =   5160
            TabIndex        =   75
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opttno 
            Caption         =   "T.No."
            Height          =   375
            Left            =   3720
            TabIndex        =   74
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optcid 
            Caption         =   "CID"
            Height          =   375
            Left            =   6480
            TabIndex        =   73
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optfname 
            Caption         =   "Farmer Name"
            Height          =   375
            Left            =   1920
            TabIndex        =   72
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optfcode 
            Caption         =   "Farmer Code"
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox TXTSEARCHID 
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
         Left            =   2880
         TabIndex        =   68
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
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
         Height          =   495
         Left            =   4200
         Picture         =   "frmfarmerreg.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4320
         Width           =   1215
      End
      Begin VSFlex7Ctl.VSFlexGrid fgrid 
         Height          =   2655
         Left            =   120
         TabIndex        =   66
         Top             =   1560
         Width           =   9135
         _cx             =   16113
         _cy             =   4683
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
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmfarmerreg.frx":11CC
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
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   615
      Left            =   13200
      TabIndex        =   58
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New Land Reg..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13080
      Picture         =   "frmfarmerreg.frx":12B1
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   14160
      Picture         =   "frmfarmerreg.frx":1963
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   14160
      Picture         =   "frmfarmerreg.frx":1CED
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3000
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "CONTRACT INFORMATION"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   6015
      Begin VB.CommandButton Command5 
         Height          =   360
         Left            =   4800
         Picture         =   "frmfarmerreg.frx":239F
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox CHKISCONTRACTSIGNED 
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
         Left            =   3360
         TabIndex        =   30
         Top             =   360
         Width           =   495
      End
      Begin MSComCtl2.DTPicker TXTCONTRACTDATE 
         Height          =   375
         Left            =   3360
         TabIndex        =   31
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102039553
         CurrentDate     =   41208
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "IS CONTRACT SIGNED"
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
         Width           =   2025
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "CONTRACT DATE"
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
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
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
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   7200
      Width           =   12975
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   12735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "LAND INFORMATION"
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
      Height          =   1335
      Left            =   6960
      TabIndex        =   23
      Top             =   5520
      Width           =   6255
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         Picture         =   "frmfarmerreg.frx":2729
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "POPS UP LAND DETAILS"
         Top             =   720
         Width           =   610
      End
      Begin VB.TextBox txtregland 
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
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox txttotalarea 
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
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL REGISTERED LAND ACRE"
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
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   2985
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL FALLOW LAND ACRE"
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
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FARMER INFORMATION"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   13095
      Begin VB.CommandButton Command9 
         Height          =   375
         Left            =   12240
         Picture         =   "frmfarmerreg.frx":2FF3
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton CMDINF 
         Enabled         =   0   'False
         Height          =   495
         Left            =   12480
         Picture         =   "frmfarmerreg.frx":379D
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox CHKINF 
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
         Left            =   12240
         TabIndex        =   56
         Top             =   2640
         Width           =   255
      End
      Begin VB.Frame Frame6 
         Caption         =   "SORT BY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   6360
         TabIndex        =   51
         Top             =   120
         Width           =   2295
         Begin VB.OptionButton OPTNAME 
            Caption         =   "NAME"
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
            Left            =   960
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OPTBYID 
            Caption         =   " ID"
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
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.TextBox txtlocation 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   43
         Top             =   1560
         Width           =   3975
      End
      Begin VB.CommandButton cmdnext 
         Enabled         =   0   'False
         Height          =   495
         Left            =   9480
         Picture         =   "frmfarmerreg.frx":3F47
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkcaretaker 
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
         Left            =   9000
         TabIndex        =   39
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkisfarmercg 
         Caption         =   "IS FARMER CG"
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
         Left            =   7200
         TabIndex        =   37
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txthouseno 
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
         Height          =   375
         Left            =   9000
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtphone2 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtphone1 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtvillage 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtcid 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtfarmername 
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
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.ComboBox cbosex 
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
         ItemData        =   "frmfarmerreg.frx":46F1
         Left            =   9000
         List            =   "frmfarmerreg.frx":46FB
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   6000
         Picture         =   "frmfarmerreg.frx":470D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "REFRESH"
         Top             =   240
         Width           =   375
      End
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "frmfarmerreg.frx":4EB7
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo cbocg 
         Bindings        =   "frmfarmerreg.frx":4ECC
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   9000
         TabIndex        =   38
         Top             =   2040
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo cbomonitor 
         Bindings        =   "frmfarmerreg.frx":4EE1
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   9000
         TabIndex        =   61
         Top             =   3000
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
      Begin MSComCtl2.DTPicker txtregdate 
         Height          =   375
         Left            =   2040
         TabIndex        =   77
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102039553
         CurrentDate     =   41429
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "REG. DATE"
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
         TabIndex        =   78
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label Label22 
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
         Left            =   7440
         TabIndex        =   62
         Top             =   3120
         Width           =   885
      End
      Begin VB.Label LBLSTATUS 
         AutoSize        =   -1  'True
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
         Left            =   9840
         TabIndex        =   60
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label21 
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
         Height          =   255
         Left            =   9000
         TabIndex        =   59
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "INFLUENTIAL PERSON"
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
         Left            =   10200
         TabIndex        =   55
         Top             =   2640
         Width           =   2040
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "LOCATION NAME"
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
         Left            =   7200
         TabIndex        =   42
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CARETAKER"
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
         Left            =   7200
         TabIndex        =   40
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "HOUSE NO."
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
         Left            =   7200
         TabIndex        =   22
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "SEX"
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
         Left            =   7200
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PHONE-2"
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
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "PHONE-1"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "VILLAGE"
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
         TabIndex        =   18
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CID"
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
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FARMER NAME"
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
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FARMER ID"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADMINISTRATIVE LOCATION"
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   13095
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   855
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
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "frmfarmerreg.frx":4EF6
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2280
         TabIndex        =   1
         Top             =   360
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "frmfarmerreg.frx":4F0B
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   6000
         TabIndex        =   2
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSDataListLib.DataCombo CBOTSHOWOG 
         Bindings        =   "frmfarmerreg.frx":4F20
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   10080
         TabIndex        =   36
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TSHOWOG"
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
         Left            =   9000
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "GEWOG"
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
         Left            =   5160
         TabIndex        =   4
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DZONGKHAG"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   2640
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":4F35
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":52CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":5669
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":6343
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":6795
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":6F4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfarmerreg.frx":72E9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   1164
      ButtonWidth     =   1217
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PRINT"
            Key             =   "PRINT"
            Object.ToolTipText     =   "PRINTS THE DISPLAYED INFORMATION OF ABSENTEE."
            ImageIndex      =   7
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   3
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label19 
      Caption         =   "PICTURE WITH PIXEL/DIMENSION 600 X 600 IS DESIRABLE."
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
      Height          =   975
      Left            =   13560
      TabIndex        =   48
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "01/01/1900 IS EQUIVALENT TO NO DATE ASSIGNED."
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
      Left            =   3480
      TabIndex        =   47
      Top             =   6960
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Image ImgPic 
      BorderStyle     =   1  'Fixed Single
      Height          =   1890
      Left            =   13320
      Top             =   960
      Width           =   2000
   End
End
Attribute VB_Name = "frmfarmerreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rsDz As New ADODB.Recordset
Dim rsGe As New ADODB.Recordset
Dim rsfr As New ADODB.Recordset
Dim rsTs As New ADODB.Recordset
Dim rsCg As New ADODB.Recordset
Dim Srs As New ADODB.Recordset
Dim CrName As String
Dim AdmLoc As String
Dim picfile As String
Private Sub cboDzongkhag_GotFocus()
cbogewog.Enabled = True
chkcf.Enabled = False
chkgrf.Enabled = False
End Sub
Private Sub cbodzongkhag_LostFocus()
On Error GoTo err
cbogewog.Text = ""
cboDzongkhag.Enabled = True
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsGe = Nothing
If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cboDzongkhag.BoundText & "' order by dzongkhagid,gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub cbofarmerid_LostFocus()
'On Error GoTo err

Dim mystream As ADODB.Stream




cbofarmerid.Enabled = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindDZ Mid(rs!idfarmer, 1, 3)
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
cboDzongkhag.Text = Mid(rs!idfarmer, 1, 3) & " " & Dzname
cbogewog.Text = Mid(rs!idfarmer, 4, 3) & " " & GEname
CBOTSHOWOG.Text = Mid(rs!idfarmer, 7, 3) & " " & TsName
txtfarmername.Text = rs!farmername
txtcid.Text = IIf(IsNull(rs!cidno), "", rs!cidno)
'If RS!sex = 0 Then
cbosex.Text = cbosex.List(rs!sex)
'ElseIf RS!sex = 1 Then
'cbosex.Text = "Female"
'End If
chkisfarmercg.Value = IIf(IsNull(rs!ISfarmercg), 0, rs!ISfarmercg)

txthouseno.Text = IIf(IsNull(rs!houseno), "", rs!houseno)
txtvillage.Text = IIf(IsNull(rs!VILLAGE), "", rs!VILLAGE)
txtlocation.Text = IIf(IsNull(rs!LocationName), "", rs!LocationName)
txtphone1.Text = IIf(IsNull(rs!phone1), "", rs!phone1)
txtphone2.Text = IIf(IsNull(rs!phone2), "", rs!phone2)
CHKISCONTRACTSIGNED.Value = IIf(IsNull(rs!ISCONTRACTSIGNED), "", rs!ISCONTRACTSIGNED)
TXTCONTRACTDATE.Value = IIf(IsNull(rs!CONTRACTDATE), "1900-01-01", rs!CONTRACTDATE)
txttotalarea.Text = Format(IIf(IsNull(rs!TOTALAREA), 0, rs!TOTALAREA), "#####0.00")
txtregland.Text = Format(IIf(IsNull(rs!REGAREA), 0, rs!REGAREA), "####0.00")
chkcaretaker.Value = IIf(IsNull(rs!ISCARETAKER), 0, rs!ISCARETAKER)
txtregdate.Value = Format(IIf(IsNull(rs!regdate), "01/01/1900", rs!regdate), "dd/MM/yyyy")
Findstatus rs!status
LBLSTATUS.Caption = Mstatus
If rs!status = "A" Then
LBLSTATUS.ForeColor = vbBlue
disfield True
Else
LBLSTATUS.ForeColor = vbRed
disfield False
End If
FindsTAFF IIf(IsNull(rs!monitor), "", rs!monitor)
cbomonitor.Text = IIf(IsNull(rs!monitor), "", rs!monitor) & " " & sTAFF

'If chkcaretaker.Value = 1 Then
'cmdnext.Enabled = False
'Else
'cmdnext.Enabled = True
'End If
txtremarks.Text = IIf(IsNull(rs!remarks), "", rs!remarks)

Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB, adOpenStatic, adLockOptimistic
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
mystream.Open
If Len(rs!picfile) > 0 Then
mystream.Write rs!picfile
mystream.SaveToFile "c:\\" & cbofarmerid.BoundText & ".jpg", adSaveCreateOverWrite
mystream.Close
ImgPic.Picture = LoadPicture("c:\\" & cbofarmerid.BoundText & ".jpg")
ImgPic.Width = 2000
ImgPic.Height = 2000

 If Dir$("c:\\" & cbofarmerid.BoundText & ".jpg") <> vbNullString Then
    Kill "c:\\" & cbofarmerid.BoundText & ".jpg"
    End If

End If





Set rs = Nothing
 
 rs.Open "select sum(regland) as regland from tbllandreg where farmerid='" & cbofarmerid.BoundText & "'", MHVDB
 If rs.EOF <> True Then
 txtregland.Text = Format(IIf(IsNull(rs!regland), 0, rs!regland) + Val(txtregland.Text), "####0.00")
 End If
 
Else

MsgBox "Record Not Found."
End If
Dim rscr As New ADODB.Recordset
Set rscr = Nothing
rscr.Open "select * from tblabsentee where caretakerid='" & cbofarmerid.BoundText & "'", MHVDB
If rscr.EOF <> True Then
chkcaretaker.Enabled = False
cmdnext.Enabled = False
ISCARETAKER = True
Else
chkcaretaker.Enabled = True
ISCARETAKER = False
End If

Dim rschk As New ADODB.Recordset
Set rschk = Nothing
rschk.Open "SELECT * FROM tbllandreg WHERE FARMERID='" & cbofarmerid.BoundText & "' ", MHVDB
If rschk.EOF <> True Then
Command1.Enabled = True
mFARID = cbofarmerid.BoundText
Else
Command1.Enabled = False
mFARID = ""
End If



Exit Sub
'err:
'MsgBox err.Description
End Sub
Private Sub disfield(tt As Boolean)
Frame1.Enabled = tt
Frame2.Enabled = tt
Frame3.Enabled = tt
Frame4.Enabled = tt
Frame5.Enabled = tt
TB.Buttons(3).Enabled = tt
End Sub

Private Sub cbogewog_GotFocus()
CBOTSHOWOG.Enabled = True
End Sub

Private Sub cbogewog_LostFocus()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsTs = Nothing
cbogewog.Enabled = False
If rsTs.State = adStateOpen Then rsTs.Close
rsTs.Open "select concat(tshewogid , ' ', tshewogname) as tshewogname,tshewogid  from tbltshewog where dzongkhagid='" & cboDzongkhag.BoundText & "' and gewogid='" & cbogewog.BoundText & "' order by dzongkhagid,gewogid", db
Set CBOTSHOWOG.RowSource = rsTs
CBOTSHOWOG.ListField = "tshewogname"
CBOTSHOWOG.BoundColumn = "tshewogid"
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub CBOTSHOWOG_LostFocus()
If chkgrf.Value = 1 Then
GRFcodemax
ElseIf chkcf.Value = 1 Then
CFcodemax
Else

Fcodemax
End If
End Sub
Private Sub Fcodemax()
On Error GoTo err
Dim id As String
CBOTSHOWOG.Enabled = False
id = 0
AdmLoc = ""
AdmLoc = cboDzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText
If Operation = "ADD" Then
Dim rs As New ADODB.Recordset
If Len(cbogewog.Text) = 0 Then
         MsgBox "Please,Select Gewog."
         cbogewog.SetFocus
        Exit Sub
        End If
        cbogewog.BackColor = vbWhite
        cbogewog.Enabled = False
Set rs = Nothing
rs.Open "select max(substring(idfarmer,11,4)+1) as MaxId from tblfarmer WHERE SUBSTRING(idfarmer,1,9)='" + AdmLoc + "'", MHVDB, adOpenForwardOnly, adLockOptimistic

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
cbofarmerid.Text = AdmLoc & "F" & id
Else
cbofarmerid.Text = AdmLoc + "F" + "0001"
End If
        
        
ElseIf Operation = "OPEN" Then

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub GRFcodemax()
On Error GoTo err
Dim id As String
CBOTSHOWOG.Enabled = False
id = 0
AdmLoc = ""
AdmLoc = cboDzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText
If Operation = "ADD" Then
Dim rs As New ADODB.Recordset
If Len(cbogewog.Text) = 0 Then
         MsgBox "Please,Select Gewog."
         cbogewog.SetFocus
        Exit Sub
        End If
        cbogewog.BackColor = vbWhite
        cbogewog.Enabled = False
Set rs = Nothing
rs.Open "select max(substring(idfarmer,11,4)+1) as MaxId from tblfarmer WHERE SUBSTRING(idfarmer,1,9)='" + AdmLoc + "'", MHVDB, adOpenForwardOnly, adLockOptimistic

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
cbofarmerid.Text = AdmLoc & "G" & id
Else
cbofarmerid.Text = AdmLoc + "G" & "0001"
End If
        
        
ElseIf Operation = "OPEN" Then

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub CFcodemax()
On Error GoTo err
Dim id As String
CBOTSHOWOG.Enabled = False
id = 0
AdmLoc = ""
AdmLoc = cboDzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText
If Operation = "ADD" Then
Dim rs As New ADODB.Recordset
If Len(cbogewog.Text) = 0 Then
         MsgBox "Please,Select Gewog."
         cbogewog.SetFocus
        Exit Sub
        End If
        cbogewog.BackColor = vbWhite
        cbogewog.Enabled = False
Set rs = Nothing
rs.Open "select max(substring(idfarmer,11,4)+1) as MaxId from tblfarmer WHERE SUBSTRING(idfarmer,1,9)='" + AdmLoc + "'", MHVDB, adOpenForwardOnly, adLockOptimistic

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
cbofarmerid.Text = AdmLoc & "C" & id
Else
cbofarmerid.Text = AdmLoc + "C" & "0001"
End If
        
        
ElseIf Operation = "OPEN" Then

Else

End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub chkcaretaker_Click()
If chkcaretaker.Value = 1 Then
CHKISCONTRACTSIGNED.Enabled = False
TXTCONTRACTDATE.Enabled = False
txttotalarea.Enabled = False
txtregland.Enabled = False
cmdnext.Enabled = True
CHKINF.Value = 0
CHKINF.Enabled = False
CMDINF.Enabled = False
ISCARETAKER = True


Else
CHKISCONTRACTSIGNED.Enabled = True
TXTCONTRACTDATE.Enabled = True
txttotalarea.Enabled = True
txtregland.Enabled = True
cmdnext.Enabled = False
CHKINF.Enabled = True
ISCARETAKER = False
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

Private Sub CHKINF_Click()
If CHKINF.Value = 1 Then
CHKISCONTRACTSIGNED.Enabled = False
TXTCONTRACTDATE.Enabled = False
txttotalarea.Enabled = False
txtregland.Enabled = False
CMDINF.Enabled = True
chkcaretaker.Value = 0
chkcaretaker.Enabled = False
cmdnext.Enabled = False



Else
CHKISCONTRACTSIGNED.Enabled = True
TXTCONTRACTDATE.Enabled = True
txttotalarea.Enabled = True
txtregland.Enabled = True
cmdnext.Enabled = False

chkcaretaker.Enabled = True

End If
End Sub

Private Sub chkisfarmercg_Click()
If chkisfarmercg.Value = 1 Then

cbocg.Enabled = True
Else

cbocg.Enabled = False

End If
End Sub

Private Sub CMDINF_Click()
If MsgBox("Do You Want To Save The Influential Person Record", vbYesNo) = vbYes Then


If Len(cbofarmerid.BoundText) = 0 Then
MsgBox "Please Check The Entries In The Farmer Registration."
Exit Sub
Else

MNU_SAVE
mbypass = True
Mcaretaker = cbofarmerid.BoundText
FATYPEINF = "F"
frminf.Show 1

End If




Else
cmdnext.Enabled = False
End If
mbypass = False
End Sub

Private Sub cmdnext_Click()
If MsgBox("Do You Want To Save The Caretaker Record", vbYesNo) = vbYes Then
If Len(cbofarmerid.BoundText) = 0 Then
MsgBox "Please Check The Entries In The Farmer Registration."
Exit Sub
Else

MNU_SAVE
mbypass = True
Mcaretaker = cbofarmerid.BoundText
FRMABSENTEE.Show 1
End If


Else
cmdnext.Enabled = False
End If
mbypass = False
End Sub

Private Sub Command1_Click()

FRMLANDDETAILS.Show 1
End Sub

Private Sub Command2_Click()
On Error GoTo ErrHandler
picfile = ""
    CD.CancelError = True
    CD.InitDir = "C:\"
    'CD.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
    CD.Filter = "All Files (*.*)|*.*"
    CD.ShowOpen
    PicPath = Mid(CD.FileName, 53, 4000)
    imagenm = PicPath
    picfile = CD.FileName
    ImgPic.Picture = LoadPicture(CD.FileName)
    ImgPic.Width = 2000
    ImgPic.Height = 2000
    
Exit Sub

ErrHandler:
'    User pressed Cancel button.
   Exit Sub
End Sub

Private Sub Command4_Click()
Label19.Visible = True
End Sub

Private Sub Command5_Click()
Label18.Visible = True
End Sub

Private Sub Command6_Click()

If TB.Buttons(3).Enabled = True Then
MsgBox "Please Save This Information First."
mbypass = False

Exit Sub
Else
mFARID = cbofarmerid.BoundText
Unload Me
mbypass = True

FRMNEWLANDREG.Show 1
End If

End Sub
Private Sub fillemp()

End Sub

Private Sub Command7_Click()
'Dim rs As New ADODB.Recordset
'Dim rsS As New ADODB.Recordset
'Dim excel_app As Object
'Dim excel_sheet As Object
'Dim i As Integer
'Dim pnastr As String
'Dim DEPTCODE As String
'pnastr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=mismukti;Initial Catalog=PNA;Data Source=NIRVANA"
'Dim Pnadb As New ADODB.Connection
'Pnadb.CursorLocation = adUseClient
'Pnadb.Open pnastr
'
'DEPTCODE = ""
'
'Screen.MousePointer = vbHourglass
'    DoEvents
'    Set excel_app = CreateObject("Excel.Application")
'    Set Excel_WBook = excel_app.Workbooks.Add
'    If Val(excel_app.Application.Version) >= 8 Then
'        Set excel_sheet = excel_app.ActiveSheet
'    Else
'        Set excel_sheet = excel_app
'    End If
'    'excel_app.Caption = "MHV"
'    Dim sl As Integer
'    sl = 1
'    'excel_app.DisplayFullScreen = True
'    excel_app.Visible = True
'    excel_sheet.Cells(3, 1) = "EMPN"
'    excel_sheet.Cells(3, 2) = "NAME"
'    excel_sheet.Cells(3, 3) = "DESIGNATION"
'     excel_sheet.Cells(3, 4) = "GRADE"
'    excel_sheet.Cells(3, 5) = "INSURED AMOUNT"
'    'excel_sheet.Cells(3, 6) = "INSURED AMOUNT"
'        i = 4
'  Set rs = Nothing
'rs.Open "select empn,a.name AS ENAME,b.name AS DNAME,A.DEPT AS DCODE,grade, a.dept,DESIG from paymas a,dept b where a.dept=b.dept and type='REG'and emp_stat is null and grade<>99 order by a.dept", Pnadb
'
'   Do While rs.EOF <> True
'
'   If DEPTCODE <> rs!DCODE Then
'
'
'     excel_sheet.Cells(i, 2) = rs!DNAME
'     excel_sheet.Cells(i, 2).Font.Bold = True
'       i = i + 1
'
'          excel_sheet.Cells(i, 1) = rs!EMPN
'          excel_sheet.Cells(i, 2) = rs!ENAME
'          sl = sl + 1
'       excel_sheet.Cells(i, 3) = rs!DESIG
'        excel_sheet.Cells(i, 4) = rs!GRADE
'        excel_sheet.Cells(i, 5) = ""
'        'DT_OF_AP
'
'     Else
'     excel_sheet.Cells(i, 1) = rs!EMPN
'     excel_sheet.Cells(i, 2) = rs!ENAME
'   excel_sheet.Cells(i, 3) = rs!DESIG
'excel_sheet.Cells(i, 4) = rs!GRADE
'        excel_sheet.Cells(i, 5) = ""
'
'sl = sl + 1
'End If
'DEPTCODE = rs!DCODE
'   i = i + 1
'
'   rs.MoveNext
'   Loop


End Sub

Private Sub Command8_Click()
' Set connRemote = New ADODB.Connection
'  connRemote.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
'                        & "SERVER=server2.mountainhazelnuts.com;" _
'                        & " DATABASE=odk_prod;" _
'                        & "UID=odk_user;PWD=none; OPTION=3"
'connRemote.Open ' remote connection
'
'
'Dim RsRemote As New ADODB.Recordset
'Set RsRemote = Nothing
'RsRemote.Open "select _URI as URI,_CREATOR_URI_USER as CREATOR_URI_USER,_CREATION_DATE as CREATION_DATE from ADVOCATEDALIYACT12_CORE", connRemote
'
'' remote connection and record selection ends here
'
'Set CONNLOCAL = New ADODB.Connection
'Dim RsLocal As New ADODB.Recordset
'Set RsLocal = Nothing
'CONNLOCAL.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=bb;Initial Catalog=odk_prodLocal" ' local connection
'Do While RsRemote.EOF <> True ' inserts remmote records into local
'CONNLOCAL.Execute "insert into ADVOCATEDALIYACT12_CORE(_URI,_CREATOR_URI_USER,_CREATION_DATE)values('" & RsRemote!uri & "','" & RsRemote!CREATOR_URI_USER & "','" & RsRemote!CREATION_DATE & "') "
'RsRemote.MoveNext
'Loop
'
'MsgBox "data transfer from colud is completed."

Frame7.Visible = False


End Sub

Private Sub Command9_Click()
Frame7.Visible = True
End Sub

Private Sub Form_Load()
On Error GoTo err

Operation = ""
TXTCONTRACTDATE.Value = Format(Now, "dd/MM/yyyy")

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cboDzongkhag.RowSource = rsDz
cboDzongkhag.ListField = "dzongkhagname"
cboDzongkhag.BoundColumn = "dzongkhagcode"

If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog order by dzongkhagid,gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"



If rsTs.State = adStateOpen Then rsTs.Close
rsTs.Open "select concat(tshewogid , ' ', tshewogname) as tshewogname,tshewogid  from tbltshewog order by dzongkhagid,gewogid", db
Set CBOTSHOWOG.RowSource = rsTs
CBOTSHOWOG.ListField = "tshewogname"
CBOTSHOWOG.BoundColumn = "tshewogid"

Set rsfr = Nothing

If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer where status='A' order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Set rsCg = Nothing

If rsCg.State = adStateOpen Then rsCg.Close
rsCg.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbocg.RowSource = rsCg
cbocg.ListField = "farmername"
cbocg.BoundColumn = "idfarmer"



Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE moniter='1'  order by STAFFCODE", db
Set cbomonitor.RowSource = Srs
cbomonitor.ListField = "STAFFNAME"
cbomonitor.BoundColumn = "STAFFCODE"




If mbypass = True Then
cboDzongkhag.Enabled = False
       TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       cboDzongkhag.Enabled = True
     cboabsenteeid.Enabled = False
       
       
       
       
       
      
CBOCARETAKER.Text = Mcaretaker
CBOCARETAKER.Enabled = False
Else


End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub Label18_Click()
Label18.Visible = False
End Sub

Private Sub Label19_Click()
Label19.Visible = False
End Sub

Private Sub OPTBYID_Click()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing


cbofarmerid.Text = ""


If OPTNAME.Value = True Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(farmername  , ' ',idfarmer) as farmername,idfarmer  from tblfarmer order by FARMERNAME", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Else
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub OPTNAME_Click()
On Error GoTo err
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing

cbofarmerid.Text = ""



If OPTNAME.Value = True Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(farmername  , ' ',idfarmer) as farmername,idfarmer  from tblfarmer order by FARMERNAME", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

Else
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
End If
Exit Sub
err:
MsgBox err.Description
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ADD"
AdmLoc = ""
chkcf.Enabled = True
chkgrf.Enabled = True
       cboDzongkhag.Enabled = False
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
         cboDzongkhag.Enabled = True
         cbofarmerid.Enabled = False
         CBOTSHOWOG.Enabled = False
         Case "OPEN"
         Operation = "OPEN"
         CLEARCONTROLL
         cbofarmerid.Enabled = True
         cboDzongkhag.Enabled = False
         cbogewog.Enabled = False
         CBOTSHOWOG.Enabled = False
         TB.Buttons(3).Enabled = True
       
       Case "SAVE"
      
       MNU_SAVE
        
       Case "DELETE"
         Case "PRINT"
         PRINTFINFO
          TB.Buttons(6).Enabled = False
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub PRINTFINFO()

'On Error Resume Next
Dim excel_app As Object
Dim excel_sheet As Object
Dim row As Long
Dim statement As String
Dim i, j, K As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Screen.MousePointer = vbHourglass

FileCopy App.Path + "\FARMERINFO.XLS", App.Path + "\" + cbofarmerid.Text + Format(Now, "ddMMyyyy") + ".XLS"
Set excel_app = CreateObject("Excel.Application")
excel_app.Workbooks.Open FileName:=App.Path + "\" + cbofarmerid.Text + Format(Now, "ddMMyyyy") + ".XLS"
If Val(excel_app.Application.Version) >= 8 Then
   Set excel_sheet = excel_app.ActiveSheet
Else
   Set excel_sheet = excel_app
End If
excel_app.Visible = True
Set rs = Nothing
rs.Open "SELECT * FROM tblfarmer WHERE IDFARMER='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindDZ Mid(rs!idfarmer, 1, 3)
excel_sheet.Cells(5, 2) = Mid(rs!idfarmer, 1, 3) & " " & Dzname 'cboDzongkhag.Text
FindGE Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3)
excel_sheet.Cells(5, 6) = Mid(rs!idfarmer, 1, 3) & Mid(rs!idfarmer, 4, 3) & " " & GEname
FindTs Mid(rs!idfarmer, 1, 3), Mid(rs!idfarmer, 4, 3), Mid(rs!idfarmer, 7, 3)
excel_sheet.Cells(6, 6) = Mid(rs!idfarmer, 1, 3) & Mid(rs!idfarmer, 4, 3) & Mid(rs!idfarmer, 7, 3) & " " & TsName
excel_sheet.Cells(7, 2) = rs!idfarmer & " " & rs!farmername
'excel_sheet.Cells(7, 6) = CBOCARETAKER.Text
excel_sheet.Cells(9, 2) = "'" & rs!cidno
If rs!sex = 0 Then
excel_sheet.Cells(9, 6) = "MALE" '
Else
excel_sheet.Cells(9, 6) = "FEMALE"
End If

excel_sheet.Cells(10, 2) = rs!houseno
excel_sheet.Cells(11, 2) = rs!VILLAGE
excel_sheet.Cells(11, 6) = rs!LocationName

excel_sheet.Cells(12, 2) = rs!phone1
excel_sheet.Cells(12, 6) = rs!phone2
If ISfarmercg = 0 Then
excel_sheet.Cells(13, 2) = "NO"
Else
excel_sheet.Cells(13, 2) = "YES"
End If
If rs!ISCARETAKER = 0 Then
excel_sheet.Cells(13, 6) = "NO"
Else
excel_sheet.Cells(13, 6) = "YES"
End If
If rs!ISCONTRACTSIGNED = 1 Then
excel_sheet.Cells(14, 2) = "YES"
If Format(rs!CONTRACTDATE, "dd/MM/yyyy") = "01/01/1900" Then
excel_sheet.Cells(14, 6) = ""
Else
excel_sheet.Cells(14, 6) = "'" & Format(rs!CONTRACTDATE, "dd/MM/yyyy") & "  " & "(DD/MM/YYYY)"
End If
Else
excel_sheet.Cells(14, 2) = "NO"
excel_sheet.Cells(14, 6) = ""

End If
excel_sheet.Cells(15, 2) = "'" & Format(IIf(IsNull(rs!TOTALAREA), 0, rs!TOTALAREA), "#####0.00")
Set rs1 = Nothing
rs1.Open "SELECT SUM(REGLAND)AS REGLAND FROM tbllandreg GROUP BY FARMERID", MHVDB
excel_sheet.Cells(15, 6) = "'" & Format(IIf(IsNull(rs!REGAREA), 0, rs!REGAREA) + IIf(IsNull(rs1!regland), 0, rs1!regland), "#####0.00")
excel_sheet.Cells(18, 1) = rs!remarks
Else
MsgBox "Record Not Found."
End If


Set rs = Nothing


'With excel_app.ActiveSheet.Pictures.Insert(App.Path + "\image\" + cboabsenteeid.BoundText & ".jpg")
'    With .ShapeRange
'        .LockAspectRatio = msoTrue
'        .Width = 60
'        .Height = 60
'    End With
'    .Left = excel_app.ActiveSheet.Cells(1, 8).Left
'    .Top = excel_app.ActiveSheet.Cells(1, 8).Top
'    .Placement = 1
'    .PrintObject = True
'End With



Screen.MousePointer = vbDefault
End Sub

Private Sub MNU_SAVE()

On Error GoTo err
Dim mystream As ADODB.Stream
Dim rs As New ADODB.Recordset

Set rs = Nothing
rs.Open "select * from tblfarmer where cidno='" & Trim(txtcid.Text) & "'", MHVDB
If rs.EOF <> True And Operation = "ADD" Then
If MsgBox("This CID already exist, Do you want to proceed?", vbYesNo) = vbYes Then
Else
Exit Sub
End If
End If

Set rs = Nothing
rs.Open "select * from tblfarmer where HOUSENO='" & Trim(txthouseno.Text) & "'", MHVDB
If rs.EOF <> True And Operation = "ADD" Then
If MsgBox("This house no already exist, Do you want to proceed?", vbYesNo) = vbYes Then
Else
Exit Sub
End If
End If

Dim id As String

id = 0
AdmLoc = ""
AdmLoc = cboDzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText
If Operation = "ADD" Then


Set rs = Nothing
rs.Open "select max(substring(idfarmer,11,4)+1) as MaxId from tblfarmer WHERE SUBSTRING(idfarmer,1,9)='" + AdmLoc + "'", MHVDB, adOpenForwardOnly, adLockOptimistic

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
End If
End If






Dim msex As Integer
If chkgrf.Value = 0 Then
If cbosex.Text = "Male" Then
msex = 0
ElseIf cbosex.Text = "Female" Then
msex = 1
Else
MsgBox "Please select The appropriate Sex."
Exit Sub
End If
End If


If Len(cbofarmerid.Text) = 0 Then
MsgBox "Please Select The Appropriate Information From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If

If chkisfarmercg.Value = 1 And Len(cbocg.Text) = 0 Then
MsgBox "Please Select The CG"
cbocg.SetFocus
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
If chkcaretaker.Value = 0 Then

MHVDB.Execute "insert into tblfarmer(idfarmer,admlocation,farmername,cidno,SEX,houseno,village,locationname,phone1," _
& "phone2,iscontractsigned,contractdate,totalarea,regarea,remarks,isfarmercg,cgid,iscaretaker,monitor,status,regdate,insertedby,inserteddate)" _
& "values('" & cbofarmerid.Text & "','" & AdmLoc & "','" & txtfarmername.Text & "','" & txtcid.Text & "','" & msex & "'" _
& " ,'" & txthouseno.Text & "','" & txtvillage.Text & "','" & txtlocation.Text & "','" & txtphone1.Text & "','" & txtphone2.Text & "','" & CHKISCONTRACTSIGNED.Value & "','" & Format(TXTCONTRACTDATE.Value, "yyyy-MM-dd") & "'" _
& ",'" & txttotalarea.Text & "','" & txtregland.Text & "','" & txtremarks.Text & "','" & chkisfarmercg.Value & "','" & cbocg.BoundText & "','" & chkcaretaker.Value & "','" & cbomonitor.BoundText & "','A','" & Format(txtregdate.Value, "yyyy-MM-dd") & "','" & MUSER & "','" & Format(Now, "yyyy-MM-dd") & "')"

Else
MHVDB.Execute "insert into tblfarmer(idfarmer,admlocation,farmername,cidno,SEX,houseno,village,locationname,phone1," _
& "phone2,iscontractsigned,contractdate,remarks,isfarmercg,cgid,iscaretaker,monitor,status,regdate,insertedby,inserteddate)" _
& "values('" & cbofarmerid.Text & "','" & AdmLoc & "','" & txtfarmername.Text & "','" & txtcid.Text & "','" & msex & "'" _
& " ,'" & txthouseno.Text & "','" & txtvillage.Text & "','" & txtlocation.Text & "','" & txtphone1.Text & "','" & txtphone2.Text & "','" & CHKISCONTRACTSIGNED.Value & "','" & Format(TXTCONTRACTDATE.Value, "yyyy-MM-dd") & "'" _
& ",'" & txtremarks.Text & "','" & chkisfarmercg.Value & "','" & cbocg.BoundText & "','" & chkcaretaker.Value & "','" & cbomonitor.BoundText & "','A','" & Format(txtregdate.Value, "yyyy-MM-dd") & "','" & MUSER & "','" & Format(Now, "yyyy-MM-dd") & "')"

End If

If Len(picfile) > 0 Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB, adOpenStatic, adLockOptimistic

mystream.Open
mystream.LoadFromFile picfile
rs!picfile = mystream.Read
rs.Update
End If

ElseIf Operation = "OPEN" Then
If chkcaretaker.Value = 0 Then
MHVDB.Execute "update tblfarmer set farmername='" & txtfarmername.Text & "',cidno='" & txtcid.Text & "',SEX='" & msex & "',houseno='" & txthouseno.Text & "',village='" & txtvillage.Text & "',locationname='" & txtlocation.Text & "',phone1='" & txtphone1.Text & "'," _
& "phone2='" & txtphone2.Text & "',iscontractsigned='" & CHKISCONTRACTSIGNED.Value & "',contractdate='" & Format(TXTCONTRACTDATE.Value, "yyyy-MM-dd") & "',totalarea='" & txttotalarea.Text & "',remarks='" & txtremarks.Text & "',isfarmercg='" & chkisfarmercg.Value & "',cgid='" & cbocg.BoundText & "',regdate='" & Format(txtregdate.Value, "yyyy-MM-dd") & "',iscaretaker='" & chkcaretaker.Value & "',monitor='" & cbomonitor.BoundText & "',updatedby='" & MUSER & "',updateddate='" & Format(Now, "yyyy-MM-dd") & "'  where idfarmer='" & cbofarmerid.BoundText & "'"

Else
MHVDB.Execute "update tblfarmer set farmername='" & txtfarmername.Text & "',regdate='" & Format(txtregdate.Value, "yyyy-MM-dd") & "',cidno='" & txtcid.Text & "',SEX='" & msex & "',houseno='" & txthouseno.Text & "',village='" & txtvillage.Text & "',locationname='" & txtlocation.Text & "',phone1='" & txtphone1.Text & "'," _
& "phone2='" & txtphone2.Text & "',iscontractsigned='" & CHKISCONTRACTSIGNED.Value & "',contractdate='" & Format(TXTCONTRACTDATE.Value, "yyyy-MM-dd") & "',remarks='" & txtremarks.Text & "',isfarmercg='" & chkisfarmercg.Value & "',cgid='" & cbocg.BoundText & "',iscaretaker='" & chkcaretaker.Value & "',monitor='" & cbomonitor.BoundText & "' ,updatedby='" & MUSER & "',updateddate='" & Format(Now, "yyyy-MM-dd") & "'  where idfarmer='" & cbofarmerid.BoundText & "'"


End If
' Val(txtregland.Text)
If Len(picfile) > 0 Then
Set mystream = New ADODB.Stream
mystream.Type = adTypeBinary


Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & cbofarmerid.BoundText & "'", MHVDB, adOpenStatic, adLockOptimistic

mystream.Open
mystream.LoadFromFile picfile
rs!picfile = mystream.Read
rs.Update
End If

Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='6'"
MHVDB.CommitTrans
mbypass = False


TB.Buttons(3).Enabled = False
        'FillGrid
       TB.Buttons(6).Enabled = True
       Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub
Private Sub CLEARCONTROLL()
txtfarmername.Text = ""
cboDzongkhag.Text = ""
cbogewog.Text = ""
CBOTSHOWOG.Text = ""
cbofarmerid.Text = ""
txtcid.Text = ""
cbosex.Text = ""
txtphone1.Text = ""
cbomonitor.Text = ""
txtphone2.Text = ""
txtvillage.Text = ""
txtlocation.Text = ""
txthouseno.Text = ""
LBLSTATUS.Caption = ""
txttotalarea.Text = ""
txtregland.Text = ""
chkisfarmercg.Value = 0
txtregdate.Value = Format(Now, "dd-MM-yyyy")
cbocg.Text = ""
chkcaretaker.Value = 0
ImgPic.Picture = Nothing
Command1.Enabled = False
TXTCONTRACTDATE.Value = Format(Now, "dd/MM/yyyy")
End Sub

Private Sub txtcid_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtfarmername_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtfarmername.SelStart + 1
    Dim sText As String
    sText = Left$(txtfarmername.Text, iPos)
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

Private Sub txthouseno_Change()
txthouseno.Text = StrConv(txthouseno.Text, vbUpperCase)
txthouseno.SelStart = Len(txthouseno)
End Sub

Private Sub txtlocation_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtlocation.SelStart + 1
    Dim sText As String
    sText = Left$(txtlocation.Text, iPos)
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

Private Sub txtphone1_KeyPress(KeyAscii As Integer)
If InStr(1, "+0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtphone2_KeyPress(KeyAscii As Integer)
If InStr(1, "+0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtregland_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TXTSEARCHID_Change()
Dim SQLSTR As String
Dim i As Integer
'If Len(TXTSEARCHID.Text) <= 3 Then Exit Sub
Dim rs As New ADODB.Recordset
If TXTSEARCHID.Text = "'" Then
MsgBox (err.Number & " : " & "Enter Valid Character for Search.")
TXTSEARCHID.Text = ""
TXTSEARCHID.SetFocus
Exit Sub
End If




    If TXTSEARCHID.Text = "" Then
        cleargrid
        i = 1
  Exit Sub
  
  End If
        Set rs = Nothing
                
        If optfcode.Value = True Then
        SQLSTR = "select * from tblfarmer where idfarmer like '" & TXTSEARCHID.Text & "%' order by idfarmer"
        ElseIf optfname.Value = True Then
        SQLSTR = "select * from tblfarmer where farmername like '" & TXTSEARCHID.Text & "%' order by idfarmer"
        ElseIf optcid.Value = True Then
        SQLSTR = "select * from tblfarmer where cidno like '" & TXTSEARCHID.Text & "%' order by idfarmer"
        ElseIf opttno.Value = True Then
        SQLSTR = "select * from tblfarmer  where  idfarmer in(select farmerid from tbllandreg where  thramno like '" & TXTSEARCHID.Text & "%') order by idfarmer"
        ElseIf opthno.Value = True Then
        SQLSTR = "select * from tblfarmer where houseno like '" & TXTSEARCHID.Text & "%' order by idfarmer"
        ElseIf optphone.Value = True Then
        SQLSTR = "select * from tblfarmer where phone1 like '" & TXTSEARCHID.Text & "%' order by idfarmer"
        Else
        MsgBox "Search option not selected."
        Exit Sub
        
        End If
        
        
        
        'SQLSTR = "" '
        
        
        
        
        rs.Open SQLSTR, MHVDB
        If rs.RecordCount > 0 Then
        rs.MoveFirst
        Else
        On Error Resume Next
        End If
         cleargrid
         i = 1
        
        Do Until rs.EOF
        fgrid.Rows = fgrid.Rows + 1
       
         fgrid.TextMatrix(i, 0) = i
         fgrid.TextMatrix(i, 1) = rs!idfarmer & "  " & rs!farmername
         fgrid.ColAlignment(1) = flexAlignLeftTop
         
         fgrid.TextMatrix(i, 2) = rs!cidno
         If opttno.Value = True Then
         fgrid.TextMatrix(i, 3) = rs!thramno
         End If
         fgrid.TextMatrix(i, 4) = rs!houseno
         fgrid.TextMatrix(i, 5) = rs!phone1
         fgrid.TextMatrix(i, 6) = rs!status
         
         
         
          i = i + 1
           rs.MoveNext
     
        Loop
'        If ListView1.ListItems.Count <> 0 Then
'        ListView1.ListItems(1).Selected = True
'        End If
End Sub
Private Sub cleargrid()
        fgrid.Clear
        fgrid.Rows = 1
        fgrid.FormatString = "^Sl.No.|^Farmer|^CID|^T.No.|^H.No.|^Phone. No.|^Status|^"
        fgrid.ColWidth(0) = 615
        fgrid.ColWidth(1) = 2535
        fgrid.ColWidth(2) = 1320
        fgrid.ColWidth(3) = 1020
        fgrid.ColWidth(4) = 960
        fgrid.ColWidth(5) = 1035
        fgrid.ColWidth(6) = 1200
        fgrid.ColWidth(7) = 270
End Sub

Private Sub txttotalarea_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtvillage_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtvillage.SelStart + 1
    Dim sText As String
    sText = Left$(txtvillage.Text, iPos)
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
