VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmplantedlist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLANTED"
   ClientHeight    =   10590
   ClientLeft      =   3105
   ClientTop       =   495
   ClientWidth     =   14880
   Icon            =   "frmplantedlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   14880
   Begin VB.Frame Frame6 
      Height          =   2055
      Left            =   0
      TabIndex        =   86
      Top             =   8400
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtincqty 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5400
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox materialid 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   600
         TabIndex        =   88
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid ItemGrd1 
         Height          =   1695
         Left            =   0
         TabIndex        =   87
         Top             =   120
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         BackColorBkg    =   16777215
         Appearance      =   0
         FormatString    =   $"frmplantedlist.frx":0E42
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Select Land Reg. Trn. No. to sattle"
      Height          =   2055
      Left            =   1320
      TabIndex        =   81
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command8 
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
         Left            =   4800
         Picture         =   "frmplantedlist.frx":0ED5
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox lsttrn 
         Appearance      =   0  'Flat
         Columns         =   10
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         ItemData        =   "frmplantedlist.frx":125F
         Left            =   240
         List            =   "frmplantedlist.frx":1266
         Style           =   1  'Checkbox
         TabIndex        =   82
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.TextBox txtfertamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5880
      TabIndex        =   84
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   6240
      Picture         =   "frmplantedlist.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Load all farmers."
      Top             =   2040
      Width           =   495
   End
   Begin VB.Frame frmaddcrate 
      Caption         =   "Add Crate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11880
      TabIndex        =   76
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox txtcrateno 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1080
         TabIndex        =   78
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1200
         Picture         =   "frmplantedlist.frx":19DC
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Crate No."
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
         TabIndex        =   79
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   4335
      Left            =   9240
      TabIndex        =   70
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
      Begin VSFlex7Ctl.VSFlexGrid fgrid 
         Height          =   2655
         Left            =   240
         TabIndex        =   74
         Top             =   960
         Width           =   5055
         _cx             =   8916
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmplantedlist.frx":2186
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
      Begin VB.CommandButton Command4 
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
         Left            =   1920
         Picture         =   "frmplantedlist.frx":221D
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   3720
         Width           =   1215
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
         Left            =   1560
         TabIndex        =   71
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "FARMER CODE"
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
         TabIndex        =   73
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   6720
      Picture         =   "frmplantedlist.frx":25A7
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox TXTMMONTH 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Crate Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   6240
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   8175
      Begin VB.CommandButton Command5 
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
         Left            =   2640
         Picture         =   "frmplantedlist.frx":2D51
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtfind 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
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
         Height          =   735
         Left            =   1440
         Picture         =   "frmplantedlist.frx":361B
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
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
         Left            =   240
         Picture         =   "frmplantedlist.frx":42E5
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox DZLIST 
         Columns         =   10
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         ItemData        =   "frmplantedlist.frx":466F
         Left            =   120
         List            =   "frmplantedlist.frx":4676
         Style           =   1  'Checkbox
         TabIndex        =   44
         Top             =   840
         Width           =   7935
      End
   End
   Begin VB.TextBox txtdsheetqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtbacktonursery 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox txtsenttofield 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtchallanqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   7440
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker txtchallandate 
      Height          =   375
      Left            =   1440
      TabIndex        =   54
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   148242433
      CurrentDate     =   41516
   End
   Begin VB.TextBox txtmonth 
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtyear 
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtchallanno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   35
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtttot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CheckBox CHKFILLIN 
      Caption         =   "FILL IN TREES"
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TXTTOTTREES 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox TXTTOTAREA 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
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
      Height          =   405
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox TXTTREES 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TXTPLANTED 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame1 
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
      Height          =   4335
      Left            =   7560
      TabIndex        =   3
      Top             =   720
      Width           =   6615
      Begin VB.TextBox TXTTREESSOFAR 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox TXTTS 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox TXTGE 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox TXTDZ 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox TXTADDLAND 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txtplantedsofar 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtregarea 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox txtfarmername 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtfarmercode 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "PLANTED TREES"
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
         Left            =   3720
         TabIndex        =   31
         Top             =   3480
         Width           =   1545
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
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ADDITIONAL LAND"
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
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "AREA PLANTED"
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
         TabIndex        =   12
         Top             =   3480
         Width           =   1425
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AREA REGISTERED"
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
         TabIndex        =   11
         Top             =   3000
         Width           =   2430
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
         TabIndex        =   10
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "FARMER CODE"
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
         TabIndex        =   9
         Top             =   600
         Width           =   1350
      End
   End
   Begin MSDataListLib.DataCombo cbofarmerid 
      Bindings        =   "frmplantedlist.frx":4682
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSComctlLib.ImageList IMG 
      Left            =   4920
      Top             =   120
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
            Picture         =   "frmplantedlist.frx":4697
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplantedlist.frx":4A31
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplantedlist.frx":4DCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplantedlist.frx":5AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplantedlist.frx":5EF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplantedlist.frx":66B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
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
      Height          =   3015
      Left            =   7560
      TabIndex        =   4
      Top             =   6240
      Width           =   7215
      _cx             =   12726
      _cy             =   5318
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmplantedlist.frx":6A4B
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
   Begin MSDataListLib.DataCombo cbotrnid 
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   2985
      Left            =   -360
      TabIndex        =   37
      Top             =   3840
      Width           =   7575
      Begin VB.TextBox txtnoofcrates 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   40
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtvarietyid 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         TabIndex        =   39
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo cbovariety 
         Bindings        =   "frmplantedlist.frx":6B07
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1440
         TabIndex        =   38
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
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   2700
         Left            =   360
         TabIndex        =   41
         Top             =   120
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4763
         _Version        =   393216
         Rows            =   210
         Cols            =   7
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "       |^|^  Variety    |^No. Of Crates    |^                Crate #                                                   |Qty.     |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   42
         Top             =   4920
         Width           =   870
      End
   End
   Begin MSDataListLib.DataCombo cbodeliveryno 
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1440
      TabIndex        =   48
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   ""
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
   Begin MSDataListLib.DataCombo cbostaff 
      Bindings        =   "frmplantedlist.frx":6B1C
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   3360
      TabIndex        =   67
      Top             =   2520
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
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
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "AMOUNT"
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
      Left            =   4920
      TabIndex        =   85
      Top             =   3600
      Width           =   810
   End
   Begin VB.Label Label26 
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
      Left            =   2520
      TabIndex        =   68
      Top             =   2640
      Width           =   840
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "D. SHEET QTY."
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
      Left            =   4080
      TabIndex        =   59
      Top             =   8040
      Width           =   1380
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "BACK TO NURSERY"
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
      TabIndex        =   58
      Top             =   7920
      Width           =   1770
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "SENT TO  FIELD"
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
      TabIndex        =   57
      Top             =   7440
      Width           =   1470
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "CHALLAN D. NO."
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
      Left            =   3960
      TabIndex        =   56
      Top             =   7560
      Width           =   1500
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "CHALLAN QTY."
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
      TabIndex        =   55
      Top             =   7080
      Width           =   1350
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "CHALLAN DATE"
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
      Left            =   0
      TabIndex        =   53
      Top             =   3120
      Width           =   1410
   End
   Begin VB.Label Label19 
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
      Left            =   4560
      TabIndex        =   50
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label Label13 
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
      Left            =   3000
      TabIndex        =   49
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "ACRE "
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
      Left            =   0
      TabIndex        =   47
      Top             =   3480
      Width           =   570
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "CHALLAN NO."
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
      Left            =   0
      TabIndex        =   36
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Line Line2 
      BorderWidth     =   10
      X1              =   7440
      X2              =   14160
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL PLANTED TREES"
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
      Left            =   10920
      TabIndex        =   28
      Top             =   5400
      Width           =   2205
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL PLANTED AREA"
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
      Left            =   7680
      TabIndex        =   27
      Top             =   5400
      Width           =   2085
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "TREES"
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
      Left            =   2760
      TabIndex        =   26
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TRN. ID"
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
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "D.NO."
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
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   10
      X1              =   7440
      X2              =   7440
      Y1              =   600
      Y2              =   9240
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
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   1035
   End
End
Attribute VB_Name = "frmplantedlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfr As New ADODB.Recordset
Dim RSTR As New ADODB.Recordset
Dim Srs As New ADODB.Recordset
Dim TOTPLANTEDAREA As Double
Dim TOTALNOOFTREES As Double
Dim ValidRow As Boolean
Dim CurrRow As Long
Dim mytot As Integer
Dim datechk As Boolean
Dim Dzstr As String
Dim chqty As Long
Dim mcratecnt As Integer
Dim trnid As String
Dim chknocrates As Boolean



Private Sub cbodeliveryno_LostFocus()
Dim i As Integer
'On Error GoTo err
Dim rs As New ADODB.Recordset
If Len(cbodeliveryno.Text) = 0 Then Exit Sub
If Operation = "ADD" Then
txtyear.Text = Mid(cbodeliveryno.Text, Len(cbodeliveryno.Text) - 4, 5)
txtyear.Text = Trim(txtyear.Text)
cbodeliveryno.Text = cbodeliveryno.BoundText
End If

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer a, tblplantdistributiondetail b where idfarmer=farmercode and b.status not in ('C','F') and challanentered<>'Y' and trnid in (select trnid from tblplantdistributionheader where status='ON') and distno='" & cbodeliveryno.BoundText & "'  order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
cbodeliveryno.Enabled = False
datechk = False
Set rs = Nothing
rs.Open "select * from tblplanted where dno='" & cbodeliveryno.BoundText & "' and year='" & txtyear.Text & "' order by trnid desc", MHVDB
If rs.EOF <> True Then
txtchallandate.Value = Format(IIf(IsNull(rs!challandate), "1999-01-01", rs!challandate), "dd/MM/yyyy")
txtchallanno.Text = rs!challanserial & rs!challanno + 1
fillstaffcode
Else
MsgBox "Check the challan date."
End If

fllvariety
fillsummary
Exit Sub
'err:
'    MsgBox "Please select the delivery no. from the dropdown menu!"
'    cbodeliveryno.SetFocus
End Sub
Private Sub fillstaffcode()
If IsNumeric(Mid(txtchallanno.Text, 1, 1)) Then
MsgBox "Please check the challan series"
Exit Sub
End If

If Not IsNumeric(Mid(txtchallanno.Text, 2, 1)) Then
MsgBox "Please check the challan series"
Exit Sub
End If

cbostaff.Text = ""
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "SELECT * FROM `tblchallanchecklist` WHERE alphabetseries ='" & Mid(txtchallanno.Text, 1, 1) & "' and ('" & Val(Mid(txtchallanno.Text, 2, 10)) & "' -`challannofrom`) between 0 and 50", MHVDB
If rs.EOF <> True Then
FindsTAFF rs!staffcode
cbostaff.Text = rs!staffcode & "  " & sTAFF
Else
cbostaff.Text = ""
MsgBox "Please check challan book!"
Exit Sub
End If
End Sub
Private Sub fllvariety()

If Operation = "OPEN" Then Exit Sub
Dim i As Integer
i = 1
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsplantvariety where status='ON' order by challanorder asc", MHVDB
Do While rs.EOF <> True
ItemGrd.TextMatrix(i, 2) = rs!Description
ItemGrd.TextMatrix(i, 1) = rs!varietyId
i = i + 1
rs.MoveNext
Loop
End Sub

Private Sub filmaterials()

If Operation = "OPEN" Then Exit Sub
Dim i As Integer
i = 1
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblincentivematlmaster where status='Active' order by incentiveid asc", MHVDB
Do While rs.EOF <> True
ItemGrd1.TextMatrix(i, 2) = rs!incentivematerial
ItemGrd1.TextMatrix(i, 3) = rs!unit
ItemGrd1.TextMatrix(i, 1) = rs!incentiveid
i = i + 1
rs.MoveNext
Loop
End Sub


Private Sub cbofarmerid_LostFocus()
Dim rstrnid As New ADODB.Recordset

trnid = ""
lsttrn.Clear

loadtrnno Mid(Trim(cbofarmerid.Text), 1, 14)

If Len(trnid) = 0 Then
Set rstrnid = Nothing
rstrnid.Open "select * from tblplantdistributiondetail where distno='" & cbodeliveryno.BoundText & "' and year='" & Val(txtyear.Text) & "'", MHVDB
If rstrnid.EOF <> True Then
Else
MsgBox "You cannot proceed, please confirm this farmer has additional/registered land!"
Exit Sub
End If

End If



If Len(cbofarmerid.Text) = 0 Then Exit Sub
TXTMMONTH.Text = ""
cbofarmerid.Enabled = False
cbodeliveryno.Enabled = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
Set rs = Nothing
rs.Open "select * from tblplantdistributiondetail where distno='" & cbodeliveryno.BoundText & "' and farmercode='" & cbofarmerid.BoundText & "' and status<>'C'", MHVDB
If rs.EOF <> True Then
txtyear.Text = rs!Year
txtmonth.Text = MonthName(rs!mnth, False)
TXTMMONTH.Text = rs!mnth
TXTPLANTED.Text = Format(rs!area, "###0.00")
TXTTREES.Text = rs!totalplant
fillgridch
End If
FillGrid "N"

Frame6.Visible = True
filmaterials
End Sub
Private Sub fillgridch()
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
rs.Open "select distinct plantvariety from tblqmsplantbatchdetail where plantbatch in(select plantbatch from tblqmssendtofielddetail )", MHVDB
Do While rs.EOF <> True
ItemGrd.TextMatrix(i, 1) = rs!plantvariety
findQmsBatchDetail rs!plantvariety
' incomplete here
rs.MoveNext
Loop

End Sub
Private Sub fillsummary()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "SELECT SUM(bcrate*35)+sum(ecrate*35) +SUM(bno*35) + SUM(crate*35) dquantity FROM  `tblplantdistributiondetail` " _
& " WHERE  distno='" & cbodeliveryno.BoundText & "' and year='" & Val(txtyear.Text) & "' and " _
& " subtotindicator='' and status not in ('C','F') and distno>0 and trnid in " _
& " (select trnid from tblplantdistributionheader where status='ON') GROUP BY distno", MHVDB
If rs.EOF <> True Then
txtdsheetqty.Text = IIf(IsNull(rs!dquantity), 0, rs!dquantity)
End If

Label17.Caption = ""
Label17.Caption = "CHALLAN ON D. NO.: " & cbodeliveryno.BoundText
Set rs = Nothing
rs.Open "SELECT SUM( challanqty ) challanqty FROM  `tblplanted` " _
& " WHERE  dno='" & cbodeliveryno.BoundText & "' and year='" & Val(txtyear.Text) & "' and status<>'C' GROUP BY dno", MHVDB
If rs.EOF <> True Then
txtchallanqty.Text = IIf(IsNull(rs!challanqty), 0, rs!challanqty)
End If

Set rs = Nothing
rs.Open "select sum(credit)sendtofield from tblqmsplanttransaction where transactiontype='4' and distributionno='" & cbodeliveryno.BoundText & "' and distyear='" & Val(txtyear.Text) & "' and status='ON'", MHVDB
If rs.EOF <> True Then
txtsenttofield.Text = IIf(IsNull(rs!sendtofield), 0, rs!sendtofield)
End If

Set rs = Nothing
rs.Open "select sum(debit)sendtofield from tblqmsplanttransaction where transactiontype='5' and distributionno='" & cbodeliveryno.BoundText & "' and distyear='" & Val(txtyear.Text) & "' and status='ON'", MHVDB
If rs.EOF <> True Then
txtbacktonursery.Text = IIf(IsNull(rs!sendtofield), 0, rs!sendtofield)
End If

End Sub



Private Sub cbotrnid_LostFocus()
cbotrnid.Text = cbotrnid.BoundText
cbotrnid.Enabled = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblplanted where trnid='" & cbotrnid.BoundText & "' and mnth>0", MHVDB
If rs.EOF <> True Then
cbodeliveryno.Text = rs!dno
txtyear.Text = rs!Year
fillsummary
txtmonth.Text = MonthName(rs!mnth, False)
TXTMMONTH.Text = rs!mnth
FindFA rs!farmercode, "F"
cbofarmerid.Text = rs!farmercode & " " & FAName
txtchallanno.Text = IIf(rs!challanserial = "", "", rs!challanserial) & rs!challanno
txtchallandate.Value = Format(IIf(IsNull(rs!challandate), "01/01/2013", rs!challandate), "dd/MM/yyyy")
TXTPLANTED.Text = rs!acreplanted
TXTTREES.Text = rs!nooftrees
txtttot.Text = rs!challanqty
txtfertamount.Text = rs!fertamount
FindsTAFF rs!staffcode
cbostaff.Text = rs!staffcode & " " & sTAFF
Else

cbodeliveryno.Text = ""
txtyear.Text = ""
txtmonth.Text = ""
TXTMMONTH.Text = ""

cbofarmerid.Text = ""
txtchallanno.Text = ""
txtchallandate.Value = Format(Now, "dd/MM/yyyy")
TXTPLANTED.Text = ""
TXTTREES.Text = ""
txtttot.Text = ""
End If

fillvariety
FillGrid "N"
If Operation = "OPEN" Then
Frame6.Visible = True
fillincentivematerial
End If
End Sub
Private Sub fillvariety()
Dim i, j As Integer
Dim mvariety As String
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblplanteddetail where trnid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
i = 1
mvariety = ""
Do While rs.EOF <> True
ItemGrd.TextMatrix(i, 1) = rs!varietyId
FindqmsPlantVariety rs!varietyId
ItemGrd.TextMatrix(i, 2) = qmsPlantVariety
mvariety = mvariety + "'" + qmsPlantVariety + "',"
ItemGrd.TextMatrix(i, 3) = rs!noofcrates
ItemGrd.TextMatrix(i, 4) = rs!cratedetail
ItemGrd.TextMatrix(i, 5) = rs!crateqty

i = i + 1
rs.MoveNext
Loop


End If

If Len(mvariety) > 0 Then
   mvariety = "(" + Left(mvariety, Len(mvariety) - 1) + ")"
   Else
   i = 1
End If
Set rs = Nothing

If Len(mvariety) = 0 Then
rs.Open "select * from tblqmsplantvariety where status='ON' order by challanorder ", MHVDB
Else
rs.Open "select * from tblqmsplantvariety where status='ON' and description not in " & mvariety & " order by challanorder", MHVDB
End If
Do While rs.EOF <> True
ItemGrd.TextMatrix(i, 2) = rs!Description
ItemGrd.TextMatrix(i, 1) = rs!varietyId
i = i + 1
rs.MoveNext
Loop
End Sub

Private Sub fillincentivematerial()
Dim i As Integer
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblincentivematltrn where headerid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
i = 1
Do While rs.EOF <> True
ItemGrd1.TextMatrix(i, 1) = rs!incentiveid
Findincentivematerialname rs!incentiveid
ItemGrd1.TextMatrix(i, 2) = incentivematerialname
ItemGrd1.TextMatrix(i, 3) = incentivematerialunit
If (rs!incentiveqty) > 0 Then
ItemGrd1.TextMatrix(i, 4) = rs!incentiveqty
Else
ItemGrd1.TextMatrix(i, 4) = ""
End If
i = i + 1
rs.MoveNext
Loop

End If
End Sub

Private Sub CHKFILLIN_Click()
If CHKFILLIN.Value = 1 Then
TXTPLANTED.Locked = True
TXTPLANTED.Enabled = False
Else
TXTPLANTED.Locked = False
TXTPLANTED.Enabled = True
End If
End Sub

Private Sub CHKFILLIN_LostFocus()
CHKFILLIN.Enabled = False
End Sub

Private Sub Command1_Click()
Frame3.Visible = False
frmaddcrate.Visible = False
ItemGrd.Enabled = True
Call SetWindowLong(hwnd, GWL_WNDPROC, lPrevWndProc)
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dzstr = ""
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + Trim(DZLIST.List(i)) + ","
         End If
    Next
If Len(Dzstr) > 0 Then
   Dzstr = Left(Dzstr, Len(Dzstr) - 1)
   ItemGrd.TextMatrix(CurrRow, 4) = ""
 ItemGrd.TextMatrix(CurrRow, 4) = Dzstr
 Frame3.Visible = False
 ItemGrd.Enabled = True
Else
   MsgBox "CRATE NOT SELECTED !!!"
   Frame3.Visible = False
     ItemGrd.Enabled = True
    
   Exit Sub
End If

End Sub


Private Sub cratecnt()
Dim i As Integer
mcratecnt = 0
For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
      mcratecnt = mcratecnt + 1
         End If
    Next
    txtcratecnt.Text = mcratecnt
End Sub

Private Sub Command3_Click()


Frame4.Visible = True

End Sub

Private Sub Command4_Click()
Frame4.Visible = False
End Sub

Private Sub Command5_Click()
frmaddcrate.Visible = True
txtcrateno.Text = ""
End Sub

Private Sub Command6_Click()
Dim rs As New ADODB.Recordset

If Len(Trim(txtcrateno.Text)) = 0 Then
frmaddcrate.Visible = False
Exit Sub
End If


Set rs = Nothing
rs.Open "select * from tblqmscrate where crateno='" & Trim(txtcrateno.Text) & "' ", MHVDB
If rs.EOF <> True Then
MsgBox "Crate Already Exists."
Else

MHVDB.Execute "insert into tblqmscrate(crateno) values('" & Trim(txtcrateno.Text) & "')"
End If
txtcrateno.Text = ""
frmaddcrate.Visible = False
Command2_Click
Frame3.Visible = False

End Sub

Private Sub Command7_Click()
cbodeliveryno.Enabled = True
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer a, tblplantdistributiondetail b where idfarmer=farmercode and b.status not in ('C','F') and trnid in (select trnid from tblplantdistributionheader where status='ON') and distno='" & cbodeliveryno.BoundText & "' order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"

End Sub

Private Sub Command8_Click()
Dim i As Integer
Dim dd
trnid = ""
For i = 0 To lsttrn.ListCount - 1
    If lsttrn.Selected(i) Then
       'trnId = trnId + Trim(lsttrn.List(i)) + ","
       dd = Split(lsttrn.List(i), "|", -1, vbTextCompare)
       trnid = trnid + dd(0) + ","
         End If
    Next
If Len(trnid) > 0 Then
   trnid = "(" + Left(trnid, Len(trnid) - 1) + ")"

   Frame5.Visible = False
Else
MsgBox "Please select the transaction no. from the list!"
    Frame5.Visible = True
   Exit Sub
End If
End Sub

Private Sub DZLIST_Click()
cratecnt
End Sub

Private Sub Form_Load()


Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsDz = Nothing

Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select concat(farmercode,' ',farmername,' ',cast(trnid as char)) as farmername,trnid  from tblplanted as a,tblfarmer as b where a.farmercode=b.idfarmer order by trnid desc", db
Set cbotrnid.RowSource = RSTR
cbotrnid.ListField = "farmername"
cbotrnid.BoundColumn = "trnid"

'tblplantdistributiondetail
Set rsfr = Nothing

If rsfr.State = adStateOpen Then rsfr.Close
'rsfr.Open "select distinct distno,concat(cast(distno as char) , '  ', cast(year as char)) dist  from tblplantdistributiondetail where  subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON')order by distno desc", db
'rsfr.Open "select distinct distno,concat(cast(distno as char) , '  ', cast(year as char)) dist  from tblplantdistributiondetail where  subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON')order by distno desc", db
rsfr.Open "select distinct distributionno as distno,concat(cast(distributionno as char) , '  ', cast(year as char)) dist  from tblqmssendtofieldhdr order by distributionno desc", db
Set cbodeliveryno.RowSource = rsfr
cbodeliveryno.ListField = "dist"
cbodeliveryno.BoundColumn = "distno"

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff where moniter='1' or advocate='1' order by STAFFCODE", db
Set cbostaff.RowSource = Srs
cbostaff.ListField = "STAFFNAME"
cbostaff.BoundColumn = "STAFFCODE"


'TXTYEAR.Value = 2013


ValidRow = True
CurrRow = 1
ItemGrd.row = 1
ItemGrd.col = 1
txtvarietyid.Left = ItemGrd.Left + ItemGrd.CellLeft
txtvarietyid.Width = ItemGrd.CellWidth
txtvarietyid.Height = ItemGrd.CellHeight

ItemGrd.col = 3
txtnoofcrates.Left = ItemGrd.Left + ItemGrd.CellLeft
txtnoofcrates.Width = ItemGrd.CellWidth
txtnoofcrates.Height = ItemGrd.CellHeight

'FillGrid "A"
TOTINFO

End Sub
Private Sub loadcrate()
Dim cnt As Integer
Dim rs As New ADODB.Recordset
Dim i, j As Integer
Dim dd As Variant


Dim crateStr As String
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If


crateStr = ""
Set rs = Nothing
rs.Open "select * from tblplanteddetail where trnid in(select trnid from tblplanted where dno='" & cbodeliveryno.BoundText & "' and year='" & txtyear.Text & "')", MHVDB
If rs.EOF <> True Then
Do While rs.EOF <> True
crateStr = crateStr + Trim(rs!cratedetail) + ","
rs.MoveNext
Loop
End If
If Len(crateStr) > 0 Then
crateStr = Left(crateStr, Len(crateStr) - 1)
crateStr = "(" & crateStr & ")"
End If

i = 1
Set rs = Nothing
txtfind.Text = ""
txtcratecnt.Text = ""
DZLIST.Clear
'If Len(cratestr) > 0 Then
'rs.Open "select * from tblqmssendtofielddetail where distributionno='" & cbodeliveryno.BoundText & "' and year='" & txtyear.Text & "' and crateno not in " & cratestr & "", MHVDB
'Else
rs.Open "select * from tblqmssendtofielddetail where distributionno='" & cbodeliveryno.BoundText & "' and year='" & txtyear.Text & "' and plantbatch in(select plantbatch from tblqmsplantbatchdetail where plantvariety ='" & Trim(ItemGrd.TextMatrix(CurrRow, 1)) & "') ", MHVDB
'End If
If rs.EOF <> True Then
Frame3.Visible = True
ItemGrd.Enabled = False
Else
'Set rs = Nothing
'rs.Open "select * from tblqmscrate Order by  crateno", MHVDB

MsgBox "This variety is not in the bill of ladding, please confirm!"
Frame3.Visible = False
ItemGrd.Enabled = True
   ItemGrd.TextMatrix(CurrRow, 3) = 0
   ItemGrd.TextMatrix(CurrRow, 5) = 0
Exit Sub
End If
With rs
Do While Not .EOF
If i > 15 Then
i = 0
End If
   DZLIST.AddItem Trim(!crateno)
   DZLIST.ItemData(DZLIST.NewIndex) = QBColor(i)
   i = i + 1
   .MoveNext
Loop
End With

If Len(ItemGrd.TextMatrix(CurrRow, 4)) > 0 Then
dd = Split(ItemGrd.TextMatrix(CurrRow, 4), ",", -1, vbTextCompare)
'dd = Split("ItemGrd.TextMatrix(CurrRow, 4)", ",")
cnt = Len(ItemGrd.TextMatrix(CurrRow, 4)) - Len(Replace(ItemGrd.TextMatrix(CurrRow, 4), ",", ""))
'Len(x) - Len(Replace(x, ",", ""))
For j = 0 To cnt
For i = 0 To DZLIST.ListCount - 1

If dd(j) = Trim(DZLIST.List(i)) Then
DZLIST.Selected(i) = True
End If

'    If DZLIST.Selected(i) Then
'       DZstr = DZstr + Trim(DZLIST.List(i)) + ","
'         End If
    Next
Next


End If
'txtfind.Visible = True
'txtfind.SetFocus
End Sub
Private Sub TOTINFO()
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select sum(acreplanted) as ap,sum(challanqty) as nt from tblplanted where status='ON'", MHVDB
TXTTOTAREA.Text = Format(IIf(IsNull(rs!ap), 0, rs!ap), "####0.00")
TXTTOTTREES.Text = IIf(IsNull(rs!nt), 0, rs!nt)
End Sub
Private Sub FillGrid(ff As String)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.rows = 1
Mygrid.FormatString = "^TRN.NO.|^FARMER CODE|^FARMER NAME|^AREA|^TREES|^YEAR"
Mygrid.ColWidth(0) = 900
Mygrid.ColWidth(1) = 1800
Mygrid.ColWidth(2) = 2250
Mygrid.ColWidth(3) = 585
Mygrid.ColWidth(4) = 753
Mygrid.ColWidth(5) = 735
TOTALNOOFTREES = 0
TOTPLANTEDAREA = 0
If ff = "A" Then
rs.Open "select * from tblplanted order by farmercode", MHVDB, adOpenForwardOnly, adLockOptimistic

Else
rs.Open "select * from tblplanted where farmercode='" & Mid(Trim(cbofarmerid.BoundText), 1, 14) & "'", MHVDB, adOpenForwardOnly, adLockOptimistic

End If
i = 1
Do While rs.EOF <> True
Mygrid.rows = Mygrid.rows + 1
Mygrid.TextMatrix(i, 0) = rs!trnid

Mygrid.TextMatrix(i, 1) = rs!farmercode
FindFA rs!farmercode, "F"
Mygrid.TextMatrix(i, 2) = FAName
Mygrid.TextMatrix(i, 3) = Format(IIf(IsNull(rs!acreplanted), 0, rs!acreplanted), "####0.00")
TOTPLANTEDAREA = TOTPLANTEDAREA + IIf(IsNull(rs!acreplanted), 0, rs!acreplanted)
Mygrid.TextMatrix(i, 4) = IIf(IsNull(rs!challanqty), 0, rs!challanqty)
TOTALNOOFTREES = TOTALNOOFTREES + IIf(IsNull(rs!challanqty), 0, rs!challanqty)
Mygrid.TextMatrix(i, 5) = IIf(IsNull(rs!Year), "", rs!Year)
rs.MoveNext
i = i + 1
Loop

rs.Close

Mygrid.MergeCol(1) = True
Mygrid.MergeCells = 1
Mygrid.MergeCol(2) = True
Mygrid.MergeCells = 2
Exit Sub
err:
Mygrid.MergeCol(1) = True
Mygrid.MergeCells = 1
Mygrid.MergeCol(2) = True
Mygrid.MergeCells = 2
MsgBox err.Description

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



Private Sub ItemGrd_Click()
Dim mrow, MCOL As Integer
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
chknocrates = False
mrow = ItemGrd.row
MCOL = ItemGrd.col
If mrow = 0 Then Exit Sub
'If mrow > 1 And Len(ItemGrd.TextMatrix(mrow - 1, 4)) = 0 Then
'   Beep
'   Exit Sub
'End If
ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = mrow
ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
Select Case MCOL
       
       Case 1
            txtvarietyid.Top = ItemGrd.Top + ItemGrd.CellTop
            txtvarietyid = ItemGrd.Text
            txtvarietyid.Visible = True
            txtvarietyid.SetFocus
       Case 2
'            cbovariety.Top = ItemGrd.Top + ItemGrd.CellTop
'            cbovariety = ItemGrd.Text
'            cbovariety.BoundText = ItemGrd.TextMatrix(CurrRow, 1)
'            cbovariety.Visible = True
'            cbovariety.SetFocus
       Case 3
            If Len(ItemGrd.TextMatrix(mrow, 1)) > 0 Then
               txtnoofcrates.Top = ItemGrd.Top + ItemGrd.CellTop
               txtnoofcrates = ItemGrd.Text
               txtnoofcrates.Visible = True
               txtnoofcrates.SetFocus
            End If
       Case 4
            If ItemGrd.col = 4 And Val(ItemGrd.TextMatrix(CurrRow, 3)) > 0 And Len(ItemGrd.TextMatrix(CurrRow, 4)) = 0 Then
            loadcrate
            
            End If
    End Select
    getsum
End Sub

Private Sub ItemGrd_DblClick()
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
Dim i, j As Integer
If ItemGrd.col = 4 And Val(ItemGrd.TextMatrix(CurrRow, 3)) > 0 And Len(ItemGrd.TextMatrix(CurrRow, 3)) > 0 Then
            loadcrate
            
            
            'Frame3.Visible = True
            'ItemGrd.Enabled = False
            
            
            End If
            
            
            
  
  If ItemGrd.col = 5 And Len(ItemGrd.TextMatrix(ItemGrd.row, 4)) > 0 And (ItemGrd.TextMatrix(ItemGrd.row, 2) = "P1" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "P" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "N" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "B" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "E" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "A" Or ItemGrd.TextMatrix(ItemGrd.row, 2) = "L1") Then
  


myinput = InputBox("Enter No. of plants of variety " & ItemGrd.TextMatrix(ItemGrd.row, 2))
            If Not IsNumeric(myinput) Then
            MsgBox "Invalid number,Double Click again to enable the input box."
            Else
            ItemGrd.TextMatrix(ItemGrd.row, 5) = CInt(myinput)
            getsum
            End If
          
End If
  

            
            
End Sub

Private Sub ItemGrd1_Click()

Dim mrow, MCOL As Integer
If Not ValidRow And CurrRow <> ItemGrd1.row Then
   ItemGrd1.row = CurrRow
   Exit Sub
End If
mrow = ItemGrd1.row
MCOL = ItemGrd1.col
If mrow = 0 Then Exit Sub
'If mrow > 1 And Len(ItemGrd.TextMatrix(mrow - 1, 4)) = 0 Then
'   Beep
'   Exit Sub
'End If
ItemGrd1.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = mrow
ItemGrd1.TextMatrix(CurrRow, 0) = Chr(174)
Select Case MCOL
       
       Case 1
            materialid.Top = ItemGrd1.Top + ItemGrd1.CellTop
            materialid = ItemGrd1.Text
            materialid.Visible = True
            materialid.SetFocus
       Case 2

       Case 3

       Case 4
              If Len(ItemGrd1.TextMatrix(mrow, 1)) > 0 Then
               txtincqty.Top = ItemGrd1.Top + ItemGrd1.CellTop
               txtincqty = ItemGrd1.Text
               txtincqty.Visible = True
               txtincqty.SetFocus
              End If
    End Select
    'getsum

End Sub

Private Sub mygrid_DblClick()
'On Error GoTo ERR
'Dim rs As New ADODB.Recordset
'Dim rs1 As New ADODB.Recordset
'Set rs = Nothing
'cbotrnid.Enabled = False
'rs.Open "select * from tblplanted where trnid='" & Mygrid.TextMatrix(Mygrid.row, 0) & "'", MHVDB
'If rs.EOF <> True Then
'cbotrnid.Text = Mygrid.TextMatrix(Mygrid.row, 0) & " " & rs!FARMERcode & " " & FAName
'FindFA rs!FARMERcode, "F"
'cbofarmerid.Text = rs!FARMERcode & " " & FAName
'
'txtfarmercode.Text = rs!FARMERcode
'txtfarmername.Text = UCase(FAName)
'TXTPLANTED.Text = IIf(IsNull(rs!acreplanted), 0, rs!acreplanted)
'
'FindDZ Mid(rs!FARMERcode, 1, 3)
'FindGE Mid(rs!FARMERcode, 1, 3), Mid(rs!FARMERcode, 4, 3)
'FindTs Mid(rs!FARMERcode, 1, 3), Mid(rs!FARMERcode, 4, 3), Mid(rs!FARMERcode, 7, 3)
'TXTDZ.Text = UCase(Dzname)
'TXTGE.Text = UCase(GEname)
'TXTTS.Text = UCase(TsName)
'Set rs1 = Nothing
'rs1.Open "select sum(acreplanted) as acreplanted from tblplanted where farmercode='" & rs!FARMERcode & "'", MHVDB
'If rs1.EOF <> True Then
'txtplantedsofar.Text = Format(IIf(IsNull(rs1!acreplanted), 0, rs1!acreplanted), "####0.00")
'End If
'Set rs1 = Nothing
'rs1.Open "select sum(regland)as rl from tbllandreg where farmerid='" & rs!FARMERcode & "'", MHVDB
'If rs1.EOF <> True Then
'txtregarea.Text = Format(IIf(IsNull(rs1!rl), 0, rs1!rl), "####0.00")
'End If
'TXTADDLAND.Text = Format(Val(txtregarea.Text) - Val(txtplantedsofar.Text), "####0.00")
'TXTTREES.Text = IIf(IsNull(rs!nooftrees), 0, rs!nooftrees)
'TXTYEAR.Value = IIf(IsNull(rs!Year), "", rs!Year)
''FillGrid rs!FARMERcode
'Else
'MsgBox "No Records Found."
'End If
'Exit Sub
'ERR:
'MsgBox ERR.Description
End Sub

Private Sub MYGRID1_Click()
If mygrid1.col = 0 Then
mygrid1.Editable = flexEDNone
ElseIf mygrid1.col = 4 Then
mygrid1.Editable = flexEDNone
Else

mygrid1.Editable = flexEDKbdMouse
End If

If mygrid1.col = 1 And mygrid1.row = 1 Then
mygrid1.Editable = flexEDKbdMouse
mygrid1.ComboList = "A|" & "B|" & "C"
ElseIf mygrid1.col = 1 And mygrid1.row > 1 And Len(mygrid1.TextMatrix(mygrid1.row - 1, 4)) <> 0 Then
mygrid1.Editable = flexEDKbdMouse
mygrid1.ComboList = "A|" & "B|" & "C"
ElseIf mygrid1.col = 1 And mygrid1.row > 1 And Len(mygrid1.TextMatrix(mygrid1.row - 1, 4)) = 0 Then
mygrid1.ComboList = ""
mygrid1.Editable = flexEDNone
Else

mygrid1.ComboList = ""

End If

If mygrid1.col > 1 And Len(mygrid1.TextMatrix(mygrid1.row, 1)) = 0 Then
mygrid1.Editable = flexEDNone
End If

findtot


End Sub
Private Sub findtot()
Dim i As Integer
mytot = 0
For i = 1 To Mygrid.rows - 1
If Len(mygrid1.TextMatrix(i, 0)) = 0 Then Exit For
mytot = mytot + Val(mygrid1.TextMatrix(i, 4))

Next
txtttot.Text = mytot

End Sub

Private Sub MYGRID1_LeaveCell()
If mygrid1.col = 2 And Val(mygrid1.TextMatrix(mygrid1.row, 2)) > 0 Then
mygrid1.TextMatrix(mygrid1.row, 4) = 35 * Val(mygrid1.TextMatrix(mygrid1.row, 2))
mygrid1.TextMatrix(mygrid1.row, 0) = mygrid1.row + 1
End If
findtot
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ADD"

      
        TB.buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       cbotrnid.Enabled = False
        CHKFILLIN.Enabled = True
      cbofarmerid.Enabled = True
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select (max(trnid)+1) as maxid from tblplanted", MHVDB
   If rs.EOF <> True Then
   cbotrnid.Text = rs!MaxId
   Else
   cbotrnid.Text = 1
   End If
   cbodeliveryno.Enabled = True
   'getDno ("ADD")
         Case "OPEN"
         Operation = "OPEN"
         CLEARCONTROLL
         cbotrnid.Enabled = True
         
         TB.buttons(3).Enabled = True
         

       'getDno ("OPEN")
       Case "SAVE"
      
       MNU_SAVE
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub getDno(op As String)

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString



'tblplantdistributiondetail
If op = "ADD" Then
Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select distinct distno  from tblplantdistributiondetail where  subtotindicator='' and status not in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON') and distno not in(select dno from tblplanted)order by distno ", db
Set cbodeliveryno.RowSource = rsfr
cbodeliveryno.ListField = "distno"
cbodeliveryno.BoundColumn = "distno"
Else
Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select distinct distno  from tblplantdistributiondetail where  subtotindicator='' and status  in ('C','F') and distno>0 and trnid in (select trnid from tblplantdistributionheader where status='ON') and distno  in(select dno from tblplanted)order by distno ", db
Set cbodeliveryno.RowSource = rsfr
cbodeliveryno.ListField = "distno"
cbodeliveryno.BoundColumn = "distno"

End If
End Sub
Private Sub MNU_SAVE()
Dim mTrnid As Long
Dim isempty As Double
Dim chkbox As String


'Dim rs As New ADODB.Recordset
Dim i, j As Integer
If IsNumeric(Mid(txtchallanno.Text, 1, 1)) Then
MsgBox "Please check the challan series"
Exit Sub
End If

If Not IsNumeric(Mid(txtchallanno.Text, 2, 1)) Then
MsgBox "Please check the challan series"
Exit Sub
End If

If Len(cbotrnid.Text) = 0 Then
MsgBox "Transaction ID is must."
Exit Sub
End If

If Val(txtfertamount.Text) = 0 Then
myinput = InputBox("Enter fertilizer amount for challan no." & txtchallanno.Text)
If Not IsNumeric(myinput) Then
MsgBox "Invalid number"
Exit Sub
Else
txtfertamount.Text = CInt(myinput)
End If
End If
If Len(trnid) = 0 And Operation = "ADD" Then
If MsgBox("Is the farmers refillin?", vbYesNo) = vbYes Then
Else
Exit Sub
End If
End If


If Len(Trim(cbostaff.Text)) = 0 Then
MsgBox "Select Monitor for this Delivery No."
Exit Sub
End If


If Val(TXTMMONTH.Text) <= 0 Then
MsgBox "Month of Distribution No. is Must."
Exit Sub
End If

If Len(cbodeliveryno.Text) = 0 Then
MsgBox "Select Distribution No."
Exit Sub
End If

If Len(cbofarmerid.Text) = 0 Then
MsgBox "Select Farmer."
Exit Sub
End If
If Len(txtchallanno.Text) = 0 Then
MsgBox "Enter Challan No."
Exit Sub
End If


If Val(txtttot.Text) <= 0 Then
MsgBox "Enter Crate details."
Exit Sub
End If

Dim tt As Integer
tt = 0
For i = 1 To ItemGrd1.rows - 1
tt = tt + Val(ItemGrd1.TextMatrix(i, 4))
Next


If tt = 0 Then
chkbox = MsgBox("Save without incentive material?", vbYesNo, "Question")
If chkbox = vbYes Then
SaveCall
Else
Exit Sub
End If
Else
SaveCall
End If

End Sub
Private Sub SaveCall()
Dim rs As New ADODB.Recordset
On Error GoTo err
MHVDB.BeginTrans
If Operation = "ADD" Then

Set rs = Nothing
rs.Open "select * from tblplanted where concat(challanserial,challanno)='" & Trim(txtchallanno.Text) & "'", MHVDB
If rs.EOF <> True Then
MsgBox "This Challan No. already entered in Dist. No: " & rs!dno & "  Farmer: " & rs!farmercode & "  " & FAName
MHVDB.RollbackTrans
Exit Sub
End If




Set rs = Nothing
rs.Open "select (max(trnid)+1) as maxid from tblplanted", MHVDB
   If rs.EOF <> True Then
   cbotrnid.Text = rs!MaxId
   Else
   cbotrnid.Text = 1
   End If


MHVDB.Execute "insert into tblplanted(trnid,farmercode,fillins,year,nooftrees,acreplanted," _
& "mnth,dno,challanno,challandate,challanqty,status,staffcode,fertamount,challanserial) values( " _
& "'" & cbotrnid.Text & "'," _
& "'" & cbofarmerid.BoundText & "'," _
& "''," _
& "'" & Val(txtyear.Text) & "'," _
& "'" & Val(TXTTREES.Text) & "'," _
& "'" & Val(TXTPLANTED.Text) & "'," _
& "'" & Val(TXTMMONTH.Text) & "'," _
& "'" & cbodeliveryno.BoundText & "'," _
& "'" & Mid(txtchallanno.Text, 2, 10) & "'," _
& "'" & Format(txtchallandate.Value, "yyyy-MM-dd") & "'," _
& "'" & Val(txtttot.Text) & "'," _
& "'ON'," _
& "'" & Mid(Trim(cbostaff.BoundText), 1, 5) & "','" & Val(txtfertamount.Text) & "','" & Mid(txtchallanno.Text, 1, 1) & "'" _
& ")"

ElseIf Operation = "OPEN" Then

'If (isempty) = 0 Then
'chkbox = MsgBox("Save without incentive material?", vbYesNo, "Question")
'If chkbox = vbYes Then

MHVDB.Execute "update tblplanted set " _
& "farmercode='" & Mid(Trim(cbofarmerid.BoundText), 1, 14) & "'," _
& "year='" & Val(txtyear.Text) & "'," _
& "nooftrees='" & Val(TXTTREES.Text) & "'," _
& "acreplanted='" & Val(TXTPLANTED.Text) & "'," _
& "mnth='" & Val(TXTMMONTH.Text) & "'," _
& "dno='" & cbodeliveryno.BoundText & "'," _
& "challanno='" & Mid(txtchallanno.Text, 2, 10) & "'," _
& "challanserial='" & Mid(txtchallanno.Text, 1, 1) & "'," _
& "challandate='" & Format(txtchallandate.Value, "yyyy-MM-dd") & "'," _
& "challanqty='" & Val(txtttot.Text) & "'," _
& "staffcode='" & Mid(Trim(cbostaff.BoundText), 1, 5) & "',fertamount='" & Val(txtfertamount.Text) & "'" _
& " where trnid='" & cbotrnid.BoundText & "'"

'End If
'End If

Else
MsgBox "Invalid Selection of operation."
MHVDB.RollbackTrans
End If

If Operation = "ADD" Then
MHVDB.Execute "delete from tblplanteddetail where trnid='" & (cbotrnid.Text) & "'"
mTrnid = cbotrnid.Text
ElseIf Operation = "OPEN" Then
MHVDB.Execute "delete from tblplanteddetail where trnid='" & (cbotrnid.BoundText) & "'"
MHVDB.Execute "delete from tblincentivematltrn where headerid='" & (cbotrnid.BoundText) & "'"
mTrnid = cbotrnid.BoundText
Else
MsgBox "Not a valid operation"
MHVDB.RollbackTrans
Exit Sub
End If

For i = 1 To ItemGrd.rows - 1
If Len(ItemGrd.TextMatrix(i, 2)) = 0 Then Exit For
If Val(ItemGrd.TextMatrix(i, 5)) > 0 Then
MHVDB.Execute "insert into tblplanteddetail(trnid,year,mnth,varietyid,noofcrates,cratedetail,crateqty) values( " _
& "'" & mTrnid & "'," _
& "'" & Val(txtyear.Text) & "'," _
& "'" & Val(TXTMMONTH.Text) & "'," _
& "'" & Val(ItemGrd.TextMatrix(i, 1)) & "'," _
& "'" & Val(ItemGrd.TextMatrix(i, 3)) & "'," _
& "'" & ItemGrd.TextMatrix(i, 4) & "'," _
& "'" & Val(ItemGrd.TextMatrix(i, 5)) & "')"
End If
Next



For i = 1 To ItemGrd1.rows - 1
If Len(ItemGrd1.TextMatrix(i, 2)) = 0 Then Exit For
If Val(ItemGrd1.TextMatrix(i, 1)) > 0 Then

MHVDB.Execute "insert into tblincentivematltrn(headerid,farmercode,incentiveid,incentiveqty) values( " _
& "'" & mTrnid & "'," _
& "'" & cbofarmerid.BoundText & "'," _
& "'" & Val(ItemGrd1.TextMatrix(i, 1)) & "'," _
& "'" & Val(ItemGrd1.TextMatrix(i, 4)) & "')"
End If
'isempty = isempty + Val(ItemGrd1.TextMatrix(i, 4))
Next

'Set rs = Nothing
'rs.Open "select count(farmerid) as cnt from "
If Len(trnid) > 0 And Operation = "ADD" And Val(TXTPLANTED.Text) > 0 Then

If Mid(Trim(cbofarmerid.Text), 10, 1) = "F" Then
MHVDB.Execute "update tbllandreg set plantedstatus='C' where trnid in " & trnid
Else
MHVDB.Execute "update tbllandregdetail set plantedstatus='P' where headerid in " & trnid
End If

End If
trnid = ""
MHVDB.Execute "update tblplantdistributiondetail set challanentered='Y' WHERE distno='" & cbodeliveryno.BoundText & "' and year='" & txtyear.Text & "' and farmercode='" & Mid(Trim(cbofarmerid.BoundText), 1, 14) & "'"
'MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='3'"
'MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='4'"
'MHVDB.Execute "UPDATE mhv.tblregistrationsettings SET updatetable ='Yes' WHERE tblid ='7'"
MHVDB.Execute "UPDATE mhv.tblfarmer set category='ActivePoor' where idfarmer='" & Mid(Trim(cbofarmerid.BoundText), 1, 14) & "'"
MHVDB.CommitTrans
TB.buttons(3).Enabled = False
FillGrid "N"
fillsummary

       
       
       TOTINFO
       'TB.Buttons(6).Enabled = True
       Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans
End Sub
Private Sub CLEARCONTROLL()
cbotrnid.Text = ""
cbofarmerid.Text = ""
TXTPLANTED.Text = ""
txtfarmercode.Text = ""
txtfarmername.Text = ""
txtregarea.Text = ""
txtplantedsofar.Text = ""
txtfertamount.Text = ""
TXTADDLAND.Text = ""
TXTDZ.Text = ""
TXTGE.Text = ""
TXTTS.Text = ""
TXTTREES.Text = ""
'TXTYEAR.Value = year(Now)
cbostaff.Text = ""
txtchallandate = Format(Now, "dd/MM/yyyy")
txtchallanno.Text = ""
cbodeliveryno.Text = ""
txtyear.Text = ""
txtmonth.Text = ""
TXTMMONTH.Text = ""


txtdsheetqty.Text = ""
Label17.Caption = "CHALLAN ON D. NO.: "
txtchallanqty.Text = ""
txtsenttofield.Text = ""
txtbacktonursery.Text = ""

txtttot.Text = ""


ItemGrd.Clear
ItemGrd.FormatString = "       |^|^  Variety    |^No. Of Crates    |^                Crate #                                                   |Qty.     |"
ItemGrd1.Clear
ItemGrd1.FormatString = "      |^|^Material                                               |^ Uom                         |^ Qty                                         "

End Sub



Private Sub Text6_Change()


End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtchallanno_KeyPress(KeyAscii As Integer)
If InStr(1, "ABCDE0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtchallanno_LostFocus()
Dim rst As New ADODB.Recordset
Set rst = Nothing
If IsNumeric(Mid(txtchallanno.Text, 1, 1)) Then
MsgBox "Please check the challan series"
Exit Sub
End If

If Not IsNumeric(Mid(txtchallanno.Text, 2, 1)) Then
MsgBox "Please check the challan series"
Exit Sub
End If

cbostaff.Text = ""
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "SELECT * FROM `tblchallanchecklist` WHERE alphabetseries ='" & Mid(txtchallanno.Text, 1, 1) & "' and ('" & Val(Mid(txtchallanno.Text, 2, 10)) & "' -`challannofrom`) between 0 and noofpages-1", MHVDB
If rs.EOF <> True Then
FindsTAFF rs!staffcode
cbostaff.Text = rs!staffcode & "  " & sTAFF
Else
cbostaff.Text = ""
MsgBox "Please check challan book!"
Exit Sub
End If

End Sub

Private Sub txtfind_DblClick()
Dim i As Integer
For i = 0 To DZLIST.ListCount - 1
If txtfind.Text = DZLIST.List(i) Then
DZLIST.Selected(i) = True
End If
Next
txtfind.Text = ""
cratecnt
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Dim i As Integer
For i = 0 To DZLIST.ListCount - 1
If txtfind.Text = DZLIST.List(i) Then
DZLIST.Selected(i) = True

End If
Next
txtfind.Text = ""
cratecnt
End If
cratecnt
End Sub

Private Sub txtnoofcrates_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 2)) > 0 Then
If Not IsNumeric(txtnoofcrates) Then
   Beep
   MsgBox "Enter a valid No."
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 3) = Val(txtnoofcrates.Text)
   ItemGrd.TextMatrix(CurrRow, 5) = Val(txtnoofcrates.Text) * 35
   ValidRow = True
   
End If
End If
txtnoofcrates.Visible = False
getsum
End Sub

Private Sub txtincqty_validate(Cancel As Boolean)
If Len(ItemGrd1.TextMatrix(CurrRow, 2)) > 0 Then
If Not IsNumeric(txtincqty) Then
Beep
MsgBox "Enter a valid No."
ValidRow = False
Cancel = True
Exit Sub
Else
ItemGrd1.TextMatrix(CurrRow, 4) = Val(txtincqty.Text)
ValidRow = True

End If
End If
txtincqty.Visible = False
End Sub

Private Sub getsum()
Dim i As Integer
chqty = 0
For i = 1 To ItemGrd.rows - 1
If Len(ItemGrd.TextMatrix(i, 2)) = 0 Then Exit For
chqty = chqty + Val(ItemGrd.TextMatrix(i, 5))

Next

txtttot.Text = chqty

End Sub

Private Sub TXTSEARCHID_Change()
Dim SQLSTR As String
Dim i As Integer
If Len(TXTSEARCHID.Text) <= 3 Then Exit Sub
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
        SQLSTR = "select concat(idfarmer , ' ', farmername) as farmername,idfarmer, " _
        & " trnid,distno from tblfarmer a, tblplantdistributiondetail b where idfarmer=farmercode" _
        & " and b.status not in ('C','F') and trnid in (select trnid from " _
        & " tblplantdistributionheader where status='ON') and " _
        & " idfarmer like '" & TXTSEARCHID.Text & "%' order by idfarmer"
     
        rs.Open SQLSTR, MHVDB
        If rs.RecordCount > 0 Then
        rs.MoveFirst
        Else
        On Error Resume Next
        End If
         cleargrid
         i = 1
        
        Do Until rs.EOF
        fgrid.rows = fgrid.rows + 1
       
        fgrid.TextMatrix(i, 0) = i
         fgrid.TextMatrix(i, 1) = rs!farmername
         fgrid.ColAlignment(1) = flexAlignLeftTop
          fgrid.TextMatrix(i, 2) = rs!trnid
           fgrid.TextMatrix(i, 3) = rs!distno
          i = i + 1
           rs.MoveNext
     
        Loop
'        If ListView1.ListItems.Count <> 0 Then
'        ListView1.ListItems(1).Selected = True
'        End If

End Sub
Private Sub cleargrid()
        fgrid.Clear
        fgrid.rows = 1
        fgrid.FormatString = "^Sl.No.|^Farmer|^Trn.Id|^Dist. No.|^"
        fgrid.ColWidth(0) = 615
        fgrid.ColWidth(1) = 2535
        fgrid.ColWidth(2) = 780
        fgrid.ColWidth(3) = 780
        fgrid.ColWidth(4) = 270
End Sub
Private Sub loadtrnno(fcode As String)
Dim rs As New ADODB.Recordset
Dim fType As String
Set rs = Nothing
fType = Mid(Trim(fcode), 10, 1)


If Len(fcode) = 0 Then Exit Sub


If fType = "F" Then
rs.Open "select count(*) as cnt,trnid,regland,regdate from tbllandreg where farmerid='" & fcode & "' and plantedstatus='N' group by trnid order by trnid desc", MHVDB
ElseIf fType = "G" Then
rs.Open "select count(*) as cnt,headerid as trnid, acre as regland,'19990101' as regdate from tbllandregdetail where farmercode='" & fcode & "' and plantedstatus='N' group by headerid order by headerid desc", MHVDB
Else
MsgBox "Invalid Farmer type."
Exit Sub
End If

If rs.EOF <> True Then
If IIf(IsNull(rs!cnt), 0, rs!cnt) <> 0 Then
With rs
Do While Not .EOF
   lsttrn.AddItem CStr(!trnid) + " | " + "Acre " & CStr(!regland) & " , Reg Date   " & CStr(!regdate)
   .MoveNext
Loop
End With
Set rs = Nothing




If fType = "F" Then
rs.Open "select count(*) as cnt ,trnid,regland,regdate from tbllandreg where farmerid='" & fcode & "' and plantedstatus='N' order by trnid desc", MHVDB
ElseIf fType = "G" Then
rs.Open "select count(*) as cnt,headerid as trnid, acre as regland,'19990101' as regdate from tbllandregdetail where farmercode='" & fcode & "' and plantedstatus='N'  order by headerid desc", MHVDB
Else
MsgBox "Invalid Farmer type."
Exit Sub
End If



If rs!cnt = 1 Then
trnid = rs!trnid
trnid = "(" + trnid + ")"
Else
trnid = ""
Frame5.Visible = True
End If
End If
End If

End Sub

