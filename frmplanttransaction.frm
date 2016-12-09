VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmplanttransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P L A N T    T R A N S A C T I O N"
   ClientHeight    =   10065
   ClientLeft      =   4215
   ClientTop       =   375
   ClientWidth     =   12330
   Icon            =   "frmplanttransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   12330
   Begin VB.CheckBox chkgrp 
      Caption         =   "Group by date"
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
      TabIndex        =   82
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox txtloc1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   10560
      TabIndex        =   79
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtcol5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtcol4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtcol3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtcol2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtcol1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Plant Batch History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   12255
      Begin VB.TextBox txttot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtdr 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtcr 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   5160
         Picture         =   "frmplanttransaction.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Exit Plant Batch History"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtplantvariety 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtnoofplants 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtrcvdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtshipmetno 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   1455
      End
      Begin VSFlex7Ctl.VSFlexGrid mygrid2 
         Height          =   3135
         Left            =   5760
         TabIndex        =   43
         Top             =   720
         Width           =   6375
         _cx             =   11245
         _cy             =   5530
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmplanttransaction.frx":1B0C
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
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Plant In Stock"
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
         Left            =   7800
         TabIndex        =   57
         Top             =   4560
         Width           =   1230
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   7800
         TabIndex        =   56
         Top             =   4080
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Plant Variety"
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
         TabIndex        =   47
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "No. of Plants Received"
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
         TabIndex        =   46
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Received Date"
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
         TabIndex        =   45
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Shipmet No."
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
         TabIndex        =   44
         Top             =   960
         Width           =   1050
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   12240
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   12240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Batch Transaction Detail"
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
         TabIndex        =   42
         Top             =   360
         Width           =   2130
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Batch Shipment Detail"
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
         TabIndex        =   41
         Top             =   360
         Width           =   1905
      End
      Begin VB.Line Line1 
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   3960
      End
   End
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
      Height          =   735
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   12135
      Begin VB.CheckBox chkvariety 
         Caption         =   "Variety"
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
         Left            =   5280
         TabIndex        =   70
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chktransaction 
         Caption         =   "Transaction Type"
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
         TabIndex        =   61
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkplantbatch 
         Caption         =   "Plant Batch"
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
         TabIndex        =   60
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkfacility 
         Caption         =   "Facility"
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
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkall 
         Caption         =   "All"
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
         TabIndex        =   58
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load Transaction"
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
         Left            =   9960
         Picture         =   "frmplanttransaction.frx":1BC0
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker txtfrmdate 
         Height          =   375
         Left            =   6840
         TabIndex        =   36
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51314689
         CurrentDate     =   41479
      End
      Begin MSComCtl2.DTPicker txttodate 
         Height          =   375
         Left            =   8520
         TabIndex        =   37
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51314689
         CurrentDate     =   41479
      End
      Begin VB.Label Label14 
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
         Left            =   6360
         TabIndex        =   39
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label15 
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
         Left            =   8280
         TabIndex        =   38
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.TextBox txttotalplant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   9600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6255
      Begin VB.ComboBox txtloc 
         Height          =   315
         ItemData        =   "frmplanttransaction.frx":1F4A
         Left            =   4440
         List            =   "frmplanttransaction.frx":1F54
         TabIndex        =   81
         Top             =   720
         Width           =   1455
      End
      Begin VB.Frame framedr 
         Height          =   2175
         Left            =   3600
         TabIndex        =   75
         Top             =   1200
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton cmdokdr 
            Height          =   495
            Left            =   2040
            Picture         =   "frmplanttransaction.frx":1F62
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton cmdcanceldr 
            Height          =   495
            Left            =   2880
            Picture         =   "frmplanttransaction.frx":22EC
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   1560
            Width           =   855
         End
         Begin VSFlex7Ctl.VSFlexGrid mgriddr 
            Height          =   1335
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   3615
            _cx             =   6376
            _cy             =   2355
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
            Rows            =   5
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmplanttransaction.frx":2676
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
      Begin VB.Frame framecr 
         Height          =   2175
         Left            =   3360
         TabIndex        =   71
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton cmdokcr 
            Height          =   495
            Left            =   2040
            Picture         =   "frmplanttransaction.frx":273D
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton cmdcancelcr 
            Height          =   495
            Left            =   2880
            Picture         =   "frmplanttransaction.frx":2AC7
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   1560
            Width           =   855
         End
         Begin VSFlex7Ctl.VSFlexGrid mgridcr 
            Height          =   1335
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   3615
            _cx             =   6376
            _cy             =   2355
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
            Rows            =   5
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmplanttransaction.frx":2E51
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
      Begin VB.TextBox txtcredit 
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
         Left            =   4800
         TabIndex        =   31
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtdebit 
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
         Left            =   2160
         TabIndex        =   8
         Top             =   3360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker txtentrydate 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51314689
         CurrentDate     =   41479
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "frmplanttransaction.frx":2F19
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo cbofacilityId 
         Bindings        =   "frmplanttransaction.frx":2F2E
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   1800
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
      Begin MSDataListLib.DataCombo cboplantBatch 
         Bindings        =   "frmplanttransaction.frx":2F43
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   1080
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
      Begin MSDataListLib.DataCombo cbotransaction 
         Bindings        =   "frmplanttransaction.frx":2F58
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Top             =   2520
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
      Begin MSDataListLib.DataCombo cboverification 
         Bindings        =   "frmplanttransaction.frx":2F6D
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   22
         Top             =   2160
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
      Begin MSDataListLib.DataCombo cbostaff 
         Bindings        =   "frmplanttransaction.frx":2F82
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   28
         Top             =   2880
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
      Begin MSDataListLib.DataCombo cbovariety 
         Bindings        =   "frmplanttransaction.frx":2F97
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Station"
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
         Left            =   3600
         TabIndex        =   80
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Plant variety"
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
         TabIndex        =   33
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Credit"
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
         Left            =   3840
         TabIndex        =   30
         Top             =   3480
         Width           =   630
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
         TabIndex        =   29
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Verification Type"
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
         TabIndex        =   27
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type"
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
         TabIndex        =   26
         Top             =   2640
         Width           =   1845
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
         TabIndex        =   16
         Top             =   240
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
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Facility Id"
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
         TabIndex        =   14
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plant batch"
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
         TabIndex        =   13
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Debit"
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
         Top             =   3480
         Width           =   570
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   12135
      Begin VSFlex7Ctl.VSFlexGrid Mygrid 
         Height          =   2775
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   12015
         _cx             =   21193
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmplanttransaction.frx":2FAC
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
   Begin VB.Frame Frame3 
      Caption         =   "Plant Inventory Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6360
      TabIndex        =   2
      Top             =   720
      Width           =   5895
      Begin VSFlex7Ctl.VSFlexGrid mygrid1 
         Height          =   2655
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   5775
         _cx             =   10186
         _cy             =   4683
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmplanttransaction.frx":30EB
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
   Begin VB.TextBox txtsumdebit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtgendebit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   9600
      Width           =   1335
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
            Picture         =   "frmplanttransaction.frx":318D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplanttransaction.frx":3527
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplanttransaction.frx":38C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplanttransaction.frx":459B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplanttransaction.frx":49ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplanttransaction.frx":51A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   1164
      ButtonWidth     =   1376
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
            Caption         =   "CANCEL"
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
   Begin VB.TextBox txtsumcredit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtgencredit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   9600
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
      Height          =   1185
      Left            =   6360
      Style           =   1  'Checkbox
      TabIndex        =   62
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Debit"
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
      TabIndex        =   69
      Top             =   9720
      Width           =   570
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Credit"
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
      Left            =   5520
      TabIndex        =   68
      Top             =   9720
      Width           =   630
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Plant In Stock"
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
      Left            =   8520
      TabIndex        =   24
      Top             =   9720
      Width           =   1425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Stock"
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
      Left            =   9120
      TabIndex        =   19
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   2160
      TabIndex        =   18
      Top             =   9720
      Width           =   555
   End
End
Attribute VB_Name = "frmplanttransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gendr, gencr As Double
Dim sumdr, sumcr As Double
Dim shortby As String
Dim Dzstr  As String
Dim plantinstock As Long


Private Sub cboFacilityId_LostFocus()

If Len(cboplantBatch.Text) = 0 Then
chkplantbatch.Value = 0
Else
chkplantbatch.Value = 1
End If


If Len(cbofacilityid.Text) = 0 Then
chkfacility.Value = 0
Else
chkfacility.Value = 1
End If

chkall.Value = 0
If chkgrp.Value = 0 Then
FillGrid "01/01/0000", "01/01/2020", shortby
Else
fillgridgrp "01/01/0000", "01/01/2020", shortby
End If

End Sub

Private Sub cboplantBatch_LostFocus()
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
If Operation <> "" Then
rsF.Open "SELECT concat(cast(c.description as char),'  ', b.description) as description,varietyid FROM  `tblqmsplantbatchdetail` a," _
 & "tblqmsplanttype b,tblqmsplantvariety c Where planttype = planttypeid" _
 & " AND plantbatch ='" & cboplantBatch.BoundText & "' and varietyid=plantvariety   order by plantbatch", db
 
 
Set cbovariety.RowSource = rsF
cbovariety.ListField = "description"
cbovariety.BoundColumn = "varietyid"


If rsF.EOF <> True Then
cbovariety.Text = rsF!Description
End If
Else




End If


'FindqmsPlanttype
'findQmsBatchDetail cboplantBatch.BoundText
'cbovariety.Text = qmsplantbatch3
'cboplantBatch.Enabled = False

If Len(cboplantBatch.Text) = 0 Then
chkplantbatch.Value = 0
Else
chkplantbatch.Value = 1
chkall.Value = 0
If chkgrp.Value = 0 Then
FillGrid "01/01/0000", "01/01/2020", shortby
Else
fillgridgrp "01/01/0000", "01/01/2020", shortby
End If
End If
If Operation = "ADD" Then
For i = 1 To mgriddr.Rows - 1
mgriddr.TextMatrix(i, 1) = ""
mgridcr.TextMatrix(i, 1) = ""
Next
End If
End Sub

Private Sub cbotransaction_LostFocus()
'txtdebit.Locked = True
'txtcredit.Locked = True
If Len(cbotransaction.BoundText) = 0 Then Exit Sub

lockDrCr cbotransaction.BoundText
End Sub
Private Sub lockDrCr(i As Integer)
Select Case i
Case 2
txtdebit.Locked = False
Case 3
txtcredit.Locked = False
Case 4
txtcredit.Locked = False
Case 5
txtdebit.Locked = False
Case 6
txtdebit.Locked = False
txtcredit.Locked = False
Case 7
txtdebit.Locked = False
Case 8
txtdebit.Locked = False
Case 9
txtdebit.Locked = False
txtcredit.Locked = False
Case 10
txtdebit.Locked = False
txtcredit.Locked = False
Case 11
txtdebit.Locked = False
txtcredit.Locked = False

Case 12
txtdebit.Locked = False
txtcredit.Locked = False

Case 13
txtdebit.Locked = False
txtcredit.Locked = False

End Select
End Sub

Private Sub cbotrnid_LostFocus()
On Error GoTo err
   
   cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select * from tblqmsplanttransaction where status='ON' and trnid='" & cbotrnid.BoundText & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
   If rs.EOF <> True Then
    txtentrydate.Value = Format(rs!entrydate, "yyyy-MM-dd")
    txtdebit.Text = IIf(rs!debit = 0, "", rs!debit)
    txtcredit.Text = IIf(rs!credit = 0, "", rs!credit)
    txtloc.Text = rs!location
    findQmsfacility rs!facilityid
    cbofacilityid.Text = rs!facilityid & "  " & qmsFacility
    findQmsBatchDetail rs!plantBatch
    
    cboplantBatch.Text = rs!plantBatch
    Findqmsverificationtype rs!verificationType
    cboverification.Text = qmsVerificationType
    Findqmstransactiontype rs!transactiontype
    cbotransaction.Text = qmsTransactionType
    FindqmsPlantVariety rs!varietyId
    cbovariety.Text = qmsPlantVariety
    FindsTAFF rs!staffid
    cbostaff.Text = Trim(rs!staffid & "  " & sTAFF)
    
    
'    subgrid cbotrnid.BoundText
   
   Else
   MsgBox "Record Not Found. Or The record is Cancelled."
    TB.Buttons(3).Enabled = False
        TB.Buttons(4).Enabled = False
   End If
   rs.Close
   Exit Sub
err:
   MsgBox err.Description
   'rs.Close
End Sub
Private Sub subgrid(trnno As Long)
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
rs.Open "select * from tblqmsnutdetail where trnid='" & trnno & "'", MHVDB
Do While rs.EOF <> True
If rs!nutdebit <> 0 Then
For i = 1 To mgriddr.Rows - 1
If Trim(mgriddr.TextMatrix(i, 2)) = Trim(rs!nuttype) Then
mgriddr.TextMatrix(i, 1) = rs!nutdebit
End If
Next

Else
For i = 1 To mgridcr.Rows - 1
If Trim(mgridcr.TextMatrix(i, 2)) = Trim(rs!nuttype) Then
mgridcr.TextMatrix(i, 1) = rs!nutcredit
End If
Next

End If



rs.MoveNext
Loop
End Sub

Private Sub cbovariety_Change()
'Dim rsF As New ADODB.Recordset
'Set db = New ADODB.Connection
'db.CursorLocation = adUseClient
'db.Open CnnString
'If rsF.State = adStateOpen Then rsF.Close
'rsF.Open "select description,varietyid  from tblqmsplantvariety order by convert(varietyid,unsigned integer)", db
'Set cbovariety.RowSource = rsF
'cbovariety.ListField = "description"
'cbovariety.BoundColumn = "varietyid"
'
'Set cbovariety.RowSource = rsF
'cbovariety.ListField = "description"
'cbovariety.BoundColumn = "varietyid"

End Sub

Private Sub cbovariety_GotFocus()
If Operation = "" Then
cbovariety.Locked = False
cbovariety.Enabled = True




End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cboVariety_LostFocus()
If Operation = "" Then
chkvariety.Value = 1
chkall.Value = 0
If chkgrp.Value = 0 Then
FillGrid "01/01/0000", "01/01/2020", shortby
Else
fillgridgrp "01/01/0000", "01/01/2020", shortby
End If
End If
cbovariety.Enabled = False
cbovariety.Locked = True
End Sub

Private Sub chkall_Click()
If chkall.Value = 1 Then
chkfacility.Enabled = False
chkplantbatch.Enabled = False
chktransaction.Enabled = False
chkvariety.Enabled = False
chkfacility.Value = 0
chkplantbatch.Value = 0
chktransaction.Value = 0
chkvariety.Value = 0
Else
chkfacility.Enabled = True
chkplantbatch.Enabled = True
chktransaction.Enabled = True
chkvariety.Enabled = True
End If



End Sub

Private Sub chkfacility_Click()
If chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
chkall.Value = 1
End If
End Sub

Private Sub chkplantbatch_Click()
If chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
chkall.Value = 1
End If
End Sub

Private Sub chktransaction_Click()
If chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
chkall.Value = 1
End If
End Sub

Private Sub cmdcancelcr_Click()
framecr.Visible = False
framedr.Visible = False
For i = 1 To mgridcr.Rows - 1
mgridcr.TextMatrix(i, 1) = ""
Next
txtcredit.Text = 0
End Sub

Private Sub cmdcanceldr_Click()
framecr.Visible = False
framedr.Visible = False
For i = 1 To mgriddr.Rows - 1
mgriddr.TextMatrix(i, 1) = ""
Next

txtdebit.Text = 0
End Sub

Private Sub cmdokcr_Click()
Dim tmpcr As Double
Dim i As Integer
tmpcr = 0
For i = 1 To mgridcr.Rows - 1
tmpcr = tmpcr + Val(mgridcr.TextMatrix(i, 1))
Next

If tmpcr = 0 Then
MsgBox "Invalid quantity."
framecr.Visible = True
Exit Sub
Else
If Val(txtdebit.Text) = 0 Then
txtcredit.Text = tmpcr
Else
txtcredit.Text = 0
End If
framecr.Visible = False
End If
End Sub

Private Sub cmdokdr_Click()
Dim tmpdr As Double
Dim i As Integer
tmpdr = 0
For i = 1 To mgriddr.Rows - 1
tmpdr = tmpdr + Val(mgriddr.TextMatrix(i, 1))
Next

If tmpdr = 0 Then
MsgBox "Invalid quantity."
framedr.Visible = True
Exit Sub
Else
If Val(txtcredit.Text) = 0 Then
txtdebit.Text = tmpdr
Else
txtdebit.Text = 0
End If
framedr.Visible = False
End If
End Sub

Private Sub Command1_Click()
Frame5.Visible = False
End Sub

Private Sub Command2_Click()
If chkgrp.Value = 0 Then
FillGrid txtfrmdate.Value, txttodate.Value, shortby
Else
fillgridgrp txtfrmdate.Value, txttodate.Value, shortby
End If
End Sub
Private Sub filltrncombo()
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnId as description  from tblqmsplanttransaction where status='ON' order by trnId", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "description"

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
'On Error GoTo err
Operation = ""
Dim rsF As New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

shortby = "A"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select trnId as description  from tblqmsplanttransaction where status='ON'  order by trnId", db
Set cbotrnid.RowSource = rsF
cbotrnid.ListField = "description"
cbotrnid.BoundColumn = "description"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select concat(facilityid,'  ',description) as description,facilityid  from tblqmsfacility order by facilityid", db
Set cbofacilityid.RowSource = rsF
cbofacilityid.ListField = "description"
cbofacilityid.BoundColumn = "facilityid"


Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select description,varietyid  from tblqmsplantvariety order by convert(varietyid,unsigned integer)", db
Set cbovariety.RowSource = rsF
cbovariety.ListField = "description"
cbovariety.BoundColumn = "varietyid"

Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
rsF.Open "SELECT plantbatch FROM  `tblqmsplantbatchdetail` where plantbatch>0 and  plantbatch not in (select plantbatch from tblqmsplanttransaction group by plantbatch having sum(debit-credit)=0)order by plantbatch", db
Set cboplantBatch.RowSource = rsF
cboplantBatch.ListField = "plantbatch"
cboplantBatch.BoundColumn = "plantbatch"


Set rsF = Nothing
If rsF.State = adStateOpen Then rsF.Close
If rsF.State = adStateOpen Then Srs.Close
rsF.Open "select concat(STAFFCODE , '  ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff WHERE nursery='1'  order by STAFFCODE", db
Set cbostaff.RowSource = rsF
cbostaff.ListField = "STAFFNAME"
cbostaff.BoundColumn = "STAFFCODE"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select  description,verificationid  from tblqmsverificationtype order by convert(verificationid,unsigned integer)", db
Set cboverification.RowSource = rsF
cboverification.ListField = "description"
cboverification.BoundColumn = "verificationid"

Set rsF = Nothing

If rsF.State = adStateOpen Then rsF.Close
rsF.Open "select  description,transitionid  from tblqmstransitiontype order by convert(transitionid,unsigned integer)", db
Set cbotransaction.RowSource = rsF
cbotransaction.ListField = "description"
cbotransaction.BoundColumn = "transitionid"

'FillGrid
'fillsummary
nursaryinventory
txtfrmdate.Value = Format(Now, "dd/MM/yyyy")
txttodate.Value = Format(Now, "dd/MM/yyyy")

Set rsF = Nothing

rsF.Open "select distinct housetype from tblqmsfacility  Order by housetype", MHVDB, adOpenStatic
With rsF
Do While Not .EOF

Select Case rsF!housetype
            Case "C"
            DZLIST.AddItem "Cold House" + " | " + Trim(!housetype)
            Case "N"
            DZLIST.AddItem "Net House" + " | " + Trim(!housetype)
            Case "H"
            DZLIST.AddItem "Hoop House" + " | " + Trim(!housetype)
            Case "T"
            DZLIST.AddItem "Terrace" + " | " + Trim(!housetype)
            Case "S"
            DZLIST.AddItem "Staging House" + " | " + Trim(!housetype)
End Select

   .MoveNext
Loop
End With



For i = 0 To DZLIST.ListCount - 1
DZLIST.Selected(i) = True

Next


Exit Sub
'err:
'MsgBox err.Description
End Sub
Private Sub nursaryinventory()
Dim SQLSTR As String
Dim i As Integer
Dim col, row As Integer
Dim rs As New ADODB.Recordset
Dim col1, col2, col3, col4, col5 As Long
Dim rs1 As New ADODB.Recordset
plantinstock = 0



mygrid1.Rows = 1
mygrid1.ColWidth(0) = 540
mygrid1.ColWidth(1) = 870
mygrid1.ColWidth(2) = 900
mygrid1.ColWidth(3) = 1005
mygrid1.ColWidth(4) = 1005
mygrid1.ColWidth(5) = 1005
mygrid1.ColWidth(6) = 195








Set rs = Nothing
i = 1
rs.Open "select * from tblqmsplantvariety where status='ON'  order by varietyid", MHVDB
'
Do While rs.EOF <> True
mygrid1.Rows = mygrid1.Rows + 1
mygrid1.TextMatrix(i, 0) = rs!varietyId
i = i + 1
rs.MoveNext
Loop
col = 1
row = 1
Dim r As Integer
Set rs = Nothing
rs.Open "select distinct housetype from tblqmsfacility where housetype in('S','N','H','T','C') ", MHVDB
Do While rs.EOF <> True
'col 1 for s,2 for net 3 for hoop
Set rs1 = Nothing
row = 1
mygrid1.TextMatrix(0, col) = rs!housetype

rs1.Open "select varietyid,sum(debit-credit) as stock from tblqmsplanttransaction where   status<>'C'  and facilityid in(select facilityid from tblqmsfacility where housetype ='" & rs!housetype & "') group by varietyid order by varietyid ", MHVDB
Do While rs1.EOF <> True
For row = 1 To mygrid1.Rows - 1
If mygrid1.TextMatrix(row, 0) = rs1!varietyId Then
mygrid1.TextMatrix(row, col) = rs1!stock
plantinstock = plantinstock + rs1!stock
Else

End If

Next

row = row + 1
rs1.MoveNext

Loop

col = col + 1
rs.MoveNext

Loop



            
             col1 = 0
             col2 = 0
             col3 = 0
             col4 = 0
             col5 = 0
            
            col = 1


                For i = 1 To mygrid1.Rows - 1

                If (Len(mygrid1.TextMatrix(i, 0))) = 0 Then Exit For
                col = 1
'                FindqmsPlantVariety CInt(Mygrid.TextMatrix(i, 0))
              
               For j = 1 To 5

                Select Case mygrid1.TextMatrix(0, col)

                Case "H"
                mygrid1.TextMatrix(0, col) = "Hoop" 'Mygrid.TextMatrix(0, col)
                Case "N"
                mygrid1.TextMatrix(0, col) = "Net" 'Mygrid.TextMatrix(0, col)
                Case "S"
                 mygrid1.TextMatrix(0, col) = "Stagging" 'Mygrid.TextMatrix(0, col)
                Case "T"
                 mygrid1.TextMatrix(0, col) = "NGT" 'Mygrid.TextMatrix(0, col)
                 Case "C"
                 mygrid1.TextMatrix(0, col) = "Cold" 'Mygrid.TextMatrix(0, col)


                End Select
               
                

                
'                xl.Cells(54 + i, 31 + j) = Mygrid.TextMatrix(i, j)
                col = col + 1
               Next
               
                  col1 = col1 + Val(mygrid1.TextMatrix(i, 1))
                  col2 = col2 + Val(mygrid1.TextMatrix(i, 2))
                  col3 = col3 + Val(mygrid1.TextMatrix(i, 3))
                  col4 = col4 + Val(mygrid1.TextMatrix(i, 4))
                  col5 = col5 + Val(mygrid1.TextMatrix(i, 5))
               
                FindqmsPlantVariety CInt(mygrid1.TextMatrix(i, 0))
                mygrid1.TextMatrix(i, 0) = qmsPlantVariety
                Next
                txtsumcredit.Text = Format(plantinstock, "###,###")
                txtcol1.Text = Format(col1, "###,###")
                txtcol2.Text = Format(col2, "###,###")
                txtcol3.Text = Format(col3, "###,###")
                txtcol4.Text = Format(col4, "###,###")
                txtcol5.Text = Format(col5, "###,###")

End Sub
Private Sub fillsummary()
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid1.Clear
sumdr = 0
sumcr = 0
mygrid1.Rows = 1
mygrid1.FormatString = "^Sl.No.|^TransactionType|^Debit|^Credit|^"
mygrid1.ColWidth(0) = 645
mygrid1.ColWidth(1) = 2910

mygrid1.ColWidth(2) = 900
mygrid1.ColWidth(3) = 960
mygrid1.ColWidth(4) = 225

rs.Open "select transactiontype,sum(debit)as dr,sum(credit) as cr from tblqmsplanttransaction where status='ON' group by transactiontype order by transactiontype", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid1.Rows = mygrid1.Rows + 1
mygrid1.TextMatrix(i, 0) = i
Findqmstransactiontype rs!transactiontype

mygrid1.TextMatrix(i, 1) = qmsTransactionType
mygrid1.ColAlignment(1) = flexAlignLeftTop

mygrid1.TextMatrix(i, 2) = IIf(rs!dr = 0, "", rs!dr)
mygrid1.ColAlignment(2) = flexAlignRightTop
mygrid1.TextMatrix(i, 3) = IIf(rs!cr = 0, "", rs!cr)
mygrid1.ColAlignment(3) = flexAlignRightTop

sumdr = sumdr + IIf(IsNull(rs!dr), 0, rs!dr)
sumcr = sumcr + IIf(IsNull(rs!cr), 0, rs!cr)
rs.MoveNext
i = i + 1
Loop
txtsumdebit.Text = sumdr
txtsumcredit.Text = sumcr
rs.Close
Exit Sub
err:
MsgBox err.Description

End Sub
Private Sub MNU_SAVE()
Dim mdr As Double
Dim mcr As Double
Dim i As Integer
Dim rs As New ADODB.Recordset
On Error GoTo err
If Len(cbofacilityid.Text) = 0 Then
MsgBox "Select facility."
Exit Sub
End If

If Len(txtloc.Text) = 0 Then
MsgBox "Select Station."
Exit Sub
End If

If Not IsNumeric(cbovariety.BoundText) Then
MsgBox "Select The variety once!"
cbovariety.Enabled = True
cbovariety.Locked = False
Exit Sub
End If

'For i = 1 To mgriddr.Rows - 1
'mdr = mdr + Val(mgriddr.TextMatrix(i, 1))
'mcr = mcr + Val(mgridcr.TextMatrix(i, 1))
'Next
'
'
'If Mid(Trim(cbovariety.Text), 1, 1) = "N" Then
'
'If mdr <> Val(txtdebit.Text) And mcr = Val(txtcredit.Text) Then
'MsgBox "Nut quantity not valid!"
'Exit Sub
'End If
'
'If mdr = 0 And mcr = 0 Then
'MsgBox "Nut quantity not valid!"
'Exit Sub
'End If
'
'
'End If


If cbotransaction.BoundText = 4 Or cbotransaction.BoundText = 5 Or cbotransaction.BoundText = 14 Then
MsgBox "You cannot proceed to save this transaction."

Exit Sub

End If

If Len(cboplantBatch.Text) = 0 Then
MsgBox "Select Plant Batch."
Exit Sub
End If

'If Val(txtdebit.Text) = 0 And Val(txtcredit.Text) = 0 Then
'MsgBox "Invalid debit/credit input."
'Exit Sub
'End If
If Len(cbotrnid.Text) = 0 Then
MsgBox "Transaction No. is Must."
Exit Sub
End If

MHVDB.BeginTrans
If Operation = "ADD" Then
LogRemarks = ""
MHVDB.Execute "INSERT INTO tblqmsplanttransaction (trnid,entrydate,plantbatch,varietyid," _
            & "facilityid,verificationtype,transactiontype,staffid,debit,credit,status,location) " _
            & "VALUEs('" & cbotrnid.Text & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & cboplantBatch.BoundText & "', " _
            & "'" & cbovariety.BoundText & "','" & cbofacilityid.BoundText & "'," _
            & "'" & cboverification.BoundText & "','" & cbotransaction.BoundText & "', " _
            & "'" & cbostaff.BoundText & "','" & Val(txtdebit.Text) & "','" & Val(txtcredit.Text) & "','ON','" & txtloc.Text & "')"
 
 
 If Val(txtdebit.Text) > 0 Then
 
 For i = 1 To mgriddr.Rows - 1
 
 If Val(mgriddr.TextMatrix(i, 1)) > 0 Then
 
 
 MHVDB.Execute "INSERT INTO tblqmsnutdetail (trnid,entrydate,plantbatch,varietyid," _
            & "facilityid,nuttype,transactiontype,nutdebit,nutcredit,status,location) " _
            & "VALUEs('" & cbotrnid.Text & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & cboplantBatch.BoundText & "', " _
            & "'" & cbovariety.BoundText & "','" & cbofacilityid.BoundText & "'," _
            & "'" & Trim(mgriddr.TextMatrix(i, 2)) & "','" & cbotransaction.BoundText & "', " _
            & "'" & Val(mgriddr.TextMatrix(i, 1)) & "','0','ON','" & txtloc.Text & "')"
            
 
 
 End If
 
 
 
 Next
 
 Else
 
   For i = 1 To mgridcr.Rows - 1
 
 If Val(mgridcr.TextMatrix(i, 1)) > 0 Then
 
 
 MHVDB.Execute "INSERT INTO tblqmsnutdetail (trnid,entrydate,plantbatch,varietyid," _
            & "facilityid,nuttype,transactiontype,nutdebit,nutcredit,status,location) " _
            & "VALUEs('" & cbotrnid.Text & "','" & Format(txtentrydate.Value, "yyyy-MM-dd") & "','" & cboplantBatch.BoundText & "', " _
            & "'" & cbovariety.BoundText & "','" & cbofacilityid.BoundText & "'," _
            & "'" & Trim(mgridcr.TextMatrix(i, 2)) & "','" & cbotransaction.BoundText & "', " _
            & "'0','" & Val(mgridcr.TextMatrix(i, 1)) & "','ON','" & txtloc.Text & "')"
            
 
 
 End If
 
 
 
 Next
 
 End If
 
 
LogRemarks = "Inserted new record" & cbotrnid.BoundText & "," & cboplantBatch.Text & "," & Mlocation
updatemhvlog Now, MUSER, LogRemarks, ""

ElseIf Operation = "OPEN" Then
MHVDB.Execute "update tblqmsplanttransaction set entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "', " _
            & "facilityid='" & cbofacilityid.BoundText & "',plantbatch='" & cboplantBatch.BoundText & "', " _
            & "debit='" & Val(txtdebit.Text) & "',credit='" & Val(txtcredit.Text) & "', " _
            & "varietyid='" & cbovariety.BoundText & "',verificationtype='" & cboverification.BoundText & "'," _
            & "transactiontype='" & cbotransaction.BoundText & "',staffid='" & cbostaff.BoundText & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and location='" & txtloc.Text & "'"


 If Val(txtdebit.Text) > 0 Then
 
 For i = 1 To mgriddr.Rows - 1
 
 If Val(mgriddr.TextMatrix(i, 1)) > 0 Then
 
 
 MHVDB.Execute "update tblqmsnutdetail  set entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "', " _
            & "facilityid='" & cbofacilityid.BoundText & "',plantbatch='" & cboplantBatch.BoundText & "', " _
            & "nutdebit='" & Val(mgriddr.TextMatrix(i, 1)) & "',nutcredit='0', " _
            & "varietyid='" & cbovariety.BoundText & "',nuttype='" & Trim(mgriddr.TextMatrix(i, 2)) & "'," _
            & "transactiontype='" & cbotransaction.BoundText & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and   location='" & txtloc.Text & "'"
            
 
 
 End If
 
 
 
 Next
 
 Else
 
   For i = 1 To mgridcr.Rows - 1
 
 If Val(mgridcr.TextMatrix(i, 1)) > 0 Then
 
 
 MHVDB.Execute "update tblqmsnutdetail  set entrydate='" & Format(txtentrydate.Value, "yyyy-MM-dd") & "', " _
            & "facilityid='" & cbofacilityid.BoundText & "',plantbatch='" & cboplantBatch.BoundText & "', " _
            & "nutdebit='0',nutcredit='" & Val(mgridcr.TextMatrix(i, 1)) & "', " _
            & "varietyid='" & cbovariety.BoundText & "',nuttype='" & Trim(mgridcr.TextMatrix(i, 2)) & "'," _
            & "transactiontype='" & cbotransaction.BoundText & "'" _
            & " where trnid='" & cbotrnid.BoundText & "' and  location='" & txtloc.Text & "'"
            
 
 
 End If
 
 
 
 Next
 
 End If


LogRemarks = "Updated  record" & cbotrnid.BoundText & "," & cboplantBatch.Text & "," & txtloc.Text
updatemhvlog Now, MUSER, LogRemarks, ""
Else
MsgBox "OPERATION NOT SELECTED."
End If

chkplantbatch.Value = 1
chkfacility.Value = 1
FillGrid "01/01/0000", "01/01/2020", shortby
Operation = ""
cbovariety.Enabled = True
MHVDB.CommitTrans
  TB.Buttons(3).Enabled = False
  TB.Buttons(4).Enabled = False
  filltrncombo
Exit Sub

err:
MsgBox err.Description
TB.Buttons(3).Enabled = False
MHVDB.RollbackTrans


End Sub

Private Sub FillGrid(frmdate As Date, todate As Date, shortby As String)
On Error GoTo err
gendr = 0
gencr = 0
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Trn. Id|^Date|^PBID|^ PVID|^Facility|^Verification Type|^Transaction Type|^Debit|^Credit|^"
Mygrid.ColWidth(0) = 645
Mygrid.ColWidth(1) = 735
Mygrid.ColWidth(2) = 1140
Mygrid.ColWidth(3) = 510
Mygrid.ColWidth(4) = 585
Mygrid.ColWidth(5) = 2025
Mygrid.ColWidth(6) = 1875
Mygrid.ColWidth(7) = 2460
Mygrid.ColWidth(8) = 855
Mygrid.ColWidth(9) = 855
Mygrid.ColWidth(10) = 180


Dzstr = ""



For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE FACILITY TYPE."
          Exit Sub
       End If
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
   MsgBox "FACILITY NOT SELECTED !!!"
   Exit Sub
End If


'If shortby = "A" Then
'rs.Open "select * from tblqmsplanttransaction where status='ON' and entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'ElseIf shortby = "F" Then
'If Len(cbofacilityId.Text) = 0 Then
'MsgBox "Select Facility From The Drop Down Combo."
'Exit Sub
'End If
'rs.Open "select * from tblqmsplanttransaction where status='ON' and  entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityId.BoundText & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'
'ElseIf shortby = "B" Then
'If Len(cboplantBatch.Text) = 0 Then
'MsgBox "Select Plant Batch From The Drop Down Combo."
'Exit Sub
'End If
'rs.Open "select * from tblqmsplanttransaction where status='ON' and  entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and plantbatch='" & cboplantBatch.BoundText & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'
'ElseIf shortby = "T" Then
'If Len(cbotransaction.Text) = 0 Then
'MsgBox "Select Transaction Type From The Drop Down Combo."
'Exit Sub
'End If
'rs.Open "select * from tblqmsplanttransaction where status='ON' and  entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and transactiontype='" & cbotransaction.BoundText & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'
'Else
'MsgBox "restart The process again."
'Exit Sub
'End If
If chkall.Value = 1 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and  status<>'C' and entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 1 And chktransaction.Value = 1 And chkvariety.Value Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' and plantbatch='" & cboplantBatch.BoundText & "' and transactiontype='" & cbotransaction.BoundText & "' order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 1 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' and plantbatch='" & cboplantBatch.BoundText & "' order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "'  order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 0 And chkvariety.Value = 0 And chktransaction.Value = 1 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' and  transactiontype='" & cbotransaction.BoundText & "' order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 0 And chkvariety.Value = 0 And chktransaction.Value = 1 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and  transactiontype='" & cbotransaction.BoundText & "' order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 1 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "'  and plantbatch='" & cboplantBatch.BoundText & "'  order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 1 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "'  and varietyid='" & cbovariety.BoundText & "'  order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 1 And chkvariety.Value = 1 Then
rs.Open "select * from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and transactiontype='" & cbotransaction.BoundText & "' and varietyid='" & cbovariety.BoundText & "'  order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
Else
MsgBox "restart The process again."
Exit Sub
End If

i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i
Mygrid.TextMatrix(i, 1) = rs!trnid
Mygrid.TextMatrix(i, 2) = Format(rs!entrydate, "dd/MM/yyyy")
Mygrid.TextMatrix(i, 3) = rs!plantBatch
FindqmsPlantVariety rs!varietyId
Mygrid.TextMatrix(i, 4) = qmsPlantVariety
findQmsfacility UCase(rs!facilityid)
Mygrid.TextMatrix(i, 5) = rs!facilityid & " " & qmsFacility
Findqmsverificationtype rs!verificationType
Mygrid.TextMatrix(i, 6) = qmsVerificationType
Mygrid.ColAlignment(6) = flexAlignLeftTop
Findqmstransactiontype rs!transactiontype
Mygrid.TextMatrix(i, 7) = qmsTransactionType
Mygrid.ColAlignment(7) = flexAlignLeftTop

Mygrid.TextMatrix(i, 8) = IIf(rs!debit = 0, "", rs!debit)
Mygrid.ColAlignment(8) = flexAlignRightTop
Mygrid.TextMatrix(i, 9) = IIf(rs!credit = 0, "", rs!credit)
Mygrid.ColAlignment(9) = flexAlignRightTop
gendr = gendr + IIf(IsNull(rs!debit), 0, rs!debit)
gencr = gencr + IIf(IsNull(rs!credit), 0, rs!credit)
rs.MoveNext
i = i + 1
Loop
txtgendebit.Text = Format(gendr, "######")
txtgencredit.Text = Format(gencr, "######")

rs.Close


txttotalplant.Text = Val(txtgendebit.Text) - Val(txtgencredit.Text)
txttotalplant.Text = IIf(Val(txttotalplant.Text) = 0, "", txttotalplant.Text)
txtgendebit.Text = Format(txtgendebit.Text, "###,###")
txtgencredit.Text = Format(txtgencredit.Text, "###,###")
txttotalplant.Text = Format(txttotalplant.Text, "###,###")
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub mgridcr_Click()
If mgridcr.col = 1 And mgridcr.row > 0 Then
mgridcr.Editable = flexEDKbdMouse
Else
mgridcr.Editable = flexEDNone
End If
End Sub

Private Sub mgriddr_Click()
If mgriddr.col = 1 And mgriddr.row > 0 Then
mgriddr.Editable = flexEDKbdMouse
Else
mgriddr.Editable = flexEDNone
End If
End Sub

Private Sub mygrid_DblClick()
If Mygrid.col = 3 And Val(Mygrid.TextMatrix(Mygrid.row, 3)) > 0 Then
Frame5.Visible = True
txtshipmetno.Text = ""
txtrcvdate.Text = ""
txtnoofplants.Text = ""
txtplantvariety.Text = ""

Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblqmsplantbatchdetail where plantbatch='" & Val(Mygrid.TextMatrix(Mygrid.row, 3)) & "'", MHVDB
If rs.EOF <> True Then
txtshipmetno.Text = rs!trnid
txtrcvdate.Text = Format(rs!entrydate, "dd/MM/yyyy")
txtnoofplants.Text = rs!shipmentsize
FindqmsPlantVariety rs!plantvariety
txtplantvariety.Text = qmsPlantVariety
End If

fillhstgrid Val(Mygrid.TextMatrix(Mygrid.row, 3))

End If
End Sub
Private Sub fillhstgrid(batchno As Integer)
On Error GoTo err
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
mygrid2.Clear
sumdr = 0
sumcr = 0
mygrid2.Rows = 1
mygrid2.FormatString = "^Sl.No.|^Date|^TransactionType|^Debit|^Credit|^"
mygrid2.ColWidth(0) = 645
mygrid2.ColWidth(1) = 1140

mygrid2.ColWidth(2) = 2475
mygrid2.ColWidth(3) = 900
mygrid2.ColWidth(4) = 960
mygrid2.ColWidth(5) = 150
rs.Open "select * from tblqmsplanttransaction where status='ON' and plantbatch='" & batchno & "' order by entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
i = 1
Do While rs.EOF <> True
mygrid2.Rows = mygrid2.Rows + 1
mygrid2.TextMatrix(i, 0) = i


mygrid2.TextMatrix(i, 1) = Format(rs!entrydate, "dd/MM/yyyy")

Findqmstransactiontype rs!transactiontype
mygrid2.TextMatrix(i, 2) = qmsTransactionType
mygrid2.TextMatrix(i, 3) = IIf(rs!debit = 0, "", rs!debit)
mygrid2.ColAlignment(3) = flexAlignRightTop
mygrid2.TextMatrix(i, 4) = IIf(rs!credit = 0, "", rs!credit)
mygrid2.ColAlignment(4) = flexAlignRightTop

sumdr = sumdr + IIf(IsNull(rs!debit), 0, rs!debit)
sumcr = sumcr + IIf(IsNull(rs!credit), 0, rs!credit)
rs.MoveNext
i = i + 1
Loop
txtdr.Text = sumdr
txtcr.Text = sumcr
txttot.Text = Val(sumdr - sumcr)
rs.Close
Exit Sub
err:
MsgBox err.Description


End Sub

Private Sub OPTALL_Click()

End Sub

Private Sub optfacility_Click()

End Sub

Private Sub Option1_Click()

End Sub

Private Sub optplantbatch_Click()

End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo err

Select Case Button.Key

       Case "ADD"

            cbovariety.Locked = True
            cbovariety.Enabled = False
            chkvariety.Value = 0
        cbotrnid.Enabled = False
        TB.Buttons(3).Enabled = True
        Operation = "ADD"
        CLEARCONTROLL
        Dim rs As New ADODB.Recordset
        Set rs = Nothing
        rs.Open "SELECT MAX(trnid)+1 AS MaxID from tblqmsplanttransaction", MHVDB, adOpenForwardOnly, adLockOptimistic
        If rs.EOF <> True Then
        cbotrnid.Text = IIf(IsNull(rs!MaxId), "1", rs!MaxId)
        Else
        cbotrnid.Text = rs!MaxId
        End If
       Case "OPEN"

            cbovariety.Locked = True
            cbovariety.Enabled = False
            chkvariety.Value = 0
        Operation = "OPEN"
        CLEARCONTROLL
        cbotrnid.Enabled = True
        TB.Buttons(3).Enabled = True
        TB.Buttons(4).Enabled = True
             
       Case "SAVE"
        MNU_SAVE
        'FillGrid
        nursaryinventory
       Case "DELETE"
         MNU_CANCEL
          nursaryinventory
       Case "EXIT"
       Unload Me
       
       
End Select
Exit Sub
err:
MsgBox err.Description
End Sub
Private Sub MNU_CANCEL()
If MsgBox("Do You Want to Cancel The Reccord No. " & cbotrnid.Text & "?", vbQuestion + vbYesNo) = vbYes Then
MHVDB.Execute "update tblqmsplanttransaction set status='C' where trnid='" & cbotrnid.BoundText & "'"
 TB.Buttons(3).Enabled = False
TB.Buttons(4).Enabled = False

LogRemarks = "Canceled Transaction No." & cbotrnid.BoundText & "," & "from table tblqmsplanttransaction"
updatemhvlog Now, MUSER, LogRemarks, ""
filltrncombo

Else

End If
End Sub

Private Sub CLEARCONTROLL()
    cbofacilityid.Text = ""
   cboplantBatch.Text = ""
     txtdebit.Text = ""
   txtcredit.Text = ""
   cboverification.Text = ""
cbotransaction.Text = ""
cbostaff.Text = ""
cbovariety.Text = ""
txtloc.Text = ""
txtentrydate.Value = Format(Now, "dd/MM/yyyy")
For i = 1 To mgriddr.Rows - 1
mgriddr.TextMatrix(i, 1) = ""
mgridcr.TextMatrix(i, 1) = ""
Next
framedr.Visible = False
framecr.Visible = False
End Sub

Private Sub txtcredit_GotFocus()
'If Mid(Trim(cbovariety.Text), 1, 1) = "N" Then
'txtcredit.Text = ""
'txtcredit.Locked = True
'framecr.Visible = True
'framedr.Visible = False
'
'framecr.Top = 2520
'framecr.Left = 2160
'End If
End Sub

Private Sub txtcredit_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtdebit_GotFocus()
'If Mid(Trim(cbovariety.Text), 1, 1) = "N" Then
'txtdebit.Text = ""
'txtdebit.Locked = True
'framedr.Visible = True
'framecr.Visible = False
'
'framedr.Top = 2520
'framedr.Left = 2160
'End If

End Sub

Private Sub txtdebit_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

Private Sub fillgridgrp(frmdate As Date, todate As Date, shortby As String)
On Error GoTo err
Dim grpstr As String
gendr = 0
gencr = 0
Dim rs As New ADODB.Recordset
Dim i As Integer
Set rs = Nothing
Mygrid.Clear
Mygrid.Rows = 1
Mygrid.FormatString = "^Sl.No.|^Trn. Id|^Date|^PBID|^ PVID|^Facility|^Verification Type|^Transaction Type|^Debit|^Credit|^"
Mygrid.ColWidth(0) = 645
Mygrid.ColWidth(1) = 735
Mygrid.ColWidth(2) = 1140
Mygrid.ColWidth(3) = 510
Mygrid.ColWidth(4) = 585
Mygrid.ColWidth(5) = 2025
Mygrid.ColWidth(6) = 1875
Mygrid.ColWidth(7) = 2460
Mygrid.ColWidth(8) = 855
Mygrid.ColWidth(9) = 855
Mygrid.ColWidth(10) = 180


Dzstr = ""



For i = 0 To DZLIST.ListCount - 1
    If DZLIST.Selected(i) Then
       Dzstr = Dzstr + "'" + Trim(Mid(DZLIST.List(i), InStr(1, DZLIST.List(i), "|") + 1)) + "',"
       Mcat = DZLIST.List(i)
       j = j + 1
    End If
    If RepName = "5" Then
       If j > 1 Then
          MsgBox "SELECT ATLEAST ONE FACILITY TYPE."
          Exit Sub
       End If
    End If
Next
If Len(Dzstr) > 0 Then
   Dzstr = "(" + Left(Dzstr, Len(Dzstr) - 1) + ")"
 
Else
   MsgBox "FACILITY NOT SELECTED !!!"
   Exit Sub
End If


'If shortby = "A" Then
'rs.Open "select * from tblqmsplanttransaction where status='ON' and entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'ElseIf shortby = "F" Then
'If Len(cbofacilityId.Text) = 0 Then
'MsgBox "Select Facility From The Drop Down Combo."
'Exit Sub
'End If
'rs.Open "select * from tblqmsplanttransaction where status='ON' and  entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityId.BoundText & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'
'ElseIf shortby = "B" Then
'If Len(cboplantBatch.Text) = 0 Then
'MsgBox "Select Plant Batch From The Drop Down Combo."
'Exit Sub
'End If
'rs.Open "select * from tblqmsplanttransaction where status='ON' and  entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and plantbatch='" & cboplantBatch.BoundText & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'
'ElseIf shortby = "T" Then
'If Len(cbotransaction.Text) = 0 Then
'MsgBox "Select Transaction Type From The Drop Down Combo."
'Exit Sub
'End If
'rs.Open "select * from tblqmsplanttransaction where status='ON' and  entrydate >='" & Format(txtfrmdate.Value, "yyyy-MM-dd") & "' and  entrydate<='" & Format(txttodate.Value, "yyyy-MM-dd") & "' and transactiontype='" & cbotransaction.BoundText & "' order by  entrydate", MHVDB, adOpenForwardOnly, adLockOptimistic
'
'Else
'MsgBox "restart The process again."
'Exit Sub
'End If
'grpstr = "group by plantbatch,varietyid,facilityid,verificationtype,transactiontype"
If chkall.Value = 1 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit " _
& "from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where " _
& " housetype in " & Dzstr & ") and  status<>'C' and " _
& " entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  " _
& " entrydate<='" & Format(todate, "yyyy-MM-dd") & "' group by entrydate,plantbatch,varietyid,facilityid, " _
& " verificationtype,transactiontype order by " _
& " entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic



ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 1 And chktransaction.Value = 1 And chkvariety.Value Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' and plantbatch='" & cboplantBatch.BoundText & "' and transactiontype='" & cbotransaction.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 1 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' and plantbatch='" & cboplantBatch.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype  order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 1 And chkplantbatch.Value = 0 And chkvariety.Value = 0 And chktransaction.Value = 1 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and facilityid='" & cbofacilityid.BoundText & "' and  transactiontype='" & cbotransaction.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 0 And chkvariety.Value = 0 And chktransaction.Value = 1 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and  transactiontype='" & cbotransaction.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 1 And chktransaction.Value = 0 And chkvariety.Value = 0 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "'  and plantbatch='" & cboplantBatch.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 0 And chkvariety.Value = 1 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "'  and varietyid='" & cbovariety.BoundText & "'  group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
ElseIf chkfacility.Value = 0 And chkplantbatch.Value = 0 And chktransaction.Value = 1 And chkvariety.Value = 1 Then
rs.Open "select trnid,entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype,sum(debit) debit,sum(credit) credit from tblqmsplanttransaction where facilityid in(select facilityid from tblqmsfacility where housetype in " & Dzstr & ") and status<>'C' and  entrydate >='" & Format(frmdate, "yyyy-MM-dd") & "' and  entrydate<='" & Format(todate, "yyyy-MM-dd") & "' and transactiontype='" & cbotransaction.BoundText & "' and varietyid='" & cbovariety.BoundText & "' group by entrydate,plantbatch,varietyid,facilityid,verificationtype,transactiontype order by  entrydate desc", MHVDB, adOpenForwardOnly, adLockOptimistic
Else
MsgBox "restart The process again."
Exit Sub
End If

i = 1
Do While rs.EOF <> True
Mygrid.Rows = Mygrid.Rows + 1
Mygrid.TextMatrix(i, 0) = i
Mygrid.TextMatrix(i, 1) = rs!trnid
Mygrid.TextMatrix(i, 2) = Format(rs!entrydate, "dd/MM/yyyy")
Mygrid.TextMatrix(i, 3) = rs!plantBatch
FindqmsPlantVariety rs!varietyId
Mygrid.TextMatrix(i, 4) = qmsPlantVariety
findQmsfacility UCase(rs!facilityid)
Mygrid.TextMatrix(i, 5) = rs!facilityid & " " & qmsFacility
Findqmsverificationtype rs!verificationType
Mygrid.TextMatrix(i, 6) = qmsVerificationType
Mygrid.ColAlignment(6) = flexAlignLeftTop
Findqmstransactiontype rs!transactiontype
Mygrid.TextMatrix(i, 7) = qmsTransactionType
Mygrid.ColAlignment(7) = flexAlignLeftTop

Mygrid.TextMatrix(i, 8) = IIf(rs!debit = 0, "", rs!debit)
Mygrid.ColAlignment(8) = flexAlignRightTop
Mygrid.TextMatrix(i, 9) = IIf(rs!credit = 0, "", rs!credit)
Mygrid.ColAlignment(9) = flexAlignRightTop
gendr = gendr + IIf(IsNull(rs!debit), 0, rs!debit)
gencr = gencr + IIf(IsNull(rs!credit), 0, rs!credit)
rs.MoveNext
i = i + 1
Loop
txtgendebit.Text = Format(gendr, "######")
txtgencredit.Text = Format(gencr, "######")

rs.Close


txttotalplant.Text = Val(txtgendebit.Text) - Val(txtgencredit.Text)
txttotalplant.Text = IIf(Val(txttotalplant.Text) = 0, "", txttotalplant.Text)
txtgendebit.Text = Format(txtgendebit.Text, "###,###")
txtgencredit.Text = Format(txtgencredit.Text, "###,###")
txttotalplant.Text = Format(txttotalplant.Text, "###,###")
Exit Sub
err:
MsgBox err.Description

End Sub


