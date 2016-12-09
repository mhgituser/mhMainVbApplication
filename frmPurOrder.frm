VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmPOeNTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   6090
   ClientLeft      =   4230
   ClientTop       =   1395
   ClientWidth     =   11835
   Icon            =   "frmPurOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11835
   Visible         =   0   'False
   Begin VB.TextBox txtr2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtr3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtiremarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtracc 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   9120
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtexciseduty 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   9120
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9120
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtChallan 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "Enter Challan No"
      Top             =   1440
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   315
      Index           =   0
      Left            =   6240
      TabIndex        =   1
      ToolTipText     =   "Enter Challan Date"
      Top             =   675
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80543745
      CurrentDate     =   36383
      MinDate         =   36161
   End
   Begin VB.Data DatBrBill 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datInvItem 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   3465
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   8775
      Begin MSDataListLib.DataCombo CboItemDesc 
         Bindings        =   "frmPurOrder.frx":076A
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         OLEDropMode     =   1
         BackColor       =   -2147483643
         ListField       =   "itemcode"
         BoundColumn     =   "ItemCode"
         Text            =   ""
      End
      Begin VB.TextBox cboItemCode 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDBCtls.DBCombo CboItemDesc1 
         Bindings        =   "frmPurOrder.frx":077F
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "itemname"
         BoundColumn     =   "itemcode"
         Text            =   ""
      End
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   3300
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5821
         _Version        =   393216
         Rows            =   201
         Cols            =   7
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         HighLight       =   0
         FormatString    =   $"frmPurOrder.frx":0798
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
         TabIndex        =   10
         Top             =   4920
         Width           =   870
      End
      Begin VB.Line Line2 
         X1              =   5280
         X2              =   8880
         Y1              =   4290
         Y2              =   4290
      End
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   315
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      ToolTipText     =   "Enter Challan Date"
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80543745
      CurrentDate     =   36383
      MinDate         =   36161
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   17
      ToolTipText     =   "Enter Challan Date"
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80543745
      CurrentDate     =   36383
      MinDate         =   36161
   End
   Begin MSDataListLib.DataCombo DBcboParty 
      Bindings        =   "frmPurOrder.frx":0838
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1800
      TabIndex        =   22
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CboBillNo 
      Bindings        =   "frmPurOrder.frx":084D
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1800
      TabIndex        =   23
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cbotransporter 
      Bindings        =   "frmPurOrder.frx":0862
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   7560
      TabIndex        =   25
      Top             =   1080
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
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   10560
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
            Picture         =   "frmPurOrder.frx":0877
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":0C11
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":0FAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":1C85
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":20D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":2891
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":2C2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":2FC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":335F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":4039
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":448B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurOrder.frx":4C45
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      TabIndex        =   36
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      TabIndex        =   35
      Top             =   2280
      Width           =   105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      TabIndex        =   34
      Top             =   1920
      Width           =   105
   End
   Begin VB.Label Nar1 
      AutoSize        =   -1  'True
      Caption         =   "Note1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   30
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Remarks To Accounts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9240
      TabIndex        =   28
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Ex-Duty"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9120
      TabIndex        =   26
      Top             =   3120
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Transporter"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6480
      TabIndex        =   24
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblYr 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1140
      TabIndex        =   20
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   6960
      TabIndex        =   19
      Top             =   5760
      Width           =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Delivery Expected On"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Reference No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ref. Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   13
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Entry Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Entry No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmPOeNTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BillDetailRec As New ADODB.Recordset
Dim rsInvItem As New ADODB.Recordset
Dim rsbrbill As New ADODB.Recordset
Dim rspobill As New ADODB.Recordset
Dim db
Dim Bill As New ADODB.Recordset
'Dim datInvItem As New ADODB.Recordset
Dim CurrRow, Jkey, ErrCTR As Long
Dim ValidRow As Boolean
Dim Operation As String
Dim ltot As Double
Const fmString = "       |^ Code      |^                               Item Name                                   |^  Unit   |^     Qty      |^   Pur. Rate |^      Amount   "

Private Sub CboBillNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub CboBillNo_LostFocus()
Dim i As Integer
Dim Issue, Recv As Double
Dim imast As New ADODB.Recordset
If Operation = "Add" Then Exit Sub
ltot = 0
Set Bill = MHVDB.Execute("select * from PurOrderHdr where procyear='" & SysYear & "' and ((purordno))=('" & CboBillNo & "') AND STATUS <> 'C'")
If Bill.EOF Then
   MsgBox CboBillNo + " Does not exists "
   CboBillNo.SetFocus
   Exit Sub
Else
   With Bill
   'lblnow = Format(!billdate, "dd/mm/yyyy") ' DatKotHead.Recordset!Time
   txtChallan = !refno
   txtDate(0) = !podate
   txtDate(1) = !refdate
   txtDate(2) = !deledate
   txtr2.Text = IIf(IsNull(!r2), "", !r2)
   txtr3.Text = IIf(IsNull(!r3), "", !r3)
   txtracc.Text = IIf(IsNull(!racc), "", !racc)
   txtiremarks.Text = IIf(IsNull(!iremarks), "", !iremarks)
   txtexciseduty.Text = !exduty
   txtRemark = IIf(IsNull(!remarks), "", !remarks)
   
   rsbrbill.Find "SuplCode='" & !suplcode & "'", , adSearchForward, 1
   If Not rsbrbill.EOF Then DBcboParty.Text = rsbrbill!Name
   
    rsbrbill.Find "SuplCode='" & !TRCODE & "'", , adSearchForward, 1
   If Not rsbrbill.EOF Then cbotransporter.Text = rsbrbill!Name
   
   
   
  ' DBcboParty = IIf(IsNull(!SUPLCode), "", !SUPLCode)
   
   
   End With
   Set BillDetailRec = MHVDB.Execute("select d.itemcode,d.qty,d.rate,B.purunit,itemname from PurOrderDtl as d,invitems as b where d.itemcode=b.itemcode and  d.procyear='" & SysYear & "' and d.purordno=('" & CboBillNo & "')")
   ItemGrd.Clear
   ItemGrd.FormatString = fmString
   With BillDetailRec
   i = 1
   Do While Not .EOF
      ItemGrd.TextMatrix(i, 0) = i
      ItemGrd.TextMatrix(i, 1) = !itemcode
      ItemGrd.TextMatrix(i, 2) = !ITEMNAME
      ItemGrd.TextMatrix(i, 3) = !purunit
      ItemGrd.TextMatrix(i, 4) = !qty
      ItemGrd.TextMatrix(i, 5) = Format(!Rate, "####0.00")
      ItemGrd.TextMatrix(i, 6) = !Rate * !qty
      ltot = ltot + !Rate * !qty
      .MoveNext
      i = i + 1
   Loop
   End With
End If
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
Frame2.Enabled = True
TB.Buttons(4).Enabled = True



'TB.Buttons(5).Enabled = True
TB.Buttons(3).Enabled = True
End Sub

Private Sub cboItemCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub cboItemCode_LostFocus()
Dim prevamt, CurrAmt, Jstock As Double
cboItemCode.Text = UCase(cboItemCode.Text)
If ItemGrd.TextMatrix(CurrRow, 1) = cboItemCode Then
   cboItemCode.Visible = False
   Exit Sub
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
'datInvItem.Recordset.FindFirst "trim(itemcode) = trim('" & cboItemCode & "')"
rsInvItem.Find "itemcode = '" & cboItemCode & "'", , adSearchForward, 1
With rsInvItem
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   ItemGrd.TextMatrix(CurrRow, 6) = ""
   CurrAmt = 0
   txtQty = ""
   txtRate = ""
   ValidRow = True
Else
   ItemGrd.TextMatrix(CurrRow, 1) = !itemcode
   ItemGrd.TextMatrix(CurrRow, 2) = !ITEMNAME
   ItemGrd.TextMatrix(CurrRow, 3) = IIf(IsNull(!unit), "Nos", !unit)
   txtQty = ItemGrd.TextMatrix(CurrRow, 4)
   If Not Val(txtQty) > 0 Then
      ValidRow = False
      ItemGrd.row = CurrRow
      txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
      txtQty.Visible = True
      txtQty.SetFocus
   End If

End If
End With
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
CboItemDesc.Visible = False
cboItemCode.Visible = False
End Sub

Private Sub CboItemDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub cboItemDesc_LostFocus()

Dim Jstock, prevamt, CurrAmt As Double
Dim imast As ADODB.Recordset
If ItemGrd.TextMatrix(CurrRow, 1) = CboItemDesc.BoundText Then
   CboItemDesc.Visible = False
   Exit Sub
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
rsInvItem.Find "itemcode = '" & CboItemDesc.BoundText & "'", , adSearchForward, 1

With rsInvItem
If .EOF Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   ItemGrd.TextMatrix(CurrRow, 6) = ""
   CurrAmt = 0
   txtQty = ""
   txtRate = ""
   ValidRow = True
Else
   ItemGrd.TextMatrix(CurrRow, 1) = rsInvItem!itemcode
   ItemGrd.TextMatrix(CurrRow, 2) = !ITEMNAME
   ItemGrd.TextMatrix(CurrRow, 3) = IIf(IsNull(!unit), "Nos", !unit)
   txtQty = ItemGrd.TextMatrix(CurrRow, 4)
   If Not Val(txtQty) > 0 Then
      ValidRow = False
      ItemGrd.row = CurrRow
      txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
      txtQty.Visible = True
      txtQty.SetFocus
   End If
End If
End With
'ltot = Round(ltot + CurrAmt - prevamt, 2)
'lblTot.Caption = Format(ltot, "###,##,##,##0.00")
CboItemDesc.Visible = False
End Sub




Private Sub DBcboParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBcboParty_Validate(Cancel As Boolean)
Dim Party As New ADODB.Recordset
If Len(Trim(DBcboParty.BoundText)) = 0 Then
   MsgBox "Supplier should not be blank "
   Exit Sub
End If
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & DBcboParty.BoundText & "'")
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
   Cancel = True
End If
End Sub

Private Sub Form_Load()

'Dim DatBrBill As New ADODB.Recordset


Set db = New ADODB.Connection
'Set datInvItem = New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open CnnString

If rsInvItem.State = adStateOpen Then rsInvItem.Close


   rsInvItem.Open "select *  from invitems order by itemname", db, adOpenForwardOnly, adLockReadOnly

Set CboItemDesc.RowSource = rsInvItem
CboItemDesc.ListField = "ITEMNAME"
CboItemDesc.BoundColumn = "ItemCode"



Set rsInvItem = MHVDB.Execute("select * from invitems order by itemname")
'
If rspobill.State = adStateOpen Then rspobill.Close
rspobill.Open ("select * from PurOrderHdr where procyear='" & SysYear & "' and status = 'ON' order by purordno desc"), db

Set CboBillNo.RowSource = rspobill
CboBillNo.ListField = "purordno"
CboBillNo.BoundColumn = "purordno"


If rsbrbill.State = adStateOpen Then rsbrbill.Close
 rsbrbill.Open ("SELECT * FROM supplieR"), db

Set DBcboParty.RowSource = rsbrbill
DBcboParty.ListField = "Name"
DBcboParty.BoundColumn = "SuplCode"

Set cbotransporter.RowSource = rsbrbill
cbotransporter.ListField = "Name"
cbotransporter.BoundColumn = "SuplCode"

ValidRow = True
CurrRow = 1
lblYr = SysYear & "\"
ItemGrd.row = 1
ItemGrd.Col = 1
cboItemCode.Left = ItemGrd.Left + ItemGrd.CellLeft
cboItemCode.Width = ItemGrd.CellWidth
cboItemCode.Height = ItemGrd.CellHeight
ItemGrd.Col = 2
CboItemDesc.Left = ItemGrd.Left + ItemGrd.CellLeft
CboItemDesc.Width = ItemGrd.CellWidth
CboItemDesc.Height = ItemGrd.CellHeight
ItemGrd.Col = 4
txtQty.Left = ItemGrd.Left + ItemGrd.CellLeft
txtQty.Width = ItemGrd.CellWidth
txtQty.Height = ItemGrd.CellHeight
ItemGrd.Col = 5
txtRate.Left = ItemGrd.Left + ItemGrd.CellLeft
txtRate.Width = ItemGrd.CellWidth
txtRate.Height = ItemGrd.CellHeight
ltot = 0
End Sub
Private Sub ItemGrd_Click()
Dim jrow, jCol As Integer
If Not ValidRow And CurrRow <> ItemGrd.row Then
   ItemGrd.row = CurrRow
   Exit Sub
End If
jrow = ItemGrd.row
jCol = ItemGrd.Col
If jrow = 0 Then Exit Sub
If jrow > 1 And Len(ItemGrd.TextMatrix(jrow - 1, 1)) = 0 Then
   Beep
   Exit Sub
End If
If CurrRow > ItemGrd.Rows - 2 Then
   ItemGrd.Rows = CurrRow + 3
End If
ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
CurrRow = jrow
ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
Select Case jCol
       Case 1
            cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
            cboItemCode = ItemGrd.Text
            cboItemCode.Visible = True
            cboItemCode.SetFocus
       Case 2
            CboItemDesc.Top = ItemGrd.Top + ItemGrd.CellTop
            CboItemDesc = ItemGrd.Text
            CboItemDesc.BoundText = ItemGrd.TextMatrix(CurrRow, 1)
            CboItemDesc.Visible = True
            CboItemDesc.SetFocus
       Case 4
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
               txtQty = ItemGrd.Text
               txtQty.Visible = True
               txtQty.SetFocus
            End If
       Case 5
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
               txtRate = ItemGrd.Text
               txtRate.Visible = True
               txtRate.SetFocus
            End If
    End Select
End Sub

Private Sub ItemGrd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 2 Then
   If CurrRow > 0 And Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
      ltot = ltot - Val(ItemGrd.TextMatrix(CurrRow, 6))
      lblTot.Caption = Format(ltot, "######0.00")
      ItemGrd.RemoveItem CurrRow
      ItemGrd.AddItem ""
   Else
      Beep
      Beep
   End If
End If
End Sub

Private Sub ItemGrd_Scroll()
'SendKeys "{TAB}", True
End Sub

Private Sub mnuadd_Click()
Dim lastbill As New ADODB.Recordset
txtDate(0) = Format(Date, "dd/mm/yyyy")
ValidRow = True
Operation = "ADD"
CurrRow = 1
txtRemark = ""
'txtNoTbl = ""
'txtpax = ""
'txtBillTo = ""
ltot = 0
ErrCTR = 0
txtDate(2) = Format(Now, "dd/mm/yyyy")
txtDate(1) = Format(Now, "dd/mm/yyyy")
cboItemCode.Visible = False
CboItemDesc.Visible = False
txtQty.Visible = False
Set lastbill = MHVDB.Execute("select max(purordno) as lno from PurOrderHdr where procyear='" & SysYear & "'")
CboBillNo = IIf(IsNull(lastbill!lno), 1, lastbill!lno + 1)
Set lastbill = Nothing
CboBillNo.Enabled = False
Frame2.Enabled = True
ItemGrd.Enabled = True
ItemGrd.Clear
ItemGrd.FormatString = fmString

TB.Buttons(3).Enabled = True

TB.Buttons(4).Enabled = False

'TB.Buttons(5).Enabled = False

End Sub
Private Sub mnuCancel_Click()
Dim UpdtStr
Dim jrec As ADODB.Recordset
On Error GoTo ERR
If MsgBox("Cancel it !!!Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
MHVDB.BeginTrans
UpdtStr = "UPDATE  PurOrderHdr SET STATUS = 'C',REMARKs = '" & txtRemark & "' WHERE procyear='" & SysYear & "' and purordno = VAL('" & CboBillNo & "')"
MHVDB.Execute UpdtStr
Frame2.Enabled = False
MHVDB.CommitTrans
DatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False

TB.Buttons(4).Enabled = False
Exit Sub
ERR:
MsgBox "error :" + IIf(IsNull(ERR.Description), " ", ERR.Description)
ERR.Clear
MHVDB.Rollback
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
Operation = "OPEN"
Frame2.Enabled = True
CboBillNo.Enabled = True

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = 0

'TB.Buttons(5).Enabled = False
ErrCTR = 0
CboBillNo.Refresh
End Sub

Private Sub mnuprn_Click()
'Dim rs As New ADODB.Recordset
'Dim TRNAME As String
'Dim tt As Double
'Set rs = Nothing
'rs.Open "SELECT * FROM SUPPLIER WHERE SUPLCODE='" & cbotransporter.BoundText & "'", MHVDB
'If rs.EOF <> True Then
'TRNAME = rs!Name
'Else
'TRNAME = ""
'End If
'tt = lblTot
'
'MHVDB.Execute "update rpthlp set billno=('" & CboBillNo & "'),mname='" & SysYear & "',TRNAME='" & TRNAME & "',TOT5=' " & tt & " '"
'frmsTOREmenu.Crp.ReportFileName = App.Path + "\purchaseorder.rpt"
'
'frmsTOREmenu.Crp.Action = 1
End Sub

Private Sub mnuSave_Click()
Dim i, j, K As Integer
Dim printNow As Boolean
Dim jrec As ADODB.Recordset
Dim InsStr, JStat, pcODE As String
If Not (Operation = "OPEN" Or Operation = "ADD") Then
   Beep
   Exit Sub
End If
If Not ValidRow Then Exit Sub
printNow = True
0:
On Error GoTo ERR

If Operation = "ADD" Then
   
   InsStr = "insert into PurOrderHdr (procyear, purordno,poDATE,suplcode,Status,refno,refdate,deledate,remarks,exduty,trcode,racc,iremarks,r2,r3) values ( '" & SysYear & "','" & CboBillNo & "'," _
                  & " '" & Format(txtDate(0), "yyyyMMdd") & "','" & DBcboParty.BoundText & "','ON','" & txtChallan & "','" & Format(txtDate(1), "yyyyMMdd") & "','" & Format(txtDate(2), "yyyyMMdd") & "','" & txtRemark & "','" & (txtexciseduty.Text) & "','" & cbotransporter.BoundText & "','" & txtracc.Text & "','" & txtiremarks.Text & "','" & txtr2.Text & "','" & txtr3.Text & "')"
   MHVDB.Execute InsStr
   For i = 1 To 994
       If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
          InsStr = "insert into PurOrderDtl (procyear, purordno,itemcode,qty,rate) values (  '" & SysYear & "','" & CboBillNo & "'," _
                  & " '" & ItemGrd.TextMatrix(i, 1) & "','" & ItemGrd.TextMatrix(i, 4) & "','" & ItemGrd.TextMatrix(i, 5) & "')"
          MHVDB.Execute InsStr
       Else
          Exit For
       End If
   Next
Else
   InsStr = "update PurOrderHdr set suplcode='" & DBcboParty.BoundText & "',podate=('" & Format(txtDate(0), "yyyyMMdd") & "'), " _
          & " refno='" & txtChallan & "',refdate='" & Format(txtDate(1), "yyyyMMdd") & "',remarks='" & txtRemark & "',trcode='" & cbotransporter.BoundText & "',racc='" & txtracc.Text & "',iremarks='" & txtiremarks.Text & "',r2='" & txtr2.Text & "',r3='" & txtr3.Text & "',exduty='" & (txtexciseduty.Text) & "' where procyear='" & SysYear & "' and purordno = ('" & CboBillNo & "')"
   MHVDB.Execute InsStr
   MHVDB.Execute "delete  from PurOrderDtl where procyear='" & SysYear & "' and purordno = ('" & CboBillNo & "')"
   For i = 1 To 994
       If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
          InsStr = "insert into PurOrderDtl (procyear, purordno,itemcode,qty,rate) values (  '" & SysYear & "','" & CboBillNo & "'," _
                  & " '" & ItemGrd.TextMatrix(i, 1) & "','" & ItemGrd.TextMatrix(i, 4) & "','" & ItemGrd.TextMatrix(i, 5) & "')"
          MHVDB.Execute InsStr
       Else
          Exit For
       End If
   Next
End If
'printNow = IIf(MsgBox("Print Now ?", vbYesNo) = vbYes, True, False)
'If printNow Then PrintBill

DatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False
Frame2.Enabled = False

TB.Buttons(3).Enabled = False

TB.Buttons(4).Enabled = 0

TB.Buttons(5).Enabled = True
ErrCTR = 0
Exit Sub
ERR:
ErrCTR = ErrCTR + 1
'If ErrCTR > 5 Then
'   If DBEngine.Errors.Count > 0 Then
'   For Each errLoop In DBEngine.Errors
'       MsgBox "Error number: " & errLoop.Number & vbCr & _
'       errLoop.Description
'   Next errLoop
''Exit Sub
'   End If
'End If
ERR.Clear

If ErrCTR < 6 Then
   For i = 1 To 1000
       For j = 1 To 9999
       Next
   Next
   GoTo 0
End If
End Sub
Private Sub Tb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
       Case "ADD"
           mnuadd_Click
       Case "OPEN"
           mnuOpen_Click
       Case "SAVE"
           mnuSave_Click
       Case "PRINT"
           mnuprn_Click
       Case "DELETE"
          ' mnuCancel_Click
       Case "EXIT"
           Unload Me
End Select
End Sub


Private Sub txtChallan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtQty)) Then
      Beep
      MsgBox "Enter a valid Quantity"
      ValidRow = False
      Exit Sub
   Else
      ItemGrd.TextMatrix(CurrRow, 4) = txtQty
      ValidRow = True
   End If
   End If
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
   ItemGrd.TextMatrix(CurrRow, 6) = Val(txtQty) * Val(ItemGrd.TextMatrix(CurrRow, 5))
   CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
   ltot = Round(ltot + CurrAmt - prevamt, 2)
   lblTot.Caption = Format(ltot, "###,##,##,##0.00")
   txtQty.Visible = False
   ItemGrd.row = CurrRow
   ItemGrd.Col = 5
   txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
   txtRate = ItemGrd.Text
   txtRate.Visible = True
   txtRate.SetFocus
End If
End Sub

Private Sub txtQty_validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtQty)) Then
   Beep
   MsgBox "Enter a valid Quantity"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 4) = txtQty
   ValidRow = True
End If
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ItemGrd.TextMatrix(CurrRow, 6) = Val(txtQty) * Val(ItemGrd.TextMatrix(CurrRow, 5))
CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
txtQty.Visible = False
End Sub


Private Sub txtRate_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtRate) Or Val(txtRate) > 0) Then
      Beep
      MsgBox "Enter a valid Rate !!!"
      ValidRow = False
      Exit Sub
   Else
      ItemGrd.TextMatrix(CurrRow, 5) = txtRate
      ValidRow = True
   End If
   End If
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
   ItemGrd.TextMatrix(CurrRow, 6) = Val(txtRate) * (ItemGrd.TextMatrix(CurrRow, 4))
   CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
   ltot = Round(ltot + CurrAmt - prevamt, 2)
   lblTot.Caption = Format(ltot, "###,##,##,##0.00")
   txtRate.Visible = False
   ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
   CurrRow = CurrRow + 1
   If CurrRow > ItemGrd.Rows - 2 Then
      ItemGrd.Rows = CurrRow + 3
   End If
   ItemGrd.row = CurrRow
   ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
   ItemGrd.Col = 1
   cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
   cboItemCode = ItemGrd.Text
   cboItemCode.Visible = True
   cboItemCode.SetFocus
End If
End Sub

Private Sub txtrate_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtRate) Or Val(txtRate) > 0) Then
   Beep
   MsgBox "Enter a valid Rate !!!"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ItemGrd.TextMatrix(CurrRow, 5) = txtRate
   ValidRow = True
End If
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ItemGrd.TextMatrix(CurrRow, 6) = Val(txtRate) * (ItemGrd.TextMatrix(CurrRow, 4))
CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
txtRate.Visible = False
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ItemGrd.Enabled = True
   CurrRow = 1
   ItemGrd.row = CurrRow
   ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
   ItemGrd.Col = 1
   cboItemCode.Top = ItemGrd.Top + ItemGrd.CellTop
   cboItemCode = ItemGrd.Text
   cboItemCode.Visible = True
   cboItemCode.SetFocus
End If
End Sub

Private Sub txtRemark_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtRemark.ToolTipText = txtRemark.Text
End Sub
