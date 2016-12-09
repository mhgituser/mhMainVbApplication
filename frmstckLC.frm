VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsTOCKCost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LANDING COST"
   ClientHeight    =   8280
   ClientLeft      =   3330
   ClientTop       =   1170
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmstckLC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12240
   Visible         =   0   'False
   Begin MSComCtl2.DTPicker txtPHDate 
      Height          =   315
      Left            =   10440
      TabIndex        =   44
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   80216065
      CurrentDate     =   36944
   End
   Begin VB.TextBox txtPHBillNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      MaxLength       =   35
      TabIndex        =   42
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtFrBillNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      MaxLength       =   35
      TabIndex        =   37
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtInvNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      MaxLength       =   35
      TabIndex        =   33
      Top             =   1080
      Width           =   2535
   End
   Begin MSComctlLib.StatusBar sT1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7905
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker LBLnOW 
      Height          =   315
      Left            =   10440
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80216065
      CurrentDate     =   36797
   End
   Begin VB.TextBox txtChallan 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   4
      ToolTipText     =   "Enter Challan No"
      Top             =   1440
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker txtChDate 
      Height          =   315
      Index           =   0
      Left            =   10440
      TabIndex        =   3
      ToolTipText     =   "Enter Challan Date"
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80216065
      CurrentDate     =   36383
      MinDate         =   36161
   End
   Begin VB.TextBox txtChallan 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "Enter Challan No"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8640
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   11895
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   45
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtph 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtacnt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10680
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtmi 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtUL 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtBSt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid ItemGrd 
         Height          =   3660
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   6456
         _Version        =   393216
         Rows            =   201
         Cols            =   15
         RowHeightMin    =   315
         ForeColorFixed  =   -2147483635
         ScrollTrack     =   -1  'True
         HighLight       =   0
         FormatString    =   $"frmstckLC.frx":076A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Remarks :"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   9
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
   Begin MSDataListLib.DataCombo CboBillNo 
      Bindings        =   "frmstckLC.frx":082D
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   1920
      TabIndex        =   46
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo CboPO 
      Bindings        =   "frmstckLC.frx":0842
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   6600
      TabIndex        =   47
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DBcboParty 
      Bindings        =   "frmstckLC.frx":0857
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1920
      TabIndex        =   48
      Top             =   1800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DBPH 
      Bindings        =   "frmstckLC.frx":086C
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1920
      TabIndex        =   49
      Top             =   2280
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      OLEDropMode     =   1
      BackColor       =   -2147483643
      ListField       =   "itemcode"
      BoundColumn     =   "ItemCode"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo lblsupl 
      Bindings        =   "frmstckLC.frx":0881
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   1920
      TabIndex        =   50
      Top             =   2880
      Width           =   3855
      _ExtentX        =   6800
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
            Picture         =   "frmstckLC.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckLC.frx":0C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckLC.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckLC.frx":1CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckLC.frx":20F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmstckLC.frx":28B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   1164
      ButtonWidth     =   1005
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OPEN"
            Key             =   "OPEN"
            Object.ToolTipText     =   "OPEN/EDIT EXISTING RECORD"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "SAVE"
            Key             =   "SAVE"
            Object.ToolTipText     =   "SAVES RECORD"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Caption         =   "Date"
      Height          =   255
      Left            =   9360
      TabIndex        =   43
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "P/ H Bill No."
      Height          =   255
      Left            =   5880
      TabIndex        =   41
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Packing/Handling Agent Code"
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label LblPh 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   6840
      TabIndex        =   39
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label Label9 
      Caption         =   "Date"
      Height          =   255
      Left            =   9360
      TabIndex        =   36
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblChDate 
      Height          =   255
      Left            =   10560
      TabIndex        =   35
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Bill Date"
      Height          =   255
      Left            =   9240
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Bill No."
      Height          =   255
      Left            =   5040
      TabIndex        =   32
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblBillDate 
      AutoSize        =   -1  'True
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3720
      TabIndex        =   31
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label7 
      Caption         =   "Freight Bill No."
      Height          =   255
      Left            =   5880
      TabIndex        =   30
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Transporter Code"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lbllc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   10680
      TabIndex        =   24
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label lblmi 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   9600
      TabIndex        =   23
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label lblul 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   8520
      TabIndex        =   22
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label lblbst 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   7680
      TabIndex        =   21
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label lblFr 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5760
      TabIndex        =   20
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label lblyr 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " "
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
      Left            =   1755
      TabIndex        =   17
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4800
      TabIndex        =   16
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label Label4 
      Caption         =   "Purchase Order no"
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
      Left            =   5040
      TabIndex        =   15
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Gate Pass No."
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Challan Date"
      Height          =   255
      Index           =   2
      Left            =   9240
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Challan No."
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblParty 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Entry No."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmsTOCKCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BillDetailRec As New ADODB.Recordset
Dim Bill As New ADODB.Recordset
Dim PORec As New ADODB.Recordset
Dim rsDatBrBill As New ADODB.Recordset
Dim dataPO As New ADODB.Recordset
Dim rsParty As New ADODB.Recordset
Dim CurrRow, Jkey, ErrCTR, JBAcNo, JFAcNo, JPAcNo As Long
Dim ValidRow As Boolean

Dim ltot, Lfr, LBST, LUL, Lmi, LLC, LPh As Double
Const fmString = "     |^  Code      ||^   Unit   |^       Qty       |^    Pur. Rate |^      Amount     |^   Freight  |^P/H Charge|^    BST      |^UL Charge|^ Add/Less |^ Landing Cost|^   L. Rate    |Acnt Code"
Private Sub PrintBill()
Dim Supp
Dim Party As ADODB.Recordset
Dim i, Lin As Integer
Dim Jamt, JFr, JPH, Jbst, Jul, JMisc As Double
Dim pfile As String
pfile = "SE" + Trim(CboBillNo) + ".txt"
On Error GoTo jerr:
Do While True
   Select Case MsgBox("Printer Ready ? ", vbYesNoCancel)
          Case vbYes
               Open "lpt1:" For Output As #1
               Exit Do
          Case vbNo
               If MsgBox("Connect/Swich on the printer and retry.", vbRetryCancel) = vbCancel Then
                  MsgBox "Bill stored to " + pfile
                  Open pfile For Output As #1
                  Exit Do
               End If
          Case vbCancel
               MsgBox "Bill stored to " + pfile
               Open pfile For Output As #1
               Exit Do
   End Select
Loop
'Open "lpt1:" For Output As #1
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & DBcboParty.BoundText & "'", dbOpenDynaset, DBReadOnly)
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
Else
   Supp = Party!Name + " ( " + DBcboParty.BoundText + " )"
End If
Set BillDetailRec = MHVDB.Execute("select d.itemcode,d.qty,d.rate,B.unit,itemname,freight,bst,ulc,misc,phcharge,lrate from tranfile as d,invitems as b where d.itemcode=b.itemcode and billno=val('" & CboBillNo & "') and d.billtype='EN' and d.procyear='" & SysYear & "'", dbOpenForwardOnly)
Lin = 61
With BillDetailRec
i = 1
Jamt = 0
JFr = 0
JPH = 0
Jbst = 0
Jul = 0
JMisc = 0
Do While Not .EOF
'For i = 1 To ItemGrd.Rows - 1
'    If Len(Trim(ItemGrd.TextMatrix(i, 1))) = 0 Then Exit For
    If Lin > 60 Then
       Print #1, Chr(18);
       Print #1, Chr(14) + "MHV" + Chr(20)
       Print #1, "MONGAR"
       Print #1, "Stock Entry Slip"
       Print #1, Chr(15) + String(131, "_")
       Print #1, PadWithChar("Entry No.", 12, " ", 0) + PadWithChar(lblyr + Trim(CboBillNo), 16, " ", 0) + " " + PadWithChar(" Date ", 5, " ", 0) + Format(LBLnOW, "dd/mm/yyyy") + "  " + PadWithChar("Purchase Order No.", 19, " ", 0) + PadWithChar(": " + CboPO, 16, " ", 0)
       Print #1, PadWithChar("Challan No.", 12, " ", 0) + PadWithChar(Trim(txtChallan(0)), 16, " ", 0) + " " + PadWithChar(" Date ", 5, " ", 0) + Format(lblChDate, "dd/mm/yyyy") + "  " + PadWithChar("Gate Pass No.", 19, " ", 0) + PadWithChar(": " + txtChallan(1), 16, " ", 0)
       Print #1, PadWithChar("Invoice No.", 12, " ", 0) + PadWithChar(Trim(txtInvNo), 16, " ", 0) + " " + PadWithChar(" Date ", 5, " ", 0) + Format(txtChDate(0), "dd/mm/yyyy")
       Print #1, "Supplier :  " + PadWithChar(lblsupl, 50, " ", 0) + " Transporter : " + Supp
       Print #1, "Remarks  :  " + txtRemark
       Print #1, String(131, "_")
       Print #1, "Code  |Description                    |    Qty.      |  Rate  |     Amount |F/P Charge|Add/Less|Freight|  BST   |Un Load|  L.Rate "
       Print #1, String(131, "_")
       Lin = 12
    End If
    Print #1, PadWithChar(!itemcode, 7, " ", 0);
    Print #1, PadWithChar(!ITEMNAME, 31, " ", 0) + PadWithChar(!qty, 7, " ", 1) + " " + PadWithChar(!unit, 7, " ", 0);
    Print #1, PadWithChar(Format(Round(!Rate, 2), "#####0.00"), 9, " ", 1) + " " + PadWithChar(Format(Round(!qty * !Rate, 2), "#######0.00"), 11, " ", 1) + " ";
    Print #1, PadWithChar(Format(Round(!phcharge, 2), "######0.00"), 10, " ", 1) + PadWithChar(Format(Round(!misc, 2), "#######0.00"), 9, " ", 1) + " ";
    Print #1, PadWithChar(Format(Round(!freight, 2), "####0.00"), 8, " ", 1) + " " + PadWithChar(Format(Round(!bst, 2), "####0.00"), 8, " ", 1);
    Print #1, PadWithChar(Format(Round(!ulc, 2), "#####0.00"), 8, " ", 1) + " " + PadWithChar(Format(Round(!lRate, 2), "####0.00"), 8, " ", 1)
    Jamt = Jamt + Round(!Rate * !qty, 2)
    JFr = JFr + !freight
    JPH = JPH + !phcharge
    Jbst = Jbst + !bst
    Jul = Jul + !ulc
    JMisc = JMisc + !misc
    If Len(ItemGrd.TextMatrix(i, 2)) > 31 Then
       Print #1, "     " + Mid(!ITEMNAME, 32)
       Lin = Lin + 1
    End If
    Lin = Lin + 1
    i = i + 1
    .MoveNext
Loop
Print #1, String(131, "_")
Print #1, PadWithChar("Total", 63, " ", 0) + PadWithChar(Format(Round(Jamt, 2), "#######0.00"), 11, " ", 1) + " ";
Print #1, PadWithChar(Format(Round(JPH, 2), "######0.00"), 10, " ", 1) + " " + PadWithChar(Format(Round(JMisc, 2), "#######0.00"), 8, " ", 1) + " ";
Print #1, PadWithChar(Format(Round(JFr, 2), "####0.00"), 8, " ", 1) + " " + PadWithChar(Format(Round(Jbst, 2), "####0.00"), 8, " ", 1);
Print #1, PadWithChar(Format(Round(Jul, 2), "####0.00"), 8, " ", 1) + " "
Print #1, String(131, "_")
Print #1, PadWithChar("Freight/Packing Handling", 63, " ", 0) + PadWithChar(Format(Round(JPH, 2), "#######0.00"), 11, " ", 1)
Print #1, PadWithChar("Add / Less ", 63, " ", 0) + PadWithChar(Format(Round(JMisc, 2), "#######0.00"), 11, " ", 1)
Print #1, String(74, "_")
Print #1, PadWithChar("Bill Amount ", 63, " ", 0) + PadWithChar(Format(Round(Jamt + JPH + JMisc, 2), "#######0.00"), 11, " ", 1)
Print #1, String(131, "_")
End With
Print #1,
Print #1,
Print #1, "   Prepared By            Checked By            Sr. Manager(S&P)            Sr. Manager(F&A)           Chief Executive"
Print #1, String(131, "_")
Print #1,
Print #1,
Print #1,
Close #1 '*/
Exit Sub
jerr:
MsgBox ERR.Description
ERR.Clear
End Sub



Private Sub CboBillNo_GotFocus()
'Frame2.Enabled = False
End Sub

Private Sub CboBillNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub CboBillNo_LostFocus()
Dim i As Integer
Dim Issue, Recv As Double
Dim imast As ADODB.Recordset
If Operation = "ADD" Then Exit Sub
ltot = 0
jvcr = 0
Set Bill = MHVDB.Execute("select * from tranhdr where procyear='" & SysYear & "' and ((billno))=('" & CboBillNo & "') AND billtype = 'EN' and status<>'C'")
If Bill.EOF Then
   MsgBox CboBillNo + " Does not exists "
   CboBillNo.SetFocus
   Exit Sub
Else
   If Bill!ProcYear <> Val(SysYear) Then
      MsgBox " Please login For  " + Str(Bill!ProcYear)
      CboBillNo.SetFocus
      Exit Sub
   End If
   With Bill
   lblBillDate = Format(!billdate, "dd/mm/yyyy")
    ' DatKotHead.Recordset!Time
   txtChallan(0) = IIf(IsNull(!challanno), "", !challanno)
   txtChallan(1) = IIf(IsNull(!gatepassno), "", !gatepassno)
   lblChDate = !challandate
   txtChDate(0) = IIf(IsNull(!invDate), !billdate, !invDate)
   txtInvNo = IIf(IsNull(!INVNO), "", !INVNO)
   JBAcNo = IIf(IsNull(!bacno), 0, !bacno)
   JFAcNo = IIf(IsNull(!facno), 0, !facno)
   JPAcNo = IIf(IsNull(!pacno), 0, !pacno)
   txtFrBillNo = IIf(IsNull(!frbillno), "", !frbillno)
   txtPHBillNo = IIf(IsNull(!phbillno), "", !phbillno)
   CboPO = IIf(IsNull(!purorderid), "", !purorderid)
   txtRemark = IIf(IsNull(!remarks), "", !remarks)
   Set imast = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & !suplcode & "'")
   If Not imast.EOF Then
      lblsupl = !suplcode
      lblsupl.ToolTipText = imast!Name + " (" + !suplcode + ")"
      lblsupl.Tag = !suplcode
   Else
      MsgBox "Supplier Code is not correct !!"
      Frame2.Enabled = False
     
      TB.Buttons(3).Enabled = False
     
      TB.Buttons(4).Enabled = False
      
      Exit Sub
   End If
   Set imast = Nothing
   'If !Status <> "F" Then
      DBcboParty = IIf(IsNull(!suplcode), "", !suplcode)
      DBPH = ""
      LBLnOW = Format(!billdate, "dd/mm/yyyy")
      txtPHDate = Format(!billdate, "dd/mm/yyyy")
   'Else
      DBcboParty = IIf(IsNull(!transportercode), !suplcode, !transportercode)
      DBPH = IIf(IsNull(!phcode), "", !phcode)
      LBLnOW = IIf(IsNull(!Frbilldate), Format(!billdate, "dd/mm/yyyy"), Format(!Frbilldate, "dd/mm/yyyy"))
      txtPHDate = IIf(IsNull(!phbilldate), Format(!billdate, "dd/mm/yyyy"), Format(!phbilldate, "dd/mm/yyyy"))
      'jvcr = IIf(IsNull(!vcrno), 0, !vcrno)
   'End If
   End With
   Set BillDetailRec = MHVDB.Execute("select d.itemcode,d.qty,d.rate,B.unit,d.freight,d.phcharge,d.bst,d.ulc,d.misc,d.lrate,d.acntcode,itemname,avgstockrate from tranfile as d,invitems as b where d.itemcode=b.itemcode and billno=('" & CboBillNo & "') and d.billtype='EN' and d.procyear='" & SysYear & "'", dbOpenForwardOnly)
   ItemGrd.Clear
   ItemGrd.FormatString = fmString
   With BillDetailRec
   i = 1
   ltot = 0
   Lfr = 0
   LBST = 0
   LUL = 0
   Lmi = 0
   LPh = 0
   LLC = 0
   Do While Not .EOF
      ItemGrd.TextMatrix(i, 0) = i
      ItemGrd.TextMatrix(i, 1) = !itemcode
      ItemGrd.TextMatrix(i, 2) = !ITEMNAME
      ItemGrd.TextMatrix(i, 3) = IIf(IsNull(!unit), "", !unit)
      ItemGrd.TextMatrix(i, 4) = IIf(IsNull(!qty), 0, !qty)
      ItemGrd.TextMatrix(i, 5) = IIf(IsNull(!Rate), 0, !Rate)
      ItemGrd.TextMatrix(i, 6) = Round(!qty * !Rate, 2)
      ItemGrd.TextMatrix(i, 7) = IIf(IsNull(!freight), 0, !freight)
      ItemGrd.TextMatrix(i, 8) = IIf(IsNull(!phcharge), 0, !phcharge)
      ItemGrd.TextMatrix(i, 9) = IIf(IsNull(!bst), 0, !bst)
      ItemGrd.TextMatrix(i, 10) = IIf(IsNull(!ulc), 0, !ulc)
      ItemGrd.TextMatrix(i, 11) = IIf(IsNull(!misc), 0, !misc)
      ItemGrd.TextMatrix(i, 12) = IIf(IsNull(!lRate), !qty * !Rate, !qty * !lRate)
      ItemGrd.TextMatrix(i, 13) = IIf(IsNull(!lRate), 0, !lRate)
      ltot = ltot + !qty * !Rate
      Lfr = Lfr + IIf(IsNull(!freight), 0, !freight)
      LPh = LPh + IIf(IsNull(!phcharge), 0, !phcharge)
      LBST = LBST + IIf(IsNull(!bst), 0, !bst)
      LUL = LUL + IIf(IsNull(!ulc), 0, !ulc)
      Lmi = Lmi + IIf(IsNull(!misc), 0, !misc)
      LLC = LLC + IIf(IsNull(!lRate), !qty * !Rate, !qty * !lRate)
      If IsNull(!acntcode) Or Len(!acntcode) = 0 Then
         Set imast = MHVDB.Execute("select acntcode from categoryfile as a,invitems as b where a.category=b.category and b.itemcode='" & !itemcode & "'", dbOpenSnapshot)
         If Not imast.EOF Then ItemGrd.TextMatrix(i, 14) = IIf(IsNull(imast!acntcode), "", imast!acntcode)
         Set imast = Nothing
      Else
         ItemGrd.TextMatrix(i, 14) = !acntcode
      End If
      .MoveNext
      i = i + 1
   Loop
   End With
'   If MsgBox("Cant be modified.You can Cancel or Print It ! Do you want Print ?", vbYesNo) = vbYes Then
'      PrintBill
'   End If
'   Operation = ""
  ' txtRemark.SetFocus
End If
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
lblFr.Caption = Lfr
lblbst.Caption = LBST
lblul.Caption = LUL
lblmi.Caption = Lmi
lbllc.Caption = LLC
Frame2.Enabled = True

TB.Buttons(2).Enabled = True

TB.Buttons(3).Enabled = True
End Sub


Private Sub cboItemDesc_LostFocus()
Dim prevamt, CurrAmt, Jstock As Double
Dim imast As ADODB.Recordset
If ItemGrd.TextMatrix(CurrRow, 1) = CboItemDesc.BoundText Then
   CboItemDesc.Visible = False
   Exit Sub
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
If Len(Trim(CboPO)) > 0 Then
   Set PORec = MHVDB.Execute("select d.itemcode from purorderdtl as d,purorderhdr as a where str(a.purordno)+'/'+str(a.procyear)='" & CboPO & "' and a.status='ON' and  a.procyear=d.procyear and a.purordno=d.purordno and d.itemcode=itemcode = '" & CboItemDesc.BoundText & "'", DBReadOnly)
   If PORec.EOF Then
      If MsgBox("This Item is not in the Purchase Order " + CboPO + " ! Do u want to Enter this ?", vbYesNo) = vbNo Then
         ItemGrd.TextMatrix(CurrRow, 1) = ""
         ItemGrd.TextMatrix(CurrRow, 2) = ""
         ItemGrd.TextMatrix(CurrRow, 3) = ""
         ItemGrd.TextMatrix(CurrRow, 4) = ""
         ItemGrd.TextMatrix(CurrRow, 5) = ""
         ItemGrd.TextMatrix(CurrRow, 6) = 0
         txtQty = ""
         txtRate = ""
         ValidRow = True
         Exit Sub
      End If
   End If
   Set PORec = Nothing
End If
datInvItem.Recordset.FindFirst "itemcode = '" & CboItemDesc.BoundText & "'"
With datInvItem.Recordset
If .NoMatch Then
   ItemGrd.TextMatrix(CurrRow, 1) = ""
   ItemGrd.TextMatrix(CurrRow, 2) = ""
   ItemGrd.TextMatrix(CurrRow, 3) = ""
   ItemGrd.TextMatrix(CurrRow, 4) = ""
   ItemGrd.TextMatrix(CurrRow, 5) = ""
   ItemGrd.TextMatrix(CurrRow, 6) = 0
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
CurrAmt = Val(ItemGrd.TextMatrix(i, 6))
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
CboItemDesc.Visible = False
End Sub




Private Sub CboPO_Validate(Cancel As Boolean)
If Len(Trim(CboPO)) > 0 Then
   ltot = 0
   Set PORec = MHVDB.Execute("select a.suplcode,d.itemcode,d.qty,d.rate,B.unit,itemname from purorderdtl as d,purorderhdr as a,invitems as b where Ltrim(str(a.purordno))+'/'+Ltrim(str(a.procyear))='" & CboPO & "' and a.status='ON' and  a.procyear=d.procyear and a.purordno=d.purordno and d.itemcode=b.itemcode ", dbOpenForwardOnly)
   If PORec.EOF Then
      MsgBox "Wrong Purchase Order No!!!"
      Cancel = True
      Exit Sub
   End If
   DBcboParty = PORec!suplcode
   ItemGrd.Clear
   ItemGrd.FormatString = fmString
   i = 1
   With PORec
   Do While Not .EOF
      ItemGrd.TextMatrix(i, 0) = i
      ItemGrd.TextMatrix(i, 1) = !itemcode
      ItemGrd.TextMatrix(i, 2) = !ITEMNAME
      ItemGrd.TextMatrix(i, 3) = !unit
      ItemGrd.TextMatrix(i, 4) = !qty
      ItemGrd.TextMatrix(i, 5) = Format(!Rate, "####0.00")
      ItemGrd.TextMatrix(i, 6) = !qty * !Rate
      ltot = ltot + !qty * !Rate
      i = i + 1
      .MoveNext
   Loop
   End With
   lblTot.Caption = Format(ltot, "###,##,##,##0.00")
   Set PORec = Nothing
End If
End Sub

Private Sub DBcboParty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub DBcboParty_Validate(Cancel As Boolean)
Dim Party As ADODB.Recordset
If Len(Trim(DBcboParty.BoundText)) = 0 Then
   MsgBox "Transporter should not be blank "
   Exit Sub
End If
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & DBcboParty.BoundText & "'")
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
   Cancel = True
End If
Set Party = Nothing
End Sub

Private Sub DBPH_Validate(Cancel As Boolean)
Dim Party As ADODB.Recordset
If Len(Trim(DBPH)) = 0 Then
   Exit Sub
End If
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & DBPH.BoundText & "'")
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
   Cancel = True
End If
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

' biill no
If rsDatBrBill.State = adStateOpen Then rsDatBrBill.Close
rsDatBrBill.Open "select * from tranhdr where status <> 'C' and billtype='EN' and procyear='" & SysYear & "' order by billno desc", db
Set CboBillNo.RowSource = rsDatBrBill
CboBillNo.ListField = "billno"
CboBillNo.BoundColumn = "billno"


'Set DatParty.Recordset = MHVDB.Execute("SELECT * FROM supplieR", dbOpenDynaset, dbReadOnly)
If rsParty.State = adStateOpen Then rsParty.Close
 rsParty.Open ("SELECT * FROM supplieR"), db
' supplier
Set lblsupl.RowSource = rsParty
lblsupl.ListField = "Name"
lblsupl.BoundColumn = "SuplCode"
 
 ' transporter
 Set DBcboParty.RowSource = rsParty
DBcboParty.ListField = "Name"
DBcboParty.BoundColumn = "SuplCode"

' agents
Set DBPH.RowSource = rsParty
DBPH.ListField = "Name"
DBPH.BoundColumn = "SuplCode"
' PO
If dataPO.State = adStateOpen Then dataPO.Close
 dataPO.Open "SELECT purOrdNo,Ltrim(CAST(purOrdNo AS CHAR))+'/'+Ltrim(CAST(procyear AS CHAR)) as po  FROM purorderhdr where status='ON' order by purordno desc  ", db
Set CboPO.RowSource = dataPO
CboPO.ListField = "PO"
CboPO.BoundColumn = "purOrdNo"

ValidRow = True
lblyr = "EN\" & SysYear & "\"
CurrRow = 1
ItemGrd.row = 1
'ItemGrd.Col = 1
'cboitemcode.Left = ItemGrd.Left + ItemGrd.CellLeft
'cboitemcode.Width = ItemGrd.CellWidth
'cboitemcode.Height = ItemGrd.CellHeight
'ItemGrd.Col = 2
'CboItemDesc.Left = ItemGrd.Left + ItemGrd.CellLeft
'CboItemDesc.Width = ItemGrd.CellWidth
'CboItemDesc.Height = ItemGrd.CellHeight
ItemGrd.Col = 7
txtQty.Left = ItemGrd.Left + ItemGrd.CellLeft
txtQty.Width = ItemGrd.CellWidth
txtQty.Height = ItemGrd.CellHeight
ItemGrd.Col = 5
txtRate.Left = ItemGrd.Left + ItemGrd.CellLeft
txtRate.Width = ItemGrd.CellWidth
txtRate.Height = ItemGrd.CellHeight
ItemGrd.Col = 8
txtph.Left = ItemGrd.Left + ItemGrd.CellLeft
txtph.Width = ItemGrd.CellWidth
txtph.Height = ItemGrd.CellHeight
ItemGrd.Col = 9
txtBSt.Left = ItemGrd.Left + ItemGrd.CellLeft
txtBSt.Width = ItemGrd.CellWidth
txtBSt.Height = ItemGrd.CellHeight
ItemGrd.Col = 10
txtUL.Left = ItemGrd.Left + ItemGrd.CellLeft
txtUL.Width = ItemGrd.CellWidth
txtUL.Height = ItemGrd.CellHeight
ItemGrd.Col = 11
txtmi.Left = ItemGrd.Left + ItemGrd.CellLeft
txtmi.Width = ItemGrd.CellWidth
txtmi.Height = ItemGrd.CellHeight

ltot = 0
   Lfr = 0
   LBST = 0
   LUL = 0
   Lmi = 0
   LLC = 0
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
sT1.SimpleText = ItemGrd.TextMatrix(CurrRow, 2)
Select Case jCol
       Case 7
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
               txtQty.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtQty.Width = ItemGrd.CellWidth
               txtQty = ItemGrd.Text
               txtQty.Visible = True
               txtQty.SetFocus
            End If
       Case 5
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
               txtRate.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtRate.Width = ItemGrd.CellWidth
               txtRate = ItemGrd.Text
               txtRate.Visible = True
               txtRate.SetFocus
            End If
       Case 6
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtAmt.Top = ItemGrd.Top + ItemGrd.CellTop
               txtAmt.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtAmt.Width = ItemGrd.CellWidth
               txtAmt = ItemGrd.Text
               txtAmt.Visible = True
               txtAmt.SetFocus
            End If
       Case 8
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtph.Top = ItemGrd.Top + ItemGrd.CellTop
               txtph.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtph.Width = ItemGrd.CellWidth
               txtph.Height = ItemGrd.CellHeight
               txtph = ItemGrd.Text
               txtph.Visible = True
               txtph.SetFocus
            End If
       Case 9
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtBSt.Top = ItemGrd.Top + ItemGrd.CellTop
               txtBSt.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtBSt = ItemGrd.Text
               txtBSt.Visible = True
               txtBSt.SetFocus
            End If
       Case 10
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtUL.Top = ItemGrd.Top + ItemGrd.CellTop
               txtUL.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtUL = ItemGrd.Text
               txtUL.Visible = True
               txtUL.SetFocus
            End If
       Case 11
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtmi.Top = ItemGrd.Top + ItemGrd.CellTop
               txtmi.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtmi.Width = ItemGrd.CellWidth
               txtmi.Height = ItemGrd.CellHeight
               txtmi = ItemGrd.Text
               txtmi.Visible = True
               txtmi.SetFocus
            End If
       Case 14
            If Len(ItemGrd.TextMatrix(jrow, 1)) > 0 Then
               txtacnt.Top = ItemGrd.Top + ItemGrd.CellTop
               txtacnt.Left = ItemGrd.Left + ItemGrd.CellLeft
               txtacnt.Width = ItemGrd.CellWidth
               txtacnt.Height = ItemGrd.CellHeight
               txtacnt = ItemGrd.Text
               txtacnt.Visible = True
               txtacnt.SetFocus
            End If
    End Select
End Sub

Private Sub ItemGrd_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox "You cant delete any row !! Delete from Inventory if required"
End Sub

Private Sub ItemGrd_Scroll()
'SendKeys "{TAB}", True
End Sub

Private Sub lblsupl_Validate(Cancel As Boolean)

Dim Party As ADODB.Recordset
If Len(Trim(lblsupl.BoundText)) = 0 Then
   MsgBox "Supplier should not be blank "
   Exit Sub
End If
Set Party = MHVDB.Execute("SELECT * FROM supplier WHERE suplCODE='" & lblsupl.BoundText & "'")
If Party.EOF Then
   MsgBox "Not a valid Supplier !!!"
   lblsupl.ToolTipText = "Enter a valid Supplier"
   lblsupl.Tag = ""
   Cancel = True
Else
   lblsupl = Party!Name
   lblsupl.ToolTipText = Party!Name + " (" + Party!suplcode + ")"
   lblsupl.Tag = Party!Name
End If
Set Party = Nothing
End Sub
Private Sub mnuadd_Click()
Dim lastbill As ADODB.Recordset
LBLnOW = Format(Now, "dd/mm/yyyy")
ValidRow = True
Operation = "ADD"
CurrRow = 1
txtRemark = ""
txtNoTbl = ""
txtpax = ""
txtBillTo = ""
ltot = 0
ErrCTR = 0
txtChDate(0) = Format(Now, "dd/mm/yyyy")
'txtChDate(1) = Format(Now, "dd/mm/yyyy")
cboItemCode.Visible = False
CboItemDesc.Visible = False
txtQty.Visible = False
Set lastbill = MHVDB.Execute("select max(billno) as lno from tranhdr where procyear='" & SysYear & "' and billtype='EN' ", dbOpenForwardOnly)
CboBillNo = IIf(IsNull(lastbill!lno), 1, lastbill!lno + 1)
Set lastbill = Nothing
CboBillNo.Enabled = False
Frame2.Enabled = True
ItemGrd.Enabled = True
ItemGrd.Clear
ItemGrd.FormatString = fmString

TB.Buttons(4).Enabled = True

TB.Buttons(3).Enabled = False
End Sub
Private Sub mnuCancel_Click()
Dim UpdtStr
Dim jrec As ADODB.Recordset
If MsgBox("Cancel it !!!Are u Sure ?", vbYesNo) = vbNo Then Exit Sub
On Error GoTo ERR
MHVDB.BeginTrans
UpdtStr = "UPDATE  tranhdr SET STATUS = 'C',REMARKs = '" & txtRemark & "' WHERE  procyear='" & SysYear & "' and billno = VAL('" & CboBillNo & "') AND billtype = 'EN'"
MHVDB.Execute UpdtStr, dbSeeChanges + dbFailOnError
Set jrec = MHVDB.Execute("select * from tranfile where  procyear='" & SysYear & "' and billno =val('" & CboBillNo & "') AND billtype = 'EN'", dbOpenDynaset)
With jrec
Do While Not .EOF
   MHVDB.Execute "update ITEMSTOCK set totpur=totpur-val('" & !qty & "') where procyear='" & SysYear & "' and ITEMCODE = '" & !itemcode & "'", dbFailOnError
   .MoveNext
Loop
End With
Frame2.Enabled = False
MHVDB.CommitTrans
DatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False
mnuCancel.Enabled = False
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
'ItemGrd.Enabled = False
Frame2.Enabled = True
CboBillNo.Enabled = True

TB.Buttons(2).Enabled = True

'TB.Buttons(4).Enabled = 0
sT1.SimpleText = ""
ErrCTR = 0
CboBillNo.Refresh
End Sub
Private Sub mnuSave_Click()
Dim i, j, K As Integer
Dim TotBill, TotFr, TotPh, jtot, JRate, Jamt As Double
Dim printNow, Final As Boolean
Dim jrec As ADODB.Recordset
Dim InsStr, JStat, pcODE, Msg As String
'If txtQty.Visible Then txtQty_validate
If Not (Operation = "OPEN" Or Operation = "ADD") Then
   Beep
   Exit Sub
End If
If Not ValidRow Then Exit Sub
printNow = True
Msg = ""
If LPh > 0 And Len(DBPH.BoundText) = 0 Then
   MsgBox "Packing Handling Agent Not defined !!!"
   DBPH.SetFocus
   Exit Sub
End If
If Lfr > 0 And Len(DBcboParty.BoundText) = 0 Then
   MsgBox "Transporter Not defined !!!"
   DBcboParty.SetFocus
   Exit Sub
End If
If Len(Trim(lblsupl.Tag)) = 0 Then
   MsgBox "Supplier Not defined !!!"
   lblsupl.SetFocus
   Exit Sub
End If
Final = False
'Final = IIf(MsgBox("Do u want to Finalise and Post this Purchase ?" + Chr(13) + Chr(10) + "(Freight/Packing Handling Charge Bill Number must for Freight/PH Posting)", vbYesNo) = vbYes, True, False)
0:
On Error GoTo ERR

For i = 1 To 994
    If Len(Trim(ItemGrd.TextMatrix(i, 1))) > 0 Then
       If Final And Len(Trim(ItemGrd.TextMatrix(i, 14))) = 0 Then
          MsgBox "Account code for all item not defined ! Record Not saved"
          
          Exit Sub
       End If
       jtot = 0
       JRate = 0
       For j = 6 To 11
           jtot = jtot + Val(ItemGrd.TextMatrix(i, j))
       Next
       If Val(ItemGrd.TextMatrix(i, 4)) > 0 Then
          JRate = jtot / Val(ItemGrd.TextMatrix(i, 4))
       End If
       InsStr = " update tranfile set freight=('" & ItemGrd.TextMatrix(i, 7) & "'),phcharge=('" & ItemGrd.TextMatrix(i, 8) & "'),bst=('" & ItemGrd.TextMatrix(i, 9) & "'),ulc=('" & ItemGrd.TextMatrix(i, 10) & "'), " _
              & " amt=('" & ItemGrd.TextMatrix(i, 6) & "'),rate=('" & ItemGrd.TextMatrix(i, 5) & "'),misc=('" & ItemGrd.TextMatrix(i, 11) & "'),lrate=('" & JRate & "'),acntcode='" & ItemGrd.TextMatrix(i, 14) & "' where " _
              & " procyear='" & SysYear & "' and billtype='EN' and billno=('" & CboBillNo & "') and itemcode='" & ItemGrd.TextMatrix(i, 1) & "'"
       MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError
    Else
       Exit For
    End If
Next
InsStr = "update tranHdR set status='F',suplcode='" & lblsupl.BoundText & "',Frbilldate='" & Format(LBLnOW, "yyyyMMdd") & "',transportercode='" & DBcboParty.BoundText & "', remarks='" & txtRemark & "', " _
       & " phcode='" & DBPH.BoundText & "',Phbillno='" & txtPHBillNo & "',phbilldate='" & Format(txtPHDate, "yyyyMMdd") & "', " _
       & " invno='" & txtInvNo & "',invdate='" & Format(txtChDate(0), "yyyyMMdd") & "',frbillno='" & txtFrBillNo & "' where  procyear='" & SysYear & "' and billno =( '" & CboBillNo & "') and billtype='EN'"
MHVDB.Execute InsStr, dbSeeChanges + dbFailOnError


printNow = IIf(MsgBox("Print Now ?", vbYesNo) = vbYes, True, False)
If printNow Then PrintBill
'rsDatBrBill.Refresh
Operation = ""
CboBillNo.Enabled = False
Frame2.Enabled = False

TB.Buttons(2).Enabled = False
TB.Buttons(3).Enabled = True

ErrCTR = 0
Exit Sub
ERR:
ErrCTR = ErrCTR + 1
If ErrCTR > 5 Then
   If DBEngine.Errors.Count > 0 Then
   For Each errLoop In DBEngine.Errors
       MsgBox "Error number: " & errLoop.Number & vbCr & _
       errLoop.Description
   Next errLoop
'Exit Sub
   End If
End If
ERR.Clear

If ErrCTR < 6 Then
   For i = 1 To 1000
       For j = 1 To 9999
       Next
   Next
   GoTo 0
End If
End Sub
Private Sub Tb_ButtonClick(ByVal Button As msComctlLib.Button)
Select Case Button.Key
       Case "ADD"
           mnuadd_Click
       Case "OPEN"
           mnuOpen_Click
       Case "SAVE"
           mnuSave_Click
       Case "DELETE"
          ' mnuCancel_Click
       Case "EXIT"
           Unload Me
End Select
End Sub


Private Sub txtacnt_Validate(Cancel As Boolean)
'Dim ARec As ADODB.Recordset
'Set ARec = MHVDB.Execute("select * from " + aCNTmAST + " where acntcode='" & txtacnt & "'", dbOpenDynaset, DBReadOnly)
'If ARec.EOF Then
'   MsgBox "Account Code Not Found !!"
'   Cancel = True
'   Exit Sub
'Else
'   ItemGrd.TextMatrix(CurrRow, 14) = txtacnt
'End If
'Set ARec = Nothing
'txtacnt.Visible = False
End Sub

Private Sub txtAmt_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtAmt) Or Val(txtAmt) > 0) Then
   Beep
   MsgBox "Enter a valid amount !!!"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   ValidRow = True
End If
End If
prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ItemGrd.TextMatrix(CurrRow, 6) = Val(txtAmt)
CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
ItemGrd.TextMatrix(CurrRow, 12) = ItemGrd.TextMatrix(CurrRow, 12) + CurrAmt - prevamt
If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
   ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
   ItemGrd.TextMatrix(CurrRow, 5) = CurrAmt / Val(ItemGrd.TextMatrix(CurrRow, 4))
Else
   ItemGrd.TextMatrix(CurrRow, 13) = 0
   ItemGrd.TextMatrix(CurrRow, 5) = 0
End If
LLC = Round(LLC + CurrAmt - prevamt, 2)
lbllc.Caption = Format(LLC, "###,##,##,##0.00")
txtAmt.Visible = False

End Sub

Private Sub txtBSt_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtBSt)) Then
      Beep
      MsgBox "Enter a valid Amount"
      ValidRow = False
      Exit Sub
   Else
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 9))
      ItemGrd.TextMatrix(CurrRow, 9) = Val(txtBSt)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 9))
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      LBST = Round(LBST + CurrAmt - prevamt, 2)
      lblbst.Caption = Format(LBST, "######0.00")
      ValidRow = True
   End If
   End If
   txtBSt.Visible = False
   ItemGrd.row = CurrRow
   ItemGrd.Col = 10
   txtUL.Top = ItemGrd.Top + ItemGrd.CellTop
   txtUL = ItemGrd.Text
   txtUL.Visible = True
   txtUL.SetFocus
End If

End Sub

Private Sub txtBSt_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtBSt)) Then
   Beep
   MsgBox "Enter a valid Amount"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 9))
      ItemGrd.TextMatrix(CurrRow, 9) = Val(txtBSt)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 9))
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      LBST = Round(LBST + CurrAmt - prevamt, 2)
      lblbst.Caption = Format(LBST, "######0.00")
      ValidRow = True
End If
End If
txtBSt.Visible = False
End Sub

Private Sub txtChallan_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub txtChDate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub



Private Sub txtmi_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtmi)) Then
      Beep
      MsgBox "Enter a valid Amount"
      ValidRow = False
      Exit Sub
   Else
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 11))
      ItemGrd.TextMatrix(CurrRow, 11) = Val(txtmi)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 11))
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      Lmi = Round(Lmi + CurrAmt - prevamt, 2)
      lblmi.Caption = Format(Lmi, "######0.00")
      ValidRow = True
   End If
   End If
   txtmi.Visible = False
   ItemGrd.TextMatrix(CurrRow, 0) = CurrRow
   CurrRow = CurrRow + 1
   If CurrRow > ItemGrd.Rows - 2 Then
      ItemGrd.Rows = CurrRow + 3
   End If
   ItemGrd.row = CurrRow
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
   ItemGrd.row = CurrRow
   ItemGrd.Col = 5
   txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
   txtRate = ItemGrd.Text
   txtRate.Visible = True
   txtRate.SetFocus
   End If
End If
End Sub

Private Sub txtmi_Validate(Cancel As Boolean)
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtmi)) Then
   Beep
   MsgBox "Enter a valid Amount"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 11))
   ItemGrd.TextMatrix(CurrRow, 11) = Val(txtmi)
   CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 11))
   ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
   If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
      ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
   Else
      ItemGrd.TextMatrix(CurrRow, 13) = 0
   End If
   LLC = Round(LLC + CurrAmt - prevamt, 2)
   lbllc.Caption = Format(LLC, "###,##,##,##0.00")
   Lmi = Round(Lmi + CurrAmt - prevamt, 2)
   lblmi.Caption = Format(Lmi, "######0.00")
   ValidRow = True
End If
End If
txtmi.Visible = False
End Sub

Private Sub txtph_Validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtph)) Then
   Beep
   MsgBox "Enter a valid Amount"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 8))
      ItemGrd.TextMatrix(CurrRow, 8) = Val(txtph)
      CurrAmt = Val(txtph)
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      LPh = Round(LPh + CurrAmt - prevamt, 2)
      LblPh.Caption = Format(LPh, "######0.00")
      ValidRow = True
End If
End If
txtph.Visible = False
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtQty)) Then
      Beep
      MsgBox "Enter a valid Amount"
      ValidRow = False
      Exit Sub
   Else
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 7))
      ItemGrd.TextMatrix(CurrRow, 7) = Val(txtQty)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 7))
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      Lfr = Round(Lfr + CurrAmt - prevamt, 2)
      lblFr.Caption = Format(Lfr, "######0.00")
      ValidRow = True
   End If
   End If
   txtQty.Visible = False
   ItemGrd.row = CurrRow
   ItemGrd.Col = 8
   txtph.Top = ItemGrd.Top + ItemGrd.CellTop
   txtph = ItemGrd.Text
   txtph.Visible = True
   txtph.SetFocus
End If
End Sub



Private Sub txtQty_validate(Cancel As Boolean)
Dim prevamt, CurrAmt As Double
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtQty)) Then
   Beep
   MsgBox "Enter a valid Amount"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 7))
   ItemGrd.TextMatrix(CurrRow, 7) = Val(txtQty)
   CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 7))
   ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
   If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
      ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
   Else
      ItemGrd.TextMatrix(CurrRow, 13) = 0
   End If
   LLC = Round(LLC + CurrAmt - prevamt, 2)
   lbllc.Caption = Format(LLC, "###,##,##,##0.00")
   Lfr = Round(Lfr + CurrAmt - prevamt, 2)
   lblFr.Caption = Format(Lfr, "######0.00")
   ValidRow = True
End If
End If
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
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 6))
      ItemGrd.TextMatrix(CurrRow, 6) = Round(Val(txtRate) * (ItemGrd.TextMatrix(CurrRow, 4)), 2)
      CurrAmt = ItemGrd.TextMatrix(CurrRow, 6)
      ltot = Round(ltot + CurrAmt - prevamt, 2)
      lblTot.Caption = Format(ltot, "###,##,##,##0.00")
      ItemGrd.TextMatrix(CurrRow, 12) = ItemGrd.TextMatrix(CurrRow, 12) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      ValidRow = True
   End If
   End If
   txtRate.Visible = False
   ItemGrd.row = CurrRow
   ItemGrd.Col = 7
   txtQty.Top = ItemGrd.Top + ItemGrd.CellTop
   txtQty = ItemGrd.Text
   txtQty.Visible = True
   txtQty.SetFocus
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
ItemGrd.TextMatrix(CurrRow, 6) = Round(Val(txtRate) * (ItemGrd.TextMatrix(CurrRow, 4)), 2)
CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 6))
ltot = Round(ltot + CurrAmt - prevamt, 2)
lblTot.Caption = Format(ltot, "###,##,##,##0.00")
ItemGrd.TextMatrix(CurrRow, 12) = ItemGrd.TextMatrix(CurrRow, 12) + CurrAmt - prevamt
If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
   ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
Else
   ItemGrd.TextMatrix(CurrRow, 13) = 0
End If
LLC = Round(LLC + CurrAmt - prevamt, 2)
lbllc.Caption = Format(LLC, "###,##,##,##0.00")
txtRate.Visible = False
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ItemGrd.Enabled = True
   CurrRow = 1
   ItemGrd.row = CurrRow
   ItemGrd.TextMatrix(CurrRow, 0) = Chr(174)
   ItemGrd.Col = 5
   txtRate.Top = ItemGrd.Top + ItemGrd.CellTop
   txtRate = ItemGrd.Text
   txtRate.Visible = True
   txtRate.SetFocus
End If
End Sub

Private Sub txtUL_KeyPress(KeyAscii As Integer)
Dim prevamt, CurrAmt As Double
If KeyAscii = 13 Then
   
   If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
   If Not (IsNumeric(txtUL)) Then
      Beep
      MsgBox "Enter a valid Amount"
      ValidRow = False
      Exit Sub
   Else
      prevamt = Val(ItemGrd.TextMatrix(CurrRow, 10))
      ItemGrd.TextMatrix(CurrRow, 10) = Val(txtUL)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 10))
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      LUL = Round(LUL + CurrAmt - prevamt, 2)
      lblul.Caption = Format(LUL, "######0.00")
      ValidRow = True
   End If
   End If
   txtUL.Visible = False
   ItemGrd.row = CurrRow
   ItemGrd.Col = 11
   txtmi.Top = ItemGrd.Top + ItemGrd.CellTop
   txtmi = ItemGrd.Text
   txtmi.Visible = True
   txtmi.SetFocus
End If
End Sub

Private Sub txtUL_Validate(Cancel As Boolean)
If Len(ItemGrd.TextMatrix(CurrRow, 1)) > 0 Then
If Not (IsNumeric(txtUL)) Then
   Beep
   MsgBox "Enter a valid Amount"
   ValidRow = False
   Cancel = True
   Exit Sub
Else
   prevamt = Val(ItemGrd.TextMatrix(CurrRow, 10))
      ItemGrd.TextMatrix(CurrRow, 10) = Val(txtUL)
      CurrAmt = Val(ItemGrd.TextMatrix(CurrRow, 10))
      ItemGrd.TextMatrix(CurrRow, 12) = Val(ItemGrd.TextMatrix(CurrRow, 12)) + CurrAmt - prevamt
      If Val(ItemGrd.TextMatrix(CurrRow, 4)) <> 0 Then
         ItemGrd.TextMatrix(CurrRow, 13) = Round(Val(ItemGrd.TextMatrix(CurrRow, 12)) / Val(ItemGrd.TextMatrix(CurrRow, 4)), 2)
      Else
         ItemGrd.TextMatrix(CurrRow, 13) = 0
      End If
      LLC = Round(LLC + CurrAmt - prevamt, 2)
      lbllc.Caption = Format(LLC, "###,##,##,##0.00")
      LUL = Round(LUL + CurrAmt - prevamt, 2)
      lblul.Caption = Format(LUL, "######0.00")
      ValidRow = True
End If
End If
txtUL.Visible = False
End Sub
