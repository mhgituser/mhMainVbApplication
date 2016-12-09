VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMFARMERQUESTION 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F A R M E R   Q U E S T I O N A R I E S "
   ClientHeight    =   8880
   ClientLeft      =   5145
   ClientTop       =   1050
   ClientWidth     =   10845
   Icon            =   "FRMFARMERQUESTION.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10845
   Begin VB.Frame Frame1 
      Caption         =   "ABSENTEE INFORMATION"
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
      Height          =   2655
      Left            =   0
      TabIndex        =   24
      Top             =   720
      Width           =   10695
      Begin VB.TextBox TXTTSHOWOG 
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
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox TXTNAME 
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
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   8055
      End
      Begin VB.TextBox TXTGEWOG 
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
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2160
         Width           =   3615
      End
      Begin VB.TextBox TXTDZONGKHAG 
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
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1680
         Width           =   8055
      End
      Begin VB.TextBox TXTCONTACT 
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
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   8055
      End
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "FRMFARMERQUESTION.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo CBOTRANID 
         Bindings        =   "FRMFARMERQUESTION.frx":0E57
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   7920
         TabIndex        =   30
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label16 
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
         Left            =   6240
         TabIndex        =   38
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LOOK UP FOR FARMER"
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
         TabIndex        =   36
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label3 
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
         TabIndex        =   34
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label4 
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
         TabIndex        =   33
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CONTACT NO."
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
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TRANSACTION ID"
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
         Left            =   6240
         TabIndex        =   31
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "QUESTION"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   10575
      Begin VB.CommandButton cmdotherquestion 
         Height          =   615
         Left            =   9480
         Picture         =   "FRMFARMERQUESTION.frx":0E6C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "SEE OTHER QUESTIONS FOR THIS ABSENTEE"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtquestion 
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
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   9255
      End
      Begin MSComCtl2.DTPicker txtquestiondate 
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   240
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
         Format          =   81920001
         CurrentDate     =   41208
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DATE"
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
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "FOLOW UP"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   10575
      Begin VB.TextBox txtanswer 
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
         Height          =   735
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtinfremarks 
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
         Left            =   4920
         TabIndex        =   7
         Top             =   2640
         Width           =   5535
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   4920
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
         Begin VB.OptionButton optansyes 
            Caption         =   "YES"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optansno 
            Caption         =   "NO"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Height          =   600
         Left            =   4920
         TabIndex        =   1
         Top             =   3120
         Width           =   1575
         Begin VB.OptionButton optanscompleteno 
            Caption         =   "NO"
            Height          =   255
            Left            =   840
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optanscompleteyes 
            Caption         =   "YES"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
      End
      Begin MSComCtl2.DTPicker txtfollowupdate 
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   240
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
         Format          =   81920001
         CurrentDate     =   41208
      End
      Begin MSDataListLib.DataCombo cbostaffcode 
         Bindings        =   "FRMFARMERQUESTION.frx":151E
         DataField       =   "ItemCode"
         Height          =   315
         Left            =   4920
         TabIndex        =   10
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
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
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PERSON RESPONSIBLE FOR FOLLOW UP"
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
         TabIndex        =   16
         Top             =   840
         Width           =   3705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "FOLLOW UP DATE"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ANSWER"
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
         TabIndex        =   14
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   " TALKED TO INFLUENTIAL AND EXPLAINED ANSWER"
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
         Top             =   2160
         Width           =   4785
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "COMMENTS FROM INFLUENTIAL"
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
         TabIndex        =   12
         Top             =   2640
         Width           =   2910
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "DOES ANSWER COMPLETE AND CLOSED ENQUERY "
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
         TabIndex        =   11
         Top             =   3360
         Width           =   4710
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   600
      Top             =   240
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
            Picture         =   "FRMFARMERQUESTION.frx":1533
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFARMERQUESTION.frx":18CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFARMERQUESTION.frx":1C67
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFARMERQUESTION.frx":2941
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFARMERQUESTION.frx":2D93
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFARMERQUESTION.frx":354D
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
      Width           =   10845
      _ExtentX        =   19129
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
Attribute VB_Name = "FRMFARMERQUESTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAb As New ADODB.Recordset
Dim Srs As New ADODB.Recordset
Dim RSTR As New ADODB.Recordset
Dim opttalkedtoinf As Integer
Dim optanswercomplete As Integer

Private Sub cboabsenteeid_LostFocus()
On Error GoTo err
Dim rs As New ADODB.Recordset
If Operation = "ADD" Then
cboabsenteeid.Enabled = False
Set rs = Nothing
rs.Open "SELECT MAX(TRNID+1) AS MAXID  FROM tblfarmerQUESTION", MHVDB
If rs.EOF <> True Then
CBOTRANID.Text = IIf(IsNull(rs!MaxID), 1, rs!MaxID)
Else
CBOTRANID.Text = 1
End If
CBOTRANID.Enabled = False
Else

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select *  from tblfarmerQUESTION WHERE FARMERID='" & cboabsenteeid.BoundText & "' order by absenteeid", db
Set CBOTRANID.RowSource = RSTR
CBOTRANID.ListField = "TRNID"
CBOTRANID.BoundColumn = "TRNID"


CBOTRANID.Enabled = True


End If


Set rs = Nothing
rs.Open "SELECT * FROM tblfarmer WHERE IDFARMER='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
TXTNAME.Text = rs!farmername
TXTCONTACT.Text = rs!phone1 & "," & rs!phone2
FindDZ Mid(Trim(cboabsenteeid.BoundText), 1, 3)

TXTDZONGKHAG.Text = Dzname
FindGE Mid(Trim(cboabsenteeid.BoundText), 1, 3), Mid(Trim(cboabsenteeid.BoundText), 4, 3)

TXTGEWOG.Text = GEname

End If






Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub Command1_Click()
frmprequestionabsentee.Show 1
End Sub




Private Sub cbofarmerid_LostFocus()
On Error GoTo err
Dim rs As New ADODB.Recordset
If Operation = "ADD" Then
cbofarmerid.Enabled = False
Set rs = Nothing
rs.Open "SELECT MAX(TRNID+1) AS MAXID  FROM tblfarmerQUESTION", MHVDB
If rs.EOF <> True Then
CBOTRANID.Text = IIf(IsNull(rs!MaxID), 1, rs!MaxID)
Else
CBOTRANID.Text = 1
End If
CBOTRANID.Enabled = False
Else

Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString


Set RSTR = Nothing
If RSTR.State = adStateOpen Then RSTR.Close
RSTR.Open "select *  from tblfarmerQUESTION WHERE FARMERID='" & cbofarmerid.BoundText & "' order by FARMERID", db
Set CBOTRANID.RowSource = RSTR
CBOTRANID.ListField = "TRNID"
CBOTRANID.BoundColumn = "TRNID"


CBOTRANID.Enabled = True


End If


Set rs = Nothing
rs.Open "SELECT * FROM tblfarmer WHERE IDFARMER='" & cbofarmerid.BoundText & "'", MHVDB
If rs.EOF <> True Then
TXTNAME.Text = rs!farmername
TXTCONTACT.Text = rs!phone1 & "," & rs!phone2
FindDZ Mid(Trim(cbofarmerid.BoundText), 1, 3)

TXTDZONGKHAG.Text = Dzname
FindGE Mid(Trim(cbofarmerid.BoundText), 1, 3), Mid(Trim(cbofarmerid.BoundText), 4, 3)
TXTGEWOG.Text = GEname

FindTs Mid(Trim(cbofarmerid.BoundText), 1, 3), Mid(Trim(cbofarmerid.BoundText), 4, 3), Mid(Trim(cbofarmerid.BoundText), 7, 3)
TXTTSHOWOG.Text = TsName

End If






Exit Sub
err:
MsgBox err.Description


End Sub

Private Sub CBOTRANID_LostFocus()
CBOTRANID.Enabled = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "SELECT * FROM tblfarmerQUESTION WHERE FARMERID='" & cbofarmerid.BoundText & "' AND TRNID='" & CBOTRANID.BoundText & "' ", MHVDB
If rs.EOF <> True Then
txtquestion.Text = rs!QUESTION
txtquestiondate.Value = rs!QDATE
txtfollowupdate.Value = rs!FOLLOWUPDATE

FindsTAFF rs!FOLLOWUPSTAFF
cbostaffcode.Text = rs!FOLLOWUPSTAFF & " " & sTAFF

txtanswer.Text = rs!ANSWER
If rs!TALKEDTOINF = 0 Then
optansno.Value = True
Else
optansyes.Value = True
End If
txtinfremarks.Text = rs!INFREMARKS

If rs!ANSWERCOMPLETE = 0 Then
optanscompleteno.Value = True
Else
optanscompleteyes.Value = True
End If


End If
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsAb = Nothing

If rsAb.State = adStateOpen Then rsAb.Close
rsAb.Open "select concat(IDFARMER , ' ', FARMERNAME) as FARMERNAME,IDFARMER  from tblfarmer order by IDFARMER", db
Set cbofarmerid.RowSource = rsAb
cbofarmerid.ListField = "FARMERNAME"
cbofarmerid.BoundColumn = "IDFARMER"

Set Srs = Nothing

If Srs.State = adStateOpen Then Srs.Close
Srs.Open "select concat(STAFFCODE , ' ', STAFFNAME) as STAFFNAME,STAFFCODE  from tblmhvstaff order by STAFFCODE", db
Set cbostaffcode.RowSource = Srs
cbostaffcode.ListField = "STAFFNAME"
cbostaffcode.BoundColumn = "STAFFCODE"
opttalkedtoinf = False
optanswercomplete = False
End Sub

Private Sub optanscompleteno_Click()
optanswercomplete = 0
End Sub

Private Sub optanscompleteyes_Click()
optanswercomplete = 1
End Sub

Private Sub optansno_Click()
opttalkedtoinf = 0
End Sub

Private Sub optansyes_Click()
opttalkedtoinf = 1
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
        Case "ADD"
        FILLCOMBO "ADD"
      cbofarmerid.Enabled = True
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       CBOTRANID.Enabled = False
       Case "OPEN"
        FILLCOMBO "OPEN"
       cbofarmerid.Enabled = True
       Operation = "OPEN"
       CLEARCONTROLL
       cbofarmerid.Enabled = True
       TB.Buttons(3).Enabled = True
       CBOTRANID.Enabled = True
       Case "SAVE"
      If Len(cbofarmerid.Text) = 0 Then
        MsgBox "PLEASE SELECT FARMER FOR THIS QUESTIONARY."
        cbofarmerid.SetFocus
        Exit Sub
        End If
       MNU_SAVE
        TB.Buttons(3).Enabled = False
      
       
       Case "DELETE"
         Case "PRINT"
         
       Case "EXIT"
       Unload Me
       
       
End Select


End Sub
Private Sub FILLCOMBO(op As String)
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsAb = Nothing
If op = "ADD" Then
If rsAb.State = adStateOpen Then rsAb.Close
rsAb.Open "select concat(IDFARMER , ' ', FARMERNAME) as FARMERNAME,IDFARMER  from tblfarmer order by IDFARMER", db
Set cbofarmerid.RowSource = rsAb
cbofarmerid.ListField = "FARMERNAME"
cbofarmerid.BoundColumn = "IDFARMER"
Else

If rsAb.State = adStateOpen Then rsAb.Close
rsAb.Open "select distinct concat(IDFARMER , ' ', FARMERNAME) as FARMERNAME,IDFARMER  from tblfarmer AS A,tblfarmerQUESTION AS B WHERE A.IDFARMER=B.FARMERID  order by IDFARMER", db
Set cbofarmerid.RowSource = rsAb
cbofarmerid.ListField = "FARMERNAME"
cbofarmerid.BoundColumn = "IDFARMER"

End If

opttalkedtoinf = False
optanswercomplete = False
End Sub

Private Sub MNU_SAVE()
On Error GoTo err



If Len(cbofarmerid.Text) = 0 Then
MsgBox "Please Select The Farmer Information From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If
MHVDB.BeginTrans
If Operation = "ADD" Then

MHVDB.Execute "insert into tblfarmerquestion (TRNID,FARMERID,qdate,question,followupdate,followupstaff,answer,talkedtoinf,infremarks,answercomplete)" _
            & "values('" & CBOTRANID.BoundText & "','" & cbofarmerid.BoundText & "','" & Format(txtquestiondate.Value, "yyyyMMdd") & "'" _
            & " , '" & txtquestion.Text & "','" & Format(txtfollowupdate.Value, "yyyyMMdd") & "','" & cbostaffcode.BoundText & "','" & txtanswer.Text & "','" & opttalkedtoinf & "','" & txtinfremarks.Text & "'" _
            & ",'" & optanswercomplete & "')"


ElseIf Operation = "OPEN" Then

MHVDB.Execute "UPDATE  tblfarmerquestion SET qdate='" & Format(txtquestiondate.Value, "yyyyMMdd") & "',question='" & txtquestion.Text & "',followupdate='" & Format(txtfollowupdate.Value, "yyyyMMdd") & "',followupstaff='" & cbostaffcode.BoundText & "',answer='" & txtanswer.Text & "',talkedtoinf='" & opttalkedtoinf & "',infremarks='" & txtinfremarks.Text & "',answercomplete='" & optanswercomplete & "' WHERE FARMERID='" & cbofarmerid.BoundText & "' AND TRNID='" & CBOTRANID.BoundText & "'"
           



Else
MsgBox "OPERATION NOT SELECTED."
End If
TB.Buttons(3).Enabled = False
MHVDB.CommitTrans
mbypass = False
Mcaretaker = ""
Exit Sub
err:
MsgBox err.Description
MHVDB.RollbackTrans

End Sub

Private Sub CLEARCONTROLL()
cbofarmerid.Text = ""
CBOTRANID.Text = ""
TXTNAME.Text = ""
TXTCONTACT.Text = ""
TXTDZONGKHAG.Text = ""
TXTGEWOG.Text = ""
txtquestiondate.Value = Now
txtquestion.Text = ""
txtfollowupdate.Value = Now
cbostaffcode.Text = ""
txtanswer.Text = ""
optansno.Value = 0
optansyes.Value = 0
txtinfremarks.Text = ""
optanscompleteyes.Value = False
optanscompleteno.Value = False
End Sub


