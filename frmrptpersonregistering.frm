VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmrptpersonregistering 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P E R S O N R E G I S T E R I N G   R E P O R T"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   Icon            =   "frmrptpersonregistering.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SHOW"
      Height          =   735
      Left            =   3000
      Picture         =   "frmrptpersonregistering.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   4200
      Picture         =   "frmrptpersonregistering.frx":15AC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4920
      Picture         =   "frmrptpersonregistering.frx":2276
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECTION OPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9375
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   4920
         TabIndex        =   12
         Top             =   2520
         Visible         =   0   'False
         Width           =   4335
         Begin MSDataListLib.DataCombo CBOSTAFF 
            Bindings        =   "frmrptpersonregistering.frx":2600
            DataField       =   "ItemCode"
            Height          =   360
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
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
      End
      Begin VB.Frame Frame8 
         Caption         =   "REGISTRATION TYPE"
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
         Height          =   4215
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   4695
         Begin VB.OptionButton Option7 
            Caption         =   "LEAD STAFF (MEETING REGISTRATION)"
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
            Top             =   2160
            Width           =   4095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "ALL REGISTRATION TYPE"
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
            Top             =   360
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.OptionButton Option12 
            Caption         =   "SUPPORT1"
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
            Top             =   2880
            Width           =   3375
         End
         Begin VB.OptionButton Option11 
            Caption         =   "SUPPORT2"
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
            Top             =   3240
            Width           =   4095
         End
         Begin VB.OptionButton Option10 
            Caption         =   "SUPPORT3"
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
            Top             =   3600
            Width           =   4095
         End
         Begin VB.OptionButton Option9 
            Caption         =   "SUPPORT4"
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
            Width           =   4095
         End
         Begin VB.OptionButton Option8 
            Caption         =   "INDIVIDUAL"
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
            Top             =   2520
            Width           =   2895
         End
         Begin VB.OptionButton Option6 
            Caption         =   "CG REGISTRATION"
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
            Top             =   1800
            Width           =   2175
         End
         Begin VB.OptionButton Option5 
            Caption         =   "MONITOR(CG REGISTRATION)"
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
            TabIndex        =   14
            Top             =   1440
            Width           =   3135
         End
         Begin VB.OptionButton Option4 
            Caption         =   "MONITOR(SHARED REGISTRATION)"
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
            Top             =   720
            Width           =   3735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "OUTREASH STAFF"
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
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4695
         Begin VB.OptionButton OPTDETAIL 
            Caption         =   "DETAIL"
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
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OPTSUMMARY 
            Caption         =   "SUMMARY"
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
            Left            =   1440
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OPTINDIVIDUAL 
            Caption         =   "INDVIDUAL"
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
            Left            =   2880
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   45
   End
End
Attribute VB_Name = "frmrptpersonregistering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSSTAFF As New ADODB.Recordset
Private Sub Command1_Click()
Select Case RptOption
Case "MSR"
MSR
Case "ALL"
all

Case "ORS"
ORS
Case "MCGR"
MCGR
Case "CGR"
CGR
Case "LS"
LS
Case "IN"
IND
Case "S1"
S1
Case "S2"
S2
Case "S3"
S3
Case "S4"
S4


End Select

End Sub
Private Sub S4()


End Sub


Private Sub S3()


End Sub


Private Sub S2()

End Sub

Private Sub S1()


End Sub
Private Sub IND()


End Sub
Private Sub LS()


End Sub
Private Sub CGR()

End Sub
Private Sub MCGR()


End Sub
Private Sub ORS()


End Sub
Private Sub all()

End Sub
Private Sub MSR()







End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set RSSTAFF = Nothing
If RSSTAFF.State = adStateOpen Then RSSTAFF.Close
RSSTAFF.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff  order by staffcode", db
Set CBOSTAFF.RowSource = RSSTAFF
CBOSTAFF.ListField = "staffname"
CBOSTAFF.BoundColumn = "staffcode"
End Sub

Private Sub OPTDETAIL_Click()
Frame3.Visible = False
Option1.Visible = False
End Sub

Private Sub OPTINDIVIDUAL_Click()
Frame3.Visible = True
Option1.Visible = True
End Sub

Private Sub Option1_Click()
RptOption = ""
RptOption = "ALL"
End Sub

Private Sub Option10_Click()
RptOption = ""
RptOption = "S3"
End Sub

Private Sub Option11_Click()
RptOption = ""
RptOption = "S2"
End Sub

Private Sub Option12_Click()
RptOption = ""
RptOption = "S1"
End Sub

Private Sub Option3_Click()
RptOption = ""
RptOption = "ORS"
End Sub

Private Sub Option4_Click()
RptOption = ""
RptOption = "MSR"
End Sub

Private Sub Option5_Click()
RptOption = ""
RptOption = "MCGR"
End Sub

Private Sub Option6_Click()
RptOption = ""
RptOption = "CGR"
End Sub

Private Sub Option7_Click()
RptOption = ""
RptOption = "LS"
End Sub

Private Sub Option8_Click()
RptOption = ""
RptOption = "IN"
End Sub

Private Sub Option9_Click()
RptOption = ""
RptOption = "S4"
End Sub

Private Sub OPTSUMMARY_Click()
Frame3.Visible = False
Option1.Visible = False
End Sub
