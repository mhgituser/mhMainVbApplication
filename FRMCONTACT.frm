VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMCONTACT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CONTACT"
   ClientHeight    =   6000
   ClientLeft      =   4635
   ClientTop       =   1185
   ClientWidth     =   13755
   Icon            =   "FRMCONTACT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   13755
   Begin VB.Frame Frame4 
      Caption         =   "CONTACT DETAILS"
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
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "setetertert"
      Top             =   2640
      Width           =   13695
      Begin VB.TextBox txtnotes 
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
         TabIndex        =   30
         Top             =   2760
         Width           =   11055
      End
      Begin VB.TextBox txtdept 
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
         TabIndex        =   29
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtmobno 
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
         Left            =   11640
         TabIndex        =   28
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtsecondname 
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
         Left            =   8040
         TabIndex        =   27
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txthomeno 
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
         Left            =   8040
         TabIndex        =   26
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtlocation 
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
         Left            =   8040
         TabIndex        =   25
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox txtrelatives 
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
         Left            =   8040
         TabIndex        =   24
         Top             =   2280
         Width           =   5535
      End
      Begin VB.TextBox txtemail 
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
         TabIndex        =   23
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtworkno 
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
         TabIndex        =   22
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtfirstname 
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
         TabIndex        =   21
         Top             =   840
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo CBOCONTACT 
         Bindings        =   "FRMCONTACT.frx":0E42
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
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
         Caption         =   "CONTACT ID"
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
         TabIndex        =   31
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "IMPORTANT NOTES:"
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
         Width           =   1860
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "IMPORTAINT RELATIVES"
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
         Left            =   5040
         TabIndex        =   19
         Top             =   2400
         Width           =   2250
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL"
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
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "HOUSE LOCATION DESCRIPTION"
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
         Left            =   5040
         TabIndex        =   17
         Top             =   1920
         Width           =   2970
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "DEPARTMENT"
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
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "WORK PHONE NO."
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
         Top             =   1440
         Width           =   1680
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "HOME PHONE NO."
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
         Left            =   5040
         TabIndex        =   14
         Top             =   1440
         Width           =   1650
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "MOBILE NO."
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
         Left            =   10560
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SECOND NAME"
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
         Left            =   5040
         TabIndex        =   12
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FIRST NAME"
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
         Top             =   960
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ROLE"
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
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   13575
      Begin MSDataListLib.DataCombo CBOROLE 
         Bindings        =   "FRMCONTACT.frx":0E57
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ROLE"
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
         TabIndex        =   9
         Top             =   360
         Width           =   510
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
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   13575
      Begin MSDataListLib.DataCombo cboDzongkhag 
         Bindings        =   "FRMCONTACT.frx":0E6C
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   2040
         TabIndex        =   1
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
      Begin MSDataListLib.DataCombo cbogewog 
         Bindings        =   "FRMCONTACT.frx":0E81
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   6000
         TabIndex        =   2
         Top             =   240
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
         Bindings        =   "FRMCONTACT.frx":0E96
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   10080
         TabIndex        =   3
         Top             =   240
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1185
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
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   720
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
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   0
      Top             =   1200
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
            Picture         =   "FRMCONTACT.frx":0EAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCONTACT.frx":1245
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCONTACT.frx":15DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCONTACT.frx":22B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCONTACT.frx":270B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCONTACT.frx":2EC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCONTACT.frx":325F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
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
End
Attribute VB_Name = "FRMCONTACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsDz As New ADODB.Recordset
Dim rsGe As New ADODB.Recordset
Dim rsCON As New ADODB.Recordset
Dim rsTs As New ADODB.Recordset
Dim rsROLE As New ADODB.Recordset
Dim mgraytype As String

Dim AdmLoc As String
Dim id As String

Private Sub CBOCONTACT_LostFocus()
On Error GoTo ERR
CBOCONTACT.Enabled = False
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblcontact where contactid='" & CBOCONTACT.BoundText & "'", MHVDB
If rs.EOF <> True Then

FindROLL rs!roleid
CBOROLE.Text = rs!roleid & " " & rOLEnAME

FindDZ Mid(rs!CONTACTID, 1, 3)
cboDzongkhag.Text = Mid(rs!CONTACTID, 1, 3) & " " & Dzname

If (Mid(rs!CONTACTID, 4, 1)) = "G" Then
FindGE Mid(rs!CONTACTID, 1, 3), Mid(rs!CONTACTID, 4, 3)
cbogewog.Text = Mid(rs!CONTACTID, 4, 3) & " " & GEname
End If

If (Mid(rs!CONTACTID, 4, 1)) = "G" And Mid(rs!CONTACTID, 7, 1) = "T" Then
FindTs Mid(rs!CONTACTID, 1, 3), Mid(rs!CONTACTID, 4, 3), Mid(rs!CONTACTID, 7, 3)
CBOTSHOWOG.Text = Mid(rs!CONTACTID, 7, 3) & " " & TsName
End If



txtfirstname.Text = IIf(IsNull(rs!firstname), "", rs!firstname)
txtsecondname.Text = IIf(IsNull(rs!secondname), "", rs!secondname)
txtworkno.Text = IIf(IsNull(rs!phonework), "", rs!phonework)
txthomeno.Text = IIf(IsNull(rs!phonehome), "", rs!phonehome)
txtmobno.Text = IIf(IsNull(rs!mobile), "", rs!mobile)
txtemail.Text = IIf(IsNull(rs!email), "", rs!email)
txtlocation = IIf(IsNull(rs!Location), "", rs!Location)
txtdept.Text = IIf(IsNull(rs!dept), "", rs!dept)
txtrelatives.Text = IIf(IsNull(rs!relatives), "", rs!relatives)
txtnotes.Text = IIf(IsNull(rs!importaintnote), "", rs!importaintnote)








 
Else

MsgBox "Record Not Found."
End If


Exit Sub
ERR:
MsgBox ERR.Description

End Sub

Private Sub cboDzongkhag_LostFocus()
Dim rs As New ADODB.Recordset
Dim rsmax As New ADODB.Recordset
On Error GoTo ERR




cboDzongkhag.Enabled = False
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsGe = Nothing
If rsGe.State = adStateOpen Then rsGe.Close
rsGe.Open "select concat(gewogid , ' ', gewogname) as gewogname,gewogid  from tblgewog where dzongkhagid='" & cboDzongkhag.BoundText & "' order by dzongkhagid,gewogid", db
Set cbogewog.RowSource = rsGe
cbogewog.ListField = "gewogname"
cbogewog.BoundColumn = "gewogid"

AdmLoc = cboDzongkhag.BoundText

If mgraytype = "11" Then
Set rs = Nothing

rs.Open "select * from tblcontact where substring(contactid,4,1)='C'", MHVDB
If rs.EOF <> True Then
' here

Set rsmax = Nothing
rsmax.Open "select max(substring(contactid,5,2)+1) as MaxId from tblcontact WHERE SUBSTRING(contactid,1,3)='" + AdmLoc + "'", MHVDB, adOpenForwardOnly, adLockOptimistic

If rsmax.EOF <> True Then
    id = IIf(IsNull(rsmax!MaxID), 1, rsmax!MaxID)
    
        If Len(id) = 1 Then
        id = "0" & id
        Else
        
        End If
    
    CBOCONTACT.Text = AdmLoc & "C" & id
Else
    CBOCONTACT.Text = AdmLoc + "C" + "01"
End If


        




' ends here

Else
CBOCONTACT.Text = AdmLoc & "C01"
End If




Else




End If


Exit Sub
ERR:

MsgBox ERR.Description




End Sub

Private Sub cbogewog_LostFocus()
Dim rs As New ADODB.Recordset
Dim rsmax As New ADODB.Recordset
On Error GoTo ERR
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
AdmLoc = cboDzongkhag.BoundText & cbogewog.BoundText

If mgraytype = "01" Then
Set rs = Nothing

rs.Open "select * from tblcontact where substring(contactid,7,1)='C'", MHVDB
If rs.EOF <> True Then



' here

Set rsmax = Nothing
rsmax.Open "select max(substring(contactid,8,2)+1) as MaxId from tblcontact WHERE SUBSTRING(contactid,1,6)='" + AdmLoc + "' and SUBSTRING(contactid,7,1)<>'T'", MHVDB, adOpenForwardOnly, adLockOptimistic

If rsmax.EOF <> True Then
    id = IIf(IsNull(rsmax!MaxID), 1, rsmax!MaxID)
    
        If Len(id) = 1 Then
        id = "0" & id
        Else
        
        End If
    
    CBOCONTACT.Text = AdmLoc & "C" & id
Else
    CBOCONTACT.Text = AdmLoc + "C" + "01"
End If


        




' ends here



Else
CBOCONTACT.Text = AdmLoc & "C01"
End If




Else


End If




Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub CBOROLE_GotFocus()
cboDzongkhag.Enabled = True
End Sub

Private Sub CBOROLE_LostFocus()
Dim rs As New ADODB.Recordset
Set rs = Nothing
cbogewog.Enabled = False
CBOTSHOWOG.Enabled = False
CBOROLE.Enabled = False
AdmLoc = cboDzongkhag.BoundText
rs.Open "select concat(cast(graygewog as char),cast(graytshowog as char)) as mgray from tblrole where roleid='" & CBOROLE.BoundText & "'", MHVDB


If rs!mgray = "11" Then
cbogewog.Enabled = False
CBOTSHOWOG.Enabled = False

ElseIf rs!mgray = "01" Then
cbogewog.Enabled = True
CBOTSHOWOG.Enabled = False
ElseIf rs!mgray = "00" Then

cbogewog.Enabled = True
CBOTSHOWOG.Enabled = True

End If

mgraytype = rs!mgray
End Sub

Private Sub CBOTSHOWOG_LostFocus()
Dim rs As New ADODB.Recordset
Dim rsmax As New ADODB.Recordset
CBOTSHOWOG.Enabled = False
AdmLoc = cboDzongkhag.BoundText & cbogewog.BoundText & CBOTSHOWOG.BoundText
If mgraytype = "00" Then
Set rs = Nothing

rs.Open "select * from tblcontact where substring(contactid,10,1)='C'", MHVDB
If rs.EOF <> True Then
' here

Set rsmax = Nothing
rsmax.Open "select max(substring(contactid,11,2)+1) as MaxId from tblcontact WHERE SUBSTRING(contactid,1,9)='" + AdmLoc + "' ", MHVDB, adOpenForwardOnly, adLockOptimistic

If rsmax.EOF <> True Then
    id = IIf(IsNull(rsmax!MaxID), 1, rsmax!MaxID)
    
        If Len(id) = 1 Then
        id = "0" & id
        Else
        
        End If
    
    CBOCONTACT.Text = AdmLoc & "C" & id
Else
    CBOCONTACT.Text = AdmLoc + "C" + "01"
End If


        




' ends here

Else
CBOCONTACT.Text = AdmLoc & "C01"
End If




Else
MsgBox "Invalid Selection,Please Try Again."

End If
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsROLE = Nothing

If rsROLE.State = adStateOpen Then rsROLE.Close
rsROLE.Open "select concat(ROLEID , ' ', roledescription) as rolename,roleid  from tblrole order by roleid", db
Set CBOROLE.RowSource = rsROLE
CBOROLE.ListField = "rolename"
CBOROLE.BoundColumn = "roleid"



Set rsDz = Nothing

If rsDz.State = adStateOpen Then rsDz.Close
rsDz.Open "select concat(dzongkhagcode , ' ', dzongkhagname) as dzongkhagname,dzongkhagcode  from tbldzongkhag order by dzongkhagcode", db
Set cboDzongkhag.RowSource = rsDz
cboDzongkhag.ListField = "dzongkhagname"
cboDzongkhag.BoundColumn = "dzongkhagcode"



Set rsCON = Nothing

If rsCON.State = adStateOpen Then rsCON.Close
rsCON.Open "select concat(contactid , ' ', firstname,' ' ,secondname) as contactname,contactid  from tblcontact order by contactid", db
Set CBOCONTACT.RowSource = rsCON
CBOCONTACT.ListField = "contactname"
CBOCONTACT.BoundColumn = "contactid"


End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Operation = "ADD" Then
If Len(CBOCONTACT.Text) = 0 Then
CBOCONTACT.ToolTipText = "jagsdjagsdg"
End If

End If
End Sub

Private Sub Tb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ADD"
AdmLoc = ""
      
       TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CBOROLE.Enabled = True
       CLEARCONTROLL
        
         Case "OPEN"
         Operation = "OPEN"
         CLEARCONTROLL
         CBOCONTACT.Enabled = True
         cboDzongkhag.Enabled = False
         cbogewog.Enabled = False
         CBOTSHOWOG.Enabled = False
         CBOROLE.Enabled = False
         TB.Buttons(3).Enabled = True
       
       Case "SAVE"
      
       MNU_SAVE
        
       Case "DELETE"
         Case "PRINT"
         'PRINTFINFO
          TB.Buttons(6).Enabled = False
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub MNU_SAVE()
On Error GoTo ERR

If Len(CBOROLE.Text) = 0 Then Exit Sub
If Len(CBOCONTACT.Text) = 0 Then Exit Sub

If txtfirstname.Text = "" Then
MsgBox "First Name Is Must To Provide."
Exit Sub
End If

If Operation = "ADD" Then

MHVDB.Execute "insert into tblcontact(roleid,dzongkhagid,gewogid,tshowogid,contactid,firstname,secondname," _
            & " phonework,phonehome,mobile,email,location,dept,relatives,importaintnote)" _
            & " values('" & CBOROLE.BoundText & "','" & cboDzongkhag.BoundText & "','" & cbogewog.BoundText & "'" _
            & " ,'" & CBOTSHOWOG.BoundText & "','" & CBOCONTACT.Text & "','" & txtfirstname.Text & "','" & txtsecondname.Text & "'" _
            & " ,'" & txtworkno.Text & "','" & txthomeno.Text & "','" & txtmobno.Text & "','" & txtemail.Text & "'" _
            & " ,'" & txtlocation.Text & "','" & txtdept.Text & "','" & txtrelatives.Text & "','" & txtnotes.Text & "')"

ElseIf Operation = "OPEN" Then

MHVDB.Execute "update tblcontact" _
            & " set roleid='" & CBOROLE.BoundText & "'" _
            & " ,firstname='" & txtfirstname.Text & "',secondname='" & txtsecondname.Text & "'" _
            & " ,phonework='" & txtworkno.Text & "',phonehome='" & txthomeno.Text & "',mobile='" & txtmobno.Text & "',email='" & txtemail.Text & "'" _
            & " ,location='" & txtlocation.Text & "',dept='" & txtdept.Text & "',relatives='" & txtrelatives.Text & "',importaintnote='" & txtnotes.Text & "' where contactid='" & CBOCONTACT.BoundText & "'"




Else
MsgBox "Invalid Operation."
End If

TB.Buttons(3).Enabled = False
TB.Buttons(6).Enabled = True
       
Exit Sub
ERR:
MsgBox ERR.Description



End Sub
Private Sub CLEARCONTROLL()
CBOROLE.Text = ""
cboDzongkhag.Text = ""
cbogewog.Text = ""
CBOTSHOWOG.Text = ""
CBOCONTACT.Text = ""
txtfirstname.Text = ""
txtsecondname.Text = ""
txtworkno.Text = ""
txthomeno.Text = ""
txtmobno.Text = ""
txtemail.Text = ""
txtlocation.Text = ""
txtdept.Text = ""
txtrelatives.Text = ""
txtnotes.Text = ""
End Sub

Private Sub txtdept_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtdept.SelStart + 1
    Dim sText As String
    sText = Left$(txtdept.Text, iPos)
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

Private Sub txtfirstname_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtfirstname.SelStart + 1
    Dim sText As String
    sText = Left$(txtfirstname.Text, iPos)
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

Private Sub txthomeno_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789+,", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
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

Private Sub txtmobno_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789+,", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtrelatives_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtrelatives.SelStart + 1
    Dim sText As String
    sText = Left$(txtrelatives.Text, iPos)
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

Private Sub txtsecondname_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtsecondname.SelStart + 1
    Dim sText As String
    sText = Left$(txtsecondname.Text, iPos)
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

Private Sub txtworkno_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789+,", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
