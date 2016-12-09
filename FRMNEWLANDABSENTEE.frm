VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMNEWLANDABSENTEE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NEW LAND REGISTRATION FOR ABSENTEE"
   ClientHeight    =   7725
   ClientLeft      =   4800
   ClientTop       =   1380
   ClientWidth     =   12615
   Icon            =   "FRMNEWLANDABSENTEE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12615
   Begin VB.Frame Frame3 
      Caption         =   "THRAM HOLDER INFORMATION"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   12135
      Begin VB.TextBox txtthramno 
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
         Left            =   1320
         TabIndex        =   23
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtthramholdername 
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
         Left            =   8640
         TabIndex        =   22
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox cborelation 
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
         ItemData        =   "FRMNEWLANDABSENTEE.frx":0E42
         Left            =   8640
         List            =   "FRMNEWLANDABSENTEE.frx":0E7C
         TabIndex        =   21
         Top             =   1080
         Width           =   3375
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
         ItemData        =   "FRMNEWLANDABSENTEE.frx":0F3C
         Left            =   1320
         List            =   "FRMNEWLANDABSENTEE.frx":0F46
         TabIndex        =   20
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "THRAM NO."
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
         TabIndex        =   27
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "THRAM HOLDER NAME"
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
         Left            =   5400
         TabIndex        =   26
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "RELATIONSHIP OF FARMER  TO THRAM HOLDER"
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
         TabIndex        =   24
         Top             =   1200
         Width           =   4425
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
      TabIndex        =   17
      Top             =   6360
      Width           =   12135
      Begin VB.TextBox txtremarks 
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
         Height          =   735
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   12015
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
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   12135
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
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cbolandtype 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FRMNEWLANDABSENTEE.frx":0F58
         Left            =   8640
         List            =   "FRMNEWLANDABSENTEE.frx":0F62
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "REGISTERED LAND ACRE"
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
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "LAND TYPE"
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
         Left            =   7080
         TabIndex        =   15
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "GENERAL INFORMATION"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   12135
      Begin MSDataListLib.DataCombo cbofarmerid 
         Bindings        =   "FRMNEWLANDABSENTEE.frx":0F6E
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1800
         TabIndex        =   6
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSComCtl2.DTPicker txtregdate 
         Height          =   375
         Left            =   8520
         TabIndex        =   7
         Top             =   600
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
         Format          =   83427329
         CurrentDate     =   41208
      End
      Begin MSDataListLib.DataCombo cbotrnid 
         Bindings        =   "FRMNEWLANDABSENTEE.frx":0F83
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ABSENTEE ID"
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
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "REGISTRATION DATE"
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
         Left            =   6480
         TabIndex        =   10
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERSON REGISTERING"
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
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   12135
      Begin MSDataListLib.DataCombo cbocgid 
         Bindings        =   "FRMNEWLANDABSENTEE.frx":0F98
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSDataListLib.DataCombo cbomhvstaff 
         Bindings        =   "FRMNEWLANDABSENTEE.frx":0FAD
         DataField       =   "ItemCode"
         Height          =   360
         Left            =   8520
         TabIndex        =   2
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MHV STAFF"
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
         Left            =   6480
         TabIndex        =   4
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ID OF CG"
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
         TabIndex        =   3
         Top             =   360
         Width           =   825
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDABSENTEE.frx":0FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDABSENTEE.frx":135C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDABSENTEE.frx":16F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDABSENTEE.frx":23D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDABSENTEE.frx":2822
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMNEWLANDABSENTEE.frx":2FDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "THRAM NO."
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
      TabIndex        =   30
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label LBLDESC 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   360
      TabIndex        =   29
      Top             =   2520
      Width           =   45
   End
End
Attribute VB_Name = "FRMNEWLANDABSENTEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfr As New ADODB.Recordset
Dim rsTr As New ADODB.Recordset
Dim rsCg As New ADODB.Recordset
Dim rsMs As New ADODB.Recordset
Dim FrName, CGname, MHVName As String

Private Sub cbotrnid_LostFocus()
On Error Resume Next
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tbllandregabsentee where trnid='" & cbotrnid.BoundText & "'", MHVDB
If rs.EOF <> True Then
FindFarmer rs!absenteeid


cbofarmerid.Text = rs!absenteeid & " " & FrName
txtthramno.Text = IIf(IsNull(rs!thramno), "", rs!thramno)
txtthramholdername.Text = IIf(IsNull(rs!thramname), "", rs!thramname)
If rs!sex = 0 Then
cbosex.Text = "Male"
ElseIf rs!sex = 1 Then
cbosex.Text = "Female"
Else
cbosex.Text = ""
End If
cborelation.Text = IIf(IsNull(rs!Relation), "", rs!Relation)
txtregland.Text = Format(IIf(IsNull(rs!regland), 0, rs!regland), "####0.00")

cbolandtype.Text = rs!LANDTYPE
If rs!cgid <> "" Then
FindCG rs!cgid
cbocgid.Text = rs!cgid & " " & CGname
Else
cbocgid.Text = ""
End If
If rs!mhvstaff <> "" Then
FindMHV rs!mhvstaff
cbomhvstaff.Text = rs!mhvstaff & " " & MHVName
End If
txtremarks.Text = rs!remarks


Else
MsgBox "No Records Found."
End If
End Sub

Private Sub Form_Load()
OPERATION = ""
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString

Set rsTr = Nothing
If rsTr.State = adStateOpen Then rsTr.Close
rsTr.Open "select concat(cast(trnid as char) ,' ', farmerid,' ',absenteename) as farmername,trnid  from tbllandregabsentee as a,tblabsentee as b where a.farmerid=b.absenteeid order by trnid", db
Set cbotrnid.RowSource = rsTr
cbotrnid.ListField = "farmername"
cbotrnid.BoundColumn = "trnid"



Set rsfr = Nothing
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(absenteeid , ' ', absenteename) as absenteename,absenteeid  from tblabsentee order by absenteeid", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "absenteename"
cbofarmerid.BoundColumn = "absenteeid"


Set rsCg = Nothing
If rsCg.State = adStateOpen Then rsCg.Close
rsCg.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbocgid.RowSource = rsCg
cbocgid.ListField = "farmername"
cbocgid.BoundColumn = "idfarmer"


Set rsMs = Nothing
If rsMs.State = adStateOpen Then rsMs.Close
rsMs.Open "select concat(staffcode , ' ', staffname) as staffname ,staffcode  from tblmhvstaff order by staffcode", db
Set cbomhvstaff.RowSource = rsMs
cbomhvstaff.ListField = "staffname"
cbomhvstaff.BoundColumn = "staffcode"
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "ADD"
       cbofarmerid.Enabled = True
        TB.Buttons(3).Enabled = True
       OPERATION = "ADD"
       CLEARCONTROLL
        cbotrnid.Enabled = False
   Dim rs As New ADODB.Recordset
   Set rs = Nothing
   rs.Open "select (max(trnid)+1) as maxid from tbllandregabsentee", MHVDB
   If rs.EOF <> True Then
   cbotrnid.Text = IIf(IsNull(rs!MaxID), 1, rs!MaxID)
   Else
   cbotrnid.Text = 1
   End If
       
       Case "OPEN"
       OPERATION = "OPEN"
       CLEARCONTROLL
       cbotrnid.Enabled = True
      cbofarmerid.Enabled = False
       'cbogewog.Enabled = True
      TB.Buttons(3).Enabled = True
       
       Case "SAVE"
       MNU_SAVE
        TB.Buttons(3).Enabled = False
       
       
       Case "DELETE"
         
       Case "EXIT"
       Unload Me
       
       
End Select
End Sub
Private Sub CLEARCONTROLL()
cbotrnid.Text = ""
cbofarmerid.Text = ""
txtregdate.Value = "01/01/1900"
txtthramno.Text = ""
txtthramholdername.Text = ""
cbosex.Text = ""
cborelation.Text = ""
txtregland.Text = ""
cbolandtype.Text = ""
cbocgid.Text = ""
cbomhvstaff.Text = ""
txtremarks.Text = ""
End Sub
Private Sub MNU_SAVE()
If OPERATION = "ADD" Then
MHVDB.Execute " insert into tbllandregabsentee (trnid,farmerid,regdate,thramno,thramname,sex,relation,regland,landtype,cgid," _
& " mhvstaff,remarks) " _
& " values('" & cbotrnid.Text & "','" & cbofarmerid.BoundText & "','" & Format(txtregdate.Value, "yyyy-MM-dd") & "','" & txtthramno.Text & "'" _
& " ,'" & txtthramholdername.Text & "','" & cbosex.ListIndex & "','" & cborelation.Text & "','" & Val(txtregland.Text) & "'" _
& " ,'" & cbolandtype.ListIndex & "','" & cbocgid.BoundText & "','" & cbomhvstaff.BoundText & "','" & txtremarks.Text & "')"

ElseIf OPERATION = "OPEN" Then
MHVDB.Execute " update  tbllandregabsentee set farmerid='" & cbofarmerid.BoundText & "',regdate='" & Format(txtregdate.Value, "yyyy-MM-dd") & "', " _
                & "thramno='" & txtthramno.Text & "',thramname='" & txtthramholdername.Text & "',sex='" & cbosex.ListIndex & "'," _
                & "relation='" & cborelation.Text & "',regland='" & Val(txtregland.Text) & "'," _
                & "landtype='" & cbolandtype.ListIndex & "',cgid='" & cbocgid.BoundText & "'," _
                & " mhvstaff='" & cbomhvstaff.BoundText & "',remarks='" & txtremarks.Text & "'  where trnid='" & cbotrnid.Text & "' "




Else
MsgBox "OPERATION NOT SELECTED."
End If
End Sub
Private Sub FindFarmer(ff As String)
FrName = ""
Dim rs As New ADODB.Recordset
Set rs = Nothing
rs.Open "select * from tblabsentee where absenteeid='" & ff & "'", MHVDB
If rs.EOF <> True Then
FrName = rs!FARMERNAME
Else
MsgBox "Record Not Found."
End If
End Sub
Private Sub FindCG(dd As String)
Dim rs As New ADODB.Recordset
CGname = ""
Set rs = Nothing
rs.Open "select * from tblfarmer where idfarmer='" & dd & "'", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
CGname = rs!FARMERNAME
Else
MsgBox "Record Not Found."
End If

End Sub
Private Sub FindMHV(dd As String)
Dim rs As New ADODB.Recordset
MHVName = "" = ""

Set rs = Nothing
rs.Open "select * from tblmhvstaff where staffcode='" & dd & "' ", MHVDB, adOpenForwardOnly, adLockOptimistic
If rs.EOF <> True Then
MHVName = rs!staffname
Else
MsgBox "Record Not Found."
End If

End Sub


Private Sub txtregland_KeyPress(KeyAscii As Integer)
If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtthramholdername_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = txtthramholdername.SelStart + 1
    Dim sText As String
    sText = Left$(txtthramholdername.Text, iPos)
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
