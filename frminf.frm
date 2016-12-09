VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frminf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I N F L  U E N T I AL       P E R S O N "
   ClientHeight    =   3405
   ClientLeft      =   5595
   ClientTop       =   1395
   ClientWidth     =   10455
   Icon            =   "frminf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   10455
   Begin VB.Frame Frame1 
      Caption         =   "INFUENTIAL INFORATION"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   10335
      Begin VB.TextBox TXTINFLUENTILPERSON 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   0
         Top             =   360
         Width           =   6375
      End
      Begin VB.TextBox TXTRELATIVES 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   1320
         Width           =   6375
      End
      Begin VB.TextBox TXTDEPT 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   840
         Width           =   6375
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   2250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DEPARTMENT/COMPANY/INSTITUTION"
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
         TabIndex        =   6
         Top             =   960
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "JOB TITLE"
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
         TabIndex        =   5
         Top             =   480
         Width           =   945
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   9480
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
            Picture         =   "frminf.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminf.frx":11DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminf.frx":1576
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminf.frx":2250
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminf.frx":26A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frminf.frx":2E5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
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
   Begin MSDataListLib.DataCombo cbofarmerid 
      Bindings        =   "frminf.frx":31F6
      DataField       =   "ItemCode"
      Height          =   360
      Left            =   3960
      TabIndex        =   9
      Top             =   720
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
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "frminf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfr As New ADODB.Recordset
Private Sub CBOJOBTITLE_Click()
If CBOJOBTITLE.Text = "OTHERS" Then
'CBOJOBTITLE.Enabled = False
TXTTITLE.Visible = True
TXTTITLE.SetFocus
Else
'CBOJOBTITLE.Enabled = True
TXTTITLE.Text = ""
TXTTITLE.Visible = False
End If
End Sub

Private Sub Form_Load()
Label2.Caption = UCase(Label2.Caption)
Label3.Caption = UCase(Label3.Caption)
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open CnnString
Set rsfr = Nothing
If FATYPEINF = "F" Then
If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(idfarmer , ' ', farmername) as farmername,idfarmer  from tblfarmer order by idfarmer", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "farmername"
cbofarmerid.BoundColumn = "idfarmer"
ElseIf FATYPEINF = "A" Then

If rsfr.State = adStateOpen Then rsfr.Close
rsfr.Open "select concat(ABSENTEEID , ' ', ABSENTEENAME) as ABSENTEENAME,ABSENTEEID  from TBLABSENTEE order by ABSENTEEID", db
Set cbofarmerid.RowSource = rsfr
cbofarmerid.ListField = "ABSENTEENAME"
cbofarmerid.BoundColumn = "ABSENTEEID"

Else
MsgBox "INFLUENTIAL SELECTION IS NOT CORRECT."
Exit Sub

End If

If mbypass = True Then

        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       'CLEARCONTROLL
       'cboDzongkhag.Enabled = True
     cbofarmerid.Enabled = False
       
      
cbofarmerid.Text = Mcaretaker
'CBOCARETAKER.Enabled = False
Else

'CBOCARETAKER.Enabled = True
End If

End Sub

Private Sub Tb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
Case "ADD"
       
        TB.Buttons(3).Enabled = True
       Operation = "ADD"
       CLEARCONTROLL
       
     cbofarmerid.Enabled = False
       
       Case "OPEN"
       Operation = "OPEN"
       CLEARCONTROLL
       cbofarmerid.Enabled = True
       
      TB.Buttons(3).Enabled = True
       
       Case "SAVE"
       If Len(cbofarmerid.Text) = 0 Then
        MsgBox "PLEASE SELECT CARETAKER FOR THIS ABSENTEE."
        cbofarmerid.SetFocus
        Exit Sub
        End If
       MNU_SAVE
        TB.Buttons(3).Enabled = False
        'FillGrid
       
       Case "DELETE"
         Case "PRINT"
         'PRINTINFO
          
       Case "EXIT"
       Unload Me
       
       
End Select

End Sub
Private Sub MNU_SAVE()
On Error GoTo ERR
If Len(cbofarmerid.Text) = 0 Then
MsgBox "Please Select The Appropriate Information From The Drop Down Controll.", , "MHV ERROR BOX"
Exit Sub
End If
MHVDB.BeginTrans
If Operation = "ADD" Then
MHVDB.Execute "INSERT INTO tblinfluential(FARMERID,JOBTITLE,DEPT,RELATION,FATYPE) VALUES('" & cbofarmerid.Text & "','" & TXTINFLUENTILPERSON.Text & "','" & txtdept.Text & "','" & txtrelatives.Text & "','" & FATYPEINF & "')"


ElseIf Operation = "OPEN" Then

MHVDB.Execute "UPDATE tblinfluential SET JOBTITLE='" & TXTINFLUENTILPERSON.Text & "',DEPT='" & txtdept.Text & "',RELATION='" & txtrelatives.Text & "' WHERE FARMERID='" & cboabsenteeid.BoundText & "' AND FATYPE='" & FATYPEINF & "'"

' Val(txtregland.Text)


Else
MsgBox "OPERATION NOT SELECTED."
End If
MHVDB.CommitTrans
mbypass = False
Mcaretaker = ""
Exit Sub
ERR:
MsgBox ERR.Description
MHVDB.RollbackTrans
End Sub
Private Sub CLEARCONTROLL()
TXTINFLUENTILPERSON.Text = ""
If mbypass = False Then
cbofarmerid.Text = ""
End If
txtdept.Text = ""
txtrelatives.Text = ""
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

Private Sub TXTINFLUENTILPERSON_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    Dim iPos As Integer
    iPos = TXTINFLUENTILPERSON.SelStart + 1
    Dim sText As String
    sText = Left$(TXTINFLUENTILPERSON.Text, iPos)
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
