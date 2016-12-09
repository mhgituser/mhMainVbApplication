VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptpersonregindall 
   Caption         =   "REPORT"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "rptpersonregindall.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "rptpersonregindall.dsx":0E42
End
Attribute VB_Name = "rptpersonregindall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MCOUNT As Integer

Private Sub ActiveReport_ReportStart()
MCOUNT = 0
End Sub

Private Sub Detail_Format()
FindFA txtFARMERID.Text, "F"
TXTFARMERNAME.Text = FAName
MCOUNT = MCOUNT + 1
End Sub

Private Sub GroupHeader1_Format()
FindsTAFF txtSTAFFCODE.Text
TXTSTAFFNAME.Text = sTAFF
End Sub

Private Sub PageHeader_Format()
Field20.Text = "INDIVIDUAL STAFF REGISTRATION TYPE INFORMATION"
End Sub

Private Sub ReportFooter_Format()
TXTTOTALFERMER.Text = MCOUNT
End Sub
