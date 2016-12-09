VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmnurserydashboard 
   Caption         =   "Nursery Dashboard"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "frmnurserydashboard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Summary"
      TabPicture(0)   =   "frmnurserydashboard.frx":0E42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Nursery"
      TabPicture(1)   =   "frmnurserydashboard.frx":0E5E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Text1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         MaxLength       =   1
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   32
         Top             =   5520
         Width           =   8295
         Begin VSFlex7Ctl.VSFlexGrid actiongrid 
            Height          =   1095
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   8175
            _cx             =   14420
            _cy             =   1931
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
            BackColorAlternate=   8438015
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
            Rows            =   9
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmnurserydashboard.frx":0E7A
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
            Begin MSComCtl2.DTPicker dtPick 
               Height          =   315
               Left            =   4080
               TabIndex        =   35
               Top             =   240
               Visible         =   0   'False
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Format          =   78905345
               CurrentDate     =   36473
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   31
         Top             =   3960
         Width           =   8295
         Begin VSFlex7Ctl.VSFlexGrid issuegrid 
            Height          =   1095
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   8175
            _cx             =   14420
            _cy             =   1931
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
            BackColorAlternate=   8438015
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
            Rows            =   4
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmnurserydashboard.frx":0F15
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
      Begin VB.Frame Frame2 
         Caption         =   " Nursery HR Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -71760
         TabIndex        =   22
         Top             =   480
         Width           =   5295
         Begin VB.TextBox txtproblems 
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
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   30
            Top             =   2760
            Width           =   5055
         End
         Begin VB.TextBox txttrainingdetails 
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
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   2040
            Width           =   5055
         End
         Begin VB.TextBox txthiringneeds 
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
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox txttrainingneeds 
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
            Height          =   405
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   600
            Width           =   5055
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Problem/Issues"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Detail of Training conducted"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   2010
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Hiring needs:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Training needs:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "MHV Nursery Dashboard Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   3015
         Begin VB.TextBox txtnutplants 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txthardplants 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtinventory 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox txtmaterials 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txtutilization 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txthealthandsafty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   2520
            Width           =   495
         End
         Begin VB.TextBox txthrneeds 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Top             =   2880
            Width           =   495
         End
         Begin VB.TextBox txttcplants 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   6
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nut Plants"
            Height          =   195
            Left            =   840
            TabIndex        =   21
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hard Plants"
            Height          =   195
            Left            =   840
            TabIndex        =   20
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Materials"
            Height          =   195
            Left            =   960
            TabIndex        =   19
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Inventory"
            Height          =   195
            Left            =   960
            TabIndex        =   18
            Top             =   1800
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Utilization"
            Height          =   195
            Left            =   960
            TabIndex        =   17
            Top             =   2160
            Width           =   675
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Health and Safety"
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "HR Needs"
            Height          =   195
            Left            =   840
            TabIndex        =   15
            Top             =   2880
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TC Plants"
            Height          =   195
            Left            =   960
            TabIndex        =   14
            Top             =   360
            Width           =   690
         End
      End
   End
   Begin VB.ComboBox cbomnth 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmnurserydashboard.frx":0FA7
      Left            =   3000
      List            =   "frmnurserydashboard.frx":0FCF
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo cboyear 
      Bindings        =   "frmnurserydashboard.frx":100F
      DataField       =   "ItemCode"
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label10 
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   405
   End
   Begin VB.Label Label9 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frmnurserydashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub actiongrid_Click()
Select Case actiongrid.col
Case 1, 2, 3
actiongrid.Editable = flexEDKbdMouse

End Select
End Sub

Private Sub issuegrid_Click()
Select Case issuegrid.col
Case 1, 2
issuegrid.Editable = flexEDKbdMouse

End Select
End Sub

Private Sub txthardplants_Change()
Select Case UCase(txthardplants.Text)
    Case "G"
    txthardplants.BackColor = vbGreen
    Case "Y"
    txthardplants.BackColor = vbYellow
    Case "R"
    txthardplants.BackColor = vbRed
    Case ""
    txthardplants.BackColor = vbWhite

End Select
End Sub

Private Sub txthardplants_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txthealthandsafty_Change()
Select Case UCase(txthealthandsafty.Text)
    Case "G"
    txthealthandsafty.BackColor = vbGreen
    Case "Y"
    txthealthandsafty.BackColor = vbYellow
    Case "R"
    txthealthandsafty.BackColor = vbRed
    Case ""
    txthealthandsafty.BackColor = vbWhite

End Select
End Sub

Private Sub txthealthandsafty_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txthrneeds_Change()
Select Case UCase(txthrneeds.Text)
    Case "G"
    txthrneeds.BackColor = vbGreen
    Case "Y"
    txthrneeds.BackColor = vbYellow
    Case "R"
    txthrneeds.BackColor = vbRed
    Case ""
    txthrneeds.BackColor = vbWhite

End Select
End Sub

Private Sub txthrneeds_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtinventory_Change()
Select Case UCase(txtinventory.Text)
    Case "G"
    txtinventory.BackColor = vbGreen
    Case "Y"
    txtinventory.BackColor = vbYellow
    Case "R"
    txtinventory.BackColor = vbRed
    Case ""
    txtinventory.BackColor = vbWhite

End Select
End Sub

Private Sub txtinventory_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtmaterials_Change()
Select Case UCase(txtmaterials.Text)
    Case "G"
    txtmaterials.BackColor = vbGreen
    Case "Y"
    txtmaterials.BackColor = vbYellow
    Case "R"
    txtmaterials.BackColor = vbRed
    Case ""
    txtmaterials.BackColor = vbWhite

End Select
End Sub

Private Sub txtmaterials_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtnutplants_Change()
Select Case UCase(txtnutplants.Text)
    Case "G"
    txtnutplants.BackColor = vbGreen
    Case "Y"
    txtnutplants.BackColor = vbYellow
    Case "R"
    txtnutplants.BackColor = vbRed
    Case ""
    txtnutplants.BackColor = vbWhite

End Select
End Sub

Private Sub txtnutplants_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub



Private Sub txttcplants_Change()
Select Case UCase(txttcplants.Text)
    Case "G"
    txttcplants.BackColor = vbGreen
    Case "Y"
    txttcplants.BackColor = vbYellow
    Case "R"
    txttcplants.BackColor = vbRed
    Case ""
    txttcplants.BackColor = vbWhite

End Select
End Sub

Private Sub txttcplants_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtutilization_Change()
Select Case UCase(txtutilization.Text)
    Case "G"
    txtutilization.BackColor = vbGreen
    Case "Y"
    txtutilization.BackColor = vbYellow
    Case "R"
    txtutilization.BackColor = vbRed
    Case ""
    txtutilization.BackColor = vbWhite

End Select
End Sub

Private Sub txtutilization_KeyPress(KeyAscii As Integer)
If InStr(1, "GYRgyr", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
