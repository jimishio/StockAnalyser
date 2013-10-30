VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Object = "{EE211B73-72AE-4772-9E9A-A261C54FB94A}#1.0#0"; "lvbuttons.ocx"
Begin VB.Form Portfolio 
   Caption         =   "StockSoft - Portfolio ( Analysis Service )"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   Icon            =   "Portfolio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "General Actions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   12720
      TabIndex        =   14
      Top             =   120
      Width           =   7575
      Begin LVbuttons.LaVolpeButton LaVolpeButton9 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Analyse"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":164A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton5 
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Save && Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   65280
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":1666
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton7 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Find Company Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   65280
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":1682
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton10 
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Clear Portfolio"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":169E
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Colour Code Meaning"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   7920
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin LVbuttons.LaVolpeButton LaVolpeButton6 
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":16BA
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sell"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   390
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3600
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caution"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   840
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1680
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keep"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":16D6
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":16F2
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton3 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":170E
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton4 
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Find"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":172A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton8 
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Plot"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   16776960
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Portfolio.frx":1746
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid GRD 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   20415
      _cx             =   36010
      _cy             =   15901
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12632256
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65280
      ForeColorSel    =   0
      BackColorBkg    =   4210752
      BackColorAlternate=   12632256
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   1
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Data are (in Cr. Rs.)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   2
      Left            =   17760
      TabIndex        =   12
      Top             =   1200
      Width           =   2490
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   20400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   20400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analysis Service - Stock Performance Dashboard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   7515
      TabIndex        =   11
      Top             =   1200
      Width           =   5325
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   4  'Upward Diagonal
      Height          =   495
      Left            =   0
      Top             =   1080
      Width           =   20415
   End
   Begin VB.Image Image1 
      Height          =   10815
      Left            =   0
      Picture         =   "Portfolio.frx":1762
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20415
   End
End
Attribute VB_Name = "Portfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
Portfolio.savegrid
End Sub

Private Sub LaVolpeButton10_Click()
Dim x As String
x = MsgBox("Are You Sure you want to clear your portfolio?", vbYesNo)

If x = vbYes Then
GRD.Rows = 0
End If
End Sub

Private Sub LaVolpeButton7_Click()
HelpFind.Show
End Sub

Private Sub Form_Load()
With GRD
    .Row = 0
    .Col = 0
    .Text = "Company"
    .Col = 1
    .Text = "Face Value"
    .Col = 2
    .Text = "Market Value"
    .Col = 3
    .Text = "Current EPS"
    .Col = 4
    .Text = "Previous EPS"
    .Col = 5
    .Text = "P/E Ratio"
    .Col = 6
    .Text = "P/B Ratio"
    .Col = 7
    .Text = "Growth Rate"
    .Col = 8
    .Text = "PEG"
    .Col = 9
    .Text = "ROCE"
    .Col = 10
    .Text = "Net Asset"
    .Col = 11
    .Text = "Market Cap"
    .Col = 12
    .Text = "Net Profit"
    .Col = 13
    .Text = "Debt Ratio"
    .Col = 14
    .Text = "Net Asset / Market Cap"
    .Col = 15
    .Text = "Last Year Dividend / Share"
    
    



loadgrid

    ' This lines of code resizes the cols irrespective of the savd format
    
    .ColWidth(0) = 2800
    .ColWidth(1) = 1200
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 700
    .ColWidth(9) = 700
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    .ColWidth(14) = 1200
    .ColWidth(15) = 1200

End With

refreshdatA
End Sub

Private Sub LaVolpeButton1_Click()
AddCompany.Show
End Sub

Private Sub LaVolpeButton2_Click()
Dim str As String
str = MsgBox("Are You Sure You Want to Delete this Company from your Portfolio?", vbYesNo, "Confirm Delete")
If str = vbYes Then
GRD.RemoveItem (GRD.Row)
End If
End Sub

Private Sub LaVolpeButton3_Click()


With Portfolio.GRD
    
    .Col = 0
    EditCompany.Text1.Text = .Text
    .Col = 1
     EditCompany.Text2.Text = .Text
    .Col = 2
     EditCompany.Text3.Text = .Text
    .Col = 3
     EditCompany.Text4.Text = .Text
    .Col = 4
     EditCompany.Text5.Text = .Text
    
    .Col = 6
     EditCompany.Text6.Text = .Text
    .Col = 9
     EditCompany.Text7.Text = .Text
    .Col = 10
     EditCompany.Text8.Text = .Text
    .Col = 11
     EditCompany.Text9.Text = .Text
    .Col = 12
     EditCompany.Text10.Text = .Text
    
    .Col = 13
     EditCompany.Text11.Text = .Text
    .Col = 15
     EditCompany.Text12.Text = .Text
    
    
    
End With
EditCompany.Show

End Sub

Private Sub LaVolpeButton4_Click()
FindData.Show

End Sub

Private Sub LaVolpeButton5_Click()
savegrid
End
End Sub

Private Sub LaVolpeButton6_Click()
refreshdatA
End Sub















Public Function refreshdatA()
Dim i As Integer, j As Integer
GRD.Visible = False
For i = 1 To GRD.Rows - 1
    GRD.Row = i
    
    For j = 0 To GRD.Cols - 1
        GRD.Col = j
        GRD.CellBackColor = &HC0C0C0
    Next j
    
Next i

For i = 1 To GRD.Rows - 1
    GRD.Row = i
    
    For j = 5 To GRD.Cols - 2
        GRD.Col = j
        GRD.CellBackColor = vbGreen
    Next j
    
Next i


filldata
GRD.Visible = True


End Function




Public Function savegrid()
On Error Resume Next
Dim str As String
str = App.Path & "\data.atman"
GRD.savegrid str, flexFileCommaText
End Function

Public Function loadgrid()
On Error Resume Next
Dim str As String
str = App.Path & "\data.atman"
GRD.loadgrid str, flexFileCommaText
End Function

Public Function filldata()
On Error Resume Next
Dim i As Integer
Dim cval As Double, ceps As Double, peps As Double, peratio As Double, gr As Double, na As Double, mcap As Double

For i = 1 To GRD.Rows - 1
    GRD.Row = i
    
    GRD.Col = 2
    cval = GRD.Text
    GRD.Col = 3
    ceps = GRD.Text
    GRD.Col = 4
    peps = GRD.Text
    peratio = cval / ceps
    gr = ((ceps - peps) / peps) * 100
    
    GRD.Col = 5
    GRD.ColFormat(5) = "#.###"
    GRD.Text = peratio
    
    GRD.Col = 7
    GRD.ColFormat(7) = "#.###"
    GRD.Text = gr
    
    GRD.Col = 8
    GRD.ColFormat(8) = "#.###"
    GRD.Text = peratio / gr
    
    GRD.Col = 10
    GRD.ColFormat(10) = "#.##"
    na = GRD.Text
    
    GRD.Col = 11
    GRD.ColFormat(11) = "#.##"
    mcap = GRD.Text
    
    GRD.Col = 14
    GRD.ColFormat(14) = "#.###"
    GRD.Text = (na / mcap) * 100
    
    
    
Next

colordata
End Function

Public Function colordata()
Dim i As Integer
Dim j As Integer

Dim peratio As Double, gr As Double, peg As Double, roce As Double, mcap As Double, dratio As Double, pmarg As Double

For i = 1 To GRD.Rows - 1
GRD.Row = i


'code for orange

GRD.Col = 5
If CDbl(GRD.Text) > 15 Then
colorrow i, 5, "orange"
End If

GRD.Col = 6
If CDbl(GRD.Text) > 1 Then
colorrow i, 6, "orange"
End If

GRD.Col = 7
If CDbl(GRD.Text) < 30 Then
colorrow i, 7, "orange"
End If

GRD.Col = 8
If CDbl(GRD.Text) = 1 Then
colorrow i, 8, "orange"
End If
If CDbl(GRD.Text) < 0 Then
colorrow i, 8, "red"
End If


GRD.Col = 9
If CDbl(GRD.Text) < 15 Then
colorrow i, 9, "orange"
End If

GRD.Col = 11
If CDbl(GRD.Text) < 100 Then
colorrow i, 11, "orange"
End If

GRD.Col = 12
If CDbl(GRD.Text) < 12 Then
colorrow i, 12, "orange"
End If

GRD.Col = 13
If CDbl(GRD.Text) > 1 Then
colorrow i, 13, "orange"
End If

GRD.Col = 14
If CDbl(GRD.Text) < 20 Then
colorrow i, 14, "orange"
End If






GRD.Col = 5
If CDbl(GRD.Text) > 20 Then
colorrow i, 5, "red"
End If

GRD.Col = 6
If CDbl(GRD.Text) > 2 Then
colorrow i, 6, "red"
End If

GRD.Col = 7
If CDbl(GRD.Text) < 20 Then
colorrow i, 7, "red"
End If

GRD.Col = 8
If CDbl(GRD.Text) > 1 Then
colorrow i, 8, "red"
End If
If CDbl(GRD.Text) < 0 Then
colorrow i, 8, "red"
End If


GRD.Col = 9
If CDbl(GRD.Text) < 10 Then
colorrow i, 9, "red"
End If

GRD.Col = 11
If CDbl(GRD.Text) < 75 Then
colorrow i, 11, "red"
End If

GRD.Col = 12
If CDbl(GRD.Text) < 10 Then
colorrow i, 12, "red"
End If

GRD.Col = 13
If CDbl(GRD.Text) > 2 Then
colorrow i, 13, "red"
End If

GRD.Col = 14
If CDbl(GRD.Text) < 5 Then
colorrow i, 14, "red"
End If





Next i
End Function


Public Function colorrow(i As Integer, k As Integer, color As String)
Dim j As Integer
If color = "red" Then
GRD.Row = i
GRD.Col = k
GRD.CellBackColor = vbRed
    
End If

If color = "orange" Then
GRD.Row = i
GRD.Col = k
GRD.CellBackColor = vbYellow
    
End If

End Function
