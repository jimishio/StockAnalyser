VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FindData 
   BackColor       =   &H00000000&
   Caption         =   "Find Company Data"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   13155
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      ExtentX         =   22675
      ExtentY         =   13361
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "FindData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Silent = True
WebBrowser1.Navigate "http://www.moneycontrol.com/stocks/cptmarket/charting_tech.php?"
End Sub
