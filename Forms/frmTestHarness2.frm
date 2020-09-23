VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTestHarness2 
   Caption         =   "ucPickBox - TestHarness"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Welcome "
      TabPicture(0)   =   "frmTestHarness2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblWelcome(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblWelcome(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLink"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblAuthorMessage"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Color "
      TabPicture(1)   =   "frmTestHarness2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTextEntry"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line1(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblMessageOn"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblProperties"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblChkBoxOptions"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblColorPick"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "pbColorPick"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtTextEntry"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmbPrintStatusMsg"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lstProperties"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkUseDialog"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Font "
      TabPicture(2)   =   "frmTestHarness2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Open"
      TabPicture(3)   =   "frmTestHarness2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Save "
      TabPicture(4)   =   "frmTestHarness2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Print "
      TabPicture(5)   =   "frmTestHarness2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.CheckBox chkUseDialog 
         Caption         =   "UseDialogColor "
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ListBox lstProperties 
         Height          =   1620
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbPrintStatusMsg 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtTextEntry 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin prjTestHarness.ucPickBox pbColorPick 
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Tag             =   "x"
         Top             =   1800
         Width           =   2535
         _extentx        =   4471
         _extenty        =   556
         color           =   0
         dialogmsg       =   ""
         dialogmsg       =   ""
         dialogmsg       =   ""
         dialogmsg       =   ""
         dialogmsg       =   ""
         printer         =   ""
         printstatusmsg  =   ""
         printstatusmsg  =   ""
         Object.tooltiptext     =   ""
         Object.tooltiptext     =   ""
         Object.tooltiptext     =   ""
         Object.tooltiptext     =   ""
         Object.tooltiptext     =   ""
      End
      Begin VB.Label lblColorPick 
         Caption         =   "BackColor:"
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Tag             =   "x"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblChkBoxOptions 
         Caption         =   "UseDialog Color/Text:"
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "ucPickBox Properties"
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblProperties 
         Caption         =   "Property Type:"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblMessageOn 
         Caption         =   "Message On:"
         Height          =   315
         Left            =   4080
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   $"frmTestHarness2.frx":00A8
         Height          =   975
         Left            =   -74640
         TabIndex        =   5
         Top             =   3045
         Width           =   4455
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to ucPickBox!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTestHarness2.frx":0144
         ForeColor       =   &H80000008&
         Height          =   2055
         Index           =   1
         Left            =   -74640
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   -72120
         Picture         =   "frmTestHarness2.frx":0288
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2040
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "click here"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71640
         MouseIcon       =   "frmTestHarness2.frx":0612
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblAuthorMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To provide feedback on this control, please                 ...."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   4200
         Width           =   5415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   840
         X2              =   4920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   840
         X2              =   4920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblTextEntry 
         Caption         =   "Dialog Message:"
         Height          =   315
         Left            =   2520
         TabIndex        =   16
         Top             =   2280
         Width           =   2535
      End
   End
   Begin prjTestHarness.ucPickBox ucPickBox1 
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   4200
      Width           =   2535
      _extentx        =   4471
      _extenty        =   556
   End
End
Attribute VB_Name = "frmTestHarness2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'   Link URL address which searches for our control submission on PCS
Const sLink As String = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&?lngWId=1&grpCategories=&txtMaxNumberOfEntriesPerPage=10&optSort=Alphabetical&chkThoroughSearch=&blnTopCode=False&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=1&intLastRecordOnPage=10&intMaxNumberOfEntriesPerPage=10&intLastRecordInRecordset=499&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=ucPickBox"


Option Explicit

Private Sub lblLink_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hwnd, "open", sLink, vbNull, vbNull, 1)
    End With
End Sub

