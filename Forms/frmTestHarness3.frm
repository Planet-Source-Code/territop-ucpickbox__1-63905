VERSION 5.00
Begin VB.Form frmTestHarness3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ucPickBox v1.8.134 - TestHarness"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "frmTestHarness3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Step 3 - Select Item using ucPickBox"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   35
      Top             =   3600
      Width           =   5535
      Begin prjTestHarness.ucPickBox ucPickBox1 
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   555
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Click the Button to load the Common Dialog, or Enter the value directly onto the control."
         Height          =   615
         Left            =   2880
         TabIndex        =   38
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "ucPickBox Control:"
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   320
         Width           =   2535
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Step 4: Review ucPickBox Results"
      ForeColor       =   &H00FF0000&
      Height          =   1400
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   5535
      Begin VB.PictureBox picFolder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   5385
         TabIndex        =   45
         Top             =   320
         Width           =   5380
         Begin VB.TextBox txtFolderPath 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label lblOpenLabel 
            Caption         =   "Folder:"
            Height          =   315
            Index           =   2
            Left            =   0
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picOpenSave 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   5385
         TabIndex        =   15
         Top             =   320
         Width           =   5380
         Begin VB.TextBox txtOpenPath 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   480
            Width           =   4455
         End
         Begin VB.TextBox txtSaveFilename 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   0
            Width           =   2535
         End
         Begin VB.TextBox txtSavePath 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   480
            Width           =   4455
         End
         Begin VB.TextBox txtFileExists 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtOpenFilename 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label lblSaveLabel 
            Caption         =   "Save File:"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblSaveLabel 
            Caption         =   "Save Path:"
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblOpenLabel 
            Caption         =   "Open File:"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblOpenLabel 
            Caption         =   "Open Path:"
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblFileExistsLabel 
            Caption         =   "File Exists:"
            Height          =   315
            Left            =   3480
            TabIndex        =   21
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picFont 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   5385
         TabIndex        =   31
         Top             =   320
         Width           =   5380
         Begin VB.ListBox lstFontSettings 
            BackColor       =   &H8000000F&
            Height          =   645
            Left            =   0
            TabIndex        =   32
            Top             =   240
            Width           =   5295
         End
         Begin VB.Label lblFontLabel 
            Caption         =   "Font:"
            Height          =   315
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   5385
         TabIndex        =   26
         Top             =   320
         Width           =   5380
         Begin VB.TextBox txtColorValue 
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   2760
            TabIndex        =   27
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   0
            TabIndex        =   29
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblColorLabel 
            Caption         =   "Color Vaue (Hex):"
            Height          =   315
            Index           =   1
            Left            =   2760
            TabIndex        =   28
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblColorLabel 
            Caption         =   "Color:"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Label lblPrintStatus 
         Caption         =   "Print Status:"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   320
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step 2 - Set ucPickBox Properties"
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5535
      Begin VB.CheckBox chkAppear3D 
         Caption         =   "3D"
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   560
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkMulitSel 
         Caption         =   "MultiSelect"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   560
         Width           =   1095
      End
      Begin VB.TextBox txtTextEntry 
         Height          =   315
         Left            =   2880
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkUseDialog 
         Caption         =   "Color "
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   560
         Width           =   1575
      End
      Begin VB.ListBox lstProperties 
         Height          =   1425
         Left            =   120
         TabIndex        =   6
         Top             =   560
         Width           =   2655
      End
      Begin VB.ComboBox cmbPrintStatusMsg 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   975
      End
      Begin prjTestHarness.ucPickBox pbColorPick 
         Height          =   315
         Left            =   2880
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Color           =   0
         Printer         =   ""
         PrintStatusMsg  =   ""
         PrintStatusMsg  =   ""
      End
      Begin VB.Label lblChkBoxOptions2 
         Caption         =   "Appearance:"
         Height          =   315
         Left            =   4320
         TabIndex        =   42
         Top             =   315
         Width           =   975
      End
      Begin VB.Label lblColorPick 
         Caption         =   "BackColor:"
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblChkBoxOptions 
         Caption         =   "UseDialog:"
         Height          =   315
         Left            =   2880
         TabIndex        =   11
         Top             =   315
         Width           =   975
      End
      Begin VB.Label lblProperties 
         Caption         =   "Property Type:"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   320
         Width           =   2535
      End
      Begin VB.Label lblMessageOn 
         Caption         =   "Message On:"
         Height          =   315
         Left            =   4440
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblTextEntry 
         Caption         =   "Dialog Message:"
         Height          =   315
         Left            =   2880
         TabIndex        =   13
         Top             =   1440
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Step 1 - Select Dialog Type"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "To begin, Please select the dialog type to use from the Droplist to the left, then proceed to Step 2."
         Height          =   615
         Left            =   2880
         TabIndex        =   39
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Dialog Type:"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label lblLink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "click here"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4155
      MouseIcon       =   "frmTestHarness3.frx":038A
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblAuthorMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To provide constructive feedback on this control, please                 ...."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   6240
      Width           =   5415
   End
End
Attribute VB_Name = "frmTestHarness3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'       ucPickBox - Enhanced File/Color/Font/Printer Picker Control
'
'   Product Name:
'       ucPickBox.ctl
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Common Dialog API Calls - Paul Mather)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=3592&lngWId=1
'       (TrimPathLen Function - Wastingtape)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=23456&lngWId=1
'       (FileExists - Eric Russell)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=829&lngWId=1
'       (ComboBox Open/Visible - Francesco Balena)
'           URL: http://www.devx.com/vb2themax/Tip/18336
'       (Max Raskin - Flat Button)
'           http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6517&lngWId=1
'
'   Legal Copyright & Trademarks:
'       Copyright © 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2006, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       sold for use as a license in accordance with the terms of the License
'       Agreement in the accompanying the documentation.
'
'       Many thanks to my friend Paul Turcksin for his careful review, suggestions,
'       and support of this UserControl and TestHarness prior to public release. In
'       addtion, I wish to thank the numerous open source authors who provide code
'       and inspiration to make such work possible.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       05Nov05 - Initial TestHarness and UserControl finished
'       06Nov05 - Cleaned up bugs in the ShowSave and ShowOpen routines.
'               - Consolidated calls for the Show Open/Save subs to make
'                 param and error handling cleaner.
'               - Added addtional API params to the ShowFont routine.
'               - Updated the ToolBox Image to a more professional image.
'               - Added addtional error handling to the TestHarness...
'       19Nov05 - Added Additional Author Credits to the Header
'               - Added UseDialogColor, UseDialogText, ForeColor, and
'                 BackColor properties to the Control and required code to
'                 allow these routines to work...
'               - Added PrintStatusMsg property to allow the user to specify
'                 what the message should say when the printer returns a value.
'               - Added PrintStaus property to provide the user feedback about
'                 if the Printer dialog "Ok"(1) or "Cancel"(0) button was pressed.
'               - Fixed bug in ShowSave routine which inconssistently computes the
'                 nFileOffset values for a file. We simply set this to "0" and then
'                 extract the values from outside of this of this routine.
'               - Changes Color from Long to OLE_COLOR property to allow for
'                 vb stanard palette.
'               - Added TranslateColor sub to wrap the OleTranslateColor method
'                 for mapping of colors to the current RGB palette.
'       20Nov05 - Added Color RollBack if the value entered is invalid.
'       04Dec05 - Changed the TestHarness layout to make it easier to follow the
'                 flow of the controls and how to use it....
'       06Dec05 - Added MultiFile selection for the ShowOpen routine and fixed several
'                 bugs with the single vs mutiple file selections.
'               - Added a ComboBox to serve at the conatiner and windowing mechanism for
'                 the list and its events....this is a hack, pure and simple. This
'                 approach was selected as it allowes a floating window and list functionality
'                 without the need for building this via API. The combobox is hidden
'                 behind the textbox at runtime and has Visiable = False. Since we
'                 call the droplist window via SendMessage this allows us to have a
'                 floating window like the ComboBox, but none of the overhead to manage ;-D
'               - Add the ability to programmatically Open the MultiFile ComboBox
'                 and check the state of the Droplist.
'               - Added cmdDrop button to simulate the drop button of the ComboBox. The
'                 key feature here being that the button is to the left of the ellipes
'                 button and is resizable with the dialog, unlike the VB ComboBox.
'       13Dec05 - Fixed minor TestHareness bug which displayed the wrong properties when
'                 selecting the lstProperties index.
'       14Dec05 - Fixed single/multiple file open bug in the ShowOpen routine which caused the
'                 the sub to enter into the wrong conditional section when a single file
'                 was selected and the MultiSelect = False.
'               - Fixed PropertyChanged calls for DialogMsg and ToolTipTexts which now supports
'                 individual item settings.
'       15Dec05 - More optimization on the ShowSave and ShowFont routines. These routines now
'                 handle missing extensions and provide a mechanism to enter them. In addtion,
'                 the FontColor property has been added to allow direct color picking of the
'                 font ForeColor, which is not appart of the StdFont structure.
'       16Dec05 - Added Appearance Property and associated API and VB routines to allow for true
'                 3D or Flat appearances of the textbox and buttons.
'       18Dec05 - Fixed Minor bugs in the ShowFont dialog routines which did not preserve the
'                 previous selections by the user. The new addtions resolve all but one known
'                 bug. At the current time, the iPointSize of the FontDialog type structure is
'                 not correctly set via code and the dialog does not respond the changes in this
'                 parameter despite accounting for the size and weight of the font. Verified the
'                 ShowFont code against www.allapi.net example and neither resulted in the pointsize
'                 being selected. For more details see http://mentalis.org/apilist/CHOOSEFONT.shtml
'       25Dec05 - Added Events: DropClick, KeyDown, KeyPress, KeyUp, MouseDown, MouseMove, MouseUp.
'               - Added GetCursorPosition function to allow reporting of the Cursor position via
'                 GetCursorPosition and ScreenToClient API's regardless of which part of the control
'                 the cursor is over. This effectively bypasses the native Event Handlers for each
'                 control, and provides a uniform reporting of the cursor position on the control surface.
'               - Added additional documentation at the Method and Property levels to provide added
'                 clarity of what the functionality is...
'       26Dec05 - Added Filter Property and associated routines to the ShowOpen, ShowSave routines,
'                 see Filter Let property for correct format of the filter string....
'               - Added ProcessFilter to replace string Pipes (|) with vbNullChar and fix the
'                 final size of the passed string to the dialogs.
'               - Added error handling for none initialized Filters to read All Files (*.*)
'       27Dec05 - Added Color, Font, File, and PrinterFlags as Public Enums along with properties
'                 to allow the developer set the styles more easily.
'               - Added SHOWCOLOR_DEFAULT, SHOWFONT_DEFAULT, SHOWOPEN_DEFAULT, SHOWSAVE_DEFAULT,
'                 and SHOWPRINTER_DEFAULT custom Non-Win32 flags to allow for rapid dialog setting
'                 which encompass the most common flags used with this control.
'               - Updated the TestHarness in the UpdatePropertiesDialog to reflect these changes.
'       28Dec05 - Added UseAutoForeColor and associated routines to allow the developer to choose
'                 if the ForeColor is to be selected automatically. The value for the new ForeColor
'                 is based on the XOr of the BackColor and should always produce high contrast text
'                 in the dialog regardless of the color selected.
'       03Jan06 - Added BrowseForFolder functionality and associated routines to round out the collection
'                 based on the request from Richard Mewett.
'       07Mar06 - Added Let Property for Path to pass data to txtResult and m_Path parameter. The displayed
'                 Path is trimmed using the TrimPathLen routine.
'               - Fixed bug which causes the txtResult to display the incorrect message when ucFolder was the
'                 dialog type.
'       16Mar06 - Add Paul Caton's SelfSubclass Thunk code to allow for BrowseForFolder CallBack without the
'                 need for an external bas module. The long point (address) of the z_SubclassProc is held in
'                 in the sc_aSubData(0).nAddrSub provided this is the only item we are subclassing....if we are
'                 subclassing multiple items (i.e. Usercontrol, Parent) then the address for each is stored in
'                 order in the sc_aSubData(n).nAddrSub, where n = 0, 1....n
'
'   Force Declarations
Option Explicit
'
'   Build Date & Time: 3/16/2006 10:31:08 AM
Const Major As Long = 1
Const Minor As Long = 8
Const Revision As Long = 132
Const BuildDateTime As String = "3/16/2006 10:31:08 AM"
'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
Private bLoading    As Boolean      'TestHarness Only -  Used to prevent recursion on Form_Load Events
Private m_ListIndex As Long         'TestHarness Only -  Used to keep track of ListIndex
'
'   Link URL address which searches for our control submission on PCS
Private Const sLink As String = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&?lngWId=1&grpCategories=&txtMaxNumberOfEntriesPerPage=10&optSort=Alphabetical&chkThoroughSearch=&blnTopCode=False&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=1&intLastRecordOnPage=10&intMaxNumberOfEntriesPerPage=10&intLastRecordInRecordset=499&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=ucPickBox"

Private Sub AutoSelText(TxtBox As TextBox)
    With TxtBox
        '   Select the text
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
End Sub

Private Sub chkAppear3D_Click()
    With Me
        .ucPickBox1.Appearance = Abs(.chkAppear3D.Value)
    End With
End Sub

Private Sub chkMulitSel_Click()
    With Me
        .ucPickBox1.MultiSelect = Abs(.chkMulitSel.Value)
    End With
End Sub

Private Sub chkUseDialog_Click()
    With Me
        If Not bLoading Then
            '   Check for UseDialogColor Option to be Selected
            Select Case m_ListIndex
                Case 8
                    '   This is for UseDialogText
                    .ucPickBox1.UseAutoForeColor = Abs(.chkUseDialog.Value)
                Case 9
                    '   This is for UseDialogColor
                    .ucPickBox1.UseDialogColor = Abs(.chkUseDialog.Value)
                    If .chkUseDialog.Value = vbChecked Then
                        '   Update the Color from ucPickBox
                        .pbColorPick.Color = .ucPickBox1.Color
                    Else
                        '   Update the BackColor
                        .lblColorPick.Caption = "BackColor:"
                        .pbColorPick.Color = .ucPickBox1.BackColor
                    End If
                Case 10
                    '   This is for UseDialogText
                    .ucPickBox1.UseDialogText = Abs(.chkUseDialog.Value)
            End Select
        End If
    End With
End Sub

Private Sub cmbPrintStatusMsg_Click()
    With Me
        If Not bLoading Then
            '   Call our custom updating routine
            Call UpdatePropertyDialogs(.cmbType.ListIndex)
        End If
    End With
End Sub

Private Sub cmbType_Click()
    With Me
        If Not bLoading Then
            '   Reset the Controls
            .pbColorPick.Reset
            '   If we are not using ucColor then Reset the BackColor via UseDialogColor
            If (.ucPickBox1.UseDialogColor) And (.cmbType.ListIndex <> 0) Then
                .ucPickBox1.UseDialogColor = False
            End If
            '   Pass the Type to the Control
            .ucPickBox1.DialogType = .cmbType.ListIndex
            '   Update the Dialogs...
            Call UpdatePropertyDialogs(.cmbType.ListIndex)
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim lAppear As ucAppearanceConstants
    With Me
        '   Prevent recursion by flagging the start...
        bLoading = True
        lAppear = .ucPickBox1.Appearance
        '   Set the Open / Save Filters
        With .ucPickBox1
            '   This is an example of how to format the filter strings
            '   for the ucPickBox version of the CommonDialog Controls
            .Filters = "Supported Files|*.bmp;*.doc;*.jpg;*.rtf;*.txt;*.tif|Bitmap Files (*.bmp)|*.bmp|Mircosoft Word Files (*.doc)|*.doc|JPEG Files (*.jpg)|*.jpg|Rich Text Format Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
        End With
        With .cmbType
            '   Give them some choices
            .AddItem "ucColor"              '0
            .AddItem "ucFolder"             '1
            .AddItem "ucFont"               '1 2
            .AddItem "ucOpen"               '2 3
            .AddItem "ucSave"               '3 4
            .AddItem "ucPrint"              '4 5
            '   This is our default
            .ListIndex = 0
        End With
        .ucPickBox1.Appearance = lAppear
        With .lstProperties
            .Clear
            .AddItem "Appearance"           '0
            .AddItem "BackColor"            '1
            .AddItem "DialogMsg"            '2
            .AddItem "ForeColor"            '3
            .AddItem "MultiSelect"          '4
            .AddItem "Path"                 '-  5
            .AddItem "PrintStatusMsg"       '5  6
            .AddItem "ToolTipTexts"         '6  7
            .AddItem "UseAutoForeColor"     '7  8
            .AddItem "UseDialogColor"       '8  9
            .AddItem "UseDialogText"        '9  10
            .AddItem "Version"              '10 11
            '   This is our default
            .ListIndex = 0
        End With
        With .cmbPrintStatusMsg
            .AddItem "Failure"              '0
            .AddItem "Success"              '1
            '   This is our default
            .ListIndex = 0
            .Visible = False
        End With
        .lblMessageOn.Visible = False
        '   Full Width
        .txtTextEntry.Width = 2535
        '   Init the Condtions on the ucColor Dialogs
        Call UpdatePropertyDialogs(ucColor)
        '   Done loading, so turn off the flag
        bLoading = False
    End With
End Sub

Private Sub lblLink_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hWnd, "open", sLink, vbNull, vbNull, 1)
    End With
End Sub

Private Sub lstProperties_Click()
    With Me
        If Not bLoading Then
            '   Call our custom Updating routine....
            Call UpdatePropertyDialogs(.cmbType.ListIndex)
        End If
    End With
End Sub

Private Sub pbColorPick_Click()
    With Me
        On Error Resume Next
        Select Case m_ListIndex
            Case 1
                '   Set the New Selected Color
                .ucPickBox1.BackColor = .pbColorPick.Color
            Case 3
                '   Set the New Selected Color
                .ucPickBox1.ForeColor = .pbColorPick.Color
        End Select
    End With
End Sub

Private Sub pbColorPick_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If KeyCode = vbKeyReturn Then
            '   Forward the call to maximize code reuse...
            pbColorPick_Click
        End If
    End With
End Sub

Private Sub txtTextEntry_GotFocus()
    With Me
        '   Auto select all the text
        Call AutoSelText(txtTextEntry)
    End With
End Sub

Private Sub txtTextEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If KeyCode = vbKeyReturn Then
            '   Forward the call to maximize code reuse...
            Call txtTextEntry_LostFocus
        End If
    End With
End Sub

Private Sub txtTextEntry_LostFocus()
    With Me
        Select Case m_ListIndex '(2 = DialogMsg; 5=Path; 6 = PrintStatusMsg; 7 = ToolTipTexts)
            Case 2
                '   Set the PickBox Dialog Message
                .ucPickBox1.DialogMsg(.cmbType.ListIndex) = .txtTextEntry.Text
            Case 5
                '   Set the PickBox Initial Path for BrowseForFolder
                .ucPickBox1.Path = .txtTextEntry.Text
            Case 6
                '   Set the PickBox Print Status Message
                .ucPickBox1.PrintStatusMsg(.cmbPrintStatusMsg.ListIndex) = .txtTextEntry.Text
            Case 7
                '   Set the PickBox ToolTipText Message
                .ucPickBox1.ToolTipTexts(.cmbType.ListIndex) = .txtTextEntry.Text
        End Select
    End With
End Sub

Private Sub ucPickBox1_Click()
    Dim NewFont As StdFont
    Dim nFile As Long
    
    '   This entrire section of the TestHarness is used to update the
    '   "Step 4 - Review xxxxx Results" Frame for the various dialogs....
    '
    
    With Me
        Select Case .ucPickBox1.DialogType
            Case [ucColor]
                '   Update the new Color of the Label
                .lblColor.BackColor = .ucPickBox1.Color
                '   Set the hex color value to the control
                .txtColorValue.Text = .ucPickBox1.LongToHexColor(.ucPickBox1.Color)
            Case [ucFolder]
                '   Set the new path in the results TextBox
                .txtFolderPath.Text = .ucPickBox1.TrimPathByLen(.ucPickBox1.Path, .txtOpenPath.Width)
            Case [ucFont]
                '   Set the new Font in the Listbox
                Set NewFont = .ucPickBox1.Font
                If Not NewFont Is Nothing Then
                    With .lstFontSettings
                        '   Clear the list to start
                        .Clear
                        '   Now add all of the Font properties...
                        .AddItem "Name " & vbTab & "(" & NewFont.Name & ")"
                        .AddItem "Bold " & vbTab & "(" & NewFont.Bold & ")"
                        .AddItem "Italic " & vbTab & "(" & NewFont.Italic & ")"
                        .AddItem "Size " & vbTab & "(" & Round(NewFont.Size, 0) & ")"
                        .AddItem "Strikethrough " & vbTab & "(" & NewFont.Strikethrough & ")"
                        .AddItem "Underline " & vbTab & "(" & NewFont.Underline & ")"
                        .AddItem "Weight " & vbTab & "(" & NewFont.Weight & ")"
                        '   Set the font
                        Set .Font = ucPickBox1.Font
                        '   Set the font color
                        .ForeColor = ucPickBox1.FontColor
                        '   Set our height
                        .Height = 700
                    End With
                Else
                    MsgBox "No New Font Selected....", vbExclamation + vbOKOnly, "ucPickBox"
                End If
            Case [ucOpen]
                If ucPickBox1.FileCount > 0 Then
                    '   Pass the Open Filename data to a text box for display
                    .txtOpenFilename.Text = .ucPickBox1.ExtractFilename(.ucPickBox1.Filename)
                    .txtOpenPath.Text = .ucPickBox1.TrimPathByLen(.ucPickBox1.ExtractPath(.ucPickBox1.Filename), .txtOpenPath.Width)
                    '   Does this file exist? Should be, this is an Open
                    .txtFileExists.Text = .ucPickBox1.FileExists(.ucPickBox1.Filename)
                End If
            Case [ucSave]
                If ucPickBox1.FileCount > 0 Then
                    '   Pass the Save Filename data to a text box for display
                    .txtSaveFilename.Text = .ucPickBox1.ExtractFilename(.ucPickBox1.Filename)
                    .txtSavePath = .ucPickBox1.TrimPathByLen(.ucPickBox1.ExtractPath(.ucPickBox1.Filename), .txtOpenPath.Width)
                    '   Does this file exist? Might not, if we just changed the name...
                    .txtFileExists.Text = .ucPickBox1.FileExists(.ucPickBox1.Filename)
                    '   Create a File with the new name
                    nFile = FreeFile
                    '   Append the File
                    Open .ucPickBox1.Filename For Append As #nFile
                    Print #nFile, "Appended File @ " & Now()
                    '   Close it...
                    Close #nFile
                End If
            Case [ucPrint]
                '   Print the form image...
                .lblPrintStatus.Visible = True
                .lblPrintStatus.Caption = "Print Status: " & .ucPickBox1.PrintStatusMsg(Abs(.ucPickBox1.PrintStatus))
                If .ucPickBox1.PrintStatus Then
                    '   Just to show it works
                    If MsgBox("Do You Really Want to Print a Copy of the Form?", vbQuestion + vbYesNo) = vbYes Then
                        .PrintForm
                    Else
                        '   Indicate the job was aborted...
                        .ucPickBox1.DialogMsg(.cmbType.ListIndex) = "Print Job Aborted!"
                        .lblPrintStatus.Caption = "Print Status: Print Job Aborted!"
                        MsgBox "Print Job Aborted!", vbInformation, "ucPickBox"
                    End If
                End If
        End Select
    End With
End Sub

Private Sub ucPickBox1_ColorChanged(NewColor As Long)
    '   Fire the Event
    Debug.Print "ColorChanged: " & NewColor
End Sub

Private Sub ucPickBox1_DblClick()
    '   Fire the Event
    Debug.Print "DblClick: ucPickBox"
End Sub

Private Sub ucPickBox1_DropClick()
    '   Fire the Event
    Debug.Print "DropClick: ucPickBox"
End Sub

Private Sub ucPickBox1_FontChanged(FontName As String)
    '   Fire the Event
    Debug.Print "FontChanged: " & FontName
End Sub

Private Sub ucPickBox1_GotFocus()
    '   Fire the Event
    Debug.Print "GotFocus: ucPickBox"
End Sub

Private Sub ucPickBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '   ToDo...add additional validation code to validate the data.
        Call ucPickBox1_Click
    End If
End Sub

Private Sub UpdatePropertyDialogs(lDialogType As ucDialogConstant)
    Dim sColor As String
    
    With Me
        '   Update the Controls in the Results Area....
        '
        '   Turn off all dialogs and labels....
        .lblPrintStatus.Caption = "Print Status: "
        .picColor.Visible = False
        .picFont.Visible = False
        .picOpenSave.Visible = False
        .picFolder.Visible = False
        .lblOpenLabel(0).Visible = False
        .lblOpenLabel(1).Visible = False
        .lblOpenLabel(2).Visible = False
        .txtOpenFilename.Visible = False
        .txtOpenPath.Visible = False
        .lblSaveLabel(0).Visible = False
        .lblSaveLabel(1).Visible = False
        .txtSaveFilename.Visible = False
        .txtSavePath.Visible = False
        .txtFolderPath.Visible = False
        .lblPrintStatus.Visible = False
        '   Now select only the ones needed based on the DialogType in
        '   the Results section...
        Select Case lDialogType
            Case [ucColor]
                .Frame4.Caption = "Step 4: Review - Color Results"
                .picColor.Visible = True
                .ucPickBox1.ColorFlags = ShowColor_Default
            Case [ucFolder]
                .Frame4.Caption = "Step 4: Review - Folder Results"
                .lblOpenLabel(2).Visible = True
                .txtFolderPath.Visible = True
                .picFolder.Visible = True
                .ucPickBox1.FolderFlags = ShowFolder_Default
            Case [ucFont]
                .Frame4.Caption = "Step 4: Review - Font Results"
                .lblFontLabel.Visible = True
                .lstFontSettings.Visible = True
                .picFont.Visible = True
                .ucPickBox1.FontFlags = ShowFont_Default
            Case [ucOpen]
                .Frame4.Caption = "Step 4: Review - Open Results"
                .lblOpenLabel(0).Visible = True
                .lblOpenLabel(1).Visible = True
                .txtOpenFilename.Visible = True
                .txtOpenPath.Visible = True
                .picOpenSave.Visible = True
                .ucPickBox1.FileFlags = ShowOpen_Default
            Case [ucSave]
                .Frame4.Caption = "Step 4: Review - Save Results"
                .lblSaveLabel(0).Visible = True
                .lblSaveLabel(1).Visible = True
                .txtSaveFilename.Visible = True
                .txtSavePath.Visible = True
                .picOpenSave.Visible = True
                .ucPickBox1.FileFlags = ShowSave_Default
            Case [ucPrint]
                .Frame4.Caption = "Step 4: Review - Print Results"
                .lblPrintStatus.Visible = True
                .ucPickBox1.PrinterFlags = ShowPrinter_Default
        End Select
        '   Update the Controls in the Properties Area....
        '
        '   Start out with all disabled
        .lblChkBoxOptions.Enabled = False
        .chkUseDialog.Enabled = False
        .lblColorPick.Enabled = False
        .pbColorPick.BackColor = vbButtonFace
        .pbColorPick.Enabled = False
        .lblTextEntry.Enabled = False
        .txtTextEntry.BackColor = vbButtonFace
        .txtTextEntry.Enabled = False
        .lblMessageOn.Enabled = False
        .cmbPrintStatusMsg.Enabled = False
        .chkUseDialog.Enabled = False
        .chkMulitSel.Enabled = False
        .lblChkBoxOptions2.Enabled = False
        .chkAppear3D.Enabled = False
        '   Store the value for later...
        m_ListIndex = .lstProperties.ListIndex
        '   Reset the Dialog Type
        .ucPickBox1.DialogType = lDialogType
        Select Case m_ListIndex
            Case 0              '(0 = Appearance)
                .chkMulitSel.Visible = False
                .lblChkBoxOptions2.Enabled = True
                .lblChkBoxOptions2.Caption = "Appearance:"
                .chkAppear3D.Enabled = True
                .chkAppear3D.Visible = True
                '   Set the Draw State of the Control
                .ucPickBox1.Appearance = Abs(.chkAppear3D.Value)
            Case 1, 3           '(1 = BackColor; 3 = ForeColor)
                '   Now check to see if this is BackColor or ForeColor
                If m_ListIndex = 1 Then
                    '   This is for BackColor....so indicate it
                    .lblColorPick.Caption = "BackColor"
                    '   Get the value
                    If .chkUseDialog.Value = vbChecked Then
                        '   Since the Color is the BackColor
                        .pbColorPick.DialogMsg(0) = .ucPickBox1.LongToHexColor(.ucPickBox1.Color)
                    Else
                        '   Just the BackColor
                        .pbColorPick.DialogMsg(0) = .ucPickBox1.LongToHexColor(.ucPickBox1.BackColor)
                    End If
                Else
                    '   This is for ForeColor....so indicate it
                    .lblColorPick.Caption = "ForeColor"
                    '   Get the value
                    .pbColorPick.DialogMsg(0) = .ucPickBox1.LongToHexColor(.ucPickBox1.ForeColor)
                End If
                '   Enable the Label and ucPickBox
                .lblColorPick.Enabled = True
                .pbColorPick.Enabled = True
                '   Get the backcolor to white
                .pbColorPick.BackColor = &HFFFFFF
            Case 4             '(4 = MultiSelect)
                .lblChkBoxOptions2.Caption = "Select Mode:"
                .chkAppear3D.Visible = False
                .chkMulitSel.Visible = True
                Select Case lDialogType
                    Case [ucOpen]
                        .lblChkBoxOptions2.Enabled = True
                        .chkMulitSel.Enabled = True
                End Select
            Case 2, 5, 6, 7       '(2 = DialogMsg; 5=Path; 6 = PrintStatusMsg; 7 = ToolTipTexts)
                .lblMessageOn.Visible = False
                '   Set the size of the TextBox
                .txtTextEntry.Width = 2535
                Select Case m_ListIndex
                    Case 2
                        '   Set the Label Caption
                        .lblTextEntry.Caption = "Dialog Message"
                        '   Get the Current Values from the ucPickBox
                        .txtTextEntry.Text = .ucPickBox1.DialogMsg(.cmbType.ListIndex)
                        '   Enalbe the Label & Textbox
                        .lblTextEntry.Enabled = True
                        .txtTextEntry.Enabled = True
                        '   Set the BackColor to White
                        .txtTextEntry.BackColor = &HFFFFFF
                    Case 5
                        If .cmbType.ListIndex = 1 Then
                            '   Set the Label Caption
                            .lblTextEntry.Caption = "Starting Path:"
                            '   Get the Current Values from the ucPickBox
                            If .ucPickBox1.Path = "\" Then
                                .ucPickBox1.Path = App.Path
                            End If
                            .txtTextEntry.Text = .ucPickBox1.Path
                            '   Enalbe the Label & Textbox
                            .lblTextEntry.Enabled = True
                            .txtTextEntry.Enabled = True
                            '   Set the BackColor to White
                            .txtTextEntry.BackColor = &HFFFFFF
                        End If
                    Case 6
                        '   Set the Label Caption
                        .lblTextEntry.Caption = "Print Status Message"
                        '   Resize the Textbox to Accomidate the Droplist width
                        .txtTextEntry.Width = 1455
                        '   Make Visible the Label and DropList
                        .lblMessageOn.Visible = True
                        .cmbPrintStatusMsg.Visible = True
                        .txtTextEntry.Text = .ucPickBox1.PrintStatusMsg(.cmbPrintStatusMsg.ListIndex)
                        If .cmbType.ListIndex = 4 Then
                            '   Enable the Label & Textbox
                            .lblTextEntry.Enabled = True
                            .txtTextEntry.Enabled = True
                            '   Set the BackColor to White
                            .txtTextEntry.BackColor = &HFFFFFF
                            .lblMessageOn.Enabled = True
                            .cmbPrintStatusMsg.Enabled = True
                        Else
                            '   Disable the Label & Textbox
                            .lblTextEntry.Enabled = False
                            .txtTextEntry.Enabled = False
                            '   Set the BackColor to look Disabled
                            .txtTextEntry.BackColor = vbButtonFace
                            '   Disable the Label and DropList
                            .lblMessageOn.Enabled = False
                            .cmbPrintStatusMsg.Enabled = False
                        End If
                    Case Else
                        '   Set the Label Caption
                        .lblTextEntry.Caption = "ToolTipText"
                        .txtTextEntry.Text = .ucPickBox1.ToolTipTexts(.cmbType.ListIndex)
                        '   Enalbe the Label & Textbox
                        .lblTextEntry.Enabled = True
                        .txtTextEntry.Enabled = True
                        '   Set the BackColor to White
                        .txtTextEntry.BackColor = &HFFFFFF
                End Select
            Case 8, 9, 10    '(8 = UseAutoForeColor; 9 = UseDialogColor; 10 = UseDialogText)
                .lblChkBoxOptions.Enabled = True
                '   If this is ucColor & UseDialogColor
                Select Case m_ListIndex
                    Case 8
                        '   Set the caption
                        .chkUseDialog.Caption = "AutoForeColor"
                        '   Get the Value
                        .chkUseDialog.Value = Abs(.ucPickBox1.UseAutoForeColor)
                        '   Enable the Checkbox
                        .chkUseDialog.Enabled = True
                    Case 9
                        If (.cmbType.ListIndex = 0) Then
                            '   This is for ucColor
                            '   Set the caption
                            .chkUseDialog.Caption = "Color"
                            '   Get the Value
                            .chkUseDialog.Value = Abs(.ucPickBox1.UseDialogColor)
                            '   Enable the Checkbox
                            .chkUseDialog.Enabled = True
                        Else
                            '   All other Dailog Types...
                            '   Disable it
                            .lblChkBoxOptions.Enabled = False
                            '   Set the caption
                            .chkUseDialog.Caption = "Color"
                            '   Get the Value
                            .chkUseDialog.Value = Abs(.ucPickBox1.UseDialogColor)
                            '   Disable the Checkbox
                            .chkUseDialog.Enabled = False
                        End If
                    Case 10
                        '   Set the caption
                        .chkUseDialog.Caption = "Text"
                        '   Get the Value
                        .chkUseDialog.Value = Abs(.ucPickBox1.UseDialogText)
                        '   Enable the Checkbox
                        .chkUseDialog.Enabled = True
                End Select
            Case 11      '(11 = Version)
                .SetFocus
                MsgBox "Product:" & vbTab & vbTab & "ucPickBox.ctl" & vbNewLine & "Version:" & vbTab & vbTab & .ucPickBox1.Version & vbNewLine & "Build Date; Time:" & vbTab & BuildDateTime & vbNewLine & "Developer:" & vbTab & "Paul R. Territo, Ph.D", vbInformation + vbOKOnly, "ucPickBox - Build Information"
        End Select
    End With
End Sub

Private Sub ucPickBox1_KeyPress(KeyAscii As Integer)
    '   Fire the event
    Debug.Print "KeyPress: " & KeyAscii
End Sub

Private Sub ucPickBox1_KeyUp(KeyCode As Integer, Shift As Integer)
    '   Fire the event
    Debug.Print "KeyUp: " & KeyCode & ", " & Shift
End Sub

Private Sub ucPickBox1_LostFocus()
    '   Fire the event
    Debug.Print "LostFocus: ucPickBox"
End Sub

Private Sub ucPickBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Fire the event
    Debug.Print "MouseDown: " & Button & ", " & Shift & ", " & X & ", " & Y
End Sub

Private Sub ucPickBox1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Fire the event
    Debug.Print "MouseMove: " & Button & ", " & Shift & ", " & X & ", " & Y
End Sub

Private Sub ucPickBox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '   Fire the event
    Debug.Print "MouseUp: " & Button & ", " & Shift & ", " & X & ", " & Y
End Sub

Private Sub ucPickBox1_PathChanged()
    Dim i As Long
    '   Fire the event
    If ucPickBox1.FileCount > 0 Then
        '   List all of the Paths that changed....
        If ucPickBox1.MultiSelect Then
            For i = 1 To ucPickBox1.FileCount
                Debug.Print "PathChanged: " & ucPickBox1.Filename(i)
            Next
        Else
            Debug.Print "PathChanged: " & ucPickBox1.Filename(1)
        End If
    End If
End Sub

Private Sub ucPickBox1_PrintStatus(Status As String)
    '   Fire the event
    Debug.Print "PrintStaus: " & Status
End Sub
