VERSION 5.00
Begin VB.Form frmTestHarness 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ucPickBox - TestHarness"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   Icon            =   "frmTestHarness.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTextEntry 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox cmbPrintStatusMsg 
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstProperties 
      Height          =   1035
      Left            =   120
      TabIndex        =   33
      Top             =   2400
      Width           =   2655
   End
   Begin prjTestHarness.ucPickBox pbColorPick 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Top             =   2400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      Color           =   0
      DialogMsg       =   ""
      DialogMsg       =   ""
      DialogMsg       =   ""
      DialogMsg       =   ""
      DialogMsg       =   ""
      Printer         =   ""
      PrintStatusMsg  =   ""
      PrintStatusMsg  =   ""
      Object.ToolTipText     =   ""
      Object.ToolTipText     =   ""
      Object.ToolTipText     =   ""
      Object.ToolTipText     =   ""
      Object.ToolTipText     =   ""
   End
   Begin VB.CheckBox chkUseDialog 
      Caption         =   "Color "
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin prjTestHarness.ucPickBox ucPickBox1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
   End
   Begin VB.PictureBox picOpenSave 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   5415
      TabIndex        =   22
      Top             =   3960
      Width           =   5415
      Begin VB.TextBox txtOpenFilename 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox txtFileExists 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
      Begin VB.TextBox txtSavePath 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtSaveFilename 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   0
         Width           =   2535
      End
      Begin VB.TextBox txtOpenPath 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label lblFileExistsLabel 
         Caption         =   "File Exists:"
         Height          =   315
         Left            =   3480
         TabIndex        =   31
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblOpenLabel 
         Caption         =   "Open Path:"
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   30
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
      Begin VB.Label lblSaveLabel 
         Caption         =   "Save Path:"
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblSaveLabel 
         Caption         =   "Save File:"
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   24
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
      ScaleWidth      =   5415
      TabIndex        =   14
      Top             =   3960
      Width           =   5415
      Begin VB.TextBox txtColorValue 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblColorLabel 
         Caption         =   "Color Vaue (Hex):"
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblColorLabel 
         Caption         =   "Color:"
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox picFont 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   5415
      TabIndex        =   19
      Top             =   3960
      Width           =   5415
      Begin VB.ListBox lstFontSettings 
         BackColor       =   &H8000000F&
         Height          =   645
         Left            =   0
         TabIndex        =   20
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label lblFontLabel 
         Caption         =   "Font:"
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Label lblMessageOn 
      Caption         =   "Message On:"
      Height          =   315
      Left            =   4440
      TabIndex        =   36
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblProperties 
      Caption         =   "Property Type:"
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label13 
      Caption         =   $"frmTestHarness.frx":038A
      Height          =   1095
      Left            =   2760
      TabIndex        =   12
      Top             =   110
      Width           =   2775
   End
   Begin VB.Label Label12 
      Caption         =   "ucPickBox Properties"
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   1200
      X2              =   5280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   1200
      X2              =   5280
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Label lblChkBoxOptions 
      Caption         =   "UseDialog Color/Text:"
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblColorPick 
      Caption         =   "BackColor:"
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblPickResults 
      AutoSize        =   -1  'True
      Caption         =   "ucPickBox Results"
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   1200
      X2              =   5280
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   1200
      X2              =   5280
      Y1              =   3765
      Y2              =   3765
   End
   Begin VB.Label lblTextEntry 
      Caption         =   "Dialog Message:"
      Height          =   315
      Left            =   2880
      TabIndex        =   5
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "ucPickBox Control:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Dialog Type:"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblPrintStatus 
      Caption         =   "Print Status:"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
   End
End
Attribute VB_Name = "frmTestHarness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+  File Description:
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
'
'   Legal Copyright & Trademarks:
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
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
'       05Nov05 - Initial test harness and usercontrol finished
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
'               - Added addtional features to the TestHarness to allow for
'                 full exploring of the control.
'               - Changes Color from Long to OLE_COLOR property to allow for
'                 vb stanard palette
'       20Nov05 - Reworked the flow of the TestHareness to reduce the confusion of
'                 which properties work with which DialogType.
'               - Moved all Control Handling for the Properties and Results sections
'                 to the UpdatePropertyDialogs subroutine to simplify and reuse common
'                 code sections.
'               - Added Color RollBack if the value entered is invalid.
'
'   Force Declarations
Option Explicit
Private bLoading    As Boolean
Private m_ListIndex As Long

Private Sub AutoSelText(TxtBox As TextBox)
    With TxtBox
        '   Select the text
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
End Sub

Private Sub chkUseDialog_Click()
    With Me
        If Not bLoading Then
            If m_ListIndex = 5 Then
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
            Else
                '   All else...
                .ucPickBox1.UseDialogText = Abs(.chkUseDialog.Value)
            End If
        End If
    End With
End Sub

Private Sub cmbPrintStatusMsg_Click()
    With Me
        '   Call our custom updating routine
        Call UpdatePropertyDialogs(.cmbType.ListIndex)
    End With
End Sub

Private Sub cmbType_Click()
    With Me
        If Not bLoading Then
            '   Pass the Type to the Control
            .ucPickBox1.DialogType = .cmbType.ListIndex
            '   Update the Dialogs...
            Call UpdatePropertyDialogs(.cmbType.ListIndex)
        End If
    End With
End Sub

Private Sub Form_Load()
    With Me
        '   Prevent recursion by flagging the start...
        bLoading = True
        With .cmbType
            '   Give them some choices
            .AddItem "ucColor"              '0
            .AddItem "ucFont"               '1
            .AddItem "ucOpen"               '2
            .AddItem "ucSave"               '3
            .AddItem "ucPrint"              '4
            '   This is our default
            .ListIndex = 0
        End With
        With .lstProperties
            .Clear
            .AddItem "BackColor"            '0
            .AddItem "DialogMsg"            '1
            .AddItem "ForeColor"            '2
            .AddItem "PrintStatusMsg"       '3
            .AddItem "ToolTipTexts"         '4
            .AddItem "UseDialogColor"       '5
            .AddItem "UseDialogText"        '6
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

Private Sub lstProperties_Click()
    With Me
        '   Call our custom Updating routine....
        Call UpdatePropertyDialogs(.cmbType.ListIndex)
    End With
End Sub

Private Sub pbColorPick_Click()
    With Me
        On Error Resume Next
        Select Case m_ListIndex
            Case 0
                '   Set the New Selected Color
                .ucPickBox1.BackColor = .pbColorPick.Color
            Case 2
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
        Select Case m_ListIndex
            Case 1
                '   Set the PickBox Dialog Message
                .ucPickBox1.DialogMsg(.cmbType.ListIndex) = .txtTextEntry.Text
            Case 3
                '   Set the PickBox Print Status Message
                .ucPickBox1.PrintStatusMsg(.cmbPrintStatusMsg.ListIndex) = .txtTextEntry.Text
            Case 4
                '   Set the PickBox ToolTipText Message
                .ucPickBox1.ToolTipTexts(.cmbType.ListIndex) = .txtTextEntry.Text
        End Select
    End With
End Sub

Private Sub ucPickBox1_Click()
    Dim NewFont As StdFont
    
    With Me
        Select Case .ucPickBox1.DialogType
            Case [ucColor]
                '   Update the new Color of the Label
                .lblColor.BackColor = .ucPickBox1.Color
                '   Set the hex color value to the control
                .txtColorValue.Text = .ucPickBox1.LongToHexColor(.ucPickBox1.Color)
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
                        Set .Font = ucPickBox1.Font
                        '   Set our height
                        .Height = 700
                    End With
                Else
                    MsgBox "No New Font Selected....", vbExclamation + vbOKOnly, "ucPickBox"
                End If
            Case [ucOpen]
                '   Pass the Open Filename data to a text box for display
                .txtOpenFilename.Text = .ucPickBox1.ExtractFilename(.ucPickBox1.Filename)
                .txtOpenPath.Text = .ucPickBox1.TrimPathByLen(.ucPickBox1.ExtractPath(.ucPickBox1.Filename), .txtOpenPath.Width)
                '   Does this file exist? Should be, this is an Open
                .txtFileExists.Text = .ucPickBox1.FileExists(.ucPickBox1.Filename)
            Case [ucSave]
                '   Pass the Save Filename data to a text box for display
                .txtSaveFilename.Text = .ucPickBox1.ExtractFilename(.ucPickBox1.Filename)
                .txtSavePath = .ucPickBox1.ExtractPath(.ucPickBox1.Filename)
                '   Does this file exist? Might not, if we just changed the name...
                .txtFileExists.Text = .ucPickBox1.FileExists(.ucPickBox1.Filename)
            Case [ucPrint]
                '   Print the form image...
                .lblPrintStatus.Caption = "Print Status: " & .ucPickBox1.PrintStatusMsg(Abs(.ucPickBox1.PrintStatus))
                '   Just to show it works
                '.PrintForm
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
    With Me
        '   Update the Controls in the Results Area....
        '
        '   Turn off all dialogs and labels....
        .lblPrintStatus.Caption = "Print Status: "
        .picColor.Visible = False
        .picFont.Visible = False
        .picOpenSave.Visible = False
        .lblOpenLabel(0).Visible = False
        .lblOpenLabel(1).Visible = False
        .txtOpenFilename.Visible = False
        .txtOpenPath.Visible = False
        .lblSaveLabel(0).Visible = False
        .lblSaveLabel(1).Visible = False
        .txtSaveFilename.Visible = False
        .txtSavePath.Visible = False
        .lblPrintStatus.Visible = False
        .chkUseDialog.Enabled = False
        '   Now select only the ones needed based on the DialogType in
        '   the Results section...
        Select Case lDialogType
            Case [ucColor]
                .lblPickResults.Caption = "ucPickBox - Color Results "
                .picColor.Visible = True
            Case [ucFont]
                .lblPickResults.Caption = "ucPickBox - Font Results "
                .lblFontLabel.Visible = True
                .lstFontSettings.Visible = True
                .picFont.Visible = True
            Case [ucOpen]
                .lblPickResults.Caption = "ucPickBox - Open Results "
                .lblOpenLabel(0).Visible = True
                .lblOpenLabel(1).Visible = True
                .txtOpenFilename.Visible = True
                .txtOpenPath.Visible = True
                .picOpenSave.Visible = True
            Case [ucSave]
                .lblPickResults.Caption = "ucPickBox - Save Results "
                .lblSaveLabel(0).Visible = True
                .lblSaveLabel(1).Visible = True
                .txtSaveFilename.Visible = True
                .txtSavePath.Visible = True
                .picOpenSave.Visible = True
            Case [ucPrint]
                .lblPickResults.Caption = "ucPickBox - Print Results "
                .lblPrintStatus.Visible = True
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
        '   Store the value for later...
        m_ListIndex = .lstProperties.ListIndex
        Select Case m_ListIndex
            Case 0, 2           '(0 = BackColor; 2 = ForeColor)
                If m_ListIndex = 0 Then
                    '   This is for BackColor....so indicate it
                    .lblColorPick.Caption = "BackColor"
                    '   Get the value
                    If .chkUseDialog.Value = vbChecked Then
                        '   Since the Color is the BackColor
                        .pbColorPick.DialogMsg(.cmbType.ListIndex) = .ucPickBox1.LongToHexColor(.ucPickBox1.Color)
                    Else
                        '   Just the BackColor
                        .pbColorPick.DialogMsg(.cmbType.ListIndex) = .ucPickBox1.LongToHexColor(.ucPickBox1.BackColor)
                    End If
                Else
                    '   This is for ForeColor....so indicate it
                    .lblColorPick.Caption = "ForeColor"
                    '   Get the value
                    .pbColorPick.DialogMsg(.cmbType.ListIndex) = .ucPickBox1.LongToHexColor(.ucPickBox1.ForeColor)
                End If
                '   Enable the Label and ucPickBox
                .lblColorPick.Enabled = True
                .pbColorPick.Enabled = True
                '   Get the backcolor to white
                .pbColorPick.BackColor = &HFFFFFF
            Case 1, 3, 4       '(1 = DialogMsg; 3 = PrintStatusMsg; 4 = ToolTipTexts)
                .lblMessageOn.Visible = False
                .txtTextEntry.Width = 2535
                Select Case m_ListIndex
                    Case 1
                        '   Set the Label Caption
                        .lblTextEntry.Caption = "Dialog Message"
                        '   Get the Current Values from the ucPickBox
                        .txtTextEntry.Text = .ucPickBox1.DialogMsg(.cmbType.ListIndex)
                        '   Enalbe the Label & Textbox
                        .lblTextEntry.Enabled = True
                        .txtTextEntry.Enabled = True
                        '   Set the BackColor to White
                        .txtTextEntry.BackColor = &HFFFFFF
                    Case 3
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
            Case 5, 6      '(5 = UseDialogColor; 6 = UseDialogText)
                .lblChkBoxOptions.Enabled = True
                '   If this is ucColor & UseDialogColor
                If (m_ListIndex = 5) Then
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
                Else
                    '   Set the caption
                    .chkUseDialog.Caption = "Text"
                    '   Get the Value
                    .chkUseDialog.Value = Abs(.ucPickBox1.UseDialogText)
                    '   Enable the Checkbox
                    .chkUseDialog.Enabled = True
                End If
        End Select
    End With
End Sub

Private Sub ucPickBox1_LostFocus()
    '   Fire the event
    Debug.Print "LostFocus: ucPickBox"
End Sub

Private Sub ucPickBox1_PathChanged(NewPath As String)
    '   Fire the event
    Debug.Print "PathChanged: " & NewPath
End Sub

Private Sub ucPickBox1_PrintStatus(Status As String)
    '   Fire the event
    Debug.Print "PrintStaus: " & Status
End Sub
