VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ucPickBox - TestHarness"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaveFilename 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox txtOpenFilename 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   5175
   End
   Begin VB.ListBox lstFontSettings 
      Height          =   645
      Left            =   2880
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2535
   End
   Begin Project1.ucPickBox ucPickBox1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _extentx        =   4471
      _extenty        =   556
   End
   Begin VB.Label Label6 
      Caption         =   "ucPickBox Control:"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Dialog Type:"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Save File:"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Open File:"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblFont 
      Caption         =   "Font:"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Color:"
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
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
'
'   Force Declarations
Option Explicit

Private Sub cmbType_Click()
    With Me
        .ucPickBox1.DialogType = .cmbType.ListIndex
    End With
End Sub

Private Sub Form_Load()
    With Me
        With .cmbType
            .AddItem "ucColor"
            .AddItem "ucFont"
            .AddItem "ucOpen"
            .AddItem "ucSave"
            .AddItem "ucPrint"
            .ListIndex = 0
        End With
    End With
End Sub

Private Sub ucPickBox1_Click()
    Dim NewFont As StdFont
    
    With Me
        Select Case .ucPickBox1.DialogType
            Case [ucColor]
                .lblColor.BackColor = .ucPickBox1.Color
            Case [ucFont]
                Set NewFont = .ucPickBox1.Font
                If Not NewFont Is Nothing Then
                    With .lstFontSettings
                        .Clear
                        .AddItem "Name " & vbTab & "(" & NewFont.Name & ")"
                        .AddItem "Bold " & vbTab & "(" & NewFont.Bold & ")"
                        .AddItem "Italic " & vbTab & "(" & NewFont.Italic & ")"
                        .AddItem "Size " & vbTab & "(" & Round(NewFont.Size, 0) & ")"
                        .AddItem "Strikethrough " & vbTab & "(" & NewFont.Strikethrough & ")"
                        .AddItem "Underline " & vbTab & "(" & NewFont.Underline & ")"
                        .AddItem "Weight " & vbTab & "(" & NewFont.Weight & ")"
                        Set .Font = ucPickBox1.Font
                        .Height = 700
                    End With
                Else
                    MsgBox "No New Font Selected....", vbExclamation + vbOKOnly, "ucPickBox"
                End If
            Case [ucOpen]
                .txtOpenFilename.Text = .ucPickBox1.Filename
            Case [ucSave]
                .txtSaveFilename.Text = .ucPickBox1.Filename
            Case [ucPrint]
                
        End Select
    End With
End Sub

Private Sub ucPickBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '   ToDo...add additional validation code to validate the data.
        Call ucPickBox1_Click
    End If
End Sub
