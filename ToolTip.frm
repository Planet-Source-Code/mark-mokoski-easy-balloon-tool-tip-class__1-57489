VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmToolTips 
   Caption         =   "Easy Tool Tips Class"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9015
   Icon            =   "ToolTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Select Tip Icon"
      ForeColor       =   &H00000080&
      Height          =   2295
      Left            =   6480
      TabIndex        =   27
      Top             =   2280
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "ToolTip.frx":1272
         MousePointer    =   99  'Custom
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1560
         MouseIcon       =   "ToolTip.frx":157C
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":1886
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1440
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ERROR"
         Height          =   255
         Index           =   3
         Left            =   1320
         MouseIcon       =   "ToolTip.frx":1CC8
         MousePointer    =   99  'Custom
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         MouseIcon       =   "ToolTip.frx":1FD2
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":22DC
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1440
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Information"
         Height          =   255
         Index           =   2
         Left            =   120
         MouseIcon       =   "ToolTip.frx":271E
         MousePointer    =   99  'Custom
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1560
         MouseIcon       =   "ToolTip.frx":2A28
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":2D32
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Warning"
         Height          =   255
         Index           =   1
         Left            =   1320
         MouseIcon       =   "ToolTip.frx":3174
         MousePointer    =   99  'Custom
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tool Tip Style"
      ForeColor       =   &H00000080&
      Height          =   2175
      Left            =   6480
      TabIndex        =   20
      Top             =   0
      Width           =   2415
      Begin VB.CheckBox Check2 
         Caption         =   "Center Tool Tip"
         Height          =   195
         Left            =   840
         MouseIcon       =   "ToolTip.frx":347E
         MousePointer    =   99  'Custom
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Rectangular Tip"
         Height          =   195
         Index           =   0
         Left            =   840
         MouseIcon       =   "ToolTip.frx":3788
         MousePointer    =   99  'Custom
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Balloon Tip"
         Height          =   195
         Index           =   1
         Left            =   840
         MouseIcon       =   "ToolTip.frx":3A92
         MousePointer    =   99  'Custom
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   80
         MouseIcon       =   "ToolTip.frx":3D9C
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":40A6
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   80
         MouseIcon       =   "ToolTip.frx":43B0
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":46BA
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   80
         MouseIcon       =   "ToolTip.frx":49C4
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":4CCE
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Send Mark Email about Easy Tool Tip Class"
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   4320
      TabIndex        =   13
      Top             =   4680
      Width           =   4575
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         MouseIcon       =   "ToolTip.frx":4FD8
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":52E2
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   720
         Locked          =   -1  'True
         MouseIcon       =   "ToolTip.frx":5BAC
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "ToolTip.frx":5EB6
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MouseIcon       =   "ToolTip.frx":5EE4
         MousePointer    =   99  'Custom
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "ToolTip.frx":61EE
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tool Tip Title"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   6015
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   140
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "(Title needed for Icons)"
         Height          =   200
         Left            =   120
         TabIndex        =   19
         Top             =   200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tool Tip Text"
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   6015
      Begin VB.TextBox BackColorRGB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "Text6"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox ForeColorRGB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Background Color"
         Height          =   375
         Left            =   120
         MouseIcon       =   "ToolTip.frx":6227
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Text Color"
         Height          =   375
         Left            =   120
         MouseIcon       =   "ToolTip.frx":6531
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   1245
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "ToolTip.frx":683B
         Top             =   240
         Width           =   3975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Test Tool Tip Parameters"
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6255
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Click To Generate Code"
         Height          =   950
         Left            =   4300
         MouseIcon       =   "ToolTip.frx":684C
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":6B56
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1870
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test Your Tool Tip Here!"
         CausesValidation=   0   'False
         Height          =   950
         Left            =   2225
         MouseIcon       =   "ToolTip.frx":6F98
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":72A2
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1870
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apply Changes"
         Height          =   950
         Left            =   120
         MouseIcon       =   "ToolTip.frx":79E4
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":7CEE
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1870
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Disable All Tool Tips"
      Height          =   915
      Left            =   2280
      MouseIcon       =   "ToolTip.frx":8130
      MousePointer    =   99  'Custom
      Picture         =   "ToolTip.frx":843A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enable All Tool Tips"
      Height          =   915
      Left            =   240
      MouseIcon       =   "ToolTip.frx":8744
      MousePointer    =   99  'Custom
      Picture         =   "ToolTip.frx":8A4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Close Program"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuTray 
      Caption         =   "TrayMnu"
      Visible         =   0   'False
      Begin VB.Menu mnuTrestore 
         Caption         =   "Restore Easy Tool Tips Window"
      End
      Begin VB.Menu mnuThide 
         Caption         =   "Hide Easy Tool Tips Window in SysTray"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTabout 
         Caption         =   "About Tool Tip Class"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTclose 
         Caption         =   "Close this Menu <ESC>"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTexit 
         Caption         =   "Exit Easy Tool Tip Class"
      End
   End
End
Attribute VB_Name = "frmToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '**************************************************************
    '
    '   Custom Tool Tip Demo
    '
    '   Mark Mokoski
    '   16-NOV-2004
    '
    '   See clsToolTips Class Module for details
    '
    '**************************************************************

    Option Explicit

    'Make new tool tip objects for this project

    Dim command1tip                As New clsTooltips
    Dim Command2Tip                As New clsTooltips
    Dim Command3Tip                As New clsTooltips
    Dim Command4Tip                As New clsTooltips
    Dim Command7Tip                As New clsTooltips
    Dim Picture1tip                As New clsTooltips
    Dim Text1Tip                   As New clsTooltips
    Dim Text2Tip                   As New clsTooltips
    Dim Text3Tip                   As New clsTooltips
    Dim Text4Tip                   As New clsTooltips
    Dim ForeColorRGBtip            As New clsTooltips
    Dim BackColorRGBtip            As New clsTooltips
    Dim Check2Tip                  As New clsTooltips
    Dim Option2Tip(1)              As New clsTooltips
    Dim Picture5Tip                As New clsTooltips
    Dim Picture6Tip                As New clsTooltips
    Dim Picture7Tip                As New clsTooltips
    
    'Public Var's used in this and other modules
    Public TipText                 As String
    Public TipTitleText            As String
    Public TipCentered             As Boolean
    Public TipStyle                As toolStyleEnum
    Public TipIcon                 As toolIconType
    Public TipForeColor            As Long
    Public TipBackColor            As Long
    
    

Private Sub Check2_Click()

    TipCentered = CBool(Check2.Value)
    
End Sub

Private Sub Command1_Click()

    ' Activate the custom tooltip to the controls
    Command2Tip.Active = True
    Command3Tip.Active = True
    Command4Tip.Active = True
    Text1Tip.Active = True
    Text2Tip.Active = True
    Text3Tip.Active = True
    Text4Tip.Active = True
    
    'If you recreate a tool tip in code, you must set or reset the color values
    'or else the last colors are retained.  the properties are keyed on the
    'parent object (control), not the tool tip it self.
    Picture1tip.CreateBalloon Picture1, "Yup, I'm a Bug!" + vbCrLf + "Now let go of me!"
    Picture1tip.ForeColor = 0   'Default forecolor (Black)
    Picture1tip.BackColor = 0   'Default Backcolor ('Off' Yellow)
    
    Command1.Enabled = False
    Command1.BackColor = &H8000000F
    Command3.Enabled = True
    Command3.BackColor = &HC0C0C0
    
End Sub

Private Sub Command2_Click()

        If Text1.Text = "" And Text4.Text = "" Then
            MsgBox "Tool Tip Text ERROR" & vbCrLf & vbCrLf & _
            "Tool Tip Text and Tool Tip Title are Blank" & vbCrLf & _
            "For proper Tool Tip operation, Tip Text and/or a Tip Title is needed", vbCritical, "Tool Tip ERROR"
            Exit Sub
        Else

            'Change Tool Tip Text and other properties

                If Text1.Text = "" Then
                    Command4Tip.TipText = " "
                Else
                    Command4Tip.TipText = Text1.Text
                End If
            
            Command4Tip.Title = TipTitleText
            Command4Tip.Style = TipStyle
            Command4Tip.Centered = TipCentered
            Command4Tip.Icon = TipIcon
            Command4Tip.ForeColor = TipForeColor
            Command4Tip.BackColor = TipBackColor
        End If

End Sub

Private Sub Command3_Click()

    ' Deactivate the custom tooltip to the controls
    Command2Tip.Active = False
    Command3Tip.Active = False
    Command4Tip.Active = False
    Text1Tip.Active = False
    Text2Tip.Active = False
    Text3Tip.Active = False
    Text4Tip.Active = False
    
    'This code below was generated with the CodeGen feature, cut and pasted here
    Picture1tip.CreateBalloon Picture1, "Ooooooooch!" & vbCrLf & "Now put me down!", "It's not nice to touch the Bug!", 3
    Picture1tip.BackColor = &H4080FF


    Command1.Enabled = True
    Command1.BackColor = &HC0C0C0
    Command3.Enabled = False
    Command3.BackColor = &H8000000F
    
End Sub

Private Sub Command4_Click()

    Command2.SetFocus
    
End Sub

Private Sub Command5_Click()

    'Set new tip fore color
    'Set Cancel to True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Set the Flags property
    CommonDialog1.FLAGS = cdlCCRGBInit
    
    'Display the Color Dialog box
    CommonDialog1.ShowColor
    
    'Set the form's foreground color to selected color
    Text1.ForeColor = CommonDialog1.Color
    Text4.ForeColor = CommonDialog1.Color
    TipForeColor = CommonDialog1.Color
    
    ForeColorRGB.Text = "&H" & Hex(TipForeColor)
    'BackColorRGB.Text = "&H" & Hex(TipBackColor)
    
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    ForeColorRGB.Text = "&H" & Hex(Text1.ForeColor)
    'BackColorRGB.Text = "&H" & Hex(Text1.BackColor)

End Sub

Private Sub Command6_Click()

    'Set new tip back color
    'Set Cancel to True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Set the Flags property
    CommonDialog1.FLAGS = cdlCCRGBInit
    
    'Display the Color Dialog box
    CommonDialog1.ShowColor
    
    'Set the form's background color to selected color
    Text1.BackColor = CommonDialog1.Color
    Text4.BackColor = CommonDialog1.Color
  
    'Since 0 is Black (no RGB), and the API thinks 0 is
    'the default color ("off" yeleow),
    'we need to "fudge" Black a bit (yes set bit "1" to "1",)
    'I couldn't resist the pun!
    
        If CommonDialog1.Color = 0 Then
            TipBackColor = &H80000008
        Else
            TipBackColor = CommonDialog1.Color
        End If
    
    'ForeColorRGB.Text = "&H" & Hex(TipForeColor)
    BackColorRGB.Text = "&H" & Hex(TipBackColor)
    
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    'ForeColorRGB.Text = "&H" & Hex(Text1.ForeColor)
    BackColorRGB.Text = "&H" & Hex(Text1.BackColor)

End Sub

Private Sub Command7_Click()

        If Text1.Text = "" Then
            MsgBox "Tool Tip Text ERROR" & vbCrLf & vbCrLf & _
            "Tool Tip Text is Blank" & vbCrLf & _
            "For proper Tool Tip operation, Tip Text is needed", vbCritical, "Tool Tip ERROR"
            Exit Sub
        Else
            Call frmCodeGen.GenCode
            frmCodeGen.Visible = True
        End If
    
End Sub


Private Sub Form_Load()
    
    Dim X            As Integer

    'Make Tool Tip objects
    command1tip.CreateBalloon Command1, "OK, I turned off all the Tool Tips but this one" + vbCrLf + "Click to restore Tool Tips", "Tool Tips are OFF", tipIconWarning
    Command2Tip.CreateBalloon Command2, "Type in new Tip Text, Title and" + vbCrLf + "choose the other parameters." + vbCrLf + "Use more than one line of text if you want." + vbCrLf + "Click to apply your changes " + vbCrLf + "and test the Tool Tip", "Balloon Tip", tipiconinfo
    Command3Tip.CreateBalloon Command3, "Click to Hide all Tool Tips"
    Command4Tip.CreateTip Command4, "Go Ahead, Change ME!"
    Picture1tip.CreateBalloon Picture1, "Yup, I'm a Bug!" + vbCrLf + "Now let go of me!"
    Text1Tip.CreateBalloon Text1, "Enter Tool Tip Text Here" + vbCrLf + "Double Click to restore default colors", "Test Tip Text", tipiconinfo
    Text2Tip.CreateBalloon Text2, "Click to goto my web site"
    Text3Tip.CreateBalloon Text3, "Click to send Mark email"
    Text4Tip.CreateBalloon Text4, "Enter Tool Tip Title Here" + vbCrLf + "By entering a Title," + vbCrLf + "you enable the Tip Icon selection", "Test Title Text", tipiconinfo
    ForeColorRGBtip.CreateBalloon ForeColorRGB, "ForeColor RGB Hex Code"
    BackColorRGBtip.CreateBalloon BackColorRGB, "Backcolor RGB Hex Code"
    Check2Tip.CreateTip Check2, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
    Check2Tip.Centered = True
    Picture5Tip.CreateTip Picture5, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
    Picture5Tip.Centered = True
    Option2Tip(0).CreateTip Option2(0), "Tip looks like this."
    Picture6Tip.CreateTip Picture6, "Tip looks like this."
    Option2Tip(1).CreateBalloon Option2(1), "Tip looks like this."
    Picture7Tip.CreateBalloon Picture7, "Tip looks like this."
    
    'Code below make with this App's CodeGen feature!
    Command7Tip.CreateBalloon Command7, _
    "Click Here to make the code" & vbCrLf & _
    "for your custom ToolTip." & vbCrLf & _
    "" & vbCrLf & _
    "Cut and Paste into your project!", _
    "Create Code", 1

    Command7Tip.ForeColor = &HEFEFEF
    Command7Tip.BackColor = &HC08000

    'Set up what controls are active
    Command1.Enabled = False
    Command1.BackColor = &H8000000F
    Command2.Enabled = True
    Command3.Enabled = True

        For X = 0 To 3
            Option1(X).Enabled = False
        Next X
        
    'Put Icon in the SysTray
    Call SystrayOn(Me, "Form is ready to hide in the Tray!")
    PopupBalloon Me, "Form is now ready to be hidden in SysTray" & vbCrLf & "Right Click for Menu", "SysTray Icon Ready!"
    ForeColorRGB.Text = "&H" & Hex(Text1.ForeColor)
    
    'Set start values
    BackColorRGB.Text = "&H" & Hex(Text1.BackColor)
    TipForeColor = 0
    TipBackColor = 0
    TipText = Text1.Text
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static lngMsg            As Long
    Dim blnflag              As Boolean

    lngMsg = X / Screen.TwipsPerPixelX

        If blnflag = False Then

            blnflag = True
        
                Select Case lngMsg
                    Case WM_RBUTTONCLK      'to popup menu on right-click
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        'Reference the menu object of the form below for popup
                        PopupMenu Me.mnuTray

                    Case WM_LBUTTONDBLCLK   'SHow form on left-dblclick
                        'Use line below if you want to remove tray icon on dbclick show form.
                        'If not, be sure to put Systrayoff in form unload and terminate events.
                        'Call SystrayOff(Me)
                        Call SetForegroundWindow(Me.hWnd)
                        'Call RemoveBalloon(Me)
                        Me.WindowState = vbNormal
                        Me.Show
                        Me.SetFocus
            
                End Select

            blnflag = False
        
        End If
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'OK, I got lazy.  You can put all the Tool Tips in a collection
    'and do a For - Next loop to do this.
    'But there is a small number of controls and it shows the Remove Method
    'used to "kill" the object.  Good idea to do this just incase to help
    'prevent memory leaks (Microsoft says that won't happen, but we know better!)
    
    Command2Tip.Remove
    Command3Tip.Remove
    Command4Tip.Remove
    Picture1tip.Remove
    Text1Tip.Remove
    Text2Tip.Remove
    Text3Tip.Remove
    
End Sub

Private Sub Form_Resize()

        If Me.WindowState = vbMinimized Then
            'Use next line if the Tray Icon is removed  on Restore, see Mouse_Move sub
            'Call SystrayOn(Me, "Put your tool text tip here")
            Me.Hide
            PopupBalloon Me, "Form is hidden in SysTray." & vbCrLf & "Double click to restore" & vbCrLf & "Right Click for Menu", "Form is Hidden in the Tray"
            ChangeSystrayToolTip Me, "Form is hidden in SysTray."
        Else
            PopupBalloon Me, "Form is now ready to be hidden in SysTray" & vbCrLf & "Right Click for Menu", "SysTray Icon Ready!"
            ChangeSystrayToolTip Me, "Form is ready to hide in the SysTray."
        End If

End Sub

Private Sub Form_Terminate()

    Unload frmAbout
    Unload frmCodeGen

    Call SystrayOff(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmAbout
    Unload frmCodeGen
    
    Call SystrayOff(Me)

End Sub

Private Sub mnuAbout_Click()

    'Bring up the About info window
    frmAbout.Visible = True
    
End Sub

Private Sub mnuExit_Click()

    'From "Files" menu, "Exit"
    Unload frmToolTips
    
End Sub

Private Sub mnuTabout_Click()

    mnuAbout_Click

End Sub

Private Sub mnuTexit_Click()

    mnuExit_Click

End Sub

Private Sub mnuThide_Click()

    Me.WindowState = vbMinimized

End Sub

Private Sub mnuTrestore_Click()

    Call SetForegroundWindow(Me.hWnd)
    'Call RemoveBalloon(Me)
    Me.WindowState = vbNormal
    Me.Show
    Me.SetFocus
    
End Sub

Private Sub Option1_Click(Index As Integer)
    
    'FInd out what Tool Tip Icon was selected

        Select Case Index
            Case 0
                TipIcon = tipNoIcon
            Case 1
                TipIcon = tipIconWarning
            Case 2
                TipIcon = tipiconinfo
            Case 3
                TipIcon = tipIconError
        End Select
    
End Sub

Private Sub Option2_Click(Index As Integer)

    'Find out what Tool Tip Style was selected

        Select Case Index
            Case 0
                TipStyle = styleStandard
                Check2Tip.Remove
                Check2Tip.CreateTip Check2, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                Check2Tip.Centered = True
                Picture5Tip.CreateTip Picture5, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                Picture5Tip.Centered = True
            Case 1
                TipStyle = styleBalloon
                Check2Tip.Remove
                Check2Tip.CreateBalloon Check2, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                Check2Tip.Centered = True
                Picture5Tip.CreateBalloon Picture5, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                Picture5Tip.Centered = True
        End Select
        
End Sub




Private Sub Picture2_Click()

        If Option1(1).Enabled = True Then Option1(1).Value = True

End Sub

Private Sub Picture3_Click()

        If Option1(2).Enabled = True Then Option1(2).Value = True

End Sub

Private Sub Picture4_Click()

        If Option1(3).Enabled = True Then Option1(3).Value = True

End Sub

Private Sub Picture5_Click()

        If Check2.Value = 0 Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If


End Sub

Private Sub Picture6_Click()
Option2(0).Value = True
End Sub

Private Sub Picture7_Click()
Option2(1).Value = True
End Sub

Private Sub Text1_Change()

    TipText = Text1.Text

End Sub

Private Sub Text1_DblClick()

    'Restore text controls colors to default
    Text1.ForeColor = &H80000008
    Text4.ForeColor = &H80000008
    '"0" = default forecolor in API
    TipForeColor = 0
    Text1.BackColor = &H80000018
    Text4.BackColor = &H80000018
    '"0" = default backcolor in API
    ForeColorRGB.Text = "&H" & Hex(Text1.ForeColor)
    BackColorRGB.Text = "&H" & Hex(Text1.BackColor)
    TipForeColor = 0
    TipBackColor = 0
    
    Command2.SetFocus
    
End Sub

Private Sub Text2_click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus

    ShellExecute hWnd, vbNullString, "http://www.rjillc.com", vbNullString, vbNullString, vbNormalFocus
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub Text3_click()

    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    
    ShellExecute hWnd, vbNullString, "mailto:markm@cmtelephone.com?subject=Questions or Comments on Easy Balloon ToolTip Code.", vbNullString, vbNullString, vbNormalFocus
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)

End Sub

Private Sub Text4_Change()

    Dim X                    As Integer
    
    'See if text box is empty

        If Text4.Text <> "" Then
            'If not empty enable the Icon option buttons and set variable to the text
            TipTitleText = Text4.Text

                For X = 0 To 3
                    Option1(X).Enabled = True
                Next X

        Else
            'Text is empty, disable Icon option buttons and null out text variable
            TipTitleText = vbNullString

                For X = 0 To 3
                    Option1(X).Enabled = False
                Next X

        End If
    
End Sub
