VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCodeGen 
   Caption         =   "Easy Tool Tips Class - Custom Tool Tip Code"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8415
   Icon            =   "frmCodeGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      MouseIcon       =   "frmCodeGen.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frmCodeGen.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customize Your Tool Tip"
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         MouseIcon       =   "frmCodeGen.frx":0A56
         MousePointer    =   99  'Custom
         Picture         =   "frmCodeGen.frx":0D60
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Customize !"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4940
         MouseIcon       =   "frmCodeGen.frx":11A2
         MousePointer    =   99  'Custom
         Picture         =   "frmCodeGen.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   990
         Width           =   2400
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   530
         Width           =   2400
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Parent Control Name"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   990
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tool Tip Name"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   530
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Put your Tool Tip Object name and Parent Object name in the boxes below to customize your code"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   7695
      End
   End
   Begin RichTextLib.RichTextBox CodeGenText 
      CausesValidation=   0   'False
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   32767
      TextRTF         =   $"frmCodeGen.frx":1D76
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuTooltip 
         Caption         =   "Easy Tool Tip Class"
      End
      Begin VB.Menu mnuCodegen 
         Caption         =   "Code Generator"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAboutETC 
         Caption         =   "About Easy Tool Tip Class"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSelectall 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo Selection"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmCodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '**************************************************************
    '
    '   Custom Tool Tip Demo
    '   Code Generation form
    '
    '   Mark Mokoski
    '   27-NOV-2004
    '
    '   Takes Tool Tip info from frmToolTips and produces code for
    '   cut and paste into your project
    '
    '**************************************************************
    
    Option Explicit
    Dim Picture1tip            As New clsTooltips



Private Sub CodeGenText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        'Look for Right Click
        If Button = 2 Then

                If CodeGenText.SelText <> "" Then
                    mnuCopy.Enabled = True
                    mnuSelectall.Enabled = False
                    mnuUndo.Enabled = True
                    PopupMenu mnuEdit
                Else
                    mnuCopy.Enabled = False
                    mnuSelectall.Enabled = True
                    mnuUndo.Enabled = False
                    PopupMenu mnuEdit
                End If

        End If

End Sub

Private Sub Command1_Click()

    Call GenCode

End Sub


Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Command1.Enabled = False
    Command1.BackColor = &H8000000F
    
    
    Picture1tip.CreateBalloon Picture1, _
    "Put the Tool Tip name and Parent Conrol name in the text boxes to the right." & _
    vbCrLf & _
    vbCrLf & _
    "Then Click the Customize buttom to update the code snippet", _
    "Easy Tool Tip Code Generator", _
    tipiconinfo
    
End Sub

Private Sub mnuAbout_Click()

    'Bring up the About info window
    frmAbout.Visible = True
    

End Sub

Private Sub mnuAboutETC_Click()

    frmAbout.Visible = True
    frmAbout.SetFocus

End Sub

Private Sub mnuClose_Click()

    Unload Me

End Sub

Private Sub mnuCopy_Click()

    Clipboard.Clear
    SendKeys "^C", True
    'To copy selection and put cursor at end of selected text, use below
    'CodeGenText.SelStart = CodeGenText.SelStart + (Len(CodeGenText.SelText) - 1)
    
    'To copy selection and put cursor at beginning of selected text, use below
    CodeGenText.SelStart = CodeGenText.SelStart
    
    'To copy selection and put cursor at end of all text, use below
    'CodeGenText.SelStart = (Len(CodeGenText.Text) + 1)
    
    'To copy selection and put cursor at beginning of all text, use below
    'CodeGenText.SelStart = 0

End Sub

Private Sub mnuEdit_Click()

        If CodeGenText.SelText <> "" Then
            mnuCopy.Enabled = True
            mnuSelectall.Enabled = False
            mnuUndo.Enabled = True
        Else
            mnuCopy.Enabled = False
            mnuSelectall.Enabled = True
            mnuUndo.Enabled = False
        End If

End Sub

Private Sub mnuSelectall_Click()

    CodeGenText.SetFocus
    CodeGenText.SelStart = 0

    CodeGenText.SelLength = Len(CodeGenText.Text)
    

    SendKeys "^A", True

End Sub

Private Sub mnuTooltip_Click()

    frmToolTips.Visible = True
    frmToolTips.SetFocus

End Sub

Private Sub mnuUndo_Click()

    'To cancel selection and put cursor at end of selected text, use below
    'CodeGenText.SelStart = CodeGenText.SelStart + (Len(CodeGenText.SelText) - 1)
    
    'To cancel selection and put cursor at beginning of selected text, use below
    CodeGenText.SelStart = CodeGenText.SelStart
    
    'To cancel selection and put cursor at end of all text, use below
    'CodeGenText.SelStart = (Len(CodeGenText.Text) + 1)
    
    'To cancel selection and put cursor at beginning of all text, use below
    'CodeGenText.SelStart = 0
    
End Sub

Private Sub Text1_Change()

        If Text1.Text <> "" And Text2.Text <> "" Then
            Command1.Enabled = True
            Command1.BackColor = &HC0C0C0
        Else
            Command1.Enabled = False
            Command1.BackColor = &H8000000F
        End If
        
End Sub

Private Sub Text2_Change()

        If Text1.Text <> "" And Text2.Text <> "" Then
            Command1.Enabled = True
            Command1.BackColor = &HC0C0C0
        Else
            Command1.Enabled = False
            Command1.BackColor = &H8000000F
        End If

End Sub

Public Sub GenCode()

    Dim TipName              As String
    Dim TipParent            As String
    Dim TipText              As String
    
    frmCodeGen.Visible = True
    Text1.SetFocus
    
    'Replace vbCrLF code (Chr$(10)+Chr$(13)) with " & vbCrLf & " text
    'for proper string format in code generation
    TipText = ReplaceText(frmToolTips.TipText)
    
    'Clean out any current text
    CodeGenText.SelStart = 0
    CodeGenText.SelLength = Len(CodeGenText.Text) + 1
    CodeGenText.SelText = ""
    'Get the boilerplate text and insert date and time
    CodeGenText.LoadFile App.Path & "\codegen.rtf"
    CodeGenText.SelStart = 168
    CodeGenText.SelText = Date & " at " & Time

        If Text1.Text = "" Then
            TipName = "<Your Tip Name>"
        Else
            TipName = Text1.Text
        End If
        
        If Text2.Text = "" Then
            TipParent = "<Your Parent Control Name>"
        Else
            TipParent = Text2.Text
        End If
        
    'Write out the Declarations section

        With CodeGenText
            .SelStart = 726
            .SelColor = vbBlue
            .SelText = vbCrLf & "Dim "
            .SelColor = vbBlack
            .SelText = TipName & vbTab & vbTab
            .SelColor = vbBlue
            .SelText = "As New "
            .SelColor = vbBlack
            .SelText = "  clsTooltips"
            .SelText = vbCrLf
    
            'Write out the Code section
            .SelStart = 1044
            .SelColor = vbBlack
            .SelText = vbCrLf
            
                If frmToolTips.TipStyle = styleBalloon Then
                    .SelText = TipName & ".CreateBalloon " & TipParent & ", _" & vbCrLf & """" & TipText & """"
                Else
                    .SelText = TipName & ".CreateTip " & TipParent & ", _" & vbCrLf & """" & TipText & """"
                End If
        
                If frmToolTips.TipTitleText <> "" Then
                    .SelText = ", _" & vbCrLf & """" & frmToolTips.TipTitleText & """, " & Val(frmToolTips.TipIcon)
                End If
                
            .SelText = vbCrLf
                
                If frmToolTips.TipCentered = True Then
                    .SelText = TipName & ".Centered = "
                    .SelColor = vbBlue
                    .SelText = "True" & vbCrLf
                    .SelColor = vbBlack
                    
                End If
                
                If frmToolTips.TipForeColor <> 0 Then
                    .SelColor = vbBlack
                    .SelText = TipName & ".ForeColor = " & "&H" & Hex(frmToolTips.TipForeColor) & vbCrLf
                End If
                
                If frmToolTips.TipBackColor <> 0 Then
                    .SelColor = vbBlack
                    .SelText = TipName & ".BackColor = " & "&H" & Hex(frmToolTips.TipBackColor) & vbCrLf
                End If
                
            .SelText = vbCrLf
            .SelStart = 0
            
        End With
                                

End Sub

Private Function ReplaceText(rText As String)
    
    'Replace Tool Tip Text with more verbose string. Add "& vbCrLf &"
    'string in place of Chr$(10)+Chr$(13)
    ReplaceText = Replace(rText, vbCrLf, """ & vbCrLf &  _" & vbCrLf & """")

End Function
