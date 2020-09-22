VERSION 5.00
Begin VB.Form MsgBx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MsgB.jcFrames MBXFrame 
      Height          =   3015
      Left            =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5318
      BackColor       =   15783104
      FillColor       =   15783104
      MoverForm       =   -1  'True
      MoverControle   =   -1  'True
      Caption         =   "My Message Box"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin MsgB.isButton MButton 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Icon            =   "MsgBx.frx":0000
         Style           =   8
         Caption         =   "isButton1"
         iNonThemeStyle  =   8
         USeCustomColors =   -1  'True
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin MsgB.isButton MButton 
         Height          =   615
         Index           =   1
         Left            =   1920
         TabIndex        =   4
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Icon            =   "MsgBx.frx":001C
         Style           =   8
         Caption         =   "isButton1"
         iNonThemeStyle  =   8
         USeCustomColors =   -1  'True
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin MsgB.isButton MButton 
         Height          =   615
         Index           =   2
         Left            =   3720
         TabIndex        =   5
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Icon            =   "MsgBx.frx":0038
         Style           =   8
         Caption         =   "isButton1"
         iNonThemeStyle  =   8
         USeCustomColors =   -1  'True
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   630
      End
      Begin MsgB.aicAlphaImage Logo 
         Height          =   735
         Left            =   120
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Image           =   "MsgBx.frx":0054
      End
   End
   Begin VB.CommandButton Default 
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "MsgBx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Credits:
'           1) La Volpe for Alpha Image Control
'           2) Juan Carlos San Román Arias for jcFrames
'           3) Body_of_Rays who modified jcFrames and posted X jcFrame
'           4) Fred.cpp for isButton
'           5) stephane swertvaegher for basic concept used in his MBox code

' This is just an example of creating custom message box
' Please vote for above said persons who worked very hard to develop such great controls
' I just tried to show what can be done if u try to combine such great controls

'Bugs:
'       one know bug is that you can not set Default and Cancel Buttons as well as can't detect key pressed on the form
'       I tried to write code for both these things but its not working
'       If anybody knows any solution, please let me know
'       My Mail ID is divyeshparikh@gmail.com

Private Sub Cancel_Click()
    
    MButton_Click (1)
    
End Sub

Private Sub Default_Click()
    
    MButton_Click (0)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
            
        If Default.Enabled Then
            
            MButton_Click (0)
        
        End If
        
    ElseIf KeyCode = vbKeyEscape Then
        
        If Cancel.Enabled Then
            
            MButton_Click (1)
            
        End If
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    MsgBox KeyAscii
    
End Sub

Private Sub MButton_Click(Index As Integer)
                            
    mbxReturn = Index
    Unload Me
    
End Sub

