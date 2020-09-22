VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    MsgB "My Message" & vbCrLf & "Line2" & vbCrLf & "Line 3" & vbCrLf & "Line 4" & vbCrLf & "Line 5" & vbCrLf & "Line 6" & vbCrLf & "Line 7", mbxMessenger, mbxMoney, mbxYesNoCancel, mbxTrash, "My Title"
    MsgB "Trial Message", mbxGradient, mbxMetallic, mbxOKOnly, mbxAlert, "Trial Message"
    MsgB "Trial Message to chech if this message box changes its width according to length of message", mbxMessenger, mbxOlive, mbxPrintDontPrint, mbxPrint, "Print"
    
End Sub
