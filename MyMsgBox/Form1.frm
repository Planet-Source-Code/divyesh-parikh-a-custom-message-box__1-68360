VERSION 5.00
Object = "*\AMsgB.vbp"
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
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin MsgB.MyMsgBox MyMsgBox1 
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      Theme           =   4
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Label1.Caption = MyMsgBox1.ShowMessage("Trial", mbxOKOnly, mbxAlert, "Trial of Control")
    Label1.Caption = MyMsgBox1.ShowMessage("Another Trial", mbxEnterExitCancel, mbxQuestion, "Question with 3 buttons")
    
End Sub
