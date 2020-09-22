VERSION 5.00
Begin VB.UserControl MyMsgBox 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   ScaleHeight     =   690
   ScaleWidth      =   690
   Begin VB.Image Image1 
      Height          =   690
      Left            =   0
      Picture         =   "MyMsgBox.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "MyMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum mbxStyleConst
    
    mbxGradient = 0
    mbxWindows = 1
    mbxMessenger = 2
    
End Enum

Public Enum mbxThemeConst
    
    mbxblue = 0
    mbxSilver = 1
    mbxOlive = 2
    mbxVisual2005 = 3
    mbxNorton2004 = 4
    mbxDarkBlue = 5
    mbxGreen = 6
    mbxOffice2003Style2 = 7
    mbxMetallic = 8
    mbxOrange = 9
    mbxTurquoise = 10
    mbxGray = 11
    mbxDarkBlue2 = 12
    mbxMoney = 13
    mbxOffice2003Style1 = 14
    
End Enum

Public Enum mbxButtonsConst
    
    mbxOKOnly
    mbxAgreeOnly
    mbxOkCancel
    mbxYesNo
    mbxYesNoCancel
    mbxQuitDontQuit
    mbxExitDontExit
    mbxSaveDontSave
    mbxLoadDontLoad
    mbxPrintDontPrint
    mbxEnterExit
    mbxEnterExitCancel
    mbxAgreeDontAgree
    mbxRetryAbort
    
End Enum

Public Enum mbxIconsConst

    mbxInfo
    mbxQuestion
    mbxAlert
    mbxSave
    mbxOpen
    mbxPrint
    mbxCritical
    mbxTrash
    mbxForbidden
    mbxSearch
    mbxLock
    
End Enum

Private m_style As mbxStyleConst
Private m_Theme As mbxThemeConst
Private m_TextColor As OLE_COLOR
Private m_FontName As String
Private m_FontSize As Integer
Private m_FontBold As Boolean
Private m_FontItalics As Boolean
Private m_FontUnderline As Boolean
Private m_FontStrikeThrough As Boolean

Public mbxRetVal%

Private Sub UserControl_Initialize()
    
    m_style = mbxMessenger
    m_Theme = mbxblue
    m_TextColor = vbBlack
    m_FontName = "Verdana"
    m_FontSize = 10
    m_FontBold = False
    m_FontItalics = False
    m_FontUnderline = False
    m_FontStrikeThrough = False
    
End Sub

Private Sub UserControl_InitProperties()
    
    m_style = mbxMessenger
    m_Theme = mbxblue
    m_TextColor = vbBlack
    m_FontName = "Verdana"
    m_FontSize = 10
    m_FontBold = False
    m_FontItalics = False
    m_FontUnderline = False
    m_FontStrikeThrough = False
    
End Sub

Public Property Let Style(ByRef New_Style As mbxStyleConst)
    
    m_style = New_Style
    PropertyChanged "Style"
    
End Property

Public Property Get Style() As mbxStyleConst
    
    Style = m_style
    
End Property

Public Property Let Theme(ByRef New_Theme As mbxThemeConst)

    m_Theme = New_Theme
    PropertyChanged "Theme"

End Property

Public Property Get Theme() As mbxThemeConst

    Theme = m_Theme
    
End Property

Public Property Let FontName(ByRef New_FontName As String)
    
    m_FontName = New_FontName
    PropertyChanged "FontName"
    
End Property

Public Property Get FontName() As String
    
    FontName = m_FontName
    
End Property

Public Property Let FontSize(ByRef New_FontSize As Integer)
    
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
    
End Property

Public Property Get FontSize() As Integer

    FontSize = m_FontSize
    
End Property

Public Property Let FontBold(ByRef New_FontBold As Boolean)
    
    m_FontBold = New_FontBold
    PropertyChanged "FontBold"
    
End Property

Public Property Get FontBold() As Boolean
    
    FontBold = m_FontBold
    
End Property

Public Property Let FontItalics(ByRef New_FontItalics As Boolean)
    
    m_FontItalics = New_FontItalics
    PropertyChanged "FontItalics"
    
End Property

Public Property Get FontItalics() As Boolean
    
    FontItalics = m_FontItalics
    
End Property

Public Property Let FontUnderline(ByRef New_FontUnderline As Boolean)
    
    m_FontUnderline = New_FontUnderline
    PropertyChanged "FontUnderline"
    
End Property

Public Property Get FontUnderline() As Boolean
    
    FontUnderline = m_FontUnderline
    
End Property

Public Property Let FontStrikeThru(ByRef New_FontStrikeThru As Boolean)
    
    m_FontStrikeThrough = New_FontStrikeThru
    PropertyChanged "FontStrikeThru"
    
End Property

Public Property Get FontStrikeThru() As Boolean
    
    FontStrikeThru = m_FontStrikeThrough
    
End Property

Public Property Let TextColor(ByRef New_Color As OLE_COLOR)
    
    m_TextColor = New_Color
    PropertyChanged "TextColor"
    
End Property

Public Property Get TextColor() As OLE_COLOR
    
    TextColor = m_TextColor
    
End Property

Public Function ShowMessage(ByVal Message As Variant, Optional Buttons As mbxButtonsConst, Optional Icons As mbxIconsConst, Optional Title As Variant)
    
    ShowMessage = mMyMsg.MsgB(Message, m_style, m_Theme, Buttons, Icons, Title)
    
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        
        m_style = .ReadProperty("Style", mbxMessenger)
        m_Theme = .ReadProperty("Theme", mbxblue)
        m_FontName = .ReadProperty("FontName", "Verdana")
        m_FontSize = .ReadProperty("FontSize", 10)
        m_FontBold = .ReadProperty("FontBold", False)
        m_FontItalics = .ReadProperty("FontItalics", False)
        m_FontUnderline = .ReadProperty("FontUnderline", False)
        m_FontStrikeThrough = .ReadProperty("FontStrikeThru", False)
        m_TextColor = .ReadProperty("TextColor", vbBlack)
        
    End With
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        
        .WriteProperty "Style", m_style, mbxMessenger
        .WriteProperty "Theme", m_Theme, mbxblue
        .WriteProperty "FontName", m_FontName, "Verdana"
        .WriteProperty "FontSize", m_FontSize, 10
        .WriteProperty "FontBold", m_FontBold, False
        .WriteProperty "FontItalics", m_FontItalics, False
        .WriteProperty "FontUnderline", m_FontUnderline, False
        .WriteProperty "FontStrikeThru", m_FontStrikeThrough, False
        .WriteProperty "TextColor", m_TextColor, vbBlack
        
    End With
    
End Sub
