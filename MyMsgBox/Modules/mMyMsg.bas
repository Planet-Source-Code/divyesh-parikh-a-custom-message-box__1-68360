Attribute VB_Name = "mMyMsg"
Public mbxReturn%

Public Enum MBXButtons
    
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

Public Enum MBXStyle
    
    mbxGradient
    mbxWindows
    mbxMessenger
    
End Enum
    
Public Enum MBXTheme
    
    mbxblue
    mbxSilver
    mbxOlive
    mbxVisual2005
    mbxNorton2004
    mbxDarkBlue
    mbxGreen
    mbxOffice2003Style2
    mbxMetallic
    mbxOrange
    mbxTurquoise
    mbxGray
    mbxDarkBlue2
    mbxMoney
    mbxOffice2003Style1
    
End Enum

Public Enum MBXIcons

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

Public Function MsgB(ByVal Message As Variant, Optional Style As MBXStyle, Optional Theme As MBXTheme, Optional Buttons As MBXButtons, Optional Icons As MBXIcons, Optional Title As Variant) As Integer

    On Error Resume Next
    
    If IsMissing(Title) Then Title = App.Title
    
    With MsgBx
        
        .MBXFrame.Caption = Title
        
        .MBXFrame.Top = 0
        .MBXFrame.Left = 0
        .lblMsg.Caption = Message
        
        If .lblMsg.Width > 4015 Then
        
            .Width = 1080 + .lblMsg.Width + 200
            
        Else
            
            .Width = 5295
            
        End If
        
        .MBXFrame.Width = .Width
        
        If .lblMsg.Height > 315 Then
            
            .Height = 600 + .lblMsg.Height + 1500 '735
            
        Else
        
            .Height = 2415
            
        End If
        
        .MBXFrame.Height = .Height
        
    
    Select Case Style
        
        Case mbxGradient
            
            .MBXFrame.Style = jcGradient
            
        Case mbxMessenger
            
            .MBXFrame.Style = Messenger
            
        Case mbxWindows
            
            .MBXFrame.Style = Windows
            
    End Select
    
    Select Case Theme
        
        Case mbxblue
            
            .MBXFrame.ThemeColor = Blue
            
        Case mbxSilver
            
            .MBXFrame.ThemeColor = Silver
            
        Case mbxOlive
        
            .MBXFrame.ThemeColor = Olive
            
        Case mbxVisual2005
            
            .MBXFrame.ThemeColor = Visual2005
            
        Case mbxNorton2004
            
            .MBXFrame.ThemeColor = Norton2004
                
        Case mbxDarkBlue
            
            .MBXFrame.ThemeColor = xThemeDarkBlue
            
        Case mbxGreen
            
            .MBXFrame.ThemeColor = xThemeGreen
            
        Case mbxOffice2003Style2
            
            .MBXFrame.ThemeColor = xThemeOffice2003Style2
            
        Case mbxMetallic
            
            .MBXFrame.ThemeColor = xThemeMetallic
            
        Case mbxOrange
            
            .MBXFrame.ThemeColor = xThemeOrange
            
        Case mbxTurquoise
            
            .MBXFrame.ThemeColor = xThemeTurquoise
            
        Case mbxGray
            
            .MBXFrame.ThemeColor = xThemeGray
            
        Case mbxDarkBlue2
            
            .MBXFrame.ThemeColor = xThemeDarkBlue2
            
        Case mbxMoney
        
            .MBXFrame.ThemeColor = xThemeMoney
            
        Case mbxOffice2003Style1
            
            .MBXFrame.ThemeColor = xThemeOffice2003Style1
            
    End Select
    
    ButtonColor Theme, Style
    
    Select Case Buttons
        
        Case mbxOKOnly
            
            ButtonSetup (1)
            
            .MButton(0).Caption = "&OK"
            .MButton(0).Default = True
            
        Case mbxOkCancel
            
            ButtonSetup (2)
            
            .MButton(0).Caption = "&OK"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Cancel"
            .MButton(1).Cancel = True
        
        Case mbxAgreeOnly
            
            ButtonSetup (1)
            .MButton(0).Caption = "&Agree"
            .MButton(0).Default = True
            
        Case mbxYesNo
            
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Yes"
            .MButton(0).Default = True
            .MButton(1).Caption = "&No"
            .MButton(1).Cancel = True
        
        Case mbxYesNoCancel
        
            ButtonSetup (3)
            
            .MButton(0).Caption = "&Yes"
            .MButton(0).Default = True
            .MButton(1).Caption = "&No"
            .MButton(1).Cancel = True
            .MButton(2).Caption = "&Cancel"
            
        Case mbxQuitDontQuit
            
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Quit"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Dont Quit"
            .MButton(1).Cancel = True
        
        Case mbxExitDontExit
            
            ButtonSetup (2)
            
            .MButton(0).Caption = "E&xit"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Dont Exit"
            .MButton(1).Cancel = True
            
        Case mbxSaveDontSave
        
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Save"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Dont Save"
            .MButton(1).Cancel = True
            
        Case mbxLoadDontLoad
        
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Load"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Dont Load"
            .MButton(1).Cancel = True
            
        Case mbxPrintDontPrint
        
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Print"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Dont Print"
            .MButton(1).Cancel = True
            
        Case mbxEnterExit
        
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Enter"
            .MButton(0).Default = True
            .MButton(1).Caption = "E&xit"
            .MButton(1).Cancel = True
            
        Case mbxEnterExitCancel
        
            ButtonSetup (3)
            
            .MButton(0).Caption = "&Enter"
            .MButton(0).Default = True
            .MButton(1).Caption = "E&xit"
            .MButton(1).Cancel = True
            .MButton(2).Caption = "&Cancel"
            
        Case mbxAgreeDontAgree
        
            ButtonSetup (2)
            
            .MButton(0).Caption = "&Agree"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Dont Agree"
            .MButton(1).Cancel = True
            
        Case mbxRetryAbort
        
            ButtonSetup (2)
            
            .MButton(0).Caption = "&ReTry"
            .MButton(0).Default = True
            .MButton(1).Caption = "&Abort"
            .MButton(1).Cancel = True
            
        End Select
    
                                
        ShowIcons (Icons)
        
    End With
    
    MsgBx.Show 1
    
    MsgB = mbxReturn
    
End Function

Private Sub ShowIcons(mIcon As MBXIcons)
    
    With MsgBx
    
        Select Case mIcon
            
            Case mbxInfo
                
                .Logo.LoadImage_FromFile App.Path & "\Icons\info.png"
                
            Case mbxQuestion
                
                .Logo.LoadImage_FromFile App.Path & "\Icons\Question.png"
                
            Case mbxAlert
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Alert.png"
                
            Case mbxSave
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Save.png"
                
            Case mbxOpen
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Open1.png"
                
            Case mbxPrint
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Printer.png"
                
            Case mbxCritical
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Critical.png"
                
            Case mbxTrash
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Trash1.png"
                
            Case mbxForbidden
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Forbid.png"
                
            Case mbxSearch
            
                .Logo.LoadImage_FromFile App.Path & "\Icons\Search.png"
                
            Case mbxLock
                
                .Logo.LoadImage_FromFile App.Path & "\Icons\Lock.png"
                
        End Select
            
    End With
    
End Sub
Private Sub ButtonColor(mButtonStyle As MBXTheme, mStyle As MBXStyle)
    
    Dim i As Integer
    Dim mColor As Long
    
    Select Case mButtonStyle
        
        Case mbxblue
            
            mColor = RGB(129, 169, 226)
            
        Case mbxSilver
            
            mColor = RGB(153, 151, 180)
            
        Case mbxOlive
        
            mColor = RGB(181, 197, 143)
            
        Case mbxVisual2005
            
            mColor = RGB(194, 194, 171)
            
        Case mbxNorton2004
            
            mColor = RGB(217, 172, 1)
                
        Case mbxDarkBlue
            
            mColor = RGB(137, 170, 224)
            
        Case mbxGreen
            
            mColor = RGB(228, 235, 200)
            
        Case mbxOffice2003Style2
            
            mColor = RGB(249, 249, 255)
            
        Case mbxMetallic
            
            mColor = RGB(219, 220, 232)
            
        Case mbxOrange
            
            mColor = RGB(255, 122, 0)
            
        Case mbxTurquoise
            
            mColor = RGB(72, 209, 204)
            
        Case mbxGray
            
            mColor = RGB(192, 192, 192)
            
        Case mbxDarkBlue2
            
            mColor = RGB(81, 128, 208)
            
        Case mbxMoney
        
            mColor = RGB(160, 160, 160)
            
        Case mbxOffice2003Style1
            
            mColor = RGB(209, 227, 251)
            
    End Select
    
    With MsgBx
    
        For i = 0 To 2
        
            .MButton(i).BackColor = mColor
            .MButton(i).HighlightColor = mColor
            
        Next
        
        If mStyle = mbxWindows Then
            
            .MBXFrame.TextBoxColor = mColor
        
        End If
        
    End With
    
End Sub

Private Sub ButtonSetup(mButtonCount As Integer)
    
    Dim i As Integer
    
    With MsgBx
        
        .Default.Enabled = False
        .Cancel.Enabled = False
        
        For i = 0 To 2
            
            .MButton(i).Visible = False
            .MButton(i).Default = False
            .MButton(i).Cancel = False
            
        Next
        
        For i = 0 To mButtonCount - 1
            
            .MButton(i).Visible = True
            .MButton(i).Top = .Height - 735
            
        Next
        
        Select Case mButtonCount
            
            Case 1
                
                .MButton(0).Left = .MBXFrame.Width / 2 - .MButton(0).Width / 2
                .Default.Enabled = True
                
            Case 2
                
                .MButton(0).Left = .MBXFrame.Width / 4 - .MButton(0).Width / 2
                .MButton(1).Left = .MButton(0).Left + (.MBXFrame.Width / 2)
                .Default.Enabled = True
                .Cancel.Enabled = True
                
            Case 3
                
                .MButton(0).Left = .MBXFrame.Width / 6 - .MButton(0).Width / 2
                .MButton(1).Left = .MButton(0).Left + (.MBXFrame.Width / 3)
                .MButton(2).Left = .MButton(1).Left + (.MBXFrame.Width / 3)
                .Default.Enabled = True
                .Cancel.Enabled = True
                
        End Select
        
    End With
        
End Sub
