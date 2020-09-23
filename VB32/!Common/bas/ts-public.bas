Attribute VB_Name = "ts"
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This code was written by The Frog Prince
'
' If you have questions or comments, I can be reached at
'        TheFrogPrince@hotmail.com
' If you wanna see more cool vb user controls, classes, code,
' and add-ins like this one, or updates to this code, go to
' my web page at
'        http://members.tripod.com/the__frog__prince/
' You are free to use, re-write, or otherwise do as you wish
' with this code.  However, if you do a cool enhancement, I
' would appreciate it if you could e-mail it to me.  I like
' to see what people do with my stuff.  =)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


Option Explicit

Public Function rectMake(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long) As Rect
    Dim tRet As Rect
    tRet.Bottom = lBottom
    tRet.Top = lTop
    tRet.Left = lLeft
    tRet.Right = lRight
    rectMake = tRet
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlKeyPress
'    This function is handy for wrapping input to textboxes
'    or other controls that have the KeyPress event to implement
'    standard types of input masks.
'       Example:
'            Private Sub txtPlaceOfEmployment_KeyPress(KeyAscii As Integer)
'                KeyAscii = ts.wrapKeyPress(KeyAscii, Uppercase + NoDoubleQuotes)
'            End Sub
Public Function ctlKeyPress(ByVal KeyAscii As KeyCodeConstants, ByVal TypeToAllow As enumKeyPressAllowTypes) As Integer
    
    Dim ltrKeyAscii As Integer
    ltrKeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    ' By default pass the keystroke through and then optionally kill it
    ctlKeyPress = KeyAscii
    
    ' Default Keystrokes to allow (enter, backspace, delete, escape)
    If _
        KeyAscii = vbKeyReturn Or _
        KeyAscii = vbKeyEscape Or _
        KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Then
        
        Exit Function
    End If
    
    ' NumbersOnly
    If (TypeToAllow And NumbersOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case (KeyAscii = vbKeySubtract Or KeyAscii = Asc("-")) And (TypeToAllow And AllowNegative)
            Case KeyAscii = Asc("#") And (TypeToAllow And AllowPounds)
            Case KeyAscii = Asc("*") And (TypeToAllow And AllowStars)
            Case KeyAscii = vbKeyDecimal And (TypeToAllow And AllowDecimal)
            Case KeyAscii = vbKeySpace And (TypeToAllow And AllowSpaces)
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' DatesOnly
    If (TypeToAllow And DatesOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = vbKeyDivide Or KeyAscii = Asc("/")
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' TimesOnly
    If (TypeToAllow And TimesOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = Asc(":") Or KeyAscii = Asc(";")
                ctlKeyPress = Asc(":")
            Case ltrKeyAscii = vbKeyA Or ltrKeyAscii = vbKeyP Or ltrKeyAscii = vbKeyM
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' LettersOnly
    If (TypeToAllow And LettersOnly) Then
        Select Case True
            Case ltrKeyAscii >= vbKeyA And ltrKeyAscii <= vbKeyZ
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' UpperCase
    If (TypeToAllow And Uppercase) Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    ' No Spaces
    If (TypeToAllow And NoSpaces) And KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If
    
    ' No Double Quotes
    If (TypeToAllow And NoDoubleQuotes) And KeyAscii = Asc("""") Then
        KeyAscii = Asc("'")
    End If
    
    ' No Single Quotes
    If (TypeToAllow And NoSingleQuotes) And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    
    ctlKeyPress = KeyAscii
    
End Function

