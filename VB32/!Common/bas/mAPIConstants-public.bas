Attribute VB_Name = "mAPIConstants"
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


Public Enum enumBorderFlags
    BF_ADJUST = &H2000
    BF_BOTTOM = &H8
    BF_DIAGONAL = &H10
    BF_FLAT = &H4000
    BF_LEFT = &H1
    BF_MIDDLE = &H800
    BF_MONO = &H8000
    BF_RIGHT = &H4
    BF_SOFT = &H1000
    BF_TOP = &H2
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
End Enum

Public Enum enumBorderEdges
    BDR_RAISEDINNER = &H4
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENINNER = &H8
    BDR_SUNKENOUTER = &H2
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum


Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Declare Function DrawEdge _
                                    Lib "user32" ( _
                                ByVal hdc As Long, _
                                qrc As Rect, _
                                ByVal edge As enumBorderEdges, _
                                ByVal grfFlags As enumBorderFlags) _
                            As Long


' List of different styles of keyboard entry allowed.
' Goes with the function ctlKeyPress()
Public Enum enumKeyPressAllowTypes
    NumbersOnly = 2 ^ 0
    Uppercase = 2 ^ 1
    NoSpaces = 2 ^ 2
    NoSingleQuotes = 2 ^ 3
    NoDoubleQuotes = 2 ^ 4
    AllowDecimal = 2 ^ 5
    AllowNegative = 2 ^ 6
    DatesOnly = 2 ^ 7
    TimesOnly = 2 ^ 8
    LettersOnly = 2 ^ 9
    AllowSpaces = 2 ^ 10
    AllowStars = 2 ^ 11
    AllowPounds = 2 ^ 12
End Enum


