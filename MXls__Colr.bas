Attribute VB_Name = "MXls__Colr"
Option Compare Database
Option Explicit
Private Const ColrLines_1$ = "ActiveBorder -4934476" & _
vbCrLf & "ActiveCaption -6703919" & _
vbCrLf & "ActiveCaptionText -16777216" & _
vbCrLf & "AliceBlue -984833" & _
vbCrLf & "AntiqueWhite -332841" & _
vbCrLf & "AppWorkspace -5526613" & _
vbCrLf & "Aqua -16711681" & _
vbCrLf & "Aquamarine -8388652" & _
vbCrLf & "Azure -983041" & _
vbCrLf & "Beige -657956" & _
vbCrLf & "Bisque -6972" & _
vbCrLf & "Black -16777216" & _
vbCrLf & "BlanchedAlmond -5171" & _
vbCrLf & "Blue -16776961" & _
vbCrLf & "BlueViolet -7722014" & _
vbCrLf & "Brown -5952982" & _
vbCrLf & "BurlyWood -2180985" & _
vbCrLf & "ButtonFace -986896" & _
vbCrLf & "ButtonHighlight -1" & _
vbCrLf & "ButtonShadow -6250336"
Private Const ColrLines_2$ = "CadetBlue -10510688" & _
vbCrLf & "Chartreuse -8388864" & _
vbCrLf & "Chocolate -2987746" & _
vbCrLf & "Control -986896" & _
vbCrLf & "ControlDark -6250336" & _
vbCrLf & "ControlDarkDark -9868951" & _
vbCrLf & "ControlLight -1842205" & _
vbCrLf & "ControlLightLight -1" & _
vbCrLf & "ControlText -16777216" & _
vbCrLf & "Coral -32944" & _
vbCrLf & "CornflowerBlue -10185235" & _
vbCrLf & "Cornsilk -1828" & _
vbCrLf & "Crimson -2354116" & _
vbCrLf & "Cyan -16711681" & _
vbCrLf & "DarkBlue -16777077" & _
vbCrLf & "DarkCyan -16741493" & _
vbCrLf & "DarkGoldenrod -4684277" & _
vbCrLf & "DarkGray -5658199" & _
vbCrLf & "DarkGreen -16751616" & _
vbCrLf & "DarkKhaki -4343957"
Private Const ColrLines_3$ = "DarkMagenta -7667573" & _
vbCrLf & "DarkOliveGreen -11179217" & _
vbCrLf & "DarkOrange -29696" & _
vbCrLf & "DarkOrchid -6737204" & _
vbCrLf & "DarkRed -7667712" & _
vbCrLf & "DarkSalmon -1468806" & _
vbCrLf & "DarkSeaGreen -7357301" & _
vbCrLf & "DarkSlateBlue -12042869" & _
vbCrLf & "DarkSlateGray -13676721" & _
vbCrLf & "DarkTurquoise -16724271" & _
vbCrLf & "DarkViolet -7077677" & _
vbCrLf & "DeepPink -60269" & _
vbCrLf & "DeepSkyBlue -16728065" & _
vbCrLf & "Desktop -16777216" & _
vbCrLf & "DimGray -9868951" & _
vbCrLf & "DodgerBlue -14774017" & _
vbCrLf & "Firebrick -5103070" & _
vbCrLf & "FloralWhite -1296" & _
vbCrLf & "ForestGreen -14513374" & _
vbCrLf & "Fuchsia -65281"
Private Const ColrLines_4$ = "Gainsboro -2302756" & _
vbCrLf & "GhostWhite -460545" & _
vbCrLf & "Gold -10496" & _
vbCrLf & "Goldenrod -2448096" & _
vbCrLf & "GradientActiveCaption -4599318" & _
vbCrLf & "GradientInactiveCaption -2628366" & _
vbCrLf & "Gray -8355712" & _
vbCrLf & "GrayText -9605779" & _
vbCrLf & "Green -16744448" & _
vbCrLf & "GreenYellow -5374161" & _
vbCrLf & "Highlight -16746281" & _
vbCrLf & "HighlightText -1" & _
vbCrLf & "Honeydew -983056" & _
vbCrLf & "HotPink -38476" & _
vbCrLf & "HotTrack -16750900" & _
vbCrLf & "InactiveBorder -722948" & _
vbCrLf & "InactiveCaption -4207141" & _
vbCrLf & "InactiveCaptionText -16777216" & _
vbCrLf & "IndianRed -3318692" & _
vbCrLf & "Indigo -11861886"
Private Const ColrLines_5$ = "Info -31" & _
vbCrLf & "InfoText -16777216" & _
vbCrLf & "Ivory -16" & _
vbCrLf & "Khaki -989556" & _
vbCrLf & "Lavender -1644806" & _
vbCrLf & "LavenderBlush -3851" & _
vbCrLf & "LawnGreen -8586240" & _
vbCrLf & "LemonChiffon -1331" & _
vbCrLf & "LightBlue -5383962" & _
vbCrLf & "LightCoral -1015680" & _
vbCrLf & "LightCyan -2031617" & _
vbCrLf & "LightGoldenrodYellow -329006" & _
vbCrLf & "LightGray -2894893" & _
vbCrLf & "LightGreen -7278960" & _
vbCrLf & "LightPink -18751" & _
vbCrLf & "LightSalmon -24454" & _
vbCrLf & "LightSeaGreen -14634326" & _
vbCrLf & "LightSkyBlue -7876870" & _
vbCrLf & "LightSlateGray -8943463" & _
vbCrLf & "LightSteelBlue -5192482"
Private Const ColrLines_6$ = "LightYellow -32" & _
vbCrLf & "Lime -16711936" & _
vbCrLf & "LimeGreen -13447886" & _
vbCrLf & "Linen -331546" & _
vbCrLf & "Magenta -65281" & _
vbCrLf & "Maroon -8388608" & _
vbCrLf & "MediumAquamarine -10039894" & _
vbCrLf & "MediumBlue -16777011" & _
vbCrLf & "MediumOrchid -4565549" & _
vbCrLf & "MediumPurple -7114533" & _
vbCrLf & "MediumSeaGreen -12799119" & _
vbCrLf & "MediumSlateBlue -8689426" & _
vbCrLf & "MediumSpringGreen -16713062" & _
vbCrLf & "MediumTurquoise -12004916" & _
vbCrLf & "MediumVioletRed -3730043" & _
vbCrLf & "Menu -986896" & _
vbCrLf & "MenuBar -986896" & _
vbCrLf & "MenuHighlight -13395457" & _
vbCrLf & "MenuText -16777216" & _
vbCrLf & "MidnightBlue -15132304"
Private Const ColrLines_7$ = "MintCream -655366" & _
vbCrLf & "MistyRose -6943" & _
vbCrLf & "Moccasin -6987" & _
vbCrLf & "NavajoWhite -8531" & _
vbCrLf & "Navy -16777088" & _
vbCrLf & "OldLace -133658" & _
vbCrLf & "Olive -8355840" & _
vbCrLf & "OliveDrab -9728477" & _
vbCrLf & "Orange -23296" & _
vbCrLf & "OrangeRed -47872" & _
vbCrLf & "Orchid -2461482" & _
vbCrLf & "PaleGoldenrod -1120086" & _
vbCrLf & "PaleGreen -6751336" & _
vbCrLf & "PaleTurquoise -5247250" & _
vbCrLf & "PaleVioletRed -2396013" & _
vbCrLf & "PapayaWhip -4139" & _
vbCrLf & "PeachPuff -9543" & _
vbCrLf & "Peru -3308225" & _
vbCrLf & "Pink -16181" & _
vbCrLf & "Plum -2252579"
Private Const ColrLines_8$ = "PowderBlue -5185306" & _
vbCrLf & "Purple -8388480" & _
vbCrLf & "Red -65536" & _
vbCrLf & "RosyBrown -4419697" & _
vbCrLf & "RoyalBlue -12490271" & _
vbCrLf & "SaddleBrown -7650029" & _
vbCrLf & "Salmon -360334" & _
vbCrLf & "SandyBrown -744352" & _
vbCrLf & "ScrollBar -3618616" & _
vbCrLf & "SeaGreen -13726889" & _
vbCrLf & "SeaShell -2578" & _
vbCrLf & "Sienna -6270419" & _
vbCrLf & "Silver -4144960" & _
vbCrLf & "SkyBlue -7876885" & _
vbCrLf & "SlateBlue -9807155" & _
vbCrLf & "SlateGray -9404272" & _
vbCrLf & "Snow -1286" & _
vbCrLf & "SpringGreen -16711809" & _
vbCrLf & "SteelBlue -12156236" & _
vbCrLf & "Tan -2968436"
Private Const ColrLines_9$ = "Teal -16744320" & _
vbCrLf & "Thistle -2572328" & _
vbCrLf & "Tomato -40121" & _
vbCrLf & "Transparent 16777215" & _
vbCrLf & "Turquoise -12525360" & _
vbCrLf & "Violet -1146130" & _
vbCrLf & "Wheat -663885" & _
vbCrLf & "White -1" & _
vbCrLf & "WhiteSmoke -657931" & _
vbCrLf & "Window -1" & _
vbCrLf & "WindowFrame -10197916" & _
vbCrLf & "WindowText -16777216" & _
vbCrLf & "Yellow -256" & _
vbCrLf & "YellowGreen -6632142"
Const ColrLines$ = ColrLines_1 & vbCrLf & ColrLines_2 & vbCrLf & ColrLines_3 & vbCrLf & ColrLines_4 & vbCrLf & ColrLines_5 & vbCrLf & ColrLines_6 & vbCrLf & ColrLines_7 & vbCrLf & ColrLines_8 & vbCrLf & ColrLines_9

Function ColrLy() As String()
ColrLy = SplitCrLf(ColrLines)
End Function

Function ColrSq() As Variant()
Dim J%, O(), Ly$(), Nm$, Colr&
Ly = ColrLy
ReDim O(1 To Sz(Ly), 1 To 2)
For J = 1 To Sz(Ly)
    LinTRstAsg Ly(J - 1), Nm, Colr
    O(J, 1) = Nm
    O(J, 2) = Colr
Next
ColrSq = O
End Function
Function CvColr(A)
If CanCvLng(A) Then CvColr = CLng(A): Exit Function
CvColr = ColrStrColr(A)
End Function
Function ColrStr$(A&)
Dim L
For Each L In ColrLy
    With Brk(L, " ")
        If .S2 = A Then ColrStr = .S1: Exit Function
    End With
Next
End Function

Function ColrStrColr(A)
Dim X$
X = AyFstRmvT1(ColrLy, A)
If X = "" Then Exit Function
ColrStrColr = CLng(X)
End Function

Function ColrWb() As Workbook
Dim Ws As Worksheet, Sq(), J%
Sq = ColrSq
Set Ws = RgWs(SqRg(ColrSq, NewA1))
For J = 1 To UBound(Sq, 1)
    WsRC(Ws, J, 3).Interior.Color = Sq(J, 2)
Next
WsCC(Ws, 1, 2).EntireColumn.AutoFit
Set ColrWb = WsWb(WsVis(Ws))
End Function

