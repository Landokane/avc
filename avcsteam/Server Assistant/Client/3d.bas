
Private Sub CalcAng(x, y, Z, X2, Y2, Z2, center)
'Takes X, Y, Z as input, then calculates the X,Y,Z output after angles have been implemented
    X1 = x - center
    Y1 = y - center
    Z1 = Z
    
    'find existing angle
    If Y1 <> 0 Then m = Z1 / Y1
    a2 = Atn(m) * (180 / 3.14159)
    
    If Y1 >= 0 Then a2 = a2 + 180
    
    a2 = a2 + VAngle
    Do Until a2 <= 360: a2 = a2 - 360: Loop
    a1 = a2 * (3.14159 / 180)
    
    Y3 = Cos(a1) * Sqr(Z1 ^ 2 + Y1 ^ 2)  ' Cos * Distance from center
    Z2 = Sin(a1) * Sqr(Z1 ^ 2 + Y1 ^ 2)  ' Sin * Distance from center
        
    'find existing angle
    If X1 <> 0 Then m = Y3 / X1
    a2 = Atn(m) * (180 / 3.14159)
    
    If X1 >= 0 Then a2 = a2 + 180
    
    a2 = a2 + Angle
    Do Until a2 <= 360: a2 = a2 - 360: Loop
    a1 = a2 * (3.14159 / 180)
    
    X2 = Cos(a1) * Sqr(X1 ^ 2 + Y3 ^ 2)  ' Cos * Distance from center
    Y2 = Sin(a1) * Sqr(X1 ^ 2 + Y3 ^ 2)  ' Sin * Distance from center
    
End Sub
Private Sub CalcRot()
'Draws the 3d view
EnDis False

Dim MapT(0 To 257, 0 To 257) As Integer
MapView.DrawWidth = ScrollDrawWidth.Value

center = (Con.MapSize / 2) + 0.2
MapView.Cls

q = ScrollRes.Value
ProgressBar1.Max = ((Con.MapSize + q) ^ 2)

'Now find the view dir and decide which way to draw
'to ensure constant back-to-front drawing
If Angle >= 0 And Angle < 45 Then DrawMode = 1
If Angle >= 45 And Angle < 135 Then DrawMode = 2
If Angle >= 135 And Angle < 225 Then DrawMode = 3
If Angle >= 225 And Angle < 315 Then DrawMode = 4
If Angle >= 315 And Angle <= 360 Then DrawMode = 5

If DrawMode = 1 Then q = -q: tpe = 1
If DrawMode = 2 Then q = -q: tpe = 2
If DrawMode = 3 Then q = q: tpe = 1
If DrawMode = 4 Then q = q: tpe = 2
If DrawMode = 5 Then q = -q: tpe = 1

If VAngle <= 90 Or VAngle > 270 Then q = -q: Debug.Print q ': tpe = 2
'If tpe = 2 And VAngle >= 0 And VAngle < 180 Then q = -q: tpe = 1

If tpe = 1 Then

    If q > 0 Then
        For y = 1 To Con.MapSize Step q '* 2
        For x = 1 To Con.MapSize Step q '* 2
    
            If DrawLines.Value = True Then DrawHer x, y, q, center
            If DrawDots.Value = True Then DrawHerDots x, y, q, center
    
        Next x
        MapView.Refresh
        DoEvents
        If StopVal = 1 Then StopVal = 0: Exit Sub
        ProgressBar1.Value = x * y
        Next y
    End If

    If q < 0 Then
        For y = Con.MapSize To 1 Step q '* 2
        For x = Con.MapSize To 1 Step q '* 2
    
            If DrawLines.Value = True Then DrawHer x, y, q, center
            If DrawDots.Value = True Then DrawHerDots x, y, q, center
    
        Next x
        If x < 0 Then x = -x
        If y < 0 Then y = -y
        MapView.Refresh
        DoEvents
        If StopVal = 1 Then StopVal = 0: Exit Sub
        
        ProgressBar1.Value = (Con.MapSize - x) * (Con.MapSize - y)
        Next y
    End If

End If

If tpe = 2 Then
    If q > 0 Then
        For x = 1 To Con.MapSize Step q '* 2
        For y = 1 To Con.MapSize Step q '* 2
    
            If DrawLines.Value = True Then DrawHer x, y, q, center
            If DrawDots.Value = True Then DrawHerDots x, y, q, center
    
        Next y
        MapView.Refresh
        DoEvents
        If StopVal = 1 Then StopVal = 0: Exit Sub
        
        ProgressBar1.Value = x * y
        Next x
    End If

    If q < 0 Then
        For x = Con.MapSize To 1 Step q '* 2
        For y = Con.MapSize To 1 Step q '* 2
    
            If DrawLines.Value = True Then DrawHer x, y, q, center
            If DrawDots.Value = True Then DrawHerDots x, y, q, center
    
        Next y
        
        If x < 0 Then x = -x
        If y < 0 Then y = -y
        MapView.Refresh
        DoEvents
        If StopVal = 1 Then StopVal = 0: Exit Sub
        
        ProgressBar1.Value = (Con.MapSize - x) * (Con.MapSize - y)
        Next x
    End If
End If

MapView.DrawWidth = 1
EnDis True
End Sub