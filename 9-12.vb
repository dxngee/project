    Next q2
    
    If (q + tY1 + 1) > (y + cx + d5 / d) Then GoTo s_g
    sp1 = w - 1: ep1 = w + tX1 + 1
    sp2 = e - 1: ep2 = e + tZ1 + 1
    If sp1 < 1 Then sp1 = 1
    If sp2 < 1 Then sp2 = 1
    If ep1 > (x + X1) Then ep1 = x + X1
    If e2 > z Then ep2 = tZ1cs = 0: tNum = 0
    For w2 = sp1 to ep1
    If j(q + tY1 + 1, w2, sp2).rNumber <> 0 And tNum <> 0 Then tNum = j(q + tY1 + 1, w2, sp2).rNumber: cs = 1
    If tNum <> j(q + tY1 + 1, w2, sp2).rNumber Then
        If cs = 2 Then
            tX1 = w2 - sp1 - 1: GoTo tZ1
        End If
        tNum = j(q + tY1 + 1, w2, sp2).rNumber: cs = cs + 1
    End If
    Next w2

z1:
    cs = 0: tNum = 0
    For r2 = sp2 To ep2
    If j(q + tY1 + 1, sp1, r2).rNumber <> 0 And tNum <> 0 Then tNum = j(q + tY1 + 1, sp1, r2).rNumber: cs = 1
    If tNum <> j(q + tY1 + 1, sp1, r2).rNumber Then
        If cs = 2 Then
            tZ1 = r2 - sp2 - 1: GoTo z2
        End If
        tNum = j(q + tY1 + 1, sp1, r2).rNumber: cs = cs + 1
    End If
    Next r2

z2:
    sp1 = w - 1: ep1 = w + tX1 + 1
    sp2 = e - 1: ep2 = e + tZ1 + 1
    If sp1 < 1 Then sp1 = 1
    If sp2 < 1 Then sp2 = 1
    If ep1 > (x + X1) Then ep1 = x + X1
    If ep2 > z Then ep2 = z1
    For w2 = sp1 To ep1
    For r2 = sp2 To ep2
    If j(q + tY1 + 1, w2, r2).rNumber <> 0 Then
        q1 = q + tY1 + 1 - j(q + tY1 + 1, w2, r2).rY: w1 = w2 - j(q + tY1 + 1, w2, r2).rX: r1 = r2 - j(q + tY1 + 1, w2, r2).rNumber
        GoSub f_sz
        If tX1 >= tX And tZ1 >= tZ And tY1 >= tY Then
            cs = 0
            q3 = q1: w3 = w1: r3 = r1
            tNum3 = tNum
            GoTo cnt_p
        End If

    End If
    Next r2: Next w2
    If j(q + tY1 + 1, w, e).rNumber = 0 Then GoTo s_g
    q1 = q + tY1 + 1 - j(q + tY1 + 1, w, e).rY: w1 = w - j(q + tY1 + 1, w, e).rX: r1 = e - j(q+ tY1 + 1, w, e).rZ
    If w <> wl Or e <> rl Then GoTo s_g 
    tNum = j(q + tYl + 1, w, e).rNumber 
    GoSub f_sz
    If tXl >= tX And tZ1l >= tZ Then 
        'ep = tYl 
        GoTo q_l 
    Else 
        GoTo s_g 
    End If 
    cnt_p: 
        cs = cs + l 
        If cs = 10 Then 
        ql = q3: wl = w3: rl = r3: tNum = tNum3 
        GoTo q_2 
    End If
    v = Rnd(1)
    If v <= a3 Then xs = -1: zs = 1: dX = 0: dZ = tZ1
    If v > a3 And v <= (a3 + a2) Then xs = -1: zs = 0: dX = 0: dZ = 0
    If v > (a3 + a2) And v <= (2 * a3 + a2) Then xs = -1: zs = -1: dX = 0: dZ = 0
    If v > (2 * a3 + a2) And v <= (2 * a3 + 2 * a2) Then xs = 0: zs = 1: dX = 0: dZ = tZ1
    If v > (2 * a3 + 2 * a2) And v <= (a1 + 2 * a3 + 2 * a2) Then xs = 0: zs = 0: dX = 0: dZ = 0
    If v > (a1 + 2 * a3 + 2 * a2) And v <= (a1 + 2 * a3 + 3 * a2) Then xs = 0: zs = -1: dX = 0: dZ = 0
    If v > (a1 + 2 * a3 + 3 * a2) And v <= (a1 + 3 * a3 + 3 * a2) Then xs = 1: zs = 1: dX = tX1: dZ = tZ1
    If v > (a1 + 3 * a3 + 3 * a2) And v <= (a1 + 3 * a3 + 4 * a2) Then xs = 1: zs = 0: dX = tX1: dZ = 0
    If v > (a1 + 3 * a3 + 4 * a2) And v <= 1 Then xs = 1: zs = -1: dX = tX1: dZ = 0
    If w = 1 And xs = -1 Then xs = 1: dX = tX1
    If (w + dX + xs) > (x + X1) Then xs = -1: dX = 0
    If e = 1 And zs = -1 Then zs = 1: dZ = tZ1
    If (e + dZ + zs) > z Then zs = -1: dZ = 0
    If j(q + tY1 + 1, w + dX + xs, e + dZ + zs).rNumber = 0 Then GoTo cnt_p
    q1 = q + tY1 + 1 - j(q + tY1 + 1, w + dX + xs, e + dZ + zs).rY: w1 = w + dx + xs - j(q + tY1 + 1, w + dX + xs, e + dZ + zs).rX: r1 = e + dZ + zs - j(q + tY1 + 1, w + dX + xs, e + dZ + zs).rZ
    tNum = j(q + tY1 + 1, w + dX + xs, e + dZ + zs).rNumber
q_2: 
    'Debug.Assert Not (j(q1, w1, r1).rNumber <> tNum)
    If j(q1, w1, r1).rNumber <> tNum Then GoTo s_g
    GoSub f_sz
    If tX1 < tX Or tZ1 < tZ Or tY1 < tY Then GoTo cnt_p
For q2 = 0 To tY
    For w2 = 0 to tX
    For r2 = 0 To tZ
    j(q + q2, w + w2, e + r2) = j(q1 + q2, w1 + w2, r1 + r2)
    j(q1 + q2, w1 + w2, r1 + r2).rNumber = 0
    Next r2: Next w2: Next a2
End If
s_g:
    Next e: Next w: Next q

    'заполнение кровли
    p_zvs:
    For q = y + cx + 1 To y + cx + d5 / d Step d5 / d
    For w = 1 To x + X1 Step d5 / d
    For e = 1 To z Step d5 / d
    If j(q, w, e).rNumber = 0 Then
        q1 = q: w1 = w: r1 = e: tNum = 0: GoSub f_sz
        If (tY + 1) >= d5 / d Then e1 = d5 / d Else ep1 = tY + 1
        If (tX + 1) >= d5 / d Then ep2 = d5 / d Else ep2 = tX + 1
        If (tZ + 1) >= d5 / d Then ep3 = d5 / d Else ep3 = tZ + 1
        For q1 = 0 To ep1 - 1: For w1 = 0 To ep2 - 1: For r1 = 0 To ep3 - 1
        If (q + q1) <= (y + cx + d5 / d) And (w + w1) <= x And (e + r1) <= z Then j(q + q1, w + w1, e + r1).rType = t2: j(q + q1, w + w1, e + r1).rNumber = n:
        j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1: j(q + q1, w + w1, e + r1).rZ = r1
        End If
        Next r1: Next w1: Next q1
        n = n + 1
    End If
    Next e: Next w: Next q
    'текущее состояние разреза
    If Me.WindowState <> vbMinimized Then GoSub razr
    If p_flg = True Then
    'фигура выпуска
        For q = 1 To y + cx
        For w = 1 To x + X1
        If j1(q, w) = v_num Then pb0.ForeColor = &HFF0000: pb0.FillColor = &HFF0000:
    gp0.fCircle (d * (w - 0.5)) * 8, 360 - (d * (q - 0.5)) * 8, d * 4
    Next w
    Next q
    'показатели
    ki1.Caption = Round(Ki * 100, 2)
    ki1.Refresh
    pt1.Caption = Round(Pt * 100, 2)
    pt1.Refresh
    rz1.Caption = Round(rZ * 100, 2)
    rz1.Refresh
    e1.Caption = Round(Ey * 100, 2)
    e1.Refresh
    ki2.Caption = Round(Ko * 100, 2)
    ki2.Refresh
    pt2.Caption = Round(Po * 100, 2)
    pt2.Refresh
    rz2.Caption = Round(Ro * 100, 2)
    rz2.Refresh
    e2.Caption = Round(Eo * 100, 2)
    e2.Refresh
    pb2.ForeColor = &HFFFF00: pb2.FillColor = &HFFFF00: gp2.fCircle Ki * 100, 113 - Pt*100, 0.5
    pb2.ForeColor = &H8080FF: pb2.FillColor = &H8080FF: gp2.fCircle Ki * 100, 113 - rZ*100, 0.5
    gp1.fCircle Ki * 100, 113 - Ey * 100, 0.5
    pb1.Refresh: pb2.Refresh
    If Me.WindowState = vbMinimized Then Me.Caption = Round(Ki * 100, 2) & "%выпущено - " & App.Title
    End If
    pb0.Refresh
    tmrMain.Enabled = True
    Exit Sub

    f_sz:

    For tY = q1 To (y + cx + d5 / d)
    If j(tY, w1, r1).rNumber <> tNum Then GoTo wt0
    Next tY
    wt0:
    tY = tY - q1 - 1
    For tX = w1 To (x + X1)
    If j(q1, tX, r1).rNumber <> tNum Then GoTo wt1
    Next tX
    wt1:
    tX = tX - w1 - 1
    For tZ = r1 To z
    If j(q1, w1, tZ).rNumber <> tNum Then GoTo wt2
    Next tZ
    wt2:
    tZ = tZ - r1 - 1
    Return
    razr:
    For q = 1 To y + cx
    For w = 1 To x + X1
    If J(q, w, u).rNumber = 0 Then pb0.ForeColor = vbRed: pb0.FillColor = vbRed:
    gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4: GoTo
    cnt_pnt
    If j(q, w, u).rType = t2 Then pb0.ForeColor = &H4080&: pb0.FillColor = &H4080&:
    gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
    If j(q, w, u).rType = t1 Then pb0.ForeColor = &HC000&: pb0.FillColor = &HC000&:
    gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
    If j(q, w, u).rType = t3 Then pb0.ForeColor = &H404040: pb0.FillColor = &H404040:
    gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
    cnt_pnt: 
    Next w
    Next q
    Return
    End Sub

    


