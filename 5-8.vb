'5
gpO.fLine w * d * 8 + 200, 352, w * d * 8 + 200, 360 - (с - cx * d) * 8
Next w 
d_frm = «175359175354190354190359185359185360185355185359180359180360180358180359175359190359190356192356192358199358199354199355192355»
For q = 1 To Len(d frm) Step 6 
gpO.f Point Val(Mid(d_fm, q, 3)) - 5, Val(Mid (d frm, q + 3, 3))
Next q 
pbO.FillColor = vbWhite: pbO.ForeColor  = ForeColor = &H80000012: gp0.fDraw
pbl.Cis: pb2.Cls 
pbl.ForeColor = &H80000012: pbl.FillColor = &H80000012 
pb2.ForeColor = &H80000012: pb2.FillColor = &H80000012 
gpl.fLine 0, 112, 201, 112: gp2.fLine 0, 112, 201, 112 
gpl.fLine 0, 113, 0,0: gp2.fLine 0, 113, 0, 0
For q = 1 To 4
For w = 0 To 200 Step 5 
gpl.fLine w, 113 - 25 * q, w + 2, 113 - 25 * q
gp2.fLine w, 113 - 25 * q, w + 2, 113 - 25 * q
Next w
Next q
For q = 1 To 8
For w = 0 To 100 Step 5
Eo = 100
If imgMain.Visible = True Then imgMain.Visible = False
cmdRpt.Enabled = False
cmdPlay.Enabled = True 
cmdPlay.SetFocus
k_step = 1
pbO.Refresh: pbl.Refresh: pb2.Refresh
Screen.MousePointer = 0
End Sub

Private Sub cmdRpt_Click()
Dim tmr_flg As Boolean 
tmr_flg = tmrMain.Enabled 
tmrMain.Enabled = False
If MsgBox(«Повторить?», vbQuestion + vbYesNo) = vbNo Then 
    tmrMain.Enabled = tmr_flg 
    Exit Sub
End If
Screen.MousePointer = 11 
tmrMain.Enabled = False 
rpt_flg = True S
creen.MousePointer = 0 
Call cmdPrm_Click 
End Sub

Private Sub Form_Activate()
If ac_flg = True Then Exit Sub
Me.Refresh
ac_flg = True
Call cmdPrm_Click
End Sub

'6
Private Sub Form_Load()
Dim Ssize As Long, ini_tmp() As String
e_£lg - False
Ssize = Screen.Width / 15
If Ssize < 800 Then
    MsgBox "Программа не работает на разрешениях" + vbCrLf + "менее 800x600.", vbCritical
    Unload Me 
    Exit Sub 
End If
Screen.MousePointer = 11 
On Error GoTo prm_err 
q = FreeFile
Open App.Path + "\model.ini" For Input As q
w = 1
Do Until EOF(q)
ReDim Preserve ini_tmp(w)
Line Input #q, ini_tmp(w)
w = w + 1
Loop
Close q
On Error GoTo 0
al = 0: a2 = 0: a3 = 0: d = 0
For q = 1 To UBound(ini tmp)
If LCase(Left(ini_tmp(q), 2 )) = "a1" And Len(ini_tmp(q)) > 3 Then al = Val(Mid(ini_tmp(q), 4))
If LCase(Left(ini_tmp(q), 2 )) = "a2" And Len(ini_tmp(q)) > 3 Then al = Val(Mid(ini_tmp(q), 4))
If LCase(Left(ini_tmp(q), 2 )) = "a3" And Len(ini_tmp(q)) > 3 Then al = Val(Mid(ini_tmp(q), 4))
If LCase(Left(ini_tmp(q), 2 )) = "d" And Len(ini_tmp(q)) > 2 Then al = Val(Mid(ini_tmp(q), 3))
Next q
If (al + 4 * a2 + 4 * a3) <> 1 Then 
    Screen.MousePointer = 0
    MsgBox "Ошибка в параметрах:" + vbCrLf + _
            "А1 + 4 * A2 + 4 * A3 <> 1", vbCritical
    Unload Me
    Exit Sub
End If
Randomize (1)
Set gpO.Device = pbO 
Set gpl.Device = pbl 
Set gp2.Device = pb2 
ac_flg = False 
rpt_flg = False
a0 = 0: b0 = 0: c0 = 0: dl0 = 0: l20 = 0: v_dm0 = 0: v_h0 = 0
v_h0 =0: d20 = 0: d30 = 0: d40 = 0: d5 = d50 'd5 = d50 --?
Screen.MousePointer = 0
Exit Sub
prm_err:
Screen.MousePointer = 0
'7
MsgBox "Ошибка в файле 'model.ini'" + vbCrLf + _
        "или файл не найден.", vbCritical
Unload Me
Exit Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim tmr_flg As Boolean 
If e_flg = True Then 
    tmr_flg = tmrMain.Enabled 
    tmrMain.Enabled = False
    If MsgBox("Выйти из программы?", vbQuestion + vbYesNo) = vbNo Then 
        Cancel = 1
        mrMain.Enabled = tmr_flg 
    End If
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then Me.Caption = "Торцевой выпуск руды (объемная задача)"
End Sub

Private Sub tmrMain_Timer()

tmrMain.Enabled = False

Dim p_flg As Boolean, nk As Integer
Dim nk1 As Integer, v_num As Long
Dim v_type As rType, tNum As Long
Dim tY As Integer, tX As Integer, tZ As Integer
Dim tY1 As Integer, tX1 As Integer, tZ1 As Integer
Dim dX As Integer, dZ As Integer, nD As Double
Dim q2 As Long, w2 As Long, r2 As Long
Dim ep As Long, ep1 As Long, ep2 As Long, ep3 As Long
Dim sp1 As Long, sp2 As Long, cs As Byte
Dim q3 As Long, w3 As Long, r3 As Long, tNum3 As Long
Dim tM() As Integer

If t >= 2 * RdO Then 
    cmdPlay.Enabled = False
    If Me.WindowState = vbMinimized Then GoSub razr
    Exit Sub
End If

'выпускаем кусок
For q = 1 To v_h 
For w = 1 To d5 / d
For e = u - Int(v_dm /2) To u + Int(v_dm / 2)
If j(q, w, e).rNumber <> 0 Then
    cs = 0
    q3 = q: w3 = w: r3 = e 
    GoTo cnt_pr 
End If
Next e: Next w: Next q
p_flg = False
GoTo ch_ms
cnt_pr:
cs = cs + 1

'8
If cs = 3 Then
    nk = q3: w = w3: nk1 = r3 - u 
    GoTo q_3 
End If
nk = Int(Rnd(1) * (v_h - 1)) + 1 
nk1 = Int(Rnd(1) * v_dm / 2) 
nD = Rnd(1)
If nD < 0.5 Then nk1 = -nk1
If j(nk, w, u + nk1).rNumber = 0 Then GoTo cnt_pr
q_3:
v_num = j(nk, w, u + nk1).rNumber: v_type = j(nk, w, u + nk1).rType
q1 = nk - j(nk, w, u + nk1).rY: w1 = w - j(nk, w, u + nk1).rX: r1 = u + nk1 - j(nk, w, u + nk1).rZ
tNum = v_num
GoSub f_sz
For q = q1 To q1 + tY: For w = w1 To w1 + tX: For e = r1 To r1 + tZ 
j(q, w, e).rNumber = 0
Next e: Next w: Next q 
t = t + (tY + 1) * (tX + 1) * (tZ + 1)
Ki = t / RdO
If v_type = t1 Then rd = rd + (tY + 1) * (tX + 1) * (tZ + 1) Else pr = pr + (tY + 1) * (tX + 1) * (tZ + 1)
rZ = pr / t
Ft = (RdO - rd) / RdO
Ey = Pt + pr / (n - 5 * z * (x + X1)) * 0.4 
If Ey < Ec Then
Eo = Ey: Po = Pt: Ro = rZ: Ko = Ki
End If
p_flg = True

'nepебираем массив
ch ms:
For q = 1 To y + cx 
For w = 1 To x + X1 
For e = 1 To z
If j(q, w, e).rNumber = 0 Then
    For tY1 = q To (y + cx + d5 / d)
    If j(tY1, w, e).rNumber <> 0 Then GoTo wt0_0
    Next tY1
wt0_0:
    tY1 = tY1 - q - 1
    tX1 = x + X1
    For q2 = 0 To tY1
    For q3 = w To (x + X1)
    If j(q + q2, q3, e).rNumber <> 0 Then GoTo wt0_1
    Next q3
wt0_1:
    q3 = q3 - w - 1
    If tX1 > q3 ThentX1 = q3
    Next q2
    tZ1 = dZ
    For q2 = 0 To tY1
    For q3 = e To z
    If j(q + q2, w, q3).rNumber <> 0 Then GoTo wt0_2
    Next q3
wt0_2:
    q3 = q3 - e - 1
    If tZ1 > q3 Then tZ1 = q3