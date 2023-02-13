Option Explicit
Public a As Double, b As Double, c As Double, d1 As Double
Public l2 As Double, wrk_flg As Boolean, v_dm As Double
Public v_h As Double, d2 As Double, d3 As Double
Public d4 As Double, d5 As Double, d50 As Double
Dim d10 As Double, d20 As Double, d30 As Double, d40 As Double
Dim a1 As Double, a2 As Double, a3 As Double
Dim a0 As Double, b0 As Double, c0 As Double, v_h0 As Double
Dim l20 As Double, v_dm0 As Double
Dim cx As Double, e_flg As Boolean
Dim x As Double, y As Double, z As Double, n As Double
Dim u As Double, v As Double, d As Double
Dim xs As Integer, zs As Integer
Dim X1 As Double, Rd0 As Double
Dim q As Double, w As Double, e As Double, t As Double
Dim q1 As Double, w1 As Double, r1 As Double
Dim gp0 As New Graphics, gp1 As New Graphics, gp2 As New Graphics
Dim ac_flg As Boolean, j() As Rock, j1() As Long
Dim k_step As Integer, rpt_flg As Boolean
Dim Ki As Double, Pt As Double, rZ As Double
Dim rd As Double, pr As Double, Ey As Double
Dim Ko As Double, Po As Double, Ro As Double, Eo As Double

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPlay_Click()
If k_step = 1 Then
  Screen.MousePointer = 11
  pb0.ForeColor = pb0.BackColor: pb0.FillColor = pb0.BackColor
  gp0.fLineB 0, 360 - (c - cx * d) * 8, a * 8, 352
  gp0.fLineB 200, 360 - (c - cx * d) * 8, a * 8 + 200, 352
  For q = 1 To y + cx
  If q > y Then
    pb0.ForeColor = &H4080&: pb0.FillColor = &4080&
    For w = 1 To x + X1
    gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8. d * 4
    gp0.fCircle (d * (w - 0.5)) * 8, 360 - (d * (q - 0.5)) * 8, d * 4
    Next w
  Else
    If X1 > 0 Then
      pb0.ForeColor = &H404040: pb0.FillColor = &H404040
      For w = x + 1 To x + X1
      gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
      gp0.fCircle (d * (w - 0.5)) * 8, 360 - (d * (q - 0.5)) * 8, d * 4
      Next w
    End If
  End If
  Next q
  pb0.ForeColor = &HC000&: pb0.FillColor = &HC000&
  For q = 1 To y
  For w = 1 To x
  gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
  gp0.fCircle (d * (w - 0.5)) * 8, 360 - (d * (q - 0.5)) * 8, d * 4
  Next w
  Next q
  k_step = 2
  pb0.Refresh
  Screen.MousePointer = 0
  Exit Sub
End If
If k_step = 2 Then
  Screen.MousePointer = 11
  rd = 0: pr = 0
  t = 0
  k_step = 3
  pb1.ForeColor = &HC0&: pb1.FillColor = &HC0&
  cmdRpt.Enabled = True
  tmrMain.Enabled = True
  Screen.MousePointer = 0
  Exit Sub
End If
If k_step = 3 Then tmrMain.Enabled = Not tmrMain.Enabled
End Sub

Private Sub cmdPrm_Click()
Dim d_frm As String
Screen.MousePointer = 11
tmrMain.Enabled = False
rpt_prm:
a = a0: b = b0: c = c0: l2 = l20: v_dm = v_dm0: v_h = v_h0
d1 = d10: d2 = d20: d3 = d30: d4 = d40: d5 = d50
Screen.MousePointer = 0
If rpt_flg = False Then
' frmPrm.Show 1
' If wrk_flg = False Then Exit Sub

a = 6: b = 12: c = 20: v_h = 3.8: v_dm = 4: l2 = 0
d1 = 0.6: d2 = 1: d3 = 0.8: d4 = 1: d5 = 1.4

End If
Screen.MousePointer = 11
ki1.Caption = <<: ki2.Caption = <<
pt1.Caption = <<: pt2.Caption = <<
rz1.Caption = <<: rz2.Caption = <<
e1.Caption = <<: e2.Caption = <<
Me.Refresh
rpt_flg = False
If e_flg = False Then e_flg = True
a0 = a: b0 = b: c0 = c: l20 = l2: v_dm0 = v_dm: v_h0 = v_h
d10 = d1: d20 = d2: d30 = d3: d40 = d4: d50 = d5

d1 = Int(d1 / d) * d: d2 = Int(d2 / d) * d: d3 = Int(d3 / d) * d
d4 = Int(d4 / d) * d: d5 = Int(d5 / d) * d

x = Int(a / d): y = Int(c / d): z = Int(b / d)
X1 = Int(12 / d): cx = 2 * d5 / d
a = x * d: c = y * d + cx * d: b = z * d: 12 = X1 * d
v_dm = Int(v_dm / d): If v_dm = 0 Then v_dm = 1
v_h = Int(v_h / d): If v_h = 0 Then v_h = 1
ReDim j(y + cx + d5 / d, x + X1, z)
n = 1: RdO = 0

'заполнение рудой
For q = 1 To Int(y / 5) Step d1 /d
For w = 1 To x Step d1 / d
For e = 1 To z Step d1 / d
For q1 = 0 To d1 / d - 1: For w1 = 0 To d1 / d - 1: For r1 = 0 To d1 / d - 1
If (q + q1) <= Int(y / 5) And (w + w1) <= x And (e + r1) <= z Then
  j(q + q1, w + w1, e + r1).rType = t1: j(q + q1, w + w1, e + r1).rNumber = n: j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1: j(q + q1, w + w1, e + r1).rZ = r1
  RdO = RdO + 1
End If
Next r1: Next w1: Next q1
n = n + 1
Next e: Next w: Next q
For q = Int(y / 5) + 1 To 2 * Int(y / 5) Step d2 / d
For w = 1 To x Step d2 / d
For e = 1 To z Step d2 / d
For q1 = 0 To d2 / d - 1: For w1 = 0 To d2 / d - 1: For r1 = 0 To d2 / d - 1
If (q + q1) <= 2 * Int(y / 5) And (w + w1) <= x And (e + r1) <= z Then
  j(q + q1, w + w1, e + r1).rType = t1: j(q + q1, w + w1, e + r1).rNumber = n: j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1: j(q + q1, w + w1, e + r1).rZ = r1
  RdO = RdO + 1
End If
Next r1: Next w1: Next q1
n = n + 1
Next e: Next w: Next q
For q = 2 * Int(y / 5) + 1 To 4 * Int(y / 5) Step d3 / d
For w = 1 To x Step d3 / d
For e = 1 To z Step d3 / d
For q1 = 0 To d3 / d - 1: For w1 = 0 To d3 / d - 1: For r1 = 0 To d5 / d - 1
If (q + q1) <= 4 * Int(y / 5) And (w + w1) <= x And (e + r1) <= z Then
  j(q + q1, w + w1, e + r1).rType = t1: j(q + q1, w + w1, e + r1).rNumber = n: j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1: j(q + q1, w + w1, e + r1).rZ = r1
  RdO = RdO + 1
End If
Next r1: Next w1: Next q1
n = n + 1
Next e: Next w: Next q
For q = 4 * Int(y / 5) + 1 To y Step d4 / d
For w = 1 To x Step d4 / d
For e = 1 To z Step d4 / d
For q1 = 0 To d4 / d - 1: For w1 = 0 To d4 / d - 1: For r1 = 0 To d4 / d - 1
If (q + q1) <= y And (w + w1) <= x And (e + r1) <= z Then
  j(q + q1, w + w1, e + r1).rType = t1: j(q + q1, w + w1, e + r1).rNumber = n: j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1: j(q + q1, w + w1, e + r1).rZ = r1
  RdO = RdO + 1
End If
Next r1: Next w1: Next q1
n = n + 1
Next e: Next w: Next q
'заполнение породой кровли
For q = y + 1 To y + cx + d5 / d Step d5 / d
For w = 1 To x + X1 Step d5 / d
For e = 1 To z Step d5 / d
For q1 = 0 To d5 / d - 1: For w1 = 0 To d5 / d - 1: For r1 = 0 To d5 / d - 1
If (q + q1) <= (y + cx + d5 / d) And (w + w1) <= (x + X1) And (e + r1) <= z Then
  j(q + q1, w + w1, e + r1).rType = t2: j(q + q1, w + w1, e + r1).rNumber = n: j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1: j(q + q1, w + w1, e + r1).rZ = r1
End If
Next r1: Next w1: Next q1
n = n + 1
Next e: Next w: Next q
'заполнение боковыми породами
If X1 <> 0 Then
  For q = 1 To y Step d5 / d
  For w = x + 1 To x + X1 Step d5 / d
  For e = 1 To z Step d5 / d
  For q1 = 0 To d5 / d - 1: For w1 = 0 To d5 / d - 1: For r1 = 0 To d5 / d - 1
  If (q + q1) <= y And (w + w1) <= (x + X1) And (e + r1) <= z Then
    j(q + q1, w + w1, e + r1).rType = t3: j(q + q1, w + w1, e + r1).rNumber = n: j(q + q1, w + w1, e + r1).rY = q1: j(q + q1, w + w1, e + r1).rX = w1:j(q + q1, w + w1, e + r1).rZ = r1
  End If
  Next r1: Next w1: Next q1
  n = n + 1
  Next e: Next w: Next q
End If
'подготовка
u = Int(z / 2)
ReDim j1(y + cx, x + X1)
For q = 1 To y + cx
For w = 1 To x + X1
j1(q, w) = j(q, w, u).rNumber
Next w
Next q
pb0.Cls
For q = 1 To y + cx
If q > y Then
  pb0.ForeColor = &H4080&: pb0.FillColor = &H4080&
  For w = 1 To x + X1
  gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
  gp0.fCircle (d * (w - 0.5)) * 8, 360 - (d * (q - 0.5)) * 8, d * 4
  Next w
Else
  If X1 > 0 Then
    pb0.ForeColor = &H404040: pb0.FillColor = &H404040
    For w = x + 1 To x + X1
    gp0.fCircle (d * w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4
    gp0.fCircle (d * w - 0.5)) * 8, 360 - (d * (q - 0.5)) * 8, d * 4
    Next w
  End If
End If
Next q
pb0.ForeColor = &HFFFF00: pb0.FillColor = &HFFFF00
gp0.fLineB 200, 352 - (v_h - 2) * d * 8, 160, 360 - c * 8
gp0.fLineB 160, 360, (a + 12) * 8 + 200, 380
gp0.fLineB 0, 360, (a + 12) * 8, 380
gp0.fLineB 200, 352, a * 8 + 200, 360 - (c - cx * d) * 8
gp0.fLineB 0, 252, a * 8, 360 - (c - cx * d) * 8
pb0.ForeColor = &H80000012: pb0.FillColor = &H80000012
For w = 0 To x - 1 Step 8
gp0.fLine w * d * 8, 352, w * d * 8, 360 - (c - cx * d) * 8
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
Screen.MousePointer = 0 
Call cmdPrm_Click 
End Sub

Private Sub Form_Activate()
If ac_flg = True Then Exit Sub
Me.Refresh
ac_flg = True
Call cmdPrm_Click
End Sub

Private Sub Form_Load()
Dim Ssize As Long, ini_tmp() As String
e_flg = False
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
For q = 1 To UBound(ini_tmp)
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
If d = 0 Then
  Screen.MousePointer = 0
  MsgBox "Ошибка в параметрах.", vbCritical
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
gp0.fCircle (d * (w - 0.5)) * 8 + 200, 360 - (d * (q - 0.5)) * 8, d * 4: GoTo cnt_pnt
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

    


