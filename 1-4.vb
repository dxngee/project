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
End Id
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
