VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   Caption         =   "Stereoscopic function visualization 1.0"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsc 
      Height          =   1575
      Index           =   3
      LargeChange     =   30
      Left            =   120
      Max             =   10
      Min             =   100
      TabIndex        =   10
      Top             =   600
      Value           =   100
      Width           =   255
   End
   Begin VB.CheckBox chkfrozenscape 
      BackColor       =   &H80000007&
      Caption         =   "Freeze landscape"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtin 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   365
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   110
      Width           =   6495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O O "
      Height          =   255
      Index           =   6
      Left            =   8040
      TabIndex        =   7
      Top             =   6000
      Width           =   490
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cube"
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.HScrollBar hsbipd 
      Height          =   255
      LargeChange     =   18
      Left            =   4560
      Max             =   190
      Min             =   100
      TabIndex        =   5
      Top             =   6000
      Value           =   130
      Width           =   1695
   End
   Begin VB.VScrollBar vsc 
      Height          =   1575
      Index           =   2
      LargeChange     =   12
      Left            =   120
      Max             =   15
      Min             =   80
      TabIndex        =   4
      Top             =   4680
      Value           =   54
      Width           =   255
   End
   Begin VB.VScrollBar vsc 
      Height          =   1575
      Index           =   1
      LargeChange     =   30
      Left            =   7920
      Max             =   330
      Min             =   10
      TabIndex        =   3
      Top             =   600
      Value           =   100
      Width           =   255
   End
   Begin VB.VScrollBar vsc 
      Height          =   1575
      Index           =   0
      LargeChange     =   15
      Left            =   8280
      Max             =   20
      Min             =   250
      TabIndex        =   2
      Top             =   600
      Value           =   80
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "landy"
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H80000008&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   389
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   573
      TabIndex        =   0
      Top             =   480
      Width           =   8655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'fluoats@hotmail.com
'Yaay, i can paste Angle transformation formulas

'if you want to play with the 3d plots formula, find
'Case sineland'


Dim strfula As String, fulaselect As Byte
Dim xr As Single, yr As Single, lastformula As Byte, vel As Byte
Dim xr2 As Single, yr2 As Single
Dim npoints As Integer, ipdr As Single, cube As Boolean
Dim changednpoints As Boolean, eyemode As Boolean, clear As Boolean
Dim expan As Single, freq As Single, frozenland As Boolean
Dim ipd As Byte, oscrate As Single, radius As Single, vscr As Single
Dim roll As Single, pitch As Single, yaw As Single 'center point angles
Dim kpts As Boolean, pressed As Boolean, osc As Boolean, sineland As Boolean
Dim yawi As Single, pitchi As Single, rolli As Single 'rotation speed
Dim cx As Integer, cy As Integer, cz As Single 'center point coords
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Dim breakloop As Boolean, gbye As Boolean


Private Sub Stuff()
Static test1 As Single, test2 As Single, test3 As Single

'holds coordinate data of cube corners
Static x3(0 To 4000) As Single, y3(0 To 4000) As Single, z3(0 To 4000) As Single

'after pixel is drawn, these store the pixel's location for reasons explained _
 within the loop
Static prevx(0 To 4000) As Integer, prevy(0 To 4000) As Integer
Static prev2x(0 To 4000) As Integer, prev2y(0 To 4000) As Integer
 
Dim vpd As Single  'vanishing-point distortion

'perform initial computations for 3d rotation
Dim y2 As Single

'a second set also speeds things up
Dim z1 As Single

'finally, these points get plotted
Dim x As Single, y As Single, z As Single

Static red As Integer, grn As Integer, blu As Integer
Static n As Integer, a As Single, b As Single, c As Single
Static landx As Single, landz As Single
Static incred As Single, incgrn As Single, incblu As Single
Static n1 As Integer, n2 As Integer
Static ix As Single, iz As Single
Static csyaw As Single, snyaw As Single
Static csyaw2 As Single, snyaw2 As Single
Static csroll As Single, snroll As Single
Static cspitch As Single, snpitch As Single
Static eye As Single, firstrun As Boolean
Const pi As Single = 3.14159265
Const twopi = 2 * pi
Const hp = pi * 0.5


Select Case firstrun 'I do some initializations here to free up Form_Load
Case 0: firstrun = 1 'This makes sure that this only happens once
Randomize
sineland = True: eye = 200
cx = 285: cy = 170: ipd = 92
red = 255: grn = Rnd * 255: blu = 255: vscr = 50
rolli = 0!: yawi = 0.023!: pitchi = 0!
a = yaw: b = pitch: c = roll: ipdr = 8 / 13
incred = -0.24: incgrn = -0.37: incblu = -0.52
npoints = vsc(0).Value: expan = 1: freq = 0.1: vel = 100
End Select

radius = vscr + 5

'Stuff function points into x3(),y3(),z3() arrays
Select Case sineland
Case True: n = 0
  'control arrays vsc(0) and vsc(1) adjust freq and expan
  For ix = -expan To expan Step freq * expan
  For iz = -expan To expan Step freq * expan
  x3(n) = ix / expan: z3(n) = iz / expan 'leave these

  'landz and landx increase just outside Sel Case sineland ..
  'adds dimension of time

  'These are some example formulas
  Select Case fulaselect
  Case 0: y3(n) = Sin(ix + landx + 4 * Sin(Cos(landx + iz)))
  Case 1: y3(n) = 0.15 * Sin(3 * ix ^ 2 + 3 * iz ^ 2 + 0.014 + landz)
  Case 2: y3(n) = 0.23 * Sin(ix * 3 + landz) + 0.23 * Sin(iz * 3 + landx)
  Case 3: y3(n) = Cos(Sqr(ix ^ 2 + iz ^ 2) + landx) * (-(ix ^ 2 + iz ^ 2) + 1)
  Case 4: y3(n) = Sin(ix ^ 2 + landx) ^ 2 / (0.134 + iz ^ 2) + Sin(iz + (iz * Sin(landx)))
  Case 5: y3(n) = 0.4 * Cos((ix + Cos(landx + 2 * iz + 2))) + 0.4 * Sin(2 * iz * ix + 1.1 + landx)
  Case 6: y3(n) = 0.1 / Sin(ix * 5 + landz) + 0.1 * 1 / Cos(iz * 5 + landx + 0.014)
  Case 7: y3(n) = 1 / (Cos(3 * ix ^ 2 + landx) + Sin(iz ^ 3 + 1))
  End Select
  lastformula = 7 'accessed by command2_Click Index Case 4 .. _
                   increase this if u add any cases and wanna c them
    
  'Final adjustments
  Select Case y3(n) 'This adds a safety ceiling and floor
  Case Is > 1.32: y3(n) = 1.32
  Case Is < -1.32: y3(n) = -1.32
  End Select
  y3(n) = radius * y3(n)
  x3(n) = radius * x3(n)
  z3(n) = radius * z3(n)
   
  n = n + 1
  Next iz
  Next ix: npoints = n - 1 '# points now in x3(),y3(),z3() arrays
End Select


'-----------------
Select Case frozenland
Case 0
 'rates of sine-scape travel
 landx = landx + 0.13 * vel * 0.01
 landz = landz + 0.32 * vel * 0.01
 Select Case landx
  Case Is > twopi: landx = landx - twopi: End Select
 Select Case landz
  Case Is > twopi: landz = landz - twopi: End Select
End Select

Select Case cube
Case True 'x3(0) thru x3(7) is the actual cube.  All points beyond _
are added for style
 x3(0) = -0.7: y3(0) = 0.7: z3(0) = -0.7
 x3(1) = 0.7: y3(1) = 0.7: z3(1) = -0.7
 x3(2) = 0.7: y3(2) = 0.7: z3(2) = 0.7
 x3(3) = -0.7: y3(3) = 0.7: z3(3) = 0.7
 x3(4) = -0.7: y3(4) = -0.7: z3(4) = -0.7
 x3(5) = 0.7: y3(5) = -0.7: z3(5) = -0.7
 x3(6) = 0.7: y3(6) = -0.7: z3(6) = 0.7
 x3(7) = -0.7: y3(7) = -0.7: z3(7) = 0.7
 
 x3(8) = -0.7: y3(8) = -0.7: z3(8) = 0.7
 x3(9) = -0.5: y3(9) = -0.7: z3(9) = 0.7
 x3(10) = -0.4: y3(10) = -0.7: z3(10) = 0.7
 x3(11) = -0.3: y3(11) = -0.7: z3(11) = 0.5
 x3(12) = -0.2: y3(12) = -0.7: z3(12) = 0.4
 x3(13) = -0.2: y3(13) = -0.7: z3(13) = 0.2
 x3(14) = -0.2: y3(14) = -0.7: z3(14) = 0!
 x3(15) = -0.2: y3(15) = -0.7: z3(15) = -0.1
 x3(16) = -0.2: y3(16) = -0.7: z3(16) = -0.3
 x3(17) = -0.2: y3(17) = -0.5: z3(17) = 0.7
 x3(18) = -0.2: y3(18) = -0.3: z3(18) = 0.7
 x3(19) = 0.2: y3(19) = -0.1: z3(19) = 0.7
 npoints = 19
 For n = 0 To npoints Step 1
  x3(n) = x3(n) * radius: y3(n) = y3(n) * radius: z3(n) = z3(n) * radius
 Next n
End Select

ipd = vscr / ipdr

Select Case eyemode 'offset yaw for either parallel or cross-eye
Case 0: snyaw2 = Sin(a - 0.08): csyaw2 = Cos(a - 0.08)
Case 1: snyaw2 = Sin(a + 0.08): csyaw2 = Cos(a + 0.08)
End Select


'Final preparations just before rendering loop
ix = 0.5 * (radius + eye) 'this will be used to dim points "farther from eye"
snyaw = Sin(a): csyaw = Cos(a)
snpitch = Sin(b): cspitch = Cos(b)
snroll = Sin(c): csroll = Cos(c)



'    Here is the LOOP!
'compute 3d rotation on point3(n), store the x and y coords
For n = 0 To npoints Step 1
x = x3(n) * cspitch + y3(n) * snpitch
y = -x3(n) * snpitch + y3(n) * cspitch
y = y * csroll + z3(n) * snroll
z1 = -y * snroll + z3(n) * csroll
z = -x * snyaw + z1 * csyaw

'only draw poins within given distance to and from eye
Select Case z
Case Is < 120
 Select Case z
 Case Is > -159
  test3 = (z + 159) / ix
  vpd = eye / (eye - z)  'vanishing-point distortion
  x = vpd * (x * csyaw + z1 * snyaw)
  y = vpd * y

 'Final adjustments to x for left-of-center, and y to half screen height
  x = x + cx - ipd: y = y + cy
 
  'erase previous point then draw new
  SetPixelV pic1.hdc, prevx(n), prevy(n), vbBlack
  SetPixelV pic1.hdc, x, y, RGB(red * test3, grn * test3, blu * test3)
  prevx(n) = x: prevy(n) = y 'store new
End Select: End Select
 
'second set for other half of screen
x = x3(n) * cspitch + y3(n) * snpitch
y = -x3(n) * snpitch + y3(n) * cspitch
y = y * csroll + z3(n) * snroll
z1 = -y * snroll + z3(n) * csroll
z = -x * snyaw2 + z1 * csyaw2

Select Case z
Case Is < 120
 Select Case z
 Case Is > -159
  vpd = eye / (eye - z)
  x = vpd * (x * csyaw2 + z1 * snyaw2)
  y = vpd * y
  x = x + cx + ipd: y = y + cy
  SetPixelV pic1.hdc, prev2x(n), prev2y(n), vbBlack
  SetPixelV pic1.hdc, x, y, RGB(red * test3, grn * test3, blu * test3)  'plots new point
  prev2x(n) = x: prev2y(n) = y
End Select: End Select
Next n


'increment rotation
a = a + yawi
b = b + pitchi
c = c + rolli

If a > 2 * pi Then  'keeps rotation inside the first octave
 a = a - 2 * pi
ElseIf a < 0 Then
 a = a + 2 * pi
End If

'roll
If b > 2 * pi Then
 b = b - 2 * pi
ElseIf b < 0 Then
 b = b + 2 * pi
End If

'pitch
If c > 2 * pi Then
 c = c - 2 * pi
ElseIf c < 0 Then
 c = c + 2 * pi
End If



'Integers red, grn, blu stay within byte range
red = red + incred
If red > 255 Then
red = 510 - red: incred = -incred
ElseIf red < 150 Then
red = 150: incred = -incred
End If

grn = grn + incgrn
If grn > 255 Then
grn = 510 - grn: incgrn = -incgrn
ElseIf grn < 150 Then
grn = 150: incgrn = -incgrn
End If

blu = blu + incblu
If blu > 255 Then
blu = 510 - blu: incblu = -incblu
ElseIf blu < 150 Then
blu = 150: incblu = -incblu
End If
End Sub

Private Sub Form_Activate()
Do While Not breakloop
DoEvents
Select Case breakloop
Case True
 Exit Do
End Select

Call Stuff

Loop
breakloop = False

Select Case gbye
Case True
 Unload Me: End
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
breakloop = True
gbye = True
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = 1
xr = x
yawi = 0
yr = y
pitchi = 0
xr2 = 0: yr2 = 0
End Sub
Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case pressed
Case 1
 xr = xr - x
 If xr > 0 Then
  If xr > xr2 Then
   xr2 = xr
  End If
  yawi = yawi * 0.8
  If xr > 6 Then
   yawi = yawi + xr2 * 2 / (2300 - npoints * 0.5)
  Else
   yawi = yawi + xr * 2 / (2300 - npoints): xr2 = xr2 - xr
  End If
 
 ElseIf xr < 0 Then
  If xr < xr2 Then
   xr2 = xr
  End If
  yawi = yawi * 0.8
  If xr < -6 Then
   yawi = yawi + xr2 * 2 / (2300 - npoints * 0.5)
  Else
   yawi = yawi + xr * 2 / (2300 - npoints): xr2 = xr2 - xr
  End If
 End If
 
 yr = yr - y
 If yr > 0 Then
  If yr > yr2 Then
   yr2 = yr
  End If
  pitchi = pitchi * 0.8
  If yr > 6 Then
   pitchi = pitchi + yr2 * 2 / (2300 - npoints * 0.5)
  Else
   pitchi = pitchi + yr * 2 / (2300 - npoints): yr2 = yr2 - yr
  End If
 
 ElseIf yr < 0 Then
  If yr < yr2 Then
   yr2 = yr
  End If
  pitchi = pitchi * 0.8
  If yr < -6 Then
   pitchi = pitchi + yr2 * 2 / (2300 - npoints * 0.5)
  Else
   pitchi = pitchi + yr * 2 / (2300 - npoints): yr2 = yr2 - yr
  End If
 End If
 xr = x: yr = y
End Select
End Sub

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
pressed = 0
Select Case yawi
Case Is > 1
yawi = 1
End Select
Select Case pitchi
Case Is > 1
pitchi = 1
End Select
End Sub
Private Sub vsc_Change(Index As Integer)
Select Case Index
Case 0: freq = 8 / vsc(0).Value    'Min:250 Value:80 Max:20
 radius = (80 + vsc(0).Value) / 250
 yawi = yawi * radius: pitchi = pitchi * radius
Case 1: expan = vsc(1).Value / 100 'Min:10 Value:100 Max:330
Case 2: vscr = vsc(2).Value
Case 3: vel = vsc(3).Value
End Select
pic1.Cls
End Sub
Private Sub chkfrozenscape_Click()
frozenland = chkfrozenscape.Value
End Sub
Private Sub hsbipd_Change()
ipdr = 80 / hsbipd.Value
End Sub


Private Sub Command2_Click(Index As Integer)
Select Case Index

Case 4: sineland = True: cube = False: pic1.Cls: freq = 8 / vsc(0).Value
 vsc(1).Visible = True: vsc(0).Visible = True: vsc(3).Visible = True
 Command2(5).Visible = True
 If chkfrozenscape.Visible Then
  fulaselect = fulaselect + 1
 End If
 chkfrozenscape.Visible = True
 If fulaselect > lastformula Then
  fulaselect = 0: Command2(4).Caption = "landy"
 ElseIf fulaselect > 0 Then
  Command2(4).Caption = fulaselect + 1
 End If
 
Case 5: pic1.Cls: cube = True: sineland = False ': freq = 0.0138
 vsc(1).Visible = False: vsc(0).Visible = False: vsc(3).Visible = False
 chkfrozenscape.Visible = False: Command2(5).Visible = False
Case 6
 Select Case eyemode
 Case 0: eyemode = 1: Command2(6).Caption = " >.< "
 Case 1: eyemode = 0: Command2(6).Caption = " O O "
 End Select
 pic1.SetFocus
End Select
End Sub
