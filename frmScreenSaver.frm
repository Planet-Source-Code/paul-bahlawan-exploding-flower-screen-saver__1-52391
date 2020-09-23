VERSION 5.00
Begin VB.Form frmScreenSaver 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7695
   FillColor       =   &H0000FF00&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   840
   End
End
Attribute VB_Name = "frmScreenSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' EXPLODING FLOWER Screen Saver (screen saver form)
''' By Paul Bahlawan
''' March 8, 2004 (original flower design)
'''
''' Revisions:
''' March 12, 2004 - add screen saver stuff
'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Elem(9) As Long
Private cx(9) As Long
Private cy(9) As Long
Private dx(9) As Long
Private dy(9) As Long
Private ca(9) As Long
Private da(9) As Long
Private l1(9) As Long
Private l2(9) As Long
Private w(9) As Long
Private bnc(9) As Single
Private bnr(9) As Single
Private r(9) As Long
Private g(9) As Long
Private b(9) As Long
Private dr(9) As Long
Private dg(9) As Long
Private db(9) As Long

Private Px As Long
Private Py As Long
Private oldpos As Long


Private Const Pi180 As Double = 3.14159265 / 180

Private Sub Form_Load()
Randomize
BorderStyle = 0
Caption = sMode
WindowState = 2
FillStyle = 0
FillColor = 0
BackColor = 0
ScaleMode = vbPixels
 
If sMode <> "Preview" Then
    SystemParametersInfo SPI_SCREENSAVERRUNNING, True, vbNull, 0
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    ShowCursor False
End If
Timer1.Interval = 25 + 5 * userSpeed
End Sub

'Main Flower routine, fired by timer
Private Sub Timer1_Timer()
Dim cnt As Long
Dim angle As Long
Dim PolySet(3) As POINTAPI
Dim i As Long
'Timer1.Enabled = False 'uncomment this + Do + DoEvents + Loop for full throttle test.
'Do

'Create Flower
i = Int(Rnd(1) * 100)
If i < userQty Then
    If Elem(i) = 0 Then
        Elem(i) = Int(Rnd(1) * userPetals * 2) + 3  'number of petals
        cx(i) = Int(Rnd(1) * ScaleWidth)            'center of flower
        cy(i) = Int(Rnd(1) * ScaleHeight)
        Do
            dx(i) = Int(Rnd(1) * 9) - 4             'direction
            dy(i) = Int(Rnd(1) * 9) - 4
        Loop While dx(i) = 0 And dy(i) = 0
        ca(i) = Int(Rnd(1) * 360)
        da(i) = Int(Rnd(1) * 7) - 3                 'spin
        l1(i) = Int(Rnd(1) * 100 * userSize) + 20   'petal length
        If sMode = "Preview" Then l1(i) = l1(i) / 4
        l2(i) = Int(Rnd(1) * (l1(i) * 0.19 * userForked))
        bnc(i) = Rnd(1)                             'bounce
        bnr(i) = Rnd(1) / 20                        'bounce rate
        w(i) = Int(Rnd(1) * 4 * userWidth) + 1      'petalwidth
        r(i) = Int(Rnd(1) * 252) + 3                'colours
        g(i) = Int(Rnd(1) * 252) + 3
        b(i) = Int(Rnd(1) * 252) + 3
        dr(i) = Int(Rnd(1) * 7) - 3
        dg(i) = Int(Rnd(1) * 7) - 3
        db(i) = Int(Rnd(1) * 7) - 3
    End If
End If


'Draw Flower(s)
For i = 0 To userQty - 1
    If Elem(i) > 0 Then
        For cnt = 0 To Elem(i) - 1
            angle = ca(i) + (360 / Elem(i) * cnt)
        
            PolySet(0).X = cx(i)
            PolySet(0).Y = cy(i)
            
            Polar l2(i) * bnc(i), angle + w(i)
            PolySet(1).X = cx(i) + Px
            PolySet(1).Y = cy(i) + Py
            
            Polar l1(i) * bnc(i), angle
            PolySet(2).X = cx(i) + Px
            PolySet(2).Y = cy(i) + Py
            
            Polar l2(i) * bnc(i), angle - w(i)
            PolySet(3).X = cx(i) + Px
            PolySet(3).Y = cy(i) + Py
            
            ForeColor = RGB(r(i), g(i), b(i))
            Polygon Me.hdc, PolySet(0), 4
        Next

'Move Flowers
        cx(i) = cx(i) + dx(i)
        cy(i) = cy(i) + dy(i)
        'rotate
        ca(i) = (ca(i) + da(i)) Mod 360
        'If ca(i) < 0 Then ca(i) = 360
        'If ca(i) > 360 Then ca(i) = 0
        'bounce
        bnc(i) = bnc(i) + bnr(i)
        If bnc(i) > 1 Or bnc(i) < 0.05 Then bnr(i) = -bnr(i)
        'colours
        r(i) = r(i) + dr(i)
        If r(i) > 253 Or r(i) < 3 Then dr(i) = -dr(i)
        g(i) = g(i) + dg(i)
        If g(i) > 253 Or g(i) < 3 Then dg(i) = -dg(i)
        b(i) = b(i) + db(i)
        If b(i) > 253 Or b(i) < 3 Then db(i) = -db(i)

'Terminate Flowers
        If cx(i) < -l1(i) * bnc(i) Or cx(i) > ScaleWidth + l1(i) * bnc(i) Then Elem(i) = 0
        If cy(i) < -l1(i) * bnc(i) Or cy(i) > ScaleHeight + l1(i) * bnc(i) Then Elem(i) = 0
    
    End If
Next

'DoEvents
'Loop
End Sub

'Convert polar coord to cartisian coord
Private Sub Polar(ByVal r As Single, ByVal t As Single)
t = t * Pi180
Px = r * Cos(t)
Py = r * Sin(t)
End Sub

'Exit screen saver on key press
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ExitSaver
End Sub

'Exit screen saver on mouse move
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If oldpos = 0 Then oldpos = X
If oldpos < X - 10 Or oldpos > X + 10 Then ExitSaver
End Sub

Private Sub ExitSaver()
If sMode <> "Preview" Then
    If bUsePassword() = True Then
        ShowCursor True
        If VerifyScreenSavePwd(Me.hwnd) = False Then
            oldpos = 0
            ShowCursor False
            Exit Sub
        End If
    End If
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If sMode <> "Preview" Then
    SystemParametersInfo SPI_SCREENSAVERRUNNING, False, vbNull, 0
    ShowCursor True
End If
End Sub

