VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Rotating Circle Sample "
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrButtons 
      Interval        =   1
      Left            =   3840
      Top             =   720
   End
   Begin VB.Shape PointerL 
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape ManL 
      Height          =   255
      Left            =   240
      Shape           =   2  'Oval
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Outside 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape Pointer 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label lblInfo 
      Caption         =   "X=? T=? DEG=?"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Man 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   720
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Degrees As Integer

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Form_Load()
Degrees = 1
PlacePointer
End Sub

Private Sub lblInfo_Click()
Dim x As Variant
x = InputBox("Write a number between 0 and 360")
If IsNumeric(x) = False Then Exit Sub
If x > 360 Or x < 0 Then Exit Sub
Degrees = Int(x)
End Sub

Private Sub tmrButtons_Timer()
If GetAsyncKeyState(vbKeyUp) <> 0 Then ' Trigger the under only if key UP is pressed
  Man.Top = Man.Top + 50 * Sin(Degrees * (3.14 / 180)) ' Move it up +50 timed the sinus of the degrees timed with pi/180
  Man.Left = Man.Left + 50 * Cos(Degrees * (3.14 / 180)) ' Move it left +50 timed the cosinus of the degrees timed with pi/180
End If
If GetAsyncKeyState(vbKeyLeft) <> 0 Then ' Trigger the under only if key LEFT is pressed
  If Degrees <= 0 Then Degrees = 365 ' If its under 0 set it to 362
  Degrees = Degrees - 5 ' Change the degrees -2
End If
If GetAsyncKeyState(vbKeyRight) <> 0 Then ' Trigger the under only if key RIGHT is pressed
  If Degrees >= 360 Then Degrees = -5 ' If its over 260 then set it to -2
  Degrees = Degrees + 5 ' Change the degrees +2
End If
lblInfo = "X=" & Man.Left & " Y=" & Man.Top & " DEG=" & Degrees ' Sets the label..
Call PlacePointer ' Placint the pointer just some stuff i created to show how that is done :P
Call OutsideShower ' Call the showing if its outside. Dont care about this if you thing all over is TOO hard ;)
Call DirectionPointer ' Call the directionpointer small, for showing u the way ur running like if ur outside the wall ;)
End Sub

Public Sub PlacePointer()
Dim Left As Integer ' Making a left integer (only to make the math pice look smaller)|
Dim Top As Integer ' Making a right integer (only to make the math pice look smaller)|
Top = (ManL.Top + ManL.Height / 2) - PointerL.Height / 2 ' Placing it in the middle of the "Man"
Left = (ManL.Left + ManL.Width / 2) - PointerL.Width / 2 ' Placing it in the middle of the "Man"
Top = Top + 200 * Sin(Degrees * (3.14 / 180)) ' Adds 200 to get it "Nearly" outside of the figure then some .. Simple MATH :P not ;)
Left = Left + 200 * Cos(Degrees * (3.14 / 180)) ' Adds 200 to get it "Nearly" outside of the figure then some. .. Simple MATH :P not ;)
PointerL.Left = Left ' The placing it to the final location we got
PointerL.Top = Top ' Then placing it to the final location we got

End Sub

Public Sub DirectionPointer()
Dim Left As Integer ' Making a left integer (only to make the math pice look smaller)|
Dim Top As Integer ' Making a right integer (only to make the math pice look smaller)|
Top = (Man.Top + Man.Height / 2) - Pointer.Height / 2 ' Placing it in the middle of the "Man"
Left = (Man.Left + Man.Width / 2) - Pointer.Width / 2 ' Placing it in the middle of the "Man"
Top = Top + 200 * Sin(Degrees * (3.14 / 180)) ' Adds 200 to get it "Nearly" outside of the figure then some .. Simple MATH :P not ;)
Left = Left + 200 * Cos(Degrees * (3.14 / 180)) ' Adds 200 to get it "Nearly" outside of the figure then some. .. Simple MATH :P not ;)
Pointer.Left = Left ' The placing it to the final location we got
Pointer.Top = Top ' Then placing it to the final location we got
End Sub


'This is just for fun, instead of making a collision detection for the
'walls i add a circle who shows u where you are :)
'i dident comment this becouse i dident think it was needed
'if u really need, add a collisiondetection =D
Public Sub OutsideShower()
Outside.Visible = True
If Pointer.Left < 0 Then
  Outside.Left = 0
  If Outside.Top > 0 Then Outside.Top = Pointer.Top Else Outside.Top = 1 '
  If Outside.Top + Outside.Height > Me.Height - 510 Then Outside.Top = Me.Height - 510 - Outside.Height
ElseIf Pointer.Top < 0 Then
  Outside.Top = 0
  If Outside.Left > 0 Then Outside.Left = Pointer.Left Else Outside.Left = 1
  If Outside.Left + Outside.Width > Me.Width Then Outside.Left = (Me.Width - 510)
ElseIf Pointer.Left > Me.Width Then
  Outside.Left = (Me.Width - 510)
  If Outside.Top > 0 Then Outside.Top = Pointer.Top Else Outside.Top = 1
  If Outside.Top + Outside.Height > Me.Height - 510 Then Outside.Top = Me.Height - 510 - Outside.Height
ElseIf Pointer.Top > (Me.Height - 510 - Outside.Height) Then
  Outside.Top = (Me.Height - 510 - Outside.Height)
  If Outside.Left > 0 Then Outside.Left = Pointer.Left Else Outside.Left = 1
  If Outside.Left + Outside.Width > Me.Width Then Outside.Left = (Me.Width - 510)
Else: Outside.Visible = False
End If
End Sub
