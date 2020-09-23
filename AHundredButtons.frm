VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   ClientHeight    =   3636
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8256
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**A hundred buttons**

'If you need a lot of buttons on a form then you can
'fairly easily draw them on instead of
'using lots of resources with separate controls

'I know the buttons aren't very pretty - that's not the point of the
'code -- which is to provide simple base code you can adapt,
'prettier buttons, whatever

'Instead of buttons you could also adapt to draw option buttons,
'check boxes,labels, or a mix

'the code's straight forward so not many comments

'Yes, I should use option explicit etc and the code could be
'cleaned up a little - but I'm just naturally lazy and untidy

'jeremyxtz

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Const DT_SINGLELINE = &H20
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&
Const printformat = DT_SINGLELINE Or DT_CENTER Or DT_VCENTER Or DT_END_ELLIPSIS

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type buttontype
r As RECT
type As Integer
caption As String
End Type

Dim buttons() As buttontype

Private Enum buttonstate
statenormal = 1
statedown = 2
stateDisabled = 4
stateHasFocus = 8
End Enum

Dim selbutton As Integer

Dim btnheight As Integer
Dim btngap As Integer
Dim btnwidth As Integer

Private Sub Form_Load()
Me.AutoRedraw = True
Me.FillStyle = 0
setsomebuttons
drawbuttons
selbutton = 10 'choose an initial button or set to -1
drawbutton selbutton, buttonstate.stateHasFocus
End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
Select Case keycode
Case vbKeySpace, vbKeyReturn
If selbutton <> -1 Then drawbutton selbutton, buttonstate.statedown Or buttonstate.stateHasFocus
Case vbKeyTab
goNextNonDisabled
Case Else
goNextLetter keycode
End Select
End Sub
Sub goNextLetter(keycode)
nextletter = StrConv(Chr(keycode), vbUpperCase)
If selbutton = -1 Or selbutton = UBound(buttons) Then btnstart = 0 Else btnstart = selbutton + 1

For i = btnstart To UBound(buttons)
pos = InStr(1, buttons(i).caption, "&")
If pos <> 0 Then
If StrConv(Mid(buttons(i).caption, pos + 1, 1), vbUpperCase) = nextletter _
And Not buttons(i).type And buttonstate.stateDisabled Then
found = True
Exit For
End If
End If
Next

If btnstart <> 0 And found <> True Then
For i = 0 To btnstart
pos = InStr(1, buttons(i).caption, "&")
If pos <> 0 Then
If StrConv(Mid(buttons(i).caption, pos + 1, 1), vbUpperCase) = nextletter _
And Not buttons(i).type And buttonstate.stateDisabled Then
found = True
Exit For
End If
End If
Next
End If

If found <> True Then Exit Sub
If selbutton <> -1 Then drawbutton selbutton, buttonstate.statenormal
selbutton = i
drawbutton selbutton, buttonstate.statenormal Or buttonstate.stateHasFocus
End Sub


Sub goNextNonDisabled()
If selbutton = -1 Or selbutton = UBound(buttons) Then btnstart = 0 Else btnstart = selbutton + 1
For i = btnstart To UBound(buttons)
If Not buttons(i).type And buttonstate.stateDisabled Then
found = True
Exit For
End If
Next
If btnstart <> 0 And found <> True Then
For i = 0 To btnstart
If Not buttons(i).type And buttonstate.stateDisabled Then
found = True
Exit For
End If
Next
End If
If found <> True Then Exit Sub
If selbutton <> -1 Then drawbutton selbutton, buttonstate.statenormal
selbutton = i
drawbutton selbutton, buttonstate.statenormal Or buttonstate.stateHasFocus

End Sub


Private Sub Form_KeyUp(keycode As Integer, Shift As Integer)
Select Case keycode
Case vbKeySpace, vbKeyReturn
If selbutton <> -1 Then
buttonAction selbutton
drawbutton selbutton, buttonstate.statenormal Or buttonstate.stateHasFocus
End If
End Select

End Sub


Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Me.caption = ""
btn = getbutton(X, Y)
If btn <> -1 Then
If buttons(btn).type And buttonstate.stateDisabled Then Exit Sub
If btn <> selbutton And selbutton <> -1 Then
drawbutton selbutton, buttonstate.statenormal
End If
selbutton = btn
drawbutton selbutton, buttonstate.statedown Or buttonstate.stateHasFocus
End If
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
btn = getbutton(X, Y)
If btn = selbutton And selbutton <> -1 Then
buttonAction btn
Else
Me.caption = ""
End If
If selbutton <> -1 Then drawbutton selbutton, buttonstate.statenormal Or buttonstate.stateHasFocus

End Sub
Sub drawbuttons()
For i = 0 To UBound(buttons)
drawbutton i
Next
End Sub
Sub drawbutton(btn, Optional state = 1)
'a simple button - you can modify for something fancier
If state And buttonstate.statedown Then
Me.FillColor = vbBlue
ElseIf state And buttonstate.stateHasFocus Then
Me.FillColor = RGB(200, 200, 200)
Else
Me.FillColor = vbButtonFace
End If

btnstr = buttons(btn).caption
If buttons(btn).type And buttonstate.stateDisabled Then
Me.ForeColor = &H808080
Else
Me.ForeColor = vbBlack
End If

With buttons(btn)
With .r
Rectangle Me.hdc, .Left, .Top, .Right, .Bottom
If state And 1 Then Rectangle Me.hdc, .Left, .Top, .Right - 1, .Bottom - 1
End With

If state And buttonstate.statedown Then Me.ForeColor = vbWhite
DrawText Me.hdc, btnstr, Len(btnstr), .r, printformat

If state And buttonstate.stateHasFocus Then
Me.ForeColor = vbRed
Dim r As RECT
r = buttons(btn).r
InflateRect r, -3, -3
DrawFocusRect Me.hdc, r
End If

End With
Me.Refresh
End Sub
Function getbutton(X, Y)
'Buttons are nicely ordered so we could do things more efficiently
'by using the x,y to get the button directly
'(especially if we were using the mousemove event)
'however I'm just keeping the code simple
'and this is how we'd find buttons if they were in the
'more likely disordered arangement
For i = 0 To UBound(buttons)
If PtInRect(buttons(i).r, X, Y) Then getbutton = i: Exit Function Else getbutton = -1
Next
End Function
Sub setsomebuttons()
'create some buttons
'you'd change this completely to suit

'choose how many buttons
ReDim buttons(99)

'set the captions
For i = 0 To UBound(buttons)
buttons(i).caption = "&Button " & i + 1
Next

'set some different values
buttons(3).type = buttonstate.stateDisabled
buttons(0).caption = "F&ish"

'get biggest
For i = 0 To UBound(buttons)
newwidth = Me.TextWidth(buttons(i).caption)
If newwidth > btnwidth Then btnwidth = newwidth
Next
btnwidth = btnwidth + 20 'add a bit

'decide how many rows
buttonrows = 15

'set button height
btnheight = Me.TextHeight("") * 2

'set gap between buttons
btngap = 5

'set button positions
For i = 0 To UBound(buttons)
If i Mod buttonrows = 0 Then
c = 0
If i <> 0 Then lft = lft + btnwidth + btngap
cols = cols + 1
Else
c = c + 1
End If
With buttons(i).r
.Top = (c * (btngap + btnheight)) + btngap
.Left = lft
.Right = lft + btnwidth
.Bottom = .Top + btnheight
End With

Next

'set form size
'you'd really need to get the form border/caption sizes using api
'to be exact but very approximately...
Me.Width = (((cols) * btnwidth) + ((cols + 1) * btngap)) * Screen.TwipsPerPixelX
Me.Height = (buttonrows + 1) * (btnheight + btngap) * Screen.TwipsPerPixelY
End Sub

Sub buttonAction(btn)
'you'd probably do a select case here
Me.caption = buttons(btn).caption & " pressed"
End Sub
