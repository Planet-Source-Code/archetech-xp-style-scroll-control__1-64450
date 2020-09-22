VERSION 5.00
Begin VB.UserControl TextScroller 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ScaleHeight     =   4110
   ScaleWidth      =   660
   Begin VB.Timer tmrEvent 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   -240
      Top             =   1800
   End
   Begin VB.TextBox t1 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "TextScroller.ctx":0000
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "TextScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* I must give credit to Pamela RAI's RichText Box Scroll Event submission
'That submission uses SendMessage call to detect the event
'I just adapted it to the textbox and modified it a bit to create this control.
'Feel free to improve upon this.  There is plenty of room for improvement, I just
'wanted to keep the code simple to get accross the general idea.
'As far as I am concerned, you can do whatever you want with this code since
'Microsoft has really done all the work
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Public Event Scroll()
Private glbIntValue As Integer
Private glbIntMax As Integer

Public Property Let Value(intValue As Integer)
   glbIntValue = intValue
End Property

Public Property Get Value() As Integer
   Value = glbIntValue
End Property

Public Property Let Max(intMax As Integer)
   glbIntMax = intMax
   SetMax
End Property

Public Property Get Max() As Integer
   Max = glbIntMax
End Property

'this exists essentially so the timer does not start until runtime
'there may be a better way
Public Sub Enable()
   tmrEvent.Enabled = True
End Sub

Public Sub Disable()
   tmrEvent.Enabled = False
End Sub

Private Sub UserControl_Initialize()
   glbIntMax = 100
   ResControl
   SetMax
End Sub

Private Sub SetMax()
   Dim i As Integer
   Dim intLines As Integer
   Dim arr() As String
   
   'this is the approx number of lines in the display of the textbox calculated using
   'an autosize label.  This is used so that any value will cause the scroll bar to display
   'on the text box
   intLines = (t1.Height / Label1.Height)
   
   'put into an array so you don't loose time concatinating strings in large values
   For i = 2 To (glbIntMax + intLines)
      ReDim Preserve arr(i)
      arr(i) = CStr(i) & vbCrLf
   Next i
   
   t1.Text = Join(arr)
End Sub

Private Sub tmrEvent_Timer()
   Static curLine As Integer, preLine As Integer
   
   'API Call returns the first visible line of the text box into the
   'variable holding the current line
   curLine = CInt(SendMessage(t1.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0))
   
   'If the first visible current line is different from the previous first visible line
   'the text box has been scrolled and the event should be raised
   If preLine <> curLine Then
      glbIntValue = curLine
      RaiseEvent Scroll
      UserControl.SetFocus
   End If
   
   'set the preline variable so it will hold the previous value the next time
   'the timer fires
   preLine = CInt(SendMessage(t1.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0))
End Sub

Private Sub ResControl()
   t1.Height = UserControl.Height
   t1.Left = -300
End Sub

Private Sub UserControl_Resize()
   ResControl
End Sub
