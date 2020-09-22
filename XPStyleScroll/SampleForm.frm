VERSION 5.00
Begin VB.Form SampleForm 
   Caption         =   "SampleForm"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin TextScroll.TextScroller TextScroller1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   360
      _extentx        =   635
      _extenty        =   8070
   End
   Begin VB.Label Label1 
      Caption         =   "Scroll Value:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblScrollValue 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "SampleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note, I suggest leaving a little room in the width of the control so that you
'can get select and move it

Private Sub Form_Load()
   TextScroller1.Max = 12
End Sub

Private Sub TextScroller1_GotFocus()
   'This enable method is set up so to turn the timer on.
   'I did this because I did not like the timer firing at design time
   TextScroller1.Enable
End Sub

Private Sub TextScroller1_Scroll()
   lblScrollValue.Caption = TextScroller1.Value
End Sub
