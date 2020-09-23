VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   LinkTopic       =   "Form2"
   MousePointer    =   15  'Size All
   ScaleHeight     =   3195
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "THIS FORM CAN DOCK WITH SCREEN ENDS AND THE MAIN FORM"
      Height          =   735
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "THIS IS ANOTHER FORM THAT DOCKS WITH THE MAIN FORM."
      Height          =   495
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MoveForm As Boolean
Dim MouseX As Long, MouseY As Long
Dim PresentX As Long, PresentY As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveForm = True
MouseX = X
MouseY = Y
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveForm Then
PresentX = Me.Left - MouseX + X
PresentY = Me.Top - MouseY + Y
Me.Move PresentX, PresentY
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm = False
If Form2.WindowState = vbNormal Then
    If Form2.Top + Form2.Height > Screen.Height - 500 Then
    Form2.Top = (Screen.Height - Form2.Height) - 1
    ElseIf Form2.Top < 500 Then
    Form2.Top = 0 + 1
    End If
    If Form2.Left + Form2.Width > Screen.Width - 500 Then
    Form2.Left = (Screen.Width - Form2.Width) - 1
    ElseIf Form2.Left < 500 Then
    Form2.Left = 0 + 1
    End If
    
    If Form2.Left > (Form1.Left + Form1.Width) Then
        If Form2.Left < (Form1.Left + Form1.Width + 500) Then
            Form2.Left = Form1.Left + Form1.Width
        End If
    ElseIf (Form2.Left + Form2.Width) < Form1.Left Then
        If (Form2.Left + Form2.Width) > (Form1.Left - 500) Then
            Form2.Left = Form1.Left - Form2.Width
        End If
    End If
    If Form2.Top > (Form1.Top + Form1.Height) Then
        If Form2.Top < (Form1.Top + Form1.Height + 500) Then
            Form2.Top = Form1.Top + Form1.Height
        End If
    ElseIf (Form2.Top + Form2.Height) < Form1.Top Then
        If (Form2.Top + Form2.Height) > (Form1.Top - 500) Then
            Form2.Top = Form1.Top - Form2.Height
        End If
    End If
    
End If
End Sub

Private Sub Timer1_Timer()
If Form2.WindowState = vbNormal Then
    If Form2.Top + Form2.Height > Screen.Height - 500 Then
    Form2.Top = (Screen.Height - Form2.Height) - 1
    ElseIf Form2.Top < 500 Then
    Form2.Top = 0 + 1
    End If
    If Form2.Left + Form2.Width > Screen.Width - 500 Then
    Form2.Left = (Screen.Width - Form2.Width) - 1
    ElseIf Form2.Left < 500 Then
    Form2.Left = 0 + 1
    End If
End If
End Sub
