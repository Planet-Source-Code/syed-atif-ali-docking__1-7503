VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MousePointer    =   15  'Size All
   ScaleHeight     =   5070
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3600
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Docking to screen ends..."
      ForeColor       =   &H00C0C0FF&
      Height          =   1335
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   720
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dock when user releases mouse button."
         Height          =   375
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dock with a timer."
         Height          =   375
         Left            =   120
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "YOU COULD DOCK THE OTHER SMALL FORM WITH THIS ONE!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   $"real.frx":0000
      Height          =   855
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "You can move this form by just holding it from anywhere (except from the controls) and dragging it."
      Height          =   495
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|--------------------------------------|
'|--------------------------------------|
'|--------------------------------------|
'HUTELL (HUman inTELLigence) Productions'
'|--------------------------------------|
'|------------HELLO EVERYONE!-----------|
'|--------------------------------------|
'|--OKAY,THIS CODE IS SPECIFICALLY FOR--|
'|----DOCKING OR WHAT YOU WOULD CALL----|
'|-"SNAPPING".IT ALSO SHOWS HOW TO DOCK-|
'|-------FORMS WITH OTHER FORMS!--------|
'|--------------------------------------|
'|--------------------------------------|
'|--BY SYED ATIF ALI (owner of HUTELL)--|
'|----EMAIL ME AT hutell@hotmail.com----|
'|--------------------------------------|
'HUTELL (HUman inTELLigence) Productions'
'|--------------------------------------|
'|--------------------------------------|
'|--------------------------------------|

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'NOTES: the form2 coding does not include any comments
'because that code is just a duplicate of the form1's
'so I suggest that you look into this coding only.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'WARNING: I, Syed Atif Ali, should and could not be held
'responsible or liable for any kind of damages or anything
'bad arising from the use of this. That means, use it at
'your own risk.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'DISCLAIMER: You could, without permission, include this
'code in your non-commercial applications given that you
'mention me and my company somewhere in the credits
'section of your non-commercial software. However, you
'DO need to obtain written permission of mine before you
'include this code within your commercial software. YOU
'COULD NOT(!), however, reproduce, publicise, or throw
'this code (into a web-site) without the given permission
'of Syed Atif Ali.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'COPYRIGHT: This source code is a copyright of Syed Atif
'Ali of HUTELL© Productions.
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\



Option Explicit
Dim MoveForm As Boolean
Dim MouseX As Long, MouseY As Long
Dim PresentX As Long, PresentY As Long

Private Sub Command1_Click()
Unload Form2
Unload Me
End Sub

Private Sub Form_Load()
Form2.Show vbModeless, Me
'^show the other form^
Me.Top = 0
Me.Left = 0
Form2.Top = 0
Form2.Left = Me.Width
'^set their positions^
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveForm = True
'^Moveform is a kind of permission for the MouseMove...
'to work or not. This gives the permission.^
MouseX = X
MouseY = Y
'^Set their values to the current mouse positions^
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MoveForm Then
'If it has permission...^
PresentX = Me.Left - MouseX + X
PresentY = Me.Top - MouseY + Y
'^Set the position where the form is going to be moved.^
Me.Move PresentX, PresentY
'^Move the form.^
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm = False
'^Switch off the permission to move the from
If Form1.WindowState = vbNormal Then
'^you can't move a form while in Max. or Min. mode.^
    If Form1.Top + Form1.Height > Screen.Height - 500 Then
    'if the bottom end of the form is bigger than the...
    'bottom of the screen - 500 then^
    Form1.Top = (Screen.Height - Form1.Height) - 1
    '^dock the form to the bottom screen.^
    ElseIf Form1.Top < 500 Then
    '^if the top of the form is smaller than the screen...
    'top - 500, then...^
    Form1.Top = 0 + 1
    '^dock the form to the top of the screen.^
    End If
    If Form1.Left + Form1.Width > Screen.Width - 500 Then
    '^if the right end of the form is bigger than the...
    'right end of the screen - 500 then...^
    Form1.Left = (Screen.Width - Form1.Width) - 1
    '^dock the form to the right end of the screen.^
    ElseIf Form1.Left < 500 Then
    'if the left end of the form is smaller than the screen...
    'left - 500, then...^
    Form1.Left = 0 + 1
    '^dock the form to the left end of the screen.^
    End If
    
    If Form2.Left > (Form1.Left + Form1.Width) Then
    '^if the form2's left end is bigger than the right end...
    'of the form1 then...^
        If Form2.Left < (Form1.Left + Form1.Width + 500) Then
        '^but if the form2's left is smaller than the right...
        'end of the form1 - 500 then...^
            Form2.Left = Form1.Left + Form1.Width
            '^dock the form2 to the left end of form1.^
        End If
    ElseIf (Form2.Left + Form2.Width) < Form1.Left Then
    '^if the form2's right end is bigger than the left end...
    'of the form1, then...^
        If (Form2.Left + Form2.Width) > (Form1.Left - 500) Then
        '^but if the form2's right is smaller than the left...
        'end of the form1 - 500 then...^
            Form2.Left = Form1.Left - Form2.Width
            '^dock the form2 to the right end of form1.^
        End If
    End If
    If Form2.Top > (Form1.Top + Form1.Height) Then
    '^if the form2's top end is bigger than the bottom end...
    'of the form1, then...^
        If Form2.Top < (Form1.Top + Form1.Height + 500) Then
        '^but if the form2's top is smaller than the bottom...
        'end of the form1 - 500 then...^
            Form2.Top = Form1.Top + Form1.Height
            '^dock the form2 to the bottom end of form1.^
        End If
    ElseIf (Form2.Top + Form2.Height) < Form1.Top Then
    '^if the form2's bottom end is bigger than the top end...
    'of the form1, then...^
        If (Form2.Top + Form2.Height) > (Form1.Top - 500) Then
        '^but if the form2's bottom is smaller than the top...
        'end of the form1 - 500 then...^
            Form2.Top = Form1.Top - Form2.Height
            '^dock the form2 to the top end of form1.^
        End If
    End If
End If
End Sub

Private Sub Option1_Click()
Timer1.Enabled = False
Form2.Timer1.Enabled = False
'^disable the timers
End Sub

Private Sub Option2_Click()
Timer1.Enabled = True
Form2.Timer1.Enabled = True
'enable the timers
End Sub

Private Sub Timer1_Timer()
If Form1.WindowState = vbNormal Then
'^you can't move a form while in Max. or Min. mode.^
    If Form1.Top + Form1.Height > Screen.Height - 500 Then
    'if the bottom end of the form is bigger than the...
    'bottom of the screen - 500 then^
    Form1.Top = (Screen.Height - Form1.Height) - 1
    '^dock the form to the bottom screen.^
    ElseIf Form1.Top < 500 Then
    '^if the top of the form is smaller than the screen...
    'top - 500, then...^
    Form1.Top = 0 + 1
    '^dock the form to the top of the screen.^
    End If
    If Form1.Left + Form1.Width > Screen.Width - 500 Then
    '^if the right end of the form is bigger than the...
    'right end of the screen - 500 then...^
    Form1.Left = (Screen.Width - Form1.Width) - 1
    '^dock the form to the right end of the screen.^
    ElseIf Form1.Left < 500 Then
    'if the left end of the form is smaller than the screen...
    'left - 500, then...^
    Form1.Left = 0 + 1
    '^dock the form to the left end of the screen.^
    End If
End Sub
