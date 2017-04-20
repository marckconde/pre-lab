VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Bolita2.0.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6480
      Top             =   2400
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6480
      Top             =   1920
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   6480
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4080
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4080
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   4080
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detener"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   1920
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   0
      Shape           =   3  'Circle
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Shape1.Top <= 4400 Then
Shape1.Top = Shape1.Top + 100
Else
Shape1.Left = Shape1.Left + 100
End If
If Shape1.Left >= 840 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If Shape1.Top >= 120 Then
Shape1.Top = Shape1.Top - 100
Else
Shape1.Left = Shape1.Left + 100
End If
If Shape1.Left >= 1680 Then
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
If Shape1.Top <= 4400 Then
Shape1.Top = Shape1.Top + 100
Else
Shape1.Left = Shape1.Left + 100
End If
If Shape1.Left >= 2520 Then
Timer3.Enabled = False
Timer4.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
If Shape1.Top >= 120 Then
Shape1.Top = Shape1.Top - 100
Else
Shape1.Left = Shape1.Left + 100
End If
If Shape1.Left >= 3360 Then
Timer4.Enabled = False
Timer5.Enabled = True
End If
End Sub

Private Sub Timer5_Timer()
If Shape1.Top <= 4400 Then
Shape1.Top = Shape1.Top + 100
Else
Shape1.Left = Shape1.Left + 100
End If
If Shape1.Left >= 4200 Then
Timer5.Enabled = False
Timer6.Enabled = True
End If
End Sub

Private Sub Timer6_Timer()
If Shape1.Top >= 120 Then
Shape1.Top = Shape1.Top - 100
Else
Shape1.Left = Shape1.Left + 100
End If
If Shape1.Left >= 5040 Then
Timer6.Enabled = False
End If
End Sub
