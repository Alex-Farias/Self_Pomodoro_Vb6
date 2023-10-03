VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   2400
   End
   Begin VB.CommandButton btn_pomodoro 
      Caption         =   "Começar!"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lbl_pomodoro 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timer As String
Dim min, sec As Integer
Dim activate As Boolean

Private Function activateTimer(activated As Boolean)
    If activated = True Then
        activate = False
    Else
        activate = True
    End If
End Function
Private Sub btn_pomodoro_Click()
    If btn_pomodoro.Caption = "Começar!" Then
        btn_pomodoro.Caption = "Parar!"
    Else
        btn_pomodoro.Caption = "Começar!"
    End If
    
    activateTimer (activate)
End Sub

Private Sub Timer1_Timer()
    If activate = True Then
        sec = sec + 1
        
        If sec = 60 Then
            min = min + 1
            sec = 0
            
            If min = 25 Then
                min = 0
                
            End If
        End If
        
        If Format(min, "00") = "" Then
            min = 0
            timer = "00"
        Else
            timer = Format(min, "00")
        End If
        
        timer = timer + ":"
        
        If Format(sec, "00") = "" Then
            timer = timer + "00"
        Else
            timer = timer + Format(sec, "00")
        End If
        
        lbl_pomodoro.Caption = timer
    End If
End Sub
