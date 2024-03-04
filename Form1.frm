VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Self Pomodoro"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   2820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_pausar 
      Caption         =   "Pausar!"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin VB.CommandButton btn_pomodoro 
      Caption         =   "Começar!"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lbl_pomodoro 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const COMECAR As String = "Começar!"
Const PARAR As String = "Parar!"
Const DESCANCAR As String = "Descançar!"
Const PAUSAR As String = "Pausar!"
Const RETOMAR As String = "Retomar!"
Const MUSIC_PATH As String = "\Audio\Hit_Pomodoro.wav"
Const MAX_MIN_ACTIVE As Integer = 25
Const MAX_MIN_REST As Integer = 5
Const JUMP_SEC As Integer = 1  '1|30
Const JUMP_MIN As Integer = 1 '1|5

Dim caption_index, timer_sec, timer_min As String
Dim max_min, min, sec As Integer
Dim activate As Boolean
Dim rest As Boolean

Private Declare Function playSound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long

Private Sub Form_Load()
    min = 0
    rest = True
    Call activateTimer(True, True)
End Sub

Public Function playFinalSound(path As String)
    playSound App.path & path, 0, 0 'Funciona apenas com tipos de arquivo wav
End Function

Public Function activateTimer(activated As Boolean, rested As Boolean)
    If activated = True Then
        If rested = True Then
            activateRest (rest)
        End If
        
        activate = False
        btn_pomodoro.Caption = caption_index
    Else
        activate = True
    End If
End Function

Public Function activateRest(rested As Boolean)
    If rested = False Then
        rest = True
        max_min = MAX_MIN_REST
        caption_index = DESCANCAR
    Else
        rest = False
        max_min = MAX_MIN_ACTIVE
        caption_index = COMECAR
    End If
    
End Function

Private Sub btn_pomodoro_Click()
    If btn_pomodoro.Caption <> PARAR Then
        btn_pomodoro.Caption = PARAR
        Call activateTimer(activate, True)
    Else
        sec = 0
        min = 0
        Call activateTimer(activate, False)
        btn_pomodoro.Caption = caption_index
    End If
End Sub

Private Sub btn_pausar_Click()
    Call activateTimer(activate, False)
    
    If btn_pausar.Caption = PAUSAR Then
        btn_pausar.Caption = RETOMAR
    Else
        btn_pausar.Caption = PAUSAR
    End If
    
End Sub

Private Sub Timer1_Timer()
    Dim Retval As Long
    
    If activate = True Then
        sec = sec + JUMP_SEC
        
        If sec = 60 Then
            min = min + JUMP_MIN
            sec = 0
            
            If min = max_min Then
                min = 0
                
                playFinalSound (MUSIC_PATH)
                Call activateTimer(activate, True)
            End If

        End If
        
        If Format(sec, "00") = "" Or IsNull(Format(sec, "00")) Then
            timer_sec = "00"
        Else
            timer_sec = Format(sec, "00")
        End If
        
        If Format(min, "00") = "" Or IsNull(Format(min, "00")) Then
            timer_min = "00"
        Else
            timer_min = Format(min, "00")
        End If
        
        lbl_pomodoro.Caption = timer_min + ":" + timer_sec
    End If
End Sub
