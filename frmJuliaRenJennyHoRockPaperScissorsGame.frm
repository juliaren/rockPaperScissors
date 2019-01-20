VERSION 5.00
Begin VB.Form frmJuliaRenJennyHoRockPaperScissorsGame 
   Caption         =   "Rock Paper Scissors Game"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScissors 
      Caption         =   "Scissors"
      Height          =   495
      Left            =   9480
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPaper 
      Caption         =   "Paper"
      Height          =   495
      Left            =   9480
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdRock 
      Caption         =   "Rock"
      Height          =   495
      Left            =   9480
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picUserScissors 
      AutoSize        =   -1  'True
      Height          =   2820
      Left            =   4680
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":0000
      ScaleHeight     =   2760
      ScaleWidth      =   3480
      TabIndex        =   14
      Top             =   1440
      Width           =   3540
   End
   Begin VB.PictureBox picUserPaper 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   4680
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":6B83
      ScaleHeight     =   3000
      ScaleWidth      =   3480
      TabIndex        =   13
      Top             =   1440
      Width           =   3540
   End
   Begin VB.PictureBox picUserRock 
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   2
      Left            =   4680
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":D9E3
      ScaleHeight     =   2670
      ScaleWidth      =   3480
      TabIndex        =   12
      Top             =   2760
      Width           =   3480
   End
   Begin VB.PictureBox picUserRock 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   1
      Left            =   4680
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":1404E
      ScaleHeight     =   3495
      ScaleWidth      =   3480
      TabIndex        =   11
      Top             =   1080
      Width           =   3480
   End
   Begin VB.PictureBox picPCPaper 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   480
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":1A6B9
      ScaleHeight     =   3000
      ScaleWidth      =   3480
      TabIndex        =   10
      Top             =   1440
      Width           =   3540
   End
   Begin VB.PictureBox picPCScissors 
      AutoSize        =   -1  'True
      Height          =   2820
      Left            =   480
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":214FC
      ScaleHeight     =   2760
      ScaleWidth      =   3480
      TabIndex        =   9
      Top             =   1440
      Width           =   3540
   End
   Begin VB.PictureBox picPCRock 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   2
      Left            =   480
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":280B0
      ScaleHeight     =   2670
      ScaleWidth      =   3480
      TabIndex        =   8
      Top             =   2640
      Width           =   3480
   End
   Begin VB.PictureBox picPCRock 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   1
      Left            =   480
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":2E5C2
      ScaleHeight     =   2670
      ScaleWidth      =   3480
      TabIndex        =   7
      Top             =   1080
      Width           =   3480
   End
   Begin VB.PictureBox picUserRock 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   0
      Left            =   4680
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":34AD4
      ScaleHeight     =   3375
      ScaleWidth      =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   3480
   End
   Begin VB.PictureBox picPCRock 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   0
      Left            =   480
      Picture         =   "frmJuliaRenJennyHoRockPaperScissorsGame.frx":3B13F
      ScaleHeight     =   2670
      ScaleWidth      =   3480
      TabIndex        =   5
      Top             =   1920
      Width           =   3480
   End
   Begin VB.Timer tmrRock 
      Interval        =   100
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End Game"
      Height          =   495
      Left            =   9600
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lblUser 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   19
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblPC 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   18
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label lblUserScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblTies 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblPCScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblInstruction 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmJuliaRenJennyHoRockPaperScissorsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer Name: Julia Ren & Jenny Ho
'Date: nov 30, 2015
'Course: ICS 2O1
'
'User Input: choose hand signal rock, paper or scissors
'Process: compare user input to computer generated random signal according to the rock paper scissors' rule
'Output: display result and record the scores

Option Explicit
Const intHighNum As Integer = 2
Const intLowNum As Integer = 0
Private intUserWins As Integer
Private intUserLosses As Integer
Private intPCWins As Integer
Private intPCLosses As Integer
Private intTies As Integer
Private intRound As Integer
Private intIncrement As Integer
Private intCount As Integer
Private intSignal As Integer

Private Sub cmdEndGame_Click()
    Unload Me
End Sub

Private Sub cmdPaper_Click()
    intCount = 0
    For intCount = 0 To 2
        picPCRock(intCount).Visible = False
    Next intCount
    
    picPCPaper.Visible = False
    picPCScissors.Visible = False
    
    intCount = 0
    For intCount = 0 To 2
        picUserRock(intCount).Visible = False
    Next intCount
    
    picUserScissors.Visible = False
    
    picUserPaper.Visible = True
    
    intSignal = Int((intHighNum - intLowNum + 1) * Rnd + intLowNum)
    
    If intSignal = 0 Then
        picPCRock(0).Visible = True
        MsgBox "You Win!"
        intUserWins = intUserWins + 1
        intPCLosses = intPCLosses + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserPaper.Visible = False
        picPCRock(0).Visible = False
    ElseIf intSignal = 1 Then
        picPCPaper.Visible = True
        MsgBox "You Are Tied!"
        intTies = intTies + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserPaper.Visible = False
        picPCPaper.Visible = False
    Else
        picPCScissors.Visible = True
        MsgBox "You Lost!"
        intUserLosses = intUserLosses + 1
        intPCWins = intPCWins + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserPaper.Visible = False
        picPCScissors.Visible = False
    End If
  
    If intRound = 10 Then
        Unload Me
    End If
End Sub

Private Sub cmdRock_Click()
    intCount = 0
    For intCount = 0 To 2
        picPCRock(intCount).Visible = False
    Next intCount
    
    picPCPaper.Visible = False
    picPCScissors.Visible = False
    
    intCount = 0
    For intCount = 0 To 2
        picUserRock(intCount).Visible = False
    Next intCount
    
    picUserPaper.Visible = False
    picUserScissors.Visible = False
    
    picUserRock(0).Visible = True
    
    intSignal = Int((intHighNum - intLowNum + 1) * Rnd + intLowNum)
    
    If intSignal = 0 Then
        picPCRock(0).Visible = True
        MsgBox "You Are Tied!"
        intTies = intTies + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserRock(0).Visible = False
        picPCRock(0).Visible = False
    ElseIf intSignal = 1 Then
        picPCPaper.Visible = True
        MsgBox "You Lost!"
        intUserLosses = intUserLosses + 1
        intPCWins = intPCWins + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserRock(0).Visible = False
        picPCPaper.Visible = False
    Else
        picPCScissors.Visible = True
        MsgBox "You Win!"
        intUserWins = intUserWins + 1
        intPCLosses = intPCLosses + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserRock(0).Visible = False
        picPCScissors.Visible = False
    End If

    If intRound = 10 Then
        Unload Me
    End If
End Sub

Private Sub cmdScissors_Click()
    intCount = 0
    For intCount = 0 To 2
        picPCRock(intCount).Visible = False
    Next intCount
    
    picPCPaper.Visible = False
    picPCScissors.Visible = False
    
    intCount = 0
    For intCount = 0 To 2
        picUserRock(intCount).Visible = False
    Next intCount
    
    picUserPaper.Visible = False
    
    picUserScissors.Visible = True
    
    intSignal = Int((intHighNum - intLowNum + 1) * Rnd + intLowNum)
    
    If intSignal = 0 Then
        picPCRock(0).Visible = True
        MsgBox "You Lost!"
        intUserLosses = intUserLosses + 1
        intPCWins = intPCWins + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserScissors.Visible = False
        picPCRock(0).Visible = False
    ElseIf intSignal = 1 Then
        picPCPaper.Visible = True
        MsgBox "You Win!"
        intUserWins = intUserWins + 1
        intPCLosses = intPCLosses + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserScissors.Visible = False
        picPCPaper.Visible = False
    Else
        picPCScissors.Visible = True
        MsgBox "You Are Tied!"
        intTies = intTies + 1
        lblPCScore.Caption = "Wins: " & intPCWins & vbCrLf & "Losses: " & intPCLosses
        lblTies.Caption = "Ties: " & intTies
        lblUserScore.Caption = "Wins: " & intUserWins & vbCrLf & "Losses: " & intUserLosses
        intRound = intRound + 1
        picUserScissors.Visible = False
        picPCScissors.Visible = False
    End If
    
    If intRound = 10 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Randomize
    
    intUserWins = 0
    intUserLosses = 0
    intPCWins = 0
    intPCLosses = 0
    intTies = 0
    intRound = 0
    intIncrement = 0
    intCount = 0
    
    lblInstruction.Caption = "Please select a hand signal to play with the computer.To end the game, press 'End Game'. The game will automatically end after 10 rounds."
    lblPC.Caption = "Computer"
    lblUser.Caption = "You"
    
    picPCScissors.Visible = False
    picPCPaper.Visible = False
    picUserScissors.Visible = False
    picUserPaper.Visible = False
End Sub

Private Sub tmrRock_Timer()
    intCount = 0
    Do While intCount < 3
        picPCRock(intCount).Visible = False
        intCount = intCount + 1
    Loop
        
    If intIncrement = 2 Then
        intIncrement = 0
    Else
        intIncrement = intIncrement + 1
    End If
           
    picPCRock(intIncrement).Visible = True
    
    intCount = 0
    Do While intCount < 3
        picUserRock(intCount).Visible = False
        intCount = intCount + 1
    Loop
    
    picUserRock(intIncrement).Visible = False
        
    If intIncrement = 2 Then
        intIncrement = 0
    Else
        intIncrement = intIncrement + 1
    End If
      
    picUserRock(intIncrement).Visible = True
End Sub
