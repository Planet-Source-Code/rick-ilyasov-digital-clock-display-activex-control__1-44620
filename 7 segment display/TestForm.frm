VERSION 5.00
Object = "*\ADisplays.vbp"
Begin VB.Form Form1 
   Caption         =   "Digital Clock Display Fiesta!"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Run the Timers!"
      Height          =   555
      Left            =   30
      TabIndex        =   7
      Top             =   1560
      Width           =   7935
   End
   Begin Displays.DigitalDisplay ddBig 
      Height          =   495
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      DisplayColor    =   0
      DisplaySize     =   0
      NumberOfDigits  =   6
      NumberOfDecimals=   2
      Value           =   0
      BorderStyle     =   1
   End
   Begin Displays.DigitalDisplay ddCounter 
      Height          =   225
      Index           =   0
      Left            =   5340
      TabIndex        =   0
      Top             =   30
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   397
      DisplayColor    =   0
      DisplaySize     =   1
      NumberOfDigits  =   15
      NumberOfDecimals=   2
      Value           =   0
      BorderStyle     =   0
   End
   Begin Displays.DigitalDisplay ddCounter 
      Height          =   225
      Index           =   1
      Left            =   5340
      TabIndex        =   1
      Top             =   270
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   397
      DisplayColor    =   1
      DisplaySize     =   1
      NumberOfDigits  =   15
      NumberOfDecimals=   2
      Value           =   0
      BorderStyle     =   0
   End
   Begin Displays.DigitalDisplay ddCounter 
      Height          =   225
      Index           =   2
      Left            =   5340
      TabIndex        =   2
      Top             =   510
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   397
      DisplayColor    =   2
      DisplaySize     =   1
      NumberOfDigits  =   15
      NumberOfDecimals=   2
      Value           =   0
      BorderStyle     =   0
   End
   Begin Displays.DigitalDisplay ddCounter 
      Height          =   225
      Index           =   3
      Left            =   5340
      TabIndex        =   3
      Top             =   750
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   397
      DisplayColor    =   3
      DisplaySize     =   1
      NumberOfDigits  =   17
      NumberOfDecimals=   0
      Value           =   0
      BorderStyle     =   0
   End
   Begin Displays.DigitalDisplay ddCounter 
      Height          =   225
      Index           =   4
      Left            =   5340
      TabIndex        =   4
      Top             =   990
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   397
      DisplayColor    =   4
      DisplaySize     =   1
      NumberOfDigits  =   17
      NumberOfDecimals=   0
      Value           =   0
      BorderStyle     =   0
   End
   Begin Displays.DigitalDisplay ddCounter 
      Height          =   225
      Index           =   5
      Left            =   5340
      TabIndex        =   5
      Top             =   1230
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   397
      DisplayColor    =   5
      DisplaySize     =   1
      NumberOfDigits  =   17
      NumberOfDecimals=   0
      Value           =   0
      BorderStyle     =   0
   End
   Begin Displays.DigitalDisplay ddBig 
      Height          =   495
      Index           =   1
      Left            =   30
      TabIndex        =   8
      Top             =   510
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      DisplayColor    =   1
      DisplaySize     =   0
      NumberOfDigits  =   6
      NumberOfDecimals=   2
      Value           =   0
      BorderStyle     =   1
   End
   Begin Displays.DigitalDisplay ddBig 
      Height          =   495
      Index           =   2
      Left            =   30
      TabIndex        =   9
      Top             =   1020
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      DisplayColor    =   2
      DisplaySize     =   0
      NumberOfDigits  =   6
      NumberOfDecimals=   2
      Value           =   0
      BorderStyle     =   1
   End
   Begin Displays.DigitalDisplay ddBig 
      Height          =   495
      Index           =   3
      Left            =   2670
      TabIndex        =   10
      Top             =   0
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      DisplayColor    =   3
      DisplaySize     =   0
      NumberOfDigits  =   6
      NumberOfDecimals=   0
      Value           =   0
      BorderStyle     =   1
   End
   Begin Displays.DigitalDisplay ddBig 
      Height          =   495
      Index           =   4
      Left            =   2670
      TabIndex        =   11
      Top             =   510
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      DisplayColor    =   4
      DisplaySize     =   0
      NumberOfDigits  =   6
      NumberOfDecimals=   0
      Value           =   0
      BorderStyle     =   1
   End
   Begin Displays.DigitalDisplay ddBig 
      Height          =   495
      Index           =   5
      Left            =   2670
      TabIndex        =   12
      Top             =   1020
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   873
      DisplayColor    =   5
      DisplaySize     =   0
      NumberOfDigits  =   6
      NumberOfDecimals=   0
      Value           =   0
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Try double-clicking the Displays too!"
      Height          =   675
      Left            =   120
      TabIndex        =   13
      Top             =   2220
      Width           =   7665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'****************************************************************************************
'CREDITS:
'WRITTEN BY: RICK ILYASOV
'****************************************************************************************
'YOU ARE FREE TO USE THIS CONTROL ANYWAY YOU WANT, JUST LEAVE MY NAME IN THE CREDITS :)
'GOOD CODIN'!

                                             
Private Sub Command1_Click()

    Dim I As Long
    
    For I = 0 To 10000
        ddBig(0).Value = I
        ddCounter(0).Value = I
        DoEvents
    Next

End Sub

Private Sub ddBig_Change(Index As Integer)
    
    'This event is fired when the value of display changes
    On Error Resume Next
    ddBig(Index + 1).Value = ddBig(Index).Value / 2
    
End Sub

Private Sub ddCounter_Change(Index As Integer)
    
    'This event is fired when the value of display changes
    On Error Resume Next
    ddCounter(Index + 1).Value = ddCounter(Index).Value / 2

End Sub

Private Sub ddCounter_DblClick(Index As Integer)
    
    'This event shows how to change color of LEDs
    If ddCounter(Index).DisplayColor = 5 Then
        ddCounter(Index).DisplayColor = 0
    Else
        ddCounter(Index).DisplayColor = ddCounter(Index).DisplayColor + 1
    End If

End Sub

Private Sub ddTotal_DblClick()

    'This event shows how to change color of LEDs
    If ddTotal.DisplayColor = 5 Then
        ddTotal.DisplayColor = 0
    Else
        ddTotal.DisplayColor = ddTotal.DisplayColor + 1
    End If
    
End Sub

'THATS ALL FOLKS!

'THE DISPLAY CODE IS IN THE SECOND PROJECT OF THIS GROUP.
'DON'T FORGET TO VOTE IF YOU LIKE THIS ;p
