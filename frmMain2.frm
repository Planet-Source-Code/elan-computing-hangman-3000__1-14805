VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hangman"
   ClientHeight    =   7410
   ClientLeft      =   3585
   ClientTop       =   2130
   ClientWidth     =   7080
   HasDC           =   0   'False
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   1605
      TabIndex        =   31
      Top             =   1305
      Width           =   3870
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   45
         TabIndex        =   43
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   42
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   660
         TabIndex        =   41
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   975
         TabIndex        =   40
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1275
         TabIndex        =   39
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1590
         TabIndex        =   38
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   1905
         TabIndex        =   37
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2205
         TabIndex        =   36
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   2520
         TabIndex        =   35
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   2820
         TabIndex        =   34
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   3135
         TabIndex        =   33
         Top             =   195
         Width           =   270
      End
      Begin VB.Label lblWord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   3450
         TabIndex        =   32
         Top             =   195
         Width           =   270
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   975
      Top             =   495
   End
   Begin VB.Label lblCategory 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1920
      TabIndex        =   30
      Top             =   1140
      Width           =   3255
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   90
      TabIndex        =   29
      Top             =   7110
      Width           =   6900
   End
   Begin VB.Label lblHowLong 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   90
      TabIndex        =   28
      Top             =   6420
      Width           =   6900
   End
   Begin VB.Label lblRemaining 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   90
      TabIndex        =   27
      Top             =   6750
      Width           =   6900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Letters Remaining:"
      Height          =   195
      Left            =   1920
      TabIndex        =   26
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      Index           =   3
      X1              =   2970
      X2              =   2415
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line4 
      BorderWidth     =   4
      Index           =   2
      X1              =   3705
      X2              =   3705
      Y1              =   2790
      Y2              =   2115
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      Index           =   1
      X1              =   4635
      X2              =   2460
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      Index           =   0
      X1              =   2415
      X2              =   2415
      Y1              =   5955
      Y2              =   2115
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      Height          =   195
      Index           =   25
      Left            =   4935
      TabIndex        =   25
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   195
      Index           =   24
      Left            =   4680
      TabIndex        =   24
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Index           =   23
      Left            =   4440
      TabIndex        =   23
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      Height          =   195
      Index           =   22
      Left            =   4170
      TabIndex        =   22
      Top             =   750
      Width           =   165
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      Height          =   195
      Index           =   21
      Left            =   3975
      TabIndex        =   21
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      Height          =   195
      Index           =   20
      Left            =   3735
      TabIndex        =   20
      Top             =   750
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   195
      Index           =   19
      Left            =   3495
      TabIndex        =   19
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   195
      Index           =   18
      Left            =   3270
      TabIndex        =   18
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   195
      Index           =   17
      Left            =   3030
      TabIndex        =   17
      Top             =   750
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      Height          =   195
      Index           =   16
      Left            =   2790
      TabIndex        =   16
      Top             =   750
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   195
      Index           =   15
      Left            =   2535
      TabIndex        =   15
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      Height          =   195
      Index           =   14
      Left            =   2310
      TabIndex        =   14
      Top             =   750
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      Height          =   195
      Index           =   13
      Left            =   2085
      TabIndex        =   13
      Top             =   750
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   195
      Index           =   12
      Left            =   4920
      TabIndex        =   12
      Top             =   420
      Width           =   135
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   11
      Top             =   420
      Width           =   90
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      Height          =   195
      Index           =   10
      Left            =   4440
      TabIndex        =   10
      Top             =   420
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      Height          =   195
      Index           =   9
      Left            =   4215
      TabIndex        =   9
      Top             =   420
      Width           =   75
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   195
      Index           =   8
      Left            =   4005
      TabIndex        =   8
      Top             =   420
      Width           =   45
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      Height          =   195
      Index           =   7
      Left            =   3735
      TabIndex        =   7
      Top             =   420
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   195
      Index           =   6
      Left            =   3495
      TabIndex        =   6
      Top             =   420
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   195
      Index           =   5
      Left            =   3270
      TabIndex        =   5
      Top             =   420
      Width           =   90
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   195
      Index           =   4
      Left            =   3030
      TabIndex        =   4
      Top             =   420
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   195
      Index           =   3
      Left            =   2790
      TabIndex        =   3
      Top             =   420
      Width           =   120
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   195
      Index           =   2
      Left            =   2535
      TabIndex        =   2
      Top             =   420
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   195
      Index           =   1
      Left            =   2310
      TabIndex        =   1
      Top             =   420
      Width           =   105
   End
   Begin VB.Label lblLetters 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   195
      Index           =   0
      Left            =   2085
      TabIndex        =   0
      Top             =   420
      Width           =   105
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1965
      X2              =   5130
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line LegLeft 
      BorderWidth     =   2
      Index           =   1
      X1              =   3180
      X2              =   3375
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line LegRight 
      BorderWidth     =   2
      Index           =   1
      X1              =   4245
      X2              =   4440
      Y1              =   5655
      Y2              =   5655
   End
   Begin VB.Line ArmRight 
      BorderWidth     =   2
      Index           =   1
      X1              =   4395
      X2              =   4560
      Y1              =   3960
      Y2              =   3900
   End
   Begin VB.Line ArmRight 
      BorderWidth     =   2
      Index           =   2
      X1              =   4440
      X2              =   4605
      Y1              =   3960
      Y2              =   4005
   End
   Begin VB.Line ArmRight 
      BorderWidth     =   2
      Index           =   3
      X1              =   4410
      X2              =   4515
      Y1              =   3960
      Y2              =   4080
   End
   Begin VB.Line ArmLeft 
      BorderWidth     =   2
      Index           =   3
      X1              =   2850
      X2              =   2985
      Y1              =   4260
      Y2              =   4215
   End
   Begin VB.Line ArmLeft 
      BorderWidth     =   2
      Index           =   2
      X1              =   2835
      X2              =   3000
      Y1              =   4155
      Y2              =   4200
   End
   Begin VB.Line ArmLeft 
      BorderWidth     =   2
      Index           =   1
      X1              =   2880
      X2              =   2985
      Y1              =   4050
      Y2              =   4170
   End
   Begin VB.Line LegLeft 
      BorderWidth     =   2
      Index           =   0
      X1              =   3150
      X2              =   3705
      Y1              =   5640
      Y2              =   5190
   End
   Begin VB.Line LegRight 
      BorderWidth     =   2
      Index           =   0
      X1              =   3720
      X2              =   4230
      Y1              =   5175
      Y2              =   5655
   End
   Begin VB.Line ArmLeft 
      BorderWidth     =   2
      Index           =   0
      X1              =   3015
      X2              =   3690
      Y1              =   4200
      Y2              =   4350
   End
   Begin VB.Line ArmRight 
      BorderWidth     =   2
      Index           =   0
      X1              =   3735
      X2              =   4365
      Y1              =   4350
      Y2              =   3960
   End
   Begin VB.Line MyBody 
      BorderWidth     =   4
      X1              =   3720
      X2              =   3720
      Y1              =   3855
      Y2              =   5160
   End
   Begin VB.Shape EyeLeft 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   1
      Left            =   3510
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   90
   End
   Begin VB.Shape EyeRight 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   90
      Index           =   1
      Left            =   3915
      Shape           =   2  'Oval
      Top             =   3120
      Width           =   90
   End
   Begin VB.Shape MyNose 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   3645
      Shape           =   2  'Oval
      Top             =   3270
      Width           =   135
   End
   Begin VB.Shape MyMouth 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   210
      Left            =   3510
      Shape           =   2  'Oval
      Top             =   3510
      Width           =   405
   End
   Begin VB.Shape EyeRight 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   210
      Index           =   0
      Left            =   3750
      Shape           =   2  'Oval
      Top             =   3060
      Width           =   315
   End
   Begin VB.Shape EyeLeft 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   210
      Index           =   0
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   3060
      Width           =   315
   End
   Begin VB.Shape MyHead 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   3030
      Shape           =   3  'Circle
      Top             =   2820
      Width           =   1410
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   630
      Left            =   1935
      Top             =   390
      Width           =   3225
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strUsername As String ' Username of player
Dim strWord As String ' Word being shown
Dim strWordWorking As String ' Word with only letters guessed

Dim iPlayed As Integer ' Number of times user has played
Dim iTime As Integer ' Time played this round

Dim iWrong As Integer ' Guesses wrong per word
Dim iTotalRight As Integer, iTotalWrong As Integer 'Total Right/Wrong

Dim IsEnabled As Boolean 'Can user guess a letter?

Dim Cn As Connection
Dim X As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim iLetter As Integer
    Dim iPlace As Integer
    
    If KeyAscii = 27 Then
        WindowState = vbMinimized
        Exit Sub
    End If
    
    'Already working on a letter
    If IsEnabled = False Then Exit Sub
    
    IsEnabled = False
    
    'Get letter pressed as int
    iLetter = Asc(UCase(Chr(KeyAscii))) - 65
    
    'Not a letter
    If iLetter < 0 Or iLetter > 25 Then
        IsEnabled = True
        Exit Sub
    End If
    
    'Letter already used
    If lblLetters(iLetter).Visible = False Then
        Beep
        IsEnabled = True
        Exit Sub
    End If
    
    'Hide letter
    lblLetters(iLetter).Visible = False
       
    iPlace = InStr(iPlace + 1, strWord, Chr(KeyAscii), vbTextCompare)
    'Wrong letter
    If iPlace = 0 Then
        WrongLetter
        IsEnabled = True
        Exit Sub
    End If
        
    'Put letters on board
    iPlace = 0
    For X = 1 To Len(strWord)
        iPlace = InStr(iPlace + 1, strWord, Chr(KeyAscii), vbTextCompare)
               
        'All letters filled in
        If iPlace = 0 Then Exit For
        
        'Right letter
        If lblWord(iPlace - 1) = "_" Then
            Call PutLetter(iLetter, iPlace - 1)
        End If
    Next
    
    IsEnabled = True
    
    'Check for word completion
    If strWordWorking = String(Len(strWord), " ") Then WordDone
End Sub

Private Sub Form_Load()
    Set Cn = New Connection
    Dim Rs As New Recordset
    
    ChDir (App.Path)
    Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Hangman.mdb"
        
    Rs.Open "[User]", Cn, adOpenKeyset
    Rs.MoveFirst
    
    strUsername = InputBox("Please enter your name", "Username", Rs.Fields("Username"))
    
    If strUsername = "" Then
        Rs.Close
        Set Rs = Nothing
        
        Unload Me
        Exit Sub
    End If
        
    Rs.Find "UserName = '" & strUsername & "'"
    
    If Rs.EOF Or Rs.BOF Then
        Cn.Execute ("INSERT INTO [User] (UserName) VALUES ('" & strUsername & "')")
        Rs.Requery
        Rs.Find "UserName = '" & strUsername & "'"
    End If
    
    iTotalRight = Rs.Fields("Right")
    iTotalWrong = Rs.Fields("Wrong")
    iPlayed = Rs.Fields("Played")
    iTime = Rs.Fields("Time")
            
    Rs.Close
    Set Rs = Nothing
    
    Caption = "Hangman   -   " & strUsername
    
    Randomize
    GetWord
End Sub

Private Sub PutLetter(iLetter As Integer, iPlace As Integer)
    Dim y As Long
    
    For y = 0 To iLetter
        lblWord(iPlace) = Chr(y + 65)
        Pause (0.01)
    Next
    
    lblWord(iPlace) = Chr(iLetter + 65)
    
    Mid(strWordWorking, iPlace + 1) = " "
End Sub

Private Sub GetWord()
    Dim Rs As Recordset, Rs2 As Recordset
    Dim sRecordCount As Long
    Dim sCount As Long ' Number of attempts to get a word not used
    
    Set Rs = New Recordset
    Set Rs2 = New Recordset
       
    Rs.Open "Words", Cn, adOpenStatic
    Rs2.Open "Category", Cn, adOpenForwardOnly
    
    Rs.MoveLast
    
    'Find word that hasn't been used
    sRecordCount = Rs.RecordCount - 1
    Do
        sCount = sCount + 1
       
        X = Int((sRecordCount + 1) * Rnd)
        
        'All words used, so reset
        If sCount = sRecordCount Then Cn.Execute ("UPDATE Words SET Used = False")
        
        Call Rs.Move(X, 1)
    Loop Until Rs.Fields(2) = False
    
    strWord = Rs.Fields(1)
    
    Rs2.Move (Rs.Fields(3) - 1)
    lblCategory = UCase(Rs2.Fields(1))
    Cn.Execute ("UPDATE Words Set Used = True WHERE Word = '" & strWord & "'")
        
    Rs.Close
    Rs2.Close
    
    Set Rs = Nothing
    Set Rs2 = Nothing
    
    'Get word from database
    strWordWorking = strWord
    lblHowLong = Len(strWord) & " Letter Word"
    
    'Reset hangman
    LegLeft(0).Visible = False
    LegLeft(1).Visible = False
    LegRight(0).Visible = False
    LegRight(1).Visible = False
    ArmLeft(0).Visible = False
    ArmLeft(1).Visible = False
    ArmLeft(2).Visible = False
    ArmLeft(3).Visible = False
    ArmRight(0).Visible = False
    ArmRight(1).Visible = False
    ArmRight(2).Visible = False
    ArmRight(3).Visible = False
    MyBody.Visible = False
    EyeLeft(0).Visible = False
    EyeLeft(1).Visible = False
    EyeRight(0).Visible = False
    EyeRight(1).Visible = False
    MyNose.Visible = False
    MyMouth.Visible = False
    MyHead.Visible = False

    'reset other variables
    lblRemaining = "10 Guesses Remaining Until Your Doom"
    lblTotal = "Total Right -  " & iTotalRight & "       Total Wrong - " & iTotalWrong & "          Total Logins - " & iPlayed & "        Total Minutes Played - " & iTime
    iWrong = 0
    
    'Reset word
    For X = 0 To 11
        lblWord(X) = ""
    Next
    
    'Put letter back on board
    For X = 0 To 25
        lblLetters(X).Visible = True
    Next
    
    'Format word
    For X = 0 To Len(strWord) - 1
        lblWord(X) = "_"
        If Mid(strWord, X + 1, 1) = " " Then lblWord(X) = " "
    Next
    
    'Centre word
    Frame1.Left = (Width / 2) - (Len(strWord) * 310) / 2
    
    IsEnabled = True
End Sub

Private Sub Pause(sngTime As Single)
    Dim sngStart As Single
    
    sngStart = Timer
    
    Do Until sngStart + sngTime < Timer
        DoEvents
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Update stats if they actually played
    If strUsername <> "" Then
        Cn.Execute ("UPDATE [User] SET [Played] = [Played] + 1, [Right] = " & iTotalRight & ", [Wrong] = " & iTotalWrong & ", [Time] = " & iTime & " WHERE UserName = '" & strUsername & "'")
    End If
               
    Cn.Close
    Set Cn = Nothing
End Sub

Private Sub lblLetters_Click(Index As Integer)
    Form_KeyPress (Index + 65)
End Sub

Private Sub WrongLetter()
    iWrong = iWrong + 1
    
    lblRemaining = 10 - iWrong & " Guesses Remaining Until Your Doom"
    
    'All guesses used up
    If iWrong = 10 Then
        LegRight(0).Visible = True
        LegRight(1).Visible = True
        
        For X = 0 To Len(strWord)
            lblWord(X) = Mid(strWord, X + 1, 1)
        Next
                
        iTotalWrong = iTotalWrong + 1
        MyHead.BackColor = vbRed
        EyeLeft(0).BackColor = vbRed
        EyeRight(0).BackColor = vbRed
        Pause (3)
        MyHead.BackColor = &HC0E0FF
        EyeLeft(0).BackColor = &HFF8080
        EyeRight(0).BackColor = &HFF8080
        GetWord
    End If
    
    Select Case iWrong
        Case 1:
            MyHead.Visible = True
        Case 2:
            MyMouth.Visible = True
        Case 3:
            MyNose.Visible = True
        Case 4:
            EyeRight(0).Visible = True
            EyeRight(1).Visible = True
        Case 5:
            EyeLeft(0).Visible = True
            EyeLeft(1).Visible = True
        Case 6:
            MyBody.Visible = True
        Case 7:
            ArmLeft(0).Visible = True
            ArmLeft(1).Visible = True
            ArmLeft(2).Visible = True
            ArmLeft(3).Visible = True
        Case 8:
            ArmRight(0).Visible = True
            ArmRight(1).Visible = True
            ArmRight(2).Visible = True
            ArmRight(3).Visible = True
        Case 9:
            LegLeft(0).Visible = True
            LegLeft(1).Visible = True
    End Select
End Sub

Private Sub WordDone()
    iTotalRight = iTotalRight + 1
    Pause (1)
    GetWord
End Sub

Private Sub Timer1_Timer()
    iTime = iTime + 1
    lblTotal = "Total Right -  " & iTotalRight & "        Total Wrong - " & iTotalWrong & "          Total Logins - " & iPlayed & "        Total Minutes Played - " & iTime
End Sub
