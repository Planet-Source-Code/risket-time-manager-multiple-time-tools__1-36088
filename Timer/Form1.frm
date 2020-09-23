VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Manager"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command16 
      Caption         =   "Test Beep"
      Height          =   315
      Left            =   2100
      TabIndex        =   40
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Set Time"
      Height          =   315
      Left            =   2100
      TabIndex        =   39
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4440
      Top             =   2940
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Alarm Clock"
      Height          =   330
      Left            =   60
      TabIndex        =   29
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Hide Settings"
      Height          =   315
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop Time/Beep"
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Timer"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Countdown"
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   4020
      Top             =   2940
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   3600
      Top             =   2940
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   4080
      TabIndex        =   23
      Top             =   -60
      Width           =   4155
      Begin VB.Label Label7 
         Caption         =   $"Form1.frx":1272
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1995
         Left            =   60
         TabIndex        =   25
         Top             =   720
         Width           =   4035
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   4080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Manager"
         BeginProperty Font 
            Name            =   "Stop"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   240
         TabIndex        =   24
         Top             =   60
         Width           =   3660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Countdown Settings"
      Height          =   2895
      Left            =   4080
      TabIndex        =   1
      Top             =   -60
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton Command1 
         Caption         =   "Begin Countdown"
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   2460
         Width           =   3375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Reset Clock"
         Height          =   315
         Left            =   1440
         TabIndex        =   26
         Top             =   2100
         Width           =   1155
      End
      Begin VB.CommandButton Command10 
         Caption         =   "STOP!"
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   2100
         Width           =   1035
      End
      Begin VB.TextBox Second 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         TabIndex        =   20
         Text            =   "00"
         Top             =   840
         Width           =   435
      End
      Begin VB.TextBox Minute 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   19
         Text            =   "00"
         Top             =   840
         Width           =   435
      End
      Begin VB.TextBox Hour 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         TabIndex        =   18
         Text            =   "00"
         Top             =   840
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Set Time"
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   2100
         Width           =   1095
      End
      Begin VB.TextBox Day 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Text            =   "00"
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Day    Hour     Minute Second"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   21
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label lblZero 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H8000000F&
         Height          =   450
         Left            =   2700
         TabIndex        =   11
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label Label3 
         Caption         =   $"Form1.frx":1559
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Label Label2 
         Caption         =   "Start Countdown at:"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Alarm Clock Settings"
      Height          =   2895
      Left            =   4080
      TabIndex        =   30
      Top             =   -60
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton Command13 
         Caption         =   "Set Alarm"
         Height          =   315
         Left            =   3060
         TabIndex        =   36
         Top             =   2460
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   1035
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   1380
         Width           =   3915
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Text            =   "1:23:45 PM"
         Top             =   600
         Width           =   3915
      End
      Begin VB.Label lblASet 
         AutoSize        =   -1  'True
         Caption         =   "(not set)"
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   780
         TabIndex        =   38
         Top             =   2520
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Set For:"
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label Label9 
         Caption         =   "Enter reminder message to display:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1140
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "Set time for alarm to go off:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Timer Settings"
      Height          =   2835
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton Command2 
         Caption         =   "Begin Timer"
         Height          =   315
         Left            =   2700
         TabIndex        =   27
         Top             =   2280
         Width           =   1155
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Reset Timer"
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop Timer"
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   $"Form1.frx":1608
         Height          =   1755
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   3795
      End
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00 XX"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   60
      TabIndex        =   31
      Top             =   2220
      Width           =   3975
   End
   Begin VB.Label d 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   60
      TabIndex        =   17
      ToolTipText     =   "Days"
      Top             =   60
      Width           =   915
   End
   Begin VB.Label h 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   1080
      TabIndex        =   16
      ToolTipText     =   "Hours"
      Top             =   60
      Width           =   915
   End
   Begin VB.Label m 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   2100
      TabIndex        =   15
      ToolTipText     =   "Minutes"
      Top             =   60
      Width           =   915
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   60
      X2              =   4020
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   4020
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   60
      X2              =   4020
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   60
      X2              =   4020
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   3120
      TabIndex        =   0
      ToolTipText     =   "Seconds"
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Command10_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
End Sub

Private Sub Command11_Click()
    s.Caption = "00"
    m.Caption = "00"
    h.Caption = "00"
    d.Caption = "00"
End Sub

Private Sub Command12_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = True
    Me.Width = "8355"
End Sub

Private Sub Command13_Click()
    lblASet.Caption = Text1.Text
End Sub

Private Sub Command14_Click()
    lblASet.Caption = Text1.Text
End Sub

Private Sub Command16_Click()
    Beep
End Sub

Private Sub Command2_Click()
    Frame1.Visible = False
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Timer1.Enabled = True
    Timer2.Enabled = False
End Sub

Private Sub Command3_Click()
    d.Caption = Day.Text
    h.Caption = Hour.Text
    m.Caption = Minute.Text
    s.Caption = Second.Text
End Sub

Private Sub Command4_Click()
    Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Me.Width = "8355"
End Sub

Private Sub Command6_Click()
    Frame1.Visible = False
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
    Me.Width = "8355"
End Sub

Private Sub Command7_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    s.Caption = "00"
End Sub

Private Sub Command8_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Me.Width = "4155"
End Sub

Private Sub Command9_Click()
    s.Caption = "00"
    m.Caption = "00"
    h.Caption = "00"
    d.Caption = "00"
End Sub

Private Sub Label6_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
    Frame4.Visible = False
    Me.Width = "8355"
End Sub

Private Sub Timer1_Timer()
If s.Caption = "60" Then
    s.Caption = "0"
    m.Caption = m.Caption + 1
Else
    s.Caption = s.Caption + 1
End If

If m.Caption = "60" Then
    m.Caption = "0"
    h.Caption = h.Caption + 1
End If

If h.Caption = "60" Then
    h.Caption = "0"
    d.Caption = d.Caption + 1
End If

If d.Caption = "99" Then
    s.Caption = "0"
    m.Caption = "0"
    h.Caption = "0"
    d.Caption = "0"
    Timer1.Enabled = False
    MsgBox "How can you possibly let your computer run for over 100 days!? Get a life! Or have you already gotten one and you're now too busy to simply turn your computer off or at least close this program!?", vbCritical, "Are you seriouse!?"
End If

End Sub

Private Sub Timer2_Timer()
    s.Caption = s.Caption - 1

If s.Caption = "0" Then
    s.Caption = "59"
    m.Caption = m.Caption - 1
End If
If m.Caption = "0" Then
    m.Caption = "59"
    h.Caption = h.Caption - 1
End If
If h.Caption = "0" Then
    h.Caption = "59"
    d.Caption = d.Caption - 1
End If
If m.Caption = "00" Then
    'does nothing
End If
If h.Caption = "00" Then
    'does nothing
End If
If d.Caption = "00" Then
    'does nothing
End If
End Sub

Private Sub Timer3_Timer()
    lblTime.Caption = Time()
If lblTime.Caption = lblASet.Caption Then
    MsgBox Text2.Text, vbExclamation, "The alarm was set for: " + Text1.Text + " to remind you:"
End If
End Sub
