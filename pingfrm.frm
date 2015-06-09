VERSION 5.00
Begin VB.Form pingfrm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ip auto ping"
   ClientHeight    =   1815
   ClientLeft      =   3900
   ClientTop       =   2355
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4455
   Begin VB.ListBox List2 
      BackColor       =   &H00FF8080&
      Height          =   1815
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "60"
      Top             =   1525
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "255.255.255.255"
      Top             =   1525
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FF8080&
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnulist1 
      Caption         =   "&list1"
      Visible         =   0   'False
      Begin VB.Menu mnuremove1 
         Caption         =   "&remove"
      End
      Begin VB.Menu mnuclear1 
         Caption         =   "&clear"
      End
   End
   Begin VB.Menu mnulist2 
      Caption         =   "&list2"
      Visible         =   0   'False
      Begin VB.Menu mnuadd 
         Caption         =   "&add"
      End
      Begin VB.Menu mnuping 
         Caption         =   "&ping"
      End
      Begin VB.Menu mnuremove2 
         Caption         =   "&remove"
      End
      Begin VB.Menu mnuclear2 
         Caption         =   "&clear"
      End
   End
End
Attribute VB_Name = "pingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxCount As Integer
Dim IntervalCount() As Integer

Public Sub GetPingReply(HostIP, Current)
On Error Resume Next
    Dim s() As String
    Dim ECHO As ICMP_ECHO_REPLY
    Dim Temp As Integer
    
    s() = Split(HostIP, " ", 2)
    For i = 1 To 4
        Call Ping(s(0), ECHO) 'Ping HostIP
        DoEvents
        If Str(ECHO.RoundTripTime) <> 0 Then Temp = Temp + 1
        'Wait for Response, Get Time Elapsed
        Label1 = Label1 & Str(ECHO.RoundTripTime) & "ms" & vbCrLf
        ECHO.RoundTripTime = 0 'Reset
    Next i
    If Current <> -1 And Temp >= 2 Then
        List1.RemoveItem (Current)
        List2.AddItem HostIP
        For i = 1 To MaxCount - 1
            IntervalCount(i) = IntervalCount(i + 1)
        Next i
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    StayOnTop Me, True
    MaxCount = 5
    ReDim IntervalCount(MaxCount)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then PopupMenu mnulist1
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then PopupMenu mnulist2
End Sub

Private Sub mnuadd_Click()
On Error Resume Next
    IPrompt = "enter ip of host."
    ITitle = "add ip"
    IDefault = ""
    NewVal = InputBox(IPrompt, ITitle, IDefault)
    If Len(NewVal) <> 0 Then List2.AddItem NewVal
End Sub

Private Sub mnuclear1_Click()
On Error Resume Next
    If List1.ListCount <> 0 Then
        sure = MsgBox("are you sure?", vbYesNo, "clear")
        If sure = 6 Then List1.Clear
    End If
End Sub

Private Sub mnuclear2_Click()
On Error Resume Next
    If List2.ListCount <> 0 Then
        sure = MsgBox("are you sure?", vbYesNo, "clear")
        If sure = 6 Then List2.Clear
    End If
End Sub

Private Sub mnuping_Click()
On Error Resume Next
    If List2.ListIndex = -1 Then Exit Sub
    Label1 = List2.List(List2.ListIndex) & vbCrLf & vbCrLf
    GetPingReply List2.List(List2.ListIndex), -1
End Sub

Private Sub mnuremove1_Click()
On Error Resume Next
    If List1.ListIndex = -1 Then Exit Sub
    If List1.ListCount <> 0 Then
        sure = MsgBox("are you sure?", vbYesNo, "clear")
        If sure = 6 Then List1.RemoveItem List1.ListIndex
    End If
End Sub

Private Sub mnuremove2_Click()
On Error Resume Next
    If List2.ListIndex = -1 Then Exit Sub
    If List2.ListCount <> 0 Then
        sure = MsgBox("are you sure?", vbYesNo, "clear")
        If sure = 6 Then List2.RemoveItem List2.ListIndex
    End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn And Len(Text1) <> 0 And Len(Text2) <> 0 Then
        If List1.ListCount = MaxCount Then
            MaxCount = MaxCount + 5
            ReDim Preserve IntervalCount(MaxCount)
        End If
        List1.AddItem Text1 & ":" & Text2
        Text1 = ""
    End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn And Len(Text1) <> 0 And Len(Text2) <> 0 Then
        If List1.ListCount = MaxCount Then
            MaxCount = MaxCount + 5
            ReDim Preserve IntervalCount(MaxCount)
        End If
        List1.AddItem Text1 & ":" & Text2
        Text1 = ""
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
For i = 0 To List1.ListCount - 1
    DoEvents
    TempString = List1.List(i)
    Where = InStr(TempString, ":")
    IntervalCount(i + 1) = IntervalCount(i + 1) + 1
    Me.Caption = "ip auto ping: " & i + 1 & " - " & IntervalCount(i + 1)
    CheckInterval = Mid(TempString, Where + 1, Len(TempString))
    If IntervalCount(i + 1) >= CheckInterval Then
        IntervalCount(i + 1) = 0
        CheckIP = Mid(TempString, 1, Where - 1)
        Label1 = CheckIP & vbCrLf & vbCrLf
        GetPingReply CheckIP, i
    End If
Next i
End Sub
