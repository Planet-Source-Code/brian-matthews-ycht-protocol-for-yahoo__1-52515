VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Micro-Bot"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7680
      TabIndex        =   26
      Text            =   "1.0"
      Top             =   3960
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Save (Room, ID, Password)"
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   360
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   3000
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Save Settings (Messages, Values)"
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4200
      MaxLength       =   800
      TabIndex        =   16
      Text            =   "Good bye, check out http://ycrack.net"
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      MaxLength       =   800
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Form1.frx":08CA
      Top             =   5880
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Send Private Message on Entrance/Exit"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   3000
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Text            =   "15.0"
      ToolTipText     =   "Custom Interval"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      ToolTipText     =   "Stop"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      ToolTipText     =   "Start"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Message to All"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      ToolTipText     =   "Send Message to All"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Text            =   "Yahoo ID"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Login"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      MaxLength       =   800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Message:"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrace Message:"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pause:"
      Height          =   255
      Left            =   6960
      TabIndex        =   22
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Room:"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Yahoo ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents ycht As Protocol_YCHT
Attribute ycht.VB_VarHelpID = -1

Private Sub Command1_Click()
ycht.ychtSendPrivateMessage Text6.Text, Text7.Text
End Sub

Private Sub Command2_Click()
For X = 0 To List1.ListCount
ycht.ychtSendPrivateMessage List1.List(X), Text7.Text
Pause Text11.Text
Next X
End Sub

Private Sub Command3_Click()
    'Initialize our class module
    Set ycht = New Protocol_YCHT
    'Set the Chat and Login servers we'll be using
    ycht.Server_Chat = "jcs1.chat.dcn.yahoo.com"
    'jcs1.chat.dcn.yahoo.com
    'scsc.msg.yahoo.com*
    'cs8.chat.yahoo.com
    'cs72.dcn.sc5.yahoo.com*
    ycht.Server_Login = "login.yahoo.com"
    'Change the below Username/Password to the one you wish to use
    ycht.Login_Username = Text4.Text
    ycht.Login_Password = Text5.Text
    'Connect to YCHT Server
    ycht.ychtConnect
End Sub

Private Sub Command4_Click()
Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
Timer1.Enabled = False

End Sub

Private Sub Form_Load()
Text4.Text = GetSetting("YCht", "Usernames", "YCht", Text4)
Text5.Text = GetSetting("YCht", "Passwords", "YCht", Text5)
Text3.Text = GetSetting("YCht", "Room", "YCht", Text3)
Check2.Value = GetSetting("YCht", "Check", "Check2", Check2.Value)
If Check2.Value = vbChecked Then
Text2.Text = GetSetting("YCht", "Messages", "YCht2", Text2)
Text7.Text = GetSetting("YCht", "Messages", "YCht", Text7)
'Text11.Text = GetSetting("YCht", "Intervals", "YCht", Text11)
Else
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Check2.Value = vbChecked Then
SaveSetting "YCht", "Check", "Check2", Check2.Value
SaveSetting "YCht", "Messages", "YCht", Text7
SaveSetting "YCht", "Messages", "YCht2", Text2
SaveSetting "YCht", "Intervals", "YCht", Text11
SaveSetting "YCht", "Usernames", "YCht", Text4
SaveSetting "YCht", "Passwords", "YCht", Text5
SaveSetting "YCht", "Room", "YCht", Text3
SaveSetting "YCht", "Check", "Check2", Check2.Value
Else
End If
End Sub

Private Sub Text11_Change()
Text11.ToolTipText = "Pause " & "[" & Text11 & "]" & " " & "Between Message to be sent to ""All"""
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Then
Label1.Caption = "Enter A Room Name"
Else
Label1.Caption = "Room: " & Text3.Text
End If
End Sub

Private Sub Timer1_Timer()
For X = 1 To List1.ListCount
ycht.ychtSendPrivateMessage List1.List(X), Text7.Text
Pause Text8.Text
Next X
End Sub

'Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
'ycht.ychtSendPrivateMessage Text6.Text, Text7.Text
'End Sub

Private Sub ycht_Away(strMsg As String)
    'Event is fired when a user goes away or comes back
    '----
    'strMsg includes the username and message such as "coozzzzz is back."
    MsgBox strMsg, vbExclamation
End Sub
Public Sub ycht_Connected(isConnected As Boolean)
    'Event is fired when our connection state changes
    '----
    'isConnected=true   : when connected
    'isConnected=false  : when disconnected
    If isConnected = True Then ycht.ychtJoinRoom Text3.Text
End Sub
Private Sub ycht_FriendStatus(strFriend As String, fStatus As FriendStatus)
    'Event is fired when a friend's status changes to Online/Offline/Chat/Games
    '----
    'I've used an Enum to handle the Statuses and should be self explanatory
    'in the below usage
    Select Case fStatus
        Case FriendStatus.ChatJoined
            Text1.Text = Text1.Text & strFriend & " in in chat." & vbCrLf & vbCrLf
        Case FriendStatus.ChatLeft
            Text1.Text = Text1.Text & strFriend & " left chat." & vbCrLf & vbCrLf
        Case FriendStatus.GamesJoined
            Text1.Text = Text1.Text & strFriend & " is in games." & vbCrLf & vbCrLf
        Case FriendStatus.GamesLeft
            Text1.Text = Text1.Text & strFriend & " left games." & vbCrLf & vbCrLf
        Case FriendStatus.OnlineFalse
            Text1.Text = Text1.Text & strFriend & " is offline." & vbCrLf & vbCrLf
        Case FriendStatus.OnlineTrue
            Text1.Text = Text1.Text & strFriend & " is online." & vbCrLf & vbCrLf
    End Select
End Sub
Private Sub ycht_ReceivedEmail(emailCount As String)
    'Event is fired when you receive a new e-mail on the name you're currently using
    '----
    'emailCount     : Count of how many emails you have
    MsgBox "We've received a new e-mail! (total of " & emailCount & " email(s)).", vbInformation
End Sub
Private Sub ycht_ReceivedInvite(strRoom As String, strUser As String)
    'Event is fired when you receive an invitation. (from testing i've noticed there
    'are problems with this between YMSG/YCHT.. however it works if invited from a
    'YMSG protocol user."
    '----
    'strRoom    : The room you are invited to
    'strUser    : The user who invited you
    Dim lRet As Long
    lRet = MsgBox("Youve been invited to join " & strRoom & " By " & strUser & vbCrLf & vbCrLf & "Join Now?", vbQuestion + vbYesNo, "Invitation")
    If lRet = vbYes Then ycht.ychtJoinRoom strRoom
End Sub
Private Sub ycht_ReceivedPrivateMessage(strUser As String, strMsg As String)
    'strUser
    'strMsg
    Text1.Text = Text1.Text & "Private Message - " & strUser & ": " & parse_HTML(strMsg) & vbCrLf
    'frmMsg.Show
    'frmMsg.Caption = strUser & " -- " & "Instant Message"
    'frmMsg.Text1.Text = Text1.Text & vbCrLf & strMsg
    'frmMsg.Text1.Text = frmMsg.Text1.Text & vbCrLf & strUser & ": " & parse_HTML(strMsg) & vbCrLf
End Sub
Public Sub ycht_Error(strError As String)
    'Event is fired when a handled error occurs within the Class Module
    '----
    'strError   : The error message
    MsgBox strError, vbCritical
End Sub
Private Sub ycht_RoomJoined(strRoom As String, strRoomTopic As String)
    'Event is fired when you join a new room. I've added this so you know when you
    'should clear your User List to add new Users.
    '----
    'strRoom        : The room you joined
    'strRoomTopic   : The topic of the room you joined
    List1.Clear
    Form1.Caption = "Connected: " & " - " & strRoom
End Sub
Private Sub ycht_UserEntered(strUser As String)
If Check1.Value = vbChecked Then
ycht.ychtSendPrivateMessage strUser, Text9.Text
Else
End If
    Dim i As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        If StrComp(strUser, List1.List(i), vbTextCompare) = 0 Then Exit Sub
    Next i
    List1.AddItem strUser
End Sub
Private Sub ycht_UserLeft(strUser As String)
If Check1.Value = vbChecked Then
ycht.ychtSendPrivateMessage strUser, Text10.Text
Else
End If
    Dim i As Integer
    For i = List1.ListCount - 1 To 0 Step -1
        'If the user exists in our list then remove them
        '(suggested that you should compare each list(i) with strUser in lowercase)
        If StrComp(strUser, List1.List(i), vbTextCompare) = 0 Then List1.RemoveItem i
    Next i
End Sub
Private Sub ycht_ReceivedMessage(strUser As String, strMsg As String)
    Text1.Text = Text1.Text & strUser & ": " & parse_HTML(strMsg) & vbCrLf
End Sub
Private Sub ycht_ReceivedEmote(strUser As String, strMsg As String)
    'Event is fired when you receive a new emote within the room you're currently in
    '----
    'strUser    : The user that sent the message
    'strMsg     : The message the user sent
    Text1.Text = Text1.Text & strUser & " " & parse_HTML(strMsg) & vbCrLf
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    'The below is just to handle chat commands such as /join, /goto, /invite, etc
    Dim tSplit() As String
    If KeyCode = "13" Then
        If Len(Text2.Text) > 0 Then
            If Mid(Text2.Text, 1, 5) = "/join" Then
                ycht.ychtJoinRoom Mid(Text2.Text, 7)
            ElseIf Mid(Text2.Text, 1, 3) = "/pm" Then
                tSplit = Split(Mid(Text2.Text, 4), " ")
                If IsArray(tSplit) Then
                    If UBound(tSplit) = 2 Then ycht.ychtSendPrivateMessage tSplit(1), tSplit(2)
                End If
            ElseIf Mid(Text2.Text, 1, 5) = "/goto" Then
                ycht.ychtGotoUser Mid(Text2.Text, 7)
            ElseIf Mid(Text2.Text, 1, 1) = ":" Then
                ycht.ychtSendEmote Mid(Text2.Text, 2)
            ElseIf Mid(Text2.Text, 1, 7) = "/invite" Then
                ycht.ychtSendInvite Mid(Text2.Text, 9)
            Else
                ycht.ychtSendMessage Text2.Text
            End If
            Text2.Text = ""
        End If
    End If
End Sub
Private Sub Text1_Change()
    'The below just keeps the last line in the text box visible (auto-scrolling)
    Text1.SelStart = Len(Text1.Text)
    'Call LimitLines(Text1, , 10)
End Sub
Private Function parse_HTML(strCheck As String) As String
    'The below just parses certain html tags from strings.. it's a bit rough but
    'it's not that important to optimize at the moment
    Dim Pos1 As Integer, Pos2 As Integer
reparse1:
    Pos1 = InStr(1, LCase(strCheck), "")
    If Pos1 > 0 Then
        Pos2 = InStr(Pos1, LCase(strCheck), "m")
        If Pos2 > 0 Then
            If Pos1 = 1 Then
                strCheck = Mid(strCheck, Pos2 + 1)
            Else
                strCheck = Mid(strCheck, 1, Pos1 - 1) & Mid(strCheck, Pos2 + 1)
            End If
        Else
            parse_HTML = strCheck
            GoTo reparse2
        End If
    Else
        parse_HTML = strCheck
        GoTo reparse2
    End If
    GoTo reparse1
reparse2:
    Pos1 = InStr(1, LCase(strCheck), "<font")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "</")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "</")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "<b")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "<alt")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "<fade")
    If Pos1 > 0 Then
        Pos2 = InStr(Pos1, LCase(strCheck), ">")
        If Pos2 > 0 Then
            If Pos1 = 1 Then
                strCheck = Mid(strCheck, Pos2 + 1)
            Else
                strCheck = Mid(strCheck, 1, Pos1 - 1) & Mid(strCheck, Pos2 + 1)
            End If
        Else
           parse_HTML = strCheck
            Exit Function
        End If
    Else
        parse_HTML = strCheck
        Exit Function
    End If
    GoTo reparse2
End Function

Sub LimitLines(TextBox As Object, Optional Delimiter As String = vbCrLf, Optional MaxLength As Integer = 500)
    Dim uberSplit, lenFirstLine
    uberSplit = Split(TextBox.Text, Delimiter)

    If UBound(uberSplit) > MaxLength Then
        lenFirstLine = InStr(TextBox.Text, Delimiter) + 1


        With TextBox
            .SelStart = 0
            .SelLength = lenFirstLine
            .SelText = ""
            .SelStart = Len(TextBox.Text)
        End With
    End If
End Sub
