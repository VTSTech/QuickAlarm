VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickAlarm v0.2 by Nigel Todman"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Use Ding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Text            =   "00"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "www.NigelTodman.com"
      Height          =   195
      Left            =   1515
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   72
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   -240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Copyright (C) 2015  Nigel Todman (nigel.todman@gmail.com)
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Private Const SND_PURGE = &H40               '  purge non-static events for task
Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Private Const SND_SYNC = &H0         '  play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Dim Counting, AlarmSounding As Boolean
Dim Ticker As Integer
Dim SecsLeft, MinLeft

Private Sub Command1_Click()
If Counting = False Then
    Command1.Caption = "STOP"
    Counting = True
    Ticker = 0
    MinLeft = Text1.Text
    MinLeft = MinLeft - 1
    SecsLeft = (MinLeft * 60) + 60
    Label1.Caption = MinLeft & ":" & (SecsLeft - (MinLeft * 60))
Else
    Command1.Caption = "START"
    Counting = False
    AlarmSounding = False
    Label1.Caption = "0:00"
End If

End Sub

Private Sub Form_Load()
Timer1.Interval = 1000
Timer1.Enabled = True
Ticker = 0
Text1.Text = 10
Counting = False
AlarmSounding = False
End Sub

Private Sub Label3_Click()
Shell ("cmd.exe /c start http://www.NigelTodman.com")
End Sub

Private Sub Timer1_Timer()
If Counting = True Then
If Label1.Caption = "0:0" And AlarmSounding = False Then
    If Check1.Value = 0 Then
        PlaySound VB.App.Path & "\BOMB_SIREN-BOMB_SIREN-247265934.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
        'Source: http://soundbible.com/2056-Missile-Alert.html
        'Available under Creative Commons Attribution 3.0 License
        'License: https://creativecommons.org/licenses/by/3.0/legalcode
        AlarmSounding = True
    ElseIf Check1.Value = 1 Then
        Beep
        Beep
        Beep
        Beep
        Beep
        AlarmSounding = True
    End If
Else
    If Ticker = 60 Then
        MinLeft = MinLeft - 1
        If MinLeft <= -1 Then
            MinLeft = 0
        End If
        Ticker = 0
        Ticker = Ticker + 1
    Else
        If MinLeft <= -1 Then
            MinLeft = 0
        End If
        Ticker = Ticker + 1
        SecsLeft = SecsLeft - 1
        If SecsLeft <= 0 Then
            SecsLeft = 0
        End If
        Label1.Caption = MinLeft & ":" & (SecsLeft - (MinLeft * 60))
    End If
End If
End If
End Sub
