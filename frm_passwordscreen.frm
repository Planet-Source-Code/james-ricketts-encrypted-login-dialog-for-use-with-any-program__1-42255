VERSION 5.00
Begin VB.Form frm_passwordscreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1920
   ClientLeft      =   1140
   ClientTop       =   3660
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frm_passwordscreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txt_user 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txt_Password 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   8
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm_passwordscreen.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Username Here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password Here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "frm_passwordscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim recarray(20) As String * 20
Dim username(20) As String * 10
Dim password(20) As String * 8
Dim usernamepass, usernamepassactual As String * 20
Dim user As String * 10
Dim pass As String * 8
Dim temp As String * 20
Dim exitloop As Boolean
Dim filename As String

exitsubvar = False
filename = App.Path & "\login.dat"

Open filename For Random As #1 Len = 20

For i = 1 To 20
Get #1, i, recarray(i)
Next

If recarray(1) = temp Then
    If MsgBox("No usernames are detected, would you like to input one?", vbYesNo, "Input new password?") = vbNo Then
    txt_Password = ""
    txt_user.Text = ""
    txt_user.SetFocus
    Close #1
    Exit Sub
    Else:
    newpasswordlist = True
    'Unload Me
    Close #1
    frm_newpassword.Show 1
    Close #1
    Exit Sub
    End If
End If


user = LCase(txt_user.Text)
pass = LCase(txt_Password.Text)

usernamepass = (user & ": " & pass)


For i = 1 To 20
    username(i) = (Left$(recarray(i), 10))
    result = Decrypt(username(i), 10)
    
    password(i) = (Right$(recarray(i), 8))
    result = Decrypt(password(i), 8)
Next

For i = 1 To 20
    
    If user = username(i) And pass = password(i) Then
    'MsgBox "bingo"
    'Exit For

    exitloop = True
    End If
    
Next

If exitloop = True Then
Unload Me
loggedin = user
frm_main.Show
Call frm_main.updatecaption
exitloop = False
Close #1
txt_user.Text = ""
txt_Password.Text = ""
Exit Sub
End If

'if passowrd and username not found then:
For i = 1 To 20
    If user = username(i) Then
    result = MsgBox("The password for this username is invalid", vbCritical, "Invalid Password")
    exitloop = True
    End If
Next


If exitloop = True Then
    exitloop = False
    txt_Password = ""
    txt_Password.SetFocus
    Close #1
    Exit Sub
End If

result = MsgBox("The username is invalid", vbCritical, "Invalid Username")
txt_user.Text = ""
txt_Password.Text = ""
txt_user.SetFocus
Close #1

    
End Sub

Private Sub Command3_Click()
Unload frm_main
Unload Me
End
End Sub

Private Sub Form_Load()

filename = App.Path & "\Login.dat"
filenametemp = (App.Path & "\temp.jpg")

documentCount = 0
lbl_version = "v" & App.Major & "." & App.Minor
End Sub


Private Sub txt_Password_Change()
If Len(txt_Password.Text) = 8 Then
lbl1.Caption = "Password must 8 characters or less"
ElseIf Len(txt_Password.Text) < 8 Then
lbl1.Caption = ""
End If

End Sub

Private Sub txt_user_Change()

If Len(txt_user.Text) = 10 Then
lbl1.Caption = "Name must be 10 characters or less"
ElseIf Len(txt_user.Text) < 10 Then
lbl1.Caption = ""
End If

End Sub
