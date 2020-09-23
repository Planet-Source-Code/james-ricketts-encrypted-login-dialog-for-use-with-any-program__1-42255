VERSION 5.00
Begin VB.Form frm_newpassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter new user..."
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5115
   ControlBox      =   0   'False
   Icon            =   "frm_newpassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_authorisation 
      ForeColor       =   &H000000FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "$"
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txt_newpassword2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txt_newpassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The default authorisation password is ""BOB"""
      Height          =   615
      Left            =   3360
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl1 
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Authorisation Password"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   -360
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-enter password"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frm_newpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim change_password As String * 15
Dim linenumber As Integer
Dim old_password As String * 15
Dim newusernameandpass As String * 20
Dim user As String * 10
Dim pass As String * 8
Dim holding(20) As String * 20
Dim temp As String * 20
Dim freespace As Integer

Private Sub cmd_cancel_Click()
Unload Me
If switchtomainform = True Then
    frm_passwordscreen.txt_user.SetFocus
ElseIf switchtomainform = False Then
'frm_users.Show
'frm_users.SetFocus
End If

End Sub

Private Sub cmd_ok_Click()

'The point of authorisation is so that if a user tries to get into the program by simply deleting the Users.dat file, then
'They will need the authorisation password to create a new user. Preventing them from accessing the program.


''''****OBVIOUSLY there wouldnt be the label saying the default authorisation password in the real program,
''''****that is just to allow people access before they know what it is for***


If authorisationneeded = True Then
   
    If LCase(txt_authorisation.Text) <> authorisationpass Then
        result = MsgBox("The authorisation password is invalid, please re-type", vbCritical, "Authorisation Error")
    Exit Sub
    ElseIf LCase(txt_authorisation.Text) = authorisationpass Then
    End If

ElseIf authorisationneeded = False Then
End If

If LCase(txt_newpassword.Text) <> LCase(txt_newpassword2.Text) Then
result = MsgBox("The passwords you entered do not match, please re-type", vbCritical, "Password matching error")
txt_newpassword.Text = ""
txt_newpassword2.Text = ""
txt_newpassword.SetFocus
Exit Sub
End If

'check for blank fields
If Text1.Text = "" Or txt_newpassword.Text = "" Then
result = MsgBox("Each field must contain data", vbCritical, "Error...")
Text1.SetFocus
Exit Sub
End If

'check for fields with spaces
If Left(Text1.Text, 1) = " " Or Left(txt_newpassword.Text, 1) = " " Then
result = MsgBox("The username or password cannot be spaces ONLY, or BEGIN with a space", vbCritical, "Error...")
txt_newpassword.Text = ""
txt_newpassword2.Text = ""
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If

Open filename For Random As #1 Len = 20

user = LCase(Trim(Text1.Text))
pass = LCase(Trim(txt_newpassword.Text))
newusernameandpass = (user & ": " & pass)
result = Encrypt(newusernameandpass, 20)


'find next available slot in #1
For i = 1 To 20
Get #1, i, holding(i)
Next

For i = 1 To 20
    If holding(i) = temp Then
    freespace = i
    Exit For
    End If
Next

'find out whether the new username is already present in the list (to prevent doubles)

Put #1, freespace, newusernameandpass

Close #1

result = MsgBox("The New Username and Password have been added", vbExclamation, "New User Added")

If switchtomainform = True Then
    switchtomainform = False
    loggedin = user
    Unload Me
    Unload frm_passwordscreen
    frm_main.Show
    frm_main.SetFocus
ElseIf switchtomainform = False Then
    Unload Me
    frm_users.Refreshlist
    frm_users.Show
    frm_users.SetFocus
End If

End Sub

Private Sub Form_Load()

Text1.Text = ""
txt_authorisation.Text = ""
txt_newpassword.Text = ""
txt_newpassword2.Text = ""
authorisationpass = "bob"
authorisationneeded = False
switchtomainform = False

If newpasswordlist = True Then
newpasswordlist = False
authorisationneeded = True
switchtomainform = True
txt_authorisation.Visible = True
Label5.Visible = True
Me.Height = 2790

ElseIf newpasswordlist = False Then
txt_authorisation.Visible = False
Label5.Visible = False
Me.Height = 2145
End If

End Sub

Private Sub txt_newpassword_Change()
If Len(txt_newpassword.Text) = 8 Then
lbl1.Caption = "The password may only be upto 8 characters long"
ElseIf Len(txt_newpassword.Text) < 8 Then
lbl1.Caption = ""
End If

End Sub


Private Sub txt_newpassword2_Change()
If Len(txt_newpassword2.Text) = 8 Then
lbl1.Caption = "The password may only be upto 8 characters long"
ElseIf Len(txt_newpassword2.Text) < 8 Then
lbl1.Caption = ""
End If

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 10 Then
lbl1.Caption = "The username may only be upto 10 characters long"
ElseIf Len(Text1.Text) < 10 Then
lbl1.Caption = ""
End If

End Sub
