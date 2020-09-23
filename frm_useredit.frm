VERSION 5.00
Begin VB.Form frm_useredit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit User - "
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_oldpass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txt_pass2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txt_pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Password"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Re-type New password"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Username"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frm_useredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim newuserandpass As String * 20
Dim user As String * 10
Dim pass As String * 8
Dim existinguserandpass As String * 20
Dim existingpass As String * 8
Dim typedoldpass As String * 8

Open FileName For Random As #1 Len = 20
Get #1, Index, existinguserandpass
result = Decrypt(existinguserandpass, 20)
existingpass = Right(existinguserandpass, 8)
typedoldpass = txt_oldpass.Text
Close #1

If LCase(existingpass) <> LCase(typedoldpass) Then
    result = MsgBox("The current password you entered is wrong, please re-type", vbCritical, "Password matching error")
    txt_oldpass.Text = ""
    txt_pass.Text = ""
    txt_pass2.Text = ""
    txt_oldpass.SetFocus
    Exit Sub
End If

If LCase(txt_pass.Text) <> LCase(txt_pass.Text) Then
    result = MsgBox("The passwords you entered do not match, please re-type", vbCritical, "Password matching error")
    txt_oldpass.Text = ""
    txt_pass.Text = ""
    txt_pass2.Text = ""
    txt_pass.SetFocus
    Exit Sub
End If
    
'check for blank fields
If Text1.Text = "" Or txt_pass.Text = "" Then
    result = MsgBox("Each field must contain data", vbCritical, "Error...")
    Text1.SetFocus
    Exit Sub
End If

'check for fields with spaces
If Left(Text1.Text, 1) = " " Or Left(txt_pass.Text, 1) = " " Then
    result = MsgBox("The username or password cannot be spaces ONLY, or BEGIN with a space", vbCritical, "Error...")
    txt_oldpass.Text = ""
    txt_pass.Text = ""
    txt_pass2.Text = ""
    Text1.Text = ""
    Text1.SetFocus
    Exit Sub
End If
    
If MsgBox("Are you sure that you want to edit this username?", vbYesNo, "Are you sure?") = vbNo Then
    Text1.Text = ""
    txt_oldpass.Text = ""
    txt_pass.Text = ""
    txt_pass2.Text = ""
    Text1.SetFocus
    Exit Sub
End If

user = LCase(Trim(Text1.Text))
pass = LCase(Trim(txt_pass.Text))
newuserandpass = (user & ": " & pass)

Open FileName For Random As #1 Len = 20
result = Encrypt(newuserandpass, 20)
Put #1, Index, newuserandpass
Close #1

Unload Me
frm_users.SetFocus
Call frm_users.Refreshlist

End Sub

Private Sub Command2_Click()
'frm_users.SetFocus
Unload Me
End Sub


Private Sub txt_pass_Change()
If Len(txt_pass.Text) = 8 Then
lbl1.Caption = "The password may only be upto 8 characters"
ElseIf Len(txt_pass.Text) < 8 Then
lbl1.Caption = ""

End If

End Sub


Private Sub txt_pass2_Change()
If Len(txt_pass2.Text) = 8 Then
lbl1.Caption = "The password may only be upto 8 characters"
ElseIf Len(txt_pass2.Text) < 8 Then
lbl1.Caption = ""

End If

End Sub

Private Sub txt_oldpass_Change()
If Len(txt_oldpass.Text) = 8 Then
lbl1.Caption = "The password may only be upto 8 characters"
ElseIf Len(txt_oldpass.Text) < 8 Then
lbl1.Caption = ""

End If

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 10 Then
lbl1.Caption = "The username may only be upto 10 characters"
ElseIf Len(Text1.Text) < 10 Then
lbl1.Caption = ""

End If

End Sub
