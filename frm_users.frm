VERSION 5.00
Begin VB.Form frm_users 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   1455
   Icon            =   "frm_users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      ItemData        =   "frm_users.frx":0442
      Left            =   0
      List            =   "frm_users.frx":0444
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu mnu_users 
      Caption         =   "&Options"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnu_user_new 
         Caption         =   "New User"
      End
      Begin VB.Menu mnu_bar 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_user_reset 
         Caption         =   "Reset All Users"
      End
   End
   Begin VB.Menu mnu_popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnu_popup_edit 
         Caption         =   "Edit User"
      End
      Begin VB.Menu mnu_popup_delete 
         Caption         =   "Delete User"
      End
   End
End
Attribute VB_Name = "frm_users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recarray(20) As String * 20
Dim buffer(20) As String
Dim username(20) As String * 10
Dim password(20) As String * 8
Dim username2(20) As String
Dim password2(20) As String
Dim temp As Integer
Dim holduser As String
Dim location As Integer
Dim blank As String * 20
Dim movingdata(20) As String * 20

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And List1.ListCount <> 0 Then
Me.PopupMenu mnu_popup
End If
End Sub

Private Sub mnu_popup_delete_Click()

If List1.ListIndex = "-1" Then
result = MsgBox("You must select a user first", vbExclamation, "Select user to delete")
Exit Sub
End If

holduser = List1.List(List1.ListIndex)

If MsgBox("Are you sure you wish to delete the user " & holduser & "?", vbYesNo, "Are you sure?") = vbNo Then
    Exit Sub
End If

If List1.ListCount = 1 Then
    If MsgBox("If you are to delete this user, " & vbCrLf & "and you do not create another before you restart the program " & vbCrLf & "you will need the product activation code supplied with this software to run it." & vbCrLf & vbCrLf & "Are you sure you want to proceed?", vbYesNo, "WARNING...") = vbNo Then
    Exit Sub
    End If
End If



'find place in file
Open filename For Random As #1 Len = 20
For i = 1 To 20
Get #1, i, recarray(i)
Next

For i = 1 To 20
    username(i) = (Left$(recarray(i), 10))
    result = Decrypt(username(i), 10)
    username2(i) = Trim(username(i))
    If username2(i) = holduser Then
    location = i
    End If
Next

Put #1, location, blank

If location = 1 Then
    
    'move all data back a place in the file so as not to confuse program that the dat file is empty
    For i = 2 To 20
        Get #1, i, movingdata(i)
        Put #1, (i - 1), movingdata(i)
    Next
End If

List1.RemoveItem (List1.ListIndex)
Close #1

Call Refreshlist
End Sub

Private Sub mnu_popup_edit_Click()
If List1.ListIndex = "-1" Then
result = MsgBox("You must select a user first", vbExclamation, "Select user to edit")
Exit Sub
End If

For i = 1 To 20
If username2(i) = List1.List(List1.ListIndex) Then
Index = i
End If
Next

frm_useredit.Caption = "Edit User - " & List1.List(List1.ListIndex)
frm_useredit.Text1.Text = List1.List(List1.ListIndex)
frm_useredit.Show 1
End Sub

Private Sub mnu_user_new_Click()

If List1.ListCount = 19 Then
result = MsgBox("A maximum of 20 users are allowed on this program", vbCritical, "Maximum user no. reached")
Exit Sub
End If

frm_newpassword.Show 1
'Unload Me

End Sub

Private Sub mnu_user_reset_Click()
Beep
If MsgBox(("WARNING..." & vbCrLf & vbCrLf & "This will delete all usernames stored by this program." & vbCrLf & vbCrLf & "To be able to create a new user you will need yo know the Authorisation Code supplied with this software" & vbCrLf & "Are you sure you want to continue?"), vbYesNo, "Warning...") = vbYes Then
Kill filename
Unload Me

'Hiding forms seems to load each form if they are not visible and then close them
'which changes a variable wich i dont want changed
'frm_newpassword.Hide <--on loading changes variables

Unload frm_useredit
frm_main.Hide

frm_passwordscreen.txt_user.Text = ""
frm_passwordscreen.txt_Password.Text = ""
frm_passwordscreen.Show 1


Else:
Exit Sub
End If

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()

Call Refreshlist
Me.Left = frm_main.Left - Me.Width
Me.Top = frm_main.Top

End Sub


Public Sub Refreshlist()
Open filename For Random As #1 Len = 20

List1.Clear

For i = 1 To 20
Get #1, i, recarray(i)
Next

For i = 1 To 20
    username(i) = (Left$(recarray(i), 10))
    result = Decrypt(username(i), 10)
    username2(i) = Trim(username(i))
    temp = Asc(username2(i))
    If temp <> "0" Then
    List1.AddItem (username2(i))
    End If

    password(i) = (Right$(recarray(i), 8))
    result = Decrypt(password(i), 8)
    password2(i) = Trim(password(i))
Next
Close #1
'If List1.ListCount >= 1 Then
'List1.ListIndex = 0
'ElseIf List1.ListCount <= 0 Then
'End If
End Sub

Private Sub Form_Terminate()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

