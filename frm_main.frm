VERSION 5.00
Begin VB.Form frm_main 
   ClientHeight    =   1800
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frm_main.frx":0000
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Developed and Programmed by James Ricketts"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnu_users 
      Caption         =   "Users"
   End
   Begin VB.Menu mnu_logoff 
      Caption         =   "Log Off"
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
updatecaption
End Sub

Private Sub mnu_logoff_Click()
Unload frm_useredit         '
Unload frm_users

'unloading main form = END (so dont!)
Me.Hide
frm_passwordscreen.txt_user.Text = ""
frm_passwordscreen.txt_Password.Text = ""
frm_passwordscreen.Show 1
End Sub

Private Sub mnu_users_Click()
frm_users.Show
End Sub
Public Sub updatecaption()
Me.Caption = loggedin
End Sub
