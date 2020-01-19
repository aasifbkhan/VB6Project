VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Restaurant Billing System Managment"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11325
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnupass 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnumodify 
         Caption         =   "Modify Data"
         Begin VB.Menu mnufooditems 
            Caption         =   "Add Food Items"
         End
         Begin VB.Menu mnufoodtype 
            Caption         =   "Add Food Type"
         End
      End
   End
   Begin VB.Menu mnubill 
      Caption         =   "Bill"
      Begin VB.Menu mnugeneratebill 
         Caption         =   "Generate Bill"
      End
      Begin VB.Menu mnureport 
         Caption         =   "Report"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   "Option"
      Begin VB.Menu mnulogout 
         Caption         =   "LogOut"
      End
      Begin VB.Menu mnunewuser 
         Caption         =   "New User"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnufooditems_Click()
AddFoodItems.Show
End Sub

Private Sub mnufoodtype_Click()
AddFoodType.Show
End Sub

Private Sub mnugeneratebill_Click()
GenrateBill.Visible = True

End Sub

Private Sub mnupassword_Click()

End Sub

Private Sub mnuprinter_Click()

End Sub

Private Sub mnulogout_Click()
Unload Me
Login.Show
End Sub

Private Sub mnunewuser_Click()
Newuser.Show
End Sub

Private Sub mnupass_Click()
Newuser.Visible = True
End Sub

Private Sub mnureport_Click()
Report.Show
End Sub
