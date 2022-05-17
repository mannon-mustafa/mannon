VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Blood Bank"
   ClientHeight    =   9375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   19095
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mstart 
      Caption         =   "Start"
      Begin VB.Menu mopen 
         Caption         =   "OPen"
      End
      Begin VB.Menu mclose 
         Caption         =   "Close"
      End
      Begin VB.Menu maccass 
         Caption         =   "Access"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub maccass_Click()
Unload home
Load access
access.Visible = True
End Sub

Private Sub mclose_Click()
Unload Me
End Sub

Private Sub mopen_Click()
Load home
home.Visible = True
End Sub

Private Sub Picture1_Click()
Image2.Picture = LoadPicture("C:\Users\SONY\project\events\" & 1 & ".jpg")
End Sub
