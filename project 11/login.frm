VERSION 5.00
Begin VB.Form access 
   Caption         =   "Access"
   ClientHeight    =   6420
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8685
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame guest 
      Caption         =   "Guest"
      Height          =   4212
      Left            =   7080
      TabIndex        =   14
      Top             =   4080
      Width           =   6012
      Begin VB.CommandButton userclear 
         Caption         =   "Clear Exesting"
         Height          =   732
         Left            =   720
         TabIndex        =   21
         Top             =   3240
         Width           =   2292
      End
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "Submit"
         Height          =   732
         Left            =   3600
         TabIndex        =   19
         Top             =   3240
         Width           =   2172
      End
      Begin VB.TextBox Text4 
         Height          =   732
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   2040
         Width           =   3612
      End
      Begin VB.TextBox Text3 
         Height          =   732
         Left            =   2280
         TabIndex        =   17
         Top             =   840
         Width           =   3612
      End
      Begin VB.Label lblpassword 
         Caption         =   "Password"
         Height          =   975
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lbluser 
         Caption         =   "User Name"
         Height          =   972
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1932
      End
   End
   Begin VB.Frame login 
      Caption         =   "Login"
      Height          =   1452
      Left            =   4920
      TabIndex        =   11
      Top             =   4800
      Width           =   6255
      Begin VB.CommandButton bttnguest 
         Caption         =   "Guest"
         Height          =   492
         Left            =   3240
         TabIndex        =   13
         Top             =   600
         Width           =   2292
      End
      Begin VB.CommandButton bttnadmin 
         Caption         =   "Admin"
         Height          =   492
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   2292
      End
   End
   Begin VB.Frame admin1 
      Caption         =   "Admin"
      Height          =   3612
      Left            =   6600
      TabIndex        =   0
      Top             =   4440
      Width           =   6732
      Begin VB.CommandButton clear 
         Caption         =   "Clear Exesting"
         Height          =   732
         Left            =   600
         TabIndex        =   20
         Top             =   2640
         Width           =   2172
      End
      Begin VB.CommandButton submit 
         Caption         =   "Submit"
         Height          =   492
         Left            =   3840
         TabIndex        =   5
         Top             =   2760
         Width           =   1692
      End
      Begin VB.TextBox Text2 
         Height          =   612
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1680
         Width           =   3252
      End
      Begin VB.TextBox Text1 
         Height          =   612
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Width           =   3252
      End
      Begin VB.Label password 
         Caption         =   "Password"
         Height          =   612
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   2052
      End
      Begin VB.Label username 
         Caption         =   "User Name"
         Height          =   612
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1812
      End
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   17160
      Picture         =   "login.frx":16CBF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2640
   End
   Begin VB.Label Label2 
      Caption         =   "9906626711/9622639683"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   23
      Top             =   9960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   360
      OLEDropMode     =   1  'Manual
      Picture         =   "login.frx":2C98A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SMHS Hospital Srinagar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4320
      TabIndex        =   22
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label address 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10095
      TabIndex        =   10
      Top             =   840
      Width           =   105
   End
   Begin VB.Label copyright 
      AutoSize        =   -1  'True
      Caption         =   "@Copyright:AHMF Solutions"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   9
      Top             =   9600
      Width           =   2685
   End
   Begin VB.Label email 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "E-mail:bloodbankpul@gmail.com"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8640
      TabIndex        =   8
      Top             =   9960
      Width           =   3150
   End
   Begin VB.Label fax 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Contact:-01933245121 / Fax:-432167543"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8520
      TabIndex        =   7
      Top             =   9600
      Width           =   3540
   End
   Begin VB.Label bloodbank 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Bank"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5760
      TabIndex        =   6
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub bttnadmin_Click()
login.Visible = False
admin1.Visible = True
End Sub

Private Sub bttnguest_Click()
login.Visible = False
guest.Visible = True
End Sub

Private Sub clear_Click()
Dim p As String
Dim u As String
p = Text1.Text
u = Text2.Text
cn.Open ("DSN=bb")
Set rs = cn.Execute("delete from permit where uname = '" & u & "' ")
MsgBox ("Deleted"), vbExclamation, bloodbank
Text1.Text = ""
Text2.Text = ""
cn.Close
End Sub

Private Sub cmdsubmit_Click()
Dim u As String
Dim p As String
Dim Y As String
Dim z As String
u = Text3.Text
p = Text4.Text
cn.Open ("DSN=bb")
Set rs = cn.Execute("select count(*) from permit2 where uname = '" & u & "' and paswrd = '" & p & "'")
While Not rs.EOF
Y = rs.Fields(0)
'z = rs.Fields(1)
'If Y = u And z = p Then
If Y = 1 Then
user = "G"
Unload Me
Load MDIForm1
MDIForm1.Visible = True
home.homeframe2.Visible = False
home.user.Visible = False
Else
MsgBox ("Sorry Wrong Password"), vbCritical, "Wrong!"
End If
rs.MoveNext
Wend
cn.Close
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
admin1.Visible = False
guest.Visible = False
login.Visible = True
End Sub




Private Sub mexit_Click()
Unload Me
Load MDIForm1
MDIForm1.Visible = True
End Sub

Private Sub submit_Click()
Dim u As String
Dim p As String
Dim Y As String
Dim z As String
u = Text1.Text
p = Text2.Text
cn.Open ("DSN=bb")
Set rs = cn.Execute("select count(*) from permit where uname = '" & u & "'and paswrd='" & p & "'")
While Not rs.EOF
Y = rs.Fields(0)
'z = rs.Fields(1)
'If Y = u And z = p Then
If Y = 1 Then
Unload Me
Load MDIForm1
MDIForm1.Visible = True
Else
MsgBox ("Sorry Wrong Password"), vbCritical, "Wrong!"
End If
rs.MoveNext
Wend
cn.Close
End Sub

Private Sub userclear_Click()
Dim p As String
Dim u As String
p = Text1.Text
u = Text2.Text
cn.Open ("DSN=bb")
Set rs = cn.Execute("delete from permit2 where uname = '" & u & "' ")
MsgBox ("Deleted"), vbExclamation, bloodbank
cn.Close
Text3.Text = ""
Text4.Text = ""
End Sub
