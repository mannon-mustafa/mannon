VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmdnrlist 
   Caption         =   "Donar List"
   ClientHeight    =   6990
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   10.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdnrlist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid dgrid 
      Height          =   6975
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   12303
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Donar List"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8265
      TabIndex        =   0
      Top             =   360
      Width           =   2610
   End
   Begin VB.Menu mhome 
      Caption         =   "Home"
   End
   Begin VB.Menu mallow 
      Caption         =   "Allow"
   End
End
Attribute VB_Name = "frmdnrlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mallow_Click()
Unload Me
Load frmrecipient
frmrecipient.Visible = True
End Sub

Private Sub mhome_Click()
Unload Me
Load home
home.Visible = True
End Sub
