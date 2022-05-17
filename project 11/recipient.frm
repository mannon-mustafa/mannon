VERSION 5.00
Begin VB.Form frmrecipient 
   Caption         =   "Recipient"
   ClientHeight    =   5550
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8895
   Icon            =   "recipient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Ticket"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1932
      Left            =   8040
      TabIndex        =   22
      Top             =   4680
      Width           =   5532
      Begin VB.CommandButton cmdcheck 
         Caption         =   "Check"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1080
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.CommandButton cmdsub 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   3480
         TabIndex        =   25
         Top             =   1200
         Width           =   1692
      End
      Begin VB.TextBox txttkt 
         Height          =   612
         Left            =   2760
         TabIndex        =   24
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Ticket No."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1812
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "check"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1812
      Left            =   7320
      TabIndex        =   18
      Top             =   4800
      Width           =   6852
      Begin VB.CommandButton cmdno 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3720
         TabIndex        =   21
         Top             =   1200
         Width           =   1812
      End
      Begin VB.CommandButton cmdyes 
         Caption         =   "YES"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1320
         TabIndex        =   20
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Have You Blood Reserved Here ? "
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Width           =   5892
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Interview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3372
      Left            =   7320
      TabIndex        =   16
      Top             =   4320
      Width           =   6972
      Begin VB.TextBox textb 
         Height          =   372
         Left            =   2520
         TabIndex        =   13
         Top             =   2880
         Width           =   1332
      End
      Begin VB.ComboBox combog 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "recipient.frx":16CBF
         Left            =   2520
         List            =   "recipient.frx":16CDB
         TabIndex        =   10
         Text            =   "select"
         Top             =   1560
         Width           =   4092
      End
      Begin VB.TextBox Text10 
         Height          =   372
         Left            =   2520
         TabIndex        =   12
         Top             =   2520
         Width           =   4092
      End
      Begin VB.TextBox Text8 
         Height          =   372
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "1"
         Top             =   2040
         Width           =   4092
      End
      Begin VB.TextBox Text6 
         Height          =   372
         Left            =   2520
         TabIndex        =   8
         Top             =   600
         Width           =   4092
      End
      Begin VB.TextBox Text5 
         Height          =   372
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
         Width           =   4092
      End
      Begin VB.Label Label15 
         Caption         =   "Bag No."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   37
         Top             =   2880
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   30
         Top             =   2520
         Width           =   1812
      End
      Begin VB.Label Label7 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   1812
      End
      Begin VB.Label Label6 
         Caption         =   "HB"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1812
      End
      Begin VB.Label Label5 
         Caption         =   "Ticket No."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   27
         Top             =   1080
         Width           =   1812
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Queries for General Recipient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   5052
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Interview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3732
      Left            =   7200
      TabIndex        =   14
      Top             =   4080
      Width           =   6972
      Begin VB.ComboBox combor 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "recipient.frx":16D11
         Left            =   2520
         List            =   "recipient.frx":16D2D
         TabIndex        =   3
         Text            =   "select"
         Top             =   1680
         Width           =   4092
      End
      Begin VB.OptionButton optno 
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4320
         TabIndex        =   7
         Top             =   3120
         Width           =   1452
      End
      Begin VB.OptionButton optyes 
         Caption         =   "YES"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2520
         TabIndex        =   6
         Top             =   3120
         Width           =   1452
      End
      Begin VB.TextBox Text9 
         Height          =   372
         Left            =   2520
         TabIndex        =   5
         Top             =   2520
         Width           =   4092
      End
      Begin VB.TextBox Text7 
         Height          =   372
         Left            =   2520
         TabIndex        =   4
         Top             =   2040
         Width           =   4092
      End
      Begin VB.TextBox Text2 
         Height          =   372
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   4092
      End
      Begin VB.TextBox Text1 
         Height          =   372
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   4092
      End
      Begin VB.Label Label14 
         Caption         =   "Issued"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   36
         Top             =   3120
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   35
         Top             =   2640
         Width           =   1812
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   34
         Top             =   2160
         Width           =   1812
      End
      Begin VB.Label Label11 
         Caption         =   "HB"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   33
         Top             =   1680
         Width           =   1812
      End
      Begin VB.Label Label10 
         Caption         =   "Ticket No."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   32
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label Label9 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label reserved 
         Alignment       =   2  'Center
         Caption         =   "Queries for Reserved Blood"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   5052
      End
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   10080
      Picture         =   "recipient.frx":16D63
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label rinterview 
      Alignment       =   2  'Center
      Caption         =   "Recipient Interview"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      TabIndex        =   0
      Top             =   2160
      Width           =   7215
   End
   Begin VB.Menu mcheck 
      Caption         =   "Check"
      Index           =   1
      Begin VB.Menu mallowgeneral 
         Caption         =   "Allow General"
      End
      Begin VB.Menu mallowreserved 
         Caption         =   "Allow Reserved"
      End
   End
   Begin VB.Menu mback 
      Caption         =   "Back"
      Index           =   2
      Begin VB.Menu mhome 
         Caption         =   "Home"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmrecipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdcheck_Click()
Dim s As String
s = txttkt.Text
cn.Open ("DSN=bb")
Set rs = cn.Execute("select reserved from donate where ticketno= '" & s & "'")
If rs.Fields(0) = "Yes" Then
 MsgBox ("you have blood reserved"), vbInformation, bloodbank
   Frame1.Visible = True
   Frame2.Visible = False
   Frame4.Visible = False
   Frame3.Visible = False
 Else
   MsgBox ("you have no blood reserved"), vbExclamation, bloodbank
   Frame2.Visible = True
   Frame4.Visible = False
   Frame3.Visible = False
End If
cn.Close
End Sub

Private Sub cmdno_Click()
Frame1.Visible = False
Frame3.Visible = False
Frame2.Visible = True
End Sub

Private Sub cmdsub_Click()
Dim s As String
Dim a As String
Dim b As String
'Dim c As String
'c = "yes"
s = txttkt.Text
cn.Open ("DSN=bb")
Set rs = cn.Execute("select  issued from donate where ticketno= '" & s & "'")
  If rs.EOF Then
  MsgBox ("invalid"), vbCritical, bloodbank
  ElseIf rs.Fields(0) = "yes" Then
  MsgBox ("Blood Issued"), vbCritical, bloodbank
  Else
  MsgBox ("ok check it"), vbExclamation, bloodbank
  End If
  cmdcheck.Visible = True
cn.Close
'cn.Execute ("update donate set reserved='yes', issued='no', addrtes='jjj' where tktno=1")
End Sub

Private Sub cmdyes_Click()
Frame3.Visible = False
Frame2.Visible = False
Frame1.Visible = False
Frame4.Visible = True
End Sub

Private Sub Form_Load()
Frame3.Visible = True
Frame1.Visible = False
Frame2.Visible = False
Frame4.Visible = False
End Sub

Private Sub mallowgeneral_Click()
Dim name2 As String, ticketno As String, rdate2 As String, hb2 As String
Dim tk2 As String
Dim qt2 As Integer, bag As Integer
Dim s As String
If bag = 0 Then
MsgBox ("Bag Number Needed"), vbCritical, bloodbank
Else
bag = textb.Text
s = txttkt.Text
ticketno = s
name2 = Text6.Text
tk2 = Text5.Text
hb2 = combog.Text
qt2 = Text8.Text
rdate2 = Text10.Text

cn.Open ("DSN=bb")
cn.Execute ("insert into recipient values ('" & name2 & "' , '" & tk2 & "' , '" & hb2 & "' , " & qt2 & " , '" & rdate2 & "')")
cn.Execute ("update donate set issued='yes' where bagno= " & bag & "")
cn.Execute ("update valid set issued ='yes'where bagno=" & bag & "")
Set rs = cn.Execute("select quantity from blood where bgroup='" & hb2 & "'")
If rs.Fields(0) = 0 Then
 MsgBox ("NoT AvaiL in StocK"), vbCritical, bloodbank
Else
cn.Execute ("update blood set quantity=quantity-1 where bgroup ='" & hb & "'")
End If
        'cn.Execute ("select bagno from donate where hb='" & hb2 & "'")

cn.Close
MsgBox ("Blood Provided"), vbExclamation, bloodbank
End If
combog.Text = ""
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""
Text10.Text = ""
End Sub
'cn.Execute ("update donate set reserved='yes', issued='no', addrtes='jjj' where tktno=1")
Private Sub mallowreserved_Click()
Dim name As String, ticketno As String, rdate As String, hb As String, issued As String
Dim tk As Integer
Dim qt As Integer
Dim s As String
s = txttkt.Text
ticketno = s
name = Text1.Text
tk = Text2.Text
hb = combor.Text
qt = Text7.Text
rdate = Text9.Text
If optyes.Value = True Then
issued = optyes.Caption
Else
issued = optno.Caption
End If
cn.Open ("DSN=bb")
cn.Execute (" insert into recipient values ('" & name & "' , " & tk & " , '" & hb & "' , '" & qt & "' , " & rdate & ")")
'cn.Execute ("insert into donate (issued) values('" & issued & "')")
cn.Execute ("update donate set reserved='no', issued='yes' where ticketno= '" & s & "'")
cn.Execute ("update valid set issued ='yes'where bagno=" & bag & "")
Set rs = cn.Execute("select quantity from blood where bgroup='" & hb2 & "'")
If rs.Fields(0) = 0 Then
 MsgBox ("NoT AvaiL in StocK"), vbCritical, bloodbank
Else
cn.Execute ("update blood set quantity=quantity-1 where bgroup ='" & hb & "'")
End If
cn.Close
MsgBox ("Blood Provided"), vbExclamation, bloodbank
Text1.Text = ""
Text2.Text = ""
combor.Text = ""
Text7.Text = ""
Text9.Text = ""
End Sub

Private Sub mexit_Click()
Unload Me
Load MDIForm1
MDIForm1.Visible = True
End Sub

Private Sub mhome_Click()
Unload Me
Load home
home.Visible = True
End Sub


