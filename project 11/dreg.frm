VERSION 5.00
Begin VB.Form dreg 
   Caption         =   "Donar Regestration"
   ClientHeight    =   6540
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dreg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox invisible 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   32
      Text            =   "no"
      Top             =   -480
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Frame sex 
      Height          =   492
      Left            =   7800
      TabIndex        =   31
      Top             =   6600
      Width           =   4092
      Begin VB.OptionButton optfemale 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   120
         Width           =   1332
      End
      Begin VB.OptionButton optmale 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1212
      End
   End
   Begin VB.OptionButton rno 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10560
      TabIndex        =   16
      Top             =   8520
      Width           =   1332
   End
   Begin VB.OptionButton ryes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7800
      TabIndex        =   15
      Top             =   8520
      Width           =   1212
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "dreg.frx":16CBF
      Left            =   7800
      List            =   "dreg.frx":16CDB
      TabIndex        =   3
      Text            =   "select"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   14
      Top             =   7920
      Width           =   4092
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   13
      Top             =   7320
      Width           =   4092
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   10
      Top             =   6120
      Width           =   4092
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   9
      Top             =   5640
      Width           =   4092
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   8
      Top             =   5160
      Width           =   4092
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   7
      Top             =   4680
      Width           =   4092
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   6
      Top             =   4200
      Width           =   4092
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   5
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   4
      Text            =   "0"
      Top             =   3120
      Width           =   4092
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   2
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   7680
      Left            =   15360
      Picture         =   "dreg.frx":16D11
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3720
   End
   Begin VB.Label Label16 
      Caption         =   "mm/dd/yy"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   33
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Reserved"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   30
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "HB"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Weight"
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
      Left            =   4920
      TabIndex        =   28
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "BP"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Ward No./Optional"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Patient Reserved for/Opt"
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
      Left            =   4920
      TabIndex        =   24
      Top             =   5760
      Width           =   2265
   End
   Begin VB.Label Label8 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   23
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Ticket NO."
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
      Left            =   4920
      TabIndex        =   21
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label Label6 
      Caption         =   "Date of bleeding"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Bag No."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Donar Regestration"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Menu mstore 
      Caption         =   "Store"
      Index           =   1
      Begin VB.Menu msave 
         Caption         =   "Save"
      End
      Begin VB.Menu mdelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnext 
         Caption         =   "Next"
      End
      Begin VB.Menu mprevious 
         Caption         =   "Previous"
      End
   End
   Begin VB.Menu mhome 
      Caption         =   "Home"
      Index           =   2
   End
   Begin VB.Menu madd 
      Caption         =   "Add"
   End
End
Attribute VB_Name = "dreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub ClearBoxes()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text12.Text = ""
Text13.Text = ""

End Sub

Private Sub madd_Click()
Dim bagno As Integer
ClearBoxes
bagno = GetPatientCount
bagno = bagno + 1
Text4.Text = bagno
End Sub

Private Sub mdelete_Click()
Dim n As Integer
    'n = InputBox("Enter name to delete")
    n = Text4.Text
    cn.Open ("DSN=bb")
    Set rs = cn.Execute("delete from donate where bagno = " & n & "")
    cn.Execute ("delete from valid where bagno=" & n & "")
    MsgBox ("Deleted"), vbExclamation, bloodbank
    cn.Close
End Sub

Private Sub mhome_Click(Index As Integer)
Unload Me
Load home
home.Visible = True
End Sub

Private Sub mnext_Click()
 Dim n As Integer
    n = Text4.Text
    n = n + 1
    GetPatientInfo n
End Sub

Private Sub mprevious_Click()
Dim n As Integer
    n = Text4.Text
    n = n - 1
    GetPatientInfo n
End Sub

Private Sub msave_Click()
Dim bagno As Integer, age As Integer, weight As Integer
Dim name As String, hb As String, bp As String, address As String, phone As String, bdate As String, sex As String, ticketno As String, wardno As String, issued As String, reserved As String
   If Text1.Text = "" Or Text2.Text = "" Then
   MsgBox ("unable to save"), vbCritical, bloodbank
   ElseIf Text5.Text = "" Or Text6.Text = "" Then
   MsgBox ("unable to save"), vbCritical, bloodbank
   ElseIf Text7.Text = "" Or Text8.Text = "" Then
   MsgBox ("unable to save"), vbCritical, bloodbank
     Else
name = Text1.Text
address = Text2.Text
hb = Combo1.Text
bagno = Text4.Text
bdate = Text5.Text
ticketno = Text6.Text
phone = Text7.Text
age = Text8.Text
patientreserved = Text9.Text
wardno = Text10.Text
issued = invisible
If optmale.Value = True Then
sex = optmale.Caption
Else
sex = optfemale.Caption
End If
bp = Text12.Text
weight = Text13.Text
If ryes.Value = True Then
reserved = ryes.Caption
Else
reserved = rno.Caption
End If
cn.Open ("DSN=bb")
' cn.execute ("insert into table values (1,'sd',5,'tgh')")
' cn.execute ("insert into table(sno,name) values (1,'sd')")
cn.Execute ("insert into donate values(" & bagno & ",'" & name & "','" & address & "','" & phone & "','" & bdate & "'," & age & ",'" & bp & "', " & weight & ",'" & sex & "','" & patientreserved & "','" & ticketno & "','" & wardno & "','" & hb & "','" & issued & "','" & reserved & "')")
cn.Execute ("update blood set quantity = quantity + 1 where bgroup = '" & hb & "'")
'cn.Execute ("insert into reserved values(" & ticketno & ",'" & issued & "')")
cn.Execute ("insert into valid values('" & hb & "','" & bdate & "'," & bagno & ",'" & issued & "')")
cn.Close
MsgBox "SAVED!", vbExclamation, bloodbank
End If
ClearBoxes
End Sub
Private Function GetPatientCount()
    Dim ct As Integer
    cn.Open ("DSN=bb")
    Set rs = cn.Execute("select count(*) from donate")
    ct = rs.Fields(0)
    cn.Close
    GetPatientCount = ct            ' return ct
End Function
Private Sub GetPatientInfo(n As Integer)
    cn.Open ("DSN=bb")
    Set rs = cn.Execute("select * from donate where bagno = " & n)
    If rs.EOF Then
        MsgBox "No More Records", vbExclamation, bloodbank
    Else
Text1.Text = rs.Fields(1)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(12)
Text4.Text = rs.Fields(0)
Text5.Text = rs.Fields(4)
Text6.Text = rs.Fields(10)
Text7.Text = rs.Fields(3)
Text8.Text = rs.Fields(5)
Text9.Text = rs.Fields(9)
Text10.Text = rs.Fields(11)
If rs.Fields(8) = "Male" Then
    optmale.Value = True
Else
    optfemale.Value = True
End If
Text12.Text = rs.Fields(6)
Text13.Text = rs.Fields(7)
    End If
    cn.Close
End Sub

