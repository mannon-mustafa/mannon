VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form home 
   Caption         =   "Home"
   ClientHeight    =   7245
   ClientLeft      =   3465
   ClientTop       =   1980
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "home.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexpire 
      Caption         =   "Expire"
      Height          =   615
      Left            =   13440
      TabIndex        =   14
      Top             =   2400
      Width           =   1932
   End
   Begin MSFlexGridLib.MSFlexGrid g1 
      Height          =   5175
      Left            =   13440
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   12360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   11760
      Top             =   2400
   End
   Begin VB.Frame homeframe2 
      Height          =   5892
      Left            =   15840
      TabIndex        =   7
      Top             =   2280
      Width           =   2052
      Begin VB.CommandButton bloodavailable 
         Caption         =   "Blood Available"
         Height          =   732
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   1812
      End
      Begin VB.CommandButton donar2 
         Caption         =   "Donar List"
         Height          =   732
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1812
      End
      Begin VB.CommandButton recipient2 
         Caption         =   "Recipient List"
         Height          =   732
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1812
      End
      Begin VB.CommandButton Admin 
         Caption         =   "Admin"
         Height          =   852
         Left            =   120
         TabIndex        =   8
         Top             =   4560
         Width           =   1812
      End
   End
   Begin VB.Frame homeframe1 
      Height          =   5892
      Left            =   3960
      TabIndex        =   2
      Top             =   2520
      Width           =   2052
      Begin VB.CommandButton user 
         Caption         =   "User"
         Height          =   852
         Left            =   120
         TabIndex        =   6
         Top             =   4680
         Width           =   1812
      End
      Begin VB.CommandButton recipient 
         Caption         =   "Recipient"
         Height          =   732
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1812
      End
      Begin VB.CommandButton donar 
         Caption         =   "Donar"
         Height          =   732
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1812
      End
      Begin VB.CommandButton gallery 
         Caption         =   "Gallery"
         Height          =   732
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   1812
      End
   End
   Begin VB.Label lblevents 
      Caption         =   "Gallery/Events"
      Height          =   495
      Left            =   8640
      TabIndex        =   12
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   6480
      Picture         =   "home.frx":16CBF
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   6495
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   13320
      Picture         =   "home.frx":1824C
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   7440
      Picture         =   "home.frx":19A46
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label address 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "SMHS Hospital Srinagar"
      Height          =   315
      Left            =   9360
      TabIndex        =   1
      Top             =   1440
      Width           =   3435
   End
   Begin VB.Label bloodbank 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Blood Bank"
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Menu madd 
      Caption         =   "Add"
      Visible         =   0   'False
   End
   Begin VB.Menu mdel 
      Caption         =   "Del"
      Visible         =   0   'False
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As ADODB.Recordset
Dim i As Integer

Private Sub admin_Click()
Load Change
Change.Visible = True
Change.usersetnew.Visible = False
Change.setnew.Visible = True
End Sub

Private Sub bloodavailable_Click()
Unload Me
Load frmblood
frmblood.Visible = True
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim c As Integer
r = 1
c = 2
frmblood.bgrid.Rows = r
frmblood.bgrid.Cols = c
frmblood.bgrid.Row = 0
frmblood.bgrid.Col = 0
frmblood.bgrid.Text = "BGroup"
frmblood.bgrid.Col = 1
frmblood.bgrid.Text = "Quantity"
cn.Open ("DSN=bb")
Set rs = cn.Execute("select * from blood")
While Not rs.EOF
frmblood.bgrid.Rows = frmblood.bgrid.Rows + 1
frmblood.bgrid.Row = r
For c = 0 To rs.Fields.Count - 1
frmblood.bgrid.Col = c
frmblood.bgrid.Text = rs.Fields(c)
Next
rs.MoveNext
r = r + 1
Wend
r = r + 1
cn.Close
End Sub

Private Sub cmdexpire_Click()
Dim r As Integer
Dim c As Integer
Dim d As Date
d = Date - 30
r = 1
g1.Rows = 1
g1.Cols = 1
g1.Row = 0
g1.Col = 0
g1.Text = "Bagno"
cn.Open ("DSN=bb")
Set rs = cn.Execute("select bagno from valid where bdate < #" & d & "# and issued = 'no'")
While Not rs.EOF
g1.Rows = g1.Rows + 1
g1.Row = r
g1.Col = 0
g1.Text = rs.Fields(0)
rs.MoveNext
r = r + 1
Wend
cn.Close
End Sub

Private Sub donar_Click()
Unload Me
Load dreg
dreg.Visible = True
End Sub

Private Sub donar2_Click()
Unload Me
Load frmdnrlist
frmdnrlist.Visible = True
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim c As Integer
r = 1
c = 15
frmdnrlist.dgrid.Rows = r
frmdnrlist.dgrid.Cols = c
frmdnrlist.dgrid.Row = 0
frmdnrlist.dgrid.Col = 0
frmdnrlist.dgrid.Text = "Bagno"
frmdnrlist.dgrid.Col = 1
frmdnrlist.dgrid.Text = "Name"
frmdnrlist.dgrid.Col = 2
frmdnrlist.dgrid.Text = "Address"
frmdnrlist.dgrid.Col = 3
frmdnrlist.dgrid.Text = "Phone"
frmdnrlist.dgrid.Col = 4
frmdnrlist.dgrid.Text = "Bdate"
frmdnrlist.dgrid.Col = 5
frmdnrlist.dgrid.Text = "Age"
frmdnrlist.dgrid.Col = 6
frmdnrlist.dgrid.Text = "BP"
frmdnrlist.dgrid.Col = 7
frmdnrlist.dgrid.Text = "Weight"
frmdnrlist.dgrid.Col = 8
frmdnrlist.dgrid.Text = "Sex"
frmdnrlist.dgrid.Col = 9
frmdnrlist.dgrid.Text = "Patient reserved for"
frmdnrlist.dgrid.Col = 10
frmdnrlist.dgrid.Text = "Ticketno."
frmdnrlist.dgrid.Col = 11
frmdnrlist.dgrid.Text = "Wardno"
frmdnrlist.dgrid.Col = 12
frmdnrlist.dgrid.Text = "HB"
frmdnrlist.dgrid.Col = 13
frmdnrlist.dgrid.Text = "issued"
frmdnrlist.dgrid.Col = 14
frmdnrlist.dgrid.Text = "reserved"
cn.Open ("DSN=bb")
Set rs = cn.Execute("select * from donate")
While Not rs.EOF
frmdnrlist.dgrid.Rows = frmdnrlist.dgrid.Rows + 1
frmdnrlist.dgrid.Row = r
For c = 0 To rs.Fields.Count - 1
frmdnrlist.dgrid.Col = c
frmdnrlist.dgrid.Text = rs.Fields(c)
'frmdnrlist.dgrid.CellTop
Next
rs.MoveNext
r = r + 1
Wend
r = r + 1
cn.Close
End Sub

Private Sub Form_Click()
Dim f As String
cd1.ShowOpen
cd1.Filter = "images|.jpg"
f = cd1.FileName
If f = "" Then
MsgBox "NO file was selected", vbCritical, "Warning!"
End If
End Sub

Private Sub Form_Load()
    i = 1
    
    If Module1.user = "G" Then
     Load MDIForm1
     MDIForm1.Visible = True
     home.homeframe2.Visible = False
     home.user.Visible = False
    Else
        Load MDIForm1
        MDIForm1.Visible = True
        home.homeframe2.Visible = True
        home.user.Visible = True
    End If
End Sub

Private Sub gallery_Click()
Timer1.Enabled = False
Image1.Visible = False
lblevents.Visible = False
End Sub
Private Sub gallery_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Timer1.Enabled = True
Image1.Visible = True
lblevents.Visible = True
End If
End Sub

Private Sub recipient_Click()
MsgBox ("First See Blood Available"), vbInformation, bloodbank
End Sub

Private Sub recipient2_Click()
Unload Me
Load frmdnrlist
frmrcplist.Visible = True
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim c As Integer
r = 1
c = 5
frmrcplist.rgrid.Rows = r
frmrcplist.rgrid.Cols = c
frmrcplist.rgrid.Row = 0
frmrcplist.rgrid.Col = 0
frmrcplist.rgrid.Text = "Patient Name"
frmrcplist.rgrid.Col = 1
frmrcplist.rgrid.Text = "Ticket No."
frmrcplist.rgrid.Col = 2
frmrcplist.rgrid.Text = "HB"
frmrcplist.rgrid.Col = 3
frmrcplist.rgrid.Text = "Quantity"
frmrcplist.rgrid.Col = 4
frmrcplist.rgrid.Text = "Date"
cn.Open ("DSN=bb")
Set rs = cn.Execute("select * from recipient")
While Not rs.EOF
frmrcplist.rgrid.Rows = frmrcplist.rgrid.Rows + 1
frmrcplist.rgrid.Row = r
For c = 0 To rs.Fields.Count - 1
frmrcplist.rgrid.Col = c
frmrcplist.rgrid.Text = rs.Fields(c)
Next
rs.MoveNext
r = r + 1
Wend
r = r + 1
cn.Close
End Sub

Private Sub Timer1_Timer()
If i = 13 Then
i = 1
End If
Image1.Picture = LoadPicture("C:\Users\SONY\project\events\" & i & ".jpg")
i = i + 1
End Sub

Private Sub user_Click()
Unload Me
Load Change
Change.Visible = True
Change.setnew.Visible = False
End Sub
