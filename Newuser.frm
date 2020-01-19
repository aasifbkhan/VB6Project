VERSION 5.00
Begin VB.Form Newuser 
   Caption         =   "Create New User"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   Picture         =   "Newuser.frx":0000
   ScaleHeight     =   3900
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   4320
      Picture         =   "Newuser.frx":43092
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2760
      Picture         =   "Newuser.frx":43AFC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   240
      Picture         =   "Newuser.frx":45566
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UID"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Create New User"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "Newuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
rs.Open "select * from login", con, adOpenDynamic, adLockOptimistic
'call data
rs.AddNew
rs.Fields("UName").Value = Text1.Text
rs.Fields("Password").Value = Text2.Text
MsgBox "New Account Created Successfully", vbInformation
rs.Update
con.Close
End Sub

Private Sub Command2_Click()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
rs.Open "select * from login where UID=" & Val(Text3.Text), con, adOpenDynamic, adLockOptimistic
rs.Fields("UName").Value = Text1.Text
rs.Fields("Password").Value = Text2.Text
MsgBox "Password Changed Successfully", vbInformation
rs.Update
con.Close
End Sub

Private Sub Command3_Click()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
rs.Open "select * from login where UID=" & (Text3.Text), con, adOpenDynamic, adLockOptimistic
If Text3.Text = rs.Fields("UID") Then
Text1.Text = rs.Fields("UName").Value
Text2.Text = rs.Fields("Password").Value
End If
End Sub
