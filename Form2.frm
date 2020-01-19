VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form GenrateBill 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genrate Bill"
   ClientHeight    =   10095
   ClientLeft      =   615
   ClientTop       =   330
   ClientWidth     =   12285
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10095
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   10680
      TabIndex        =   28
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   66650113
      CurrentDate     =   43004
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   4680
      Picture         =   "Form2.frx":1F9902
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   26
      Top             =   6240
      Width           =   1695
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   7080
      TabIndex        =   24
      Top             =   5520
      Width           =   1815
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   5160
      TabIndex        =   21
      Top             =   5520
      Width           =   1815
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   3240
      TabIndex        =   20
      Top             =   5520
      Width           =   1815
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   1320
      TabIndex        =   17
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text6 
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
      Left            =   8400
      TabIndex        =   16
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   6480
      Picture         =   "Form2.frx":1FAC2C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   7320
      Picture         =   "Form2.frx":1FD296
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8040
      Width           =   1455
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   12225
      TabIndex        =   12
      Top             =   9720
      Width           =   12285
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
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
      Left            =   6600
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Picture         =   "Form2.frx":1FF918
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3600
      TabIndex        =   3
      Top             =   1005
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1200
      TabIndex        =   1
      Top             =   1005
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "----Category----"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   9120
      TabIndex        =   25
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   22
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Table.No."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   18360
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Qty."
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   18360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Restaurant Billing System Managment"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   13815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill.No"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "GenrateBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rsmax As New ADODB.Recordset

Public Sub data()
If con.State = adStateOpen Then
con.Close
End If

Set rs5 = New ADODB.Recordset
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
rs.Open "select * from Food_Item", con, adOpenDynamic, adLockOptimistic
rs1.Open "select * from Bill", con, adOpenDynamic, adLockOptimistic
rs3.Open "select * from Food_type", con, 3, 3
rs2.Open "select Item_name from Food_Item where F_type='" & Combo1.Text & "'", con, adOpenDynamic, adLockOptimistic
rs4.Open "select Rate from Food_Item where Item_name='" & List1.Text & "'", con, adOpenDynamic, adLockOptimistic
rsmax.Open "select max(Bill_no) from Bill ", con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub Combo1_Click()
List1.Clear
Call data
Do Until rs2.EOF
List1.AddItem rs2!Item_name
rs2.MoveNext
Loop
Set rs2 = Nothing
con.Close

End Sub

Private Sub Command1_Click()
Call data
n = 0
For i = 0 To List2.ListCount - 1
rs1.AddNew
rs1!Bill_no = Text1.Text
If n = 0 Then
rs1!T_no = Text2.Text
rs1!Total_Amount = Text4.Text
rs1!Bill_Date = DTPicker1.Value
rs1!dummy = Text1.Text
End If
n = 1
rs1!Item = List2.List(i)
rs1!Qty = List3.List(i)
rs1!Rate = List4.List(i)
rs1!Amount = List5.List(i)
rs1.Update

rs1.MoveFirst
While Not rs1.EOF
If rs1!Item = List2.List(i) Then
rs1!Qty = rs1!Qty
rs1.Update
End If
rs1.MoveNext
Wend

Next
MsgBox "Order successfully", vbInformation

rs5.Open "select * from Bill where Bill_no=" & Text1.Text & " ", con, adOpenDynamic, adLockBatchOptimistic

Set DataReport1.DataSource = rs5
DataReport1.Sections("Section1").Controls(1).DataField = rs1.Fields(8).Name
DataReport1.Sections("Section1").Controls(2).DataField = rs1.Fields(1).Name
DataReport1.Sections("Section1").Controls(3).DataField = rs1.Fields(2).Name
DataReport1.Sections("Section1").Controls(4).DataField = rs1.Fields(3).Name
DataReport1.Sections("Section1").Controls(5).DataField = rs1.Fields(4).Name
DataReport1.Sections("Section1").Controls(6).DataField = rs1.Fields(5).Name
DataReport1.Sections("Section1").Controls(7).DataField = rs1.Fields(6).Name

DataReport1.Show
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Call data

Text1.Text = rsmax.Fields(0).Value + 1
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


Combo1.Text = "----Catagory----"


con.Close
End Sub

Private Sub Command4_Click()

List2.AddItem (List1.Text)
List3.AddItem (Text3.Text)

List4.AddItem (Text5.Text)
Text6.Text = Text5.Text * Text3.Text
List5.AddItem (Text6.Text)
Text4.Text = Text6.Text + Val(Text4.Text)

List1.Clear
Text3.Text = " "
Text5.Text = " "


End Sub

Private Sub Form_Load()
Call data
rs3.MoveFirst
While Not rs3.EOF

Combo1.AddItem (rs3.Fields(1))

rs3.MoveNext
Wend
Text1.Text = rsmax.Fields(0).Value + 1
con.Close


End Sub

Private Sub List1_Click()
Call data
Text5.Text = rs4.Fields("Rate").Value
con.Close
End Sub
Private Sub Timer1_Timer()
If Timer1.Enabled = True Then



Text7.Text = Date
End If

End Sub

Private Sub Text3_Change()
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= 32 And KeyAscii <= 47 Or KeyAscii >= 58 And KeyAscii <= 127 Then
MsgBox "Character and Special Characters Not Allowed", vbCritical
Text8.Text = ""
Text8.SetFocus
KeyAscii = 0
End If
End Sub
