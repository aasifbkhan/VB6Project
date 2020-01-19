VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Report 
   Caption         =   "Report"
   ClientHeight    =   2610
   ClientLeft      =   4425
   ClientTop       =   5250
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   Picture         =   "Report.frx":0000
   ScaleHeight     =   2610
   ScaleWidth      =   6135
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   3000
      Picture         =   "Report.frx":10916
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Bright"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   67698689
      CurrentDate     =   43005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Public Sub data()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
End Sub

Private Sub Command1_Click()
Call data

Dim str As String
str = DTPicker1.Value
rs.Open "select * from Bill where Bill_Date= '" & str & "' ", con, adOpenDynamic, adLockOptimistic

If rs.EOF = True And rs.BOF = True Then
MsgBox "Record Not Found"
End If

Set DataReport2.DataSource = rs
DataReport2.Sections("Section1").Controls(1).DataField = rs.Fields(8).Name
DataReport2.Sections("Section1").Controls(2).DataField = rs.Fields(6).Name
DataReport2.Show
End Sub

