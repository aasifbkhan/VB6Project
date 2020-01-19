VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AddFoodItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Food Items"
   ClientHeight    =   10665
   ClientLeft      =   810
   ClientTop       =   525
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ModifyData.frx":0000
   ScaleHeight     =   10665
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
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
      Height          =   735
      Left            =   6360
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
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
      Left            =   6360
      TabIndex        =   13
      Top             =   3840
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   11760
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Restaurant Billing System\project\Restaurant1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Restaurant Billing System\project\Restaurant1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Food_Item"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ModifyData.frx":1F9902
      Height          =   3615
      Left            =   2760
      TabIndex        =   9
      Top             =   6840
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Bright"
         Size            =   12.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   9480
      Picture         =   "ModifyData.frx":1F9917
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   9000
      Picture         =   "ModifyData.frx":1FBA49
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   6480
      Picture         =   "ModifyData.frx":1FD133
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   4320
      Picture         =   "ModifyData.frx":1FEB9D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   5
      Top             =   4800
      Width           =   2895
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
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "F_type"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "ModifyData.frx":200223
      Left            =   6360
      List            =   "ModifyData.frx":200225
      TabIndex        =   1
      Text            =   " --Select-Category--"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Maintain  Your Record Or Data Here "
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
      Height          =   615
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Width           =   11055
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   14400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FIID.-"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14400
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rate -"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item name -"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Category"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "AddFoodItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsmax As New ADODB.Recordset
Public Sub data()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
rs.Open "select * from Food_Item", con, adOpenDynamic, adLockOptimistic
rs1.Open "select * from Food_type", con, adOpenDynamic, adLockOptimistic
rs2.Open "select F_type from Food_type where F_type='" & Combo1.Text & "'", con, adOpenDynamic, adLockOptimistic
rsmax.Open "select max(FI_Id) from Food_Item", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Combo1_Click()
Call data
Text4.Text = rs2.Fields("F_type").Value
con.Close
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Please Enter Some Data "
Else

Call data
rs.AddNew
rs!FI_Id = Text1.Text
rs!Item_name = Text2.Text
rs!F_type = Text4.Text
rs!Rate = Text3.Text
rs.Update
MsgBox "Food Item Added Sucessfully", vbInformation


Text2.Text = ""
Text4.Text = ""
Text3.Text = ""
End If
Adodc2.Refresh
con.Close
Call data
Text1.Text = rsmax.Fields(0).Value + 1

con.Close
End Sub

Private Sub Command2_Click()
Dim n As Integer
n = 0
If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text3.Text = "" Then
MsgBox "Null Value is not Allowed Please Fill The Details", vbCritical

Else

Call data
rs.MoveFirst
While Not rs.EOF
If Text1.Text = rs.Fields(0).Value Then
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text4.Text
rs.Fields(3).Value = Text3.Text

n = 1

rs.Update
MsgBox "Update Successfully", vbInformation
Text2.Text = ""
Text4.Text = ""
Text3.Text = ""

End If

rs.MoveNext
Wend

If n = 0 Then

MsgBox "Search Record First", vbCritical

End If
con.Close
End If
Adodc2.Refresh
Call data
Text1.Text = rsmax.Fields(0).Value + 1

con.Close
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text3.Text = "" Then
MsgBox "Null Value is not Allowed Please Fill The Details", vbCritical
Else
Call data
rs.MoveFirst
While Not rs.EOF

If Text1.Text = rs.Fields(0).Value Then
rs.Delete
End If
rs.MoveNext
Wend
con.Close
Text2.Text = ""
Text4.Text = ""
Text3.Text = ""
End If
Adodc2.Refresh
End Sub

Private Sub Command4_Click()
Dim n As Integer
n = 0
If Text1.Text = "" Then
MsgBox "Please Enter Food ID ", vbCritical
Else



Call data
rs.MoveFirst
While Not rs.EOF

If Text1.Text = rs.Fields(0).Value Then

Text2.Text = rs.Fields(1).Value
Text4.Text = rs.Fields(2).Value
Text3.Text = rs.Fields(3).Value
n = 1
MsgBox "Record Found", vbInformation
 End If


rs.MoveNext

Wend
 
If n = 0 Then

        MsgBox "Record Not Found", vbCritical
End If

con.Close
End If
Adodc2.Refresh
End Sub

Private Sub Command5_Click()
Call data

rs.MoveFirst
Text1.Text = rsmax.Fields(0).Value + 1

rs.MoveNext

con.Close
End Sub

Private Sub Form_Load()
Call data
rs1.MoveFirst
While Not rs1.EOF

Combo1.AddItem (rs1.Fields(1))

rs1.MoveNext
Wend



Text1.Text = rsmax.Fields(0).Value + 1
con.Close

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= 32 And KeyAscii <= 47 Or KeyAscii >= 58 And KeyAscii <= 127 Then
MsgBox "Character And Special Characters Not Allowed", vbCritical
Text1.Text = ""
Text1.SetFocus
KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 64 Or KeyAscii >= 91 And KeyAscii <= 96 Or KeyAscii >= 123 And KeyAscii <= 127 Then
MsgBox "Number And Special Characters Not Allowed", vbCritical
Text2.Text = " "
Text2.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub Text3_Change()
If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= 32 And KeyAscii <= 47 Or KeyAscii >= 58 And KeyAscii <= 127 Then
MsgBox "Character And Special Characters Not Allowed", vbCritical
Text1.Text = ""
Text1.SetFocus
KeyAscii = 0
End If
End Sub

Private Sub Text4_Change()
If KeyAscii >= 33 And KeyAscii <= 64 Or KeyAscii >= 91 And KeyAscii <= 96 Or KeyAscii >= 123 And KeyAscii <= 127 Then
MsgBox "Number And Special Characters Not Allowed", vbCritical
Text2.Text = " "
Text2.SetFocus
KeyAscii = 0
End If
End Sub
