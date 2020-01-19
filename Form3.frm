VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AddFoodType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Food Type"
   ClientHeight    =   6285
   ClientLeft      =   3465
   ClientTop       =   2790
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":12EF82
      Height          =   1935
      Left            =   3120
      TabIndex        =   8
      Top             =   3720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   23
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
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6720
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      RecordSource    =   "Food_type"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   6120
      Picture         =   "Form3.frx":12EF97
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   6000
      Picture         =   "Form3.frx":1310C9
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   4080
      Picture         =   "Form3.frx":1327B3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   2160
      Picture         =   "Form3.frx":13421D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add Food Type"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   7335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FI_Type"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F_ID"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "AddFoodType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rsmax As New ADODB.Recordset
Public Sub data()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Restaurant1.mdb; persist Security info=false"
con.Open
rs.Open "select * from Food_type", con, adOpenDynamic, adLockOptimistic
rsmax.Open "select max(Food_Id) from Food_type", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please Enter Some Data "
Else

Call data
rs.AddNew
rs!Food_Id = Text1.Text
rs!F_type = Text2.Text
rs.Update
MsgBox "Food Type Added Sucessfully", vbInformation


Text2.Text = ""
End If
Adodc1.Refresh
con.Close
Call data
Text1.Text = rsmax.Fields(0).Value + 1

con.Close
End Sub

Private Sub Command2_Click()
Dim n As Integer
n = 0
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Null Value is not Allowed Please Fill The Details", vbCritical

Else

Call data
rs.MoveFirst
While Not rs.EOF
If Text1.Text = rs.Fields(0).Value Then
rs.Fields(1).Value = Text2.Text


n = 1

rs.Update
MsgBox "Update Successfully", vbInformation
Text2.Text = ""

End If

rs.MoveNext
Wend

If n = 0 Then

MsgBox "Search Record First", vbCritical

End If
con.Close
End If
Adodc1.Refresh
Call data
Text1.Text = rsmax.Fields(0).Value + 1

con.Close
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Then
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

End If
Adodc1.Refresh
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
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Call data
Text1.Text = rsmax.Fields(0).Value + 1
con.Close

End Sub

