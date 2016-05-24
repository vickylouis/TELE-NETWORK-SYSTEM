VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Particular Customer Report"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_view 
      Caption         =   "VIEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Picture         =   "Form1.frx":06AB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cus_code 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "PARTICULAR CUSTOMER REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_view_Click()
DataEnvironment1.Connection1.Close
DataEnvironment1.Connection1.Open
DataEnvironment1.Command7 cus_code.Text
DataReport7.Show
'DataEnvironment1.Connection1.Close
End Sub

Private Sub Form_Load()
'DataEnvironment1.Connection1.Close
DataEnvironment1.Connection1.Open
cn.ConnectionString = "Provider=microsoft.jet.oledb.4.0; data source = " & App.Path & "\tele.mdb"
cn.Open

rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
cus_code.Clear
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
cus_code.AddItem (rs(1))
rs.MoveNext
Wend
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'DataEnvironment1.Connection1.Close
rs.Close
cn.Close
'DataEnvironment1.Connection1.Open
End Sub
