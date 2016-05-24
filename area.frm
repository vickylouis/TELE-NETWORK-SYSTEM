VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frm_area_det 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Area Details"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form8"
   ScaleHeight     =   7680
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
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
      Left            =   8040
      Picture         =   "area.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
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
      Left            =   8040
      Picture         =   "area.frx":06AB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   4695
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   16761024
      ForeColor       =   192
      BackColorFixed  =   16711680
      ForeColorFixed  =   16777215
      Appearance      =   0
      FormatString    =   "Customer Code     |Customer Name            |Tele Number         "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox area_code 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox area_name 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox tot_con 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Area Code"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Area Name"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "TELECOMMUNICATION SYSTEM OF NETWORKS"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Connections"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "frm_area_det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim a
Private Sub area_code_Click()
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
If area_code.Text = rs(5) Then
area_name.Text = rs(4)
End If
rs.MoveNext
Wend
End If

tot_con.Text = ""
a = 0
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
If area_code.Text = rs(5) And area_name.Text = rs(4) Then
a = a + 1
tot_con.Text = a
End If
rs.MoveNext
Wend
End If

End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub cmd_view_Click()
Flex1.Clear
Dim i As Integer
i = 1
clear_flex
With Flex1
    .Row = 0
    .Col = 0
    .Text = "Customer Code"
    
    .Col = 1
    .Text = "Customer Name"
        
    .Col = 2
    .Text = "Tele Number"
    
                    
End With
rs.MoveFirst
If rs.RecordCount = 0 Then
Else
rs.MoveFirst
    While Not rs.EOF
        With Flex1
        If rs(5) = area_code.Text Then
            .AddItem ""
            .Row = i
            .Col = 0
            .Text = rs(1)
            
            .Col = 1
            .Text = rs(2)
            
            .Col = 2
            .Text = rs(6)
                     
            i = i + 1
        End If
End With
rs.MoveNext
Wend
End If

End Sub
Private Sub clear_flex()
Dim k As Long
For k = Flex1.Rows To 3 Step -1
    Flex1.RemoveItem k
Next k
End Sub

Private Sub Flex1_Click()

End Sub

Private Sub Form_Load()
a = 0
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic

Dim j As Integer
 count1 = 0

If rs.RecordCount = 0 Then
Else
    rs.MoveFirst
    While Not rs.EOF
        For j = 0 To area_code.ListCount
            If rs(5) = area_code.List(j) Then
                count1 = 1
            End If
        Next j
        If count1 = 0 Then
            area_code.AddItem (rs(5))
        End If
        If count1 = 1 Then
            count1 = 0
        End If
        rs.MoveNext
    Wend
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
cn.Close
End Sub
