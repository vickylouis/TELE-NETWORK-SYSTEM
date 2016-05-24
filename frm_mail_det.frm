VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mail"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox from 
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
      Left            =   4560
      TabIndex        =   7
      Text            =   "                          "
      Top             =   240
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   -1560
      Picture         =   "frm_mail_det.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   4155
      TabIndex        =   6
      Top             =   1560
      Width           =   4215
      Begin VB.PictureBox Picture4 
         Height          =   615
         Left            =   2640
         Picture         =   "frm_mail_det.frx":72D7
         ScaleHeight     =   555
         ScaleWidth      =   1515
         TabIndex        =   12
         Top             =   2880
         Width           =   1575
      End
   End
   Begin VB.TextBox b1 
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
      Height          =   2175
      Left            =   2880
      TabIndex        =   5
      Text            =   "Message"
      Top             =   2640
      Width           =   5055
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   1080
      Picture         =   "frm_mail_det.frx":7BB9
      ScaleHeight     =   555
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox s1 
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   -120
      Picture         =   "frm_mail_det.frx":849B
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   -240
      Width           =   2775
   End
   Begin VB.ComboBox a1 
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
      Left            =   4560
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox t1 
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
      Left            =   4560
      TabIndex        =   0
      Text            =   "                          "
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Attachment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim oOApp As Outlook.Application
Dim oOMail As Outlook.MailItem

Private Sub Form_Load()
a1.AddItem ("Yes")
a1.AddItem ("No")


cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Tele.mdb;Persist Security Info=False"
cn.CursorLocation = adUseClient

rs.Open "select * from cus_info", cn, adOpenDynamic, adLockPessimistic
From.Text = ""
t1.Text = ""

If rs.RecordCount = 0 Then
Else
rs.MoveFirst
While Not rs.EOF
t1.AddItem (rs(8))
rs.MoveNext
Wend
End If

End Sub

Private Sub Picture4_Click()
Set oOApp = CreateObject("Outlook.Application")
Set oOMail = oOApp.CreateItem(olMailItem)

With oOMail
.To = t1.Text
.Subject = s1.Text
.Body = b1.Text
'.Display
If a1.Text = "Yes" Then
.Attachments.Add "c:\Attachments.Zip", olByValue, 1
End If
.Send
Set oOMail = Nothing
Set oOApp = Nothing

End With
End Sub
