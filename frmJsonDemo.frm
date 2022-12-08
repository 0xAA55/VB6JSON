VERSION 5.00
Begin VB.Form frmJsonDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB JSON"
   ClientHeight    =   6495
   ClientLeft      =   450
   ClientTop       =   1875
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11415
   Begin VB.CommandButton cmdFormat 
      Caption         =   "格式化JSON =>"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtFormatJSON 
      Height          =   5175
      Left            =   6600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox txtSourceJSON 
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblInputJson 
      AutoSize        =   -1  'True
      Caption         =   "测试输入JSON"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmJsonDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFormat_Click()
On Error GoTo ErrHander
txtFormatJSON.Text = JSONToString(ParseJSONString(txtSourceJSON.Text), 4)
Exit Sub
ErrHander:
MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Private Sub Form_Load()
txtSourceJSON.Text = "[{""asd"": 123, ""fsd"": 456}, {""sdf"": {""asdf"": 1234}, ""ghi"": ""ih\u0038g""}]"
End Sub
