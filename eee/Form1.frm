VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Object
Dim epmask As Long
Dim context As String
Dim sUser As String

context = "8A187894-3B6B-4473-81D7-4633309F58F4"

epmask = 0

Set a = CreateObject("E1LRUC.e1LRConfig")

a.GetModulePermissions epmask, context, InputBox("USer", , "nagendra.prasad")

MsgBox epmask
End Sub
