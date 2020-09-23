VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show FindWord Form"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   900
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmFind.Show
End Sub
