VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   180
      Top             =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   315
      Left            =   2340
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "Search Complete..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx _
    As Long, ByVal cy As Long, ByVal wFlags As Long)
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
Private find As New FindWord
Dim pos As Integer
Dim place As Integer

Private Sub Form_Load()
  pos = 1
  place = 0
  SetWindowPos frmFind.hWnd, HWND_TOPMOST, frmFind.Left / 15, _
        frmFind.Top / 15, frmFind.Width / 15, _
        frmFind.Height / 15, SWP_SHOWWINDOW
 chk = GetSetting("FindWord", "Check", "Firstrun")
 If chk = "" Then
    FileCopy App.Path & "\TextFunctions.dll", "c:\windows\system\TextFunctions.dll"
    Shell "c:\windows\system\REGSVR32.EXE c:\windows\system\TextFunctions.dll", vbMinimizedNoFocus
    SaveSetting "FindWord", "Check", "Firstrun", "Registered"
    find.FindWord 12487, "Find", "text1"
 End If
Exit Sub
errs:
  MsgBox "Error -" & Err.Description
End Sub

Private Sub Command1_Click()
Label1.Visible = False
If txtFind <> "" Then
  With frmNabber.Intxt
    place = InStr(pos, .Text, txtFind, vbTextCompare)
    If place > 0 Then
       .SetFocus: .SelStart = place - 1
       .SelLength = Len(txtFind)
       pos = place + Len(txtFind)
    Else
       pos = 1: place = 0: Label1.Visible = True
    End If
  End With
End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
 txtFind.SetFocus
 Timer1.Enabled = False
End Sub
