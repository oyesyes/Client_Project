VERSION 5.00
Begin VB.Form RO��½�� 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7245
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   1860
      TabIndex        =   1
      Top             =   1560
      Width           =   1515
   End
   Begin VB.CommandButton StartGame_buttom 
      Caption         =   "��ʼ��Ϸ"
      Height          =   615
      Left            =   1860
      TabIndex        =   0
      Top             =   480
      Width           =   1515
   End
End
Attribute VB_Name = "RO��½��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute Lib "App.Path & shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1


Private Sub StartGame_buttom_Click()

    If Dir(App.Path & "\oyesyes.exe") = "" Then
        MsgBox "û�д��ļ�,��ѱ������Ƶ���Ϸ��Ŀ¼�£�����"
    Else
    End If
    ShellExecute Me.hwnd, "Open", App.Path & "\oyesyes.exe", vbNullString, vbNullString, SW_SHOWNORMAL
    
    End
    
End Sub

Private Sub Command2_Click()
    Shell App.Path & "\oyesyes.exe"
    End
End Sub

Private Sub Form_Click()
    Dim x, y
    x = 20
    y = 30
    Print App.Path

End Sub

