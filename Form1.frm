VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   3900
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   840
   End
   Begin VB.Label ��F2������ɫ 
      BackColor       =   &H8000000B&
      Caption         =   "��F2������ɫ"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "Label2"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Label1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
x As Long
y As Long
End Type
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
'�жϺ�������ʱָ���������״̬
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Function color(R As Integer, G As Integer, B As Integer) As String
    Dim se(8) As String
    se(1) = "��ɫ": se(2) = "��ɫ": se(3) = "��ɫ": se(4) = "��ɫ": se(5) = "��ɫ": se(6) = "���ɫ": se(7) = "��ɫ": se(8) = "��ɫ"
    If R <= 20 And G <= 20 And B <= 20 Then
        color = se(1)
    ElseIf R >= 235 And G >= 235 And B >= 235 Then color = se(2)
    ElseIf R - G >= 50 And R - B >= 50 And Abs(B - G) <= 50 Then color = se(3)
    ElseIf B - G >= 50 And B - R >= 50 And Abs(R - G) <= 50 Then color = se(4)
    ElseIf G - R >= 50 And G - B >= 50 And Abs(R - B) <= 50 Then color = se(5)
    ElseIf R - G >= 50 And B - G >= 50 And Abs(B - R) <= 50 Then color = se(6)
    ElseIf R - B >= 50 And G - B >= 50 And Abs(R - G) <= 50 Then color = se(7)
    ElseIf G - R >= 50 And B - R >= 50 And Abs(G - B) <= 50 Then color = se(8)
    Else: color = "��ɫ"
    End If
End Function
Private Sub Form_Load()
Timer1.Interval = 1
End Sub

Private Sub Timer1_Timer()
Dim hdc As Long
Dim A As POINTAPI
Dim quyanse As Long
Call GetCursorPos(A) 'ȡ�����λ��
Text1.Text = "X��" & A.x & "   Y��" & A.y
hdc = GetDC(0) 'ȡ��������Ļ��hDC
Form1.BackColor = GetPixel(hdc, A.x, A.y) 'ȡ��ɫ
ReleaseDC 0, hdc '�ͷ�hDC
If MyHotKey(vbKeyF2) Then '�������F2,�ͻ�ȡ��ɫֵ��������
quyanse = GetPixel(Me.hdc, 2, 2) 'ȡ��ɫֵ
Text2.Text = Str(quyanse)
Label1.Caption = "R:" & CLng("&H" & Right(Hex(GetPixel(Me.hdc, 2, 2)), 2)) & " G:" & CLng("&H" & Right(Left(Hex(GetPixel(Me.hdc, 2, 2)), 4), 2)) & " B:" & CLng("&H" & Left(Hex(GetPixel(Me.hdc, 2, 2)), 2))
Label2.Caption = color(CLng("&H" & Right(Hex(GetPixel(Me.hdc, 2, 2)), 2)), CLng("&H" & Right(Left(Hex(GetPixel(Me.hdc, 2, 2)), 4), 2)), CLng("&H" & Left(Hex(GetPixel(Me.hdc, 2, 2)), 2)))
MsgBox Text1.Text & "��ɫֵ:" & quyanse
End If
End Sub
Private Function MyHotKey(vKeyCode) As Boolean
MyHotKey = (GetAsyncKeyState(vKeyCode) < 0)
End Function
