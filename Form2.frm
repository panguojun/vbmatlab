VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6765
   ClientLeft      =   2055
   ClientTop       =   2490
   ClientWidth     =   9645
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Top             =   6000
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim objShell As Object
    Dim strPythonScript As String
    Dim strOutput As String
    Dim i As Integer
    Dim x(100) As Single
    Dim y(100) As Single

    ' ��Text1�ı����л�ȡPython�ű�����
    strPythonScript = Text1.Text

    ' ��VB6������ִ��Python�ű�����ȡ��������
    Set objShell = CreateObject("WScript.Shell")
    strOutput = objShell.Exec("python -c """ & strPythonScript & """").StdOut.ReadAll

    ' ���������ݻ�����VB6������
    Dim curveData() As String
    curveData = Split(strOutput, vbNewLine)
   
    For i = 1 To 99
        Dim values() As String
        values = Split(curveData(i), " ")
        x(i) = CSng(values(0))
        y(i) = CSng(values(1))
    Next i
    
    ScaleWidth_ = ScaleWidth / 10
    ScaleHeight_ = ScaleHeight / 10
    X0 = 1000
    Y0 = 2500
    
    ' ��VB6�����л�������
    For i = 1 To UBound(x)
        Form1.Picture1.Line (X0 + x(i - 1) * ScaleWidth_, Y0 + y(i - 1) * ScaleHeight_)-(X0 + x(i) * ScaleWidth_, Y0 + y(i) * ScaleHeight_), vbWhite
    Next i
End Sub

