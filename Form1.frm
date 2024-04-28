VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   11805
   ClientTop       =   2490
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      FillColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6555
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   120
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form2.Show
End Sub

