VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ProperCase Dll Tester"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReRun 
      Caption         =   "Re-Run"
      Height          =   360
      Left            =   495
      TabIndex        =   5
      Top             =   1725
      Width           =   1650
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   2370
      TabIndex        =   4
      Top             =   1725
      Width           =   1650
   End
   Begin VB.Label Label4 
      Caption         =   "Your Last name entered was:"
      Height          =   270
      Left            =   30
      TabIndex        =   3
      Top             =   840
      Width           =   2205
   End
   Begin VB.Label Label3 
      Caption         =   "Your First name entered was:"
      Height          =   270
      Left            =   30
      TabIndex        =   2
      Top             =   90
      Width           =   2265
   End
   Begin VB.Label Label2 
      Height          =   270
      Left            =   30
      TabIndex        =   1
      Top             =   1140
      Width           =   4200
   End
   Begin VB.Label Label1 
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   4170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim here makes the x instance visible to all subs on the form
Dim x As New ProperCaseName.clsName

Private Sub cmdExit_Click()
Unload Form1 'or you could use Unload Me also here
End Sub

Private Sub cmdReRun_Click()
x.EnterName

Label1.Caption = x.FirstName
Label2.Caption = x.LastName
End Sub

Private Sub Form_Load()
'Test looking at properties of x, notice that the strPrivateVar and the function
'are not visible in this project

Form1.Show

x.EnterName

Label1.Caption = x.FirstName
Label2.Caption = x.LastName

End Sub
