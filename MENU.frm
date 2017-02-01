VERSION 5.00
Begin VB.Form MENU 
   Caption         =   "PROGRAMACION NO LINEAL"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   14925
   LinkTopic       =   "Form2"
   ScaleHeight     =   10500
   ScaleWidth      =   14925
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "&Busqueda Fibonacci"
      Height          =   435
      Left            =   3960
      TabIndex        =   3
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Metodo de la Seccion Dorado"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Metodos"
      Height          =   5295
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Metodo &Dicotomico"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Command1.Visible = True

Form1.Show
End Sub

Private Sub Command2_Click()
Unload Me
Form1.dorado.Visible = True

Form1.Show
End Sub

Private Sub Command3_Click()
Unload Me
Form1.fibonacci.Visible = True
Form1.Show

End Sub
