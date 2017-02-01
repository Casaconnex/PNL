VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PROGRAMACION NO LINEAL"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton fibonacci 
      Caption         =   "&Calcular"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton dorado 
      Caption         =   "&Calcular"
      Height          =   375
      Left            =   2880
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   8040
      Width           =   1455
   End
   Begin VB.ListBox List7 
      Height          =   5100
      Left            =   8760
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ListBox List6 
      Height          =   5100
      Left            =   6840
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ListBox List5 
      Height          =   5100
      Left            =   5400
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Height          =   5100
      Left            =   3960
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List3 
      Height          =   5100
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   5100
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   5100
      ItemData        =   "Form1.frx":0442
      Left            =   240
      List            =   "Form1.frx":0444
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calcular"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   9360
      TabIndex        =   30
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "="
      Height          =   240
      Index           =   2
      Left            =   9000
      TabIndex        =   29
      Top             =   7440
      Width           =   135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   7680
      TabIndex        =   28
      Top             =   7200
      Width           =   75
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   6600
      TabIndex        =   27
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   240
      Index           =   2
      Left            =   8160
      TabIndex        =   26
      Top             =   7680
      Width           =   120
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   7560
      X2              =   8880
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   5520
      TabIndex        =   25
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   3120
      TabIndex        =   24
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   240
      Index           =   1
      Left            =   3840
      TabIndex        =   23
      Top             =   7200
      Width           =   75
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3720
      X2              =   5040
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   240
      Index           =   1
      Left            =   4320
      TabIndex        =   22
      Top             =   7680
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "="
      Height          =   240
      Index           =   1
      Left            =   5280
      TabIndex        =   21
      Top             =   7440
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   2040
      TabIndex        =   20
      Top             =   7440
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "="
      Height          =   240
      Index           =   0
      Left            =   1680
      TabIndex        =   19
      Top             =   7440
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   18
      Top             =   7680
      Width           =   120
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   1560
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   7200
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "f(landa)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   6960
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "f(miu)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   8880
      TabIndex        =   12
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Miu(k)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   5640
      TabIndex        =   10
      Top             =   1440
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Landa(k)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   3960
      TabIndex        =   8
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "b(k)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a(k)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Función: 1-4X+3X^2"
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'limpia las listas
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
'llama al procedimiento principal
PasoInicial
End Sub

Private Sub Command2_Click()

End Sub

Private Sub dorado_Click()
'limpia las listas
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
'llama al procedimiento principal
PASOINDORADO
End Sub

Private Sub fibonacci_Click()
'limpia las listas
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear
'llama al procedimiento principal
pasoinfibona

End Sub

Private Sub salir_Click()
Unload Me
MENU.Show
End Sub
