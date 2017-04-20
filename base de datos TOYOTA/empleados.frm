VERSION 5.00
Begin VB.Form empleados 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   Picture         =   "empleados.frx":0000
   ScaleHeight     =   8250
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11160
      ScaleHeight     =   675
      ScaleWidth      =   2595
      TabIndex        =   16
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "AÑADIR NUEVO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIGUIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ANTERIOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Puesto"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      DataField       =   "Direccion"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2400
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   3120
      Width           =   6375
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   2895
      Left            =   2760
      ScaleHeight     =   2835
      ScaleWidth      =   8235
      TabIndex        =   9
      Top             =   5160
      Width           =   8295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "EMPLEADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   11520
      TabIndex        =   15
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "APELLIDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "PUESTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "DIRECCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8370
      Left            =   -960
      Picture         =   "empleados.frx":5510F
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   15240
   End
End
Attribute VB_Name = "empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveLast
End If
End Sub
Private Sub Command2_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
End Sub

