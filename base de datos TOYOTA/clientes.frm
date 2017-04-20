VERSION 5.00
Begin VB.Form clientes 
   Caption         =   "Form2"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   LinkTopic       =   "Form2"
   ScaleHeight     =   8700
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      DataField       =   "Telefono"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   3120
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      DataField       =   "Direccion"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Apellido"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   6375
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
      Left            =   720
      TabIndex        =   3
      Top             =   3960
      Width           =   2295
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
      Left            =   3840
      TabIndex        =   2
      Top             =   3960
      Width           =   2055
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
      Left            =   6720
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
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
      Left            =   9720
      TabIndex        =   0
      Top             =   3960
      Width           =   2175
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11160
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   2895
      Left            =   2400
      ScaleHeight     =   2835
      ScaleWidth      =   7755
      TabIndex        =   8
      Top             =   5160
      Width           =   7815
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   15
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "TELEFONO"
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
      Left            =   840
      TabIndex        =   13
      Top             =   3120
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
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label3 
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
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8730
      Left            =   0
      Picture         =   "clientes.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15240
   End
End
Attribute VB_Name = "clientes"
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
