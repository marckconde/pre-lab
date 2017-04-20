VERSION 5.00
Begin VB.Form segunda 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   2985
   ClientTop       =   1965
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   Picture         =   "segunda.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   15150
   Begin VB.PictureBox DataGrid1 
      Height          =   2775
      Left            =   2760
      ScaleHeight     =   2715
      ScaleWidth      =   7995
      TabIndex        =   14
      Top             =   4920
      Width           =   8055
   End
   Begin VB.TextBox Text1 
      DataField       =   "Clase"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   240
      Width           =   6375
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
      Left            =   10320
      TabIndex        =   7
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
      Left            =   7320
      TabIndex        =   6
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
      Left            =   4440
      TabIndex        =   5
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
      Left            =   1320
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      DataField       =   "Color"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Modelo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      DataField       =   "Serie"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2400
      Width           =   6375
   End
   Begin VB.TextBox Text5 
      DataField       =   "Potencia"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   3120
      Width           =   6375
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11760
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "VEHICULOS DE SEGUNDA GENERACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   555
      Left            =   10320
      TabIndex        =   15
      Top             =   1920
      Width           =   9915
   End
   Begin VB.Label Label1 
      Caption         =   "CLASE"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "COLOR"
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
      Left            =   1440
      TabIndex        =   11
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "MODELO"
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
      Left            =   1440
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "SERIE"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "POTENCIA"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8370
      Left            =   0
      Picture         =   "segunda.frx":C4797
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15240
   End
End
Attribute VB_Name = "segunda"
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



