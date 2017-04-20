VERSION 5.00
Begin VB.Form primera 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox DataGrid1 
      Height          =   3135
      Left            =   2040
      ScaleHeight     =   3075
      ScaleWidth      =   9795
      TabIndex        =   15
      Top             =   4920
      Width           =   9855
   End
   Begin VB.TextBox Text5 
      DataField       =   "Potencia"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   3240
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      DataField       =   "Serie"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2520
      Width           =   6375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Combustible"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Modelo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   360
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
      Left            =   1320
      TabIndex        =   3
      Top             =   4080
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
      Left            =   4440
      TabIndex        =   2
      Top             =   4080
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
      Left            =   7320
      TabIndex        =   1
      Top             =   4080
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
      Left            =   10320
      TabIndex        =   0
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   11760
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "VEHICULOS PRIMERA GENERACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   11280
      TabIndex        =   14
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   13
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COMBUSTIBLE"
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
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   8370
      Left            =   0
      Picture         =   "primera.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15240
   End
End
Attribute VB_Name = "primera"
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


