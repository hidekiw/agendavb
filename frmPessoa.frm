VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPessoa 
   Caption         =   "Novo Contato"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscaNome 
      Caption         =   "Busca"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid flexTelefones 
      Height          =   1575
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2778
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalvaTelefone 
      Caption         =   "+"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtTelefone 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalvaPessoa 
      Caption         =   "+"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid flexPessoas 
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Telefone"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalvaPessoa_Click()
    Dim opessoa As New Pessoa
    Dim p As ADODB.Recordset
    Dim id As Integer
    If txtNome.Text = "" Then
        Exit Sub
    End If
    id = opessoa.novo(txtNome.Text)
    Set p = opessoa.buscaID(id, True)
    flexPessoas.AddItem p!id & vbTab & p!nome
End Sub


