VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPessoa 
   Caption         =   "Novo Contato"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid flexAgenda 
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdBuscaNome 
      Caption         =   "Busca"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   360
      Width           =   615
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
Private Sub cmdBuscaNome_Click()

    Dim opessoa As New Pessoa
    Dim p As ADODB.Recordset
    flexAgenda.Rows = 1
    Set p = opessoa.buscaComTelefone(txtNome.Text)
    While p.EOF = False
        flexAgenda.AddItem p!id_pessoa & vbTab & p!nome & vbTab & p!id_telefone & vbTab & p!numero
        p.MoveNext
    Wend
End Sub

Private Sub cmdSalvaPessoa_Click()
    Dim opessoa As New Pessoa
    Dim p As ADODB.Recordset
    Dim id As Integer
    If txtNome.Text = "" Then
        Exit Sub
    End If
    id = opessoa.novo(txtNome.Text)
    Set p = opessoa.buscaID(id, True)
    flexAgenda.AddItem p!id & vbTab & p!nome
End Sub


Private Sub cmdSalvaTelefone_Click()

End Sub

Private Sub Form_Load()
    flexAgenda.Rows = 1
End Sub
