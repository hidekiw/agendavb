VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "AgendaVB"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9015
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Agenda"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton cmdSalvaTelefone 
         Caption         =   "Salva"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscaNome 
         Caption         =   "Busca Nome"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdNovaPessoa 
         Caption         =   "Nova Pessoa"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   7320
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNomePessoa 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
      Begin MSFlexGridLib.MSFlexGrid flexAgenda 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   4
      End
      Begin VB.Label Label2 
         Caption         =   "Telefone:"
         Height          =   255
         Left            =   6600
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Menu mnuPessoas 
      Caption         =   "Pessoas"
      Begin VB.Menu mnuApagarPessoa 
         Caption         =   "Apagar"
      End
   End
   Begin VB.Menu mnuTelefones 
      Caption         =   "Telefones"
      Begin VB.Menu mnuApagarTelefone 
         Caption         =   "Apagar"
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "Sobre"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Sub cmdBuscaNome_Click()
    busca txtNomePessoa.Text
End Sub
Private Sub busca(nome As String)
    Dim opessoa As New Pessoa
    Dim p As ADODB.Recordset
    flexAgenda.Rows = 1
    Sleep 500
    Set p = opessoa.buscaComTelefone(nome)
    While p.EOF = False
        flexAgenda.AddItem p!id_pessoa & vbTab & p!nome & vbTab & p!id_telefone & vbTab & p!numero
        p.MoveNext
    Wend
End Sub
Private Sub cmdNovaPessoa_Click()
    Dim opessoa As New Pessoa
    Dim p As ADODB.Recordset
    Dim id As Integer
    If txtNomePessoa.Text = "" Then
        Exit Sub
    End If
    id = opessoa.novo(txtNomePessoa.Text)
    Set p = opessoa.buscaID(id, True)
    flexAgenda.AddItem p!id & vbTab & p!nome
End Sub

Private Sub cmdSalvaTelefone_Click()
    Dim t As New Telefone
    If flexAgenda.Row < 1 Then
        Exit Sub
    End If
    If txtTelefone.Text = "" Then
        MsgBox "Digite o número do telefone!", vbCritical, "Erro"
        Exit Sub
    End If
    t.novo txtTelefone, flexAgenda.TextMatrix(flexAgenda.Row, 0)
    busca ""
End Sub

Private Sub flexAgenda_Click()
    txtNomePessoa.Text = flexAgenda.TextMatrix(flexAgenda.Row, 1)
End Sub

Private Sub Form_Load()
    configGrid
    busca ""
End Sub
Private Sub configGrid()
    flexAgenda.Rows = 1
    flexAgenda.SelectionMode = flexSelectionByRow
    flexAgenda.MergeCells = flexMergeRestrictRows
    flexAgenda.ColWidth(0) = 500    'id pessoa
    flexAgenda.ColWidth(1) = 5800    'nome
    flexAgenda.ColWidth(2) = 500    'id telefone
    flexAgenda.ColWidth(3) = 1500    ' telefone
    flexAgenda.TextMatrix(0, 0) = "ID"
    flexAgenda.TextMatrix(0, 1) = "Nome"
    flexAgenda.TextMatrix(0, 2) = "ID"
    flexAgenda.TextMatrix(0, 3) = "Telefone"
End Sub

Private Sub mnuApagarPessoa_Click()
    Dim p As New Pessoa
    If flexAgenda.Row < 1 Then
        Exit Sub
    End If
    If flexAgenda.TextMatrix(flexAgenda.Row, 0) = "" Then
        Exit Sub
    End If
    p.apagaPessoaSoft flexAgenda.TextMatrix(flexAgenda.Row, 0)
End Sub

Private Sub mnuApagarTelefone_Click()
    Dim t As New Telefone
    If flexAgenda.Row < 1 Then
        Exit Sub
    End If
    If flexAgenda.TextMatrix(flexAgenda.Row, 2) = "" Then
        Exit Sub
    End If

    t.apagaTelefone flexAgenda.TextMatrix(flexAgenda.Row, 2)
    busca ""
End Sub
