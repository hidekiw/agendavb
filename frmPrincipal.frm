VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "AgendaVB"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Agenda"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.CommandButton Command3 
         Caption         =   "TELEFONE"
         Height          =   255
         Left            =   5880
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscaNome 
         Caption         =   "BUSCA"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSalvaPessoa 
         Caption         =   "ADD"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtTelefone 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNomePessoa 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "a"
         Top             =   360
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flexAgenda 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   4
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
Private Sub cmdBuscaNome_Click()
    Dim opessoa As New Pessoa
    Dim p As ADODB.Recordset
    flexAgenda.Rows = 1
    Set p = opessoa.buscaComTelefone(txtNomePessoa.Text)
    While p.EOF = False
        flexAgenda.AddItem p!id_pessoa & vbTab & p!nome & vbTab & p!id_telefone & vbTab & p!numero
        p.MoveNext
    Wend
End Sub

Private Sub cmdSalvaPessoa_Click()
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



Private Sub Form_Load()
    flexAgenda.Rows = 1
flexAgenda.SelectionMode = flexSelectionByRow
'flexAgenda.FillStyle = flexFillSingle
'flexAgenda.AllowBigSelection = False
flexAgenda.MergeCells = flexMergeRestrictRows
End Sub
Private Sub listaPessoas()
'    Dim rspessoas As New ADODB.Recordset
'    Dim opessoas As New Pessoa
'    configFlexPessoas
'    Set rspessoas = opessoas.busca("")
'    While rspessoas.EOF = False
'        flexPessoas.AddItem rspessoas!id & vbTab & rspessoas!nome
'        rspessoas.MoveNext
'    Wend
End Sub
'Private Sub configFlexPessoas()
'flexPessoas.Rows = 1
'flexPessoas.TextMatrix(0, 0) = "ID"
'flexPessoas.TextMatrix(0, 1) = "Nome"
'flexPessoas.ColWidth(0) = 500
'flexPessoas.ColWidth(1) = 6000
'flexPessoas.SelectionMode = flexSelectionByRow
'flexPessoas.AllowBigSelection = False
'End Sub
