VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "AgendaVB"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   27
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      HighLight       =   0
      ScrollBars      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flexTelefones 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid flexPessoas 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2355
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim opessoa As New Pessoa
Dim otelefone As New Telefone

'Dim rsPessoas As New ADODB.Recordset
'opessoa.novo "ALEX"
'Set rsPessoas = opessoa.busca("W")
'opessoa.apagaPessoaSoft 4
listaPessoas
otelefone.novo "77777777", 6

End Sub
Private Sub listaPessoas()
    Dim rspessoas As New ADODB.Recordset
    Dim opessoas As New Pessoa
    configFlexPessoas
    Set rspessoas = opessoas.busca("")
    While rspessoas.EOF = False
        flexPessoas.AddItem rspessoas!id & vbTab & rspessoas!nome
        rspessoas.MoveNext
    Wend
End Sub
Private Sub configFlexPessoas()
flexPessoas.Rows = 1
flexPessoas.TextMatrix(0, 0) = "ID"
flexPessoas.TextMatrix(0, 1) = "Nome"
flexPessoas.ColWidth(0) = 500
flexPessoas.ColWidth(1) = 6000
flexPessoas.SelectionMode = flexSelectionByRow
flexPessoas.AllowBigSelection = False
End Sub
