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
   Begin MSFlexGridLib.MSFlexGrid flexAgenda 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
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
    frmPessoa.Show
    'Dim rsPessoas As New ADODB.Recordset
    'opessoa.novo "ALEX"
    'Set rsPessoas = opessoa.busca("W")
    'opessoa.apagaPessoaSoft 4
    'listaPessoas
    'otelefone.novo "77777777", 6

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
