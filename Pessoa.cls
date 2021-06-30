VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pessoa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As ADODB.Connection
Private Sub Class_Initialize()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & App.Path & "\" & db_file
    conn.Open
End Sub
Public Function novo(nome As String) As Integer
    Dim rs As ADODB.Recordset
    Dim consulta As String
    novo = -1
    consulta = "INSERT INTO pessoa(nome,created_at) VALUES('" & UCase(nome) & "','" & Now & "')"
    Set rs = conn.Execute(consulta, , adCmdText)
    consulta = "SELECT id FROM pessoa WHERE nome='" & nome & "'"
    Set rs = conn.Execute(consulta, , adCmdText)
    novo = rs!id

End Function
Public Function busca(nome As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM pessoa WHERE nome LIKE '%" & nome & "%' AND deleted_at IS NULL"
    Set rs = conn.Execute(consulta, , adCmdText)
    Set busca = rs
End Function
Public Function buscaComTelefone(nome As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select p.id as id_pessoa, t.id as id_telefone,* from pessoa p left join telefone t on (p.id=t.id_pessoa) WHERE p.nome LIKE '%" & nome & "%' AND p.deleted_at IS NULL AND t.deleted_at IS NULL"
    Set rs = conn.Execute(consulta, , adCmdText)
    Set buscaComTelefone = rs
End Function
Public Function buscaID(id As Integer, lixeira As Boolean) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM pessoa WHERE id=" & id
    If lixeira = False Then
        consulta = consulta & " AND deleted_at IS NULL"
    End If
    Set rs = conn.Execute(consulta, , adCmdText)
    Set buscaID = rs
End Function
Public Function apagaPessoaSoft(id As Integer) As Boolean
    Dim consulta As String
    consulta = "UPDATE pessoa SET deleted_at=now WHERE deleted_at IS NULL AND id = " & id
    conn.Execute consulta
End Function
Public Function lixeira() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select p.id as id_pessoa, t.id as id_telefone,* from pessoa p left join telefone t on (p.id=t.id_pessoa) WHERE p.nome LIKE '%" & nome & "%' "
    Set rs = conn.Execute(consulta, , adCmdText)
    Set lixeira = rs
End Function
Public Function existeNome(nome As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM pessoa WHERE nome = '" & nome & "' AND deleted_at IS NULL"
    Set rs = conn.Execute(consulta, , adCmdText)
    Set existeNome = rs
End Function

