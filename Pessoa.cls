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
Const db_file = "agenda.accdb"
Dim conn As ADODB.Connection
Private Sub Class_Initialize()

    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & App.Path & "\" & db_file
    conn.Open
End Sub
Public Sub novo(nome As String)
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "INSERT INTO pessoa(nome,created_at) VALUES('" & UCase(nome) & "','" & Now & "')"
    Set rs = conn.Execute(consulta, , adCmdText)

End Sub
Public Function busca(nome As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM pessoa WHERE nome LIKE '%" & nome & "%' AND deleted_at IS NULL"
    Set rs = conn.Execute(consulta, , adCmdText)
   Set busca = rs
End Function
Public Function apagaPessoa(id As Integer) As Boolean
    Dim consulta As String
    consulta = "DELETE FROM pessoa WHERE id=" & id
    conn.Execute consulta
End Function
Public Function apagaPessoaSoft(id As Integer) As Boolean
    Dim consulta As String
    consulta = "UPDATE pessoa SET deleted_at=now WHERE id = " & id
    conn.Execute consulta
End Function
Public Function lixeira() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM pessoa WHERE deleted_at IS NULL"
    Set rs = conn.Execute(consulta, , adCmdText)
   Set lixeira = rs
End Function
