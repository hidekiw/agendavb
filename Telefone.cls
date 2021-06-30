VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Telefone"
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
Public Sub novo(numero As String, id_pessoa As Integer)
'Cria nova pessoa
    Dim consulta As String
    consulta = "INSERT INTO telefone(numero,id_pessoa,created_at) VALUES('" & numero & "'," & id_pessoa & ",'" & Now & "')"
    conn.Execute consulta
End Sub
Public Function busca(numero As String) As ADODB.Recordset
'Busca pelo nome
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM telefone WHERE numero LIKE '%" & numero & "%'"
    Set rs = conn.Execute(consulta, , adCmdText)
    Set busca = rs
End Function
Public Function apagaTelefone(id As Integer) As Boolean
'Soft DELETE
    Dim consulta As String
    consulta = "UPDATE telefone SET deleted_at=NOW WHERE deleted_at IS NULL AND id=" & id
    conn.Execute consulta
End Function

