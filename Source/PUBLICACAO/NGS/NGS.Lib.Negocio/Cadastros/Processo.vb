Public Class Processos
    Inherits List(Of Processo)
    Public Sub New(ByVal Servidor As eServidor)
        Dim ds As DataSet
        Dim sql As String = "SELECT Processo_Id, Descricao, " & IIf(Servidor.Equals(eServidor.ServidorUOL), "IFNULL", "ISNULL") & "(Manual, '') as Manual, DataAtualizacao, UsuarioAlteracao FROM Processos"

        If Servidor.Equals(eServidor.ServidorLocal) Then
            ds = New AcessaBanco().ConsultaDataSet(sql, "Processo")
        Else
            ds = New AcessaBancoMySql().ConsultaDataSet(sql, "Processo")
        End If

        For Each row As DataRow In ds.Tables(0).Rows

            Dim obj As New Processo
            obj.Processo = row("Processo_Id")
            obj.Descricao = row("Descricao")
            obj.Manual = row("Manual")

            If Not IsDBNull(row("DataAtualizacao")) Then
                obj.DataAtualizacao = row("DataAtualizacao")
            End If
            If Not IsDBNull(row("DataAtualizacao")) Then
                obj.UsuarioAlteracao = row("UsuarioAlteracao")
            End If

            Me.Add(obj)

        Next
    End Sub
End Class

Public Class Processo
    Public Sub New()

    End Sub

    Public Sub New(ByVal Servidor As eServidor, ByVal processo As String)
        Dim ds As DataSet
        Dim sql As String = "Select Processo_Id, Descricao, " & IIf(Servidor.Equals(eServidor.ServidorUOL), "IFNULL", "isnull") & "(Manual, '') as Manual, DataAtualizacao, UsuarioAlteracao from Processos Where Processo_Id = '" & processo & "'"

        If Servidor.Equals(eServidor.ServidorLocal) Then
            ds = New AcessaBanco().ConsultaDataSet(sql, "Processo")
        Else
            ds = New AcessaBancoMySql().ConsultaDataSet(sql, "Processo")
        End If

        For Each row As DataRow In ds.Tables(0).Rows
            Me.Processo = row("Processo_Id").ToString()
            Me.Descricao = row("Descricao").ToString()
            Me.Manual = row("Manual").ToString()
            Me.DataAtualizacao = row("DataAtualizacao").ToString()
            Me.UsuarioAlteracao = row("UsuarioAlteracao").ToString()
        Next
    End Sub

    Private _IUD As String
    Private _Processo As String
    Private _Descricao As String
    Private _Manual As String
    Private _DataAtualizacao As Date
    Private _UsuarioAlteracao As String

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Processo() As String
        Get
            Return _Processo
        End Get
        Set(ByVal value As String)
            _Processo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property Manual() As String
        Get
            Return _Manual
        End Get
        Set(ByVal value As String)
            _Manual = value
        End Set
    End Property

    Public Property DataAtualizacao() As Date
        Get
            Return _DataAtualizacao
        End Get
        Set(ByVal value As Date)
            _DataAtualizacao = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Function Salvar(ByVal servidor As eServidor) As Boolean
        If servidor.Equals(eServidor.ServidorLocal) Then
            If Me.IUD = Nothing Then Return True
            Dim Banco As New AcessaBanco
            Dim Sqls As New ArrayList

            Sqls.Clear()
            SalvarSql(Sqls)

            Return Banco.GravaBanco(Sqls)
        Else
            If Not Funcoes.VerificaConexaoInternet() Then Return False
            Dim param As New Dictionary(Of String, Object)
            param.Add("Processo_Id", Me.Processo)
            param.Add("Descricao", Me.Descricao)
            param.Add("Manual", Me.Manual)
            param.Add("DataAtualizacao", Me.DataAtualizacao.ToString("yyyy-MM-dd HH:mm:ss"))
            param.Add("usuarioAlteracao", UsuarioServidor.NomeUsuario)

            Return New AcessaBancoMySql().GravaBanco("sp_insertorupdate_processo", True, param)
        End If
        Return False
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO Processos(Processo_Id, Descricao, Manual, DataAtualizacao, UsuarioAlteracao) " & vbCrLf & _
                      " VALUES ('" & Me.Processo & "', '" & Me.Descricao & "', '" & Me.Manual & "', '" & Me.DataAtualizacao.ToString("yyyy-MM-dd HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Processos SET" & vbCrLf & _
                      "        Descricao        = '" & Me.Descricao & "'," & vbCrLf & _
                      "        Manual           = '" & Me.Manual & "'," & vbCrLf & _
                      "        DataAtualizacao  = '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                      "        UsuarioAlteracao = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                      "  WHERE Processo_Id  = '" & Me.Processo & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE From Processos" & vbCrLf & _
                      "  WHERE Processo_Id = '" & Me.Processo & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

End Class
