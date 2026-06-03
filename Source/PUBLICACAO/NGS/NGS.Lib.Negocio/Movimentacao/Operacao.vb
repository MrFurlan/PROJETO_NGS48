Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListOperacao
    Inherits List(Of Operacao)

#Region "Construtor"
    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        If CarregarDados Then
            Selecionar()
        End If
    End Sub
#End Region

#Region "Methods"
    Public Function Selecionar() As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Operacao_Id, Descricao, Producao, isnull(UFDepositoDestino,0) as UFDepositoDestino, isnull(classe,'') as Classe, isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioAlteracao,'') as UsuarioAlteracao " &
                                   "  FROM Operacoes " &
                                   " ORDER BY Operacao_Id"

            Dim dsOperacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "Produtos")

            For Each drOperacao As DataRow In dsOperacoes.Tables(0).Rows
                Dim objOperacao As New Operacao()

                objOperacao.Codigo = Convert.ToInt32(drOperacao("Operacao_Id"))
                objOperacao.Descricao = drOperacao("Descricao")
                objOperacao.Producao = drOperacao("Producao") = "S"
                objOperacao.UFDepositoDestino = drOperacao("UFDepositoDestino")
                objOperacao.CodigoClasse = drOperacao("Classe")
                objOperacao.UsuarioInclusao = drOperacao("UsuarioInclusao")
                objOperacao.UsuarioAlteracao = drOperacao("UsuarioAlteracao")

                Me.Add(objOperacao)
            Next

            Return True
        Catch ex As Exception
            'Me.Erro = ex
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function
#End Region

End Class

<Serializable()>
Public Class Operacao
    Public Erro As Exception

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Me.Codigo = Codigo
        Selecionar(Codigo)
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Producao As Boolean
    Private _UFDepositoDestino As Boolean
    Private _CodigoClasse As String
    Private _Classe As ClasseDeOperacao
    Private _SubOperacaoes As ListSubOperacao
    Private _UsuarioInclusao As String
    Private _UsuarioAlteracao As String
    Private _Selecionado As Boolean = False
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
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

    Public Property Producao() As Boolean
        Get
            Return _Producao
        End Get
        Set(ByVal value As Boolean)
            _Producao = value
        End Set
    End Property

    Public Property UFDepositoDestino As Boolean
        Get
            Return _UFDepositoDestino
        End Get
        Set(value As Boolean)
            _UFDepositoDestino = value
        End Set
    End Property

    Public Property CodigoClasse() As String
        Get
            Return _CodigoClasse
        End Get
        Set(ByVal value As String)
            _CodigoClasse = value
            _Classe = Nothing
        End Set
    End Property

    Public ReadOnly Property Classe() As ClasseDeOperacao
        Get
            If _Classe Is Nothing And Not String.IsNullOrWhiteSpace(_CodigoClasse) Then _Classe = New ClasseDeOperacao(_CodigoClasse)
            Return _Classe
        End Get
    End Property

    Public Property SubOperacoes As ListSubOperacao
        Get
            If _SubOperacaoes Is Nothing And _Codigo > 0 Then
                _SubOperacaoes = New ListSubOperacao(_Codigo)
            End If
            Return _SubOperacaoes
        End Get
        Set(value As ListSubOperacao)
            _SubOperacaoes = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
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

    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Selecionar(ByVal Operacao As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Operacao_Id, Descricao, Producao, isnull(UFDepositoDestino,0) as UFDepositoDestino, isnull(classe,'') as Classe, UsuarioInclusao, UsuarioAlteracao " &
                                   "  FROM Operacoes " &
                                   " WHERE Operacao_Id = " & Operacao.ToString()

            Dim dsOperacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "Produtos")

            If dsOperacoes.Tables(0).Rows.Count > 0 Then
                Dim drOperacao As DataRow = dsOperacoes.Tables(0).Rows(0)

                Me.Codigo = Convert.ToInt32(drOperacao("Operacao_Id"))
                Me.Descricao = drOperacao("Descricao").ToString()
                Me.Producao = drOperacao("Producao").ToString() = "S"
                Me.UFDepositoDestino = drOperacao("UFDepositoDestino")
                Me.CodigoClasse = drOperacao("Classe")
                Me.UsuarioInclusao = drOperacao("UsuarioInclusao")
                Me.UsuarioAlteracao = drOperacao("UsuarioAlteracao")
            End If

            Return True
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        salvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub salvarSql(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        Select Case IUD
            Case "I"
                strSql = "Insert into Operacoes (Operacao_Id,Descricao,Producao, Classe, UFDepositoDestino, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf &
                      "values (" & Me.Codigo & ",'" & UCase(Me.Descricao) & "'," & vbCrLf &
                      "'" & IIf(Me.Producao, "S", "N") & "','" & Me.CodigoClasse & "'," & CByte(Me.UFDepositoDestino) & ",'" & UsuarioServidor.NomeUsuario & "', '" & DateTime.Now.ToString("yyyy-MM-dd") & "')"
                sqls.Add(strSql)
            Case "U"
                strSql = "UPDATE Operacoes set " & vbCrLf &
                      "    Descricao            ='" & UCase(Me.Descricao) & "'" & vbCrLf &
                      "   ,Producao             ='" & IIf(Me.Producao, "S", "N") & "'" & vbCrLf &
                      "   ,UFDepositoDestino    =" & Str(Me.UFDepositoDestino) & vbCrLf &
                      "   ,Classe               ='" & Me.CodigoClasse & "'" & vbCrLf &
                      "   ,UsuarioAlteracao     ='" & UsuarioServidor.NomeUsuario & "'" & vbCrLf &
                      "   ,UsuarioAlteracaoData ='" & DateTime.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      " WHERE Operacao_id = " & Me.Codigo & vbCrLf
                sqls.Add(strSql)
            Case "D"
                strSql = "Delete from Operacoes Where Operacao_Id = " & Me.Codigo
                sqls.Add(strSql)
        End Select
    End Sub
#End Region

End Class