Imports System.Web
Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

Public Class Numeradores
    Inherits List(Of Numerador)

#Region "Contrutores"
    'Lista Vazia
    Public Sub New()

    End Sub
    'Lista Carregada por Empresa
    Public Sub New(ByVal Empresa As String, ByVal Endereco As Integer)
        Dim Banco As New AcessaBanco
        Dim strSQL As String
        strSQL = " SELECT c.Reduzido,n.Empresa_Id, n.EndEmpresa_Id,C.Fantasia + ' - ' + C.Cidade + '-' + C.Estado_Id AS Nome," & vbCrLf & _
                 "       n.Numerador_Id, n.Descricao, n.Sequencia " & vbCrLf & _
                 "  FROM Numerador n" & vbCrLf & _
                 " INNER JOIN Clientes c" & vbCrLf & _
                 "    ON n.Pais_Id       = c.Pais_Id " & vbCrLf & _
                 "   AND n.Empresa_Id    = c.Cliente_Id " & vbCrLf & _
                 "   AND n.EndEmpresa_Id = c.Endereco_Id " & vbCrLf & _
                 " WHERE n.Empresa_Id    ='" & Empresa & "'" & vbCrLf & _
                 "   AND n.EndEmpresa_Id = " & Endereco & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Numerador")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim Num As New Numerador
            Num.Reduzido = row("Reduzido")
            Num.Empresa = row("Empresa_Id")
            Num.EndEmpresa = row("EndEmpresa_Id")
            Num.Nome = row("Nome")
            Num.Numerador = row("Numerador_Id")
            Num.Descricao = row("Descricao")
            Num.Sequencia = row("Sequencia")
            Me.Add(Num)
        Next
    End Sub
#End Region

#Region "Métodos"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As ArrayList = New ArrayList()

        Sqls.Clear()
        Sqls.AddRange(Me.SalvarSQL())

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Function SalvarSQL() As ArrayList
        Dim Sqls As ArrayList = New ArrayList()
        Sqls.Clear()
        For Each Num As Numerador In Me
            Sqls.Add(Num.SalvarSql())
        Next
        Return Sqls
    End Function
#End Region

End Class

Public Class Numerador

#Region "Fields - Enum"

    Public Enum enumNumerador
        TITULO = 1
        CERTIDAONEGATIVA = 2
        ANALISEDECREDITO = 3
        PROCURACAO = 9
        PEDIDOS = 10
        RECEITAS = 13
        ADIANTAMENTOS = 15
        NOTAENTRADA = 20
        NOTAENTRADARETROATIVA = 21
        NOTASAIDA = 30
        NOTASAIDARETROATIVA = 31
        CONHECIMENTO = 35
        CONTRATODEFRETE = 40
        FATURADEFRETE = 45
        AUTORIZACAODERETIRADA = 50
        RAZAO = 60
        LAUDO = 101
        ROMANEIO = 110
        ORDEMDEPRODUCAO = 75
    End Enum

    Private _sEmpresa, _sDescricao, _sReduzido, _sNome, _sSerie As String
    Private _iEndEmpresa, _iNumerador, _iSequencia As Integer
    Private _IUD As String

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal Numerador As Integer)
        Dim Banco As New AcessaBanco
        Dim Sql As String

        Sql = " SELECT c.Reduzido,n.Empresa_Id, n.EndEmpresa_Id, C.Fantasia + ' - ' + C.Cidade + '-' + C.Estado AS Nome," & vbCrLf & _
              "        n.Numerador_Id, n.Descricao, n.Sequencia, isnull(n.Serie, '') AS Serie " & vbCrLf & _
              "   FROM Numerador n" & vbCrLf & _
              "  INNER JOIN Clientes c" & vbCrLf & _
              "     ON n.Empresa_Id    = c.Cliente_Id " & vbCrLf & _
              "    AND n.EndEmpresa_Id = c.Endereco_Id " & vbCrLf & _
              "  WHERE n.Empresa_Id    ='" & Empresa & "'" & vbCrLf & _
              "    AND n.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
              "    AND n.Numerador_Id  = " & Numerador & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Nume")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim Row As DataRow = ds.Tables(0).Rows(0)
            Me.Reduzido = Row("Reduzido")
            Me.Empresa = Row("Empresa_Id")
            Me.EndEmpresa = Row("EndEmpresa_Id")
            Me.Numerador = Row("Numerador_Id")
            Me.Descricao = Row("Descricao")
            Me.Nome = Row("Nome")
            Me.Sequencia = Row("Sequencia")
            Me.Serie = Row("Serie")
        Else

            Dim objEmpresa As New Cliente(Empresa, EndEmpresa)

            Throw New Exception("Não existe Numerador " & Numerador & " Cadastrado para a Empresa " & _
                                Funcoes.AlinharEsquerda(objEmpresa.Nome, 28, ".") & _
                              " - " & Funcoes.AlinharEsquerda(objEmpresa.Cidade, 20, ".") & " " & objEmpresa.CodigoEstado & _
                              " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(objEmpresa.Codigo), 18, ".") & _
                              "-" & objEmpresa.CodigoEndereco.ToString() & "-" & objEmpresa.Reduzido)
        End If
    End Sub

    Public Sub New(ByVal Numerador As Integer)
        Dim Banco As New AcessaBanco
        Dim strSQL As String
        Dim Servidor As String = HttpContext.Current.Session("ssNomeServidor")

        strSQL = " SELECT Empresa_Id, EndEmpresa_Id, Numerador_Id, Descricao, Sequencia, isnull(Serie, '') as Serie " & vbCrLf & _
                 "  FROM Numerador " & vbCrLf & _
                 " WHERE Empresa_Id    ='" & Servidor & "'" & vbCrLf & _
                 "   AND EndEmpresa_Id = 0" & vbCrLf & _
                 "   AND Numerador_Id  = " & Numerador & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "numer")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            Me.Empresa = row("Empresa_Id")
            Me.EndEmpresa = row("EndEmpresa_Id")
            Me.Numerador = row("Numerador_Id")
            Me.Descricao = row("Descricao")
            Me.Sequencia = row("Sequencia")
            Me.Serie = row("Serie")
        Else
            Throw New Exception("Não existe Numerador " & Numerador & " Cadastrado para o servidor " & Servidor)
        End If
    End Sub

#End Region

#Region "Propriedades"
    Public Property IUD() As Char
        Get
            Return _IUD
        End Get
        Set(ByVal value As Char)
            value = IUD
        End Set
    End Property

    Public Property Reduzido() As String
        Get
            Return _sReduzido
        End Get
        Set(ByVal value As String)
            _sReduzido = value
        End Set
    End Property

    Public Property Empresa() As String
        Get
            Return _sEmpresa
        End Get
        Set(ByVal Value As String)
            _sEmpresa = Value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _iEndEmpresa
        End Get
        Set(ByVal Value As Integer)
            _iEndEmpresa = Value
        End Set
    End Property

    Public Property Nome() As String
        Get
            Return _sNome
        End Get
        Set(ByVal value As String)
            _sNome = value
        End Set
    End Property

    Public Property Numerador() As Integer
        Get
            Return _iNumerador
        End Get
        Set(ByVal Value As Integer)
            _iNumerador = Value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _sSerie
        End Get
        Set(ByVal value As String)
            _sSerie = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _sDescricao
        End Get
        Set(ByVal Value As String)
            _sDescricao = Value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _iSequencia
        End Get
        Set(ByVal Value As Integer)
            _iSequencia = Value
        End Set
    End Property


#End Region

#Region "Métodos"

    Public Shared Erro As Exception

    Public Shared Function PegarNumero(ByVal Servidor As String, ByVal Tipo As eTiposNumerador) As Integer
        Return PegarNumero(Servidor, 0, Tipo)
    End Function

    Public Shared Function PegarNumeroRemoto(ByVal pBanco As Integer, ByVal Servidor As String, ByVal Tipo As eTiposNumerador) As Integer
        Dim Banco As New AcessaBanco(pBanco, Servidor)
        Dim intNumerador As Integer = 0

        Erro = Nothing

        Try
            Dim strSQL As String = "EXEC sp_Numerador '" & Servidor & "', " & 0 & ", " & Convert.ToString(Convert.ToInt32(Tipo))
            Dim dsBanco As DataSet = Banco.ConsultaDataSet(strSQL, "Numerador")

            If dsBanco.Tables(0).Rows.Count > 0 Then
                intNumerador = Convert.ToInt32(dsBanco.Tables(0).Rows(0)("Sequencia"))
            Else
                intNumerador = -1
            End If

            Return intNumerador
        Catch ex As Exception
            Erro = ex
            Return -2
        Finally
            Banco = Nothing
        End Try
    End Function

    Public Shared Function PegarNumero(ByVal Empresa As String, ByVal Endereco As Integer, ByVal Tipo As eTiposNumerador) As Integer
        Dim Banco As New AcessaBanco()
        Dim intNumerador As Integer = 0

        Erro = Nothing

        Try
            Dim strSQL As String = "EXEC sp_Numerador '" & Empresa & "', " & Endereco.ToString & ", " & Convert.ToString(Convert.ToInt32(Tipo))
            Dim dsBanco As DataSet = Banco.ConsultaDataSet(strSQL, "Numerador")

            If dsBanco.Tables(0).Rows.Count > 0 Then
                intNumerador = Convert.ToInt32(dsBanco.Tables(0).Rows(0)("Sequencia"))
            Else
                'intNumerador = -1
                Throw New Exception(String.Format("Numerador n:{0} não cadastrado!", Convert.ToString(Convert.ToInt32(Tipo))))
            End If

            Return intNumerador
        Catch ex As Exception
            Erro = ex
            Return -2
        Finally
            Banco = Nothing
        End Try
    End Function

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As ArrayList = New ArrayList()

        Sqls.Clear()
        Sqls.Add(SalvarSql())
        If Banco.GravaBanco(Sqls) Then
            Me.IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Function SalvarSql() As String
        Dim strSQL As String

        Select Case Me.IUD
            Case "I"
                strSQL = "	  Insert into Numerador(Empresa_Id, EndEmpresa_Id, Numerador_Id, Descricao, Sequencia, Serie)" & vbCrLf & _
                         "                  values('" & Me.Empresa & "'," & Me.EndEmpresa & "," & Me.Numerador & ",'" & Me.Descricao & "'," & Me.Sequencia & ",'" & Me.Serie & "')" & vbCrLf
                Return strSQL
            Case "U"
                strSQL = "	  update Numerador set " & vbCrLf & _
                         "	      Descricao       ='" & Me.Descricao & "'" & vbCrLf & _
                         "	     ,Sequencia       = " & Me.Sequencia & vbCrLf & _
                         "      ,Serie           ='" & Me.Serie & "'" & vbCrLf & _
                         "	  Where Empresa_id    ='" & Me.Empresa & "'" & vbCrLf & _
                         "		and EndEmpresa_id = " & Me.EndEmpresa & vbCrLf & _
                         "		and Numerador_id  = " & Me.Numerador & vbCrLf
                Return strSQL
            Case "D"
                strSQL = " Delete Numerador " & vbCrLf & _
                         " Where Empresa_id    ='" & Me.Empresa & "'" & vbCrLf & _
                         "	  and EndEmpresa_id = " & Me.EndEmpresa & vbCrLf & _
                         "	  and Numerador_id  = " & Me.Numerador & vbCrLf
                Return strSQL
        End Select
        Return String.Empty
    End Function

    Public Function IncrementarNumerador(Optional ByVal pGravaServidor As Boolean = False, Optional ByVal pIncremento As Integer = 0) As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As ArrayList = New ArrayList()

        Sqls.Clear()
        Sqls.Add(IncrementarNumeradorSql(pGravaServidor))
        If Banco.GravaBanco(Sqls) Then
            Me.IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Function IncrementarNumeradorSql(Optional ByVal pGravaServidor As Boolean = False, Optional ByVal pIncremento As Integer = 0) As String
        Dim strSQL As String

        strSQL = "	  update Numerador set " & vbCrLf & _
                 "	     Sequencia        = Sequencia + " & IIf(pIncremento = 0, "1", pIncremento.ToString) & vbCrLf
        If pGravaServidor Then
            strSQL &= " Where Empresa_id    ='" & HttpContext.Current.Session("ssNomeServidor") & "'" & vbCrLf & _
                      "	  and EndEmpresa_id = 0 " & vbCrLf
        Else
            strSQL &= " Where Empresa_id    ='" & Me.Empresa & "'" & vbCrLf & _
                      "	  and EndEmpresa_id = " & Me.EndEmpresa & vbCrLf
        End If
        strSQL &= " and Numerador_id  = " & Me.Numerador
        Return strSQL
    End Function

#End Region

End Class