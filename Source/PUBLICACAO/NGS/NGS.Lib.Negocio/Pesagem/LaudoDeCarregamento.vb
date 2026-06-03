Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

'************************************************************************************************************************************************************
'******************************************************** Lista da Classe Laudo de Carregamento *************************************************************
'************************************************************************************************************************************************************
<Serializable()>
Public Class ListLaudoDeCarregamento
    Inherits List(Of LaudoDeCarregamento)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pLaudoDeCarregamento As LaudoDeCarregamento,
                   Optional ByVal DataInicial As String = "",
                   Optional ByVal DataFinal As String = "")

        Dim Banco As New AcessaBanco()
        Dim Sql As String
        Dim strAndWhere As String = "WHERE"

        Try
            Sql = "SELECT P.Empresa_Id, P.EndEmpresa_Id, P.LaudoDeCarregamento_Id, P.Sequencia_Id," & vbCrLf &
                  "       P.OrdemDeCarregamento, " & vbCrLf &
                  "       P.Transportador, P.EndTransportador," & vbCrLf &
                  "       P.Placa, P.Motorista, P.EndMotorista," & vbCrLf &
                  "       P.ViaTransporte, " & vbCrLf &
                  "       P.Movimento, P.Situacao," & vbCrLf &
                  "       P.PrimeiraPesagem, P.SegundaPesagem, P.Liquido," & vbCrLf &
                  "       P.EntradaBalanca," & vbCrLf &
                  "	      isnull(P.SaidaBalanca,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS SaidaBalanca," & vbCrLf &
                  "       P.Observacoes," & vbCrLf &
                  "       P.PesagemManual," & vbCrLf &
                  "       P.UsuarioInclusao," & vbCrLf &
                  "       P.UsuarioInclusaoData," & vbCrLf &
                  "       isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao," & vbCrLf &
                  "       isnull(P.UsuarioAlteracaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioAlteracaoData," & vbCrLf &
                  "       isnull(P.UsuarioReimpressao,'') AS UsuarioReimpressao," & vbCrLf &
                  "       isnull(P.UsuarioReimpressaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioReimpressaoData," & vbCrLf &
                  "       isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento," & vbCrLf &
                  "       isnull(P.UsuarioCancelamentoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioCancelamentoData" & vbCrLf &
                  "  FROM LaudoDeCarregamento P" & vbCrLf

            If pLaudoDeCarregamento.CodigoEmpresa.Length > 0 Then
                Sql &= strAndWhere & " P.Empresa_Id = '" & pLaudoDeCarregamento.CodigoEmpresa & "' AND P.EndEmpresa_id = " & pLaudoDeCarregamento.EnderecoEmpresa & " " & vbCrLf
                strAndWhere = " AND "
            End If

            If pLaudoDeCarregamento.Codigo > 0 Then
                Sql &= strAndWhere & " P.LaudoDeCarregamento_Id = " & pLaudoDeCarregamento.Codigo & vbCrLf
                strAndWhere = " AND "
            Else
                If Not String.IsNullOrWhiteSpace(DataInicial) AndAlso Not String.IsNullOrWhiteSpace(DataFinal) Then
                    Sql &= strAndWhere & " P.Movimento BETWEEN '" & CDate(DataInicial).ToString("yyyy-MM-dd") & "' AND '" & CDate(DataFinal).ToString("yyyy-MM-dd") & "'" & vbCrLf
                End If
            End If

            If pLaudoDeCarregamento.CodigoOrdemDeCarregamento > 0 Then
                Sql &= strAndWhere & " P.OrdemDeCarregamento = " & pLaudoDeCarregamento.CodigoOrdemDeCarregamento & " " & vbCrLf
                strAndWhere = " AND "
            End If

            If pLaudoDeCarregamento.CodigoTransportador.Length > 0 Then
                Sql &= strAndWhere & " P.Transportador = '" & pLaudoDeCarregamento.CodigoTransportador & "' AND P.EndTransportador = " & pLaudoDeCarregamento.EnderecoTransportador & vbCrLf
                strAndWhere = " AND "
            End If

            If pLaudoDeCarregamento.CodigoMotorista.Length > 0 Then
                Sql &= strAndWhere & " P.Motorista = '" & pLaudoDeCarregamento.CodigoMotorista & "' AND P.EndMotorista = " & pLaudoDeCarregamento.EnderecoMotorista & vbCrLf
                strAndWhere = " AND "
            End If

            If pLaudoDeCarregamento.CodigoSituacao > 0 Then
                Sql &= strAndWhere & " P.Situacao = " & pLaudoDeCarregamento.CodigoSituacao & vbCrLf
                strAndWhere = " AND "
            End If

            If pLaudoDeCarregamento.CodigoPlaca.Length > 0 Then
                Sql &= strAndWhere & " P.Placa = '" & pLaudoDeCarregamento.CodigoPlaca & "' " & vbCrLf
                strAndWhere = " AND "
            End If

            Sql &= " AND P.Sequencia_Id = 0 " & vbCrLf

            Sql &= " ORDER BY P.LaudoDeCarregamento_Id ASC"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "LaudoDeCarregamento")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim P As New LaudoDeCarregamento

                P.CodigoEmpresa = row("Empresa_Id")
                P.EnderecoEmpresa = row("EndEmpresa_Id")
                P.Codigo = row("LaudoDeCarregamento_Id")
                P.Sequencia = row("Sequencia_Id")
                P.CodigoOrdemDeCarregamento = row("OrdemDeCarregamento")
                P.CodigoTransportador = row("Transportador")
                P.EnderecoTransportador = row("EndTransportador")
                P.CodigoPlaca = row("Placa")
                P.CodigoMotorista = row("Motorista")
                P.EnderecoMotorista = row("EndMotorista")
                P.CodigoViaDeTransporte = row("ViaTransporte")
                P.Movimento = row("Movimento")
                P.CodigoSituacao = row("Situacao")
                P.PrimeiraPesagem = row("PrimeiraPesagem")
                P.SegundaPesagem = row("SegundaPesagem")
                P.Liquido = row("Liquido")
                P.EntradaBalanca = row("EntradaBalanca")
                If Not IsDBNull(row("SaidaBalanca")) Then P.SaidaBalanca = row("SaidaBalanca")
                P.Observacoes = row("Observacoes").ToString()
                P.PesagemManual = row("PesagemManual")
                P.UsuarioInclusao = row("UsuarioInclusao")
                P.DataInclusao = DateTime.Now
                P.UsuarioAlteracao = row("UsuarioAlteracao")
                P.DataAlteracao = DateTime.Now
                P.UsuarioReimpressao = row("UsuarioReimpressao")
                P.DataReimpressao = DateTime.Now
                P.UsuarioCancelamento = row("UsuarioCancelamento")
                P.DataCancelamento = row("UsuarioCancelamentoData")

                Me.Add(P)
            Next
        Catch ex As Exception

        Finally
            Banco = Nothing
        End Try
    End Sub

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each P As LaudoDeCarregamento In Me
            If P.IUD <> "" Then
                P.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

'************************************************************************************************************************************************************
'*************************************************************** Classe Base LaudoDeCarregamento ************************************************************************
'************************************************************************************************************************************************************
<Serializable()>
Public Class LaudoDeCarregamento

#Region "Construtores"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pLaudoDeCarregamento As Integer, Optional ByVal pSequencia As Integer = 0)
        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT P.Empresa_Id, P.EndEmpresa_Id, P.LaudoDeCarregamento_Id, P.Sequencia_Id," & vbCrLf &
                  "       P.OrdemDeCarregamento, " & vbCrLf &
                  "       P.Transportador, P.EndTransportador," & vbCrLf &
                  "       P.Placa, P.Motorista, P.EndMotorista," & vbCrLf &
                  "       P.ViaTransporte, " & vbCrLf &
                  "       P.Movimento, P.Situacao," & vbCrLf &
                  "       P.PrimeiraPesagem, P.SegundaPesagem, P.Liquido," & vbCrLf &
                  "       P.EntradaBalanca," & vbCrLf &
                  "	      isnull(P.SaidaBalanca,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS SaidaBalanca," & vbCrLf &
                  "       P.Observacoes," & vbCrLf &
                  "       P.PesagemManual," & vbCrLf &
                  "       P.UsuarioInclusao," & vbCrLf &
                  "       P.UsuarioInclusaoData," & vbCrLf &
                  "       isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao," & vbCrLf &
                  "       isnull(P.UsuarioAlteracaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioAlteracaoData," & vbCrLf &
                  "       isnull(P.UsuarioReimpressao,'') AS UsuarioReimpressao," & vbCrLf &
                  "       isnull(P.UsuarioReimpressaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioReimpressaoData," & vbCrLf &
                  "       isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento," & vbCrLf &
                  "       isnull(P.UsuarioCancelamentoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioCancelamentoData" & vbCrLf &
                  "  FROM LaudoDeCarregamento P" & vbCrLf &
                  " WHERE P.Empresa_Id    ='" & pEmpresa & "'" & vbCrLf &
                  "   AND P.EndEmpresa_Id = " & pEndEmpresa & vbCrLf &
                  "   AND P.LaudoDeCarregamento_Id    = " & pLaudoDeCarregamento & vbCrLf &
                  "   and p.Sequencia_Id  = " & pSequencia & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "LaudoDeCarregamento")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.CodigoEmpresa = row("Empresa_Id")
            Me.EnderecoEmpresa = row("EndEmpresa_Id")
            Me.Codigo = row("LaudoDeCarregamento_Id")
            Me.Sequencia = row("Sequencia_Id")
            Me.CodigoOrdemDeCarregamento = row("OrdemDeCarregamento")
            Me.CodigoTransportador = row("Transportador")
            Me.EnderecoTransportador = row("EndTransportador")
            Me.CodigoPlaca = row("Placa")
            Me.CodigoMotorista = row("Motorista")
            Me.EnderecoMotorista = row("EndMotorista")
            Me.CodigoViaDeTransporte = row("ViaTransporte")
            Me.Movimento = row("Movimento")
            Me.CodigoSituacao = row("Situacao")
            Me.PrimeiraPesagem = row("PrimeiraPesagem")
            Me.SegundaPesagem = row("SegundaPesagem")
            Me.Liquido = row("Liquido")
            Me.EntradaBalanca = row("EntradaBalanca")
            If Not IsDBNull(row("SaidaBalanca")) Then Me.SaidaBalanca = row("SaidaBalanca")
            Me.Observacoes = row("Observacoes").ToString()
            Me.PesagemManual = row("PesagemManual")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.DataInclusao = DateTime.Now
            Me.UsuarioAlteracao = row("UsuarioAlteracao")
            Me.DataAlteracao = DateTime.Now
            Me.UsuarioReimpressao = row("UsuarioReimpressao")
            Me.DataReimpressao = DateTime.Now
            Me.UsuarioCancelamento = row("UsuarioCancelamento")
            Me.DataCancelamento = row("UsuarioCancelamentoData")

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Public Erro As Exception
    Private _SelecionaLaudoDeCarregamento As Boolean = False

    Private _IUD As String

    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente

    Private _Codigo As Integer
    Private _Sequencia As Integer
    Private _SequenciaBanco As Integer

    'Private _OrdemDeCarregamento As Integer

    Private _CodigoTransportador As String = ""
    Private _EnderecoTransportador As Integer
    Private _Transportador As Cliente

    Private _CodigoPlaca As String = ""
    Private _Placa As Placa

    Private _CodigoMotorista As String = ""
    Private _EnderecoMotorista As Integer
    Private _Motorista As Cliente

    Private _CodigoOrdemDeCarregamento As Integer
    'Private _OrdemDeCarregamento As OrdemDeCarregamento

    Private _CodigoViaDeTransporte As Integer
    Private _ViaDeTransporte As ViaDeTransporte

    Private _Movimento As DateTime = Date.Today

    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao

    Private _PrimeiraPesagem As Decimal
    Private _SegundaPesagem As Decimal
    Private _Liquido As Decimal

    Private _EntradaBalanca As DateTime
    Private _SaidaBalanca As DateTime

    Private _Observacoes As String = ""

    Private _PesagemManual As Boolean

    '****** Controle de Usuario **********
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime?
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime?
    Private _UsuarioReimpressao As String
    Private _UsuarioReimpressaoData As DateTime?
    Private _UsuarioCancelamento As String
    Private _UsuarioCancelamentoData As DateTime?

#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Trim.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
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

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    'Public Property OrdemDeCarregamento() As Integer
    '    Get
    '        Return OrdemDeCarregamento
    '    End Get
    '    Set(ByVal value As Integer)
    '        _OrdemDeCarregamento = value
    '    End Set
    'End Property

    Public Property CodigoTransportador() As String
        Get
            Return _CodigoTransportador
        End Get
        Set(ByVal value As String)
            _CodigoTransportador = value
        End Set
    End Property

    Public Property EnderecoTransportador() As Integer
        Get
            Return _EnderecoTransportador
        End Get
        Set(ByVal value As Integer)
            _EnderecoTransportador = value
        End Set
    End Property

    Public Property Transportador() As Cliente
        Get
            If _Transportador Is Nothing And _CodigoTransportador.Trim.Length > 0 Then _Transportador = New Cliente(_CodigoTransportador, _EnderecoTransportador)
            Return _Transportador
        End Get
        Set(ByVal value As Cliente)
            _Transportador = value
        End Set
    End Property

    Public Property CodigoPlaca() As String
        Get
            Return _CodigoPlaca
        End Get
        Set(ByVal value As String)
            _CodigoPlaca = value
        End Set
    End Property

    Public Property Placa() As Placa
        Get
            If _Placa Is Nothing And _CodigoPlaca.Trim.Length > 0 Then _Placa = New Placa(_CodigoPlaca)
            Return _Placa
        End Get
        Set(ByVal value As Placa)
            _Placa = value
        End Set
    End Property

    Public Property CodigoMotorista() As String
        Get
            Return _CodigoMotorista
        End Get
        Set(ByVal value As String)
            _CodigoMotorista = value
        End Set
    End Property

    Public Property EnderecoMotorista() As Integer
        Get
            Return _EnderecoMotorista
        End Get
        Set(ByVal value As Integer)
            _EnderecoMotorista = value
        End Set
    End Property

    Public Property Motorista() As Cliente
        Get
            If _Motorista Is Nothing And _CodigoMotorista.Trim.Length > 0 Then _Motorista = New Cliente(_CodigoMotorista, _EnderecoMotorista)
            Return _Motorista
        End Get
        Set(ByVal value As Cliente)
            _Motorista = value
        End Set
    End Property

    Public Property CodigoOrdemDeCarregamento() As Integer
        Get
            Return _CodigoOrdemDeCarregamento
        End Get
        Set(ByVal value As Integer)
            _CodigoOrdemDeCarregamento = value
        End Set
    End Property

    'Public Property OrdemDeCarregamento() As OrdemDeCarregamento
    '    Get
    '        If _OrdemDeCarregamento Is Nothing And _CodigoOrdemDeCarregamento > 0 Then _OrdemDeCarregamento = New OrdemDeCarregamento(_CodigoTransportador, _EnderecoTransportador, _CodigoOrdemDeCarregamento)
    '        Return _OrdemDeCarregamento
    '    End Get
    '    Set(ByVal value As OrdemDeCarregamento)
    '        _OrdemDeCarregamento = value
    '    End Set
    'End Property

    Public Property CodigoViaDeTransporte() As Integer
        Get
            Return _CodigoViaDeTransporte
        End Get
        Set(ByVal value As Integer)
            _CodigoViaDeTransporte = value
        End Set
    End Property

    Public Property ViaDeTransporte() As ViaDeTransporte
        Get
            If _ViaDeTransporte Is Nothing And _CodigoViaDeTransporte > 0 Then _ViaDeTransporte = New ViaDeTransporte(_CodigoViaDeTransporte)
            Return _ViaDeTransporte
        End Get
        Set(ByVal value As ViaDeTransporte)
            _ViaDeTransporte = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
        End Set
    End Property

    Public Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Situacao)
            _Situacao = value
        End Set
    End Property

    Public Property PrimeiraPesagem() As Decimal
        Get
            Return _PrimeiraPesagem
        End Get
        Set(ByVal value As Decimal)
            _PrimeiraPesagem = value
        End Set
    End Property

    Public Property SegundaPesagem() As Decimal
        Get
            Return _SegundaPesagem
        End Get
        Set(ByVal value As Decimal)
            _SegundaPesagem = value
        End Set
    End Property

    Public Property Liquido() As Decimal
        Get
            Return _Liquido
        End Get
        Set(ByVal value As Decimal)
            _Liquido = value
        End Set
    End Property

    Public Property EntradaBalanca() As DateTime
        Get
            Return _EntradaBalanca
        End Get
        Set(ByVal value As DateTime)
            _EntradaBalanca = value
        End Set
    End Property

    Public Property SaidaBalanca() As DateTime
        Get
            Return _SaidaBalanca
        End Get
        Set(ByVal value As DateTime)
            _SaidaBalanca = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property PesagemManual() As Boolean
        Get
            Return _PesagemManual
        End Get
        Set(ByVal value As Boolean)
            _PesagemManual = value
        End Set
    End Property

    '********** Controle de Usuarios  *************
    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime?
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioInclusaoData = value
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

    Public Property DataAlteracao() As DateTime?
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property UsuarioReimpressao() As String
        Get
            Return _UsuarioReimpressao
        End Get
        Set(ByVal value As String)
            _UsuarioReimpressao = value
        End Set
    End Property

    Public Property DataReimpressao() As DateTime?
        Get
            Return _UsuarioReimpressaoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioReimpressaoData = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime?
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioCancelamentoData = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Dim usa As DateTime = DateTime.Now
        Dim urd As DateTime = DateTime.Now

        Select Case Me.IUD
            Case "I"
                If Me.Codigo.Equals(0) Then
                    Dim N As New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 101)
                    Me.Codigo = N.Sequencia + 1
                    Sqls.Add(N.IncrementarNumeradorSql)
                End If

                sql = "Insert Into LaudoDeCarregamento " & vbCrLf &
                      " (Empresa_Id, EndEmpresa_Id, LaudoDeCarregamento_Id, Sequencia_Id, OrdemDeCarregamento, " & vbCrLf &
                      " Transportador, EndTransportador, " & vbCrLf &
                      " Placa, Motorista, EndMotorista, ViaTransporte, " & vbCrLf &
                      " Movimento, Situacao, " & vbCrLf &
                      " PrimeiraPesagem, SegundaPesagem, Liquido, " & vbCrLf &
                      " EntradaBalanca, SaidaBalanca, Observacoes, PesagemManual, " & vbCrLf &
                      " UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf &
                      " UsuarioReimpressao, UsuarioReimpressaoData, UsuarioCancelamento, UsuarioCancelamentoData) " & vbCrLf &
                      " Values ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & "," & Me.Sequencia & "," & Me.CodigoOrdemDeCarregamento & "," & vbCrLf &
                      "'" & Me.CodigoTransportador & "'," & Me.EnderecoTransportador & "," & vbCrLf &
                      "'" & Me.CodigoPlaca & "','" & Me.CodigoMotorista & "', " & Me.EnderecoMotorista & "," & Me.CodigoViaDeTransporte & "," & vbCrLf &
                      "'" & Me.Movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'," & Me.CodigoSituacao & "," & vbCrLf &
                      Str(Me.PrimeiraPesagem) & "," & Str(Me.SegundaPesagem) & "," & Str(Me.Liquido) & "," & vbCrLf &
                      "'" & Me.EntradaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "', " & IIf(Me.SegundaPesagem > 0, "'" & Me.SaidaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & "," & vbCrLf &
                      "'" & Me.Observacoes & "'," & CByte(Me.PesagemManual) & "," & vbCrLf &
                      "'" & Me.UsuarioInclusao & "', GETDATE(), " & IIf(String.IsNullOrWhiteSpace(UsuarioAlteracao), "NULL", "'" & UsuarioAlteracao & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf &
                      "" & IIf(String.IsNullOrWhiteSpace(UsuarioReimpressao), "NULL", "'" & UsuarioReimpressao & "'") & ",'" & urd.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
                      IIf(String.IsNullOrWhiteSpace(UsuarioCancelamento), "NULL", "'" & UsuarioCancelamento & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "')"

                Sqls.Add(sql)

            Case "U"
                sql = "Update LaudoDeCarregamento set " & vbCrLf &
                      "     OrdemDeCarregamento     = " & Me.CodigoOrdemDeCarregamento & vbCrLf &
                      "    ,Transportador           ='" & Me.CodigoTransportador & "'" & vbCrLf &
                      "    ,EndTransportador        = " & Me.EnderecoTransportador & vbCrLf &
                      "    ,Placa                   ='" & Me.CodigoPlaca & "'" & vbCrLf &
                      "    ,Motorista               ='" & Me.CodigoMotorista & "'" & vbCrLf &
                      "    ,EndMotorista            = " & Me.EnderecoMotorista & vbCrLf &
                      "    ,ViaTransporte           = " & Me.CodigoViaDeTransporte & vbCrLf &
                      "    ,Movimento               ='" & Me.Movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                      "    ,Situacao                = " & Me.CodigoSituacao & vbCrLf &
                      "    ,PrimeiraPesagem         = " & Me.PrimeiraPesagem & vbCrLf &
                      "    ,SegundaPesagem          = " & Me.SegundaPesagem & vbCrLf &
                      "    ,Liquido                 = " & Me.Liquido & vbCrLf &
                      "    ,SaidaBalanca            = " & IIf(Me.SegundaPesagem > 0, "'" & Me.SaidaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & vbCrLf &
                      "    ,Observacoes             ='" & Me.Observacoes & "'" & vbCrLf &
                      "    ,UsuarioAlteracao        ='" & Me.UsuarioAlteracao & "'" & vbCrLf &
                      "    ,UsuarioAlteracaoData    = '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                      "    ,UsuarioReimpressao      ='" & Me.UsuarioReimpressao & "'" & vbCrLf &
                      "    ,UsuarioReimpressaoData  ='" & String.Format("{0:yyyy-MM-dd HH:mm:ss}", Me.DataReimpressao) & "'" & vbCrLf &
                      "    ,UsuarioCancelamento     ='" & _UsuarioCancelamento & "'" & vbCrLf &
                      "    ,UsuarioCancelamentoData ='" & String.Format("{0:yyyy-MM-dd HH:mm:ss}", Me.DataCancelamento) & "'" & vbCrLf &
                      "	Where Empresa_Id             ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id          = " & Me.EnderecoEmpresa & vbCrLf &
                      "   and LaudoDeCarregamento_Id = " & Me.Codigo & vbCrLf &
                      "   and Sequencia_Id           = " & Me.Sequencia

                Sqls.Add(sql)
            Case "D"

                sql = " Update LaudoDeCarregamento Set" & vbCrLf &
                      "     Situacao                = 2" & vbCrLf &
                      "    ,Observacoes             ='" & Me.Observacoes & "'" & vbCrLf &
                      "    ,UsuarioCancelamento     = '" & Me.UsuarioCancelamento & "'" & vbCrLf &
                      "    ,UsuarioCancelamentoData = '" & CDate(Me.DataCancelamento).ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                      "	Where Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                      "   and LaudoDeCarregamento_Id    = " & Me.Codigo & vbCrLf &
                      "   and Sequencia_Id  = " & Me.Sequencia

                Sqls.Add(sql)

            Case "S" 'Gravar Sequência das Alterações

                sql = "Insert Into LaudoDeCarregamento " & vbCrLf &
                      " (Empresa_Id, EndEmpresa_Id, LaudoDeCarregamento_Id, Sequencia_Id, OrdemDeCarregamento, " & vbCrLf &
                      " Transportador, EndTransportador, " & vbCrLf &
                      " Placa, Motorista, EndMotorista, ViaTransporte, " & vbCrLf &
                      " Movimento, Situacao, " & vbCrLf &
                      " PrimeiraPesagem, SegundaPesagem, Liquido, " & vbCrLf &
                      " EntradaBalanca, SaidaBalanca, Observacoes, PesagemManual, " & vbCrLf &
                      " UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf &
                      " UsuarioReimpressao, UsuarioReimpressaoData, UsuarioCancelamento, UsuarioCancelamentoData) " & vbCrLf &
                      " Values ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & "," & Me.Sequencia & "," & Me.CodigoOrdemDeCarregamento & "," & vbCrLf &
                      "'" & Me.CodigoTransportador & "'," & Me.EnderecoTransportador & "," & vbCrLf &
                      "'" & Me.CodigoPlaca & "','" & Me.CodigoMotorista & "', " & Me.EnderecoMotorista & "," & Me.CodigoViaDeTransporte & "," & vbCrLf &
                      "'" & Me.Movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'," & Me.CodigoSituacao & "," & vbCrLf &
                      Str(Me.PrimeiraPesagem) & "," & Str(Me.SegundaPesagem) & "," & Str(Me.Liquido) & "," & vbCrLf &
                      "'" & Me.EntradaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "', " & IIf(Me.SegundaPesagem > 0, "'" & Me.SaidaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & "," & vbCrLf &
                      "'" & Me.Observacoes & "'," & CByte(Me.PesagemManual) & "," & vbCrLf &
                      "'" & Me.UsuarioInclusao & "', GETDATE(), " & IIf(String.IsNullOrWhiteSpace(UsuarioAlteracao), "NULL", "'" & UsuarioAlteracao & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf &
                      "" & IIf(String.IsNullOrWhiteSpace(UsuarioReimpressao), "NULL", "'" & UsuarioReimpressao & "'") & ",'" & urd.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
                      IIf(String.IsNullOrWhiteSpace(UsuarioCancelamento), "NULL", "'" & UsuarioCancelamento & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "')"

                Sqls.Add(sql)

        End Select
    End Sub

#End Region

End Class
