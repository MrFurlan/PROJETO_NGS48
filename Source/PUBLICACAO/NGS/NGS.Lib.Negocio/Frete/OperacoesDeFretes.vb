Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListOperacoesDeFretes
    Inherits List(Of OperacoesDeFretes)

    Public Sub New(Optional ByVal Where As String = "")
        Dim sql As String = String.Empty
        sql &= " SELECT OPF.Operacao_Id, OPF.SubOperacoes_Id, OPF.TipoPessoa_Id, OPF.OpDestino_Id, OPF.SubOpDestino_Id,  OPF.TipoOperacao, " & vbCrLf &
               "        OPF.PrestacaoDeServico, OPF.OpAnulacao, OPF.SubOpAnulacao, OPF.OpContrapartida, OPF.SubOpContrapartida," & vbCrLf &
               "        SOP.EntradaSaida, SOP.Classe, isnull(OPF.Finalidade,0) as Finalidade," & vbCrLf &
               "        OPF.UsuarioInclusao, OPF.UsuarioInclusaoData, OPF.UsuarioAlteracao, OPF.UsuarioAlteracaoData " & vbCrLf &
               "   FROM OperacoesDeFretes OPF " & vbCrLf &
               "  INNER JOIN SubOperacoes SOP " & vbCrLf &
               "     ON SOP.Operacao_Id = OPF.Operacao_Id " & vbCrLf &
               "    AND SOP.SubOperacoes_Id = OPF.SubOperacoes_Id " & vbCrLf
        If Not String.IsNullOrWhiteSpace(Where) Then
            sql &= "  WHERE " & Where & vbCrLf
        End If
        sql &= "  ORDER BY OPF.Operacao_Id, OPF.SubOperacoes_Id"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "OperacoesDeFretes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New OperacoesDeFretes
                obj.OperacaoId = row("Operacao_Id")
                obj.SubOperacaoId = row("SubOperacoes_Id")
                obj.EntradaSaida = row("EntradaSaida")
                obj.TipoPessoaId = row("TipoPessoa_Id")
                obj.Classe = row("Classe")

                If IsDBNull(row("OpDestino_Id")) Then
                    obj.OpDestinoId = New Nullable(Of Integer)
                Else
                    obj.OpDestinoId = CInt(row("OpDestino_Id"))
                End If

                If IsDBNull(row("SubOpDestino_Id")) Then
                    obj.SubOpDestinoId = New Nullable(Of Integer)
                Else
                    obj.SubOpDestinoId = CInt(row("SubOpDestino_Id"))
                End If

                obj.TipoOperacao = CType(CInt(row("TipoOperacao")), eTipoOperacao)

                If IsDBNull(row("PrestacaoDeServico")) Then
                    obj.PrestacaoDeServico = New Nullable(Of ePrestacaoServico)
                Else
                    obj.PrestacaoDeServico = CType(CInt(row("PrestacaoDeServico")), ePrestacaoServico)
                End If

                If IsDBNull(row("OpAnulacao")) Then
                    obj.OpAnulacao = New Nullable(Of Integer)
                Else
                    obj.OpAnulacao = CInt(row("OpAnulacao"))
                End If

                If IsDBNull(row("SubOpAnulacao")) Then
                    obj.SubOpAnulacao = New Nullable(Of Integer)
                Else
                    obj.SubOpAnulacao = CInt(row("SubOpAnulacao"))
                End If

                If IsDBNull(row("OpContrapartida")) Then
                    obj.OpContrapartida = New Nullable(Of Integer)
                Else
                    obj.OpContrapartida = CInt(row("OpContrapartida"))
                End If

                If IsDBNull(row("SubOpContrapartida")) Then
                    obj.SubOpContrapartida = New Nullable(Of Integer)
                Else
                    obj.SubOpContrapartida = CInt(row("SubOpContrapartida"))
                End If

                obj.CodigoFinalidade = row("Finalidade")

                Me.Add(obj)
            Next
        End If
    End Sub

End Class

<Serializable()>
Public Class OperacoesDeFretes
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal operacao As Integer, ByVal suboperacao As Integer, Optional ByVal financeiro As String = "")
        Dim db As New AcessaBanco
        Dim sql As String = String.Empty
        sql &= " SELECT OPF.Operacao_Id, OPF.SubOperacoes_Id, OPF.TipoPessoa_Id, OPF.OpDestino_Id, OPF.SubOpDestino_Id,  OPF.TipoOperacao, " & vbCrLf &
               "        OPF.PrestacaoDeServico, OPF.OpAnulacao, OPF.SubOpAnulacao, OPF.OpContrapartida, OPF.SubOpContrapartida," & vbCrLf &
               "        SUB.EntradaSaida, SUB.Classe, isnull(OPF.Finalidade,0) as Finalidade, " & vbCrLf &
               "        OPF.UsuarioInclusao, OPF.UsuarioInclusaoData, OPF.UsuarioAlteracao, OPF.UsuarioAlteracaoData " & vbCrLf &
               "   FROM OperacoesDeFretes OPF " & vbCrLf &
               "  INNER JOIN SubOperacoes SUB " & vbCrLf &
               "     ON OPF.Operacao_Id = SUB.Operacao_Id " & vbCrLf &
               "    AND OPF.SubOperacoes_Id = SUB.SubOperacoes_Id " & vbCrLf
        If Not String.IsNullOrWhiteSpace(financeiro) Then sql &= "    AND SUB.Financeiro = '" & financeiro & "'" & vbCrLf
        sql &= "  WHERE OPF.Operacao_Id = " & operacao & " AND OPF.SubOperacoes_Id = " & suboperacao

        Dim ds As DataSet = db.ConsultaDataSet(sql, "OperacoesDeFretes")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)



        Me.OperacaoId = row("Operacao_Id")
        Me.SubOperacaoId = row("SubOperacoes_Id")
        Me.TipoPessoaId = row("TipoPessoa_Id")
        Me.Classe = row("Classe")

        If IsDBNull(row("OpDestino_Id")) Then
            Me.OpDestinoId = New Nullable(Of Integer)
        Else
            Me.OpDestinoId = CInt(row("OpDestino_Id"))
        End If

        If IsDBNull(row("SubOpDestino_Id")) Then
            Me.SubOpDestinoId = New Nullable(Of Integer)
        Else
            Me.SubOpDestinoId = CInt(row("SubOpDestino_Id"))
        End If

        If Not IsDBNull(row("PrestacaoDeServico")) Then
            Me.PrestacaoDeServico = CType(row("PrestacaoDeServico"), ePrestacaoServico)
        End If

        If IsDBNull(row("OpAnulacao")) Then
            Me.OpAnulacao = New Nullable(Of Integer)
        Else
            Me.OpAnulacao = CInt(row("OpAnulacao"))
        End If

        If IsDBNull(row("SubOpAnulacao")) Then
            Me.SubOpAnulacao = New Nullable(Of Integer)
        Else
            Me.SubOpAnulacao = CInt(row("SubOpAnulacao"))
        End If

        If IsDBNull(row("OpContrapartida")) Then
            Me.OpContrapartida = New Nullable(Of Integer)
        Else
            Me.OpContrapartida = CInt(row("OpContrapartida"))
        End If

        If IsDBNull(row("SubOpContrapartida")) Then
            Me.SubOpContrapartida = New Nullable(Of Integer)
        Else
            Me.SubOpContrapartida = CInt(row("SubOpContrapartida"))
        End If

        Me.CodigoFinalidade = row("Finalidade")

        Me.TipoOperacao = CType(row("TipoOperacao"), eTipoOperacao)

    End Sub

    Public Sub New(ByVal classe As String,
               ByVal entradaSaida As String,
               ByVal fisicaJuridica As String,
               Optional ByVal financeiro As String = "",
               Optional ByVal tipo As Nullable(Of eTipoOperacao) = Nothing,
               Optional ByVal pFinalidade As Integer = 0,
               Optional ByVal Devolucao As Boolean = False,
               Optional ByVal Estado As String = "",
               Optional ByVal Regiao As String = "")

        Dim db As New AcessaBanco
        Dim sql As String = String.Empty
        sql = " SELECT TOP 1 " & vbCrLf &
              "        OPF.Operacao_Id, OPF.SubOperacoes_Id, OPF.TipoPessoa_Id, OPF.OpDestino_Id, OPF.SubOpDestino_Id,  OPF.TipoOperacao, " & vbCrLf &
              "        OPF.PrestacaoDeServico, OPF.OpAnulacao, OPF.SubOpAnulacao, OPF.OpContrapartida, OPF.SubOpContrapartida," & vbCrLf &
              "        SOP.EntradaSaida, SOP.Classe, isnull(OPF.Finalidade,0) as Finalidade, " & vbCrLf &
              "        OPF.UsuarioInclusao, OPF.UsuarioInclusaoData, OPF.UsuarioAlteracao, OPF.UsuarioAlteracaoData " & vbCrLf &
              "   FROM OperacoesDeFretes OPF " & vbCrLf &
              "  INNER JOIN SubOperacoes SOP " & vbCrLf &
              "     ON SOP.Operacao_Id = OPF.Operacao_Id " & vbCrLf &
              "    AND SOP.SubOperacoes_Id = OPF.SubOperacoes_Id " & vbCrLf

        ' só faz join com OE/ES se for usar Estado/Regiao
        Dim precisaOE As Boolean = Not String.IsNullOrWhiteSpace(Estado) OrElse Not String.IsNullOrWhiteSpace(Regiao)

        If precisaOE Then
            sql &= "    INNER JOIN VW_OperacoesVigentes OE " & vbCrLf &
               "        ON OPF.Operacao_Id      = OE.Operacao " & vbCrLf &
               "            AND OPF.SubOperacoes_Id = OE.SubOperacao " & vbCrLf &
               "        INNER JOIN Estados ES " & vbCrLf &
               "            ON ES.Estado_Id = OE.EstadoDestino " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(financeiro) Then
            sql &= "    And SOP.Financeiro = '" & financeiro & "'" & vbCrLf
        End If

        sql &= "   LEFT JOIN SubOperacoes SOPD " & vbCrLf &
               "     ON SOPD.SubOperacoes_Id = OPF.SubOpDestino_Id " & vbCrLf &
               "    AND SOPD.Operacao_Id     = OPF.OpDestino_Id " & vbCrLf &
               "  WHERE sop.EntradaSaida = '" & entradaSaida & "' " & vbCrLf

        If classe = eClassesOperacoes.AFIXAR.ToString Then
            sql &= "    AND SOP.Classe = " & IIf(entradaSaida = "E", "'" & eClassesOperacoes.COMPRAS.ToString & "'", "'" & eClassesOperacoes.VENDAS.ToString & "'") & vbCrLf
        Else
            sql &= "    AND SOP.Classe = (SELECT TOP 1 ClassificacaoDeFrete FROM ClassesDeOperacoes WHERE Classe_Id = '" & classe & "') " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(fisicaJuridica) Then sql &= " AND opf.TipoPessoa_Id = '" & fisicaJuridica & "' " & vbCrLf

        If tipo IsNot Nothing Then sql &= "    AND OPF.TipoOperacao = '" & CInt(tipo) & "'" & vbCrLf

        sql &= "    AND isnull(OPF.Finalidade,0) = " & pFinalidade

        sql &= "    AND SOP.Devolucao = '" & IIf(Devolucao, "S", "N") & "'" & vbCrLf

        If precisaOE Then
            sql &= "    AND OE.ativo = 1 " & vbCrLf
            sql &= "    AND (OE.EstadoDestino = '" & Estado & "' OR OE.EstadoDestino = '" & Regiao & "') " & vbCrLf
        End If

        sql &= "    AND isnull(OPF.Finalidade,0) = " & pFinalidade

        If precisaOE Then
            sql &= "ORDER BY " & vbCrLf
            sql &= "    CASE " & vbCrLf
            sql &= "        WHEN 'PR' <> OE.EstadoDestino THEN 0  -- interestadual (primeiro) " & vbCrLf
            sql &= "        ELSE 1                                -- mesma região (depois) " & vbCrLf
            sql &= "    END " & vbCrLf
        End If

        Dim ds As DataSet = db.ConsultaDataSet(sql, "OperacoesDeFretes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Me.OperacaoId = ds.Tables(0).Rows(0)("Operacao_Id")
            Me.SubOperacaoId = ds.Tables(0).Rows(0)("SubOperacoes_Id")
            Me.TipoPessoaId = ds.Tables(0).Rows(0)("TipoPessoa_Id")
            Me.Classe = ds.Tables(0).Rows(0)("Classe")

            If IsDBNull(ds.Tables(0).Rows(0)("OpDestino_Id")) Then
                Me.OpDestinoId = New Nullable(Of Integer)
            Else
                Me.OpDestinoId = CInt(ds.Tables(0).Rows(0)("OpDestino_Id"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("SubOpDestino_Id")) Then
                Me.SubOpDestinoId = New Nullable(Of Integer)
            Else
                Me.SubOpDestinoId = CInt(ds.Tables(0).Rows(0)("SubOpDestino_Id"))
            End If

            If Not IsDBNull(ds.Tables(0).Rows(0)("PrestacaoDeServico")) Then
                Me.PrestacaoDeServico = CType(ds.Tables(0).Rows(0)("PrestacaoDeServico"), ePrestacaoServico)
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("OpAnulacao")) Then
                Me.OpAnulacao = New Nullable(Of Integer)
            Else
                Me.OpAnulacao = CInt(ds.Tables(0).Rows(0)("OpAnulacao"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("SubOpAnulacao")) Then
                Me.SubOpAnulacao = New Nullable(Of Integer)
            Else
                Me.SubOpAnulacao = CInt(ds.Tables(0).Rows(0)("SubOpAnulacao"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("OpContrapartida")) Then
                Me.OpContrapartida = New Nullable(Of Integer)
            Else
                Me.OpContrapartida = CInt(ds.Tables(0).Rows(0)("OpContrapartida"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("SubOpContrapartida")) Then
                Me.SubOpContrapartida = New Nullable(Of Integer)
            Else
                Me.SubOpContrapartida = CInt(ds.Tables(0).Rows(0)("SubOpContrapartida"))
            End If

            Me.CodigoFinalidade = ds.Tables(0).Rows(0)("Finalidade")

            Me.TipoOperacao = CType(ds.Tables(0).Rows(0)("TipoOperacao"), eTipoOperacao)
        End If
    End Sub

    Public Sub New(ByVal entradaSaida As String,
                   ByVal bEmitente_FRT As Boolean,
                   ByVal bTomador_FRT As Boolean,
                   ByVal bEmpresa_NF As Boolean,
                   ByVal bDestino_NF As Boolean)

        Dim db As New AcessaBanco
        Dim sql As String = String.Empty
        sql = " SELECT TOP 1 " & vbCrLf &
              "        OPF.Operacao_Id, OPF.SubOperacoes_Id, OPF.TipoPessoa_Id, OPF.OpDestino_Id, OPF.SubOpDestino_Id,  OPF.TipoOperacao, " & vbCrLf &
              "        OPF.PrestacaoDeServico, OPF.OpAnulacao, OPF.SubOpAnulacao, OPF.OpContrapartida, OPF.SubOpContrapartida," & vbCrLf &
              "        SOP.EntradaSaida, SOP.Classe, isnull(OPF.Finalidade,0) as Finalidade, " & vbCrLf &
              "        OPF.UsuarioInclusao, OPF.UsuarioInclusaoData, OPF.UsuarioAlteracao, OPF.UsuarioAlteracaoData " & vbCrLf &
              "   FROM OperacoesDeFretes OPF " & vbCrLf &
              "  INNER JOIN SubOperacoes SOP " & vbCrLf &
              "     ON SOP.Operacao_Id          = OPF.Operacao_Id " & vbCrLf &
              "    AND SOP.SubOperacoes_Id      = OPF.SubOperacoes_Id  " & vbCrLf &
              "  INNER JOIN VW_OperacoesVigentes OE  " & vbCrLf &
              "     ON OPF.Operacao_Id          = OE.Operacao  " & vbCrLf &
              "     AND OPF.SubOperacoes_Id     = OE.SubOperacao " & vbCrLf

        sql &= "   LEFT JOIN SubOperacoes SOPD " & vbCrLf &
               "     ON SOPD.SubOperacoes_Id    = OPF.SubOpDestino_Id " & vbCrLf &
               "    AND SOPD.Operacao_Id        = OPF.OpDestino_Id " & vbCrLf &
               "  WHERE sop.EntradaSaida        = '" & entradaSaida & "' " & vbCrLf

        sql &= "    AND OE.ativo = 1 " & vbCrLf
        sql &= "    AND OPF.Emitente_FRT        = '" & bEmitente_FRT & "'" & vbCrLf
        sql &= "    AND OPF.Tomador_FRT         = '" & bTomador_FRT & "'" & vbCrLf
        sql &= "    AND OPF.Empresa_NF          = '" & bEmpresa_NF & "'" & vbCrLf
        sql &= "    AND OPF.Destino_NF          = '" & bDestino_NF & "'" & vbCrLf

        Dim ds As DataSet = db.ConsultaDataSet(sql, "OperacoesDeFretes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Me.OperacaoId = ds.Tables(0).Rows(0)("Operacao_Id")
            Me.SubOperacaoId = ds.Tables(0).Rows(0)("SubOperacoes_Id")
            Me.TipoPessoaId = ds.Tables(0).Rows(0)("TipoPessoa_Id")
            Me.Classe = ds.Tables(0).Rows(0)("Classe")

            If IsDBNull(ds.Tables(0).Rows(0)("OpDestino_Id")) Then
                Me.OpDestinoId = New Nullable(Of Integer)
            Else
                Me.OpDestinoId = CInt(ds.Tables(0).Rows(0)("OpDestino_Id"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("SubOpDestino_Id")) Then
                Me.SubOpDestinoId = New Nullable(Of Integer)
            Else
                Me.SubOpDestinoId = CInt(ds.Tables(0).Rows(0)("SubOpDestino_Id"))
            End If

            If Not IsDBNull(ds.Tables(0).Rows(0)("PrestacaoDeServico")) Then
                Me.PrestacaoDeServico = CType(ds.Tables(0).Rows(0)("PrestacaoDeServico"), ePrestacaoServico)
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("OpAnulacao")) Then
                Me.OpAnulacao = New Nullable(Of Integer)
            Else
                Me.OpAnulacao = CInt(ds.Tables(0).Rows(0)("OpAnulacao"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("SubOpAnulacao")) Then
                Me.SubOpAnulacao = New Nullable(Of Integer)
            Else
                Me.SubOpAnulacao = CInt(ds.Tables(0).Rows(0)("SubOpAnulacao"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("OpContrapartida")) Then
                Me.OpContrapartida = New Nullable(Of Integer)
            Else
                Me.OpContrapartida = CInt(ds.Tables(0).Rows(0)("OpContrapartida"))
            End If

            If IsDBNull(ds.Tables(0).Rows(0)("SubOpContrapartida")) Then
                Me.SubOpContrapartida = New Nullable(Of Integer)
            Else
                Me.SubOpContrapartida = CInt(ds.Tables(0).Rows(0)("SubOpContrapartida"))
            End If

            Me.CodigoFinalidade = ds.Tables(0).Rows(0)("Finalidade")

            Me.TipoOperacao = CType(ds.Tables(0).Rows(0)("TipoOperacao"), eTipoOperacao)
        End If
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _OperacaoId As Integer
    Private _SubOperacaoId As Integer
    Private _TipoPessoaId As String
    Private _Classe As String
    Private _EntradaSaida As String
    Private _OpDestinoId As System.Nullable(Of Int32)
    Private _SubOpDestinoId As System.Nullable(Of Int32)
    Private _OpAnulacao As System.Nullable(Of Int32)
    Private _SubOpAnulacao As System.Nullable(Of Int32)
    Private _OpContrapartida As System.Nullable(Of Int32)
    Private _SubOpContrapartida As System.Nullable(Of Int32)
    Private _PrestacaoDeServico As System.Nullable(Of ePrestacaoServico)
    Private _TipoOperacao As eTipoOperacao
    Private _CodigoFinalidade As Integer
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

    Public Property OperacaoId() As Integer
        Get
            Return _OperacaoId
        End Get
        Set(ByVal value As Integer)
            _OperacaoId = value
        End Set
    End Property

    Public Property SubOperacaoId() As Integer
        Get
            Return _SubOperacaoId
        End Get
        Set(ByVal value As Integer)
            _SubOperacaoId = value
        End Set
    End Property

    Public Property OpDestinoId() As System.Nullable(Of Int32)
        Get
            Return _OpDestinoId
        End Get
        Set(ByVal value As System.Nullable(Of Int32))
            _OpDestinoId = value
        End Set
    End Property

    Public Property SubOpDestinoId() As System.Nullable(Of Int32)
        Get
            Return _SubOpDestinoId
        End Get
        Set(ByVal value As System.Nullable(Of Int32))
            _SubOpDestinoId = value
        End Set
    End Property

    Public Property TipoPessoaId() As String
        Get
            Return _TipoPessoaId
        End Get
        Set(ByVal value As String)
            _TipoPessoaId = value
        End Set
    End Property

    Public Property PrestacaoDeServico() As System.Nullable(Of ePrestacaoServico)
        Get
            Return _PrestacaoDeServico
        End Get
        Set(ByVal value As System.Nullable(Of ePrestacaoServico))
            _PrestacaoDeServico = value
        End Set
    End Property

    Public Property TipoOperacao() As eTipoOperacao
        Get
            Return _TipoOperacao
        End Get
        Set(ByVal value As eTipoOperacao)
            _TipoOperacao = value
        End Set
    End Property

    Public Property OpAnulacao() As System.Nullable(Of Int32)
        Get
            Return _OpAnulacao
        End Get
        Set(ByVal value As System.Nullable(Of Int32))
            _OpAnulacao = value
        End Set
    End Property

    Public Property SubOpAnulacao() As System.Nullable(Of Int32)
        Get
            Return _SubOpAnulacao
        End Get
        Set(ByVal value As System.Nullable(Of Int32))
            _SubOpAnulacao = value
        End Set
    End Property

    Public Property OpContrapartida() As System.Nullable(Of Int32)
        Get
            Return _OpContrapartida
        End Get
        Set(ByVal value As System.Nullable(Of Int32))
            _OpContrapartida = value
        End Set
    End Property

    Public Property SubOpContrapartida() As System.Nullable(Of Int32)
        Get
            Return _SubOpContrapartida
        End Get
        Set(ByVal value As System.Nullable(Of Int32))
            _SubOpContrapartida = value
        End Set
    End Property

    Public Property Classe() As String
        Get
            Return _Classe
        End Get
        Set(ByVal value As String)
            _Classe = value
        End Set
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

    Public Property CodigoFinalidade As Integer
        Get
            Return _CodigoFinalidade
        End Get
        Set(value As Integer)
            _CodigoFinalidade = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Try
            If IUD = Nothing Then
                Return True
            End If
            Dim db As New AcessaBanco
            Dim Sqls As New ArrayList
            Sqls.Clear()
            SalvarSql(Sqls)
            Return db.GravaBanco(Sqls)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Dim Destino As String = ""
                If Me.OpDestinoId Is Nothing Then
                    Destino = "null"
                Else
                    Destino = "'" & Me.OpDestinoId & "'"
                End If

                Dim DestinoSub As String = ""
                If Me.SubOpDestinoId Is Nothing Then
                    DestinoSub = "null"
                Else
                    DestinoSub = "'" & Me.SubOpDestinoId & "'"
                End If

                Dim Anulacao As String = ""
                If Me.OpAnulacao Is Nothing Then
                    Anulacao = "null"
                Else
                    Anulacao = "'" & Me.OpAnulacao & "'"
                End If

                Dim AnulacaoSub As String = ""
                If Me.SubOpAnulacao Is Nothing Then
                    AnulacaoSub = "null"
                Else
                    AnulacaoSub = "'" & Me.SubOpAnulacao & "'"
                End If

                Dim PrestacaoServico As String = ""
                If Me.PrestacaoDeServico Is Nothing Then
                    PrestacaoServico = "null"
                Else
                    PrestacaoServico = "'" & Me.PrestacaoDeServico & "'"
                End If

                Sql = " INSERT INTO OperacoesDeFretes " & vbCrLf &
                      "        (Operacao_Id, SubOperacoes_Id, TipoPessoa_Id, OpDestino_Id, SubOpDestino_Id, OpAnulacao, SubOpAnulacao, " & vbCrLf &
                      "         OpContrapartida, SubOpContrapartida, PrestacaoDeServico, TipoOperacao, Finalidade, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf &
                      " VALUES (" & Me.OperacaoId & "," & vbCrLf &
                      "         " & Me.SubOperacaoId & ", " & vbCrLf &
                      "         '" & Me.TipoPessoaId & "', " & vbCrLf &
                      "          " & Destino & ", " & vbCrLf &
                      "          " & DestinoSub & ", " & vbCrLf &
                      "          " & Anulacao & ", " & vbCrLf &
                      "          " & AnulacaoSub & ", " & vbCrLf &
                      "          " & IIf(Me.OpContrapartida Is Nothing, "null", ("'" & Me.OpContrapartida & "'")) & ", " & vbCrLf &
                      "          " & IIf(Me.SubOpContrapartida Is Nothing, "null", ("'" & Me.SubOpContrapartida & "'")) & ", " & vbCrLf &
                      "          " & PrestacaoServico & ", " & CInt(Me.TipoOperacao) & "," & vbCrLf &
                      "          " & Me.CodigoFinalidade & ", " & vbCrLf &
                      "          '" & UsuarioServidor.NomeUsuario & "', " & vbCrLf &
                      "          '" & CDate(DateTime.Now).ToString("yyyy/MM/dd hh:mm") & "')"
                Sqls.Add(Sql)
            Case "U"
                Dim Destino As String = ""
                If Me.OpDestinoId Is Nothing Then
                    Destino = "null"
                Else
                    Destino = "'" & Me.OpDestinoId & "'"
                End If

                Dim DestinoSub As String = ""
                If Me.SubOpDestinoId Is Nothing Then
                    DestinoSub = "null"
                Else
                    DestinoSub = "'" & Me.SubOpDestinoId & "'"
                End If

                Dim Anulacao As String = ""
                If Me.OpAnulacao Is Nothing Then
                    Anulacao = "null"
                Else
                    Anulacao = "'" & Me.OpAnulacao & "'"
                End If

                Dim AnulacaoSub As String = ""
                If Me.SubOpAnulacao Is Nothing Then
                    AnulacaoSub = "null"
                Else
                    AnulacaoSub = "'" & Me.SubOpAnulacao & "'"
                End If

                Dim Contrapartida As String = ""
                If Me.OpContrapartida Is Nothing Then
                    Contrapartida = "null"
                Else
                    Contrapartida = "'" & Me.OpContrapartida & "'"
                End If

                Dim ContrapartidaSub As String = ""
                If Me.SubOpContrapartida Is Nothing Then
                    ContrapartidaSub = "null"
                Else
                    ContrapartidaSub = "'" & Me.SubOpContrapartida & "'"
                End If

                Dim PrestacaoServico As String = ""
                If Me.PrestacaoDeServico Is Nothing Then
                    PrestacaoServico = "null"
                Else
                    PrestacaoServico = "'" & Me.PrestacaoDeServico & "'"
                End If

                Sql = " UPDATE OperacoesDeFretes " & vbCrLf &
                      "    SET TipoPessoa_Id       ='" & Me.TipoPessoaId & "'," & vbCrLf &
                      "        OpDestino_Id        = " & Destino & ", " & vbCrLf &
                      "        SubOpDestino_Id     = " & DestinoSub & ", " & vbCrLf &
                      "        OpAnulacao          = " & Anulacao & ", " & vbCrLf &
                      "        SubOpAnulacao       = " & AnulacaoSub & ", " & vbCrLf &
                      "        OpContrapartida     = " & Contrapartida & ", " & vbCrLf &
                      "        SubOpContrapartida  = " & ContrapartidaSub & ", " & vbCrLf &
                      "        PrestacaoDeServico  = " & PrestacaoServico & ", " & vbCrLf &
                      "        TipoOperacao        = " & CInt(Me.TipoOperacao) & ", " & vbCrLf &
                      "        Finalidade          = " & Me.CodigoFinalidade.ToSqlNULL & ", " & vbCrLf &
                      "        UsuarioInclusao     ='" & UsuarioServidor.NomeUsuario & "' ," & vbCrLf &
                      "        UsuarioInclusaoData ='" & CDate(DateTime.Now).ToString("yyyy/MM/dd hh:mm") & "'" & vbCrLf &
                      "  WHERE Operacao_Id = " & Me.OperacaoId & vbCrLf &
                      "    AND SubOperacoes_Id = " & Me.SubOperacaoId
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE OperacoesDeFretes " & vbCrLf &
                      "  WHERE Operacao_Id = " & Me.OperacaoId & vbCrLf &
                      "    AND SubOperacoes_Id = " & Me.SubOperacaoId & vbCrLf &
                      "    AND TipoPessoa_Id = '" & Me.TipoPessoaId & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class