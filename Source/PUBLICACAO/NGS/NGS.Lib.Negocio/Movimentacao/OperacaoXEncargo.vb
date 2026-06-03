Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListOperacaoXEncargo
    Inherits List(Of OperacaoXEncargo)

#Region "Fields"
    Private _TotalLiquido As Decimal
    Private _EncProduto As OperacaoXEncargo
#End Region

#Region "Property"
    Public Property TotalLiquido() As Decimal
        Get
            Return _TotalLiquido
        End Get
        Set(ByVal value As Decimal)
            _TotalLiquido = value
        End Set
    End Property

    Public Property EncProduto() As OperacaoXEncargo
        Get
            Return _EncProduto
        End Get
        Set(ByVal value As OperacaoXEncargo)
            _EncProduto = value
        End Set
    End Property
#End Region

#Region "Construtores"
    Public Sub New(Optional ByVal CriaComEncargosObrigatorios As Boolean = False)
        If CriaComEncargosObrigatorios Then
            Dim EncProduto As New Encargo("PRODUTO")
            Dim EncLiquido As New Encargo("LIQUIDO")

            Dim OxE_Produto As New OperacaoXEncargo(EncProduto)
            Dim OxE_Liquido As New OperacaoXEncargo(EncLiquido)

            Me.EncProduto = OxE_Produto
            Me.Add(Me.EncProduto)
            Me.Add(OxE_Liquido)
        End If
    End Sub

    Public Sub New(ByVal GrupoProduto As String, ByVal Produto As String, ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal EstadoOrigem As String, ByVal EstadoDestino As String, Optional Exato As Boolean = False)
        Me.Selecionar(GrupoProduto, Produto, Operacao, SubOperacao, EstadoOrigem, EstadoDestino, Exato)
    End Sub

    Public Sub New(ByVal GrupoProduto As String, ByVal Produto As String, ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal EstadoOrigem As String, ByVal EstadoDestino As String, ByVal Movimento As String)
        Me.SelecionarEncargos(GrupoProduto, Produto, Operacao, SubOperacao, EstadoOrigem, EstadoDestino, Movimento)
    End Sub
#End Region

#Region "Methods"
    Public Function Selecionar(ByVal GrupoProduto As String, ByVal Produto As String, ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal EstadoOrigem As String, ByVal EstadoDestino As String, Optional Exato As Boolean = False) As Boolean
        Dim objBanco As New AcessaBanco()
        Dim uf As String
        Dim strSQL As String
        Dim strSQLProduto As String = Produto
        If EstadoOrigem = EstadoDestino Then
            uf = "'" & EstadoOrigem & "'"
        Else
            uf = "(SELECT Regiao " & _
                 "  FROM Estados " & _
                 " WHERE Estado_Id = '" & EstadoDestino & "') "
        End If

        Try
            strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                     "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                     "       OXE.Sinal, OXE.PercentualBase, OXE.PercentualLimite, OXE.Observacao, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, " & vbCrLf & _
                     "       OXE.UsuarioInclusao, OXE.UsuarioInclusaoData, OXE.UsuarioAlteracao, OXE.UsuarioAlteracaoData, " & vbCrLf & _
                     "		 CASE " & vbCrLf & _
                     "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                     "				THEN CASE " & vbCrLf & _
                     "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                     "							THEN 3 " & vbCrLf & _
                     "							ELSE 53 " & vbCrLf & _
                     "						END " & vbCrLf & _
                     "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                     "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                     "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                     "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                     "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                     "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                     "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                     " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                     "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                     "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                     "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                     "   AND OXE.EstadoOrigem_Id  = '" & EstadoOrigem & "' " & vbCrLf & _
                     "   AND OXE.EstadoDestino_Id = '" & EstadoDestino & "' " & vbCrLf & _
                     " ORDER BY case When OXE.Encargo_Id = 'PRODUTO' then 1 when OXE.Encargo_Id = 'LIQUIDO' then 3 else 2 end, OXE.Encargo_Id" & vbCrLf
            Dim dsEncargos As New DataSet
            dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")

            If dsEncargos.Tables(0).Rows.Count = 0 Then
                If Exato Then Return False
                strSQLProduto = ""
                strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                         "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                         "       OXE.Sinal, OXE.PercentualBase, OXE.PercentualLimite, OXE.Observacao, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, " & vbCrLf & _
                          "      OXE.UsuarioInclusao, OXE.UsuarioInclusaoData, OXE.UsuarioAlteracao, OXE.UsuarioAlteracaoData, " & vbCrLf & _
                         "		 CASE " & vbCrLf & _
                         "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                         "				THEN CASE " & vbCrLf & _
                         "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                         "							THEN 3 " & vbCrLf & _
                         "							ELSE 53 " & vbCrLf & _
                         "						END " & vbCrLf & _
                         "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                         "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                         "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                         "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                         "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                         "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                         "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                         " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                         "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                         "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.EstadoOrigem_Id  = '" & EstadoOrigem & "' " & vbCrLf & _
                         "   AND OXE.EstadoDestino_Id = '" & EstadoDestino & "' " & vbCrLf & _
                         " ORDER BY case When OXE.Encargo_Id = 'PRODUTO' then 1 when OXE.Encargo_Id = 'LIQUIDO' then 3 else 2 end, OXE.Encargo_Id" & vbCrLf
                dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")
            End If

            If dsEncargos Is Nothing OrElse dsEncargos.Tables(0).Rows.Count = 0 Then
                strSQLProduto = Produto
                strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                         "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta, " & vbCrLf & _
                         "       OXE.Sinal, OXE.PercentualBase, OXE.PercentualLimite, OXE.Observacao, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, " & vbCrLf & _
                          "      OXE.UsuarioInclusao, OXE.UsuarioInclusaoData, OXE.UsuarioAlteracao, OXE.UsuarioAlteracaoData, " & vbCrLf & _
                         "		 CASE " & vbCrLf & _
                         "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                         "				THEN CASE " & vbCrLf & _
                         "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                         "							THEN 3 " & vbCrLf & _
                         "							ELSE 53 " & vbCrLf & _
                         "						END " & vbCrLf & _
                         "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                         "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                         "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                         "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                         "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                         "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                         "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                         " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                         "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                         "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.EstadoOrigem_Id  ='" & EstadoOrigem & "' " & vbCrLf & _
                         "   AND OXE.EstadoDestino_Id = " & uf & " " & vbCrLf & _
                         " ORDER BY case When OXE.Encargo_Id = 'PRODUTO' then 1 when OXE.Encargo_Id = 'LIQUIDO' then 3 else 2 end, OXE.Encargo_Id" & vbCrLf
                dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")
            End If

            If dsEncargos.Tables(0).Rows.Count = 0 Then
                strSQLProduto = ""
                strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                         "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                         "       OXE.Sinal, OXE.PercentualBase, OXE.PercentualLimite, OXE.Observacao, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, " & vbCrLf & _
                          "      OXE.UsuarioInclusao, OXE.UsuarioInclusaoData, OXE.UsuarioAlteracao, OXE.UsuarioAlteracaoData, " & vbCrLf & _
                         "		 CASE " & vbCrLf & _
                         "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                         "				THEN CASE " & vbCrLf & _
                         "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                         "							THEN 3 " & vbCrLf & _
                         "							ELSE 53 " & vbCrLf & _
                         "						END " & vbCrLf & _
                         "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                         "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                         "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                         "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                         "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                         "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                         "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                         " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                         "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                         "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.EstadoOrigem_Id  ='" & EstadoOrigem & "' " & vbCrLf & _
                         "   AND OXE.EstadoDestino_Id = " & uf & " " & vbCrLf & _
                         " ORDER BY case When OXE.Encargo_Id = 'PRODUTO' then 1 when OXE.Encargo_Id = 'LIQUIDO' then 3 else 2 end, OXE.Encargo_Id" & vbCrLf
                dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")
            End If

            If dsEncargos.Tables(0).Rows.Count > 0 Then
                For Each drEncargo As DataRow In dsEncargos.Tables(0).Rows
                    Dim objOperacoesXEncargos As New OperacaoXEncargo()

                    objOperacoesXEncargos.Selecionado = True
                    objOperacoesXEncargos.IUD = "U"
                    objOperacoesXEncargos.CodigoGrupoProduto = GrupoProduto
                    objOperacoesXEncargos.CodigoProduto = drEncargo("Produto_Id").ToString()
                    objOperacoesXEncargos.CodigoOperacao = Operacao
                    objOperacoesXEncargos.CodigoSubOperacao = SubOperacao
                    objOperacoesXEncargos.EstadoOrigem = drEncargo("EstadoOrigem_Id").ToString()
                    objOperacoesXEncargos.EstadoDestino = drEncargo("EstadoDestino_Id").ToString()
                    objOperacoesXEncargos.CodigoEncargo = drEncargo("Encargo_Id").ToString()

                    objOperacoesXEncargos.SituacaoTributariaICMS = Convert.ToInt32(drEncargo("SituacaoTributaria"))
                    objOperacoesXEncargos.SituacaoTributariaIPI = drEncargo("SituacaoTributariaIPI")
                    objOperacoesXEncargos.SituacaoTributariaPISCOFINS = drEncargo("SituacaoTributariaPISCOFINS")
                    objOperacoesXEncargos.CodigoObservacaoPISCOFINS = drEncargo("ObservacaoPISCOFINS")
                    objOperacoesXEncargos.CodigoObservacaoGeral = Convert.ToInt32(drEncargo("Observacao"))

                    objOperacoesXEncargos.CodigoGrupoFiscal = Convert.ToInt32(drEncargo("GrupoFiscal"))
                    objOperacoesXEncargos.CodigoFiscal = Convert.ToInt32(drEncargo("CodigoFiscal"))
                    objOperacoesXEncargos.CodigoDebitaConta = drEncargo("DebitaConta").ToString()
                    objOperacoesXEncargos.CodigoCreditaConta = drEncargo("CreditaConta").ToString()
                    objOperacoesXEncargos.Sinal = drEncargo("Sinal")
                    objOperacoesXEncargos.PercentualBase = drEncargo("PercentualBase").ToString()
                    objOperacoesXEncargos.Aliquota = drEncargo("Aliquota").ToString()
                    objOperacoesXEncargos.AliquotaExibicao = drEncargo("AliquotaExibicao").ToString()
                    objOperacoesXEncargos.PercentualLimite = drEncargo("PercentualLimite").ToString()
                    objOperacoesXEncargos.UsuarioAlteracao = UsuarioServidor.Usuario.Usuario_Id
                    objOperacoesXEncargos.UsuarioAlteracaoData = DateTime.Now

                    If drEncargo("Encargo_Id") = "PRODUTO" Then EncProduto = objOperacoesXEncargos
                    Me.Add(objOperacoesXEncargos)
                Next
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Function SelecionarEncargos(ByVal GrupoProduto As String, ByVal Produto As String, ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal EstadoOrigem As String, ByVal EstadoDestino As String, ByVal Movimento As String) As Boolean
        Dim sql As String = ""

        sql &= " Select " & vbCrLf & _
               "	  case" & vbCrLf & _
               "		when OxE.Produto_Id <> '' and OxE.EstadoDestino_Id = '" & EstadoDestino & "'" & vbCrLf & _
               "		  then 1" & vbCrLf & _
               "		when OxE.Produto_Id =  '' and OxE.EstadoDestino_Id = '" & EstadoDestino & "'" & vbCrLf & _
               "		  then 2" & vbCrLf & _
               "	  end Prioridade, " & vbCrLf & _
               "      case " & vbCrLf & _
               "         When OxE.Encargo_Id = 'PRODUTO' " & vbCrLf & _
               "           Then 1 " & vbCrLf & _
               "         When OxE.Encargo_Id = 'LIQUIDO' or OxE.Encargo_Id = 'LIQUIDOAPAGAR' " & vbCrLf & _
               "           Then 3 " & vbCrLf & _
               "         Else 2 " & vbCrLf & _
               "      end ordem, " & vbCrLf & _
               "      OxE.GrupoProduto_Id, OxE.Produto_Id, OxE.Operacao_Id, OxE.SubOperacao_Id, OxE.EstadoOrigem_Id, OxE.EstadoDestino_Id, OxE.Encargo_Id, isnull(OxE.SituacaoTributaria,999) AS SituacaoTributaria, " & vbCrLf & _
               "      OxE.GrupoFiscal, OxE.CodigoFiscal, OxE.DebitaConta, OxE.CreditaConta, OxE.Sinal, OxE.PercentualBase," & vbCrLf & _
               "      isnull(encXtx.Percentual, OxE.Aliquota) AS Aliquota, " & vbCrLf & _
               "      encXtx.Percentual, encXtx.SimplesNacional, OxE.AliquotaExibicao, " & vbCrLf & _
               "      OxE.PercentualLimite, OT.Descricao as ObservacaoFiscal, " & vbCrLf & _
               "      case when isnull(prd.IPI,0) > 0 then 1 else 0 end Presumido," & vbCrLf & _
               "	  CASE " & vbCrLf & _
               "		WHEN OxE.situacaotributariaipi is null " & vbCrLf & _
               "			THEN CASE " & vbCrLf & _
               "					WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
               "						THEN 3 " & vbCrLf & _
               "						ELSE 53 " & vbCrLf & _
               "					END " & vbCrLf & _
               "			ELSE OxE.SituacaoTributariaIPI " & vbCrLf & _
               "		END AS SituacaoTributariaIPI, " & vbCrLf & _
               "     isnull(OxE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OxE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
               "   into #Temp" & vbCrLf & _
               "   from OperacoesXEncargos OxE" & vbCrLf & _
               "   INNER JOIN SubOperacoes SO " & vbCrLf & _
               "			ON  SO.Operacao_Id     = OxE.Operacao_Id " & vbCrLf & _
               "			AND SO.SubOperacoes_Id = OxE.SubOperacao_Id " & vbCrLf & _
               "   left join (SELECT ext.Encargo_Id as Encargo, ext.Percentual, ext.Estado_Id, isnull(ext.SimplesNacional,0) as SimplesNacional " & vbCrLf & _
               "				FROM EncargosXTaxas ext " & vbCrLf & _
               "				WHERE Data_Id = (SELECT MAX(Data_Id) " & vbCrLf & _
               "								   FROM EncargosXTaxas " & vbCrLf & _
               "								  WHERE EncargosXTaxas.Encargo_Id = ext.Encargo_Id" & vbCrLf & _
               "									AND EncargosXTaxas.Estado_Id  = ext.Estado_Id" & vbCrLf & _
               "									AND Data_id <= '" & CDate(Movimento).ToString("yyyy-MM-dd") & "')" & vbCrLf & _
               "              ) encXtx " & vbCrLf & _
               "     on encXtx.Encargo    = OxE.Encargo_id " & vbCrLf & _
               "    AND encXtx.Estado_Id  = OxE.EstadoOrigem_id " & vbCrLf & _
               "   left join ObservacoesTributarias OT" & vbCrLf & _
               "     on OxE.Observacao = OT.Codigo_Id" & vbCrLf & _
               "  Inner Join Encargos En" & vbCrLf & _
               "     on En.Encargo_id =  OxE.Encargo_Id" & vbCrLf & _
               "   LEFT JOIN Produtos Prd" & vbCrLf & _
               "     on OxE.Encargo_Id = 'IPI'" & vbCrLf & _
               "    and Prd.Produto_Id = '" & Produto & "'" & vbCrLf & _
               "  where OxE.Operacao_Id      = " & Operacao & vbCrLf & _
               "    and OxE.SubOperacao_Id   = " & SubOperacao & vbCrLf & _
               "    and OxE.GrupoProduto_Id  = " & GrupoProduto & vbCrLf & _
               "    and (OxE.produto_id      = '' or OxE.Produto_Id = '" & Produto & "')" & vbCrLf & _
               "    and OxE.EstadoOrigem_Id  ='" & EstadoOrigem & "'" & vbCrLf & _
               "    and OxE.EstadoDestino_Id ='" & EstadoDestino & "';" & vbCrLf & _
               "                                                        " & vbCrLf & _
               " select T.GrupoProduto_Id, T.Produto_id, T.Operacao_Id, T.SubOperacao_Id, T.EstadoOrigem_Id, T.EstadoDestino_Id, " & vbCrLf & _
               "        T.Encargo_Id, T.SituacaoTributaria, T.GrupoFiscal, T.CodigoFiscal, T.DebitaConta, T.CreditaConta, " & vbCrLf & _
               "        T.Sinal, T.PercentualBase, T.Aliquota, isnull(T.AliquotaExibicao,0) as AliquotaExibicao,  T.PercentualLimite," & vbCrLf & _
               "        isnull(T.ObservacaoFiscal,'') as ObservacaoFiscal " & vbCrLf & _
               "        T.SituacaoTributariaIPI, T.SituacaoTributariaPISCOFINS, T.ObservacaoPISCOFINS" & vbCrLf & _
               "   Into #Pre " & vbCrLf & _
               "   from #Temp T " & vbCrLf & _
               "    and NxE.Encargo_Id      = T.Encargo_Id" & vbCrLf & _
               "  where T.prioridade  = (select min(prioridade) from #Temp) " & vbCrLf & _
               "  order by ordem, case When Encargo_Id = 'PRODUTO' then 1 when Encargo_Id = 'LIQUIDO' then 3 else 2 end, Encargo_Id;" & vbCrLf

        Dim objBanco As New AcessaBanco()
        Dim dsEncargos As DataSet = objBanco.ConsultaDataSet(sql, "OperacoesXEncargos")

        If dsEncargos Is Nothing OrElse dsEncargos.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            For Each drEncargo As DataRow In dsEncargos.Tables(0).Rows
                Dim objOperacoesXEncargos As New OperacaoXEncargo()

                objOperacoesXEncargos.Selecionado = True
                objOperacoesXEncargos.CodigoGrupoProduto = drEncargo("GrupoProduto_Id")
                objOperacoesXEncargos.CodigoProduto = drEncargo("Produto_id")
                objOperacoesXEncargos.CodigoOperacao = drEncargo("Operacao_Id")
                objOperacoesXEncargos.CodigoSubOperacao = drEncargo("SubOperacao_Id")
                objOperacoesXEncargos.EstadoOrigem = drEncargo("EstadoOrigem_Id")
                objOperacoesXEncargos.EstadoDestino = drEncargo("EstadoDestino_Id")
                objOperacoesXEncargos.CodigoEncargo = drEncargo("Encargo_Id")

                objOperacoesXEncargos.SituacaoTributariaICMS = Convert.ToInt32(drEncargo("SituacaoTributaria"))
                objOperacoesXEncargos.SituacaoTributariaIPI = drEncargo("SituacaoTributariaIPI")
                objOperacoesXEncargos.SituacaoTributariaPISCOFINS = drEncargo("SituacaoTributariaPISCOFINS")
                objOperacoesXEncargos.CodigoObservacaoPISCOFINS = drEncargo("ObservacaoPISCOFINS")
                objOperacoesXEncargos.CodigoObservacaoGeral = Convert.ToInt32(drEncargo("ObservacaoFiscal"))

                objOperacoesXEncargos.CodigoGrupoFiscal = Convert.ToInt32(drEncargo("GrupoFiscal"))
                objOperacoesXEncargos.CodigoFiscal = Convert.ToInt32(drEncargo("CodigoFiscal"))
                objOperacoesXEncargos.CodigoDebitaConta = drEncargo("DebitaConta")
                objOperacoesXEncargos.CodigoCreditaConta = drEncargo("CreditaConta")
                objOperacoesXEncargos.Sinal = drEncargo("Sinal").ToString()
                objOperacoesXEncargos.PercentualBase = drEncargo("PercentualBase")
                objOperacoesXEncargos.Aliquota = drEncargo("Aliquota")
                objOperacoesXEncargos.AliquotaExibicao = drEncargo("AliquotaExibicao")
                objOperacoesXEncargos.PercentualLimite = drEncargo("PercentualLimite")

                If drEncargo("Encargo_Id") = "PRODUTO" Then EncProduto = objOperacoesXEncargos

                Me.Add(objOperacoesXEncargos)
            Next

            Return True
        End If
    End Function

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Sqls.Count = 0 Then Return True

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As OperacaoXEncargo In Me
            If Me.EncProduto.IUD = "D" Or Me.EncProduto.IUD = "I" Then item.IUD = Me.EncProduto.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function ExisteLancamentos() As Boolean
        Dim Banco As New AcessaBanco
        Dim SQL As String
        SQL = "select case" & vbCrLf & _
              "          when Exists( select 1" & vbCrLf & _
              "                         from notasfiscaisxencargos" & vbCrLf & _
              "                        Where Grupo           = " & EncProduto.CodigoGrupoProduto & vbCrLf & _
              "                          and Produto         ='" & EncProduto.CodigoProduto & "'" & vbCrLf & _
              "                          and Operacao        = " & EncProduto.CodigoOperacao & vbCrLf & _
              "                          and SubOperacao     = " & EncProduto.CodigoSubOperacao & vbCrLf & _
              "                          and EstadoOrigem    ='" & EncProduto.EstadoOrigem & "'" & vbCrLf & _
              "                          and EstadoDestino   ='" & EncProduto.EstadoDestino & "'" & vbCrLf & _
              "                      )" & vbCrLf & _
              "            then 1" & vbCrLf & _
              "            else 0" & vbCrLf & _
              "       end" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(SQL, "Existe")

        If ds.Tables(0).Rows(0)(0) = 0 Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

End Class

<Serializable()> _
Public Class OperacaoXEncargo

#Region "Fields"
    Private _IUD As String = ""
    Private _Selecionado As Boolean = True
    Private _CodigoGrupoProduto As String = ""
    Private _CodigoProduto As String = ""

    Private _CodigoOperacao As Integer
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao

    Private _EstadoOrigem As String = ""
    Private _EstadoDestino As String = ""
    Private _CodigoEncargo As String = ""
    Private _Encargo As Encargo

    Private _SituacaoTributariaICMS As Integer
    Private _SituacaoTributariaIPI As Integer
    Private _SituacaoTributariaPISCOFINS As Integer
    Private _CodigoObservacaoPISCOFINS As Integer
    Private _CodigoObservacaoGeral As Integer

    Private _CodigoGrupoFiscal As Integer
    Private _CodigoFiscal As Integer
    Private _DescricaoFiscal As String

    Private _CodigoDebitaConta As String
    Private _DebitaConta As PlanoDeConta
    Private _CodigoCreditaConta As String
    Private _CreditaConta As PlanoDeConta

    Private _Calculado As String
    Private _Sinal As String
    Private _PercentualBase As Decimal
    Private _Aliquota As Decimal
    Private _AliquotaExibicao As Decimal
    Private _PercentualLimite As Decimal
    Private _UsuarioInclusao As String = ""
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _UsuarioAlteracaoData As DateTime
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

    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property

    Public Property CodigoGrupoProduto() As String
        Get
            Return _CodigoGrupoProduto
        End Get
        Set(ByVal value As String)
            _CodigoGrupoProduto = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
            _SubOperacao = Nothing
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
            _SubOperacao = Nothing
        End Set
    End Property

    Public ReadOnly Property SubOperacao() As SubOperacao
        Get
            If _SubOperacao Is Nothing And _CodigoSubOperacao > 0 Then _SubOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)

            Return _SubOperacao
        End Get
    End Property

    Public Property EstadoOrigem() As String
        Get
            Return _EstadoOrigem
        End Get
        Set(ByVal value As String)
            _EstadoOrigem = value
        End Set
    End Property

    Public Property EstadoDestino() As String
        Get
            Return _EstadoDestino
        End Get
        Set(ByVal value As String)
            _EstadoDestino = value
        End Set
    End Property

    Public Property CodigoEncargo() As String
        Get
            Return _CodigoEncargo
        End Get
        Set(ByVal value As String)
            _CodigoEncargo = value
            _Encargo = Nothing
        End Set
    End Property

    Public Property Encargo() As Encargo
        Get
            If _Encargo Is Nothing And _CodigoEncargo.Length > 0 Then _Encargo = New Encargo(_CodigoEncargo)
            Return _Encargo
        End Get
        Set(ByVal value As Encargo)
            _Encargo = value
        End Set
    End Property

    Public Property SituacaoTributariaICMS() As Integer
        Get
            Return _SituacaoTributariaICMS
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributariaICMS = value
        End Set
    End Property

    Public Property SituacaoTributariaIPI() As Integer
        Get
            Return _SituacaoTributariaIPI
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributariaIPI = value
        End Set
    End Property

    Public Property SituacaoTributariaPISCOFINS() As Integer
        Get
            Return _SituacaoTributariaPISCOFINS
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributariaPISCOFINS = value
        End Set
    End Property

    Public Property CodigoObservacaoPISCOFINS() As Integer
        Get
            Return _CodigoObservacaoPISCOFINS
        End Get
        Set(ByVal value As Integer)
            _CodigoObservacaoPISCOFINS = value
        End Set
    End Property

    Public ReadOnly Property ValorPeso() As eValorPeso
        Get
            Return Encargo.ValorOuPeso
        End Get
    End Property

    Public Property CodigoGrupoFiscal() As Integer
        Get
            Return _CodigoGrupoFiscal
        End Get
        Set(ByVal value As Integer)
            _CodigoGrupoFiscal = value
        End Set
    End Property

    Public Property CodigoFiscal() As Integer
        Get
            Return _CodigoFiscal
        End Get
        Set(ByVal value As Integer)
            _CodigoFiscal = value
        End Set
    End Property

    Public Property CodigoDebitaConta() As String
        Get
            Return _CodigoDebitaConta
        End Get
        Set(ByVal value As String)
            _CodigoDebitaConta = value
            _DebitaConta = Nothing
        End Set
    End Property

    Public Property DebitaConta() As PlanoDeConta
        Get
            If _DebitaConta Is Nothing And _CodigoDebitaConta.Length > 0 Then _DebitaConta = New PlanoDeConta("", 0, _CodigoDebitaConta)
            Return _DebitaConta
        End Get
        Set(ByVal value As PlanoDeConta)
            _DebitaConta = value
        End Set
    End Property

    Public Property CodigoCreditaConta() As String
        Get
            Return _CodigoCreditaConta
        End Get
        Set(ByVal value As String)
            _CodigoCreditaConta = value
            _CreditaConta = Nothing
        End Set
    End Property

    Public Property CreditaConta() As PlanoDeConta
        Get
            If _CreditaConta Is Nothing And _CodigoCreditaConta.Length > 0 Then _CreditaConta = New PlanoDeConta("", 0, _CodigoCreditaConta)
            Return _CreditaConta
        End Get
        Set(ByVal value As PlanoDeConta)
            _CreditaConta = value
        End Set
    End Property

    Public Property Sinal() As String
        Get
            Return _Sinal
        End Get
        Set(ByVal value As String)
            _Sinal = value
        End Set
    End Property

    Public Property PercentualBase() As Decimal
        Get
            Return _PercentualBase
        End Get
        Set(ByVal value As Decimal)
            _PercentualBase = value
        End Set
    End Property

    Public Property Aliquota() As Decimal
        Get
            Return _Aliquota
        End Get
        Set(ByVal value As Decimal)
            _Aliquota = value
        End Set
    End Property

    Public Property AliquotaExibicao() As Decimal
        Get
            Return _AliquotaExibicao
        End Get
        Set(ByVal value As Decimal)
            _AliquotaExibicao = value
        End Set
    End Property

    Public Property PercentualLimite() As Decimal
        Get
            Return _PercentualLimite
        End Get
        Set(ByVal value As Decimal)
            _PercentualLimite = value
        End Set
    End Property

    Public Property CodigoObservacaoGeral() As Integer
        Get
            Return _CodigoObservacaoGeral
        End Get
        Set(ByVal value As Integer)
            _CodigoObservacaoGeral = value
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

    Public Property UsuarioInclusaoData() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
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

    Public Property UsuarioAlteracaoData() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal pEncargo As Encargo)

        Me.Selecionado = True
        Me.CodigoEncargo = pEncargo.Codigo
        Me.CodigoDebitaConta = pEncargo.ContaDebito
        Me.CodigoCreditaConta = pEncargo.ContaCredito
        Me.PercentualBase = pEncargo.BaseCalculo
        Me.Aliquota = pEncargo.Aliquota
        Me.AliquotaExibicao = pEncargo.Aliquota

    End Sub

    Public Sub New(ByVal GrupoProduto As String, ByVal Produto As String, ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal EstadoOrigem As String, ByVal EstadoDestino As String, ByVal Encargo As String)

        Selecionar(GrupoProduto, Produto, Operacao, SubOperacao, EstadoOrigem, EstadoDestino, Encargo)

    End Sub

    Public Sub New(ByRef EncargoNF As NotaFiscalXItemXEncargo)

        Selecionar(EncargoNF.CodigoGrupoProduto, EncargoNF.CodigoProduto, EncargoNF.CodigoOperacao, EncargoNF.CodigoSubOperacao, EncargoNF.EstadoOrigem, EncargoNF.EstadoDestino, EncargoNF.Codigo)

    End Sub

#End Region

#Region "Methods"

    Public Function Selecionar(ByVal GrupoProduto As String, ByVal Produto As String, ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal EstadoOrigem As String, ByVal EstadoDestino As String, ByVal Encargo As String) As Boolean
        Dim objBanco As New AcessaBanco()
        Dim uf As String
        Dim strSQL As String
        Dim strSQLProduto As String = Produto

        If EstadoOrigem = EstadoDestino Then
            uf = "'" & EstadoOrigem & "'"
        Else
            uf = "(SELECT Regiao " & _
                 "   FROM Estados " & _
                 "  WHERE Estado_Id = '" & EstadoDestino & "') "
        End If

        Try
            strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                     "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                     "       OXE.Sinal, OXE.PercentualBase, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao,  OXE.PercentualLimite, OXE.Observacao, " & vbCrLf & _
                     "		 CASE " & vbCrLf & _
                     "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                     "				THEN CASE " & vbCrLf & _
                     "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                     "							THEN 3 " & vbCrLf & _
                     "							ELSE 53 " & vbCrLf & _
                     "						END " & vbCrLf & _
                     "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                     "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                     "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                     "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                     "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                     "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                     "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                     " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                     "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                     "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                     "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                     "   AND OXE.EstadoOrigem_Id  ='" & EstadoOrigem & "' " & vbCrLf & _
                     "   AND OXE.EstadoDestino_Id ='" & EstadoDestino & "' " & vbCrLf & _
                     "   AND OXE.Encargo_Id       ='" & Encargo & "' " & vbCrLf & _
                     " ORDER BY OXE.Encargo_Id"

            Dim dsEncargos As DataSet = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")

            If dsEncargos.Tables(0).Rows.Count = 0 Then
                strSQLProduto = ""
                strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                         "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                         "       OXE.Sinal, OXE.PercentualBase, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, OXE.PercentualLimite, OXE.Observacao, " & vbCrLf & _
                         "		 CASE " & vbCrLf & _
                         "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                         "				THEN CASE " & vbCrLf & _
                         "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                         "							THEN 3 " & vbCrLf & _
                         "							ELSE 53 " & vbCrLf & _
                         "						END " & vbCrLf & _
                         "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                         "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                         "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                         "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                         "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                         "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                         "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                         " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                         "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                         "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.EstadoOrigem_Id  ='" & EstadoOrigem & "' " & vbCrLf & _
                         "   AND OXE.EstadoDestino_Id ='" & EstadoDestino & "' " & vbCrLf & _
                         "   AND OXE.Encargo_Id       ='" & Encargo & "' " & vbCrLf & _
                         " ORDER BY OXE.Encargo_Id"

                dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")
            End If

            If dsEncargos.Tables(0).Rows.Count = 0 Then
                strSQLProduto = Produto
                strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                         "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                         "       OXE.Sinal, OXE.PercentualBase, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, OXE.PercentualLimite, OXE.Observacao, " & vbCrLf & _
                         "		 CASE " & vbCrLf & _
                         "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                         "				THEN CASE " & vbCrLf & _
                         "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                         "							THEN 3 " & vbCrLf & _
                         "							ELSE 53 " & vbCrLf & _
                         "						END " & vbCrLf & _
                         "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                         "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                         "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                         "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                         "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                         "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                         "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                         " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                         "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                         "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.EstadoOrigem_Id  ='" & EstadoOrigem & "' " & vbCrLf & _
                         "   AND OXE.EstadoDestino_Id = " & uf & " " & vbCrLf & _
                         "   AND OXE.Encargo_Id       ='" & Encargo & "' " & vbCrLf & _
                         " ORDER BY OXE.Encargo_Id"

                dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")
            End If

            If dsEncargos.Tables(0).Rows.Count = 0 Then
                strSQLProduto = ""
                strSQL = "SELECT OXE.Produto_Id, OXE.EstadoOrigem_Id, OXE.EstadoDestino_Id, OXE.Encargo_Id, isnull(OXE.SituacaoTributaria,999) AS SituacaoTributaria," & vbCrLf & _
                         "       OXE.GrupoFiscal, OXE.CodigoFiscal, OXE.DebitaConta, OXE.CreditaConta," & vbCrLf & _
                         "       OXE.Sinal, OXE.PercentualBase, OXE.Aliquota, isnull(OXE.AliquotaExibicao,0) as AliquotaExibicao, OXE.PercentualLimite, OXE.Observacao, " & vbCrLf & _
                         "		 CASE " & vbCrLf & _
                         "			WHEN OXE.situacaotributariaipi is null " & vbCrLf & _
                         "				THEN CASE " & vbCrLf & _
                         "						WHEN SO.EntradaSaida = 'E' " & vbCrLf & _
                         "							THEN 3 " & vbCrLf & _
                         "							ELSE 53 " & vbCrLf & _
                         "						END " & vbCrLf & _
                         "				ELSE OXE.SituacaoTributariaIPI " & vbCrLf & _
                         "			END AS SituacaoTributariaIPI, " & vbCrLf & _
                         "       isnull(OXE.SituacaoTributariaPISCOFINS,0) as SituacaoTributariaPISCOFINS, isnull(OXE.ObservacaoPISCOFINS,0) as ObservacaoPISCOFINS" & vbCrLf & _
                         "  FROM OperacoesXEncargos OXE " & vbCrLf & _
                         "	INNER JOIN SubOperacoes SO " & vbCrLf & _
                         "			ON  SO.Operacao_Id     = OXE.Operacao_Id " & vbCrLf & _
                         "			AND SO.SubOperacoes_Id = OXE.SubOperacao_Id " & vbCrLf & _
                         " WHERE OXE.GrupoProduto_Id  ='" & GrupoProduto & "' " & vbCrLf & _
                         "   AND OXE.Produto_Id       ='" & strSQLProduto & "' " & vbCrLf & _
                         "   AND OXE.Operacao_Id      = " & Operacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.SubOperacao_Id   = " & SubOperacao.ToString() & " " & vbCrLf & _
                         "   AND OXE.EstadoOrigem_Id  ='" & EstadoOrigem & "' " & vbCrLf & _
                         "   AND OXE.EstadoDestino_Id = " & uf & " " & vbCrLf & _
                         "   AND OXE.Encargo_Id       ='" & Encargo & "' " & vbCrLf & _
                         " ORDER BY OXE.Encargo_Id"

                dsEncargos = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")
            End If

            If dsEncargos.Tables(0).Rows.Count > 0 Then
                Dim drEncargo As DataRow = dsEncargos.Tables(0).Rows(0)

                Dim objOperacoesXEncargos As New OperacaoXEncargo()

                Me.Selecionado = True
                Me.CodigoGrupoProduto = GrupoProduto
                Me.CodigoProduto = drEncargo("Produto_Id").ToString()
                Me.CodigoOperacao = Operacao
                Me.CodigoSubOperacao = SubOperacao
                Me.EstadoOrigem = drEncargo("EstadoOrigem_Id").ToString()
                Me.EstadoDestino = drEncargo("EstadoDestino_Id").ToString()
                Me.CodigoEncargo = drEncargo("Encargo_Id").ToString()

                Me.SituacaoTributariaICMS = drEncargo("SituacaoTributaria")
                Me.SituacaoTributariaIPI = drEncargo("SituacaoTributariaIPI")
                Me.SituacaoTributariaPISCOFINS = drEncargo("SituacaoTributariaPISCOFINS")
                Me.CodigoObservacaoPISCOFINS = drEncargo("ObservacaoPISCOFINS")
                Me.CodigoObservacaoGeral = Convert.ToInt32(drEncargo("Observacao"))

                Me.CodigoGrupoFiscal = Convert.ToInt32(drEncargo("GrupoFiscal"))
                Me.CodigoFiscal = Convert.ToInt32(drEncargo("CodigoFiscal"))
                Me.CodigoDebitaConta = drEncargo("DebitaConta").ToString()
                Me.CodigoCreditaConta = drEncargo("CreditaConta").ToString()
                Me.Sinal = drEncargo("Sinal")
                Me.PercentualBase = drEncargo("PercentualBase").ToString()
                Me.Aliquota = drEncargo("Aliquota").ToString()
                Me.AliquotaExibicao = drEncargo("AliquotaExibicao").ToString()
                Me.PercentualLimite = drEncargo("PercentualLimite").ToString()

                Return True
            Else : Return False
            End If
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "INSERT INTO OperacoesXEncargos (GrupoProduto_Id, Produto_Id, Operacao_Id, SubOperacao_Id, EstadoOrigem_Id, EstadoDestino_Id, Encargo_Id, GrupoFiscal, " & vbCrLf & _
                      "                                CodigoFiscal, DebitaConta, CreditaConta, Sinal, PercentualBase, Aliquota, PercentualLimite, AliquotaExibicao, " & vbCrLf & _
                      "                                SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, Observacao, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                      "VALUES ('" & Me.CodigoGrupoProduto & "','" & Me.CodigoProduto & "'," & Me.CodigoOperacao & "," & Me.CodigoSubOperacao & ",'" & Me.EstadoOrigem & "','" & Me.EstadoDestino & "','" & Me.CodigoEncargo & "'," & Me.CodigoGrupoFiscal & "," & vbCrLf & _
                      Me.CodigoFiscal & ",'" & Me.CodigoDebitaConta & "','" & Me.CodigoCreditaConta & "','" & Me.Sinal & "'," & Str(Me.PercentualBase) & "," & Str(Me.Aliquota) & "," & Str(Me.PercentualLimite) & "," & Str(Me.AliquotaExibicao) & "," & vbCrLf & _
                      Me.SituacaoTributariaICMS & "," & Me.SituacaoTributariaIPI & "," & Me.SituacaoTributariaPISCOFINS & "," & Me.CodigoObservacaoPISCOFINS & "," & Me.CodigoObservacaoGeral & " , '" & UsuarioServidor.Usuario.Usuario_Id & "' , '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "' ) "
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE OperacoesXEncargos SET" & vbCrLf & _
                      "     GrupoFiscal                 = " & Me.CodigoGrupoFiscal & vbCrLf & _
                      "    ,CodigoFiscal                ='" & Me.CodigoFiscal & "'" & vbCrLf & _
                      "    ,DebitaConta                 ='" & Me.CodigoDebitaConta & "'" & vbCrLf & _
                      "    ,CreditaConta                ='" & Me.CodigoCreditaConta & "'" & vbCrLf & _
                      "    ,Sinal                       ='" & Me.Sinal & "'" & vbCrLf & _
                      "    ,PercentualBase              = " & Str(Me.PercentualBase) & vbCrLf & _
                      "    ,Aliquota                    = " & Str(Me.Aliquota) & vbCrLf & _
                      "    ,PercentualLimite            = " & Str(Me.PercentualLimite) & vbCrLf & _
                      "    ,AliquotaExibicao            = " & Str(Me.AliquotaExibicao) & vbCrLf & _
                      "    ,SituacaoTributaria          = " & Me.SituacaoTributariaICMS & vbCrLf & _
                      "    ,SituacaoTributariaIPI       = " & Me.SituacaoTributariaIPI & vbCrLf & _
                      "    ,SituacaoTributariaPISCOFINS = " & Me.SituacaoTributariaPISCOFINS & vbCrLf & _
                      "    ,ObservacaoPISCOFINS         = " & Me.CodigoObservacaoPISCOFINS & vbCrLf & _
                      "    ,Observacao                  = " & Me.CodigoObservacaoGeral & vbCrLf & _
                      "    ,UsuarioAlteracao            ='" & UsuarioServidor.Usuario.Usuario_Id & "'" & vbCrLf & _
                      "    ,UsuarioAlteracaoData        ='" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                      " WHERE GrupoProduto_Id           ='" & Me.CodigoGrupoProduto & "'" & vbCrLf & _
                      "   AND Produto_Id                ='" & Me.CodigoProduto & "'" & vbCrLf & _
                      "   AND Operacao_Id               = " & Me.CodigoOperacao & vbCrLf & _
                      "   AND SubOperacao_Id            = " & Me.CodigoSubOperacao & vbCrLf & _
                      "   AND EstadoOrigem_Id           ='" & Me.EstadoOrigem & "'" & vbCrLf & _
                      "   AND EstadoDestino_Id          ='" & Me.EstadoDestino & "'" & vbCrLf & _
                      "   AND Encargo_Id                ='" & Me.CodigoEncargo & "'" & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE OperacoesXEncargos " & _
                      " WHERE GrupoProduto_Id   ='" & Me.CodigoGrupoProduto & "'" & vbCrLf & _
                      "   AND Produto_Id        ='" & Me.CodigoProduto & "'" & vbCrLf & _
                      "   AND Operacao_Id       = " & Me.CodigoOperacao & vbCrLf & _
                      "   AND SubOperacao_Id    = " & Me.CodigoSubOperacao & vbCrLf & _
                      "   AND EstadoOrigem_Id   ='" & Me.EstadoOrigem & "'" & vbCrLf & _
                      "   AND EstadoDestino_Id  ='" & Me.EstadoDestino & "'" & vbCrLf & _
                      "   AND Encargo_Id        ='" & Me.CodigoEncargo & "'" & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class