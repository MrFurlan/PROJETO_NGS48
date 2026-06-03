Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ConsistenciaDeNotas
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ConsistenciaDeNotas", "ACESSAR") Then
                CarregarUnidade()
                ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                Limpar()
                LiberaEmpresa()
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                CarregarTipoDeDocumento(False)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteCDNF" & HID.Value) Is Nothing Then
            txtCodigoCliente.Value = CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).Codigo) & " - " & CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteCDNF" & HID.Value)
        ElseIf Not Session("objClienteCDNFEmb" & HID.Value) Is Nothing Then
            txtCodigoEmbarque.Value = CType(Session("objClienteCDNFEmb" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objClienteCDNFEmb" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            txtClienteEmbarque.Text = Funcoes.FormatarCpfCnpj(CType(Session("objClienteCDNFEmb" & HID.Value), [Lib].Negocio.Cliente).Codigo) & " - " & CType(Session("objClienteCDNFEmb" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteCDNFEmb" & HID.Value)
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Private Sub CarregarTipoDeDocumento(ByVal todos As Boolean)
        chkTipoDeDocumento.Items.Clear()
        chkTipoDeDocumento.DataValueField = "Codigo"
        chkTipoDeDocumento.DataTextField = "Descricao"
        Dim lst As New [Lib].Negocio.ListTipoDeDocumento()
        chkTipoDeDocumento.DataSource = lst.ToArray()
        chkTipoDeDocumento.DataBind()
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If todos Then
                chkTipoDeDocumento.Items(i).Selected = True
            ElseIf chkTipoDeDocumento.Items(i).Value = 1 Then
                chkTipoDeDocumento.Items(i).Selected = True
            End If
        Next
    End Sub

    Protected Sub chkAllTipos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkAllTipos.Checked Then
                For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
                    chkTipoDeDocumento.Items(i).Selected = True
                Next
            Else
                For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1

                    If chkTipoDeDocumento.Items(i).Value = 1 Then
                        chkTipoDeDocumento.Items(i).Selected = True
                    Else
                        chkTipoDeDocumento.Items(i).Selected = False
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataSet(ByVal objEmpresa As [Lib].Negocio.Cliente, ByVal bConsolidado As Boolean) As DataSet

        SqlConsistenciaNota(objEmpresa, bConsolidado)

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ConsistenciaDeNotas")
        Return ds

    End Function

    Function getDataSetConsolidado(ByVal objEmpresa As [Lib].Negocio.Cliente, ByVal bConsolidado As Boolean) As DataSet

        SqlConsistenciaNota(objEmpresa, bConsolidado)

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ConsistenciaDeNotas")
        Return ds

    End Function

    Private Function SqlConsistenciaNota(objEmpresa As Cliente, ByVal bConsolidado As Boolean) As String

        Sql = "  SELECT N.Empresa_Id, N.EndEmpresa_Id, convert(varchar,N.Movimento,103) AS Movimento, convert(varchar,Pe.DataPedido,103) AS DataPedido, " & vbCrLf
        Sql &= " N.UsuarioInclusao, CONCAT(N.TipoDeDocumento, '-', TP.Descricao) as [Tipo do Documento], N.Nota_Id as Nota, N.Serie_Id as Serie, " & vbCrLf
        Sql &= " N.Pedido, NFI.CFOP_Id AS CFOP, N.Operacao, o.Descricao as OpDescricao, N.SubOperacao, so.Descricao as SubOpDescricao, " & vbCrLf
        Sql &= "        C.Cliente_Id as Cliente, C.Nome, CO.Representante_Id as Representante, CLI.Nome as NomeRepresentante, CO.ValorComissao, " & vbCrLf
        Sql &= "		Case" & vbCrLf
        Sql &= "			When C.Numero > 0" & vbCrLf
        Sql &= "				THEN C.Endereco + ', ' + convert(varchar,C.Numero)" & vbCrLf
        Sql &= "				ELSE C.Endereco" & vbCrLf
        Sql &= "			End as Endereco," & vbCrLf
        Sql &= "        C.Cidade, C.Estado, NFI.Produto_Id as Produto, P.Nome AS NomeDoProduto, P.NCM, ISNULL(NXT.Placa, '') AS Placa, ISNULL(CONVERT(VARCHAR,LP.Laudo), '') AS Laudo,  " & vbCrLf
        Sql &= "        CASE " & vbCrLf
        Sql &= "        WHEN ISNULL(LP.PesoLiquido,0) > 0 " & vbCrLf
        Sql &= "        THEN LP.PesoLiquido " & vbCrLf
        Sql &= "        ELSE NFI.PesoFiscal " & vbCrLf
        Sql &= "        END as PesoLiquido, " & vbCrLf
        Sql &= "        CASE" & vbCrLf
        Sql &= "            WHEN ISNULL(LP.Percentual,0) = 0 " & vbCrLf
        Sql &= "                THEN 'INTACTA-NAO' " & vbCrLf
        Sql &= "            WHEN ISNULL(LP.Percentual,0) = 1 " & vbCrLf
        Sql &= "                THEN '1-NEGATIVO' " & vbCrLf
        Sql &= "            WHEN ISNULL(LP.Percentual,0) = 2 " & vbCrLf
        Sql &= "                THEN '2-POSITIVO' " & vbCrLf
        Sql &= "            WHEN ISNULL(LP.Percentual,0) = 3 " & vbCrLf
        Sql &= "                THEN '3-DECLARADO' " & vbCrLf
        Sql &= "            WHEN ISNULL(LP.Percentual,0) = 4 " & vbCrLf
        Sql &= "                THEN '4-PARTICIPANTE' " & vbCrLf
        Sql &= "            WHEN ISNULL(LP.Percentual,0) = 5 " & vbCrLf
        Sql &= "                THEN '5-ORIGEM PARTICIPANTE' " & vbCrLf
        Sql &= "                ELSE '' " & vbCrLf
        Sql &= "        END as Monsanto, " & vbCrLf
        Sql &= "        isnull(P.RegMinAgr,'') AS RegMinAgr, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(NAVXIN.Codigo_Id,0) > 0" & vbCrLf
        Sql &= "				THEN convert(varchar,NAVXIN.Codigo_Id) +'-'+ NAV.Descricao" & vbCrLf
        Sql &= "				ELSE ''" & vbCrLf
        Sql &= "			END as Invoice, " & vbCrLf
        Sql &= "        Pg.Pagamento_Id as CondicaoPagamento ,Pg.Descricao as DescricaoCodPagamento," & vbCrLf
        Sql &= "        NFI.QuantidadeFiscal as Quantidade," & vbCrLf
        Sql &= "        CASE" & vbCrLf
        Sql &= "            WHEN isnull(pxi.FatorConversao,0) > 0" & vbCrLf
        Sql &= "                THEN NFI.QuantidadeFiscal * pxi.FatorConversao" & vbCrLf
        Sql &= "                ELSE NFI.QuantidadeFiscal" & vbCrLf
        Sql &= "        END as QuantidadeItemPedido," & vbCrLf
        Sql &= "        NFI.Unitario, NFI.Valor, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFunrural.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncFunrural.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Funrural, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncSenar.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncSenar.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Senar, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFethab.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncFethab.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as BaseFethab, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFacs.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncFacs.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Facs, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFethab.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncFethab.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Fethab, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIcms.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncIcms.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as BaseIcms, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIcms.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncIcms.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Icms, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIcmsST.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncIcmsST.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as BaseIcmsST, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIcmsST.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncIcmsST.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as IcmsST, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFECP.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncFECP.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END AS BaseFECP, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFECP.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncFECP.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END AS FECP, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncPis.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncPis.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as BasePis, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncPis.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncPis.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Pis, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncCofins.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncCofins.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as BaseCofins, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncCofins.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncCofins.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Cofins, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIpi.Base,0) > 0" & vbCrLf
        Sql &= "				THEN EncIpi.Base" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as BaseIpi, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIpi.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncIpi.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Ipi, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncFrete.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncFrete.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Fretes, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncDespAcess.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncDespAcess.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as [Despesa Acessoria], " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncDesconto.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncDesconto.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as Desconto, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncCsrf.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncCsrf.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as CSRF, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncInss.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncInss.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as INSS, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncIrrf.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncIrrf.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as IRRF, " & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN ISNULL(EncLiquido.Valor,0) > 0" & vbCrLf
        Sql &= "				THEN EncLiquido.Valor" & vbCrLf
        Sql &= "				ELSE 0" & vbCrLf
        Sql &= "			END as ValorTotal, " & vbCrLf
        Sql &= "		NFI.OperacaoXEstado AS OPXENCARGO," & vbCrLf
        Sql &= "		OxE.STICMS as STICMS," & vbCrLf
        Sql &= "		OxE.STPISCOFINS as STPISCOFINS," & vbCrLf
        Sql &= "		NFER.ChaveNfe as ChaveNfe," & vbCrLf
        Sql &= "		CASE" & vbCrLf
        Sql &= "			WHEN so.EntradaSaida = 'E'" & vbCrLf
        Sql &= "				THEN oep.DebitaConta" & vbCrLf
        Sql &= "				ELSE oep.CreditaConta" & vbCrLf
        Sql &= "			END as ContaProduto, " & vbCrLf
        Sql &= "			ISNULL(PLA.Titulo, '') AS [Descrição da Conta], " & vbCrLf
        Sql &= "		ISNULL(NEXP.FaturaExportacao,'') AS FaturaExportacao, " & vbCrLf
        Sql &= "		IIF(LEN(EMB.Nome) > 1, LEFT(ISNULL(EMB.Nome, ''), 28) + ' - ' + LEFT(ISNULL(EMB.Cidade, ''), 28) + ' - ' + LEFT(ISNULL(EMB.Estado, ''), 28) " & vbCrLf
        Sql &= "		+ ' - ' + STUFF(STUFF(STUFF(STUFF(Replace(EMB.Cliente_Id, ' ', ''), 3, 0, '.'), 7, 0, '.'), 11, 0, '/'), 16, 0, '-') " & vbCrLf
        Sql &= "		+ ' - ' + LEFT(ISNULL(EMP.Endereco_Id, ''), 28), '') AS LocalDeEmbarque, " & vbCrLf
        Sql &= "		Pe.Observacoes "

        If bConsolidado Then
            Sql &= "  INTO #tempNotas " & vbCrLf
        End If

        Sql &= "  FROM  NotasFiscais N" & vbCrLf
        Sql &= "		 INNER JOIN Clientes C" & vbCrLf
        Sql &= "				 ON N.Cliente_Id            = C.Cliente_Id" & vbCrLf
        Sql &= "				 AND N.EndCliente_Id        = C.Endereco_Id" & vbCrLf
        Sql &= "		 INNER JOIN Clientes EMP" & vbCrLf
        Sql &= "				 ON N.Empresa_Id            = EMP.Cliente_Id" & vbCrLf
        Sql &= "				 AND N.EndEmpresa_Id        = EMP.Endereco_Id" & vbCrLf
        Sql &= "		 LEFT JOIN Clientes EMB" & vbCrLf
        Sql &= "				 ON N.LocalEmbarque         = EMB.Cliente_Id" & vbCrLf
        Sql &= "				 AND N.EndLocalEmbarque     = EMB.Endereco_Id" & vbCrLf
        Sql &= "		 INNER JOIN NotasFiscaisXItens NFI" & vbCrLf
        Sql &= "				 ON N.Empresa_Id            = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND N.EndEmpresa_Id        = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND N.Cliente_Id           = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND N.EndCliente_Id        = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND N.EntradaSaida_Id      = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND N.Serie_Id             = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND N.Nota_Id              = NFI.Nota_Id" & vbCrLf
        Sql &= "		 INNER JOIN TipoDeDocumento TP" & vbCrLf
        Sql &= "				 ON N.TipoDeDocumento       = TP.Codigo_Id" & vbCrLf
        Sql &= "		 INNER JOIN OperacaoXEstadoXEncargo oeP" & vbCrLf
        Sql &= "			 	 On oeP.Codigo_Id           = NFI.OperacaoXEstado" & vbCrLf
        Sql &= "				And oeP.Encargo_Id          = 'PRODUTO'" & vbCrLf
        Sql &= "		 LEFT JOIN ( SELECT  NXR.Empresa_Id," & vbCrLf
        Sql &= "				            NXR.EndEmpresa_Id," & vbCrLf
        Sql &= "				            NXR.Cliente_Id," & vbCrLf
        Sql &= "				            NXR.EndCliente_Id," & vbCrLf
        Sql &= "				            NXR.EntradaSaida_Id," & vbCrLf
        Sql &= "				            NXR.Serie_Id," & vbCrLf
        Sql &= "				            NXR.Nota_Id," & vbCrLf
        Sql &= "				            LP_INT.Pesagem_Id As Laudo," & vbCrLf
        Sql &= "				            RxD.Percentual," & vbCrLf
        Sql &= "				            R.PesoLiquido" & vbCrLf
        Sql &= "				    FROM Romaneios R " & vbCrLf
        Sql &= "				LEFT Join RomaneiosXPesagens RXP " & vbCrLf
        Sql &= "				    ON R.Empresa_Id             = RXP.Empresa_Id " & vbCrLf
        Sql &= "				    AND R.EndEmpresa_Id         = RXP.EndEmpresa_Id " & vbCrLf
        Sql &= "                    AND R.Romaneio_Id           = RXP.Romaneio_Id" & vbCrLf
        Sql &= "				LEFT JOIN RomaneiosXDescontos RxD " & vbCrLf
        Sql &= "				    ON RxD.Romaneio_Id    = R.Romaneio_Id " & vbCrLf
        Sql &= "				    AND RxD.Empresa_Id    = R.Empresa_Id " & vbCrLf
        Sql &= "                    AND RxD.EndEmpresa_Id = R.EndEmpresa_Id" & vbCrLf
        Sql &= "                    AND RxD.Analise_Id    = 12" & vbCrLf
        Sql &= "				LEFT OUTER JOIN NotasFiscaisXRomaneios NXR" & vbCrLf
        Sql &= "				    ON R.Empresa_Id             = NXR.Empresa_Id " & vbCrLf
        Sql &= "				    AND R.EndEmpresa_Id         = NXR.EndEmpresa_Id " & vbCrLf
        Sql &= "                    AND R.Romaneio_Id           = NXR.Romaneio_Id" & vbCrLf
        Sql &= "				LEFT Join Pesagem LP_INT " & vbCrLf
        Sql &= "				    ON RxP.Empresa_Id           = LP_INT.Empresa_Id  " & vbCrLf
        Sql &= "				    AND RxP.EndEmpresa_Id       = LP_INT.EndEmpresa_Id" & vbCrLf
        Sql &= "                    AND RxP.Pesagem_Id          = LP_INT.Pesagem_Id" & vbCrLf
        Sql &= "				    AND LP_INT.Sequencia_Id     = 0 ) As LP" & vbCrLf
        Sql &= "				ON NFI.Empresa_Id           = LP.Empresa_Id" & vbCrLf
        Sql &= "				AND NFI.EndEmpresa_Id       = LP.EndEmpresa_Id" & vbCrLf
        Sql &= "                AND NFI.Cliente_Id          = LP.Cliente_Id" & vbCrLf
        Sql &= "				AND NFI.EndCliente_Id       = LP.EndCliente_Id" & vbCrLf
        Sql &= "                AND NFI.EntradaSaida_Id     = LP.EntradaSaida_Id" & vbCrLf
        Sql &= "				AND NFI.Serie_Id            = LP.Serie_Id" & vbCrLf
        Sql &= "                AND NFI.Nota_Id = LP.Nota_Id" & vbCrLf
        Sql &= "        LEFT JOIN NotasFiscaisXTransportadores NxT" & vbCrLf
        Sql &= "                ON NFI.Empresa_Id           = NxT.Empresa_Id" & vbCrLf
        Sql &= "                AND NFI.EndEmpresa_Id       = NxT.EndEmpresa_Id" & vbCrLf
        Sql &= "                AND NFI.Cliente_Id          = NxT.Cliente_Id" & vbCrLf
        Sql &= "                AND NFI.EndCliente_Id       = NxT.EndCliente_Id" & vbCrLf
        Sql &= "                AND NFI.EntradaSaida_Id     = NxT.EntradaSaida_Id" & vbCrLf
        Sql &= "                AND NFI.Serie_Id            = NxT.Serie_Id" & vbCrLf
        Sql &= "                AND NFI.Nota_Id             = NxT.Nota_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'FUNRURAL'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncFunrural" & vbCrLf
        Sql &= "				 ON EncFunrural.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncFunrural.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncFunrural.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id           = 'SENAR'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncSenar" & vbCrLf
        Sql &= "				 ON EncSenar.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncSenar.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncSenar.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncSenar.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncSenar.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncSenar.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncSenar.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncSenar.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncSenar.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncSenar.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where (Encargo_Id = 'ICMS' OR Encargo_Id = 'ICMS A REC.')" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncIcms" & vbCrLf
        Sql &= "				 ON EncIcms.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncIcms.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncIcms.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncIcms.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncIcms.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncIcms.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncIcms.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncIcms.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncIcms.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncIcms.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'ICMS-ST'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncIcmsST" & vbCrLf
        Sql &= "				 ON EncIcmsST.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncIcmsST.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncIcmsST.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'FUNDO FECP'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncFECP" & vbCrLf
        Sql &= "				 ON EncFECP.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncFECP.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncFECP.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncFECP.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncFECP.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncFECP.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncFECP.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncFECP.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncFECP.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncFECP.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id IN ('PIS','PIS RECUP.')" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncPis" & vbCrLf
        Sql &= "				 ON EncPis.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncPis.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncPis.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncPis.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncPis.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncPis.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncPis.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncPis.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncPis.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncPis.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id IN ('COFINS','COFINS A REC.')" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncCofins" & vbCrLf
        Sql &= "				 ON EncCofins.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncCofins.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncCofins.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncCofins.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncCofins.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncCofins.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncCofins.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncCofins.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncCofins.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncCofins.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'IPI'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncIpi" & vbCrLf
        Sql &= "				 ON EncIpi.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncIpi.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncIpi.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncIpi.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncIpi.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncIpi.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncIpi.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncIpi.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncIpi.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncIpi.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'FETHAB'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncFethab" & vbCrLf
        Sql &= "				 ON EncFethab.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncFethab.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncFethab.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncFethab.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncFethab.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncFethab.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncFethab.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncFethab.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncFethab.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncFethab.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'FRETES'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncFrete" & vbCrLf
        Sql &= "				 ON EncFrete.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncFrete.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncFrete.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncFrete.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncFrete.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncFrete.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncFrete.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncFrete.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncFrete.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncFrete.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'DESP. ACESSORIA'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncDespAcess" & vbCrLf
        Sql &= "				 ON EncDespAcess.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncDespAcess.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncDespAcess.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'DESCONTOS'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncDesconto" & vbCrLf
        Sql &= "				 ON EncDesconto.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncDesconto.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncDesconto.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'LIQUIDO'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncLiquido" & vbCrLf
        Sql &= "				 ON EncLiquido.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncLiquido.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncLiquido.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'FACS'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncFacs" & vbCrLf
        Sql &= "				 ON EncFacs.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncFacs.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncFacs.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncFacs.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncFacs.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncFacs.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncFacs.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncFacs.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncFacs.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncFacs.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'CSRF'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncCsrf" & vbCrLf
        Sql &= "				 ON EncCsrf.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncCsrf.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncCsrf.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'INSS RETIDO'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncInss" & vbCrLf
        Sql &= "				 ON EncInss.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncInss.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncInss.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncInss.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncInss.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncInss.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncInss.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncInss.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncInss.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncInss.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor" & vbCrLf
        Sql &= "					 from NotasFiscaisXEncargos" & vbCrLf
        Sql &= "					 Where Encargo_Id       = 'IRRF PJ'" & vbCrLf
        Sql &= "					   AND Valor > 0 )  as EncIrrf" & vbCrLf
        Sql &= "				 ON EncIrrf.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND EncIrrf.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND EncIrrf.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 LEFT JOIN NotaFiscalXExportacao NEXP" & vbCrLf
        Sql &= "				 ON NEXP.Empresa_Id       = N.Empresa_Id" & vbCrLf
        Sql &= "				 AND NEXP.EndEmpresa_Id   = N.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND NEXP.Cliente_Id      = N.Cliente_Id" & vbCrLf
        Sql &= "				 AND NEXP.EndCliente_Id   = N.EndCliente_Id " & vbCrLf
        Sql &= "				 AND NEXP.EntradaSaida_Id = N.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND NEXP.Serie_Id        = N.Serie_Id" & vbCrLf
        Sql &= "				 AND NEXP.Nota_Id         = N.Nota_Id" & vbCrLf
        Sql &= "		 LEFT JOIN NfeRealizadas NFER" & vbCrLf
        Sql &= "				 ON NFER.Empresa_Id       = N.Empresa_Id" & vbCrLf
        Sql &= "				 AND NFER.EndEmpresa_Id   = N.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND NFER.Cliente_Id      = N.Cliente_Id" & vbCrLf
        Sql &= "				 AND NFER.EndCliente_Id   = N.EndCliente_Id " & vbCrLf
        Sql &= "				 AND NFER.EntradaSaida_Id = N.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND NFER.Serie_Id        = N.Serie_Id" & vbCrLf
        Sql &= "				 AND NFER.Nota_Id         = N.Nota_Id" & vbCrLf
        Sql &= "		 LEFT JOIN DocumentoXML DXml" & vbCrLf
        Sql &= "				 ON DXml.Empresa_Id       = NFER.Empresa_Id" & vbCrLf
        Sql &= "				 AND DXml.Chave_Id        = NFER.ChaveNFe" & vbCrLf
        Sql &= "		 INNER JOIN Produtos P" & vbCrLf
        Sql &= "				 ON NFI.Produto_Id     = P.Produto_Id" & vbCrLf
        Sql &= "		 INNER JOIN SubOperacoes so" & vbCrLf
        Sql &= "				 ON so.Operacao_id      = NFI.Operacao" & vbCrLf
        Sql &= "				 AND so.SubOperacoes_Id = NFI.SubOperacao" & vbCrLf
        Sql &= "		 INNER JOIN Operacoes o" & vbCrLf
        Sql &= "				 ON o.Operacao_id      = NFI.Operacao" & vbCrLf
        Sql &= "         LEFT JOIN (" & vbCrLf
        Sql &= "                        SELECT Conta_Id, Titulo " & vbCrLf
        Sql &= "                        FROM PlanoDeContas " & vbCrLf
        Sql &= "                        GROUP BY Conta_Id, Titulo " & vbCrLf
        Sql &= "                   ) PLA " & vbCrLf
        Sql &= "                ON CASE " & vbCrLf
        Sql &= "         	        WHEN so.EntradaSaida = 'E' THEN " & vbCrLf
        Sql &= "         		        oep.DebitaConta " & vbCrLf
        Sql &= "         		    ELSE oep.CreditaConta " & vbCrLf
        Sql &= "         	       END = PLA.Conta_Id " & vbCrLf
        Sql &= "		 LEFT JOIN Comissoes As CO" & vbCrLf
        Sql &= "				 On CO.Empresa_Id    = N.Empresa_Id" & vbCrLf
        Sql &= "				 And CO.EndEmpresa_Id = N.EndEmpresa_Id" & vbCrLf
        Sql &= "				 And CO.Pedido_Id = N.Pedido" & vbCrLf
        Sql &= "		 LEFT JOIN Clientes As CLI" & vbCrLf
        Sql &= "				 On CLI.Cliente_Id    = CO.Representante_Id" & vbCrLf
        Sql &= "				 And CLI.Endereco_Id = CO.EndRepresentante_Id" & vbCrLf
        Sql &= "		 LEFT JOIN Pedidos As Pe" & vbCrLf
        Sql &= "				 On Pe.Empresa_Id    = N.Empresa_Id" & vbCrLf
        Sql &= "				 And Pe.EndEmpresa_Id = N.EndEmpresa_Id" & vbCrLf
        Sql &= "				 And Pe.Pedido_Id = N.Pedido" & vbCrLf
        Sql &= "         LEFT JOIN PedidoXItem pxi " & vbCrLf
        Sql &= "            On Pe.empresa_id        = pxi.empresa_id " & vbCrLf
        Sql &= "            And Pe.endempresa_id    = pxi.endempresa_id " & vbCrLf
        Sql &= "            And Pe.pedido_id        = pxi.pedido_id " & vbCrLf
        Sql &= "            And NFI.Produto_Id      = pxi.Produto_Id " & vbCrLf
        Sql &= "		 LEFT JOIN Pagamentos Pg" & vbCrLf
        Sql &= "				 On Pg.Pagamento_Id = Pe.CondicaoPagamento" & vbCrLf
        Sql &= "		 LEFT JOIN NaviosXInvoice NAVXIN" & vbCrLf
        Sql &= "				 On NAVXIN.Codigo_Id = N.InvoiceNavio" & vbCrLf
        Sql &= "		 LEFT JOIN Navios NAV" & vbCrLf
        Sql &= "				 On NAV.Codigo_Id = NAVXIN.Navio_Id" & vbCrLf
        Sql &= "		 LEFT JOIN OperacaoXEstado OxE" & vbCrLf
        Sql &= "				 On OxE.Codigo_Id = NFI.OperacaoXEstado" & vbCrLf
        Sql &= "        Where N.Situacao = 1 " & vbCrLf

        If Not chkUnificarEmpresa.Checked Then
            Sql &= "        And N.Empresa_Id = '" & objEmpresa.Codigo & "'" & vbCrLf
            Sql &= "        And N.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf
        End If

        'APENAS NOTAS DO MESMO ESTADO DA EMPRESA - FURLAN - 07/10/2024
        If rdDentroEstado.Checked Then
            Sql &= "        And EMP.Estado = C.Estado " & vbCrLf
        End If

        'APENAS NOTAS DO ESTADO DIFERENTE DA EMPRESA - FURLAN - 07/10/2024
        If rdForaEstado.Checked Then
            Sql &= "        And NOT EMP.Estado = C.Estado " & vbCrLf
        End If

        If chkNossaEmissao.Checked Then
            Sql &= "AND (N.NossaEmissao = 'S' OR isnull(DXml.Tipo,'') = 'Mic')" & vbCrLf
        End If

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= "          And N.Cliente_Id = '" & strCliente(0) & "' And N.EndCliente_Id = " & strCliente(1) & vbCrLf
        End If

        If txtCodigoEmbarque.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoEmbarque.Value.ToString.Split("-")
            Sql &= "          And N.LocalEmbarque = '" & strCliente(0) & "' And N.EndLocalEmbarque = " & strCliente(1) & vbCrLf
        End If

        Sql &= "          And N.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
        Sql &= "          And N.EntradaSaida_Id = '" & IIf(RadEntradas.Checked, "E", "S") & "'" & vbCrLf

        VerificarGrupoProduto(Sql)

        If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND N.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf
        If cmbSubOperacao.SelectedIndex > 0 Then
            Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            Sql &= "AND N.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
        End If

        If IsNumeric(txtOpercaoXEstado.Text) Or txtOpercaoXEstado.Text.Trim() <> "" Then
            Sql &= "AND NFI.OperacaoXEstado = " & txtOpercaoXEstado.Text & vbCrLf
        End If

        Dim tp As String = ""
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If chkTipoDeDocumento.Items(i).Selected Then
                tp &= IIf(tp.Length = 0, "", ",") & chkTipoDeDocumento.Items(i).Value
            End If
        Next

        If tp <> "" Then
            Sql &= " AND n.TipoDeDocumento in(" & tp & ")" & vbCrLf
        End If

        If RadMovimento.Checked Then
            Sql &= " Order By N.Empresa_Id, N.EndEmpresa_Id, N.Movimento" & vbCrLf
        Else
            Sql &= " Order By N.Empresa_Id, N.EndEmpresa_Id, C.Nome" & vbCrLf
        End If

        If bConsolidado Then

            Sql &= " SELECT  Empresa_Id,                            " & vbCrLf &
                    "		 EndEmpresa_Id,                         " & vbCrLf &
                    "		 Movimento,                             " & vbCrLf &
                    "		 DataPedido,                            " & vbCrLf &
                    "		 Nota,                                  " & vbCrLf &
                    "		 Serie,                                 " & vbCrLf &
                    "		 Pedido,                                " & vbCrLf &
                    "		 Cliente,                               " & vbCrLf &
                    "		 Nome,                                  " & vbCrLf &
                    "		 Representante,                         " & vbCrLf &
                    "		 NomeRepresentante,                     " & vbCrLf &
                    "		 ValorComissao,                         " & vbCrLf &
                    "		 Endereco,                              " & vbCrLf &
                    "		 Cidade,                                " & vbCrLf &
                    "		 Estado,                                " & vbCrLf &
                    "		 CondicaoPagamento,                     " & vbCrLf &
                    "		 DescricaoCodPagamento,                 " & vbCrLf &
                    "		 SUM(Quantidade) AS Quantidade,         " & vbCrLf &
                    "		 SUM(QuantidadeItemPedido)              " & vbCrLf &
                    "        AS QuantidadeItemPedido,               " & vbCrLf &
                    "		 --Unitario,                            " & vbCrLf &
                    "		 SUM(Valor) AS Valor,                   " & vbCrLf &
                    "		 SUM(Funrural) AS Funrural,             " & vbCrLf &
                    "		 SUM(Senar) AS Senar,                   " & vbCrLf &
                    "		 SUM(BaseIcms) AS BaseIcms,             " & vbCrLf &
                    "		 SUM(Icms) AS Icms,                     " & vbCrLf &
                    "		 SUM(BaseIcmsST) AS BaseIcmsST,         " & vbCrLf &
                    "		 SUM(IcmsST) AS IcmsST,                 " & vbCrLf &
                    "		 SUM(BasePis) AS BasePis,               " & vbCrLf &
                    "		 SUM(Pis) AS Pis,                       " & vbCrLf &
                    "		 SUM(BaseCofins) AS BaseCofins,         " & vbCrLf &
                    "		 SUM(Cofins) AS Cofins,                 " & vbCrLf &
                    "		 SUM(BaseIpi) AS BaseIpi,               " & vbCrLf &
                    "		 SUM(Ipi) AS Ipi,                       " & vbCrLf &
                    "		 SUM(Desconto) AS Desconto              " & vbCrLf &
                    " FROM #tempNotas                               " & vbCrLf &
                    " GROUP BY Empresa_Id,                          " & vbCrLf &
                    "		  EndEmpresa_Id,                        " & vbCrLf &
                    "		  Movimento,                            " & vbCrLf &
                    "		  DataPedido,                           " & vbCrLf &
                    "		  Nota,                                 " & vbCrLf &
                    "		  Serie,                                " & vbCrLf &
                    "		  Pedido,                               " & vbCrLf &
                    "		  Cliente,                              " & vbCrLf &
                    "		  Nome,                                 " & vbCrLf &
                    "		  Representante,                        " & vbCrLf &
                    "		  NomeRepresentante,                    " & vbCrLf &
                    "		  ValorComissao,                        " & vbCrLf &
                    "		  Endereco,                             " & vbCrLf &
                    "		  Cidade,                               " & vbCrLf &
                    "		  Estado,                               " & vbCrLf &
                    "		  CondicaoPagamento,                    " & vbCrLf &
                    "		  DescricaoCodPagamento                 "
        End If

        Return Sql

    End Function

    Function getDataSetPorPlaca(ByVal objEmpresa As [Lib].Negocio.Cliente) As DataSet
        Sql = "SELECT N.Empresa_Id, N.EndEmpresa_Id, n.Movimento, " & vbCrLf &
                "		case " & vbCrLf &
                "			when n.EntradaSaida_Id = 'E' " & vbCrLf &
                "				then c.Cidade " & vbCrLf &
                "				else e.Cidade " & vbCrLf &
                "			end as Origem, " & vbCrLf &
                "		case " & vbCrLf &
                "			when n.EntradaSaida_Id = 'E' " & vbCrLf &
                "				then e.Cidade " & vbCrLf &
                "				else c.Cidade " & vbCrLf &
                "			end as Destino, " & vbCrLf &
                "		nt.Placa, n.Nota_Id AS Nota, NFI.CFOP_Id as CFOP, n.Serie_Id AS Serie, n.EntradaSaida_Id AS EntSai, Max(p.Pesagem_Id) as Laudo, sum(NFI.Valor) as TotalNF " & vbCrLf &
                "FROM NotasFiscais AS n  " & vbCrLf &
                "			INNER JOIN Clientes e " & vbCrLf &
                "					ON n.Empresa_Id     = e.Cliente_Id " & vbCrLf &
                "					AND n.EndEmpresa_Id = e.Endereco_Id " & vbCrLf &
                "			INNER JOIN Clientes c " & vbCrLf &
                "					ON n.Cliente_Id     = c.Cliente_Id " & vbCrLf &
                "					AND n.EndCliente_Id = c.Endereco_Id " & vbCrLf &
                "			INNER JOIN  NotasFiscaisXItens NFI " & vbCrLf &
                "					ON n.Empresa_Id       = NFI.Empresa_Id " & vbCrLf &
                "					AND n.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf &
                "					AND n.Cliente_Id      = NFI.Cliente_Id " & vbCrLf &
                "					AND n.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf &
                "					AND n.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf &
                "					AND n.Serie_Id        = NFI.Serie_Id " & vbCrLf &
                "					AND n.Nota_Id         = NFI.Nota_Id " & vbCrLf &
                "			INNER JOIN  NotasFiscaisXTransportadores nt " & vbCrLf &
                "					ON n.Empresa_Id       = nt.Empresa_Id " & vbCrLf &
                "					AND n.EndEmpresa_Id   = nt.EndEmpresa_Id " & vbCrLf &
                "					AND n.Cliente_Id      = nt.Cliente_Id " & vbCrLf &
                "					AND n.EndCliente_Id   = nt.EndCliente_Id " & vbCrLf &
                "					AND n.EntradaSaida_Id = nt.EntradaSaida_Id " & vbCrLf &
                "					AND n.Serie_Id        = nt.Serie_Id " & vbCrLf &
                "					AND n.Nota_Id         = nt.Nota_Id " & vbCrLf &
                "			INNER JOIN  Pesagem p  " & vbCrLf &
                "					ON n.Empresa_Id       = p.Empresa_Id  " & vbCrLf &
                "					AND n.EndEmpresa_Id   = p.EndEmpresa_Id  " & vbCrLf &
                "					AND n.EntradaSaida_Id = p.EntradaSaida " & vbCrLf &
                "		    LEFT JOIN NfeRealizadas NFER" & vbCrLf &
                "				   ON NFER.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                "				   AND NFER.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                "				   AND NFER.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                "				   AND NFER.EndCliente_Id   = n.EndCliente_Id " & vbCrLf &
                "				   AND NFER.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                "				   AND NFER.Serie_Id        = n.Serie_Id" & vbCrLf &
                "				   AND NFER.Nota_Id         = n.Nota_Id" & vbCrLf &
                "		 LEFT JOIN DocumentoXML DXml" & vbCrLf &
                "				 ON DXml.Empresa_Id       = NFER.Empresa_Id" & vbCrLf &
                "				 AND DXml.Chave_Id        = NFER.ChaveNFe" & vbCrLf &
                "WHERE (n.Situacao = 1) " & vbCrLf &
                "  AND (n.Empresa_Id = '" & objEmpresa.Codigo & "') " & vbCrLf &
                "  AND (n.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
                "  AND (n.NFG = 0) " & vbCrLf &
                "  AND (n.EntradaSaida_Id = '" & IIf(RadEntradas.Checked, "E", "S") & "')" & vbCrLf

        If chkNossaEmissao.Checked Then
            Sql &= "AND (n.NossaEmissao = 'S' OR isnull(DXml.Tipo,'') = 'Mic')" & vbCrLf
        End If

        If Trim(txtPlaca.Text).Length > 0 Then
            Sql &= "  AND (nt.Placa = '" & Trim(txtPlaca.Text) & "') " & vbCrLf
        End If

        Dim tp As String = ""
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If chkTipoDeDocumento.Items(i).Selected Then
                tp &= IIf(tp.Length = 0, "", ",") & chkTipoDeDocumento.Items(i).Value
            End If
        Next

        If tp <> "" Then
            Sql &= " AND n.TipoDeDocumento in(" & tp & ")" & vbCrLf
        End If

        Sql &= "GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Movimento, e.Cidade, c.Cidade, nt.Placa, n.Nota_Id, NFI.CFOP_Id, n.Serie_Id, n.EntradaSaida_Id " & vbCrLf &
                "ORDER BY n.Movimento, n.Nota_Id "

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ConsistenciaDeNotas")
        Return ds
    End Function

    Function Validar() As Boolean
        If ddlUnidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Unidade de Negocio")
            Return False
        ElseIf DdlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Empresa")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        txtCodigoEmbarque.Value = ""
        txtClienteEmbarque.Text = ""
        txtOpercaoXEstado.Text = ""

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Unidade: " & ddlUnidade.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à: " & txtDataFinal.Text & " - "
        End If

        Return param
    End Function

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteCDNF" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmdClienteEmbarque_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteCDNFEmb" & HID.Value, "txtNome")
    End Sub

    Private Function VerificarGrupoProduto(ByRef Sql As String) As String
        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "NFI.Produto_id")
            Sql &= " AND " & retorno(0)
            Return retorno(1)
        Else
            Return ""
        End If
    End Function

    Private Function FileToByteArray(ByVal fileName As String) As Byte()
        Dim buffer As Byte() = Nothing
        Try
            Dim stream As New System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            Dim reader As New System.IO.BinaryReader(stream)
            Dim bytes As Long = New System.IO.FileInfo(fileName).Length
            buffer = reader.ReadBytes(CType(bytes, Long))
            stream.Close()
            stream.Dispose()
            reader.Close()
        Catch ex As Exception
            Debug.WriteLine("Exception caught in process: {0}", ex.ToString())
            Throw ex
        End Try
        Return buffer
    End Function

    Private Sub GerarExcel(ByVal periodo As String, ByVal objEmpresa As [Lib].Negocio.Cliente)
        Try

            Dim ds As DataSet = getDataSet(objEmpresa, False)
            Dim dsConsolidado As DataSet = getDataSetConsolidado(objEmpresa, True)

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não existem dados para gerar o relatório!", eTitulo.Info)
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    PlanilhaConsistenciaDeNotas(objEmpresa, ds, package)

                    PlanilhaConsistenciaDeNotasConsolidado(objEmpresa, dsConsolidado, package)

                    'salvando planilha do excel
                    package.Save()

                End Using
            End Using

            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub PlanilhaConsistenciaDeNotas(objEmpresa As Cliente, ds As DataSet, package As ExcelPackage)

        'criando planilha títulos
        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("CONSISTÊNCIA DE NOTAS")

        'criando linha com o cabeçalho da planilha
        Dim rowIndex As Integer = 1
        Dim columnIndex As Integer = 1

        'criando linha que informa o nome da empresa e o cnpj
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa a cidade e o estado da empresa
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa o título do relatório
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "CONSISTÊNCIA DE NOTAS")
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa o período selecionado na página
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)) & " - " & IIf(RadEntradas.Checked, "ENTRADAS", "SAÍDAS"))
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha com o cabeçalho da planilha
        For Each col As DataColumn In ds.Tables(0).Columns
            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
            columnIndex += 1
        Next

        Dim linhaNomeColuna = rowIndex

        'criando auto filtro na planilha
        worksheet.Cells("A5:AA" & rowIndex).AutoFilter = True

        'aplicando formatação nas células do cabeçalho
        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
            range.Style.Font.Bold = True
            range.Style.Fill.PatternType = ExcelFillStyle.Solid
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
            range.Style.Font.Color.SetColor(Color.White)
        End Using
        rowIndex += 1

        Dim colunaFim As Integer = 0

        ' criando conteúdo da planilha com os dados do dataset
        For Each row As DataRow In ds.Tables(0).Rows
            columnIndex = 1
            For Each col As DataColumn In ds.Tables(0).Columns

                If col.ColumnName = "Quantidade" OrElse
                    col.ColumnName = "QuantidadeItemPedido" OrElse
                    col.ColumnName = "Unitario" OrElse
                    col.ColumnName = "Valor" OrElse
                    col.ColumnName = "Funrural" OrElse
                    col.ColumnName = "Senar" OrElse
                    col.ColumnName = "BaseIcms" OrElse
                    col.ColumnName = "Icms" OrElse
                    col.ColumnName = "BaseIcmsST" OrElse
                    col.ColumnName = "IcmsST" OrElse
                    col.ColumnName = "BaseFECP" OrElse
                    col.ColumnName = "FECP" OrElse
                    col.ColumnName = "BasePis" OrElse
                    col.ColumnName = "Pis" OrElse
                    col.ColumnName = "BaseCofins" OrElse
                    col.ColumnName = "Cofins" OrElse
                    col.ColumnName = "BaseIpi" OrElse
                    col.ColumnName = "Ipi" OrElse
                    col.ColumnName = "BaseFethab" OrElse
                    col.ColumnName = "Fethab" OrElse
                    col.ColumnName = "Desconto" OrElse
                    col.ColumnName = "ValorTotal" Then
                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToDecimal(row(col.ColumnName))
                Else
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                End If

                columnIndex += 1
            Next

            'aplicando formatação nas células do conteúdo
            If rowIndex Mod 2 = 0 Then
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                End Using
            End If
            rowIndex += 1
        Next


        Dim linhaCabecalho As Integer = linhaNomeColuna
        Dim colunaTotalLetra As String = ""
        Dim colunaFimDadosLetra As String = ""

        Dim colunaQuantidadeLetra As String = ""
        Dim colunaQuantidadeItemPedidoLetra As String = ""
        Dim colunaValorLetra As String = ""
        Dim colunaFunruralLetra As String = ""
        Dim colunaSenarLetra As String = ""
        Dim colunaBaseIcmsLetra As String = ""
        Dim colunaIcmsLetra As String = ""
        Dim colunaBaseIcmsSTLetra As String = ""
        Dim colunaIcmsSTLetra As String = ""
        Dim colunaBaseFECPLetra As String = ""
        Dim colunaFECPLetra As String = ""
        Dim colunaBasePisLetra As String = ""
        Dim colunaPisLetra As String = ""
        Dim colunaBaseCofinsLetra As String = ""
        Dim colunaCofinsLetra As String = ""
        Dim colunaBaseIpiLetra As String = ""
        Dim colunaIpiLetra As String = ""
        Dim colunaBaseFethabLetra As String = ""
        Dim colunaFethabLetra As String = ""
        Dim colunaDescontoLetra As String = ""
        Dim colunaValorTotalLetra As String = ""

        Dim colunaQuantidade As String = "Quantidade"
        Dim colunaQuantidadeItemPedido As String = "QuantidadeItemPedido"
        Dim colunaValor As String = "Valor"
        Dim colunaFunrural As String = "Funrural"
        Dim colunaSenar As String = "Senar"
        Dim colunaBaseIcms As String = "BaseIcms"
        Dim colunaIcms As String = "Icms"
        Dim colunaBaseIcmsST As String = "BaseIcmsST"
        Dim colunaIcmsST As String = "IcmsST"
        Dim colunaBaseFECP As String = "BaseFECP"
        Dim colunaFECP As String = "FECP"
        Dim colunaBasePis As String = "BasePis"
        Dim colunaPis As String = "Pis"
        Dim colunaBaseCofins As String = "BaseCofins"
        Dim colunaCofins As String = "Cofins"
        Dim colunaBaseIpi As String = "BaseIpi"
        Dim colunaIpi As String = "Ipi"
        Dim colunaBaseFethab As String = "BaseFethab"
        Dim colunaFethab As String = "Fethab"
        Dim colunaDesconto As String = "Desconto"
        Dim colunaValorTotal As String = "ValorTotal"

        ' Procure o índice da coluna pelo nome no cabeçalho
        For colIndex As Integer = 1 To worksheet.Dimension.End.Column

            Dim valorCabecalho As String = worksheet.Cells(linhaCabecalho, colIndex).Text

            If valorCabecalho.Equals(colunaQuantidade, StringComparison.OrdinalIgnoreCase) Then

                'Quantidade
                ' Converta o índice para a letra correspondente
                colunaTotalLetra = IndexParaColuna(colIndex - 1)
                colunaFimDadosLetra = IndexParaColuna(columnIndex - 1)
                colunaQuantidadeLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaQuantidadeItemPedido, StringComparison.OrdinalIgnoreCase) Then

                'QuantidadeItemPedido
                colunaQuantidadeItemPedidoLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaValor, StringComparison.OrdinalIgnoreCase) Then

                'Valor
                colunaValorLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaFunrural, StringComparison.OrdinalIgnoreCase) Then

                'Funrural
                colunaFunruralLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaSenar, StringComparison.OrdinalIgnoreCase) Then

                'Senar
                colunaSenarLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBaseIcms, StringComparison.OrdinalIgnoreCase) Then

                'BaseIcms
                colunaBaseIcmsLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaIcms, StringComparison.OrdinalIgnoreCase) Then

                'Icms
                colunaIcmsLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBaseIcmsST, StringComparison.OrdinalIgnoreCase) Then

                'BaseIcmsST
                colunaBaseIcmsSTLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaIcmsST, StringComparison.OrdinalIgnoreCase) Then

                'IcmsST
                colunaIcmsSTLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBaseFECP, StringComparison.OrdinalIgnoreCase) Then

                'BaseFECP
                colunaBaseFECPLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaFECP, StringComparison.OrdinalIgnoreCase) Then

                'FECP
                colunaFECPLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBasePis, StringComparison.OrdinalIgnoreCase) Then

                'BasePis
                colunaBasePisLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaPis, StringComparison.OrdinalIgnoreCase) Then

                'Pis
                colunaPisLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBaseCofins, StringComparison.OrdinalIgnoreCase) Then

                'BaseCofins
                colunaBaseCofinsLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaCofins, StringComparison.OrdinalIgnoreCase) Then

                'Cofins
                colunaCofinsLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBaseIpi, StringComparison.OrdinalIgnoreCase) Then

                'BaseIpi
                colunaBaseIpiLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaIpi, StringComparison.OrdinalIgnoreCase) Then

                'Ipi
                colunaIpiLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaBaseFethab, StringComparison.OrdinalIgnoreCase) Then

                'BaseFethab
                colunaBaseFethabLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaFethab, StringComparison.OrdinalIgnoreCase) Then

                'Fethab
                colunaFethabLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaDesconto, StringComparison.OrdinalIgnoreCase) Then

                'Desconto
                colunaDescontoLetra = IndexParaColuna(colIndex)

            ElseIf valorCabecalho.Equals(colunaValorTotal, StringComparison.OrdinalIgnoreCase) Then

                'ValorTotal
                colunaValorTotalLetra = IndexParaColuna(colIndex)

            End If

        Next

        'Totalizadores

        '**************************************************************************************************************************
        'X -TOTAL
        'criando colunas de totalizadores
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaTotalLetra, rowIndex, colunaFimDadosLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaTotalLetra, rowIndex, colunaFimDadosLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaTotalLetra, rowIndex, colunaFimDadosLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaTotalLetra, rowIndex, colunaFimDadosLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaTotalLetra, rowIndex, colunaFimDadosLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(String.Format("{0}{1}", colunaTotalLetra, rowIndex)).Value = String.Format("{0}", "TOTAL")


        '**************************************************************************************************************************
        'Quantidade
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaQuantidadeLetra, linhaCabecalho + 1, colunaQuantidadeLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaQuantidadeLetra, linhaCabecalho + 1, colunaQuantidadeLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'QuantidadeItemPedido
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeItemPedidoLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeItemPedidoLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeItemPedidoLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeItemPedidoLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeItemPedidoLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaQuantidadeItemPedidoLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaQuantidadeItemPedidoLetra, linhaCabecalho + 1, colunaQuantidadeItemPedidoLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaQuantidadeItemPedidoLetra, linhaCabecalho + 1, colunaQuantidadeItemPedidoLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Valor
        worksheet.Cells(String.Format("{0}{1}", colunaValorLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaValorLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaValorLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaValorLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaValorLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaValorLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaValorLetra, linhaCabecalho + 1, colunaValorLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaValorLetra, linhaCabecalho + 1, colunaValorLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Funrural
        worksheet.Cells(String.Format("{0}{1}", colunaFunruralLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaFunruralLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaFunruralLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaFunruralLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaFunruralLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaFunruralLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaFunruralLetra, linhaCabecalho + 1, colunaFunruralLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaFunruralLetra, linhaCabecalho + 1, colunaFunruralLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Senar
        worksheet.Cells(String.Format("{0}{1}", colunaSenarLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaSenarLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaSenarLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaSenarLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaSenarLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaSenarLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaSenarLetra, linhaCabecalho + 1, colunaSenarLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaSenarLetra, linhaCabecalho + 1, colunaSenarLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BaseIcms
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBaseIcmsLetra, linhaCabecalho + 1, colunaBaseIcmsLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBaseIcmsLetra, linhaCabecalho + 1, colunaBaseIcmsLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Icms
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaIcmsLetra, linhaCabecalho + 1, colunaIcmsLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaIcmsLetra, linhaCabecalho + 1, colunaIcmsLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BaseIcmsST
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsSTLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsSTLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsSTLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsSTLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsSTLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIcmsSTLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBaseIcmsSTLetra, linhaCabecalho + 1, colunaBaseIcmsSTLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBaseIcmsSTLetra, linhaCabecalho + 1, colunaBaseIcmsSTLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'IcmsST
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsSTLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsSTLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsSTLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsSTLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsSTLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaIcmsSTLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaIcmsSTLetra, linhaCabecalho + 1, colunaIcmsSTLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaIcmsSTLetra, linhaCabecalho + 1, colunaIcmsSTLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BaseFECP
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFECPLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFECPLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFECPLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFECPLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFECPLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFECPLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBaseFECPLetra, linhaCabecalho + 1, colunaBaseFECPLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBaseFECPLetra, linhaCabecalho + 1, colunaBaseFECPLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'FECP
        worksheet.Cells(String.Format("{0}{1}", colunaFECPLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaFECPLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaFECPLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaFECPLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaFECPLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaFECPLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaFECPLetra, linhaCabecalho + 1, colunaFECPLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaFECPLetra, linhaCabecalho + 1, colunaFECPLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BasePis
        worksheet.Cells(String.Format("{0}{1}", colunaBasePisLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBasePisLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBasePisLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBasePisLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBasePisLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBasePisLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBasePisLetra, linhaCabecalho + 1, colunaBasePisLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBasePisLetra, linhaCabecalho + 1, colunaBasePisLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'PIS
        worksheet.Cells(String.Format("{0}{1}", colunaPisLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaPisLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaPisLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaPisLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaPisLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaPisLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaPisLetra, linhaCabecalho + 1, colunaPisLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaPisLetra, linhaCabecalho + 1, colunaPisLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BaseCofins
        worksheet.Cells(String.Format("{0}{1}", colunaBaseCofinsLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBaseCofinsLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBaseCofinsLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBaseCofinsLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBaseCofinsLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBaseCofinsLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBaseCofinsLetra, linhaCabecalho + 1, colunaBaseCofinsLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBaseCofinsLetra, linhaCabecalho + 1, colunaBaseCofinsLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Cofins
        worksheet.Cells(String.Format("{0}{1}", colunaCofinsLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaCofinsLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaCofinsLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaCofinsLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaCofinsLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaCofinsLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaCofinsLetra, linhaCabecalho + 1, colunaCofinsLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaCofinsLetra, linhaCabecalho + 1, colunaCofinsLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BaseIPI
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIpiLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIpiLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIpiLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIpiLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIpiLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBaseIpiLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBaseIpiLetra, linhaCabecalho + 1, colunaBaseIpiLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBaseIpiLetra, linhaCabecalho + 1, colunaBaseIpiLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Ipi
        worksheet.Cells(String.Format("{0}{1}", colunaIpiLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaIpiLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaIpiLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaIpiLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaIpiLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaIpiLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaIpiLetra, linhaCabecalho + 1, colunaIpiLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaIpiLetra, linhaCabecalho + 1, colunaIpiLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'BaseFethab
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFethabLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFethabLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFethabLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFethabLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFethabLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaBaseFethabLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaBaseFethabLetra, linhaCabecalho + 1, colunaBaseFethabLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaBaseFethabLetra, linhaCabecalho + 1, colunaBaseFethabLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Fethab
        worksheet.Cells(String.Format("{0}{1}", colunaFethabLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaFethabLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaFethabLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaFethabLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaFethabLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaFethabLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaFethabLetra, linhaCabecalho + 1, colunaFethabLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaFethabLetra, linhaCabecalho + 1, colunaFethabLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Desconto
        worksheet.Cells(String.Format("{0}{1}", colunaDescontoLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaDescontoLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaDescontoLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaDescontoLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaDescontoLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaDescontoLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaDescontoLetra, linhaCabecalho + 1, colunaDescontoLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaDescontoLetra, linhaCabecalho + 1, colunaDescontoLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

        'Valor
        worksheet.Cells(String.Format("{0}{1}", colunaValorTotalLetra, rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("{0}{1}", colunaValorTotalLetra, rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("{0}{1}", colunaValorTotalLetra, rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("{0}{1}", colunaValorTotalLetra, rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("{0}{1}", colunaValorTotalLetra, rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("{0}{1}", colunaValorTotalLetra, rowIndex)).Formula = String.Format("=SUM({0}{1}:{2}{3})", colunaValorTotalLetra, linhaCabecalho + 1, colunaValorTotalLetra, rowIndex - 1)
        worksheet.Cells(String.Format("{0}{1}:{2}{3}", colunaValorTotalLetra, linhaCabecalho + 1, colunaValorTotalLetra, rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

        'setando autofit nas células da planilha
        worksheet.Cells.AutoFitColumns(0)

        'congelando quinta linha (cabeçalho)
        worksheet.View.FreezePanes(6, 1)

    End Sub

    Private Sub PlanilhaConsistenciaDeNotasConsolidado(objEmpresa As Cliente, ds As DataSet, package As ExcelPackage)

        'criando planilha títulos
        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("CONSOLIDADO")

        'criando linha com o cabeçalho da planilha
        Dim rowIndex As Integer = 1
        Dim columnIndex As Integer = 1

        'criando linha que informa o nome da empresa e o cnpj
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa a cidade e o estado da empresa
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa o título do relatório
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "CONSOLIDADO")
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa o período selecionado na página
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)) & " - " & IIf(RadEntradas.Checked, "ENTRADAS", "SAÍDAS"))
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha com o cabeçalho da planilha
        For Each col As DataColumn In ds.Tables(0).Columns
            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
            columnIndex += 1
        Next

        'criando auto filtro na planilha
        worksheet.Cells("A5:O" & rowIndex).AutoFilter = True

        'aplicando formatação nas células do cabeçalho
        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
            range.Style.Font.Bold = True
            range.Style.Fill.PatternType = ExcelFillStyle.Solid
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
            range.Style.Font.Color.SetColor(Color.White)
        End Using
        rowIndex += 1

        ' criando conteúdo da planilha com os dados do dataset
        For Each row As DataRow In ds.Tables(0).Rows
            columnIndex = 1
            For Each col As DataColumn In ds.Tables(0).Columns

                If col.ColumnName = "Quantidade" OrElse
                    col.ColumnName = "QuantidadeItemPedido" OrElse
                    col.ColumnName = "Unitario" OrElse
                    col.ColumnName = "Valor" OrElse
                    col.ColumnName = "Funrural" OrElse
                    col.ColumnName = "Senar" OrElse
                    col.ColumnName = "BaseIcms" OrElse
                    col.ColumnName = "Icms" OrElse
                    col.ColumnName = "BaseIcmsST" OrElse
                    col.ColumnName = "IcmsST" OrElse
                    col.ColumnName = "BaseFECP" OrElse
                    col.ColumnName = "FECP" OrElse
                    col.ColumnName = "BasePis" OrElse
                    col.ColumnName = "Pis" OrElse
                    col.ColumnName = "BaseCofins" OrElse
                    col.ColumnName = "Cofins" OrElse
                    col.ColumnName = "BaseIpi" OrElse
                    col.ColumnName = "Ipi" OrElse
                    col.ColumnName = "BaseFethab" OrElse
                    col.ColumnName = "Fethab" OrElse
                    col.ColumnName = "Desconto" Then
                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToDecimal(row(col.ColumnName))
                Else
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                End If

                columnIndex += 1
            Next

            'aplicando formatação nas células do conteúdo
            If rowIndex Mod 2 = 0 Then
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                End Using
            End If
            rowIndex += 1
        Next

        '**************************************************************************************************************************
        'Q -TOTAL
        'criando colunas de totalizadores
        worksheet.Cells(String.Format("Q{0}:AG{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("Q{0}:AG{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("Q{0}:AG{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("Q{0}:AG{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("Q{0}:AG{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(String.Format("Q{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")


        '**************************************************************************************************************************
        'R -Quantidade
        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("R{0}", rowIndex)).Formula = String.Format("=SUM(R6:R{0})", rowIndex - 1)
        worksheet.Cells(String.Format("R6:R{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'S -QuantidadeItemPedido
        worksheet.Cells(String.Format("S{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("S{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("S{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("S{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("S{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("S{0}", rowIndex)).Formula = String.Format("=SUM(S6:S{0})", rowIndex - 1)
        worksheet.Cells(String.Format("S6:S{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'T -Valor
        worksheet.Cells(String.Format("T{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("T{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("T{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("T{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("T{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("T{0}", rowIndex)).Formula = String.Format("=SUM(T6:T{0})", rowIndex - 1)
        worksheet.Cells(String.Format("T6:T{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'U -Funrural
        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("U{0}", rowIndex)).Formula = String.Format("=SUM(U6:U{0})", rowIndex - 1)
        worksheet.Cells(String.Format("U6:U{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'V -Senar
        worksheet.Cells(String.Format("V{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("V{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("V{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("V{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("V{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("V{0}", rowIndex)).Formula = String.Format("=SUM(V6:V{0})", rowIndex - 1)
        worksheet.Cells(String.Format("V6:V{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'W -BaseIcms
        worksheet.Cells(String.Format("W{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("W{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("W{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("W{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("W{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("W{0}", rowIndex)).Formula = String.Format("=SUM(W6:W{0})", rowIndex - 1)
        worksheet.Cells(String.Format("W6:W{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'X -Icms
        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("X{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("X{0}", rowIndex)).Formula = String.Format("=SUM(X6:X{0})", rowIndex - 1)
        worksheet.Cells(String.Format("X6:X{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Y -BaseIcmsST
        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("Y{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("Y{0}", rowIndex)).Formula = String.Format("=SUM(Y6:Y{0})", rowIndex - 1)
        worksheet.Cells(String.Format("Y6:Y{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'Z -IcmsST
        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("Z{0}", rowIndex)).Formula = String.Format("=SUM(Z6:Z{0})", rowIndex - 1)
        worksheet.Cells(String.Format("Z6:Z{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AA -BasePis
        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AA{0}", rowIndex)).Formula = String.Format("=SUM(AA6:AA{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AA6:AA{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AB -PIS
        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AB{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AB{0}", rowIndex)).Formula = String.Format("=SUM(AB6:AB{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AB6:AB{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AC -BaseCofins
        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AC{0}", rowIndex)).Formula = String.Format("=SUM(AC6:AC{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AC6:AC{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AD -Cofins
        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AD{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AD{0}", rowIndex)).Formula = String.Format("=SUM(AD6:AD{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AD6:AD{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AE -BaseIPI
        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AE{0}", rowIndex)).Formula = String.Format("=SUM(AE6:AE{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AE6:AE{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AF -Ipi
        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AF{0}", rowIndex)).Formula = String.Format("=SUM(AF6:AF{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AF6:AF{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"


        '**************************************************************************************************************************
        'AG -Desconto
        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("AG{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("AG{0}", rowIndex)).Formula = String.Format("=SUM(AG6:AG{0})", rowIndex - 1)
        worksheet.Cells(String.Format("AG6:AG{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

        'formatando células numéricas
        'worksheet.Cells(String.Format("P6:X{0}", rowIndex)).Style.Numberformat.Format = "0.#########0_ ;[Red]-0.#########0"
        'worksheet.Cells(String.Format("Q6:Q{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("R6:R{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("S6:S{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("T6:T{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("U6:U{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("V6:V{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("W6:W{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
        'worksheet.Cells(String.Format("X6:X{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

        'setando autofit nas células da planilha
        worksheet.Cells.AutoFitColumns(0)

        'congelando quinta linha (cabeçalho)
        worksheet.View.FreezePanes(6, 1)
    End Sub

    ' Função para converter índice em letra da coluna
    Private Function IndexParaColuna(index As Integer) As String
        Dim coluna As String = ""
        While index > 0
            Dim resto As Integer = (index - 1) Mod 26
            coluna = Chr(65 + resto) & coluna
            index = (index - resto) \ 26
        End While
        Return coluna
    End Function

    Private Sub GerarExcelPorPlaca(ByVal periodo As String, ByVal objEmpresa As [Lib].Negocio.Cliente)
        Try
            Dim ds As DataSet = getDataSetPorPlaca(objEmpresa)
            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("CONSISTÊNCIA DE NOTAS")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "CONSISTÊNCIA DE NOTAS")
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células numéricas
                        worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Formula = String.Format("=SUM(L6:L{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("ConsistenciaDeNotas", "RELATORIO") Then
                If Validar() Then
                    Dim Periodo As String = "PERÍODO: " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)) & " - " & IIf(RadEntradas.Checked, "ENTRADAS", "SAÍDAS")

                    Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

                    If Pdf Then
                        Dim ds As DataSet = getDataSet(objEmpresa, False)
                        Dim parameters = New Dictionary(Of String, Object)()
                        parameters.Add("Parametros", getParam())

                        Funcoes.BindReport(Me.Page, ds, "Cr_ConsistenciaDeNotas", eExportType.PDF, parameters)
                    Else
                        GerarExcel(Periodo, objEmpresa)
                    End If
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir o relatorio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelPlaca_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelPlaca.Click
        Try
            If Funcoes.VerificaPermissao("ConsistenciaDeNotas", "RELATORIO") Then
                If Validar() Then
                    Dim Periodo As String = "PERÍODO: " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)) & " - " & IIf(RadEntradas.Checked, "ENTRADAS", "SAÍDAS")

                    Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

                    GerarExcelPorPlaca(Periodo, objEmpresa)

                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir o relatorio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConsistenciaDeNotas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class