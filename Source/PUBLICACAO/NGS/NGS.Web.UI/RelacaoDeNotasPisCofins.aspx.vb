Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelacaoDeNotasPisCofins
    Inherits BasePage
    Dim Sql As String = ""
    Dim Empresa() As String
    Dim PeriodoInicial As String = ""
    Dim PeriodoFinal As String = ""
    Dim EmpresaMestre As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelacaoDeNotasPisCofins", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("PlanilhasExcel")) Then
                        Directory.CreateDirectory(Server.MapPath("PlanilhasExcel"))
                    End If
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    CargaUnidade()
                    VerificaUnidade()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf & _
              "FROM Clientes C " & vbCrLf & _
              "INNER JOIN ClientesXTipos CT " & vbCrLf & _
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf & _
              "WHERE CT.Tipo_Id = 050 " & vbCrLf & _
              "ORDER BY Nome" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
            Next
        End If

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                DdlUnidade.SelectedValue = Dr("AcessoUnidade")
                CargaEmpresas()
                DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            Next
        End If
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()
        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
                Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
                Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
                Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
                Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
                Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
                DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
            Next
        End If

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function Validar()
        Dim ok As Boolean = True

        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a unidade de Negócio.")
            ok = False
        End If
        If DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a empresa.")
            ok = False
        End If
        If txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Informe o período inicial.")
            ok = False
        End If
        If txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Informe o período final.")
            ok = False
        End If

        Return ok
    End Function

    Private Sub Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getDataSet() As DataSet
        Dim ds As DataSet

        Empresa = DdlEmpresa.SelectedValue.Split("-")
        EmpresaMestre = Left(Empresa(0), 8)

        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy-MM-dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy-MM-dd")

        Sql = "SELECT Empresa.Reduzido AS RedEmpresa, NF.Empresa_Id AS CNPJEmpresa, NF.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa, Empresa.Cidade AS CidadeEmpresa, Empresa.Estado AS UFEmpresa," & vbCrLf &
              "	      NF.Cliente_Id  AS Cliente, NF.EndCliente_Id AS EndCliente, Clientes.Inscricao, Clientes.Nome AS NomeCliente, Clientes.Cidade AS CidadeCliente, Clientes.Estado AS UFCliente," & vbCrLf &
              "       isnull(NF.LocalEmbarque,'') AS LocalEmbarque, isnull(NF.EndLocalEmbarque,0) AS EndLocalEmbarque, isnull(CliEMB.Nome,'') AS NomeLocalEmbarque, isnull(CliEMB.Cidade,'') AS CidadeLocalEmbarque, isnull(CliEMB.Estado,'') AS UFLocalEmbarque," & vbCrLf &
              "	      Produtos.Grupo, nfi.Produto_Id AS Produto, isnull(Produtos.CodigoProdutoTerceiro,'') AS CodigoProdutoTerceiro, Produtos.NCM, Produtos.Nome AS NomeProduto," & vbCrLf &
              "	      NF.EntradaSaida_Id AS E_S, NF.Serie_Id  AS Serie, NF.Nota_Id AS Nota, isnull(NF.InvoiceNavio,'') AS InvoiceNavio, NV.Descricao AS Navio," & vbCrLf &
              "	      NF.Pedido, P.PedidoEfetivo," & vbCrLf &
              "	      ISNULL(RxP.Pesagem_Id, 0) AS NumeroTicket," & vbCrLf &
              "	      Case" & vbCrLf &
              "	        when ISNULL(nfi.PesoFiscal,0) = 0" & vbCrLf &
              "           then nfi.QuantidadeFisica" & vbCrLf &
              "           else nfi.PesoFiscal" & vbCrLf &
              "       end as PesoFiscal," & vbCrLf &
              "       Case" & vbCrLf &
              "	        when NF.EntradaSaida_Id = 'E'" & vbCrLf &
              "	         then nfi.QuantidadeFisica" & vbCrLf &
              "	         else ISNULL(NxD.PesoLiquido,0)" & vbCrLf &
              "	      end as PesoChegada," & vbCrLf &
              "       NF.Movimento, NF.DataDaNota, NF.NossaEmissao," & vbCrLf &
              "	      OE.CodigoFiscal AS CFOP, Cfop.Descricao AS Nome_CFOP," & vbCrLf &
              "	      isnull(OE.STICMS, 0) AS CST_ICMS," & vbCrLf &
              "	      isnull(OE.STPISCOFINS, 0) AS CST_PIS, OE.ObsPISCOFINS as ObsPisCofins," & vbCrLf &
              "	      isnull(OE.STIPI, 0) AS CST_IPI," & vbCrLf &
              "	      nfi.OperacaoXEstado AS OxE," & vbCrLf &
              "	      NF.Operacao, NF.SubOperacao,	so.Descricao," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) in ('FRETES', 'DESP.ADUANEIRAS')            THEN nfe.Valor      ELSE 0 END) + nfi.Valor AS ValorDoItem," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR') THEN nfe.Base       ELSE 0 END) AS BaseFUNRURAL," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR') THEN nfe.Percentual ELSE 0 END) AS AliquotaFUNRURAL," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR') THEN nfe.Valor      ELSE 0 END) AS FUNRURAL," & vbCrLf &
              "       SUM(CASE WHEN nfe.Encargo_Id = 'PIS CRED PRE'    THEN nfe.Base       ELSE 0 END) AS BasePisPre," & vbCrLf &
              "       SUM(CASE WHEN nfe.Encargo_Id = 'PIS CRED PRE'    THEN nfe.Percentual ELSE 0 END) AS AliquotaPisPre," & vbCrLf &
              "       SUM(CASE WHEN nfe.Encargo_Id = 'PIS CRED PRE'    THEN nfe.Valor      ELSE 0 END) AS PisPre," & vbCrLf &
              "       SUM(CASE WHEN nfe.Encargo_Id = 'COFINS CRED PRE' THEN nfe.Base       ELSE 0 END) AS BaseCofinsPre," & vbCrLf &
              "       SUM(CASE WHEN nfe.Encargo_Id = 'COFINS CRED PRE' THEN nfe.Percentual ELSE 0 END) AS AliquotaCofinsPre," & vbCrLf &
              "       SUM(CASE WHEN nfe.Encargo_Id = 'COFINS CRED PRE' THEN nfe.Valor      ELSE 0 END) AS CofinsPre," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS S/SERV.PJ'  THEN nfe.Base       ELSE 0 END) AS BaseInssServPJ," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS S/SERV.PJ'  THEN nfe.Percentual ELSE 0 END) AS AliquotaInssServPj," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS S/SERV.PJ'  THEN nfe.Valor      ELSE 0 END) AS InssServPj," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS S/SERV.PF'  THEN nfe.Base       ELSE 0 END) AS BaseInssServPf," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS S/SERV.PF'  THEN nfe.Percentual ELSE 0 END) AS AliquotaInssServPf," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS S/SERV.PF'  THEN nfe.Valor ELSE 0 END) AS InssServPf," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'    THEN nfe.Base       ELSE 0 END) AS BaseICMS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'    THEN nfe.Percentual ELSE 0 END) AS AliquotaICMS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'    THEN nfe.Valor      ELSE 0 END) AS ICMS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'     THEN nfe.Base       ELSE 0 END) AS BaseIPI," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'     THEN nfe.Percentual ELSE 0 END) AS AliquotaIPI," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'     THEN nfe.Valor      ELSE 0 END) AS IPI," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'FETHAB'  THEN nfe.Base       ELSE 0 END) AS BaseFETHAB," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'FETHAB'  THEN nfe.Percentual ELSE 0 END) AS AliquotaFETHAB," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'FETHAB'  THEN nfe.Valor      ELSE 0 END) AS FETHAB," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('PIS','PIS RECUP.') THEN nfe.Base       ELSE 0 END) AS BasePIS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('PIS','PIS RECUP.') THEN nfe.Percentual ELSE 0 END) AS AliquotaPIS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('PIS','PIS RECUP.') THEN nfe.Valor      ELSE 0 END) AS PIS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('COFINS','COFINS RECUP.') THEN nfe.Base       ELSE 0 END) AS BaseCOFINS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('COFINS','COFINS RECUP.') THEN nfe.Percentual ELSE 0 END) AS AliquotaCOFINS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) IN ('COFINS','COFINS RECUP.') THEN nfe.Valor      ELSE 0 END) AS COFINS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'CSRF'    THEN nfe.Base       ELSE 0 END) AS BaseCSRF," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'CSRF'    THEN nfe.Percentual ELSE 0 END) AS AliquotaCSRF," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'CSRF'    THEN nfe.Valor      ELSE 0 END) AS CSRF," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS'    THEN nfe.Base       ELSE 0 END) AS BaseINSS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS'    THEN nfe.Percentual ELSE 0 END) AS AliquotaINSS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'INSS'    THEN nfe.Valor      ELSE 0 END) AS INSS," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IRRF PF' THEN nfe.Base       ELSE 0 END) AS BaseIrrfPf," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IRRF PF' THEN nfe.Percentual ELSE 0 END) AS AliquotaIrrfPf," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IRRF PF' THEN nfe.Valor      ELSE 0 END) AS IrrfPf," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IRRF PJ' THEN nfe.Base       ELSE 0 END) AS BaseIrrfPJ," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IRRF PJ' THEN nfe.Percentual ELSE 0 END) AS AliquotaIrrfPJ," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'IRRF PJ' THEN nfe.Valor      ELSE 0 END) AS IrrfPJ," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ISS PJ'  THEN nfe.Base       ELSE 0 END) AS BaseIssPj," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ISS PJ'  THEN nfe.Percentual ELSE 0 END) AS AliquotaIssPJ," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ISS PJ'  THEN nfe.Valor      ELSE 0 END) AS IssPJ," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ISS'     THEN nfe.Base       ELSE 0 END) AS IssBase," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ISS'     THEN nfe.Percentual ELSE 0 END) AS AliquotaIss," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'ISS'     THEN nfe.Valor      ELSE 0 END) AS Iss," & vbCrLf &
              "       SUM(CASE WHEN isnull(E.EncargoAgrupador,nfe.Encargo_Id) = 'LIQUIDO' THEN nfe.Valor      ELSE 0 END) AS VALORLIQUIDO," & vbCrLf &
              "       isnull(NFeR.ChaveNfe,'') as ChaveNFE, (CONVERT(varchar,tpDoc.Codigo_Id) + '-' + tpDoc.Descricao) as Tipo" & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
              "    ON nfi.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND nfi.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND nfi.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND nfi.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND nfi.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND nfi.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND nfi.Nota_Id         = NF.Nota_Id" & vbCrLf &
              " INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
              "    ON nfe.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
              "   AND nfe.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
              "   AND nfe.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
              "   AND nfe.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
              "   AND nfe.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
              "   AND nfe.Serie_Id        = nfi.Serie_Id" & vbCrLf &
              "   AND nfe.Nota_Id         = nfi.Nota_Id" & vbCrLf &
              "   AND nfe.Produto_Id      = nfi.Produto_Id" & vbCrLf &
              "   AND nfe.CFOP_Id         = nfi.CFOP_Id" & vbCrLf &
              "   AND nfe.Sequencia_id    = nfi.Sequencia_Id" & vbCrLf &
              " INNER JOIN	TipoDeDocumento tpDoc" & vbCrLf &
              "    ON tpDoc.Codigo_Id    = NF.TipoDeDocumento" & vbCrLf &
              " INNER JOIN	Clientes AS Empresa" & vbCrLf &
              "    ON Empresa.Cliente_Id  = NF.Empresa_Id" & vbCrLf &
              "   AND Empresa.Endereco_Id = NF.EndEmpresa_Id" & vbCrLf &
              " INNER JOIN	Clientes AS Clientes" & vbCrLf &
              "    ON Clientes.Cliente_Id  = NF.Cliente_Id" & vbCrLf &
              "   AND Clientes.Endereco_Id = NF.EndCliente_Id" & vbCrLf &
              " Inner Join OperacaoXEstado OE" & vbCrLf &
              "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
              " INNER JOIN Produtos" & vbCrLf &
              "    ON Produtos.Produto_Id = nfi.Produto_Id" & vbCrLf &
              " INNER JOIN SubOperacoes so" & vbCrLf &
              "    ON so.Operacao_Id     = NF.Operacao" & vbCrLf &
              "   AND so.SubOperacoes_Id = NF.SubOperacao" & vbCrLf &
              " INNER JOIN Cfop" & vbCrLf &
              "    ON Cfop.Cfop_Id = OE.CodigoFiscal" & vbCrLf &
              "  Left Join Encargos E" & vbCrLf &
              "    on E.Encargo_id = nfe.Encargo_Id" & vbCrLf &
              "   and isnull(E.EncargoAgrupador,'') <> ''" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXRomaneios AS NFR" & vbCrLf &
              "    ON NFR.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND NFR.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND NFR.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND NFR.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND NFR.Nota_Id         = NF.Nota_Id" & vbCrLf &
              "  LEFT JOIN Romaneios Ro" & vbCrLf &
              "    ON Ro.Empresa_Id    = NFR.Empresa_Id" & vbCrLf &
              "   AND Ro.EndEmpresa_Id = NFR.EndEmpresa_Id" & vbCrLf &
              "   AND Ro.Romaneio_Id   = NFR.Romaneio_Id" & vbCrLf &
              "  LEFT JOIN RomaneiosXPesagens RxP" & vbCrLf &
              "    ON RxP.Empresa_Id    = Ro.Empresa_Id" & vbCrLf &
              "   AND RxP.EndEmpresa_Id = Ro.EndEmpresa_Id" & vbCrLf &
              "   AND RxP.Romaneio_Id   = Ro.Romaneio_Id" & vbCrLf &
              "  LEFT JOIN NotasXDestinos NxD" & vbCrLf &
              "    ON NxD.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
              "   AND NxD.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
              "   AND NxD.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
              "   AND NxD.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
              "   AND NxD.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
              "   AND NxD.Serie_Id        = nfi.Serie_Id" & vbCrLf &
              "   AND NxD.Nota_Id         = nfi.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfeRealizadas AS NFeR" & vbCrLf &
              "    ON NFeR.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND NFeR.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND NFeR.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND NFeR.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND NFeR.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND NFeR.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND NFeR.Nota_Id         = NF.Nota_Id" & vbCrLf &
              " LEFT JOIN	Pedidos P" & vbCrLf &
              "    ON P.Pedido_Id = NF.Pedido" & vbCrLf &
              "   AND P.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND P.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              " LEFT JOIN	Clientes AS CliEMB" & vbCrLf &
              "    ON CliEMB.Cliente_Id  = NF.LocalEmbarque" & vbCrLf &
              "   AND CliEMB.Endereco_Id = NF.EndLocalEmbarque" & vbCrLf &
              "  LEFT JOIN NaviosXInvoice AS NavXInv" & vbCrLf &
              "    ON NavXInv.Codigo_Id = NF.InvoiceNavio" & vbCrLf &
              "  LEFT JOIN Navios AS NV" & vbCrLf &
              "    ON NV.Codigo_Id = NavXInv.Navio_Id" & vbCrLf

        If chkConsolidado.Checked = True Then
            Sql &= " Where NF.Empresa_Id LIKE '" & EmpresaMestre & "%'" & vbCrLf &
                   "   AND NF.EndEmpresa_id = " & Empresa(1) & vbCrLf
        Else
            Sql &= " Where NF.Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf &
                   "   AND NF.EndEmpresa_id = " & Empresa(1) & vbCrLf
        End If

        Sql &= "   AND NF.Movimento between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
              "   AND NF.Situacao = 1" & vbCrLf

        'PARA RTGRÃOS NÃO DEVE CONSIDERAR ESSAS SÉRIES - FURLAN - 08/08/2024
        If Left(Session("ssEmpresa"), 8) = "24450490" Then
            Sql &= "   AND NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf
        End If


        Sql &= " Group By Empresa.Reduzido, NF.Empresa_Id, NF.EndEmpresa_Id, Empresa.Nome, Empresa.Cidade, Empresa.Estado," & vbCrLf &
               "	      NF.Cliente_Id, NF.EndCliente_Id, Clientes.Inscricao, Clientes.Nome, Clientes.Cidade, Clientes.Estado," & vbCrLf &
               "	      NF.LocalEmbarque, NF.EndLocalEmbarque, CliEMB.Nome, CliEMB.Cidade, CliEMB.Estado," & vbCrLf &
               "	      Produtos.Grupo, nfi.Produto_Id, Produtos.CodigoProdutoTerceiro, Produtos.NCM, Produtos.Nome," & vbCrLf &
               "	      NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id, NF.InvoiceNavio, NV.Descricao, nfi.OperacaoXEstado," & vbCrLf &
               "	      NF.Pedido, P.PedidoEfetivo, ISNULL(RxP.Pesagem_Id, 0)," & vbCrLf &
               "	      Case when ISNULL(nfi.PesoFiscal,0) = 0 then nfi.QuantidadeFisica else nfi.PesoFiscal end," & vbCrLf &
               "          Case when NF.EntradaSaida_Id = 'E' then nfi.QuantidadeFisica else ISNULL(NxD.PesoLiquido,0) end," & vbCrLf &
               "          NF.Movimento, NF.DataDaNota, NF.NossaEmissao," & vbCrLf &
               "	      OE.CodigoFiscal, Cfop.Descricao," & vbCrLf &
               "	      isnull(OE.STICMS, 0), isnull(OE.STPISCOFINS, 0), OE.ObsPISCOFINS, isnull(OE.STIPI, 0)," & vbCrLf &
               "	      NF.Operacao, NF.SubOperacao, so.Descricao, nfi.Valor, NFeR.ChaveNfe, tpDoc.Codigo_Id, tpDoc.Descricao" & vbCrLf &
               " Order By NF.EntradaSaida_Id, NF.Movimento, NF.Nota_Id" & vbCrLf


        ds = Banco.ConsultaDataSet(Sql, "Notas")

        Return ds
    End Function

    Private Sub Processar()
        Try
            If Funcoes.VerificaPermissao("RelacaoDeNotasPisCofins", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim colunas As New Dictionary(Of String, eTipoCampo)

                colunas.Add("PesoFiscal", eTipoCampo.Numerico)
                colunas.Add("PesoChegada", eTipoCampo.Numerico)
                colunas.Add("Movimento", eTipoCampo.Data)
                colunas.Add("DataDaNota", eTipoCampo.Data)
                colunas.Add("ValorDoItem", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseICMS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaICMS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("ICMS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseIPI", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaIPI", eTipoCampo.ValorSemTotalizador)
                colunas.Add("IPI", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseFUNRURAL", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaFUNRURAL", eTipoCampo.ValorSemTotalizador)
                colunas.Add("FUNRURAL", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseFETHAB", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaFETHAB", eTipoCampo.ValorSemTotalizador)
                colunas.Add("FETHAB", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BasePIS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaPIS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("PIS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseCOFINS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaCOFINS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("COFINS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BasePisPre", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaPisPre", eTipoCampo.ValorSemTotalizador)
                colunas.Add("PisPre", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseCofinsPre", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaCofinsPre", eTipoCampo.ValorSemTotalizador)
                colunas.Add("CofinsPre", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseInssServPJ", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaInssServPj", eTipoCampo.ValorSemTotalizador)
                colunas.Add("InssServPj", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseInssServPf", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaInssServPf", eTipoCampo.ValorSemTotalizador)
                colunas.Add("InssServPf", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseCSRF", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaCSRF", eTipoCampo.ValorSemTotalizador)
                colunas.Add("CSRF", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseINSS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaINSS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("INSS", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseIrrfPf", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaIrrfPf", eTipoCampo.ValorSemTotalizador)
                colunas.Add("IrrfPf", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseIrrfPJ", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaIrrfPJ", eTipoCampo.ValorSemTotalizador)
                colunas.Add("IrrfPJ", eTipoCampo.ValorSemTotalizador)
                colunas.Add("BaseIssPj", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaIssPJ", eTipoCampo.ValorSemTotalizador)
                colunas.Add("IssPJ", eTipoCampo.ValorSemTotalizador)
                colunas.Add("IssBase", eTipoCampo.ValorSemTotalizador)
                colunas.Add("AliquotaIss", eTipoCampo.ValorSemTotalizador)
                colunas.Add("Iss", eTipoCampo.ValorSemTotalizador)
                colunas.Add("VALORLIQUIDO", eTipoCampo.ValorSemTotalizador)

                Funcoes.BindExcelOffice(Me.Page, ds, "RELAÇÃO DE NOTAS PIS COFINS", colunas)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para processar registro.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Try
            Processar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelacaoDeNotasPisCofins")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class