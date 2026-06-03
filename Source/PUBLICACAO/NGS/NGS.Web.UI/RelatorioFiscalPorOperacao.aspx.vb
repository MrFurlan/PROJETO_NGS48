Imports NGS.Lib.Negocio
Public Class RelatorioFiscalPorOperacao
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioFiscalPorOperacao", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, [Lib].Negocio.CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
                    ddl.Carregar(ddlClasse, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1")
                    txtMes.Text = Now.Month()
                    txtAno.Text = Now.Year()
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioFiscalPorOperacao", "Relatorio") Then

            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioFiscalPorOperacao", "Relatorio") Then

            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                Exit Sub
            End If
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
            Funcoes.Ajuda(Me.Page, "RelatorioFiscalPorOperacao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaCliente_Click(sender As Object, e As EventArgs) Handles btnBuscaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        ddlClasse.SelectedValue = String.Empty
        ddlSafra.SelectedValue = String.Empty
        txtMes.Text = Now.Month()
        txtAno.Text = Now.Year()
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "   SELECT E.Reduzido AS EmpresaReduzido,                                                                                                                 " & vbCrLf & _
                            "          dbo.FormatarCpfCnpj(NF.Empresa_Id) as Empresa_Id,                                                                                              " & vbCrLf & _
                            "          NF.EndEmpresa_Id,                                                                                                                              " & vbCrLf & _
                            "          E.Nome AS EmpresaNome,                                                                                                                         " & vbCrLf & _
                            "          E.Cidade AS EmpresaCidade,                                                                                                                     " & vbCrLf & _
                            "          E.Estado AS EmpresaEstado,                                                                                                                     " & vbCrLf & _
                            "          NF.EntradaSaida_Id,                                                                                                                            " & vbCrLf & _
                            "   	   P.Grupo,                                                                                                                                       " & vbCrLf & _
                            "          NFI.Produto_Id AS Produto,                                                                                                                     " & vbCrLf & _
                            "          P.Nome AS NomeProduto,                                                                                                                         " & vbCrLf & _
                            "          NF.Operacao,                                                                                                                                   " & vbCrLf & _
                            "          NF.SubOperacao,                                                                                                                                " & vbCrLf & _
                            "          dbo.FormatarCpfCnpj(NF.Cliente_Id) as Cliente_Id,                                                                                              " & vbCrLf & _
                            "          NF.EndCliente_Id,                                                                                                                              " & vbCrLf & _
                            "          ISNULL(CLI.Nome, '') AS NomeCliente,                                                                                                           " & vbCrLf & _
                            "          CLI.Cidade AS CidadeCliente,                                                                                                                   " & vbCrLf & _
                            "          CLI.Estado AS EstadoCliente,                                                                                                                   " & vbCrLf & _
                            "   	   SUM(nfi.QuantidadeFiscal) PesoFiscal,                                                                                                          " & vbCrLf & _
                            "          SUM(NFI.Valor) as Valor,                                                                                                                       " & vbCrLf & _
                            "                                                                                                                                                         " & vbCrLf & _
                            "          SUM(COALESCE (Encargos.Encargo1, 0)) AS ValorEncargo1,                                                                                         " & vbCrLf & _
                            "          SUM(COALESCE (Encargos.Encargo2, 0)) AS ValorEncargo2,                                                                                         " & vbCrLf & _
                            "          SUM(COALESCE (Encargos.Encargo3, 0)) AS ValorEncargo3,                                                                                         " & vbCrLf & _
                            "          SUM(COALESCE (Encargos.Encargo4, 0)) AS ValorEncargo4,                                                                                         " & vbCrLf & _
                            "          SUM(COALESCE (Encargos.Encargo5, 0)) AS ValorEncargo5                                                                                          " & vbCrLf & _
                            "     FROM NotasFiscais AS NF WITH (NOLOCK)                                                                                                               " & vbCrLf & _
                            "    INNER JOIN NotasFiscaisXItens AS NFI WITH (NOLOCK)                                                                                                   " & vbCrLf & _
                            "       ON NF.Empresa_Id      = NFI.Empresa_Id                                                                                                            " & vbCrLf & _
                            "      AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id                                                                                                         " & vbCrLf & _
                            "      AND NF.Cliente_Id      = NFI.Cliente_Id                                                                                                            " & vbCrLf & _
                            "      AND NF.EndCliente_Id   = NFI.EndCliente_Id                                                                                                         " & vbCrLf & _
                            "      AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id                                                                                                       " & vbCrLf & _
                            "      AND NF.Serie_Id        = NFI.Serie_Id                                                                                                              " & vbCrLf & _
                            "      AND NF.Nota_Id         = NFI.Nota_Id                                                                                                               " & vbCrLf & _
                            "     LEFT Join NFERealizadas NFER WITH (NOLOCK)                                                                                                          " & vbCrLf & _
                            "       ON NF.Empresa_Id      = NFER.Empresa_Id                                                                                                           " & vbCrLf & _
                            "      AND NF.EndEmpresa_Id   = NFER.EndEmpresa_Id                                                                                                        " & vbCrLf & _
                            "      AND NF.Cliente_Id      = NFER.Cliente_Id                                                                                                           " & vbCrLf & _
                            "      AND NF.EndCliente_Id   = NFER.EndCliente_Id                                                                                                        " & vbCrLf & _
                            "      AND NF.EntradaSaida_Id = NFER.EntradaSaida_Id                                                                                                      " & vbCrLf & _
                            "      AND NF.Serie_Id        = NFER.Serie_Id                                                                                                             " & vbCrLf & _
                            "      AND NF.Nota_Id         = NFER.Nota_Id                                                                                                              " & vbCrLf & _
                            "    INNER JOIN Produtos AS P WITH (NOLOCK)                                                                                                               " & vbCrLf & _
                            "       ON P.Produto_Id = NFI.Produto_Id                                                                                                                  " & vbCrLf & _
                            "     LEFT JOIN Marca M WITH (NOLOCK)                                                                                                                     " & vbCrLf & _
                            "       ON M.Marca_id = P.Marca                                                                                                                           " & vbCrLf & _
                            "                                                                                                                                                         " & vbCrLf & _
                            "     LEFT Join Pedidos Pe WITH (NOLOCK)                                                                                                                  " & vbCrLf & _
                            "       on NF.Empresa_Id    = Pe.Empresa_Id                                                                                                               " & vbCrLf & _
                            "      and NF.EndEmpresa_Id = Pe.EndEmpresa_Id                                                                                                            " & vbCrLf & _
                            "      and NF.Pedido        = Pe.Pedido_Id                                                                                                                " & vbCrLf & _
                            "    INNER JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id,      " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='PRODUTO'         THEN isnull(SituacaoTributariaIPI,0)       else 0 end) as SituacaoTributariaIPI,      " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='PRODUTO'         THEN isnull(SituacaoTributariaPISCOFINS,0) else 0 end) as SituacaoTributariaPISCOFINS," & vbCrLf & _
                            "                       sum(case when Encargo_Id ='PRODUTO'         THEN isnull(SituacaoTributaria,0)          else 0 end) as SituacaoTributariaICMS,     " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='ICMS'            THEN isnull(Valor, 0)                      else 0 end) AS Encargo1,                   " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='FUNRURAL'        THEN isnull(Valor, 0)                      else 0 end) AS Encargo2,                   " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='FUNRURAL TERCEI' THEN isnull(Valor, 0)                      else 0 end) AS Encargo3,                   " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='FETHAB'          THEN isnull(Valor, 0)                      else 0 end) AS Encargo4,                   " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='FACS'            THEN isnull(Valor, 0)                      else 0 end) AS Encargo5,                   " & vbCrLf & _
                            "                       sum(case when Encargo_Id ='LIQUIDO'         THEN isnull(Valor, 0)                      else 0 end) AS LIQUIDO                     " & vbCrLf & _
                            "                  FROM NotasFiscaisXEncargos WITH (NOLOCK)                                                                                               " & vbCrLf & _
                            "                 Where Encargo_Id in ('LIQUIDO', 'PRODUTO','ICMS','FUNRURAL','FUNRURAL TERCEI','FETHAB','FACS')                                          " & vbCrLf & _
                            "                 group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id,  Sequencia_Id   " & vbCrLf & _
                            "                ) AS Encargos                                                                                                                            " & vbCrLf & _
                            "       ON NFI.Empresa_Id      = Encargos.Empresa_Id                                                                                                      " & vbCrLf & _
                            "      AND NFI.EndEmpresa_Id   = Encargos.EndEmpresa_Id                                                                                                   " & vbCrLf & _
                            "      AND NFI.Cliente_Id      = Encargos.Cliente_Id                                                                                                      " & vbCrLf & _
                            "      AND NFI.EndCliente_Id   = Encargos.EndCliente_Id                                                                                                   " & vbCrLf & _
                            "      AND NFI.EntradaSaida_Id = Encargos.EntradaSaida_Id                                                                                                 " & vbCrLf & _
                            "      AND NFI.Serie_Id        = Encargos.Serie_Id                                                                                                        " & vbCrLf & _
                            "      AND NFI.Nota_Id         = Encargos.Nota_Id                                                                                                         " & vbCrLf & _
                            "      AND NFI.Produto_Id      = Encargos.Produto_Id                                                                                                      " & vbCrLf & _
                            "      AND NFI.CFOP_Id         = Encargos.CFOP_Id                                                                                                         " & vbCrLf & _
                            "      AND NFI.Sequencia_Id    = Encargos.Sequencia_Id                                                                                                    " & vbCrLf & _
                            "                                                                                                                                                         " & vbCrLf & _
                            "    Inner Join Operacoes as OP WITH (NOLOCK)                                                                                                             " & vbCrLf & _
                            "       ON OP.Operacao_Id     = NFI.Operacao                                                                                                              " & vbCrLf & _
                            "    INNER JOIN SubOperacoes AS SO WITH (NOLOCK)                                                                                                          " & vbCrLf & _
                            "       ON SO.Operacao_Id     = NFI.Operacao                                                                                                              " & vbCrLf & _
                            "      AND SO.SubOperacoes_Id = NFI.SubOperacao                                                                                                           " & vbCrLf & _
                            "     LEFT JOIN NotasFiscaisXRomaneios AS NFR WITH (NOLOCK)                                                                                               " & vbCrLf & _
                            "       ON NFR.Empresa_Id      = NFI.Empresa_Id                                                                                                           " & vbCrLf & _
                            "      AND NFR.EndEmpresa_Id   = NFI.EndEmpresa_Id                                                                                                        " & vbCrLf & _
                            "      AND NFR.Cliente_Id      = NFI.Cliente_Id                                                                                                           " & vbCrLf & _
                            "      AND NFR.EndCliente_Id   = NFI.EndCliente_Id                                                                                                        " & vbCrLf & _
                            "      AND NFR.EntradaSaida_Id = NFI.EntradaSaida_Id                                                                                                      " & vbCrLf & _
                            "      AND NFR.Serie_Id        = NFI.Serie_Id                                                                                                             " & vbCrLf & _
                            "      AND NFR.Nota_Id         = NFI.Nota_Id                                                                                                              " & vbCrLf & _
                            "      AND NFR.Produto_Id      = NFI.Produto_Id                                                                                                           " & vbCrLf & _
                            "      AND NFR.CFOP_Id         = NFI.CFOP_Id                                                                                                              " & vbCrLf & _
                            "     LEFT JOIN Clientes AS CLI WITH (NOLOCK)                                                                                                             " & vbCrLf & _
                            "       ON nf.Cliente_Id    = CLI.Cliente_Id                                                                                                              " & vbCrLf & _
                            "      AND nf.EndCliente_Id = CLI.Endereco_Id                                                                                                             " & vbCrLf & _
                            "     LEFT JOIN Clientes AS E WITH (NOLOCK)                                                                                                               " & vbCrLf & _
                            "       ON nf.Empresa_Id    = E.Cliente_Id                                                                                                                " & vbCrLf & _
                            "      AND nf.EndEmpresa_Id = E.Endereco_Id                                                                                                               " & vbCrLf & _
                            "     LEFT JOIN Clientes AS DO WITH (NOLOCK)                                                                                                              " & vbCrLf & _
                            "       ON NFI.Deposito    = DO.Cliente_Id                                                                                                                " & vbCrLf & _
                            "      AND NFI.EndDeposito = DO.Endereco_Id                                                                                                               " & vbCrLf & _
                            "     LEFT JOIN Clientes AS DD WITH (NOLOCK)                                                                                                              " & vbCrLf & _
                            "       ON NF.Destino    = DD.Cliente_Id                                                                                                                  " & vbCrLf & _
                            "      AND NF.EndDestino = DD.Endereco_Id                                                                                                                 " & vbCrLf & _
                            "                                                                                                                                                         " & vbCrLf & _
                            "    WHERE NF.Situacao = 1                                                                                                                                " & vbCrLf & _
                            "    AND NF.Movimento BETWEEN '2015-08-01' AND '2015-08-31'                                                                                               " & vbCrLf & _
                            "    AND NF.Empresa_Id    ='04854422000185'                                                                                                               " & vbCrLf & _
                            "    AND NF.EndEmpresa_Id = 0                                                                                                                             " & vbCrLf & _
                            "    AND NFI.Produto_id in (Select Produto_id from Produtos where LEFT(grupo,5) in ('10101'))                                                             " & vbCrLf & _
                            "    AND SO.Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "'                                                                                                                            " & vbCrLf & _
                            "    AND OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "'                                                                                                                             " & vbCrLf & _
                            "    AND NF.TipoDeDocumento in (1,2,3,4,5,6,7,8,9,10,11,12,13,14)                                                                                         " & vbCrLf & _
                            "    Group by E.Reduzido, NF.Empresa_Id, NF.EndEmpresa_Id, E.Nome, E.Cidade, E.Estado,                                                                    " & vbCrLf & _
                            "             NF.EntradaSaida_Id,                                                                                                                         " & vbCrLf & _
                            "   	      P.Grupo, NFI.Produto_Id,                                                                                                                    " & vbCrLf & _
                            "             P.Nome, NFI.CFOP_Id,                                                                                                                        " & vbCrLf & _
                            "             NF.Operacao, NF.SubOperacao,                                                                                                                " & vbCrLf & _
                            "             NF.Cliente_Id, NF.EndCliente_Id, ISNULL(CLI.Nome, ''),                                                                                      " & vbCrLf & _
                            "   		  CLI.Cidade, CLI.Estado                                                                                                                      " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "MovimentoFiscalPorOperacao")

        Return ds
    End Function

#End Region


End Class