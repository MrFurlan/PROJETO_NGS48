Imports System.Data
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RegistrosFiscaisSintegra
    Inherits BasePage

    Private Sql1 As String
    Private Sql As String
    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Private NomeArquivo As String
    Private NomeArquivo2 As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RegistrosFiscaisSintegra", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("Sintegra")) Then
                        Directory.CreateDirectory(Server.MapPath("Sintegra"))
                    End If
                    txtPeriodoInicial.Text = Today
                    txtPeriodoFinal.Text = Today
                    CarregarUnidade()
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

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next

        CargaProcessos()
    End Sub

    Private Sub CargaEmpresas()
        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub CargaProcessos()
        Dim Registro As String = ddlEmpresa.SelectedValue
        Dim Campo = Registro.Split("-")

        Sql = "SELECT Processo_Id, PeriodoInicial, PeriodoFinal, Convenio, NaturezaDaOperacao, ArquivoMagnetico" & vbCrLf & _
              " FROM ProcessoSintegra" & vbCrLf & _
              " WHERE Empresa_Id = '" & Campo(0) & "' " & vbCrLf & _
              "Order by Processo_Id DESC " & vbCrLf

        grdProcesso.DataSource = Banco.ConsultaDataSet(Sql, "ProcessoSintegra")
        grdProcesso.DataBind()

        Dim i As Integer
        For i = 0 To grdProcesso.Rows.Count - 1
            CType(grdProcesso.Rows(i).FindControl("imdBaixarArquivo"), LinkButton).Visible = False
        Next
    End Sub

    Private Sub Limpar()
        txtPeriodoFinal.Text = ""
        txtPeriodoInicial.Text = ""
        txtProcesso.Text = ""
        ddlArquivoMagnetico.SelectedIndex = 0
        ddlConvenio.SelectedIndex = 0
        ddlNaturezaDaOperacao.SelectedIndex = 2
        txtPeriodoInicial.Text = Today
        txtPeriodoFinal.Text = Today
        grdProcesso.DataSource = New List(Of Object)
        grdProcesso.DataBind()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub txtProcesso_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaProcessos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdProcesso_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdProcesso.SelectedIndexChanged
        Try
            txtProcesso.Text = grdProcesso.SelectedRow.Cells(1).Text()
            txtPeriodoInicial.Text = grdProcesso.SelectedRow.Cells(2).Text()
            txtPeriodoFinal.Text = grdProcesso.SelectedRow.Cells(3).Text()
            ddlConvenio.SelectedIndex = ddlConvenio.Items.IndexOf(ddlConvenio.Items.FindByValue(grdProcesso.SelectedRow.Cells(4).Text()))
            ddlNaturezaDaOperacao.SelectedIndex = ddlNaturezaDaOperacao.Items.IndexOf(ddlNaturezaDaOperacao.Items.FindByValue(grdProcesso.SelectedRow.Cells(5).Text))
            ddlArquivoMagnetico.SelectedIndex = ddlArquivoMagnetico.Items.IndexOf(ddlArquivoMagnetico.Items.FindByValue(grdProcesso.SelectedDataKey.Values(0)))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtProcesso.Text = ""
            ddlArquivoMagnetico.SelectedIndex = 0
            ddlConvenio.SelectedIndex = 0
            ddlNaturezaDaOperacao.SelectedIndex = 2
            grdProcesso.DataSource = Nothing
            grdProcesso.DataBind()

            Dim Registro() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

            CargaProcessos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub GeraSintegra()
        Dim Registro As String = ddlEmpresa.SelectedValue
        Dim Campo = Registro.Split("-")

        NomeArquivo2 = "Sintegra/" & Campo(0) & "-" & Format(CDate(txtPeriodoInicial.Text), "yyyyMMdd") & "-" & Format(CDate(txtPeriodoFinal.Text), "yyyyMMdd") & ".txt"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        Dim linha As String
        Dim strm As StreamWriter
        Dim IntNota As Integer
        Dim IntSequencia As Integer

        Dim Registro10 As Integer = 0
        Dim Registro11 As Integer = 0

        Dim Registro50 As Integer = 0
        Dim Registro54 As Integer = 0
        Dim Registro70 As Integer = 0
        Dim Registro71 As Integer = 0
        Dim Registro74 As Integer = 0
        Dim Registro75 As Integer = 0

        Dim Registro88 As Integer = 0

        Dim Registro90 As Integer = 0
        Dim Registro99 As Integer = 0

        Dim CGC As String = ""
        Dim Inscricao As String = ""

        'Temporaria Controla Notas de Serviço Conjugadas

        Sql1 = "Select nfxi.Empresa_Id, " & vbCrLf & _
               "       nfxi.EndEmpresa_Id, " & vbCrLf & _
               "       nfxi.Cliente_Id, nfxi.EndCliente_Id, nfxi.EntradaSaida_Id, nfxi.Serie_Id, " & vbCrLf & _
               "       nfxi.Nota_Id, " & vbCrLf & _
               "       sum(Case" & vbCrLf & _
               "             when OxE.CodigoFiscal in (1933,2933)" & vbCrLf & _
               "               then 1 " & vbCrLf & _
               "               else 0 " & vbCrLf & _
               "           end ) ItensCfop," & vbCrLf & _
               "       sum(1) nrItens" & vbCrLf & _
               "  into #NotasCompostas" & vbCrLf & _
               "  From notasfiscaisxitens nfxi" & vbCrLf & _
               " Inner Join OperacaoXEstado OxE" & vbCrLf & _
               "    on OxE.Codigo_Id = nfxi.operacaoxestado" & vbCrLf & _
               " Inner join (SELECT NF.Empresa_Id, NF.EndEmpresa_Id, " & vbCrLf & _
               "                    NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf & _
               "                    NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id    " & vbCrLf & _
               "               FROM NotasFiscais NF " & vbCrLf & _
               "              INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
               "                 ON NF.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
               "                AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
               "                AND NF.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
               "                AND NF.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
               "                AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
               "                AND NF.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
               "                AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
               "              Inner Join Suboperacoes SO" & vbCrLf & _
               "                 on SO.Operacao_id     = NFI.Operacao" & vbCrLf & _
               "                and SO.SubOperacoes_id = NFI.suboperacao" & vbCrLf & _
               "              inner join OperacaoXEstado OE" & vbCrLf & _
               "                 on oe.Codigo_Id = NFI.operacaoxestado" & vbCrLf & _
               "              WHERE NF.EntradaSaida_Id = 'E'" & vbCrLf & _
               "                And NF.Empresa_Id = '" & Campo(0) & "' " & vbCrLf & _
               "                AND NF.Movimento BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "' " & vbCrLf & _
               "                AND NF.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf & _
               "                AND Oe.CodigoFiscal IN (1933,2933)" & vbCrLf & _
               "             ) sb" & vbCrLf & _
               "    ON sb.Empresa_Id      = nfxi.Empresa_Id " & vbCrLf & _
               "   AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id " & vbCrLf & _
               "   AND sb.Cliente_Id      = nfxi.Cliente_Id " & vbCrLf & _
               "   AND sb.EndCliente_Id   = nfxi.EndCliente_Id " & vbCrLf & _
               "   AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id " & vbCrLf & _
               "   AND sb.Serie_Id        = nfxi.Serie_Id " & vbCrLf & _
               "   AND sb.Nota_Id         = nfxi.Nota_Id   " & vbCrLf & _
               " group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id, " & vbCrLf & _
               "          nfxi.Cliente_Id, nfxi.EndCliente_Id, " & vbCrLf & _
               "          nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf & _
               " having sum(1)  <> sum(Case" & vbCrLf & _
               "                         when OxE.CodigoFiscal in (1933,2933)" & vbCrLf & _
               "                           then 1" & vbCrLf & _
               "                           else 0" & vbCrLf & _
               "                       end );" & vbCrLf
        Try
            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim Array As New ArrayList

            'Registro Tipo 10
            'Registro Tipo 10

            Sql = "SELECT Cliente_Id, Inscricao, Nome, Cidade, Estado, Isnull(Telefone,'') AS Telefone " & vbCrLf & _
                  "  FROM Clientes" & vbCrLf & _
                  " WHERE Cliente_Id = '" & Campo(0) & "' And Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "10"                                                                                                'Tipo
                        linha &= Microsoft.VisualBasic.Left(.Item("Cliente_Id"), 14)                                                'CNPF/MF
                        linha &= Funcoes.AlinharEsquerda(IIf(.Item("Inscricao").trim = "", "ISENTO", .Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "")), 14, " ")
                        linha &= Funcoes.AlinharEsquerda(.Item("Nome"), 35, " ")                                                            'Nome do Contribuinte
                        linha &= Funcoes.AlinharEsquerda(.Item("Cidade"), 30, " ")                                                          'Municipio
                        linha &= .Item("Estado")                                                                                    'Estado
                        linha &= Funcoes.AlinharDireita(.Item("Telefone").Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", ""), 10, "0")  'Fax
                        linha &= Format(CDate(txtPeriodoInicial.Text), "yyyyMMdd")                                                  'Data Inicial
                        linha &= Format(CDate(txtPeriodoFinal.Text), "yyyyMMdd")                                                    'Data Final
                        linha &= ddlConvenio.SelectedValue
                        linha &= ddlNaturezaDaOperacao.SelectedValue
                        linha &= ddlArquivoMagnetico.SelectedValue

                        CGC = Microsoft.VisualBasic.Left(.Item("Cliente_Id"), 14)                                                       'CNPF/MF
                        Inscricao = Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " ")   'Inscricao Estadual

                    End With

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Registro10 += 1

                Next
            End If

            'Registro Tipo 11   

            Sql = "SELECT Clientes.Endereco, ISNULL(Clientes.Numero, '0') AS Numero, ISNULL(Clientes.Complemento, ' ') AS Complemento, Clientes.Bairro, Clientes.Cep, " & vbCrLf & _
                  "       ClientesXEmpresas.NomeDoContador, Clientes.Telefone" & vbCrLf & _
                  "  FROM Clientes" & vbCrLf & _
                  "  LEFT JOIN ClientesXEmpresas" & vbCrLf & _
                  "    ON Clientes.Cliente_Id  = ClientesXEmpresas.Empresa_Id" & vbCrLf & _
                  "   AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
                  " WHERE Clientes.Cliente_Id = '" & Campo(0) & "'" & vbCrLf & _
                  "   And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "11"                                                                                                                    'Tipo
                        linha &= Funcoes.AlinharEsquerda(.Item("Endereco"), 34, " ")                                                                    'Municipio
                        linha &= Funcoes.AlinharEsquerda(.Item("Numero"), 5, "0")                                                                       'Nome do Contribuinte
                        linha &= Funcoes.AlinharEsquerda(.Item("Complemento"), 22, " ")                                                                 'Municipio
                        linha &= Funcoes.AlinharEsquerda(.Item("Bairro"), 15, " ")                                                                      'Municipio
                        linha &= Funcoes.AlinharDireita(.Item("Cep").Replace(".", "").Replace("-", "").Replace(",", ""), 8, "0")                        'Fax
                        linha &= Funcoes.AlinharEsquerda(.Item("NomeDoContador"), 28, " ")                                                              'Municipio
                        linha &= Funcoes.AlinharDireita(.Item("Telefone").Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", ""), 12, "0") 'Fax
                    End With

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Registro11 += 1
                Next
            End If


            'Registro Tipo 50
            '---------------------------------------------------

            'Sql = "  SELECT     Cliente, Inscricao, Estado, EntradaSaida, Serie, NumeroDaNota, Movimento, DataDaNota, CFOP, Conta, SUM(((ValorContabil+ValorIPI+ValorFrete+OutrasDespesas) - ValorDesconto))  AS ValorContabil, sum(BaseICMS) as BaseICMS, " & vbCrLf
            'Sql &= "            sum(ICMS) as ICMS " & vbCrLf
            'Sql &= " FROM       (SELECT  NotasFiscais.Nota_Id AS NumeroDaNota, NotasFiscais.EntradaSaida_Id AS EntradaSaida, NotasFiscais.Serie_Id AS Serie, " & vbCrLf
            'Sql &= "            NotasFiscais.Cliente_Id AS Cliente, NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf
            'Sql &= "            NotasFiscais.EstadoDoCliente AS Estado, SubOperacoes.GrupoDeContas AS Conta, " & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorContabil, " & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorIPI, " & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id like '%ICMS%' THEN case when NotasFiscaisXEncargos.CFOP_Id IN (1551,2551) AND YEAR(NotasFiscais.Movimento) > 2011 then 0 else NotasFiscaisXEncargos.Base end ELSE 0 END) AS BaseICMS," & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id like '%ICMS%' THEN case when NotasFiscaisXEncargos.CFOP_Id IN (1551,2551) AND YEAR(NotasFiscais.Movimento) > 2011 then 0 else NotasFiscaisXEncargos.Valor end ELSE 0 END) AS ICMS, Clientes.Inscricao, " & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorDesconto, " & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'FRETES' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorFrete, " & vbCrLf
            'Sql &= "            SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESP.ADUANEIRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS OutrasDespesas " & vbCrLf
            'Sql &= " FROM       NotasFiscais INNER JOIN" & vbCrLf
            'Sql &= "            NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf
            'Sql &= "            NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf
            'Sql &= "            NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf
            'Sql &= "            NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf
            'Sql &= "            NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf
            'Sql &= "            NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf
            'Sql &= "            NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf
            'Sql &= "            NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf
            'Sql &= "            NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf
            'Sql &= "            NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf
            'Sql &= "            NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf
            'Sql &= "            NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id " & vbCrLf
            'Sql &= "            INNER JOIN SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf
            'Sql &= "            NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf
            'Sql &= "            Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf
            'Sql &= " WHERE      (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC')))" & vbCrLf
            'Sql &= "            AND (NotasFiscais.Empresa_Id = '" & Campo(0) & "')" & vbCrLf
            'Sql &= "            AND (NotasFiscais.Movimento BETWEEN '" & Format(CDate(txtPeriodoInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtPeriodoFinal.Text), "yyyy-MM-dd") & "')" & vbCrLf
            'Sql &= "            AND (NotasFiscais.Situacao = 1)" & vbCrLf
            'Sql &= "            AND Not (NotasFiscais.EntradaSaida_Id = 'E'   " & vbCrLf
            'Sql &= "            AND  NotasFiscaisXItens.CFOP_Id IN (1933,2933)    " & vbCrLf
            'Sql &= "            ) AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NotasFiscaisXItens.CFOP_Id IN (1949,2949)) " & vbCrLf
            'Sql &= "            AND Not (NotasFiscais.EntradaSaida_Id = 'S'   " & vbCrLf
            'Sql &= "            AND  NotasFiscaisXItens.CFOP_Id IN (5933,6933)    " & vbCrLf
            'Sql &= "            ) AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NotasFiscaisXItens.CFOP_Id IN (5949,6949)) " & vbCrLf
            'Sql &= "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1350 AND 1360) AND" & vbCrLf
            'Sql &= "                  (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 2350 AND 2360) AND" & vbCrLf
            'Sql &= "                  (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 3350 AND 3360) AND" & vbCrLf
            'Sql &= "                  (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 5350 AND 5360) AND" & vbCrLf
            'Sql &= "                  (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 6350 AND 6360) AND" & vbCrLf
            'Sql &= "                  (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 7350 AND 7360) " & vbCrLf
            'Sql &= "  OR exists (Select  1 from #notascompostas where " & vbCrLf & _
            '           "     #NotasCompostas.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
            '           "     AND #NotasCompostas.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
            '           "     AND #NotasCompostas.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
            '           "     AND #NotasCompostas.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
            '           "     AND #NotasCompostas.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
            '           "     AND #NotasCompostas.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
            '           "     AND #NotasCompostas.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
            '           "              )                                  " & vbCrLf
            'Sql &= " GROUP BY   NotasFiscaisXItens.Produto_Id, NotasFiscais.Nota_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Cliente_Id, NotasFiscais.Movimento, " & vbCrLf
            'Sql &= "            NotasFiscais.DataDaNota, NotasFiscaisXItens.CFOP_Id, SubOperacoes.GrupoDeContas, NotasFiscais.EstadoDoCliente, Clientes.Inscricao ) AS Consulta" & vbCrLf
            'Sql &= " GROUP BY Cliente, Inscricao, Estado, EntradaSaida, Serie, NumeroDaNota, Movimento, " & vbCrLf
            'Sql &= " DataDaNota, CFOP, Conta " & vbCrLf
            'Sql &= " ORDER BY Movimento, Serie, NumeroDaNota" & vbCrLf

            '-----------------------

            Sql = "SELECT Cliente, Inscricao, Estado, EntradaSaida, Serie, NumeroDaNota, Movimento, CFOP, Conta, " & vbCrLf & _
                  "		  sum(((ValorContabil+ValorIPI+ValorFrete+OutrasDespesas) - ValorDesconto))  AS ValorContabil, " & vbCrLf & _
                  "		  sum(BaseICMS) as BaseICMS," & vbCrLf & _
                  "       sum(ICMS) as ICMS," & vbCrLf & _
                  "       Aliquota " & vbCrLf & _
                  "  FROM (SELECT N.Nota_Id AS NumeroDaNota, N.EntradaSaida_Id AS EntradaSaida, N.Serie_Id AS Serie, " & vbCrLf & _
                  "				  N.Cliente_Id AS Cliente, N.Movimento, N.DataDaNota, OE.CodigoFiscal AS CFOP, " & vbCrLf & _
                  "				  N.EstadoDoCliente AS Estado, So.GrupoDeContas AS Conta, C.Inscricao, " & vbCrLf & _
                  "				  sum(isnull(sbPrd.Valor,0)) AS ValorContabil, " & vbCrLf & _
                  "				  sum(isnull(sbIPI.Valor,0)) AS ValorIPI, " & vbCrLf & _
                  "				  sum(case " & vbCrLf & _
                  "					    when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(N.Movimento) > 2011 " & vbCrLf & _
                  "						  then 0 " & vbCrLf & _
                  "						  else isnull(sbIcms.Base,0) " & vbCrLf & _
                  "				      end) AS BaseICMS, " & vbCrLf & _
                  "				  case " & vbCrLf & _
                  "					when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(N.Movimento) > 2011 " & vbCrLf & _
                  "					  then 0 " & vbCrLf & _
                  "					  else isnull(sbIcms.Percentual,0) " & vbCrLf & _
                  "				  end AS Aliquota, " & vbCrLf & _
                  "				  sum(case " & vbCrLf & _
                  "					    when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(N.Movimento) > 2011 " & vbCrLf & _
                  "						  then 0 " & vbCrLf & _
                  "						  else isnull(sbIcms.Valor,0) " & vbCrLf & _
                  "				  end) AS ICMS, " & vbCrLf & _
                  "				  sum(isnull(sbDesc.Valor,0)) AS ValorDesconto, " & vbCrLf & _
                  "				  sum(isnull(sbFrt.Valor,0)) AS ValorFrete, " & vbCrLf & _
                  "				  sum(isnull(sbDAdn.Valor,0)) AS OutrasDespesas " & vbCrLf & _
                  "		     FROM NotasFiscais N " & vbCrLf & _
                  "			INNER JOIN NotasFiscaisXItens NxI " & vbCrLf & _
                  "		       ON NxI.Empresa_Id      = N.Empresa_Id " & vbCrLf & _
                  "			  AND NxI.EndEmpresa_Id   = N.EndEmpresa_Id " & vbCrLf & _
                  "			  AND NxI.Cliente_Id      = N.Cliente_Id " & vbCrLf & _
                  "			  AND NxI.EndCliente_Id   = N.EndCliente_Id " & vbCrLf & _
                  "			  AND NxI.EntradaSaida_Id = N.EntradaSaida_Id " & vbCrLf & _
                  "			  AND NxI.Serie_Id        = N.Serie_Id " & vbCrLf & _
                  "			  AND NxI.Nota_Id         = N.Nota_Id " & vbCrLf & _
                  "			INNER JOIN (SELECT NxExPRD.Empresa_Id, NxExPRD.EndEmpresa_Id, NxExPRD.Cliente_Id, NxExPRD.EndCliente_Id, " & vbCrLf & _
                  "							   NxExPRD.EntradaSaida_Id, NxExPRD.Serie_Id, NxExPRD.Nota_Id, " & vbCrLf & _
                  "							   NxExPRD.Produto_Id, NxExPRD.CFOP_Id, NxExPRD.Sequencia_Id, sum(NxExPRD.Valor) AS Valor " & vbCrLf & _
                  "					      From NotasFiscaisXEncargos NxExPRD " & vbCrLf & _
                  "						 Where NxExPRD.Encargo_id = 'PRODUTO' " & vbCrLf & _
                  "						 Group By NxExPRD.Empresa_Id, NxExPRD.EndEmpresa_Id, NxExPRD.Cliente_Id, NxExPRD.EndCliente_Id, " & vbCrLf & _
                  "						          NxExPRD.EntradaSaida_Id, NxExPRD.Serie_Id, NxExPRD.Nota_Id, NxExPRD.Produto_Id, NxExPRD.CFOP_Id, NxExPRD.Sequencia_Id" & vbCrLf & _
                  "                     ) sbPrd " & vbCrLf & _
                  "			   ON sbPrd.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                  "			  AND sbPrd.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                  "			  AND sbPrd.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                  "			  AND sbPrd.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                  "			  AND sbPrd.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                  "			  AND sbPrd.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                  "			  AND sbPrd.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                  "			  AND sbPrd.Produto_Id      = NxI.Produto_Id " & vbCrLf & _
                  "			  AND sbPrd.CFOP_Id         = NxI.CFOP_Id " & vbCrLf & _
                  "           AND sbPrd.Sequencia_id    = NxI.Sequencia_id " & vbCrLf & _
                  "			 LEFT JOIN (SELECT NxExIPI.Empresa_Id, NxExIPI.EndEmpresa_Id, NxExIPI.Cliente_Id, NxExIPI.EndCliente_Id, " & vbCrLf & _
                  "							   NxExIPI.EntradaSaida_Id, NxExIPI.Serie_Id, NxExIPI.Nota_Id, " & vbCrLf & _
                  "							   NxExIPI.Produto_Id, NxExIPI.CFOP_Id, NxExIPI.Sequencia_id, sum(NxExIPI.Valor) AS Valor " & vbCrLf & _
                  "						  From NotasFiscaisXEncargos NxExIPI " & vbCrLf & _
                  "						 Where NxExIPI.Encargo_id = 'IPI' " & vbCrLf & _
                  "						 Group By NxExIPI.Empresa_Id, NxExIPI.EndEmpresa_Id, NxExIPI.Cliente_Id, NxExIPI.EndCliente_Id, " & vbCrLf & _
                  "						          NxExIPI.EntradaSaida_Id, NxExIPI.Serie_Id, NxExIPI.Nota_Id, NxExIPI.Produto_Id, NxExIPI.CFOP_Id, NxExIPI.Sequencia_Id" & vbCrLf & _
                  "                     ) sbIPI " & vbCrLf & _
                  "			   ON sbIPI.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                  "			  AND sbIPI.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                  "			  AND sbIPI.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                  "			  AND sbIPI.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                  "			  AND sbIPI.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                  "			  AND sbIPI.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                  "			  AND sbIPI.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                  "			  AND sbIPI.Produto_Id      = NxI.Produto_Id " & vbCrLf & _
                  "			  AND sbIPI.CFOP_Id         = NxI.CFOP_Id " & vbCrLf & _
                  "           AND sbIPI.Sequencia_id    = NxI.Sequencia_id " & vbCrLf & _
                  "			 LEFT JOIN (SELECT NxExIcms.Empresa_Id, NxExIcms.EndEmpresa_Id, NxExIcms.Cliente_Id, NxExIcms.EndCliente_Id, " & vbCrLf & _
                  "							   NxExIcms.EntradaSaida_Id, NxExIcms.Serie_Id, NxExIcms.Nota_Id, " & vbCrLf & _
                  "							   NxExIcms.Produto_Id, NxExIcms.CFOP_Id, NxExIcms.sequencia_id, NxExIcms.PercentualExibicao AS Percentual, sum(NxExIcms.Base) AS Base, " & vbCrLf & _
                  "							   sum(NxExIcms.Valor) AS Valor " & vbCrLf & _
                  "						  From NotasFiscaisXEncargos NxExIcms " & vbCrLf & _
                  "                      Inner join Encargos Enc" & vbCrLf & _
                  "                         on Enc.Encargo_id = NxExIcms.Encargo_id" & vbCrLf & _
                  "						 Where NxExIcms.Encargo_id = 'ICMS' or Enc.EncargoAgrupador = 'ICMS'" & vbCrLf & _
                  "						 Group By NxExIcms.Empresa_Id, NxExIcms.EndEmpresa_Id, NxExIcms.Cliente_Id, NxExIcms.EndCliente_Id, " & vbCrLf & _
                  "						          NxExIcms.EntradaSaida_Id, NxExIcms.Serie_Id, NxExIcms.Nota_Id, " & vbCrLf & _
                  "								  NxExIcms.PercentualExibicao, NxExIcms.Produto_Id, NxExIcms.CFOP_Id, NxExIcms.sequencia_id" & vbCrLf & _
                  "                     ) sbIcms " & vbCrLf & _
                  "			   ON sbIcms.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                  "			  AND sbIcms.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                  "			  AND sbIcms.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                  "			  AND sbIcms.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                  "			  AND sbIcms.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                  "			  AND sbIcms.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                  "			  AND sbIcms.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                  "			  AND sbIcms.Produto_Id      = NxI.Produto_Id " & vbCrLf & _
                  "			  AND sbIcms.CFOP_Id         = NxI.CFOP_Id " & vbCrLf & _
                  "           AND sbIcms.Sequencia_id    = NxI.Sequencia_id " & vbCrLf & _
                  "			 LEFT JOIN (SELECT NxExDesc.Empresa_Id, NxExDesc.EndEmpresa_Id, NxExDesc.Cliente_Id, NxExDesc.EndCliente_Id, " & vbCrLf & _
                  "						       NxExDesc.EntradaSaida_Id, NxExDesc.Serie_Id, NxExDesc.Nota_Id, " & vbCrLf & _
                  "							   NxExDesc.Produto_Id, NxExDesc.CFOP_Id, NxExDesc.Sequencia_Id, sum(NxExDesc.Valor) AS Valor " & vbCrLf & _
                  "						  From NotasFiscaisXEncargos NxExDesc " & vbCrLf & _
                  "						 Where NxExDesc.Encargo_id = 'Desc' " & vbCrLf & _
                  "						 Group By NxExDesc.Empresa_Id, NxExDesc.EndEmpresa_Id, NxExDesc.Cliente_Id, NxExDesc.EndCliente_Id, " & vbCrLf & _
                  "						          NxExDesc.EntradaSaida_Id, NxExDesc.Serie_Id, NxExDesc.Nota_Id, NxExDesc.Produto_Id, NxExDesc.CFOP_Id, NxExDesc.Sequencia_Id" & vbCrLf & _
                  "                     ) sbDesc " & vbCrLf & _
                  "			   ON sbDesc.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                  "			  AND sbDesc.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                  "			  AND sbDesc.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                  "			  AND sbDesc.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                  "			  AND sbDesc.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                  "			  AND sbDesc.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                  "			  AND sbDesc.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                  "			  AND sbDesc.Produto_Id      = NxI.Produto_Id " & vbCrLf & _
                  "			  AND sbDesc.CFOP_Id         = NxI.CFOP_Id " & vbCrLf & _
                  "           AND sbDesc.Sequencia_id    = NxI.Sequencia_id " & vbCrLf & _
                  "			 LEFT JOIN (SELECT NxExFrt.Empresa_Id, NxExFrt.EndEmpresa_Id, NxExFrt.Cliente_Id, NxExFrt.EndCliente_Id, " & vbCrLf & _
                  "							   NxExFrt.EntradaSaida_Id, NxExFrt.Serie_Id, NxExFrt.Nota_Id, " & vbCrLf & _
                  "							   NxExFrt.Produto_Id, NxExFrt.CFOP_Id, NxExFrt.Sequencia_Id, sum(NxExFrt.Valor) AS Valor " & vbCrLf & _
                  "						  From NotasFiscaisXEncargos NxExFrt " & vbCrLf & _
                  "						 Where NxExFrt.Encargo_id = 'Frt' " & vbCrLf & _
                  "						 Group By NxExFrt.Empresa_Id, NxExFrt.EndEmpresa_Id, NxExFrt.Cliente_Id, NxExFrt.EndCliente_Id, " & vbCrLf & _
                  "						          NxExFrt.EntradaSaida_Id, NxExFrt.Serie_Id, NxExFrt.Nota_Id, NxExFrt.Produto_Id, NxExFrt.CFOP_Id, NxExFrt.Sequencia_Id" & vbCrLf & _
                  "                     ) sbFrt " & vbCrLf & _
                  "			   ON sbFrt.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                  "			  AND sbFrt.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                  "			  AND sbFrt.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                  "			  AND sbFrt.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                  "			  AND sbFrt.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                  "			  AND sbFrt.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                  "			  AND sbFrt.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                  "			  AND sbFrt.Produto_Id      = NxI.Produto_Id " & vbCrLf & _
                  "			  AND sbFrt.CFOP_Id         = NxI.CFOP_Id " & vbCrLf & _
                  "           AND sbFrt.Sequencia_id    = NxI.Sequencia_id " & vbCrLf & _
                  "			 LEFT JOIN (SELECT NxExDAdn.Empresa_Id, NxExDAdn.EndEmpresa_Id, NxExDAdn.Cliente_Id, NxExDAdn.EndCliente_Id, " & vbCrLf & _
                  "							   NxExDAdn.EntradaSaida_Id, NxExDAdn.Serie_Id, NxExDAdn.Nota_Id, " & vbCrLf & _
                  "							   NxExDAdn.Produto_Id, NxExDAdn.CFOP_Id, NxExDAdn.Sequencia_Id, sum(NxExDAdn.Valor) AS Valor " & vbCrLf & _
                  "						  From NotasFiscaisXEncargos NxExDAdn " & vbCrLf & _
                  "						 Where NxExDAdn.Encargo_id = 'DAdn' " & vbCrLf & _
                  "						 Group By NxExDAdn.Empresa_Id, NxExDAdn.EndEmpresa_Id, NxExDAdn.Cliente_Id, NxExDAdn.EndCliente_Id, " & vbCrLf & _
                  "								  NxExDAdn.EntradaSaida_Id, NxExDAdn.Serie_Id, NxExDAdn.Nota_Id, NxExDAdn.Produto_Id, NxExDAdn.CFOP_Id, NxExDAdn.Sequencia_Id" & vbCrLf & _
                  "                     ) sbDAdn " & vbCrLf & _
                  "			   ON sbDAdn.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                  "			  AND sbDAdn.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                  "			  AND sbDAdn.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                  "			  AND sbDAdn.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                  "			  AND sbDAdn.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                  "			  AND sbDAdn.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                  "			  AND sbDAdn.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                  "			  AND sbDAdn.Produto_Id      = NxI.Produto_Id " & vbCrLf & _
                  "			  AND sbDAdn.CFOP_Id         = NxI.CFOP_Id " & vbCrLf & _
                  "           AND sbDAdn.Sequencia_id    = NxI.Sequencia_id " & vbCrLf & _
                  "			INNER JOIN SubOperacoes So " & vbCrLf & _
                  "			   ON So.Operacao_Id = N.Operacao " & vbCrLf & _
                  "			  AND So.SubOperacoes_Id = N.SubOperacao " & vbCrLf & _
                  "			INNER JOIN Clientes C " & vbCrLf & _
                  "			   ON C.Cliente_Id  = N.Cliente_Id " & vbCrLf & _
                  "			  AND C.Endereco_Id = N.EndCliente_Id " & vbCrLf & _
                  "         Inner join OperacaoxEstado OE" & vbCrLf & _
                  "            on OE.Codigo_id = NxI.OperacaoxEstado" & vbCrLf & _
                  "		    WHERE NOT (N.Serie_Id IN ('D', 'F', 'REC')) " & vbCrLf & _
                  "			  AND N.Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
                  "			  AND N.Movimento BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                  "			  AND N.Situacao   = 1" & vbCrLf & _
                  "			  AND Not (N.EntradaSaida_Id = 'E' AND OE.CodigoFiscal IN (1933,2933)) " & vbCrLf & _
                  "			  AND Not (So.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "'  AND OE.CodigoFiscal IN (1949,2949)) " & vbCrLf & _
                  "			  AND Not (N.EntradaSaida_Id = 'S' AND OE.CodigoFiscal IN (5933,6933)) " & vbCrLf & _
                  "			  AND NOT (So.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "'  AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf & _
                  "			  AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                  "			  AND (OE.CodigoFiscal NOT BETWEEN 2350 AND 2360)" & vbCrLf & _
                  "			  AND (OE.CodigoFiscal NOT BETWEEN 3350 AND 3360)" & vbCrLf & _
                  "			  AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                  "			  AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                  "			  AND (OE.CodigoFiscal NOT BETWEEN 7350 AND 7360)" & vbCrLf & _
                  "			   OR exists (Select  1 from #notascompostas " & vbCrLf & _
                  "						   where #NotasCompostas.Empresa_Id      = N.Empresa_Id " & vbCrLf & _
                  "							 AND #NotasCompostas.EndEmpresa_Id   = N.EndEmpresa_Id " & vbCrLf & _
                  "							 AND #NotasCompostas.Cliente_Id      = N.Cliente_Id " & vbCrLf & _
                  "							 AND #NotasCompostas.EndCliente_Id   = N.EndCliente_Id " & vbCrLf & _
                  "							 AND #NotasCompostas.EntradaSaida_Id = N.EntradaSaida_Id " & vbCrLf & _
                  "							 AND #NotasCompostas.Serie_Id        = N.Serie_Id " & vbCrLf & _
                  "							 AND #NotasCompostas.Nota_Id         = N.Nota_Id) " & vbCrLf & _
                  "		    GROUP BY N.Nota_Id, N.EntradaSaida_Id, N.Serie_Id, N.Cliente_Id, N.Movimento, OE.CodigoFiscal, OE.STICMS, " & vbCrLf & _
                  "				  N.DataDaNota, NxI.CFOP_Id, isnull(sbIcms.Percentual,0), So.GrupoDeContas, N.EstadoDoCliente, C.Inscricao" & vbCrLf & _
                  "        ) AS Consulta " & vbCrLf & _
                  "  GROUP BY Cliente, Inscricao, Estado, EntradaSaida, Serie, NumeroDaNota, Movimento, CFOP, Conta, Aliquota " & vbCrLf & _
                  "  ORDER BY Movimento, Serie, NumeroDaNota "

            ds = Banco.ConsultaDataSet(Sql1 + Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "50"
                        If .Item("Estado") = "EX" Then
                            linha &= "00000000000000"                                                        'Cliente
                            linha &= "ISENTO        "                                                         'Inscrição
                        Else
                            linha &= Funcoes.AlinharDireita(.Item("Cliente").Trim(), 14, "0")                                                       'Cliente
                            'linha &= Funcoes.AlinharEsquerda(.Item("Inscricao").Trim.Replace(".", "").Replace(",", "").Replace("-", ""), 14, " ")    'Inscrição
                            linha &= Funcoes.AlinharEsquerda(IIf(.Item("Inscricao").trim = "", "ISENTO", .Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "")), 14, " ") 'Inscricao Estadual
                        End If

                        linha &= Format(dr("Movimento"), "yyyyMMdd")                                                                'Movimento                   
                        linha &= Funcoes.AlinharEsquerda(.Item("Estado"), 2, " ")                                                        'Estado
                        linha &= "01"                                                                                               'Série
                        Dim Ser = Funcoes.ValidarSerie(.Item("Serie"), .Item("EntradaSaida"))

                        linha &= Funcoes.AlinharEsquerda(Ser, 3, " ")                                                                       'Subserie
                        linha &= Funcoes.AlinharDireita(.Item("NumeroDaNota"), 6, "0")                                                           'Nota
                        linha &= Funcoes.AlinharDireita(.Item("CFOP"), 4, "0")                                                      'Cfop
                        linha &= "P"                                                                                                'Série
                        linha &= Funcoes.AlinharDireita(CStr(Math.Abs(.Item("ValorContabil"))).Replace(",", ""), 13, "0")                'Valor da Nota
                        linha &= Funcoes.AlinharDireita(CStr(Math.Abs(.Item("BaseIcms"))).Replace(",", ""), 13, "0")                        'Base de Calculo
                        linha &= Funcoes.AlinharDireita(CStr(Math.Abs(.Item("ICMS"))).Replace(",", ""), 13, "0")                            'ICMS
                        linha &= "0000000000000"                                                                                    'Isentas
                        Dim Outras As Decimal = Math.Abs(.Item("ValorContabil") - .Item("BaseICMS"))                             'Calculo
                        linha &= Funcoes.AlinharDireita(CStr(Outras).Replace(",", ""), 13, "0")                                             'Outras
                        linha &= Funcoes.AlinharDireita(CStr(.Item("Aliquota")).Replace(",", ""), 4, "0")                                   'Aliquota
                        'linha &= "0000"                                                                                                         'Aliquota
                        linha &= "N"                                                                                                'Situacao
                    End With

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Registro50 += 1

                Next
            End If

            '---------Registro Tipo 54-----------------------------------

            Sql = "SELECT Cliente, Estado, EntradaSaida, Serie, NumeroDaNota as Nota, Produto, CFOP, Quantidade, ((ValorContabil+ValorIPI+ValorFrete+OutrasDespesas) - ValorDesconto) as Valor, BaseICMS, ValorICMS, ValorIPI, AliquotaICMS" & vbCrLf & _
                  "  FROM (SELECT  NF.Cliente_Id AS Cliente, NF.EntradaSaida_Id AS EntradaSaida, NF.Serie_Id AS Serie, NF.EstadoDoCliente AS Estado," & vbCrLf & _
                  "                NF.Nota_Id AS NumeroDaNota, NFI.Produto_Id AS Produto, OE.CodigoFiscal AS CFOP, " & vbCrLf & _
                  "                NFI.QuantidadeFiscal AS Quantidade, " & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id =    'PRODUTO'         THEN NFE.Valor ELSE 0 END) AS ValorContabil, " & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id like '%ICMS%'          THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else NFE.Base end ELSE 0 END) AS BaseICMS," & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id like '%ICMS%'          THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else NFE.PercentualExibicao end ELSE 0 END) AS AliquotaICMS," & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id like '%ICMS%'          THEN case when ((OE.CodigoFiscal IN (1551,2551)) or (OE.CodigoFiscal IN (1653,2653) and OE.STICMS in(10,60))) AND YEAR(NF.Movimento) > 2011 then 0 else NFE.Valor end ELSE 0 END) AS ValorICMS," & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id =    'IPI'             THEN NFE.Valor ELSE 0 END) AS ValorIPI, " & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id =    'DESCONTOS'       THEN NFE.Valor ELSE 0 END) AS ValorDesconto, " & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id =    'FRETES'          THEN NFE.Valor ELSE 0 END) AS ValorFrete, " & vbCrLf & _
                  "                SUM(CASE WHEN NFE.Encargo_id =    'DESP.ADUANEIRAS' THEN NFE.Valor ELSE 0 END) AS OutrasDespesas " & vbCrLf & _
                  "           FROM NotasFiscais NF" & vbCrLf & _
                  "          INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
                  "             ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
                  "            AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                  "            AND NF.Cliente_Id      = NFI.Cliente_Id  " & vbCrLf & _
                  "            AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                  "            AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
                  "            AND NF.Serie_Id        = NFI.Serie_Id  " & vbCrLf & _
                  "            AND NF.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                  "          INNER JOIN OperacaoxEstado OE" & vbCrLf & _
                  "             On OE.Codigo_Id = NFI.OperacaoxEstado" & vbCrLf & _
                  "          INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf & _
                  "             ON NFI.Empresa_Id      = NFE.Empresa_Id" & vbCrLf & _
                  "            AND NFI.EndEmpresa_Id   = NFE.EndEmpresa_Id  " & vbCrLf & _
                  "            AND NFI.Cliente_Id      = NFE.Cliente_Id" & vbCrLf & _
                  "            AND NFI.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf & _
                  "            AND NFI.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf & _
                  "            AND NFI.Serie_Id        = NFE.Serie_Id" & vbCrLf & _
                  "            AND NFI.Nota_Id         = NFE.Nota_Id  " & vbCrLf & _
                  "            AND NFI.Produto_Id      = NFE.Produto_Id" & vbCrLf & _
                  "            AND NFI.Cfop_ID         = NFE.Cfop_Id" & vbCrLf & _
                  "            AND NFI.Sequencia_Id    = NFE.Sequencia_Id " & vbCrLf & _
                  "            AND NFE.Encargo_Id IN ('ISS', 'PRODUTO') " & vbCrLf & _
                  "          INNER JOIN SubOperacoes SO" & vbCrLf & _
                  "             ON NF.Operacao    = SO.Operacao_Id" & vbCrLf & _
                  "            AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                  "          INNER JOIN Clientes" & vbCrLf & _
                  "             ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf & _
                  "            AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                  "          WHERE (NOT (NF.Serie_Id IN ('D', 'F', 'REC')))" & vbCrLf & _
                  "            AND (NF.Empresa_Id = '" & Campo(0) & "')" & vbCrLf & _
                  "            AND (NF.Movimento BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "')" & vbCrLf & _
                  "            AND (NFI.CFOP_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                  "            AND (NFI.CFOP_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf & _
                  "            AND (NFI.CFOP_Id NOT BETWEEN 3350 AND 3360)" & vbCrLf & _
                  "            AND (NFI.CFOP_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                  "            AND (NFI.CFOP_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                  "            AND (NFI.CFOP_Id NOT BETWEEN 7350 AND 7360)" & vbCrLf & _
                  "            AND (NF.Situacao = 1)" & vbCrLf & _
                  "            AND Not (NF.EntradaSaida_Id = 'E' AND  OE.Codigofiscal IN (1933,2933))    " & vbCrLf & _
                  "            AND Not (SO.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "'   AND  OE.Codigofiscal IN (1949,2949)) " & vbCrLf & _
                  "            AND Not (NF.EntradaSaida_Id = 'S' AND  OE.Codigofiscal IN (5933,6933))    " & vbCrLf & _
                  "            AND Not (SO.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "'   AND  OE.Codigofiscal IN (5949,6949)) " & vbCrLf & _
                  "             OR exists (Select  1" & vbCrLf & _
                  "                          from #notascompostas" & vbCrLf & _
                  "                         where #NotasCompostas.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                  "                           AND #NotasCompostas.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                  "                           AND #NotasCompostas.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                  "                           AND #NotasCompostas.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                  "                           AND #NotasCompostas.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                  "                           AND #NotasCompostas.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                  "                           AND #NotasCompostas.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                  "                         )" & vbCrLf & _
                  "          GROUP BY NF.Cliente_Id, NF.EstadoDoCliente, NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id, NFI.Produto_Id,  OE.Codigofiscal, NFI.QuantidadeFiscal" & vbCrLf & _
                  "        ) AS Consulta" & vbCrLf & _
                  "   ORDER BY EntradaSaida, Cliente, NumeroDaNota" & vbCrLf

            IntNota = 0
            IntSequencia = 0

            ds = Banco.ConsultaDataSet(Sql1 + Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "54"                                                                                                'Tipo
                        If .Item("Estado") = "EX" Then
                            linha &= "00000000000000"                                                        'Cliente
                        Else
                            linha &= Funcoes.AlinharDireita(.Item("Cliente").Trim(), 14, "0")                                                       'Cliente
                        End If

                        linha &= "01"                                                                                               'Série

                        Dim Ser = Funcoes.ValidarSerie(.Item("Serie"), .Item("EntradaSaida"))
                        linha &= Funcoes.AlinharEsquerda(Ser, 3, " ")                                                                       'Subserie

                        linha &= Funcoes.AlinharDireita(.Item("Nota"), 6, "0")                                                           'Nota
                        linha &= Funcoes.AlinharDireita(.Item("CFOP"), 4, "0")                                                      'Cfop
                        linha &= "000" 'Situacao Tributaria
                        If CInt(.Item("Nota")) <> IntNota Then
                            IntSequencia = 0
                        End If
                        IntSequencia += 1
                        linha &= Funcoes.AlinharDireita(CStr(IntSequencia), 3, "0") 'Sequencia do Item
                        linha &= Funcoes.AlinharEsquerda(.Item("Produto"), 14, " ")                                                         'Subserie

                        Dim Valor As String = Microsoft.VisualBasic.Format(.Item("Quantidade"), "00000000.000")

                        linha &= Valor.Replace(",", "")                                                           'Nota
                        linha &= Funcoes.AlinharDireita(CStr(.Item("Valor")).Replace(",", ""), 12, "0")      'Valor da Nota
                        linha &= "000000000000"                                                      'Descontos

                        linha &= Funcoes.AlinharDireita(CStr(.Item("BaseIcms")).Replace(",", ""), 12, "0")   'Base de Calculo
                        linha &= Funcoes.AlinharDireita(CStr(.Item("ValorICMS")).Replace(",", ""), 12, "0")  'ICMS

                        linha &= Funcoes.AlinharDireita(CStr(.Item("ValorIPI")).Replace(",", ""), 12, "0")   'IPI
                        linha &= Funcoes.AlinharDireita(CStr(.Item("AliquotaIcms")).Replace(",", ""), 4, "0") 'Aliquota

                        IntNota = CInt(.Item("Nota"))                                                  'Nota Anterior
                    End With

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Registro54 += 1
                Next
            End If

            'Registro Tipo 70

            Sql = "SELECT Cliente, Inscricao, Estado, EntradaSaida, Serie, NumeroDaNota as Nota, Movimento, DataDaNota, CFOP, ((ValorContabil+ValorIPI+ValorFrete+OutrasDespesas) - ValorDesconto) as Valor, BaseICMS, " & vbCrLf & _
                  "       ROUND(CASE WHEN ValorICMS = 0 THEN 0 ELSE ValorICMS * 100 / BaseICms END, 0) AS Aliquota, ValorICMS" & vbCrLf & _
                  "  FROM (SELECT NF.Nota_Id AS NumeroDaNota, NF.EntradaSaida_Id AS EntradaSaida, NF.Serie_Id AS Serie, " & vbCrLf & _
                  "               NF.Cliente_Id AS Cliente, NF.Movimento, NF.DataDaNota, OE.CodigoFiscal AS CFOP, " & vbCrLf & _
                  "               NF.EstadoDoCliente AS Estado, SO.GrupoDeContas AS Conta, " & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id = 'PRODUTO'         THEN NFE.Valor      ELSE 0 END) AS ValorContabil," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id LIKE '%ICMS%'       THEN NFE.Base       ELSE 0 END) AS BaseICMS," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id LIKE '%ICMS%'       THEN NFE.Percentual ELSE 0 END) AS AliquotaICMS," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id LIKE '%ICMS%'       THEN NFE.Valor      ELSE 0 END) AS ValorICMS," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id = 'IPI'             THEN NFE.Valor      ELSE 0 END) AS ValorIPI," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id = 'DESCONTOS'       THEN NFE.Valor      ELSE 0 END) AS ValorDesconto," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id = 'FRETES'          THEN NFE.Valor      ELSE 0 END) AS ValorFrete," & vbCrLf & _
                  "               SUM(CASE WHEN NFE.Encargo_id = 'DESP.ADUANEIRAS' THEN NFE.Valor      ELSE 0 END) AS OutrasDespesas," & vbCrLf & _
                  "               Clientes.Inscricao" & vbCrLf & _
                  "          FROM NotasFiscais NF" & vbCrLf & _
                  "         INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
                  "            ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
                  "           AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                  "           AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                  "           AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                  "           AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
                  "           AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                  "           AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
                  "         INNER JOIN OperacaoxEstado OE" & vbCrLf & _
                  "            on OE.Codigo_Id = NFI.OperacaoXEstado" & vbCrLf & _
                  "         INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf & _
                  "            ON NFI.Empresa_Id      = NFE.Empresa_Id" & vbCrLf & _
                  "           AND NFI.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf & _
                  "           AND NFI.Cliente_Id      = NFE.Cliente_Id" & vbCrLf & _
                  "           AND NFI.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf & _
                  "           AND NFI.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf & _
                  "           AND NFI.Serie_Id        = NFE.Serie_Id" & vbCrLf & _
                  "           AND NFI.Nota_Id         = NFE.Nota_Id" & vbCrLf & _
                  "           AND NFI.Produto_Id      = NFE.Produto_Id" & vbCrLf & _
                  "           AND NFI.Sequencia_Id    = NFE.Sequencia_Id " & vbCrLf & _
                  "           AND NFE.Encargo_Id IN ('ISS', 'PRODUTO')" & vbCrLf & _
                  "         INNER JOIN SubOperacoes SO" & vbCrLf & _
                  "            ON NF.Operacao    = SO.Operacao_Id" & vbCrLf & _
                  "           AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                  "         INNER JOIN Clientes" & vbCrLf & _
                  "            ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf & _
                  "           AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                  "         WHERE (NOT (NF.Serie_Id IN ('D', 'F', 'REC')))" & vbCrLf & _
                  "           AND (NF.Empresa_Id = '" & Campo(0) & "')  " & vbCrLf & _
                  "           AND (NF.Movimento BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "') " & vbCrLf & _
                  "           AND (" & vbCrLf & _
                  "                   (OE.CodigoFiscal BETWEEN 1350 AND 1360)" & vbCrLf & _
                  "                OR (OE.CodigoFiscal BETWEEN 2350 AND 2360)" & vbCrLf & _
                  "                OR (OE.CodigoFiscal BETWEEN 3350 AND 3360)" & vbCrLf & _
                  "                OR (OE.CodigoFiscal BETWEEN 5350 AND 5360)" & vbCrLf & _
                  "                OR (OE.CodigoFiscal BETWEEN 6350 AND 6360) " & vbCrLf & _
                  "                OR (OE.CodigoFiscal BETWEEN 7350 AND 7360)" & vbCrLf & _
                  "               ) " & vbCrLf & _
                  "            And (NF.Situacao = 1)" & vbCrLf & _
                  "            AND Not (NF.EntradaSaida_Id = 'E' AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf & _
                  "            AND Not (SO.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "'   AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf & _
                  "            AND Not (NF.EntradaSaida_Id = 'S' AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf & _
                  "            AND Not (SO.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "'   AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf & _
                  "          GROUP BY NF.Nota_Id, NF.EntradaSaida_Id, NF.Serie_Id, NF.Cliente_Id, NF.Movimento, " & vbCrLf & _
                  "                   NF.DataDaNota, OE.CodigoFiscal, SO.GrupoDeContas, NF.EstadoDoCliente, Clientes.Inscricao" & vbCrLf & _
                  "        ) AS Consulta" & vbCrLf & _
                  " ORDER BY EntradaSaida, Movimento, Serie, NumeroDaNota" & vbCrLf
            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "70"                                                                                                'Tipo
                        linha &= Funcoes.AlinharDireita(.Item("Cliente"), 14, "0")                                                       'Cliente
                        linha &= Funcoes.AlinharEsquerda(IIf(.Item("Inscricao").trim = "", "ISENTO", .Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "")), 14, " ") 'Inscricao Estadual
                        linha &= Format(dr("Movimento"), "yyyyMMdd")                                                                'Movimento                   
                        linha &= Funcoes.AlinharEsquerda(.Item("Estado"), 2, " ")                                                        'Estado
                        linha &= "08"                                                                                               'Série
                        Dim Ser = Funcoes.ValidarSerie(.Item("Serie"), .Item("EntradaSaida"))

                        linha &= "U  "                                                                       'Subserie
                        linha &= Funcoes.AlinharDireita(.Item("Nota"), 6, "0")                                                           'Nota
                        linha &= Funcoes.AlinharDireita(.Item("CFOP"), 4, "0")                                                      'Cfop

                        linha &= Funcoes.AlinharDireita(CStr(.Item("Valor")).Replace(",", ""), 13, "0")                          'Valor da Nota
                        linha &= Funcoes.AlinharDireita(CStr(.Item("BaseICMS")).Replace(",", ""), 14, "0")                                  'Base de Calculo
                        linha &= Funcoes.AlinharDireita(CStr(.Item("ValorICMS")).Replace(",", ""), 14, "0")                                      'ICMS
                        linha &= "00000000000000"                                                                                    'Isentas
                        Dim Outras As Decimal = .Item("Valor") - .Item("BaseICMS")                                       'Calculo
                        linha &= Funcoes.AlinharDireita(CStr(Outras).Replace(",", ""), 14, "0")                                             'Outras
                        linha &= "2"                                                                                                'Situacao
                        linha &= "N"                                                                                                'Situacao
                    End With

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Registro70 += 1

                Next
            End If


            'Registro Tipo 74 Registro de Inventario
            If Format(CDate(txtPeriodoFinal.Text), "MM") = 2 Then

                Sql = " SELECT  'Reg-74' as Tipo,   " & vbCrLf & _
                      " 		ApuracaoDeCustos.Mes_Id, " & vbCrLf & _
                      " 		ApuracaoDeCustos.Ano_Id, " & vbCrLf & _
                      "        ApuracaoDeCustos.Produto_Id," & vbCrLf & _
                      "        ApuracaoDeCustos.Quantidade, " & vbCrLf & _
                      "        ApuracaoDeCustos.ValorDoProduto, " & vbCrLf & _
                      "        '0' as CodigoDoInventario," & vbCrLf & _
                      "        Clientes.Cliente_Id, " & vbCrLf & _
                      "        Clientes.Inscricao, " & vbCrLf & _
                      "        Clientes.Estado" & vbCrLf & _
                      " FROM   ApuracaoDeCustos INNER JOIN  Clientes ON" & vbCrLf & _
                      "        ApuracaoDeCustos.Empresa_Id = Clientes.Cliente_Id AND " & vbCrLf & _
                      "        ApuracaoDeCustos.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf & _
                      " WHERE  Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
                      "        And ApuracaoDeCustos.CodigoDeCusto_Id = 920 " & vbCrLf & _
                      "        And Quantidade <> 0 " & vbCrLf & _
                      "        And Mes_Id =  12" & vbCrLf & _
                      "        And Ano_Id = " & Format(CDate(txtPeriodoInicial.Text), "yyyy") - 1 & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "74"
                            linha &= dr("Ano_Id") & Funcoes.AlinharDireita(.Item("Mes_Id"), 2, "0") & Format(CDate(txtPeriodoFinal.Text), "dd")     'Data                                                            'Movimento                   
                            linha &= Funcoes.AlinharEsquerda(.Item("Produto_Id"), 14, " ")
                            'Produto
                            Dim Valor As String = Microsoft.VisualBasic.Format(.Item("Quantidade"), "0000000000.000")

                            linha &= Valor.Replace(",", "")                                                     'Quantidade
                            linha &= Funcoes.AlinharDireita(CStr(.Item("ValorDoProduto")).Replace(",", ""), 13, "0")                                'Valor do Produto
                            linha &= "1"                                                                        'Posse
                            If Campo(0) = .Item("Cliente_Id") Then
                                linha &= "00000000000000"
                            Else
                                linha &= Funcoes.AlinharDireita(.Item("Cliente_Id"), 14, "0")                                                              'Cliente
                            End If
                            'linha &= Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " ")        'Inscrição
                            linha &= Funcoes.AlinharEsquerda(" ", 14, " ")        'Inscrição
                            linha &= Funcoes.AlinharEsquerda(.Item("Estado"), 2, " ")                                                               'Estado
                            linha &= Funcoes.AlinharDireita(" ", 45, " ")                                                                           'Brancos
                        End With

                        strm = New StreamWriter(NomeArquivo, True)
                        strm.WriteLine(linha)
                        strm.Close()
                        Registro74 += 1
                    Next
                End If
            End If


            'Registro Tipo 75
            Sql = " SELECT DISTINCT Produtos.Produto_Id, Produtos.NCM, Produtos.Nome, Produtos.Unidade " & vbCrLf & _
                  " FROM  Produtos  " & vbCrLf & _
                  " LEFT JOIN NotasFiscaisXItens  " & vbCrLf & _
                  " ON Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id  " & vbCrLf & _
                  " LEFT JOIN NotasFiscais  " & vbCrLf & _
                  " ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id  " & vbCrLf & _
                  " AND NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
                  " AND NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id  " & vbCrLf & _
                  " AND NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id  " & vbCrLf & _
                  " AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
                  " AND NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id  " & vbCrLf & _
                  " AND NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id  " & vbCrLf & _
                  " LEFT JOIN SubOperacoes  " & vbCrLf & _
                  " ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id  " & vbCrLf & _
                  " AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) AND " & vbCrLf & _
                   "       (NotasFiscais.Empresa_Id = '" & Campo(0) & "') AND " & vbCrLf & _
                   "       (NotasFiscais.Movimento BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "') AND " & vbCrLf & _
                   "       (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1350 AND 1360) AND" & vbCrLf & _
                   "       (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 2350 AND 2360) AND" & vbCrLf & _
                   "       (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 3350 AND 3360) AND" & vbCrLf & _
                   "       (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 5350 AND 5360) AND" & vbCrLf & _
                   "       (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 6350 AND 6360) AND" & vbCrLf & _
                   "       (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 7350 AND 7360) AND" & vbCrLf & _
                   "       (NotasFiscais.Situacao = 1)" & vbCrLf & _
                   "        AND Not (NotasFiscais.EntradaSaida_Id = 'E'   " & vbCrLf & _
                   "        AND  NotasFiscaisXItens.CFOP_Id IN (1933,2933)    " & vbCrLf & _
                   "       ) AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NotasFiscaisXItens.CFOP_Id IN (1949,2949)) " & vbCrLf & _
                   "        AND Not (NotasFiscais.EntradaSaida_Id = 'S'   " & vbCrLf & _
                   "        AND  NotasFiscaisXItens.CFOP_Id IN (5933,6933)    " & vbCrLf & _
                   "       ) AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NotasFiscaisXItens.CFOP_Id IN (5949,6949)) " & vbCrLf & _
                   "  OR exists (Select  1 from #notascompostas where " & vbCrLf & _
                   "     #NotasCompostas.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                   "     AND #NotasCompostas.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                   "     AND #NotasCompostas.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                   "     AND #NotasCompostas.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                   "     AND #NotasCompostas.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                   "     AND #NotasCompostas.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                   "     AND #NotasCompostas.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                   "              )                                  " & vbCrLf

            If Format(CDate(txtPeriodoFinal.Text), "MM") = 2 Then
                Sql &= " OR Produtos.Produto_Id in( SELECT  ApuracaoDeCustos.Produto_Id " & vbCrLf & _
                       " FROM   ApuracaoDeCustos INNER JOIN  Clientes ON" & vbCrLf & _
                       "        ApuracaoDeCustos.Empresa_Id = Clientes.Cliente_Id AND " & vbCrLf & _
                       "        ApuracaoDeCustos.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf & _
                       " WHERE  Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
                       "        And ApuracaoDeCustos.CodigoDeCusto_Id = 920 " & vbCrLf & _
                       "        And Quantidade <> 0 " & vbCrLf & _
                       "        And Mes_Id =  12" & vbCrLf & _
                       "        And Ano_Id = " & Format(CDate(txtPeriodoInicial.Text), "yyyy") - 1 & vbCrLf & _
                       "                                       )          " & vbCrLf
            End If
            ds = Banco.ConsultaDataSet(Sql1 + Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "75"
                        linha &= Format(CDate(txtPeriodoInicial.Text), "yyyyMMdd")
                        linha &= Format(CDate(txtPeriodoFinal.Text), "yyyyMMdd")

                        linha &= Funcoes.AlinharEsquerda(.Item("Produto_Id"), 14, " ")
                        linha &= Funcoes.AlinharEsquerda(.Item("NCM"), 8, " ")
                        linha &= Funcoes.AlinharEsquerda(.Item("Nome"), 53, " ")
                        linha &= Funcoes.AlinharEsquerda(.Item("Unidade"), 6, " ")
                        linha &= "00000"
                        linha &= "0000"
                        linha &= "00000"
                        linha &= "0000000000000"
                    End With

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Registro75 += 1

                Next
            End If

            If HttpContext.Current.Session("ssEstadoEmpresa") = "MS" Then

                'Registro Tipo 88   CONTADOR

                Sql = "  SELECT Clientes.Endereco, ISNULL(Clientes.Numero, '0') AS Numero, ISNULL(Clientes.Complemento, ' ') AS Complemento, Clientes.Bairro, Clientes.Cep, " & vbCrLf & _
                      " ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, ClientesXEmpresas.CRCContador, Clientes.Telefone, Clientes.Email " & vbCrLf & _
                      " FROM Clientes LEFT OUTER JOIN" & vbCrLf & _
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                      " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id = '" & Campo(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "88"                                                                                                            'Tipo
                            linha &= "EC"                                                                                                            'SubTipo
                            linha &= Funcoes.AlinharEsquerda(.Item("NomeDoContador"), 39, " ")                                                      'NomeContador
                            linha &= Funcoes.AlinharEsquerda(.Item("CPFContador"), 11, "0")                                                         'CPFContador
                            linha &= Funcoes.AlinharEsquerda(.Item("CRCContador".Replace("/", "").Replace(".", "").Replace("-", "").Replace(" ", "")), 10, " ")  'CRCContador
                            linha &= Funcoes.AlinharDireita(.Item("Telefone").Replace("(", "").Replace(")", "").Replace("-", ""), 11, "0")          'Tel/Fax
                            linha &= Funcoes.AlinharEsquerda(.Item("Email"), 50, " ")                                                               'Email
                            linha &= "1"                                                                                                             'Indicador de alteracao cadastro 0-Sim 1-Não
                        End With

                        strm = New StreamWriter(NomeArquivo, True)
                        strm.WriteLine(linha)
                        strm.Close()
                        Registro88 += 1
                    Next
                End If

                'Registro Tipo 88   SOFTWARE

                'Sql = "  SELECT Clientes.Endereco, ISNULL(Clientes.Numero, '0') AS Numero, ISNULL(Clientes.Complemento, ' ') AS Complemento, Clientes.Bairro, Clientes.Cep, "
                'Sql &= " Clientes.Nome, Clientes.Cliente_id, ClientesXEmpresas.CRCContador, Clientes.Telefone, Clientes.Email "
                'Sql &= " FROM Clientes LEFT OUTER JOIN"
                'Sql &= " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
                'Sql &= " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
                'Sql &= " WHERE Clientes.Cliente_Id = '" & Campo(0) & "' And Clientes.Endereco_Id = 0"

                'ds = Banco.ConsultaDataSet(Sql, "Clientes")

                'If ds.Tables(0).Rows.Count > 0 Then
                '    For Each dr In ds.Tables(0).Rows
                '        With dr
                linha = "88"                                                                                                            'Tipo
                linha &= "SF"                                                                                                            'SubTipo
                linha &= Funcoes.AlinharEsquerda("NGS SOLUCOES EM SISTEMAS DE INFORMATICA LTDA ", 35, " ")                              'NomeEmpresaDesenv
                linha &= Funcoes.AlinharEsquerda("12420302000106", 14, "0")                                                             'CNPJEmpresa
                linha &= Funcoes.AlinharEsquerda("33263558900", 11, " ")                                                                'CpfTecnico
                linha &= Funcoes.AlinharDireita("04130187175", 11, "0")                                                                 'Tel/Fax
                linha &= Funcoes.AlinharEsquerda("neri_galvan@hotmail.com", 50, " ")                                                    'Email
                linha &= "1"                                                                                                             'Indicador de alteracao cadastro 0-Sim 1-Não
                '        End With

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                Registro88 += 1
                '    Next
                'End If

            End If

            '----Registros Tipo 90-------------------------------------------------------------------

            'Registro90 = Registro50 + Registro54 + Registro70 + Registro75

            Registro90 = 1
            If Registro50 <> 0 Then Registro90 += 1
            If Registro54 <> 0 Then Registro90 += 1
            If Registro70 <> 0 Then Registro90 += 1
            If Registro71 <> 0 Then Registro90 += 1
            If Registro74 <> 0 Then Registro90 += 1
            If Registro75 <> 0 Then Registro90 += 1
            If Registro88 <> 0 Then Registro90 += 1

            If Registro50 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "50"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro50, 8, "0")     'Total de Registro tipo 50                                             'Nota
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                      'Municipio

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            If Registro54 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "54"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro54, 8, "0")     'Total de Registro tipo 54                                             'Nota
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                      'Municipio

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            If Registro70 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "70"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro70, 8, "0")     'Total de Registro tipo 70                                             'Nota
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                      'Municipio

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            If Registro71 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "71"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro71, 8, "0")     'Total de Registro tipo 71                                             
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                     
                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            If Registro74 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "74"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro74, 8, "0")     'Total de Registro tipo 75                                            
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                     

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            If Registro75 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "75"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro75, 8, "0")     'Total de Registro tipo 75                                             'Nota
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                      'Municipio

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            If Registro88 <> 0 Then
                linha = "90"                                    'Tipo 90 
                linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
                linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
                linha &= "88"                                   'Tipo do Registro
                linha &= Funcoes.AlinharDireita(Registro88, 8, "0")     'Total de Registro tipo 88                                            'Nota
                linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
                linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                      'Municipio

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If

            '------Registro de Encerramento Tipo 90
            Registro99 = Registro10 + Registro11 + Registro50 + Registro54 + Registro70 + Registro71 + Registro74 + Registro75 + Registro88 + Registro90

            linha = "90"                                    'Tipo 90 
            linha &= Funcoes.AlinharEsquerda(CGC, 14, " ")          'CGC da Empresa                                                
            linha &= Funcoes.AlinharEsquerda(Inscricao, 14, " ")    'Inscriçao da Empresa
            linha &= "99"                                   'Tipo do Registro
            linha &= Funcoes.AlinharDireita(Registro99, 8, "0")     'Total de Registro tipo 75                                             'Nota
            linha &= Funcoes.AlinharEsquerda(" ", 85, " ")          'Espaço Livre
            linha &= Funcoes.AlinharEsquerda(Registro90, 1, "0")    'Total de Registro Tipo 90                                                      'Municipio

            strm = New StreamWriter(NomeArquivo, True)
            strm.WriteLine(linha)
            strm.Close()

            '----------------------------------------------------------------------------------------
            'MsgBox(Me.Page, "Registros Fiscais Sintegra")
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString))
        End Try
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImdBaixarArquivo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnkBaixarArquivo As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnkBaixarArquivo.NamingContainer, GridViewRow)
            Dim NomeArquivo As String = ""
            Dim emp As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim PI As String() = grdProcesso.Rows(row.RowIndex).Cells(2).Text.Split("/")
            Dim PF As String() = grdProcesso.Rows(row.RowIndex).Cells(3).Text.Split("/")

            NomeArquivo = emp(0).Trim & "-" & PI(2) & PI(1) & PI(0) & "-" & PF(2) & PF(1) & PF(0) & ".txt"
            Download(NomeArquivo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Download(ByVal NomeArquivo As String)
        Dim CaminhoCompleto As String = Server.MapPath("~/Sintegra/" & NomeArquivo)
        If File.Exists(CaminhoCompleto) And Not String.IsNullOrEmpty(CaminhoCompleto) Then
            Try
                Dim filePath As String = "~/Sintegra/" & NomeArquivo
                Response.AppendHeader("Content-Disposition", "attachment; filename=" & NomeArquivo)
                Response.ContentType = "text/plain"
                Const bufferLength As Integer = 10000
                Dim buffer As Byte() = New Byte(bufferLength - 1) {}
                Dim length As Integer = 0
                Dim download As Stream = Nothing

                Try
                    download = New FileStream(Server.MapPath("~/Sintegra/" & NomeArquivo), FileMode.Open, FileAccess.Read)
                    Do
                        If Response.IsClientConnected Then
                            length = download.Read(buffer, 0, bufferLength)
                            Response.OutputStream.Write(buffer, 0, length)
                            buffer = New Byte(bufferLength - 1) {}
                        Else
                            length = -1
                        End If
                    Loop While length > 0
                    Response.Flush()
                    Response.End()
                Finally
                    If download IsNot Nothing Then
                        download.Close()
                    End If
                End Try
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message.ToString)
            End Try
        Else
            MsgBox(Me.Page, "Arquivo não encontrado.")
        End If
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("RegistrosFiscaisSintegra", "GRAVAR") Then
                Dim sqlList As ArrayList = New ArrayList
                Dim Registro As String = ddlEmpresa.SelectedValue
                Dim Campo = Registro.Split("-")


                If txtProcesso.Text.Length > 0 Then
                    Sql = " SELECT *"
                    Sql &= "  FROM ProcessoSintegra"
                    Sql &= " WHERE Empresa_Id = '" & Campo(0) & "' and Processo_Id = " & txtProcesso.Text & ""

                    Dim dsProcesso As DataSet = Banco.ConsultaDataSet(Sql, "ProcessoSintegra")

                    If dsProcesso.Tables(0).Rows.Count = 0 Then
                        Sql = "INSERT INTO ProcessoSintegra " & vbCrLf & _
                              "(Empresa_Id " & vbCrLf & _
                              ", Processo_Id" & vbCrLf & _
                              ", PeriodoInicial" & vbCrLf & _
                              ", PeriodoFinal" & vbCrLf & _
                              ", Convenio" & vbCrLf & _
                              ", NaturezaDaOperacao" & vbCrLf & _
                              ", ArquivoMagnetico" & vbCrLf & _
                              ", Fechado, ArquivoDeSaida) " & vbCrLf

                        Sql &= "VALUES ('" & Campo(0) & "', " & txtProcesso.Text & ", '" & txtPeriodoInicial.Text.ToSqlDate() & "', " & vbCrLf & _
                        "'" & txtPeriodoFinal.Text.ToSqlDate() & "', " & ddlConvenio.SelectedValue & ", " & ddlNaturezaDaOperacao.SelectedValue & ", " & vbCrLf & _
                        "" & ddlArquivoMagnetico.SelectedValue & ", 'N', '" & "Sintegra/" & Campo(0) & "-" & Format(CDate(txtPeriodoInicial.Text), "yyyyMMdd") & "-" & Format(CDate(txtPeriodoFinal.Text), "yyyyMMdd") & ".txt" & "')" & vbCrLf
                    Else
                        Sql = "UPDATE ProcessoSintegra" & vbCrLf & _
                              "  SET PeriodoInicial = CONVERT(DATETIME, '" & txtPeriodoInicial.Text & "', 103)," & vbCrLf & _
                              "      PeriodoFinal = CONVERT(DATETIME, '" & txtPeriodoFinal.Text & "', 103)," & vbCrLf & _
                              "      Convenio = " & ddlConvenio.SelectedValue & "," & vbCrLf & _
                              "      NaturezaDaOperacao = " & ddlNaturezaDaOperacao.SelectedValue & "," & vbCrLf & _
                              "      ArquivoMagnetico = " & ddlArquivoMagnetico.SelectedValue & " " & vbCrLf & _
                              "WHERE Empresa_Id = '" & Campo(0) & "' and " & vbCrLf & _
                              "      Processo_Id = " & txtProcesso.Text & "" & vbCrLf
                    End If
                    sqlList.Add(Sql)

                    If Banco.GravaBanco(sqlList) Then
                        GeraSintegra()
                        CargaProcessos()
                    Else
                    End If

                    Dim i As Integer
                    For i = 0 To grdProcesso.Rows.Count - 1
                        If grdProcesso.Rows(i).Cells(1).Text = txtProcesso.Text Then
                            CType(grdProcesso.Rows(i).FindControl("imdBaixarArquivo"), LinkButton).Visible = True
                        Else
                            CType(grdProcesso.Rows(i).FindControl("imdBaixarArquivo"), LinkButton).Visible = False
                        End If
                    Next
                    ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
                Else
                    MsgBox(Me.Page, "Numero do processo obrigatorio.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RegistrosFiscaisSintegra")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class