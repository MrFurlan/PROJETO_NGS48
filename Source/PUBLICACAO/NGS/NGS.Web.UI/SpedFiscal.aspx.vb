Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Diagnostics.Eventing.Reader

Partial Class SpedFiscal
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim Sql As String
    Dim Empresa() As String
    Dim Cliente() As String
    Dim Codigo As String = ""
    Dim Descricao As String = ""
    Dim Crc As String
    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim IcmsARecolher As Decimal

    ''' <summary>
    ''' Controle da Produção e de Estoque
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure BlocoK

        ''' <summary>
        ''' Abertura do Bloco K
        ''' </summary>
        ''' <remarks>Este registro deve ser gerado para abertura do bloco K, indicando se há registros de informações no bloco.</remarks>
        Public RegistroK001 As String

        Public RegistroK010 As String

        ''' <summary>
        ''' Período de Apuração do ICMS/IPI
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar o período de apuração do ICMS ou do IPI, prevalecendo os períodos mais curtos.
        '''          Contribuintes com mais de um período de apuração no mês declaram um registro K100 para cada período no mesmo arquivo.
        '''          Não podem ser informados dois ou mais registros com os mesmos campos DT_INI e DT_FIN.</remarks>
        Public RegistroK100 As String

        ''' <summary>
        ''' Estoque Escriturado
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar o estoque final escriturado do período de apuração informado no registro K100,
        '''          por tipo de estoque e por participante, nos casos em que couber, das mercadorias de tipo 00 - Mercadoria para revenda,
        '''          01 - Matéria-Prima, 02 - Embalagem, 03 - Produtos em processo, 04 - Produto Acabado, 05 - Subproduto e 10 - Outros insumos - 
        '''          Campo TIPO_ITEM do registro 0200.
        '''          A unidade de medida é, obrigatoriamente, a de controle de estoque constante no campo 06 do registro 0200 - UNID_INV.
        '''          A chave deste registro são os campos: DT_EST, COD_ITEM, IND_EST e COD_PART (este quando houver).</remarks>
        Public RegistroK200 As String

        ''' <summary>
        ''' Outras Movimentações Internas Entre Mercadorias.
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar a movimentação interna entre mercadorias, que não se enquadre nas movimentações internas já informadas nos Registros
        '''          k230 - Itens Produzidos e  k235 - Insumos Consumidos: Produção acabada e consumo no processo produtivo, respectivamente.
        '''          Exemplo: reclassificação de um produto em outro códio em função do cliente a que se destina.
        '''          A unidade de medida é, obrigatóriamente, aquela constante no campo 06 do registro 0200: UNID_INV.
        '''          A quantidade movimentada deve ser expressa na unidade de medida do item de origem.</remarks>
        Public RegistroK220 As String

        ''' <summary>
        ''' Itens Produzidos
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar a produção acabada de produto em processo (tipo 03 - campo TIPO_ITEM do registro 0200)
        ''' e produto acabado (tipo 04 - campo TIPO_ITEM do registro 0200). Deverá existir mesmo que a quantidade de produção acabada seja igual a zero,
        ''' nas situações em que exista o consumo de item componente/insumo no registro filho K235. Nessa situação a produção ficou em elaboração.
        '''          Essa produção em elaboração não é quantificada, uma vez que a matéria não é mais um insumo e nem ainda um produto resultante.
        '''          A unidade de medida é, obrigatóriamente,  a de controle de estoque constante no campo 06 do registro 0200: UNID_INV.
        '''          Qunado houver identificação da ordem de produção não for identificada, a chave deste registro são os campos: COD_DOC_OP e COD_ITEM.
        '''          Nos casos em que a ordem de produção não for identificada, o campo chave passa a ser COD_ITEM.
        ''' </remarks>
        Public RegistroK230 As String

        ''' <summary>
        ''' Insumos Consumidos
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar o consumo de mercadoria no processo produtivo, vinculado ao produto resultante informado no campo COD_ITEM
        ''' do Registro K230 - Itens Produzidos.
        '''          A unidade de medida é, obrigatóriamente, a de controle de estoque constante no campo 06 do registro 0200: UNID_INV.
        '''          A chave deste registro são os campos DT_SAIDA e COD_ITEM.</remarks>
        Public RegistroK235 As String

        ''' <summary>
        ''' Industrialização Efetuada Por Terceiros - Itens Produzidos
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar os produtos que foram industrializados pro terceiros e sua quantidade.
        '''          A unidade de medida é, obrigatoriamente, a de controle de estoque no campo 06 do registro 0200: UNID_INV.
        '''          A chave deste registro são os campos DT_PROD e COD_ITEM.</remarks>
        Public RegistroK250 As String

        ''' <summary>
        ''' Industrialização Em Terceiros - Insumos Consumidos
        ''' </summary>
        ''' <remarks>Este registro tem o objetivo de informar a quantidade de consumo do insumo que foi remetido para ser industrializado em terceiro,
        ''' viculado ao produto resultante informado no campo COD_ITEM do registro K250.
        '''          A unidade de medida é, obrigatoriamente, a de controle de estoque constante no campo 06 do registro 0200: UNID_INV.
        '''          A chave desste registro são os campos DT_CONS e COD_ITEM deste registro.</remarks>
        Public RegistroK255 As String

        ''' <summary>
        ''' Encerramento do Bloco
        ''' </summary>
        ''' <remarks>Este registro destina-se a identificar o encerramento do Bloco K e a informar a quantidade de linhas(registros) existentes no bloco.</remarks>
        Public registroK990 As String

    End Structure

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Fiscal.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("SpedFiscal", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("~/SpedFiscal")) Then Directory.CreateDirectory(Server.MapPath("~/SpedFiscal"))
                    CargaUnidade()
                    VerificaUnidade()
                    imdDownload.Visible = False
                    DdlPerfil.SelectedIndex = 0
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
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
        CarregarProcesso()
        CarregarProcessoIPI()
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProcesso()
            CarregarProcessoIPI()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlProcesso_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            HttpContext.Current.Session("ProcessoIcms") = DdlProcesso.SelectedValue

            Sql = "SELECT * " & vbCrLf & _
                  "  FROM ProcessoRAICMS " & vbCrLf & _
                  " WHERE Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf & _
                  "   And Processo_Id = " & DdlProcesso.SelectedValue & vbCrLf & _
                  " order by PeriodoInicial desc " & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                txtDataInicial.Text = Dr("PeriodoInicial")
                txtDataFinal.Text = Dr("PeriodoFinal")
                txtLivro.Text = Dr("Livro")
                txtFolha.Text = Dr("PaginaInicial")
            Next

            Empresa = DdlEmpresa.SelectedValue.Split("-")
            txtArquivoDeSaida.Text = "SpedFiscalMes-" & txtDataInicial.Text.ToSqlDate() & "-" & Empresa(0) & ".txt"
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIcms") = Empresa(0)
        DdlProcesso.Items.Clear()

        Sql = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & vbCrLf & _
              "  FROM ProcessoRAICMS " & vbCrLf & _
              " WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
              " order by PeriodoInicial desc " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  "
            Descricao &= CDate(Dr("PeriodoInicial")).ToString("dd/MM/yyyy") & "  à  "
            Descricao &= CDate(Dr("PeriodoFinal")).ToString("dd/MM/yyyy") & "  Livro  "
            Descricao &= Dr("Livro").ToString("000") & "  Folha  "
            Descricao &= Dr("PaginaInicial").ToString("000")

            DdlProcesso.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcesso.Items.Insert(0, "")
        DdlProcesso.SelectedIndex = 0
    End Sub

    Private Sub CarregarProcessoIPI()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIcms") = Empresa(0)
        DdlProcessoIPI.Items.Clear()

        Sql = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & vbCrLf & _
              "  FROM ProcessoRAIPI " & vbCrLf & _
              " WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
              " order by PeriodoInicial desc " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  "
            Descricao &= CDate(Dr("PeriodoInicial")).ToString("dd/MM/yyyy") & "  à  "
            Descricao &= CDate(Dr("PeriodoFinal")).ToString("dd/MM/yyyy") & "  Livro  "
            Descricao &= Dr("Livro").ToString("000") & "  Folha  "
            Descricao &= Dr("PaginaInicial").ToString("000")

            DdlProcessoIPI.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcessoIPI.Items.Insert(0, "")
        DdlProcessoIPI.SelectedIndex = 0
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            MsgBox(Me.Page, "Informe a unidade de negocio.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe o periodo inicial.")
            Return False
        ElseIf Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Valor informado no período inicial não é uma data válida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período final.")
            Return False
        ElseIf Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Valor informado no período final não é uma data válida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtLivro.Text) Then
            MsgBox(Me.Page, "Informe o número do livro.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtFolha.Text) Then
            MsgBox(Me.Page, "Informe o número da folha.")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        VerificaUnidade()
        txtDataInicial.Text = String.Empty
        txtDataFinal.Text = String.Empty
        txtLivro.Text = String.Empty
        txtFolha.Text = String.Empty
        txtArquivoDeSaida.Text = String.Empty
        imdDownload.Visible = False
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub imdDownload_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & txtArquivoDeSaida.Text)
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(Server.MapPath("~/SpedFiscal/" & txtArquivoDeSaida.Text), FileMode.Open, FileAccess.Read)
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
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub PreencherTabelas(ByRef ds As DataSet)
        Dim sql As String
        '***********************************************************************************
        '********************************** BLOCO C ****************************************
        '***********************************************************************************
        sql = "/*Bloco C*/" & vbCrLf & _
              "SELECT count(*)" & vbCrLf & _
              "  From (" & vbCrLf & _
              "        SELECT NF.EntradaSaida_Id," & vbCrLf & _
              "               OE.CodigoFiscal as CFOP" & vbCrLf & _
              "          FROM NotasFiscais nf" & vbCrLf & _
              "         Inner join Notasfiscaisxitens nfi " & vbCrLf & _
              "            on nf.Empresa_id      = nfi.Empresa_id" & vbCrLf & _
              "           and nf.EndEmpresa_id   = nfi.EndEmpresa_id" & vbCrLf & _
              "           and nf.cliente_id      = nfi.cliente_id" & vbCrLf & _
              "           and nf.endcliente_id   = nfi.endcliente_id" & vbCrLf & _
              "           and nf.entradasaida_id = nfi.entradasaida_id" & vbCrLf & _
              "           and nf.nota_id         = nfi.nota_id" & vbCrLf & _
              "           and nf.serie_id        = nfi.serie_id" & vbCrLf & _
              "         Inner Join OperacaoxEstado OE" & vbCrLf & _
              "            on OE.Codigo_id = nfi.OperacaoXEstado" & vbCrLf & _
              "         WHERE (NOT (NF.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf & _
              "           and NF.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
              "           AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
              "       ) as Consulta" & vbCrLf & _
              " Where (CFOP Not Between 1350 and 1360) " & vbCrLf & _
              "   and (CFOP Not Between 2350 and 2360) " & vbCrLf & _
              "   and (CFOP Not in(1932,2932)) " & vbCrLf & _
              "   and (CFOP Not Between 5350 and 5360) " & vbCrLf & _
              "   and (CFOP Not Between 6350 and 6360) " & vbCrLf & _
              "   and (CFOP Not Between 1302 and 1303) " & vbCrLf

        '***********************************************************************************
        '********************************** BLOCO D ****************************************
        '***********************************************************************************
        sql &= "/*Bloco D*/" & vbCrLf & _
               "SELECT count(*)" & vbCrLf & _
               "  From (" & vbCrLf & _
               "         SELECT nf.EntradaSaida_Id," & vbCrLf & _
               "                OE.CodigoFiscal as CFOP" & vbCrLf & _
               "		   FROM NotasFiscais nf" & vbCrLf & _
               "         Inner join Notasfiscaisxitens nfi" & vbCrLf & _
               "            on nf.Empresa_id      = nfi.Empresa_id" & vbCrLf & _
               "           and nf.EndEmpresa_id   = nfi.EndEmpresa_id" & vbCrLf & _
               "           and nf.cliente_id      = nfi.cliente_id" & vbCrLf & _
               "           and nf.endcliente_id   = nfi.endcliente_id" & vbCrLf & _
               "           and nf.entradasaida_id = nfi.entradasaida_id" & vbCrLf & _
               "           and nf.nota_id         = nfi.nota_id" & vbCrLf & _
               "           and nf.serie_id        = nfi.serie_id" & vbCrLf & _
               "          Inner Join OperacaoxEstado OE" & vbCrLf & _
               "             on OE.Codigo_id = nfi.OperacaoXEstado" & vbCrLf & _
               "		  WHERE (NOT (nf.Serie_Id IN ('D', 'F','REC')))" & vbCrLf & _
               "			and nf.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
               "           AND nf.Movimento  BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
               "         ) as Consulta" & vbCrLf & _
               " Where (CFOP Between 1350 and 1360)" & vbCrLf & _
               "    Or (CFOP Between 2350 and 2360)" & vbCrLf & _
               "    Or (CFOP in (1932,2932)) " & vbCrLf & _
               "    Or (CFOP Between 5350 and 5360)" & vbCrLf & _
               "    Or (CFOP Between 6350 and 6360)" & vbCrLf & _
               "    or (CFOP Between 1300 and 1350)" & vbCrLf & _
               "    Or (CFOP Between 2300 and 2350)" & vbCrLf & _
               "    Or (CFOP Between 5300 and 5350)" & vbCrLf & _
               "    Or (CFOP Between 6300 and 6350)" & vbCrLf







        '***********************************************************************************
        '************************************* H005 ****************************************
        '************************************* H010 ****************************************
        '***********************************************************************************
        'Inventario
        If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
            sql &= "/*Inventario*/" & vbCrLf & _
                   "Select sb.Empresa_id," & vbCrLf & _
                   "       Sb.EndEmpresa_id," & vbCrLf & _
                   "       Emp.Nome as NomeEmpresa," & vbCrLf & _
                   "       Emp.Complemento," & vbCrLf & _
                   "       Emp.Cidade," & vbCrLf & _
                   "       Emp.Estado," & vbCrLf & _
                   "       sb.Classificacao as IND_PROP," & vbCrLf & _
                   "       sb.Cliente_id," & vbCrLf & _
                   "       sb.EndCliente_Id," & vbCrLf & _
                   "       cli.Nome as NomeCliente," & vbCrLf & _
                   "       sb.Produto_id," & vbCrLf & _
                   "       Prd.Nome as NomeProduto," & vbCrLf & _
                   "       sb.Conta," & vbCrLf & _
                   "       Prd.Unidade," & vbCrLf & _
                   "       cast(sum(sb.qtdeRazao) as decimal(18,4))   as QtdeRazao," & vbCrLf & _
                   "       cast(Sum(sb.Qtdefiscal) as decimal(18,4))  as QtdeFiscal," & vbCrLf & _
                   "       cast(sum(sb.qtdeRazao + sb.Qtdefiscal)as decimal(18,4)) as Qtde," & vbCrLf & _
                   "       cast(sum(sb.ValorRazao) as decimal(18,2))  as ValorRazao," & vbCrLf & _
                   "       cast(sum(sb.ValorFiscal) as decimal(18,2)) as ValorFiscal," & vbCrLf & _
                   "       cast(Sum(sb.ValorRazao + sb.ValorFiscal)   as decimal(18,2)) as Valor" & vbCrLf & _
                   "  into #Base" & vbCrLf & _
                   "  from (" & vbCrLf & _
                   "   		select 'NF' as Origem," & vbCrLf & _
                   "   		       nfi.Empresa_Id," & vbCrLf & _
                   "   		       nfi.EndEmpresa_id," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				  when SO.Deposito          = 'S' then 1" & vbCrLf & _
                   "   				  when SO.ProdutoDeTerceiro =  1  then 2" & vbCrLf & _
                   "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
                   "   			   end as Classificacao," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                   "   				   then nf.destino" & vbCrLf & _
                   "   				   else nfi.cliente_id" & vbCrLf & _
                   "   			   end Cliente_id," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                   "   				   then nf.Enddestino" & vbCrLf & _
                   "   				   else nfi.Endcliente_id" & vbCrLf & _
                   "   			   end EndCliente_id," & vbCrLf & _
                   "   			   ISNULL(PA.Produto_Id, Nfi.Produto_Id) as Produto_Id," & vbCrLf & _
                   "   			   so.GrupoDeContas as Conta," & vbCrLf & _
                   "   			   nf.pedido," & vbCrLf & _
                   "               convert(numeric(18,4),0) as QtdeRazao," & vbCrLf & _
                   "   			   SUM(Case" & vbCrLf & _
                   "   					 When So.Devolucao = 'S'" & vbCrLf & _
                   "   					   then nfi.QuantidadeFiscal * -1" & vbCrLf & _
                   "   					   else nfi.QuantidadeFiscal" & vbCrLf & _
                   "   				   end) QtdeFiscal," & vbCrLf & _
                   "               convert(numeric(18,4),0) as ValorRazao," & vbCrLf & _
                   "   			   SUM(Case" & vbCrLf & _
                   "   					 When So.Devolucao = 'S'" & vbCrLf & _
                   "   					   then nfe.valor * -1" & vbCrLf & _
                   "   					   else nfe.valor" & vbCrLf & _
                   "   				   end) as valorFiscal" & vbCrLf & _
                   "   		  from notasfiscais NF" & vbCrLf & _
                   "   		 inner join Notasfiscaisxitens NFI" & vbCrLf & _
                   "   			on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
                   "   		   and NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                   "   		   and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                   "   		   and NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                   "   		   and NF.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
                   "   		   and NF.Nota_id         = NFI.Nota_id" & vbCrLf & _
                   "   		   and NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                   "   		 Inner Join NotasFiscaisxEncargos Nfe" & vbCrLf & _
                   "   			on Nfe.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
                   "   		   and Nfe.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                   "   		   and Nfe.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                   "   		   and Nfe.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                   "   		   and Nfe.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
                   "   		   and Nfe.Nota_id         = NFI.Nota_id" & vbCrLf & _
                   "   		   and Nfe.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                   "   		   and Nfe.Produto_id      = NFI.Produto_id" & vbCrLf & _
                   "   		   and Nfe.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
                   "   		   and Nfe.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
                   "   		 Inner join produtos p" & vbCrLf & _
                   "   			on NFI.Produto_id = P.Produto_id" & vbCrLf & _
                   "          LEFT Join ProdutosAgrupados PA" & vbCrLf & _
                   "            ON PA.ProdutoAgrupado_Id = NFI.Produto_Id" & vbCrLf & _
                   "   		 Inner Join Suboperacoes SO" & vbCrLf & _
                   "   			on SO.Operacao_Id        = NFI.Operacao" & vbCrLf & _
                   "   		   and SO.Suboperacoes_Id    = NFI.Suboperacao" & vbCrLf & _
                   "   		 where (SO.produtodeterceiro = 1 or SO.deposito = 'S')" & vbCrLf & _
                   "   		   and SO.Operacao_id <> 80" & vbCrLf & _
                   "           and MONTH(NF.Movimento) <= 12" & vbCrLf & _
                   "           AND YEAR(NF.Movimento) < " & CDate(txtDataFinal.Text).Year & vbCrLf & _
                   "   		   and nf.situacao    = 1" & vbCrLf & _
                   "   		  group by nfi.Empresa_Id," & vbCrLf & _
                   "   		           nfi.EndEmpresa_id," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				  when SO.Deposito          = 'S' then 1" & vbCrLf & _
                   "   				  when SO.ProdutoDeTerceiro = 1   then 2" & vbCrLf & _
                   "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
                   "   			   end," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				 When left(nf.empresa_id, 8) = left(nfi.Cliente_id, 8)" & vbCrLf & _
                   "   				   then nf.destino" & vbCrLf & _
                   "   				   else nfi.cliente_id" & vbCrLf & _
                   "   			   end," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                   "   				   then nf.Enddestino" & vbCrLf & _
                   "   				   else nfi.Endcliente_id" & vbCrLf & _
                   "   			   end," & vbCrLf & _
                   "               nf.pedido," & vbCrLf & _
                   "               ISNULL(PA.Produto_Id, Nfi.Produto_Id)," & vbCrLf & _
                   "               so.GrupoDeContas" & vbCrLf & _
                   "        union all" & vbCrLf & _
                   "   		Select 'R' as Origem," & vbCrLf & _
                   "   		       R.Empresa_Id," & vbCrLf & _
                   "   			   R.Endempresa_id," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then 2" & vbCrLf & _
                   "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then 1" & vbCrLf & _
                   "   			   end Classificacao," & vbCrLf & _
                   "   			   R.Cliente_id," & vbCrLf & _
                   "   			   R.EndCliente_id," & vbCrLf & _
                   "   			   ISNULL(PA.Produto_Id, R.Produto) AS Produto," & vbCrLf & _
                   "   			   R.Conta_Id," & vbCrLf & _
                   "   			   isnull(R.Pedido,0) as Pedido," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then Isnull(R.CreditoQuantidade, 0.0000) - Isnull(R.DebitoQuantidade, 0.0000)" & vbCrLf & _
                   "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoQuantidade, 0.0000)  - Isnull(R.CreditoQuantidade, 0.0000)" & vbCrLf & _
                   "   			   end," & vbCrLf & _
                   "   			   convert(numeric(18,4),0)," & vbCrLf & _
                   "   			   case" & vbCrLf & _
                   "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then isnull(R.CreditoOficial, 0.00) - isnull(R.DebitoOficial, 0.00)" & vbCrLf & _
                   "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoOficial, 0.00)  - isnull(R.CreditoOficial, 0.00)" & vbCrLf & _
                   "   			   end," & vbCrLf & _
                   "               convert(numeric(18,4),0)" & vbCrLf & _
                   "   		  from Razao R" & vbCrLf & _
                   "   		 inner join ClientesxEmpresas CxE" & vbCrLf & _
                   "   			on R.Empresa_id    = CxE.Empresa_Id" & vbCrLf & _
                   "   		   and R.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
                   "   		   and (left(R.Conta_Id,7) = CxE.ContaEstoqueEmNossoPoder or Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) = Left(cxe.ContaestoqueemPoderdeTerceiros,5))" & vbCrLf & _
                   "   		 inner join Clientes C" & vbCrLf & _
                   "   			on R.Cliente_id    = C.Cliente_id" & vbCrLf & _
                   "   		   and R.EndCliente_id = C.Endereco_id" & vbCrLf & _
                   "          LEFT Join ProdutosAgrupados PA" & vbCrLf & _
                   "            ON PA.ProdutoAgrupado_Id = R.Produto" & vbCrLf & _
                   "   		 where lote_id not in (9,10,11)" & vbCrLf & _
                   "           AND YEAR(R.Movimento_Id) < " & CDate(txtDataFinal.Text).Year & vbCrLf & _
                   "   		   and len(isnull(produto,'')) > 0" & vbCrLf & _
                   "   		) as Sb" & vbCrLf & _
                   "     inner join Clientes Emp" & vbCrLf & _
                   "        on Sb.Empresa_id    = Emp.Cliente_id" & vbCrLf & _
                   "       and Sb.EndEmpresa_id = Emp.Endereco_id" & vbCrLf & _
                   "     inner join Clientes Cli" & vbCrLf & _
                   "        on Sb.cliente_id    = Cli.Cliente_id" & vbCrLf & _
                   "       and Sb.EndCliente_id = Cli.Endereco_id" & vbCrLf & _
                   "     inner Join Produtos Prd" & vbCrLf & _
                   "        on Prd.Produto_id = sb.Produto_id" & vbCrLf & _
                   "     Where Prd.TipoDoItem Not in(7,8,9,10,99)" & vbCrLf & _
                   "       and sb.Empresa_id = '" & Empresa(0) & "'" & vbCrLf & _
                   "       and Left(sb.empresa_Id,8) <> left(sb.cliente_Id,8)" & vbCrLf & _
                   "       and sb.Classificacao <> 1 " & vbCrLf
            If Empresa(0).Substring(0, 8) <> "04854422" Then
                sql &= "       AND 1=2" & vbCrLf
            End If
            sql &= "   Group by sb.Empresa_id," & vbCrLf & _
               "          Sb.EndEmpresa_id," & vbCrLf & _
               "          Emp.Nome," & vbCrLf & _
               "          Emp.Complemento," & vbCrLf & _
               "          Emp.Cidade," & vbCrLf & _
               "          Emp.Estado," & vbCrLf & _
               "          sb.Classificacao," & vbCrLf & _
               "          sb.Cliente_id," & vbCrLf & _
               "          sb.EndCliente_Id," & vbCrLf & _
               "          cli.Nome," & vbCrLf & _
               "          sb.Produto_id," & vbCrLf & _
               "          Prd.Nome," & vbCrLf & _
               "          sb.Conta," & vbCrLf & _
               "          Prd.Unidade," & vbCrLf & _
               "          Prd.Nome " & vbCrLf & _
               "   having SUM(sb.QtdeRazao + sb.QtdeFiscal)   > 0  AND Sum(sb.ValorRazao + sb.ValorFiscal) > 0" & vbCrLf & _
               "   order by cli.Nome, sb.Produto_id" & vbCrLf
            '******* H005
            sql &= "/*Registro H005*/" & vbCrLf & _
                   "Select sum(sb.Valor) as Valor " & vbCrLf & _
                   "  From (" & vbCrLf & _
                   "       SELECT isnull(Sum(isnull(ValorDoProduto,0)),0) as Valor" & vbCrLf & _
                   "         FROM ApuracaoDeCustos " & vbCrLf & _
                   "        WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                   "          And CodigoDeCusto_Id in (109, 920) " & vbCrLf & _
                   "          And Quantidade > 0" & vbCrLf & _
                   "          And ValorDoProduto > 0" & vbCrLf & _
                   "          And Mes_Id = 12" & vbCrLf & _
                   "          And Ano_Id = " & Format(CDate(txtDataFinal.Text), "yyyy") - 1 & vbCrLf & _
                   "       Union all" & vbCrLf & _
                   "       Select sum(valor) from #Base" & vbCrLf & _
                   "       ) as sb" & vbCrLf
            '******** H010

            sql &= "/*Registro H010*/" & vbCrLf & _
                   " SELECT ApuracaoDeCustos.Produto_Id AS Produto," & vbCrLf & _
                   "        ApuracaoDeCustos.Quantidade," & vbCrLf & _
                   "        CONVERT(money, ApuracaoDeCustos.ValorDoProduto / ApuracaoDeCustos.Quantidade) AS Unitario," & vbCrLf & _
                   "        ApuracaoDeCustos.ValorDoProduto, " & vbCrLf & _
                   "        CASE " & vbCrLf & _
                   "           --WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 109 THEN 2" & vbCrLf & _
                   "           WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 920 and ApuracaoDeCustos.empresa_id <> Clientes.Cliente_Id then 1 " & vbCrLf & _
                   "           ELSE 0" & vbCrLf & _
                   "        END AS IND_PROP," & vbCrLf & _
                   "        Clientes.Cliente_Id," & vbCrLf & _
                   "        Clientes.Endereco_Id, " & vbCrLf & _
                   "        PlanoDeCustos.DebitoMercadoria," & vbCrLf & _
                   "        PlanoDeCustos.CreditoMercadoria," & vbCrLf & _
                   "        Produtos.Unidade" & vbCrLf & _
                   "   FROM ApuracaoDeCustos " & vbCrLf & _
                   "  INNER JOIN PlanoDeCustos" & vbCrLf & _
                   "     ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id" & vbCrLf & _
                   "  INNER JOIN Produtos" & vbCrLf & _
                   "     ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
                   "  INNER JOIN Clientes" & vbCrLf & _
                   "     ON ApuracaoDeCustos.Deposito_Id    = Clientes.Cliente_Id" & vbCrLf & _
                   "    AND ApuracaoDeCustos.EndDeposito_Id = Clientes.Endereco_Id" & vbCrLf & _
                   "  WHERE ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                   "    And ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920) " & vbCrLf & _
                   "    And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                   "    And ApuracaoDeCustos.ValorDoProduto > 0 " & vbCrLf & _
                   "    And ApuracaoDeCustos.Mes_Id =  12" & vbCrLf & _
                   "    And ApuracaoDeCustos.Ano_Id = " & Format(CDate(txtDataFinal.Text), "yyyy") - 1 & vbCrLf & _
                   "  UNION ALL" & vbCrLf & _
                   "  Select Produto_Id," & vbCrLf & _
                   "         Qtde," & vbCrLf & _
                   "         CONVERT(money,valor / qtde)," & vbCrLf & _
                   "         valor," & vbCrLf & _
                   "         IND_PROP," & vbCrLf & _
                   "         cliente_id," & vbCrLf & _
                   "         EndCliente_id," & vbCrLf & _
                   "         Conta, Conta," & vbCrLf & _
                   "         Unidade" & vbCrLf & _
                   "   From #Base" & vbCrLf & _
                   "  where qtde > 0 " & vbCrLf & _
                   "    AND Valor > 0" & vbCr

        Else
            sql &= "Select 1" & vbCrLf 'H005
            sql &= "Select 1" & vbCrLf 'H010
        End If

        '***********************************************************************************
        '********************************** REG0150 ****************************************
        '***********************************************************************************
        'Registro 0150 - Tabela de Cadastro Do Participante
        sql &= "Declare @movimento nvarchar(10);" & vbCrLf & _
              "Set @movimento = '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf

        sql &= "/*Registro 0150*/" & vbCrLf & _
              " SELECT C.Cliente_Id, C.Endereco_Id, C.Regiao, C.Categoria, C.Estado, isnull(C.Pais,1058) as Pais, C.Nome, " & vbCrLf & _
              "        C.Fantasia, C.Endereco, C.Numero, ISNULL(C.Complemento, N'') AS Complemento, C.Bairro, C.Cep, " & vbCrLf & _
              "        C.Cidade, C.Inscricao, C.Telefone, C.Fax, C.Email, C.Imagem, C.Reduzido, " & vbCrLf & _
              "        ISNULL(C.CodigoDoMunicipio, 0) AS CodigoDoMunicipio, C.Situacao, E.Estado_Id, E.Descricao, Mun.Estadoibge" & vbCrLf & _
              "   FROM Clientes C" & vbCrLf & _
              "  INNER JOIN Estados E" & vbCrLf & _
              "     ON C.Estado = E.Estado_Id" & vbCrLf & _
              "  INNER JOIN Municipios Mun" & vbCrLf & _
              "     ON E.Estado_Id = Mun.Estado_id" & vbCrLf & _
              "    AND ISNULL(C.CodigoDoMunicipio, 0) = Mun.Codigo_id" & vbCrLf & _
              "    AND C.Cidade= Mun.Municipio_id " & vbCrLf & _
              "  Inner Join" & vbCrLf & _
              "      (" & vbCrLf & _
              "     Select sb.Cliente_id," & vbCrLf & _
              "             sb.EndCliente_Id" & vbCrLf & _
               "     from (" & vbCrLf & _
               "   		select nfi.Empresa_Id," & vbCrLf & _
               "   		       nfi.EndEmpresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   				  when SO.ProdutoDeTerceiro =  1                                                                    then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
               "   			   end as Classificacao," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.destino" & vbCrLf & _
               "   				   else nfi.cliente_id" & vbCrLf & _
               "   			   end Cliente_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.Enddestino" & vbCrLf & _
               "   				   else nfi.Endcliente_id" & vbCrLf & _
               "   			   end EndCliente_id," & vbCrLf & _
               "   			   nfi.Produto_Id," & vbCrLf & _
               "   			   isnull(nf.pedido,0) as Pedido," & vbCrLf & _
               "                  convert(numeric(18,4),0) as QtdeRazao," & vbCrLf & _
               "   			   SUM(Case" & vbCrLf & _
               "   					 When So.Devolucao = 'S'" & vbCrLf & _
               "   					   then nfi.QuantidadeFiscal * -1" & vbCrLf & _
               "   					   else nfi.QuantidadeFiscal" & vbCrLf & _
               "   				   end) QtdeFiscal," & vbCrLf & _
               "               convert(numeric(18,4),0) as ValorRazao," & vbCrLf & _
               "   			   SUM(Case" & vbCrLf & _
               "   					 When So.Devolucao = 'S'" & vbCrLf & _
               "   					   then nfe.valor * -1" & vbCrLf & _
               "   					   else nfe.valor" & vbCrLf & _
               "   				   end) as valorFiscal" & vbCrLf & _
               "   		  from notasfiscais NF" & vbCrLf & _
               "   		 inner join Notasfiscaisxitens NFI" & vbCrLf & _
               "   			on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
               "   		   and NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
               "   		   and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
               "   		   and NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
               "   		   and NF.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
               "   		   and NF.Nota_id         = NFI.Nota_id" & vbCrLf & _
               "   		   and NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
               "   		 Inner Join NotasFiscaisxEncargos Nfe" & vbCrLf & _
               "   			on Nfe.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
               "   		   and Nfe.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
               "   		   and Nfe.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
               "   		   and Nfe.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
               "   		   and Nfe.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
               "   		   and Nfe.Nota_id         = NFI.Nota_id" & vbCrLf & _
               "   		   and Nfe.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
               "   		   and Nfe.Produto_id      = NFI.Produto_id" & vbCrLf & _
               "   		   and Nfe.CFOP_Id         = NFI.CFOP_Id" & vbCrLf & _
               "   		   and Nfe.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
               "   		   and Nfe.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
               "   		 Inner join produtos p" & vbCrLf & _
               "   			on NFI.Produto_id = P.Produto_id" & vbCrLf & _
               "   		 Inner Join Suboperacoes SO" & vbCrLf & _
               "   			on SO.Operacao_Id        = NFI.Operacao" & vbCrLf & _
               "   		   and SO.Suboperacoes_Id    = NFI.Suboperacao" & vbCrLf & _
               "   		 where (SO.produtodeterceiro = 1 or SO.deposito = 'S')" & vbCrLf & _
               "   		   and (SO.EstoqueFisico = 'S' or SO.EstoqueFiscal = 'S' or SO.Devolucao = 'S')" & vbCrLf & _
               "   		   and SO.Operacao_id <> 80" & vbCrLf & _
               "   		   and nf.movimento  <= @movimento" & vbCrLf & _
               "           and nf.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
               "           and nf.Cliente_Id not in('" & Empresa(0) & "')" & vbCrLf & _
               "   		  group by nfi.Empresa_Id," & vbCrLf & _
               "   		           nfi.EndEmpresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   				  when SO.ProdutoDeTerceiro = 1                                                                     then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id, 8) = left(nfi.Cliente_id, 8)" & vbCrLf & _
               "   				   then nf.destino" & vbCrLf & _
               "   				   else nfi.cliente_id" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.Enddestino" & vbCrLf & _
               "   				   else nfi.Endcliente_id" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "               nf.pedido," & vbCrLf & _
               "               nfi.Produto_Id" & vbCrLf & _
               "        union all" & vbCrLf & _
               "   		Select R.Empresa_Id," & vbCrLf & _
               "   			   R.Endempresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   			   end Classificacao," & vbCrLf & _
               "   			   Case" & vbCrLf & _
               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
               "                   then R.Deposito" & vbCrLf & _
               "                   else R.Cliente_id" & vbCrLf & _
               "               end Cliente_id," & vbCrLf & _
               "   			   Case" & vbCrLf & _
               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
               "                   then R.EndDeposito" & vbCrLf & _
               "                   else R.EndCliente_id" & vbCrLf & _
               "               end EndCliente_id," & vbCrLf & _
               "   			   R.Produto," & vbCrLf & _
               "   			   isnull(R.Pedido,0) as Pedido," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then Isnull(R.CreditoQuantidade, 0.0000) - Isnull(R.DebitoQuantidade, 0.0000)" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoQuantidade, 0.0000)  - Isnull(R.CreditoQuantidade, 0.0000)" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   convert(numeric(18,4),0)," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then isnull(R.CreditoOficial, 0.00) - isnull(R.DebitoOficial, 0.00)" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoOficial, 0.00)  - isnull(R.CreditoOficial, 0.00)" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "               convert(numeric(18,4),0)" & vbCrLf & _
               "   		  from Razao R" & vbCrLf & _
               "   		 inner join ClientesxEmpresas CxE" & vbCrLf & _
               "   			on R.Empresa_id    = CxE.Empresa_Id" & vbCrLf & _
               "   		   and R.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
               "   		   and (left(R.Conta_Id,7) = CxE.ContaEstoqueEmNossoPoder or Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) = Left(cxe.ContaestoqueemPoderdeTerceiros,5))" & vbCrLf & _
               "   		 inner join Clientes C" & vbCrLf & _
               "   			on R.Cliente_id    = C.Cliente_id" & vbCrLf & _
               "   		   and R.EndCliente_id = C.Endereco_id" & vbCrLf & _
               "   		 where lote_id not in (9,10,11)" & vbCrLf & _
               "           and R.Movimento_Id <= @movimento" & vbCrLf & _
               "           and year(R.Movimento_Id) > 2010 " & vbCrLf & _
               "           and R.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
               "           and R.Cliente_Id not in('" & Empresa(0) & "')" & vbCrLf & _
               "   		   and len(isnull(produto,'')) > 0" & vbCrLf & _
               "   		) as Sb" & vbCrLf & _
               "     inner join Clientes Emp" & vbCrLf & _
               "        on Sb.Empresa_id    = Emp.Cliente_id" & vbCrLf & _
               "       and Sb.EndEmpresa_id = Emp.Endereco_id" & vbCrLf & _
               "     inner join Clientes Cli" & vbCrLf & _
               "        on Sb.cliente_id    = Cli.Cliente_id" & vbCrLf & _
               "       and Sb.EndCliente_id = Cli.Endereco_id" & vbCrLf & _
               "     inner Join Produtos Prd" & vbCrLf & _
               "        on Prd.Produto_id = sb.Produto_id" & vbCrLf & _
               "     Where Prd.TipoDoItem Not in(7,8,9,10,99)" & vbCrLf & _
               "    Group by sb.Cliente_id," & vbCrLf & _
               "          sb.EndCliente_Id" & vbCrLf & _
               "    having SUM(sb.QtdeRazao + sb.QtdeFiscal)   <> 0  Or Sum(sb.ValorRazao + sb.ValorFiscal) <> 0" & vbCrLf

        sql &= "union" & vbCrLf

        sql &= "         SELECT NF.Cliente_Id, NF.EndCliente_Id" & vbCrLf &
              "           FROM NotasFiscais NF  " & vbCrLf &
              "          INNER JOIN NotasFiscaisXItens NFXI" & vbCrLf &
              "             ON NF.Empresa_Id      = NFXI.Empresa_Id" & vbCrLf &
              "            AND NF.EndEmpresa_Id   = NFXI.EndEmpresa_Id" & vbCrLf &
              "            AND NF.Cliente_Id      = NFXI.Cliente_Id" & vbCrLf &
              "            AND NF.EndCliente_Id   = NFXI.EndCliente_Id" & vbCrLf &
              "            AND NF.EntradaSaida_Id = NFXI.EntradaSaida_Id" & vbCrLf &
              "            AND NF.Serie_Id        = NFXI.Serie_Id" & vbCrLf &
              "            AND NF.Nota_Id         = NFXI.Nota_Id" & vbCrLf &
              "          Inner Join OperacaoXEstado OE" & vbCrLf &
              "             on OE.Codigo_id = nfxi.OperacaoxEstado" & vbCrLf &
              "          INNER JOIN SubOperacoes Sop" & vbCrLf &
              "             ON NF.Operacao    = Sop.Operacao_Id " & vbCrLf &
              "            AND NF.SubOperacao = Sop.SubOperacoes_Id " & vbCrLf &
              "          WHERE NF.Situacao not in (9,2,10) " & vbCrLf &
              "            And NF.Operacao   > 0 " & vbCrLf &
              "            And NF.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf

        If Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
            'SE FOR ORIX TEM QUE ENTRAR NA LISTA
        Else
            sql &= "            And NF.Cliente_Id not in('" & Empresa(0) & "')" & vbCrLf
        End If

        sql &= "            And NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
              "            And NF.Serie_Id  not in('D','F','REC') " & vbCrLf

        'NÃO LEVAR TIPO DE DOCUMENTO 65, SOLICITAÇÃO JÉSSICA EM 19/12/2023 - FURLAN
        If Left(Empresa(0), 8) = "05366261" OrElse
            Left(Empresa(0), 8) = "38198213" OrElse
            Left(Empresa(0), 8) = "40938762" OrElse
            Left(Empresa(0), 8) = "49673784" Then

            sql &= "            AND NF.TipoDeDocumento NOT IN (15,65)" & vbCrLf
        Else
            sql &= "            AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf
        End If

        sql &= "            AND NF.Finalidade NOT IN (30)" & vbCrLf &
              "            AND Not (NF.EntradaSaida_Id = 'E'   AND OE.CodigoFiscal IN (1933,2933))    " & vbCrLf & _
              "            AND NOT (Sop.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf & _
              "            AND NOT (Sop.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf & _
              "          UNION " & vbCrLf & _
              "         SELECT NF.Cliente_Id, NF.EndCliente_Id" & vbCrLf & _
              "           FROM MemorandoDeExportacaoXNotaFiscal MeNF" & vbCrLf & _
              "          INNER JOIN NotasFiscais NF" & vbCrLf & _
              "             ON MeNF.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
              "            AND MeNF.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
              "            AND MeNF.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
              "            AND MeNF.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
              "            AND MeNF.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
              "            AND MeNF.Nota_Id         = NF.Nota_Id" & vbCrLf & _
              "            AND MeNF.Serie_Id        = NF.Serie_Id" & vbCrLf & _
              "          INNER JOIN NFERealizadas NfeR" & vbCrLf & _
              "             ON NF.Empresa_Id      = NfeR.Empresa_Id" & vbCrLf & _
              "            AND NF.EndEmpresa_Id   = NfeR.EndEmpresa_Id" & vbCrLf & _
              "            AND NF.Cliente_Id      = NfeR.Cliente_Id" & vbCrLf & _
              "            AND NF.EndCliente_Id   = NfeR.EndCliente_Id" & vbCrLf & _
              "            AND NF.EntradaSaida_Id = NfeR.EntradaSaida_Id" & vbCrLf & _
              "            AND NF.Serie_Id        = NfeR.Serie_Id" & vbCrLf & _
              "            AND NF.Nota_Id         = NfeR.Nota_Id" & vbCrLf & _
              "          INNER JOIN MemorandoDeExportacao ME" & vbCrLf & _
              "             ON MeNF.EmpresaMemorando_Id    = ME.EmpresaMemorando_Id" & vbCrLf & _
              "            AND MeNF.EndEmpresaMemorando_Id = ME.EndEmpresaMemorando_Id" & vbCrLf & _
              "            AND MeNF.Memorando_Id           = ME.Memorando_Id" & vbCrLf & _
              "           LEFT JOIN NotaFiscalXExportacao AS NFEx " & vbCrLf & _
              "             ON NF.Empresa_Id      = NFEx.Empresa_Id " & vbCrLf & _
              "            AND NF.EndEmpresa_Id   = NFEx.EndEmpresa_Id " & vbCrLf & _
              "            AND NF.Cliente_Id      = NFEx.Cliente_Id " & vbCrLf & _
              "            AND NF.EndCliente_Id   = NFEx.EndCliente_Id " & vbCrLf & _
              "            AND NF.EntradaSaida_Id = NFEx.EntradaSaida_Id " & vbCrLf & _
              "            AND NF.Serie_Id        = NFEx.Serie_Id " & vbCrLf & _
              "            AND NF.Nota_Id         = NFEx.Nota_Id " & vbCrLf & _
              "          WHERE NF.Situacao      not in(9,10)" & vbCrLf & _
              "            AND NF.Empresa_Id    =  '" & Empresa(0) & "'" & vbCrLf & _
              "            AND NF.Cliente_Id not in ('" & Empresa(0) & "')" & vbCrLf & _
              "            --AND ME.DataMemorando BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
              "            AND ISNULL(ME.DataAverbacao, NFEx.DataAverbacao) BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
              "            --AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf

        If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
            sql &= "          UNION" & vbCrLf & _
                    "       SELECT Deposito_Id, EndDeposito_id" & vbCrLf & _
                    "       FROM ApuracaoDeCustos  " & vbCrLf & _
                    "		        INNER JOIN Produtos " & vbCrLf & _
                    "				        ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                    "		        INNER JOIN UnidadeDeMedida UN " & vbCrLf & _
                    "				        ON Produtos.Unidade = UN.Unidade_Id" & vbCrLf & _
                    "       Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                    "           And ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                    "           And ApuracaoDeCustos.Mes_Id =  " & PeriodoInicial.Month & vbCrLf & _
                    "           And ApuracaoDeCustos.Ano_Id = " & PeriodoInicial.Year & vbCrLf & _
                    "           And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                    "           And ApuracaoDeCustos.ValorDoProduto > 0" & vbCrLf & _
                    "           And  ISNULL(Produtos.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf & _
                    "           And ApuracaoDeCustos.Deposito_Id <> ApuracaoDeCustos.Empresa_Id" & vbCrLf & _
                    "       GROUP BY Deposito_Id, EndDeposito_id" & vbCrLf
        End If

        If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
            sql &= "          UNION" & vbCrLf & _
                   "         SELECT APC.Deposito_Id, APC.EndDeposito_Id" & vbCrLf & _
                   "           FROM ApuracaoDeCustos APC" & vbCrLf & _
                   "          WHERE APC.CodigoDeCusto_Id IN (109,920)" & vbCrLf & _
                   "            And APC.Empresa_Id       =  '" & Empresa(0) & "'" & vbCrLf & _
                   "            And APC.Deposito_Id      <> '" & Empresa(0) & "'" & vbCrLf & _
                   "            And APC.Ano_id           =" & PeriodoInicial.Year - 1 & vbCrLf & _
                   "            And APC.Mes_id           = 12" & vbCrLf & _
                   "            And APC.Quantidade       <> 0 " & vbCrLf

            'TESTE INFORMADO PARA ISOLAR PARA O LEANDRO - FURLAN - 09/03/2021
            If Empresa(0).Substring(0, 8) = "05272759" Then
                sql &= "        And APC.Deposito_Id      <> '95384459000142'" & vbCrLf
            End If

            sql &= "           UNION" & vbCrLf & _
                   "           Select Cliente_id, EndCliente_Id from #Base" & vbCrLf
        End If
        sql &= "      ) as Consulta" & vbCrLf & _
               "    on C.Cliente_ID  = Consulta.Cliente_ID " & vbCrLf & _
               "   and C.Endereco_ID = Consulta.EndCliente_ID " & vbCrLf & _
               " ORDER BY Nome, Estado, Cidade " & vbCrLf

        '***********************************************************************************
        '********************************** REG0190 ****************************************
        '***********************************************************************************
        'Registro 0190 - Identificação das Unidades de Medidas    
        sql &= "/*Registro 0190*/" & vbCrLf & _
               "SELECT Unidade, UN.Descricao " & vbCrLf & _
               " into #Unidades " & vbCrLf & _
               "  FROM (" & vbCrLf & _
               "        SELECT Produtos.Unidade " & vbCrLf & _
               "          FROM NotasFiscais" & vbCrLf & _
               "         INNER JOIN NotasFiscaisXItens" & vbCrLf & _
               "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
               "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
               "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
               "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
               "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
               "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
               "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
               "         INNER JOIN Produtos " & vbCrLf & _
               "            ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
               "          LEFT JOIN NFERealizadas" & vbCrLf & _
               "            ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
               "           AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
               "           AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
               "           AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
               "           AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
               "           AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
               "           AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
               "         INNER JOIN SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf & _
               "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
               "         Inner Join OperacaoXEstado OE" & vbCrLf & _
               "            ON OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf & _
               "         WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
               "           AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & vbCrLf & _
               "           AND NotasFiscais.Cliente_Id not in('" & Empresa(0) & "')" & vbCrLf & _
               "           AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
               "           AND NotasFiscais.TipoDeDocumento NOT IN (15)" & vbCrLf &
               "           AND NotasFiscais.Finalidade NOT IN (30)" & vbCrLf &
               "           AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT IN (1932,2932))" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
               "           AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
               "           AND Not (NotasFiscais.EntradaSaida_Id = 'E'   " & vbCrLf & _
               "           AND OE.CodigoFiscal IN (1933,2933)) " & vbCrLf & _
               "           AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf & _
               "           AND Not (NotasFiscais.EntradaSaida_Id = 'S'   " & vbCrLf & _
               "           AND OE.CodigoFiscal IN (5933,6933)) " & vbCrLf & _
               "           AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf & _
               "           AND ((NFERealizadas.Chavenfe is null) or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null) )" & vbCrLf & _
               "            OR (NotasFiscais.Empresa_Id =  '" & Empresa(0) & "' " & vbCrLf & _
               "                AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf & _
               "                AND NotasFiscais.TipoDeDocumento = 2" & vbCrLf & _
               "                AND NotasFiscais.NossaEmissao    ='S') " & vbCrLf & _
               "         GROUP BY Produtos.Unidade " & vbCrLf & _
               "        having sum(NotasFiscaisxitens.quantidadefiscal) > 0" & vbCrLf

        If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
            sql &= " Union" & vbCrLf & _
                   " SELECT Produtos.Unidade " & vbCrLf & _
                   "   FROM Produtos " & vbCrLf & _
                   "  INNER JOIN ApuracaoDeCustos " & vbCrLf & _
                   "     ON Produtos.Produto_Id = ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                   "  Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                   "    And ApuracaoDeCustos.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf & _
                   "    And ApuracaoDeCustos.Ano_id     = " & PeriodoInicial.Year - 1 & vbCrLf & _
                   "    And ApuracaoDeCustos.Mes_id     = 12 " & vbCrLf & _
                   "    And ApuracaoDeCustos.Quantidade <> 0 " & vbCrLf & _
                   "    AND ApuracaoDeCustos.ValorDoProduto > 0 " & vbCrLf & _
                   "  Union " & vbCrLf & _
                   "  Select Unidade from #Base" & vbCrLf
        End If

        sql &= " ) AS Consulta " & vbCrLf & _
               " INNER JOIN UnidadeDeMedida UN " & vbCrLf & _
               "    ON Unidade = UN.Unidade_Id " & vbCrLf & _
               " GROUP BY Unidade, UN.Descricao " & vbCrLf

        If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
            sql &= "INSERT INTO #Unidades" & vbCrLf & _
                    "SELECT Produtos.Unidade, UN.Descricao" & vbCrLf & _
                    "FROM ApuracaoDeCustos  " & vbCrLf & _
                    "		INNER JOIN Produtos " & vbCrLf & _
                    "				ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                    "		INNER JOIN UnidadeDeMedida UN " & vbCrLf & _
                    "				ON Produtos.Unidade = UN.Unidade_Id" & vbCrLf & _
                    "Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                    "    And ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                    "    And ApuracaoDeCustos.Mes_Id =  " & PeriodoInicial.Month & vbCrLf & _
                    "    And ApuracaoDeCustos.Ano_Id = " & PeriodoInicial.Year & vbCrLf & _
                    "    And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                    "    And ApuracaoDeCustos.ValorDoProduto > 0" & vbCrLf & _
                    "    And  ISNULL(Produtos.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf & _
                    "GROUP BY Produtos.Unidade, UN.Descricao" & vbCrLf
        End If

        sql &= "select Unidade, Descricao from #Unidades " & vbCrLf & _
               "GROUP BY Unidade, Descricao" & vbCrLf


        '***********************************************************************************
        '********************************** REG0200 ****************************************
        '***********************************************************************************
        'Registro 0200 - Tabela de Identificação do Item (Produtos e Serviços)
        sql &= "/*Registro 0200*/" & vbCrLf &
               "SELECT P.Produto_Id," & vbCrLf &
               "       P.Nome," & vbCrLf &
               "       P.Unidade," & vbCrLf &
               "       ISNULL(P.TipoDoItem, 0) AS TipoDoItem," & vbCrLf &
               "       ISNULL(P.NCM, '') AS NCM," & vbCrLf &
               "       ISNULL(P.CodigoEX, '') AS CodigoEX," & vbCrLf &
               "       ISNULL(P.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf &
               "       ISNULL(P.CodigoDoServico, 0) AS CodigoDoServico," & vbCrLf &
               "       ISNULL(P.ICMS, 0) AS ICMS" & vbCrLf &
               " into #BASEPRODUTOS" & vbCrLf &
               "  FROM PRODUTOS P " & vbCrLf &
               " Where produto_Id" & vbCrLf &
               "    in (" & vbCrLf &
               "          --Select isnull(pa.Produto_Id, p.Produto_Id)" & vbCrLf &
               "           -- from Produtos p" & vbCrLf &
               "           -- left join ProdutosAgrupados pa" & vbCrLf &
               "             -- on p.Produto_Id = pa.ProdutoAgrupado_Id" & vbCrLf &
               "           --Where p.Produto_id in (" & vbCrLf &
               "                                    SELECT ISNULL(PA.Produto_id, NotasFiscaisXItens.Produto_Id) AS Produto_Id" & vbCrLf &
               "                                      FROM NotasFiscais" & vbCrLf &
               "                                     INNER JOIN NotasFiscaisXItens" & vbCrLf &
               "                                        ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
               "                                       AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
               "                                       AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
               "                                       AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
               "                                       AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
               "                                       AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
               "                                       AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
               "                                     Inner Join OperacaoXEstado OE" & vbCrLf &
               "                                        ON OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
               "                                      LEFT JOIN NotaFiscalXExportacao AS NFEx " & vbCrLf &
               "                                        ON NotasFiscais.Empresa_Id       = NFEx.Empresa_Id " & vbCrLf &
               "                                       AND NotasFiscais.EndEmpresa_Id    = NFEx.EndEmpresa_Id " & vbCrLf &
               "                                       AND NotasFiscais.Cliente_Id       = NFEx.Cliente_Id " & vbCrLf &
               "	                                   AND NotasFiscais.EndCliente_Id    = NFEx.EndCliente_Id " & vbCrLf &
               "                                       AND NotasFiscais.EntradaSaida_Id  = NFEx.EntradaSaida_Id " & vbCrLf &
               "                                       AND NotasFiscais.Serie_Id         = NFEx.Serie_Id " & vbCrLf &
               "                                       AND NotasFiscais.Nota_Id          = NFEx.Nota_Id " & vbCrLf &
               "                                      LEFT JOIN MemorandoDeExportacao AS ME " & vbCrLf &
               "                                        ON NotasFiscais.Empresa_Id = ME.Empresa " & vbCrLf &
               "                                       AND NotasFiscais.EndEmpresa_Id = ME.EndEmpresa " & vbCrLf &
               "                                       AND NotasFiscais.Cliente_Id = ME.Cliente " & vbCrLf &
               "                                       AND NotasFiscais.EndCliente_Id = ME.EndCliente" & vbCrLf &
               "                                       AND NotasFiscais.EntradaSaida_Id = ME.EntradaSaida " & vbCrLf &
               "                                       AND NotasFiscais.Serie_Id = ME.Serie " & vbCrLf &
               "                                       AND NotasFiscais.Nota_Id = ME.Nota" & vbCrLf &
               "                                      --LEFT JOIN ProdutosAgrupados PA" & vbCrLf &
               "                                        --ON PA.ProdutoAgrupado_Id = NotasFiscaisXItens.Produto_Id" & vbCrLf &
               "                                     --INNER JOIN Produtos " & vbCrLf &
               "                                        --ON Produtos.Produto_Id = isnull(PA.Produto_id,NotasFiscaisXItens.Produto_Id)" & vbCrLf &
               "                                      LEFT JOIN ProdutosAgrupados pa" & vbCrLf &
               "                                        ON NotasFiscaisXItens.Produto_Id = pa.ProdutoAgrupado_Id" & vbCrLf &
               "                                      LEFT JOIN NFERealizadas" & vbCrLf &
               "                                        ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf &
               "                                       AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf &
               "                                       AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf &
               "                                       AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf &
               "                                       AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf &
               "                                       AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf &
               "                                       AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf &
               "                                     INNER JOIN SubOperacoes" & vbCrLf &
               "                                        ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf &
               "                                       AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
               "                                     WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf &
               "                                       AND NotasFiscais.Situacao not in(2,9,10) " & vbCrLf &
               "                                       AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & vbCrLf &
               "                                       AND NotasFiscais.Cliente_Id not in ('" & Empresa(0) & "')" & vbCrLf &
               "                                       AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf

        If Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
            sql &= "                                       AND NotasFiscais.TipoDeDocumento NOT IN(2,8,10,11,14,15,16,57,65)" & vbCrLf

        Else
            sql &= "                                       AND NotasFiscais.TipoDeDocumento NOT IN (15)" & vbCrLf
        End If

        sql &= "                                       AND NotasFiscais.Finalidade NOT IN (30)" & vbCrLf &
               "                                       AND NotasFiscaisXItens.Operacao > 0" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT IN(1932,2932))" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf &
               "                                       AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf &
               "                                       AND Not (NotasFiscais.EntradaSaida_Id = 'E'  AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf &
               "                                       AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
               "                                       --AND Not (NotasFiscais.EntradaSaida_Id = 'S'  AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf &
               "                                       AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf &
               "                                       AND ((NFERealizadas.Chavenfe is null) " & vbCrLf &
               "                                              or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null)" & vbCrLf &
               "                                              or (ISNULL(ME.DataAverbacao, NFEx.DataAverbacao) BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'))" & vbCrLf &
               "                                        OR (NotasFiscais.Empresa_Id =  '" & Empresa(0) & "' AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' AND NotasFiscais.TipoDeDocumento=2 AND NotasFiscais.NossaEmissao='S') " & vbCrLf &
               "                                     group by ISNULL(PA.Produto_id, NotasFiscaisXItens.Produto_Id), OE.CodigoFiscal" & vbCrLf &
               "                                    having sum(NotasFiscaisxitens.quantidadefiscal) > 0" & vbCrLf

        If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
            sql &= "                                         Union" & vbCrLf & _
                   "                                        SELECT ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                   "                                          FROM ApuracaoDeCustos  " & vbCrLf & _
                   "                                                Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                   "                                                  And ApuracaoDeCustos.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf & _
                   "                                                  And ApuracaoDeCustos.Ano_id     = " & PeriodoInicial.Year - 1 & vbCrLf & _
                   "                                                  And ApuracaoDeCustos.Mes_id     = 12 " & vbCrLf & _
                   "                                                  And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                   "                                                  and ApuracaoDeCustos.ValorDoProduto > 0 " & vbCrLf & _
                   "                                         Union" & vbCrLf & _
                   "                                        Select Produto_Id from #Base"
        End If
        sql &= " )--)" & vbCrLf

        If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
            sql &= "INSERT INTO #BASEPRODUTOS" & vbCrLf & _
                    "SELECT P.Produto_Id," & vbCrLf & _
                    "       P.Nome," & vbCrLf & _
                    "       P.Unidade," & vbCrLf & _
                    "       ISNULL(P.TipoDoItem, 0) AS TipoDoItem," & vbCrLf & _
                    "       ISNULL(P.NCM, '') AS NCM," & vbCrLf & _
                    "       ISNULL(P.CodigoEX, '') AS CodigoEX," & vbCrLf & _
                    "       ISNULL(P.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf & _
                    "       ISNULL(P.CodigoDoServico, 0) AS CodigoDoServico," & vbCrLf & _
                    "       ISNULL(P.ICMS, 0) AS ICMS" & vbCrLf & _
                    "FROM ApuracaoDeCustos " & vbCrLf & _
                    "		INNER JOIN PRODUTOS P" & vbCrLf & _
                    "				ON P.Produto_id = ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                    "Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                    "    And ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                    "    And ApuracaoDeCustos.Mes_Id =  " & PeriodoInicial.Month & vbCrLf & _
                    "    And ApuracaoDeCustos.Ano_Id = " & PeriodoInicial.Year & vbCrLf & _
                    "    And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                    "    And ApuracaoDeCustos.ValorDoProduto > 0" & vbCrLf & _
                    "    And  ISNULL(P.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf
        End If

        sql &= "INSERT INTO #BASEPRODUTOS" & vbCrLf & _
               "     Select sb.Produto_id," & vbCrLf & _
               "            Prd.Nome," & vbCrLf & _
               "            Prd.Unidade," & vbCrLf & _
               "            ISNULL(Prd.TipoDoItem, 0) AS TipoDoItem," & vbCrLf & _
               "            ISNULL(Prd.NCM, '') AS NCM," & vbCrLf & _
               "            ISNULL(Prd.CodigoEX, '') AS CodigoEX," & vbCrLf & _
               "            ISNULL(Prd.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf & _
               "            ISNULL(Prd.CodigoDoServico, 0) AS CodigoDoServico," & vbCrLf & _
               "            ISNULL(Prd.ICMS, 0) AS ICMS" & vbCrLf & _
               "     from (" & vbCrLf & _
               "   		select nfi.Empresa_Id," & vbCrLf & _
               "   		       nfi.EndEmpresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   				  when SO.ProdutoDeTerceiro =  1                                                                    then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
               "   			   end as Classificacao," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.destino" & vbCrLf & _
               "   				   else nfi.cliente_id" & vbCrLf & _
               "   			   end Cliente_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.Enddestino" & vbCrLf & _
               "   				   else nfi.Endcliente_id" & vbCrLf & _
               "   			   end EndCliente_id," & vbCrLf & _
               "   			   nfi.Produto_Id," & vbCrLf & _
               "   			   isnull(nf.pedido,0) as Pedido," & vbCrLf & _
               "                  convert(numeric(18,4),0) as QtdeRazao," & vbCrLf & _
               "   			   SUM(Case" & vbCrLf & _
               "   					 When So.Devolucao = 'S'" & vbCrLf & _
               "   					   then nfi.QuantidadeFiscal * -1" & vbCrLf & _
               "   					   else nfi.QuantidadeFiscal" & vbCrLf & _
               "   				   end) QtdeFiscal," & vbCrLf & _
               "               convert(numeric(18,4),0) as ValorRazao," & vbCrLf & _
               "   			   SUM(Case" & vbCrLf & _
               "   					 When So.Devolucao = 'S'" & vbCrLf & _
               "   					   then nfe.valor * -1" & vbCrLf & _
               "   					   else nfe.valor" & vbCrLf & _
               "   				   end) as valorFiscal" & vbCrLf & _
               "   		  from notasfiscais NF" & vbCrLf & _
               "   		 inner join Notasfiscaisxitens NFI" & vbCrLf & _
               "   			on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
               "   		   and NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
               "   		   and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
               "   		   and NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
               "   		   and NF.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
               "   		   and NF.Nota_id         = NFI.Nota_id" & vbCrLf & _
               "   		   and NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
               "   		 Inner Join NotasFiscaisxEncargos Nfe" & vbCrLf & _
               "   			on Nfe.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
               "   		   and Nfe.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
               "   		   and Nfe.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
               "   		   and Nfe.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
               "   		   and Nfe.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
               "   		   and Nfe.Nota_id         = NFI.Nota_id" & vbCrLf & _
               "   		   and Nfe.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
               "   		   and Nfe.Produto_id      = NFI.Produto_id" & vbCrLf & _
               "   		   and Nfe.CFOP_Id         = NFI.CFOP_Id" & vbCrLf & _
               "   		   and Nfe.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
               "   		   and Nfe.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
               "   		 Inner join produtos p" & vbCrLf & _
               "   			on NFI.Produto_id = P.Produto_id" & vbCrLf & _
               "   		 Inner Join Suboperacoes SO" & vbCrLf & _
               "   			on SO.Operacao_Id        = NFI.Operacao" & vbCrLf & _
               "   		   and SO.Suboperacoes_Id    = NFI.Suboperacao" & vbCrLf & _
               "   		 where (SO.produtodeterceiro = 1 or SO.deposito = 'S')" & vbCrLf & _
               "   		   and (SO.EstoqueFisico = 'S' or SO.EstoqueFiscal = 'S' or SO.Devolucao = 'S')" & vbCrLf & _
               "   		   and SO.Operacao_id <> 80" & vbCrLf & _
               "   		   and nf.movimento  <= @movimento" & vbCrLf & _
               "           and nf.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
               "           and nf.Cliente_Id not in('" & Empresa(0) & "')" & vbCrLf & _
               "   		  group by nfi.Empresa_Id," & vbCrLf & _
               "   		           nfi.EndEmpresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   				  when SO.ProdutoDeTerceiro = 1                                                                     then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id, 8) = left(nfi.Cliente_id, 8)" & vbCrLf & _
               "   				   then nf.destino" & vbCrLf & _
               "   				   else nfi.cliente_id" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.Enddestino" & vbCrLf & _
               "   				   else nfi.Endcliente_id" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "               nf.pedido," & vbCrLf & _
               "               nfi.Produto_Id" & vbCrLf & _
               "        union all" & vbCrLf & _
               "   		Select R.Empresa_Id," & vbCrLf & _
               "   			   R.Endempresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   			   end Classificacao," & vbCrLf & _
               "   			   Case" & vbCrLf & _
               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
               "                   then R.Deposito" & vbCrLf & _
               "                   else R.Cliente_id" & vbCrLf & _
               "               end Cliente_id," & vbCrLf & _
               "   			   Case" & vbCrLf & _
               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
               "                   then R.EndDeposito" & vbCrLf & _
               "                   else R.EndCliente_id" & vbCrLf & _
               "               end EndCliente_id," & vbCrLf & _
               "   			   R.Produto," & vbCrLf & _
               "   			   isnull(R.Pedido,0) as Pedido," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then Isnull(R.CreditoQuantidade, 0.0000) - Isnull(R.DebitoQuantidade, 0.0000)" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoQuantidade, 0.0000)  - Isnull(R.CreditoQuantidade, 0.0000)" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   convert(numeric(18,4),0)," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then isnull(R.CreditoOficial, 0.00) - isnull(R.DebitoOficial, 0.00)" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoOficial, 0.00)  - isnull(R.CreditoOficial, 0.00)" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "               convert(numeric(18,4),0)" & vbCrLf & _
               "   		  from Razao R" & vbCrLf & _
               "   		 inner join ClientesxEmpresas CxE" & vbCrLf & _
               "   			on R.Empresa_id    = CxE.Empresa_Id" & vbCrLf & _
               "   		   and R.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
               "   		   and (left(R.Conta_Id,7) = CxE.ContaEstoqueEmNossoPoder or Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) = Left(cxe.ContaestoqueemPoderdeTerceiros,5))" & vbCrLf & _
               "   		 inner join Clientes C" & vbCrLf & _
               "   			on R.Cliente_id    = C.Cliente_id" & vbCrLf & _
               "   		   and R.EndCliente_id = C.Endereco_id" & vbCrLf & _
               "   		 where lote_id not in (9,10,11)" & vbCrLf & _
               "           and R.Movimento_Id <= @movimento" & vbCrLf & _
               "           and year(R.Movimento_Id) > 2010 " & vbCrLf & _
               "           and R.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
               "   		   and len(isnull(produto,'')) > 0" & vbCrLf & _
               "   		) as Sb" & vbCrLf & _
               "     inner join Clientes Emp" & vbCrLf & _
               "        on Sb.Empresa_id    = Emp.Cliente_id" & vbCrLf & _
               "       and Sb.EndEmpresa_id = Emp.Endereco_id" & vbCrLf & _
               "     inner join Clientes Cli" & vbCrLf & _
               "        on Sb.cliente_id    = Cli.Cliente_id" & vbCrLf & _
               "       and Sb.EndCliente_id = Cli.Endereco_id" & vbCrLf & _
               "     inner Join Produtos Prd" & vbCrLf & _
               "        on Prd.Produto_id = sb.Produto_id" & vbCrLf & _
               "     Where Prd.TipoDoItem Not in(7,8,9,10,99)" & vbCrLf & _
               "    Group by sb.Produto_id," & vbCrLf & _
               "             Prd.Nome," & vbCrLf & _
               "             Prd.Unidade," & vbCrLf & _
               "             Prd.TipoDoItem," & vbCrLf & _
               "             Prd.NCM," & vbCrLf & _
               "             Prd.CodigoEX," & vbCrLf & _
               "             Prd.CodigoDoGenero," & vbCrLf & _
               "             Prd.CodigoDoServico," & vbCrLf & _
               "             Prd.ICMS" & vbCrLf & _
               "    having SUM(sb.QtdeRazao + sb.QtdeFiscal)   <> 0  Or Sum(sb.ValorRazao + sb.ValorFiscal) <> 0" & vbCrLf

        sql &= "SELECT Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf & _
                "FROM #BASEPRODUTOS" & vbCrLf & _
                "GROUP BY Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf


        ds = Banco.ConsultaDataSet(sql, "Consulta")
        ds.Tables(0).TableName = "BLOCOC"
        ds.Tables(1).TableName = "BLOCOD"
        ds.Tables(2).TableName = "H005"
        ds.Tables(3).TableName = "H010"
        ds.Tables(4).TableName = "REG0150"
        ds.Tables(5).TableName = "REG0190"
        ds.Tables(6).TableName = "REG0200"



    End Sub


    Private Sub Processar()
        If String.IsNullOrWhiteSpace(txtArquivoDeSaida.Text) Then
            MsgBox(Me.Page, "É necessário selecionar o caminho do arquivo de saída!")
            Exit Sub
        End If

        Dim NomeArquivo As String
        Dim NomeArquivo2 As String = "SpedFiscal/" & txtArquivoDeSaida.Text.ToUpper
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Using strm As New StreamWriter(Server.MapPath("~/SpedFiscal/" & txtArquivoDeSaida.Text.Trim().ToUpper()), True)
            Try
                Empresa = DdlEmpresa.SelectedValue.Split("-")
                PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
                PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
                imdDownload.Visible = False

                Dim linha As String
                Dim IntSequencia As Integer = 0
                Dim BlocoC As Integer = 1
                Dim BlocoD As Integer = 1

                Dim Nota As String = ""
                Dim Serie As String = ""
                Dim EntradaSaida As String = ""
                Dim Cliente As String = ""
                Dim EndCliente As String = ""
                Dim Emissao As String = ""
                Dim Modelo As String = ""
                Dim Cancelada As String = ""
                Dim NossaEmissao As Boolean = False

                Dim Registro0000 As Integer = 0     'Abertura do Arquivo Digital e Identificação e Identificação da Entidade
                Dim Registro0001 As Integer = 0     'Abertura do Bloco 0
                Dim Registro0002 As Integer = 0     'CLASSIFICAÇÃO DO ESTABELECIMENTO INDUSTRIAL OU EQUIPARADO A INDUSTRIAL
                Dim Registro0005 As Integer = 0     'Dados Complementares da Entidade
                Dim Registro0015 As Integer = 0     'DADOS DO CONTRIBUINTE SUBSTITUTO OU RESPONSÁVEL PELO ICMS DESTINO

                Dim Registro0100 As Integer = 0     'Dados do Contabilista
                Dim Registro0150 As Integer = 0     'Tabela de Cadastro do Participante
                Dim Registro0190 As Integer = 0     'Tabela de Unidades de Medidas
                Dim Registro0200 As Integer = 0     'Tabela de Identificação do Item (Produto e Serviços)
                Dim Registro0220 As Integer = 0     'Tabela de Fatores de Conversão de Unidades
                Dim Registro0400 As Integer = 0     'Tabela de Natureza da Operação/Prestação
                Dim Registro0450 As Integer = 0     'Tabela de Informações Complementares do Documento Fiscal
                Dim Registro0460 As Integer = 0     'Tabela de Observações do Lancamento Fiscal
                Dim Registro0990 As Integer = 0     'Encerramento do Bloco 0

                Dim RegistroB001 As Integer = 0     'Abertura do Broco B
                Dim RegistroB990 As Integer = 0     'Encerramento do Bloco B

                Dim RegistroC001 As Integer = 0     'Abertura do Broco C
                Dim RegistroC100 As Integer = 0     'Nota Fiscal
                Dim RegistroC101 As Integer = 0     'Nota Fiscal X Encargos - Informação complementar dos documentos fiscais quando das operações interestaduais destinadas a consumidor final não contribuinte EC 87/15 (código 55)
                Dim RegistroC170 As Integer = 0     'Nota Fiscal X Itens
                Dim RegistroC172 As Integer = 0     'OPERAÇÕES COM ISSQN (CÓDIGO 01)
                Dim RegistroC190 As Integer = 0     'Registro Analitico do Documento
                'Dim RegistroC195 As Integer = 0     'OBSERVAÇOES DO LANÇAMENTO FISCAL (CÓDIGO 01, 1BE 55)
                'Dim RegistroC197 As Integer = 0     'OUTRAS OBRIGAÇÕES TRIBUTÁRIAS, AJUSTES E INFORMAÇÕES DE VALORES PROVENIENTES DE DOCUMENTO FISCAL.

                Dim RegistroC500 As Integer = 0     'Nota Fiscal
                Dim RegistroC510 As Integer = 0     'Nota Fiscal X Itens
                Dim RegistroC590 As Integer = 0     'Registro Analitico do Documento

                Dim RegistroC990 As Integer = 0     'Encerramento do Bloco C

                Dim RegistroD001 As Integer = 0     'Abertura do Bloco D
                Dim RegistroD100 As Integer = 0     'Nota Fiscal de Serviço de Transporte
                Dim RegistroD190 As Integer = 0     'Registro Analitico dos Documentos
                Dim RegistroD500 As Integer = 0     'Nota Fiscal de Serviço de Comunicação
                Dim RegistroD590 As Integer = 0     'Registro Analitico dos Documentos
                Dim RegistroD990 As Integer = 0     'Encerramento do Bloco D

                Dim RegistroE001 As Integer = 0     'Abertura do Bloco D
                Dim RegistroE100 As Integer = 0     'Periodo da Apuração
                Dim RegistroE110 As Integer = 0     'Registro Analitico dos Documentos
                Dim RegistroE111 As Integer = 0     'AJUSTE/BENEFÍCIO/INCENTIVO DA APURAÇÃO DO ICMS
                Dim RegistroE112 As Integer = 0     'INFORMAÇÕES ADICIONAIS DOS AJUSTES DA APURAÇÃO DO ICMS
                Dim RegistroE113 As Integer = 0     'INFORMAÇÕES ADICIONAIS DOS AJUSTES DA APURAÇÃO DO ICMS - IDENTIFICAÇÃO DOS DOCUMENTOS FISCAIS
                Dim RegistroE115 As Integer = 0     'Dar - Diferencial de Alíquota - (Observação ICMS)
                Dim RegistroE116 As Integer = 0     'Obrigações do Icms A Recvolher - Operações Próprias

                Dim RegistroE200 As Integer = 0     'PERÍODO DA APURAÇÃO DO ICMS - SUBSTITUIÇÃO TRIBUTÁRIA
                Dim RegistroE210 As Integer = 0     'APURAÇÃO DO ICMS – SUBSTITUIÇÃO TRIBUTÁRIA
                Dim RegistroE220 As Integer = 0     'APURAÇÃO DO ICMS – SUBSTITUIÇÃO TRIBUTÁRIA
                Dim RegistroE240 As Integer = 0     'INFORMAÇÕES ADICIONAIS DOS AJUSTES DA APURAÇÃO DO ICMS SUBSTITUIÇÃO TRIBUTÁRIA – IDENTIFICAÇÃO DOS DOCUMENTOS FISCAIS
                Dim RegistroE250 As Integer = 0     'OBRIGAÇÕES DO ICMS RECOLHIDO OU A RECOLHER – SUBSTITUIÇÃO TRIBUTÁRIA

                Dim RegistroE300 As Integer = 0     'PERÍODO DE APURAÇÃO DO FUNDO DE COMBATE À POBREZA E DO ICMS DIFERENCIAL DE ALÍQUOTA – UF ORIGEM/DESTINO EC 87/15
                Dim RegistroE310 As Integer = 0     'APURAÇÃO DO ICMS DIFERENCIAL DE ALÍQUOTA – UF ORIGEM/DESTINO EC 87/15. (VÁLIDO ATÉ 31/12/2016)
                Dim RegistroE316 As Integer = 0     'OBRIGAÇÕES RECOLHIDAS OU A RECOLHER – FUNDO DE COMBATE À POBREZA E ICMS DIFERENCIAL DE ALÍQUOTA UF ORIGEM/DESTINO EC 87/15

                Dim RegistroE500 As Integer = 0     'Nota Fiscal de Serviço de Transporte
                Dim RegistroE510 As Integer = 0     'Apuração do IPI - Operações Próprias
                Dim RegistroE520 As Integer = 0     'APURAÇÃO DO IPI. Registro Analitico dos Documentos
                Dim RegistroE530 As Integer = 0     'AJUSTES DA APURAÇÃO DO IPI.
                Dim RegistroE990 As Integer = 0     'Registro Analitico dos Documentos

                Dim RegistroH001 As Integer = 0     'Abertura do Bloco H
                Dim RegistroH005 As Integer = 0     'Total do Inventário
                Dim RegistroH010 As Integer = 0     'Inventário 
                Dim RegistroH990 As Integer = 0     'Encerramento do Bloco H

                Dim RegistroG001 As Integer = 0     'Abertura do Bloco G
                Dim RegistroG110 As Integer = 0     'Total do Inventário
                Dim RegistroG125 As Integer = 0     'Inventário 
                Dim RegistroG126 As Integer = 0     'Encerramento do Bloco G
                Dim RegistroG130 As Integer = 0     'Encerramento do Bloco G
                Dim RegistroG140 As Integer = 0     'Encerramento do Bloco G
                Dim RegistroG990 As Integer = 0     'Encerramento do Bloco G

                Dim RegistroK001 As Integer = 0     'Abertura do Bloco K
                Dim RegistroK010 As Integer = 0     'Informação sobre o tipo de Leiaute(0-Simplificado, 1-Completo, 2-Restrito aos saldos de Estoque) Obrigatório à partir de 2023. Furlan - 08/02/2023.
                Dim RegistroK100 As Integer = 0     'Período de Apuração do ICMS/IPI
                Dim RegistroK200 As Integer = 0     'Estoque Escriturado K
                Dim RegistroK990 As Integer = 0     'Encerramento do Bloco K

                Dim Registro1001 As Integer = 0     'Abertura do Bloco 1
                Dim Registro1010 As Integer = 0     'Obrigatoriedade de Registros do Bloco 1
                Dim Registro1100 As Integer = 0     'Registro de Informação sobre exportação
                Dim Registro1105 As Integer = 0     'Documentos Fiscais de Exportação
                Dim Registro1110 As Integer = 0     'Operações De Exportação Indireta - Mercadorias De Terceiros
                Dim Registro1400 As Integer = 0     'Informações sobre Valores Agregados
                Dim Registro1990 As Integer = 0     'Encerramento do Bloco 1

                Dim Registro9001 As Integer = 0     'Abertura do Bloco 9
                Dim Registro9900 As Integer = 0     'Registros do Arquivo
                Dim Registro9990 As Integer = 0     'Encerramento do Bloco 9
                Dim Registro9999 As Integer = 0     'Encerramento do Arquivo Digital

                Dim RegistroGeral As Integer = 0    'Registro Geral

                Dim ds As DataSet
                Dim ds1 As DataSet
                Dim ds2 As DataSet
                Dim dsNew As DataSet
                dsNew = Nothing
                PreencherTabelas(dsNew)


                ' Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
                ' Teste de Movimento do Periodo
                ' Bloco C
                If dsNew.Tables("BLOCOC").Rows(0)(0) > 0 Then
                    BlocoC = 0
                End If
                ' Bloco D
                If dsNew.Tables("BLOCOD").Rows(0)(0) > 0 Then
                    BlocoD = 0
                End If

                'Registro 0000 - Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária

                Sql = "SELECT distinct  Clientes.Cliente_Id, Clientes.Nome, Clientes.CodigoDoMunicipio,Clientes.Estado, Clientes.Inscricao, Municipios.EstadoIbge, ClientesXEmpresas.InscricaoMunicipal" & vbCrLf & _
                      "  FROM Clientes " & vbCrLf & _
                      " INNER JOIN ClientesXEmpresas" & vbCrLf & _
                      "    ON Clientes.Cliente_Id  = ClientesXEmpresas.Empresa_Id " & vbCrLf & _
                      "   AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id " & vbCrLf & _
                      " INNER JOIN Estados" & vbCrLf & _
                      "    ON Clientes.Estado = Estados.Estado_Id " & vbCrLf & _
                      " INNER JOIN Municipios" & vbCrLf & _
                      "    ON Clientes.Estado            = Municipios.Estado_id" & vbCrLf & _
                      "   AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id" & vbCrLf & _
                      "   AND Clientes.Cidade = Municipios.Municipio_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id  ='" & Empresa(0) & "'" & vbCrLf & _
                      "   And Clientes.Endereco_Id = " & Empresa(1) & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                'Leiaute.
                Dim Leiaute As String = GetLeiaute(PeriodoFinal.Year, PeriodoFinal.Month)

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|0000"                                                                                                'Tipo
                            'linha &= IIf(PeriodoFinal.Year < 2011, "|003", IIf(PeriodoFinal.Year < 2012, "|004", IIf(PeriodoFinal.Month < 7 And PeriodoFinal.Year = 2012, "|005", IIf(PeriodoFinal.Month > 7 And PeriodoFinal.Year = 2012, "|006", IIf(PeriodoFinal.Year < 2014, "|007", "|008")))))
                            linha &= "|" & Leiaute                                                                                          'COD_VER
                            linha &= "|" & DdlFinalidade.SelectedValue
                            linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                                   'Data Inicial
                            linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                                     'Data Final
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                            linha &= "|" & Microsoft.VisualBasic.Left(Trim(.Item("Cliente_Id")), 14)
                            linha &= "|"
                            linha &= "|" & .Item("Estado")                                                                                   'Estado
                            linha &= "|" & Trim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
                            linha &= "|" & .Item("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                            linha &= "|" & Trim(.Item("InscricaoMunicipal"))
                            linha &= "|"                            'Suframa
                            linha &= "|" & DdlPerfil.SelectedValue
                            linha &= "|" & DdlAtividade.SelectedValue
                            linha &= "|"

                            HttpContext.Current.Session("Estado") = .Item("Estado")

                        End With

                        strm.WriteLine(linha)
                        Registro0000 += 1
                        RegistroGeral += 1
                    Next
                End If

                ' Registro 0001  - Abertura do Bloco 0

                linha = "|0001"
                linha &= "|0"
                linha &= "|"

                strm.WriteLine(linha)
                Registro0001 += 1
                RegistroGeral += 1

                ' Isola a verde agricola pois não é industria
                If PeriodoFinal.Year > 2019 AndAlso Left(Empresa(0), 8) <> "44979506" AndAlso Left(Empresa(0), 8) <> "62747840" AndAlso Left(Empresa(0), 8) <> "62780383" AndAlso Left(Empresa(0), 8) <> "63358210" Then
                    'Registro 0002  - CLASSIFICAÇÃO DO ESTABELECIMENTO INDUSTRIAL OU EQUIPARADO A INDUSTRIAL
                    linha = "|0002"

                    If Left(Empresa(0), 8) = "03189063" Then
                        linha &= "|01"
                    ElseIf Left(Empresa(0), 8) = "05272759" Then
                        linha &= "|00"
                    ElseIf Left(Empresa(0), 8) = "40938762" Then
                        linha &= "|00"
                    ElseIf Left(Empresa(0), 8) = "05366261" Then
                        linha &= "|01"
                    Else
                        linha &= "|09"
                    End If

                    linha &= "|"

                    strm.WriteLine(linha)
                    Registro0002 += 1
                    RegistroGeral += 1
                End If

                'Registro 0005 - Dados Complementares da Entidade
                Sql = "SELECT Clientes.Fantasia, Clientes.Endereco, Clientes.Numero, Clientes.Complemento, Clientes.Bairro, Clientes.Cep, Clientes.Telefone, isnull(Clientes.Fax,'') as Fax, isnull(Clientes.Email,'') as Email" & vbCrLf & _
                      "  FROM Clientes " & vbCrLf & _
                      " INNER JOIN ClientesXEmpresas" & vbCrLf & _
                      "    ON Clientes.Cliente_Id  = ClientesXEmpresas.Empresa_Id" & vbCrLf & _
                      "   AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
                      " INNER JOIN Estados" & vbCrLf & _
                      "    ON Clientes.Estado = Estados.Estado_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id ='" & Empresa(0) & "'" & vbCrLf & _
                      "  And Clientes.Endereco_Id = " & Empresa(1) & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|0005"
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Fantasia"))
                            linha &= "|" & .Item("Cep").ToString.Trim.Replace("-", "").Replace(".", "")
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Endereco"))
                            linha &= "|" & Trim(.Item("Numero"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Complemento"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Bairro"))
                            linha &= "|" & Funcoes.AlinharDireita(.Item("Telefone").ToString.Trim.Replace("(", "").Replace(")", "").Replace("-", ""), 10, "0")
                            linha &= "|" & Funcoes.AlinharDireita(.Item("Fax").ToString.Trim.Replace("(", "").Replace(")", "").Replace("-", ""), 10, "0")
                            linha &= "|" & Trim(.Item("Email"))
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0005 += 1
                        RegistroGeral += 1
                    Next
                End If


                'Registro 0015 - DADOS DO CONTRIBUINTE SUBSTITUTO OU RESPONSÁVEL PELO ICMS DESTINO
                Sql = "SELECT Estado_Id, IESubstitutoTributario" & vbCrLf &
                            "from ClienteXSubstitutoTributario" & vbCrLf &
                            "order by Estado_Id"

                ds = Banco.ConsultaDataSet(Sql, "ClientesST")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|0015"
                            linha &= "|" & Trim(.Item("Estado_Id"))
                            linha &= "|" & Trim(.Item("IESubstitutoTributario"))
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0015 += 1
                        RegistroGeral += 1
                    Next
                End If


                'Registro 0100 - Dados do Contabilista
                Sql = "SELECT ClientesXEmpresas.CPFContador, ClientesXEmpresas.CRCContador" & vbCrLf & _
                      "  FROM Clientes" & vbCrLf & _
                      " INNER JOIN ClientesXEmpresas" & vbCrLf & _
                      "    ON Clientes.Cliente_Id  = ClientesXEmpresas.Empresa_Id" & vbCrLf & _
                      "   AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id = '" & Empresa(0) & "'" & vbCrLf & _
                      "   And Clientes.Endereco_Id = " & Empresa(1) & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        Codigo = dr("CpfContador")
                        Crc = dr("CRCContador").Replace(".", "").Replace("-", "")
                    Next
                End If

                Sql = " SELECT DISTINCT Clientes.Nome," & vbCrLf & _
                      "                 Clientes.Cliente_Id," & vbCrLf & _
                      "                 Clientes.Cep, Clientes.Endereco, Clientes.Numero, Clientes.Complemento, Clientes.Bairro, isnull(Clientes.Telefone,'')as Telefone, isnull(Clientes.Fax,'')as Fax, isnull(Clientes.Email,'') as Email," & vbCrLf & _
                      "                 Municipios.Estadoibge, Clientes.CodigoDoMunicipio" & vbCrLf & _
                      "  FROM Clientes " & vbCrLf & _
                      " INNER JOIN Municipios" & vbCrLf & _
                      "    ON Clientes.Estado            = Municipios.Estado_id" & vbCrLf & _
                      "   AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id" & vbCrLf & _
                      "   AND Clientes.Cidade = Municipios.Municipio_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id = '" & Codigo & "'" & vbCrLf & _
                      "   And Clientes.Endereco_Id = 0"

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|0100"
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                            linha &= "|" & Trim(.Item("Cliente_Id"))
                            linha &= "|" & Crc
                            linha &= "|"
                            linha &= "|" & .Item("CEP").ToString.Trim.Replace("-", "").Replace(".", "")
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Endereco"))
                            linha &= "|" & Trim(.Item("Numero"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Complemento"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Bairro"))
                            linha &= "|" & Trim(.Item("Telefone").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""))
                            linha &= "|" & Trim(.Item("Fax").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""))
                            linha &= "|" & Trim(.Item("Email"))
                            linha &= "|" & .Item("Estadoibge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0100 += 1
                        RegistroGeral += 1
                    Next
                Else
                    MsgBox(Me.Page, "Contador não Cadastrado. Favor cadastrar em Cadastro de Clientes.")
                    Exit Sub
                End If

                '********************************************************************************************************************************************************************************************************************************
                '********************************************************************************************************************************************************************************************************************************
                '********************************************************************************************************************************************************************************************************************************
                '********************************************************************************************************************************************************************************************************************************
                '********************************************************************************************************************************************************************************************************************************
                '********************************************************************************************************************************************************************************************************************************
                'Registro 0150 - Tabela de Cadastro Do Participante

                If dsNew.Tables("REG0150").Rows.Count > 0 Then
                    For Each dr As DataRow In dsNew.Tables("REG0150").Rows
                        With dr
                            linha = "|0150"
                            linha &= "|" & Trim(.Item("Cliente_Id")) & Trim(.Item("Endereco_Id"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                            linha &= "|" & .Item("Pais")
                            If Microsoft.VisualBasic.Len(.Item("Cliente_Id")) = 11 And .Item("Pais") = 1058 Then
                                linha &= "|"
                                linha &= "|" & Trim(.Item("Cliente_Id"))
                            End If
                            If Microsoft.VisualBasic.Len(.Item("Cliente_Id")) = 14 And .Item("Pais") = 1058 Then
                                linha &= "|" & Trim(.Item("Cliente_Id"))
                                linha &= "|"
                            End If
                            If .Item("Pais") <> 1058 Then
                                linha &= "|"
                                linha &= "|"
                            End If

                            'linha &= "|" & .Item("Inscricao")
                            linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "").Replace("ISENTO", "").Replace("ISENTA", ""), 14, " "))    'Inscricao Estadual

                            If .Item("Pais") = 1058 Then
                                linha &= "|" & .Item("EstadoIBGE") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                            Else
                                linha &= "|"
                            End If

                            linha &= "|"   'Suframa
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Endereco"))
                            linha &= "|" & Trim(.Item("Numero"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Complemento"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Bairro"))
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0150 += 1
                        RegistroGeral += 1
                    Next
                End If


                If Left(Empresa(0), 8) = "05366261" AndAlso Format(CDate(txtDataFinal.Text), "MM") = 9 Then
                    'linha = "|0150"
                    'linha &= "|620099409970"
                    'linha &= "|CLAUDIMIR VEZZARO E LEDE FACHI VEZZARO"
                    'linha &= "|1058"
                    'linha &= "|"
                    'linha &= "|62009940997"
                    'linha &= "|9511092318"
                    'linha &= "|4104808"

                    'linha &= "|"   'Suframa
                    'linha &= "|LTS 31 E 29 4PER SAO FCO OU LOPEI"
                    'linha &= "|"
                    'linha &= "|SITIO VEZZARO"
                    'linha &= "|SEDE ALVORADA"
                    'linha &= "|"

                    'strm.WriteLine(linha)
                    'Registro0150 += 1
                    'RegistroGeral += 1

                    'linha = "|0150"
                    'linha &= "|007532199800"
                    'linha &= "|MAURO LUIS BONATTO"
                    'linha &= "|1058"
                    'linha &= "|"
                    'linha &= "|00753219980"
                    'linha &= "|9580764107"
                    'linha &= "|4104808"

                    'linha &= "|"   'Suframa
                    'linha &= "|RODOVIA PR-180"
                    'linha &= "|"
                    'linha &= "|LOTE RURAL 53-B-3 GLEBA 06 COLONIA TORMENTA"
                    'linha &= "|SAO SALVADOR"
                    'linha &= "|"

                    'strm.WriteLine(linha)
                    'Registro0150 += 1
                    'RegistroGeral += 1


                    linha = "|0150"
                    linha &= "|003354609680"
                    linha &= "|VALDIR KUCINSKI"
                    linha &= "|1058"
                    linha &= "|"
                    linha &= "|00335460968"
                    linha &= "|9516819542"
                    linha &= "|4104808"

                    linha &= "|"   'Suframa
                    linha &= "|FAZENDA GAUCHA"
                    linha &= "|"
                    linha &= "|LTS 51,52,53,56,57,58,71 E 72"
                    linha &= "|SEDE ALVORADA"
                    linha &= "|"

                    strm.WriteLine(linha)
                    Registro0150 += 1
                    RegistroGeral += 1

                End If


                'Registro 0190 - Identificação das Unidades de Medidas              
                If dsNew.Tables("REG0190").Rows.Count > 0 Then
                    For Each dr As DataRow In dsNew.Tables("REG0190").Rows
                        With dr
                            linha = "|0190"
                            linha &= "|" & Trim(.Item("Unidade"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0190 += 1
                        RegistroGeral += 1
                    Next
                End If

                'Registro 0200 - Tabela de Identificação do Item (Produtos e Serviços)
                If dsNew.Tables("REG0200").Rows.Count > 0 Then
                    For Each dr As DataRow In dsNew.Tables("REG0200").Rows
                        With dr
                            linha = "|0200"
                            linha &= "|" & Trim(.Item("Produto_Id"))
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                            linha &= "|"                    'Código de Barra
                            linha &= "|"                    'Código Anterior
                            linha &= "|" & Trim(.Item("Unidade"))
                            linha &= "|" & Funcoes.AlinharDireita(.Item("TipoDoItem"), 2, "0")
                            linha &= "|" & Left(Trim(.Item("NCM").Replace(".", "")), 8)
                            linha &= "|" & Trim(.Item("CodigoEX"))
                            If .Item("CodigoDoGenero") = 0 Then
                                linha &= "|"
                            Else
                                linha &= "|" & CInt(.Item("CodigoDoGenero")).ToString("00")
                            End If
                            linha &= "|" & IIf(PeriodoFinal.Year > 2014, String.Format("00.00", CInt(Trim(.Item("CodigoDoServico")))), Trim(.Item("CodigoDoServico")))
                            linha &= "|" & Trim(.Item("ICMS"))
                            linha &= "|"
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0200 += 1
                        RegistroGeral += 1
                    Next
                End If

                'Acrescentado em 07/02/2018 - furlan
                'Registro 0220 - Fatores de Conversão de Unidades
                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                    Sql = "SELECT ISNULL(PA.Produto_id, NotasFiscaisXItens.Produto_Id) AS Produto_Id" & vbCrLf & _
                            "INTO #Base" & vbCrLf & _
                            "FROM NotasFiscais" & vbCrLf & _
                            "		INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                            "				ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                            "				AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                            "				AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                            "				AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                            "				AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                            "				AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                            "				AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                            "		Inner Join OperacaoXEstado OE" & vbCrLf & _
                            "				ON OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf & _
                            "		LEFT JOIN NotaFiscalXExportacao AS NFEx " & vbCrLf & _
                            "				ON NotasFiscais.Empresa_Id       = NFEx.Empresa_Id " & vbCrLf & _
                            "				AND NotasFiscais.EndEmpresa_Id    = NFEx.EndEmpresa_Id " & vbCrLf & _
                            "				AND NotasFiscais.Cliente_Id       = NFEx.Cliente_Id " & vbCrLf & _
                            "				AND NotasFiscais.EndCliente_Id    = NFEx.EndCliente_Id " & vbCrLf & _
                            "				AND NotasFiscais.EntradaSaida_Id  = NFEx.EntradaSaida_Id " & vbCrLf & _
                            "				AND NotasFiscais.Serie_Id         = NFEx.Serie_Id " & vbCrLf & _
                            "				AND NotasFiscais.Nota_Id          = NFEx.Nota_Id " & vbCrLf & _
                            "		LEFT JOIN MemorandoDeExportacao AS ME " & vbCrLf & _
                            "				ON NotasFiscais.Empresa_Id = ME.Empresa " & vbCrLf & _
                            "				AND NotasFiscais.EndEmpresa_Id = ME.EndEmpresa " & vbCrLf & _
                            "				AND NotasFiscais.Cliente_Id = ME.Cliente " & vbCrLf & _
                            "				AND NotasFiscais.EndCliente_Id = ME.EndCliente" & vbCrLf & _
                            "				AND NotasFiscais.EntradaSaida_Id = ME.EntradaSaida " & vbCrLf & _
                            "				AND NotasFiscais.Serie_Id = ME.Serie " & vbCrLf & _
                            "				AND NotasFiscais.Nota_Id = ME.Nota" & vbCrLf & _
                            "		LEFT JOIN ProdutosAgrupados pa" & vbCrLf & _
                            "			ON NotasFiscaisXItens.Produto_Id = pa.ProdutoAgrupado_Id" & vbCrLf & _
                            "		LEFT JOIN NFERealizadas" & vbCrLf & _
                            "				ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                            "				AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                            "				AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                            "				AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                            "				AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                            "				AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                            "				AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                            "		INNER JOIN SubOperacoes" & vbCrLf & _
                            "				ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf & _
                            "				AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                            "WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                            "AND NotasFiscais.Situacao not in(2,9,10) " & vbCrLf & _
                            "AND NotasFiscais.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                            "AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                            "AND NotasFiscais.TipoDeDocumento NOT IN (15)" & vbCrLf &
                            "AND NotasFiscais.Finalidade NOT IN (30)" & vbCrLf &
                            "AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT IN (1932,2932))" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                            "AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
                            "AND Not (NotasFiscais.EntradaSaida_Id = 'E'  AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf & _
                            "AND NOT (SubOperacoes.Classe IN ('SERVICOS') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf & _
                            "AND NOT (SubOperacoes.Classe IN ('SERVICOS') AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf & _
                            "AND ((NFERealizadas.Chavenfe is null) " & vbCrLf & _
                            "        or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null)" & vbCrLf & _
                            "        or (ISNULL(ME.DataAverbacao, NFEx.DataAverbacao) BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'))" & vbCrLf & _
                            "OR (NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' AND NotasFiscais.TipoDeDocumento=2 AND NotasFiscais.NossaEmissao='S') " & vbCrLf & _
                            "group by ISNULL(PA.Produto_id, NotasFiscaisXItens.Produto_Id), OE.CodigoFiscal" & vbCrLf & _
                            "having sum(NotasFiscaisxitens.quantidadefiscal) > 0" & vbCrLf

                    Sql &= "INSERT INTO #Base" & vbCrLf & _
                            "SELECT ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                            "FROM ApuracaoDeCustos  " & vbCrLf & _
                            "		INNER JOIN PRODUTOS P" & vbCrLf & _
                            "				ON P.Produto_id = ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                            "Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                            "    And ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                            "    And ApuracaoDeCustos.Mes_Id =  " & CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Month & vbCrLf & _
                            "    And ApuracaoDeCustos.Ano_Id = " & CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year & vbCrLf & _
                            "    And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                            "    And ApuracaoDeCustos.ValorDoProduto > 0" & vbCrLf & _
                            "    And ISNULL(P.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf

                    Sql &= "SELECT DISTINCT PxU.Unidade_Id as Unidade," & vbCrLf &
                            "	   case" & vbCrLf &
                            "	      When (PxU.Unidade_Id = 'UN' OR PxU.Unidade_Id = 'UND')" & vbCrLf &
                            "	         then 1" & vbCrLf &
                            "	         else PxU.FatorConversao_Id" & vbCrLf &
                            "	   end as FatorConversao" & vbCrLf &
                            "  FROM #Base P " & vbCrLf &
                            "		INNER JOIN Produtos Prd" & vbCrLf &
                            "				ON Prd.Produto_id = P.Produto_Id" & vbCrLf &
                            "		LEFT JOIN ProdutosxUnidadeDeComercializacao PxU" & vbCrLf &
                            "				ON PxU.Produto_id = P.Produto_Id" & vbCrLf &
                            "GROUP BY PxU.Unidade_Id, PxU.FatorConversao_Id" & vbCrLf

                    ds = Banco.ConsultaDataSet(Sql, "Consulta0220")

                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr As DataRow In ds.Tables(0).Rows

                            If dsNew.Tables("REG0190").Select("Unidade = '" + dr("Unidade") + "'").Count > 0 Then

                                With dr
                                    linha = "|0220"
                                    linha &= "|" & .Item("Unidade")
                                    linha &= "|" & Math.Round(CDec(.Item("FatorConversao")), 6).ToString

                                    'CÓDIGO DE BARRAS LAYOUT À PARTIR DE 2022 - FURLAN - 09/02/2022
                                    If PeriodoFinal.Year > 2021 Then linha &= "|"

                                    linha &= "|"
                                End With

                                strm.WriteLine(linha)
                                Registro0220 += 1
                                RegistroGeral += 1

                            End If
                        Next
                    End If
                End If

                '******************************************************************************************************************************************************************************************************************************************
                '******************************************************************************************************************************************************************************************************************************************
                '******************************************************************************************************************************************************************************************************************************************
                '******************************************************************************************************************************************************************************************************************************************
                '******************************************************************************************************************************************************************************************************************************************
                '******************************************************************************************************************************************************************************************************************************************
                '******************************************************************************************************************************************************************************************************************************************

                'Registro 0400 - Tabela de natureza da operação/prestação
                Sql = "  SELECT distinct SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao" & vbCrLf &
                      "    FROM NotasFiscais" & vbCrLf &
                      "   Inner Join NotasFiscaisXItens" & vbCrLf &
                      "      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf &
                      "     AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                      "     AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
                      "     AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                      "     AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                      "     AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                      "     AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                      "   Inner Join OperacaoXEstado OE" & vbCrLf &
                      "      on OE.Codigo_id =  NotasFiscaisXItens.OperacaoXEstado" & vbCrLf &
                      "    LEFT JOIN NFERealizadas" & vbCrLf &
                      "      ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf &
                      "     AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf &
                      "     AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf &
                      "     AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf &
                      "     AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf &
                      "     AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf &
                      "     AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id     " & vbCrLf &
                      "   INNER JOIN SubOperacoes " & vbCrLf &
                      "      ON isnull(NotasFiscaisXItens.Operacao,NotasFiscais.Operacao) = SubOperacoes.Operacao_Id " & vbCrLf &
                      "     AND isnull(NotasFiscaisXItens.SubOperacao,NotasFiscais.SubOperacao) = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                      "   WHERE NotasFiscais.Serie_Id Not IN ('D', 'F','REC')" & vbCrLf &
                      "     AND (NotasFiscais.Empresa_Id =  '" & Empresa(0) & "')" & vbCrLf &
                      "     AND (NotasFiscais.Operacao > 0) " & vbCrLf &
                      "     AND (NotasFiscais.Situacao NOT IN(2,4,9)) " & vbCrLf

                If Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "     AND (NotasFiscais.TipoDeDocumento NOT IN(2,8,10,11,14,15,16,57,58,65))" & vbCrLf

                Else
                    Sql &= "     AND (NotasFiscais.TipoDeDocumento NOT IN (15))" & vbCrLf
                End If

                Sql &= "     AND (NotasFiscais.Finalidade NOT IN (30))" & vbCrLf &
                      "     AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT IN (1932,2932))" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                      "     AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                      "     AND (NotasFiscais.NossaEmissao = 'N')" & vbCrLf &
                      "     AND Not (NotasFiscais.EntradaSaida_Id = 'E'  AND  OE.CodigoFiscal IN (1933,2933))    " & vbCrLf &
                      "     AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                      "   --AND Not (NotasFiscais.EntradaSaida_Id = 'S'   AND  OE.CodigoFiscal IN (5933,6933)    " & vbCrLf &
                      "     AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                      "   Group by SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao" & vbCrLf &
                      "  having(SUM(NotasFiscaisXItens.QuantidadeFiscal) > 0)"

                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|0400"
                            linha &= "|" & .Item("Operacao_Id") & .Item("SubOperacoes_Id")
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0400 += 1
                        RegistroGeral += 1
                    Next
                End If

                'Registro 0450 - Tabela de informações complementares do documento fiscal
                'Sql = " SELECT   codigo_id, Descricao"
                'Sql &= " FROM    ObservacoesTributarias"
                'Sql &= " WHERE   (Estado = '" & HttpContext.Current.Session("Estado") & "')"


                Sql = "SELECT codigo_id, Descricao" & vbCrLf & _
                      "  FROM ObservacoesTributarias" & vbCrLf & _
                      " WHERE codigo_id IN      (" & vbCrLf & _
                      "                          SELECT OT.codigo_id" & vbCrLf & _
                      "                            FROM NotasFiscais " & vbCrLf & _
                      "						   INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                      "						      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                      "                        Inner Join OperacaoXEstado OE" & vbCrLf & _
                      "                           on OE.Codigo_id =  NotasFiscaisXItens.OperacaoXEstado" & vbCrLf & _
                      "						   INNER JOIN ObservacoesTributarias AS OT " & vbCrLf & _
                      "                              ON CONVERT(varchar(400), OT.Descricao) = CONVERT(varchar(400), NotasFiscais.ObservacoesDeEmbarque)" & vbCrLf & _
                      "							LEFT JOIN NFERealizadas" & vbCrLf & _
                      "							  ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                      "						   WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                      "							 AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & _
                      "							 AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & _
                      "							 AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                      "                          AND (NotasFiscais.Situacao not in(9,10)) " & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT BETWEEN 2350 AND 2360)" & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT IN (1932,2932))" & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                      "							 AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                      "							 AND ((NFERealizadas.Chavenfe is null) or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null))" & vbCrLf & _
                      "                           Group by OT.codigo_id" & vbCrLf & _
                      "                          having sum(NotasFiscaisxitens.quantidadefiscal) > 0)" & vbCrLf

                'Sql = "select Codigo_id, Descricao from ObservacoesTributarias where codigo_id in (" & _
                '      "select distinct OT.Codigo_id" & _
                '      "  from notasfiscais" & _
                '      " inner Join ObservacoesTributarias OT" & _
                '      "    on convert(varchar,OT.Descricao) = convert(varchar,notasfiscais.observacoesdeembarque)" & _
                '      "  WHERE NotasFiscais.Operacao > 0 And NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND " & _
                '      "        NotasFiscais.Movimento BETWEEN '" & Format(PeriodoInicial, "yyyy/MM/dd") & "' AND '" & Format(PeriodoFinal, "yyyy/MM/dd") & "'" & _
                '      "    And NotasFiscais.Serie_Id not in('D','F') )"

                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|0450"
                            linha &= "|" & .Item("Codigo_Id")
                            linha &= "|" & Left(Funcoes.EliminarCaracteresEspeciais(.Item("Descricao")), 200)
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0450 += 1
                        RegistroGeral += 1
                    Next
                End If

                'Registro 0460 - Tabela de informações complementares do documento fiscal
                Sql = "SELECT codigo_id, Descricao" & vbCrLf & _
                      "  FROM ObservacoesDoLancamentoFiscais " & vbCrLf & _
                      " WHERE codigo_id IN      (" & vbCrLf & _
                      "                          SELECT OT.codigo_id" & vbCrLf & _
                      "                            FROM NotasFiscais " & vbCrLf & _
                      "						   INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                      "						      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                      "						   INNER JOIN ObservacoesDoLancamentoFiscais AS OT " & vbCrLf & _
                      "                              ON OT.Codigo_Id = NotasFiscais.ObsLanctoFiscal " & vbCrLf & _
                      "							LEFT JOIN NFERealizadas" & vbCrLf & _
                      "							  ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                      "							 AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                      "							 AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                      "						   WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                      "							 AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & _
                      "							 AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & _
                      "							 AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                      "                          AND (NotasFiscais.Situacao not in(9,10)) " & vbCrLf & _
                      "                           Group by OT.codigo_id" & vbCrLf & _
                      "                          having sum(NotasFiscaisxitens.quantidadefiscal) > 0)" & vbCrLf

                Dim ds13 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds13.Tables(0).Rows.Count > 0 Then
                    For Each dr13 As DataRow In ds13.Tables(0).Rows
                        With dr13
                            linha = "|0460"
                            linha &= "|" & .Item("Codigo_Id")
                            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        Registro0460 += 1
                        RegistroGeral += 1
                    Next
                End If

                '' Registro 0990  - Encerramento do Bloco 0
                Registro0990 += 1

                linha = "|0990"

                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2023 Then
                    linha &= "|" & Registro0000 + Registro0001 + Registro0002 + Registro0005 + Registro0015 + Registro0100 + Registro0150 + Registro0190 + Registro0200 + Registro0220 + Registro0400 + Registro0450 + Registro0460 + Registro0990
                ElseIf CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2019 Then
                    linha &= "|" & Registro0000 + Registro0001 + Registro0002 + Registro0005 + Registro0100 + Registro0150 + Registro0190 + Registro0200 + Registro0220 + Registro0400 + Registro0450 + Registro0460 + Registro0990
                ElseIf CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                    linha &= "|" & Registro0000 + Registro0001 + Registro0005 + Registro0100 + Registro0150 + Registro0190 + Registro0200 + Registro0220 + Registro0400 + Registro0450 + Registro0460 + Registro0990
                Else
                    linha &= "|" & Registro0000 + Registro0001 + Registro0005 + Registro0100 + Registro0150 + Registro0190 + Registro0200 + Registro0400 + Registro0450 + Registro0460 + Registro0990
                End If

                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1

                '  Fim do Bloco 0


                ' Abertura do Bloco B

                ' Registro B001  - Abertura do Bloco 
                linha = "|B001"
                linha &= "|" & 1       'Indicador de Movimento - 0 - Bloco com Dados Informados
                '                                                1 - Bloco sem dados informados
                linha &= "|"

                strm.WriteLine(linha)
                RegistroB001 += 1
                RegistroGeral += 1


                ' Registro B990  - Encerramento do Bloco B
                RegistroB990 += 1

                linha = "|B990"
                linha &= "|" & RegistroB001 + RegistroB990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1



                ' Abertura do Bloco C

                'Registro C001 - Confere se existe registros no C100 e C500
                Sql = "SELECT case" & vbCrLf &
                      "         when (not nf.TipoDeDocumento in (57)) and Count(*) > 0" & vbCrLf &
                      "           then 0" & vbCrLf &
                      "           else 1" & vbCrLf &
                      "       end Existe" & vbCrLf &
                      "  FROM NotasFiscais nf" & vbCrLf &
                      " Inner Join NotasFiscaisXitens nfi" & vbCrLf &
                      "    ON nf.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "   AND nf.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "   AND nf.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "   AND nf.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "   AND nf.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "   AND nf.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      "   AND nf.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      " Inner Join OperacaoXEstado OE" & vbCrLf &
                      "    on OE.Codigo_id =  nfi.OperacaoXEstado" & vbCrLf &
                      " INNER JOIN SubOperacoes so" & vbCrLf &
                      "    ON nfi.Operacao    = so.Operacao_Id" & vbCrLf &
                      "   AND nfi.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                      " WHERE nf.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "   AND nf.Operacao  > 0" & vbCrLf &
                      "   AND nf.Situacao not in(9,10)" & vbCrLf &
                      "   AND nf.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                      "   AND nf.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                      "   AND nf.TipoDeDocumento NOT IN (15)" & vbCrLf &
                      "   AND nf.Finalidade NOT IN (30)" & vbCrLf &
                      "   --AND(OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                      "   AND OE.CodigoFiscal NOT BETWEEN 1350 AND 1360" & vbCrLf &
                      "   AND OE.CodigoFiscal NOT BETWEEN 2300 AND 2360" & vbCrLf &
                      "   AND OE.CodigoFiscal NOT IN (1932,2932)" & vbCrLf &
                      "   AND OE.CodigoFiscal NOT BETWEEN 5350 AND 5360" & vbCrLf &
                      "   AND OE.CodigoFiscal NOT BETWEEN 6350 AND 6360" & vbCrLf &
                      "   AND OE.CodigoFiscal NOT BETWEEN 1302 AND 1303" & vbCrLf &
                      "   AND NOT(nfi.EntradaSaida_Id = 'E' AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf &
                      "   AND NOT(so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                      "   --AND NOT(nfi.EntradaSaida_Id = 'S' AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf &
                      "   AND NOT(so.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                      "   GROUP BY nf.TipoDeDocumento"
                ' AND(OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)  esta linha esta no bloco c100 + nao estava nesse where entao adicionei ele acima

                ds = Banco.ConsultaDataSet(Sql, "Consulta") ' tem que verificar se o campo não está vindo com algum dado, por que é feito apenas uma validação simples de count > 0

                RegistroC001 += 1

                linha = "|C001"

                If ds.Tables(0).Rows.Count > 0 Then
                    linha &= "|" & ds.Tables(0).Rows(0)(0)  '0 - Com Dados  /  1 - Sem Dados
                Else
                    linha &= "|1"  '0 - Com Dados  /  1 - Sem Dados
                End If

                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1


                'RegistroC100 - Nota Fiscal (Código 01), Nota Fiscal Avulsa (Código 1B), 
                '                Nota Fiscal de Produtor (Código 04) e Nfe (código 55)

                ds = ConsultaRegistroC100_C170_C172_C190()

                If ds.Tables("C100").Rows.Count > 0 Then

                    Dim C100NF As Integer = 0
                    Dim C100Serie As String = ""
                    Dim C100EntradaSaida As String = ""
                    Dim C100Cliente As String = ""
                    Dim C100EndCliente As Integer = 0

                    For Each dr As DataRow In ds.Tables("C100").Rows
                        With dr

                            linha = "|C100"  'Registro 01 - Texto fixo contendo "C100"

                            'Registro 02 - Indicador do tipo de operação: 0 - Entrada; 1 - Saída
                            If .Item("EntradaSaida_Id") = "E" Then
                                linha &= "|0"
                            Else
                                linha &= "|1"
                            End If

                            'Registro 03 - Indicador do emitente do documento fiscal: 0 - Emissão própria; 1 - Terceiros
                            If .Item("NossaEmissao") = "S" Then
                                NossaEmissao = True
                                linha &= "|1"
                            Else
                                NossaEmissao = False
                                linha &= "|0"
                            End If

                            'Registro 04 - Código do participante (campo 02 do Registro 0150): - do emitente do documento ou do remetente das mercadorias, no caso de entradas; - do adquirente, no caso de saídas
                            'If .Item("ValorTotal") <> 0 Then
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf .Item("Situacao") = 1 Then    'Situação do Documento Fiscal
                                linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                            Else
                                linha &= "|"
                            End If
                            'Else
                            '    linha &= "|"
                            'End If

                            'Registro 05 - Código do modelo do documento fiscal, conforme a Tabela 4.1.1
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|65"
                                Modelo = "65"
                            ElseIf Len(.Item("ChaveNfe")) < 10 Then
                                linha &= "|01"
                                Modelo = "01"
                            Else
                                linha &= "|55"
                                Modelo = "55"
                            End If

                            'Registro 06 - Código da situação do documento fiscal, conforme a Tabela 4.1.2
                            If .Item("Situacao") = 1 Then    'Situação do Documento Fiscal
                                'Testando tipo de emissor sefaz
                                If Len(.Item("ChaveNfe")) = 44 AndAlso RTrim(.Item("Serie_Id")) = "890" AndAlso
                                                                        (Mid(.Item("ChaveNfe"), 7, 14) = "02935843000105" Or
                                                                        Mid(.Item("ChaveNfe"), 7, 14) = "78393592000146" Or
                                                                        Mid(.Item("ChaveNfe"), 7, 14) = "82951310000156" Or
                                                                        Mid(.Item("ChaveNfe"), 7, 14) = "87958674000181" Or
                                                                        Mid(.Item("ChaveNfe"), 7, 14) = "58290502000184" Or
                                                                        Mid(.Item("ChaveNfe"), 7, 14) = "76416890000189" Or
                                                                        Mid(.Item("ChaveNfe"), 7, 14) = "03507415000578") Then
                                    linha &= "|08"
                                ElseIf .Item("QuantidadeFiscal") = 0 Then
                                    linha &= "|06"
                                Else
                                    linha &= "|00"
                                End If
                                Cancelada = "N"
                            Else
                                linha &= "|02"
                                Cancelada = "S"
                            End If

                            'Registro 07 - Série do documento fiscal
                            linha &= "|" & RTrim(.Item("Serie_Id"))
                            Dim n As String = ""
                            n = .Item("Nota_Id")
                            'Registro 08 - Número do documento fiscal
                            linha &= "|" & .Item("Nota_Id")

                            'Registro 09 - Chave da Nota Fiscal Eletrônica
                            linha &= "|" & .Item("ChaveNfe")

                            'Registro 10 - Data da emissão do documento fiscal
                            If Cancelada = "N" Then
                                linha &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy")
                            Else
                                linha &= "|"
                            End If

                            'Registro 11 - Data da entrada ou da saída
                            If Cancelada = "N" Then
                                linha &= "|" & Format(CDate(.Item("Movimento")), "ddMMyyyy")
                            Else
                                linha &= "|"
                            End If

                            'Registro 12 - Valor total do documento fiscal
                            If Cancelada = "N" Then
                                linha &= "|" & .Item("ValorTotal")
                            Else
                                linha &= "|"
                            End If

                            'Registro 13 - Indicador do tipo de pagamento: 0 - À vista; 1 - A prazo; 9 - Sem pagamento. C 001* - O O Obs.: A partir de 01/07/2012 passará a ser: Indicador do tipo de pagamento: 0 - À vista; 1 - A prazo; 2 - Outros
                            If Cancelada = "N" Then
                                linha &= "|1"
                            Else
                                linha &= "|"
                            End If

                            'Registro 14 - Valor total do desconto
                            If Cancelada = "N" Then
                                linha &= "|" & .Item("ValorDesconto")   'Valor Total do Desconto
                            Else
                                linha &= "|"
                            End If

                            'Registro 15 - Abatimento não tributado e não comercial Por exemplo: desconto ICMS nas remessas para ZFM.
                            linha &= "|"                        'Valor Total do Abatimento

                            'Registro 16 - Valor total das mercadorias e serviços
                            If Cancelada = "N" Then
                                linha &= "|" & .Item("ValorBruto")
                            Else
                                linha &= "|"
                            End If

                            'Registro 17 - Indicador do tipo do frete:
                            'Obs: A partir de 01/01/2018 passará a ser: Indicador do tipo de frete:
                            '0 - Contratação do Frete por conta do Remetente (CIF);
                            '1 - Contratação do Frete por conta do Destinatário (FOB);
                            '2 - Contratação do Frete por conta de Terceiros;
                            '3 - Transporte Próprio por conta do Remetente;
                            '4 - Transporte Próprio por conta do Destinatário;
                            '9 - Sem Ocorrência de Transporte.
                            If Cancelada = "N" Then
                                'linha &= "|1"                       'Identificador do Tipo de Frete

                                'Entrada
                                If .Item("EntradaSaida_Id") = "E" Then
                                    'Frete por conta da Empresa
                                    If .Item("CIFFobNf") = "FOB" Then
                                        linha &= "|1"
                                    ElseIf .Item("CIFFobNf") = "CIF" Then 'Frete por conta do Fornecedor
                                        linha &= "|0"
                                    Else
                                        linha &= "|9"
                                    End If
                                Else 'Saída

                                    'Frete por conta da Empresa
                                    If .Item("CIFFobNf") = "CIF" Then
                                        linha &= "|1"
                                    ElseIf .Item("CIFFobNf") = "FOB" Then 'Frete por conta do Cliente
                                        linha &= "|0"
                                    Else
                                        linha &= "|9"
                                    End If
                                End If
                            Else
                                linha &= "|"
                            End If

                            'Registro 18 - Valor do frete indicado no documento fiscal
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorFrete")
                            Else
                                linha &= "|"
                            End If

                            'Registro 19 - Valor do seguro indicado no documento fiscal
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorSeguro")
                            Else
                                linha &= "|"
                            End If

                            'Registro 20 - Valor de outras despesas acessórias
                            linha &= "|"

                            'Registro 21 - Valor da base de cálculo do ICMS
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("BaseIcms")
                            Else
                                linha &= "|"
                            End If

                            'Registro 22 - Valor do ICMS
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorIcms")
                            Else
                                linha &= "|"
                            End If

                            'Registro 23 - Valor da base de cálculo do ICMS substituição tributária
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("BaseIcmsST")
                            Else
                                linha &= "|"
                            End If

                            'Registro 24 - Valor do ICMS retido por substituição tributária
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorIcmsST")
                            Else
                                linha &= "|"
                            End If

                            'Registro 25 - Valor total do IPI
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorIPI")
                            Else
                                linha &= "|"
                            End If

                            'Registro 26 - Valor total do PIS
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorPIS")
                            Else
                                linha &= "|"
                            End If

                            'Registro 27 - Valor total da COFINS
                            If .Item("TipoDeDocumento") = 65 Then
                                linha &= "|"
                            ElseIf Cancelada = "N" Then
                                linha &= "|" & .Item("ValorCofins")
                            Else
                                linha &= "|"
                            End If

                            'Registro 28 - Valor total do PIS retido por substituição tributária
                            linha &= "|"
                            'Registro 29 - Valor total da COFINS retido por sustituição tributária
                            linha &= "|"
                            linha &= "|"

                            strm.WriteLine(linha)
                            RegistroC100 += 1
                            RegistroGeral += 1

                            C100NF = .Item("Nota_Id")
                            C100Serie = .Item("Serie_Id")
                            C100EntradaSaida = .Item("EntradaSaida_Id")
                            C100Cliente = .Item("Cliente_Id")
                            C100EndCliente = .Item("EndCliente_Id")


                            'Registro C101 - Notas Fiscais X Encargos  - Informação complementar dos documentos fiscais quando das operações interestaduais destinadas a consumidor final não contribuinte EC 87/15 (código 55)
                            If ds.Tables("C101").Rows.Count > 0 Then
                                CompoeRegistroC101(ds, strm, linha, RegistroC101, RegistroGeral, C100NF, C100Cliente, C100EndCliente, C100Serie, C100EntradaSaida)
                            End If

                            'Registro C170 - Notas Fiscais X Itens
                            If ds.Tables("C170").Rows.Count > 0 Then
                                CompoeRegistroC170(ds, strm, linha, RegistroC170, RegistroGeral, C100NF, C100Cliente, C100EndCliente, C100Serie, C100EntradaSaida)
                            End If

                            'Registro C172 - OPERAÇÕES COM ISSQN (CÓDIGO 01)
                            If ds.Tables("C172").Rows.Count > 0 Then
                                CompoeRegistroC172(ds, strm, linha, RegistroC172, RegistroGeral, C100NF, C100Cliente, C100EndCliente, C100Serie, C100EntradaSaida)
                            End If

                            'Registro C190 - Notas Fiscais X Itens
                            If ds.Tables("C190").Rows.Count > 0 Then
                                CompoeRegistroC190(ds, strm, linha, RegistroC190, RegistroGeral, C100NF, C100Cliente, C100EndCliente, C100Serie, C100EntradaSaida)
                            End If
                        End With
                    Next
                End If


                ''Registro C195 - OBSERVAÇOES DO LANÇAMENTO FISCAL (CÓDIGO 01, 1B E 55)
                'Sql = "SELECT codigo_id, Descricao" & vbCrLf & _
                '                  "  FROM ObservacoesDoLancamentoFiscais " & vbCrLf & _
                '                  " WHERE codigo_id IN      (" & vbCrLf & _
                '                  "                          SELECT OT.codigo_id" & vbCrLf & _
                '                  "                            FROM NotasFiscais " & vbCrLf & _
                '                  "						   INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                '                  "						      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                '                  "						   INNER JOIN ObservacoesDoLancamentoFiscais AS OT " & vbCrLf & _
                '                  "                              ON OT.Codigo_Id = NotasFiscais.ObsLanctoFiscal " & vbCrLf & _
                '                  "							LEFT JOIN NFERealizadas" & vbCrLf & _
                '                  "							  ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf

                'Sql &= " WHERE NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
                'Sql &= "       NotasFiscais.Cliente_Id = '" & Cliente & "' AND NotasFiscais.EndCliente_Id = " & EndCliente & " And "
                'Sql &= "       NotasFiscais.Nota_Id = " & Nota & " AND NotasFiscais.Serie_Id = '" & Serie & "' And "
                'Sql &= "       NotasFiscais.EntradaSaida_Id = '" & EntradaSaida & "' )"
                ''NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                ''                              "							 AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & _
                ''                              "							 AND NotasFiscais.Movimento BETWEEN '" & Format(PeriodoInicial, "yyyy/MM/dd") & "' AND '" & Format(PeriodoFinal, "yyyy/MM/dd") & "'" & _
                ''                              "							 AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                ''                              "                          AND (NotasFiscais.Situacao <> 9) " & vbCrLf & _
                ''                              "                           Group by OT.codigo_id" & vbCrLf & _
                ''                              "                          having sum(NotasFiscaisxitens.quantidadefiscal) > 0)" & vbCrLf

                'ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                'If ds2.Tables(0).Rows.Count > 0 Then
                '    For Each dr2 As DataRow In ds2.Tables(0).Rows
                '        With dr2
                '            linha = "|C195"
                '            linha &= "|" & .Item("Codigo_Id")
                '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
                '            linha &= "|"
                '        End With

                '        strm.WriteLine(linha)
                '        RegistroC195 += 1
                '        RegistroGeral += 1
                '    Next
                'End If


                ''Registro C197 - 'OUTRAS OBRIGAÇÕES TRIBUTÁRIAS, AJUSTES E INFORMAÇÕES DE VALORES PROVENIENTES DE DOCUMENTO FISCAL.
                'Sql = "SELECT COD_AJ, DESCR_COMPL_APUR, COD_ITEM, VL_BC_ICMS, ALIQ_ICMS, VL_ICMS, VL_OUTROS " & vbCrLf & _
                '                  "  FROM " & vbCrLf & _
                '                  " (" & vbCrLf & _
                '                  "                 SELECT NotasFiscais.CodigoAjusteApuracao as COD_AJ, OT.Descricao as DESCR_COMPL_APUR, " & vbCrLf & _
                '                  "                       NotasFiscaisXEncargos.Produto_Id as COD_ITEM, " & vbCrLf & _
                '"                                         isnull(sum(case " & vbCrLf & _
                '"											 when UPPER(Encargo_Id) like '%OUTROSCREDITOSICMS%'" & vbCrLf & _
                '"											   then isnull(NotasFiscaisXEncargos.Base,0)" & vbCrLf & _
                '"											   else 0" & vbCrLf & _
                '"										   end),0) as VL_BC_ICMS,  " & vbCrLf & _
                '"                                         isnull(sum(case " & vbCrLf & _
                '"											 when UPPER(Encargo_Id) like '%OUTROSCREDITOSICMS%'" & vbCrLf & _
                '"											   then isnull(NotasFiscaisXEncargos.Percentual,0)" & vbCrLf & _
                '"											   else 0" & vbCrLf & _
                '"										   end),0) as ALIQ_ICMS," & vbCrLf & _
                '"                                         isnull(sum(case " & vbCrLf & _
                '"											 when UPPER(Encargo_Id) like '%OUTROSCREDITOSICMS%'" & vbCrLf & _
                '"											   then isnull(NotasFiscaisXEncargos.Valor,0)" & vbCrLf & _
                '"											   else 0" & vbCrLf & _
                '"										   end),0) as VL_ICMS," & vbCrLf & _
                '                  "                       OT.Descricao as VL_OUTROS " & vbCrLf & _
                '                  "                            FROM NotasFiscais " & vbCrLf & _
                '                  "						   INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                '                  "						      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                '                   "						   INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
                '                  "						      ON NotasFiscais.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Serie_Id        = NotasFiscaisXEncargos.Serie_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Nota_Id         = NotasFiscaisXEncargos.Nota_Id " & vbCrLf & _
                '                  "						   INNER JOIN AjustesDaApuracaoIcms AS OT " & vbCrLf & _
                '                  "                              ON OT.Codigo_Id = NotasFiscais.CodigoAjusteApuracao " & vbCrLf & _
                '                  "							LEFT JOIN NFERealizadas" & vbCrLf & _
                '                  "							  ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                '                  "							 AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf

                'Sql &= " WHERE NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
                'Sql &= "       NotasFiscais.Cliente_Id = '" & Cliente & "' AND NotasFiscais.EndCliente_Id = " & EndCliente & " And "
                'Sql &= "       NotasFiscais.Nota_Id = " & Nota & " AND NotasFiscais.Serie_Id = '" & Serie & "' And "
                'Sql &= "       NotasFiscais.EntradaSaida_Id = '" & EntradaSaida & "'  "
                'Sql &= "     AND (NotasFiscais.Operacao > 0) "
                'Sql &= "     AND (NotasFiscais.Situacao <> 9) "
                'Sql &= "  Group by NotasFiscais.CodigoAjusteApuracao , OT.Descricao, NotasFiscaisXEncargos.Produto_Id, OT.Descricao ) AS CONSULTA"

                'ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                'If ds2.Tables(0).Rows.Count > 0 Then
                '    For Each dr2 As DataRow In ds2.Tables(0).Rows
                '        With dr2
                '            linha = "|C197"
                '            linha &= "|" & Funcoes.AlinharDireita(.Item("COD_AJ"), 10, " ")
                '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(LTrim(RTrim(.Item("DESCR_COMPL_APUR"))))
                '            linha &= "|" & IIf((Modelo <> "55" And Cancelada = "N") Or (Modelo = "55" And Not NossaEmissao), "", "") '.Item("COD_ITEM"), " ")
                '            linha &= "|" & .Item("VL_BC_ICMS")
                '            linha &= "|" & .Item("ALIQ_ICMS")
                '            linha &= "|" & .Item("VL_ICMS")
                '            linha &= "|0" '& .Item("VL_OUTROS")
                '            linha &= "|"
                '        End With

                '        strm.WriteLine(linha)
                '        RegistroC197 += 1
                '        RegistroGeral += 1
                '    Next
                'End If


                'Next
                'End If

                'Registro C500 - Nota Fiscal (Código 06), Nota Fiscal de Energia Eletrica), 
                Sql = "SELECT NF.EntradaSaida_Id," & vbCrLf &
                      "       NF.Cliente_Id, NF.EndCliente_Id," & vbCrLf &
                      "	      NF.Serie_Id, NF.Nota_Id," & vbCrLf &
                      "       NF.NossaEmissao, ISNULL(NFERealizadas.ChaveNfe, '') AS ChaveNfe," & vbCrLf &
                      "	      NF.DataDaNota, NF.Movimento, NF.TipoDeDocumento," & vbCrLf &
                      "       sum(Enc.ValorTotal)   as ValorTotal," & vbCrLf &
                      "	      sum(Enc.BaseIcms)     as BaseIcms," & vbCrLf &
                      "	      sum(Enc.Aliquota)     as Aliquota," & vbCrLf &
                      "	      sum(Enc.ValorIcms)    as ValorIcms," & vbCrLf &
                      "	      sum(Enc.ValorIPI)     as ValorIPI," & vbCrLf &
                      "	      sum(Enc.ValorPIS)     as ValorPIS," & vbCrLf &
                      "	      sum(Enc.ValorCofins)  as ValorCofins" & vbCrLf &
                      "  FROM NotasFiscais NF" & vbCrLf &
                      " Inner Join NotasFiscaisXItens nfi" & vbCrLf &
                      "    on NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                      "   and NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                      "   and NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                      "   and NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                      "   and NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                      "   and NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                      "   and NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                      " INNER JOIN OperacaoXEstado OE" & vbCrLf &
                      "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                      " INNER JOIN (" & vbCrLf &
                      "				SELECT NFxE.Empresa_Id, NFxE.EndEmpresa_Id," & vbCrLf &
                      "				       NFxE.Cliente_Id, NFxE.EndCliente_Id," & vbCrLf &
                      "					   NFxE.EntradaSaida_Id, NFxE.Nota_Id, NFxE.Serie_Id," & vbCrLf &
                      "					   NFxE.Produto_id, NFxE.Sequencia_id," & vbCrLf &
                      "					   isnull(sum(case" & vbCrLf &
                      "									 when isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id) = 'PRODUTO'" & vbCrLf &
                      "									   then NFxE.Valor" & vbCrLf &
                      "									   else 0" & vbCrLf &
                      "								   end),0) as ValorTotal," & vbCrLf &
                      "					   isnull(sum(case" & vbCrLf &
                      "									 when isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id) = 'ICMS'" & vbCrLf &
                      "									   then isnull(NFxE.Base,0)" & vbCrLf &
                      "									   else 0" & vbCrLf &
                      "								   end),0) as BaseIcms," & vbCrLf &
                      "					   0 AS Aliquota," & vbCrLf &
                      "					   isnull(sum(case" & vbCrLf &
                      "									 when isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id)= 'ICMS'" & vbCrLf &
                      "									   then NFxE.Valor" & vbCrLf &
                      "									   else 0" & vbCrLf &
                      "								   end),0) as ValorIcms," & vbCrLf &
                      "					   isnull(sum(case" & vbCrLf &
                      "									 when isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id) = 'IPI'" & vbCrLf &
                      "									   then NFxE.Valor" & vbCrLf &
                      "									   else 0" & vbCrLf &
                      "								   end),0) as ValorIPI," & vbCrLf &
                      "					   isnull(sum(case" & vbCrLf &
                      "									 when isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id) = 'PIS'" & vbCrLf &
                      "									   then NFxE.Valor" & vbCrLf &
                      "									   else 0" & vbCrLf &
                      "								   end),0) as ValorPIS," & vbCrLf &
                      "					   isnull(sum(case" & vbCrLf &
                      "									 when isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id) = 'COFINS'" & vbCrLf &
                      "									   then NFxE.Valor" & vbCrLf &
                      "									   else 0" & vbCrLf &
                      "								   end),0) as ValorCOFINS" & vbCrLf &
                      "				  FROM NotasFiscaisXEncargos AS NFxE" & vbCrLf &
                      "				  Left join Encargos AS Enc" & vbCrLf &
                      "				    on NFxE.Encargo_Id = Enc.Encargo_id" & vbCrLf &
                      "				   and isnull(Enc.EncargoAgrupador,'') <> ''" & vbCrLf &
                      "                 Where isnull(Enc.EncargoAgrupador,nfxe.Encargo_Id) in ('PRODUTO','PIS','COFINS','ICMS')" & vbCrLf &
                      "				 Group By NFxE.Empresa_Id, NFxE.EndEmpresa_Id, NFxE.Cliente_Id, NFxE.EndCliente_Id, NFxE.EntradaSaida_Id, NFxE.Nota_Id, NFxE.Serie_Id, NFxE.Produto_id, NFxE.Sequencia_id" & vbCrLf &
                      "             ) Enc" & vbCrLf &
                      "    ON nfi.Empresa_Id      = Enc.Empresa_Id" & vbCrLf &
                      "   AND nfi.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf &
                      "   AND nfi.Cliente_Id      = Enc.Cliente_Id" & vbCrLf &
                      "   AND nfi.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf &
                      "   AND nfi.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf &
                      "   AND nfi.Serie_Id        = Enc.Serie_Id" & vbCrLf &
                      "   AND nfi.Nota_Id         = Enc.Nota_Id" & vbCrLf &
                      "   AND nfi.Produto_id      = Enc.Produto_Id" & vbCrLf &
                      "   and nfi.Sequencia_Id    = Enc.Sequencia_id" & vbCrLf &
                      "  LEFT JOIN NFERealizadas" & vbCrLf &
                      "    ON NF.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf &
                      "   AND NF.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf &
                      "   AND NF.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf &
                      "   AND NF.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf &
                      "   AND NF.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf &
                      "   AND NF.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf &
                      "   AND NF.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf &
                      " WHERE NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                      "   AND NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                      "   AND NF.Movimento  BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                      "   AND NF.Operacao > 0" & vbCrLf &
                      "   AND NF.Situacao not in(9,10)" & vbCrLf &
                      "   AND OE.CodigoFiscal Between 1252 and 1253" & vbCrLf &
                      " Group by NF.EntradaSaida_Id," & vbCrLf &
                      "          NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
                      " 	     NF.Serie_Id, NF.Nota_Id, " & vbCrLf &
                      "          NF.NossaEmissao, ISNULL(NFERealizadas.ChaveNfe, ''), " & vbCrLf &
                      " 	     NF.DataDaNota, NF.Movimento, NF.TipoDeDocumento " & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            Nota = .Item("Nota_Id")
                            Serie = .Item("Serie_Id")
                            EntradaSaida = .Item("EntradaSaida_Id")
                            Cliente = .Item("Cliente_Id")
                            EndCliente = .Item("EndCliente_Id")
                            Emissao = .Item("NossaEmissao")

                            linha = "|C500"  '01 - Texto fixo contendo "C500" 
                            linha &= "|0"    '02 - Indicador do tipo de operação: 0 - Entrada; 1 - Saída
                            linha &= "|1"    '03 - Indicador do emitente do documento fiscal: 0 - Emissão própria; 1 - Terceiros

                            linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id") '04 - Código do participante (campo 02 do Registro 0150)

                            If .Item("TipoDeDocumento") = 16 Then
                                linha &= "|" & "06" '06 - Código do modelo do documento fiscal para Energia Elétrica - Furlan 16/03/2026
                            Else
                                linha &= "|" & .Item("TipoDeDocumento") '05 - Código do modelo do documento fiscal, conforme a Tabela 4.1.1 
                            End If

                            linha &= "|00"  '06 - Código da situação do documento fiscal, conforme a Tabela 4.1.2 

                            linha &= "|" & Trim(.Item("Serie_Id")) '07 - Série do documento fiscal
                            linha &= "|"                           '08 - Subsérie do documento fiscal

                            'Removido conforme orientação Jéssica - Furlan - 12/12/2025
                            'linha &= "|04"                        '09 -  Código de classe de consumo de energia elétrica ou gás
                            linha &= "|"                           '09 -  Código de classe de consumo de energia elétrica ou gás

                            linha &= "|" & .Item("Nota_Id")        '10 - Número do documento fiscal
                            linha &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy") '11 - Data da emissão do documento fiscal
                            linha &= "|" & Format(CDate(.Item("Movimento")), "ddMMyyyy")  '12 - Data da entrada ou da saída
                            linha &= "|" & .Item("ValorTotal")                            '13 - Valor total do documento fiscal
                            linha &= "|"                                                  '14 - Valor total do desconto
                            linha &= "|" & .Item("ValorTotal")                            '15 - Valor total fornecido/consumido

                            linha &= "|"                                                  '16 - Valor total dos serviços não-tributados pelo ICMS
                            linha &= "|"                                                  '17 - Valor total cobrado em nome de terceiros
                            linha &= "|"                                                  '18 - Valor total de despesas acessórias indicadas no documento fiscal

                            linha &= "|" & .Item("BaseIcms")                              '19 - Valor acumulado da base de cálculo do ICMS
                            linha &= "|" & .Item("ValorIcms")                             '20 - Valor acumulado do ICMS

                            linha &= "|"                                                  '21 - Valor da base de cálculo do ICMS substituição tributária
                            linha &= "|"                                                  '22 - Valor do ICMS retido por substituição tributária

                            linha &= "|"                                                  '23 - Código da informação complementar do documento fiscal (campo 02 do Registro 0450)
                            linha &= "|"                                                  '24 - Valor do Pis
                            linha &= "|"                                                  '25 - Valor da Cofins

                            'Removido conforme orientação Jéssica - Furlan - 12/12/2025
                            'linha &= "|3"                                                 '26 - Código de tipo de Ligação 1 - Monofásico 2 - Bifásico 3 - Trifásico
                            'linha &= "|12"                                                '27 - Código de Grupo de Tensão
                            linha &= "|"                                                   '26 - Código de tipo de Ligação 1 - Monofásico 2 - Bifásico 3 - Trifásico
                            linha &= "|"                                                   '27 - Código de Grupo de Tensão

                            If PeriodoFinal.Year > 2019 Then

                                If .Item("ChaveNfe").ToString.Length > 0 Then
                                    linha &= "|" & .Item("ChaveNfe")                      '28 - Chave da Nota Fiscal de Energia Elétrica Eletrônica
                                    linha &= "|1"                                         '29 - Finalidade da emissão do documento eletrônico: 1 Normal - 2 Substituição - 3 Normal com ajuste
                                Else
                                    linha &= "|"                                          '28 - Chave da Nota Fiscal de Energia Elétrica Eletrônica
                                    linha &= "|"                                          '29 - Finalidade da emissão do documento eletrônico: 1 Normal - 2 Substituição - 3 Normal com ajuste
                                End If

                                linha &= "|"                                              '30 - Chave da nota referenciada, substituída.
                                linha &= "|"                                              '31 - Indicador do Destinatário/Acessante: 1 Contribuinte do ICMS - 2 Contribuinte Isento de Inscrição no Cadastro de Contribuintes do ICMS - 9 Não Contribuinte
                                linha &= "|"                                              '32 - Código do município do destinatário conforme a tabela do IBGE
                                linha &= "|"                                              '33 - Código da conta analítica contábil debitada/creditada
                            End If

                            If PeriodoFinal.Year > 2021 Then
                                linha &= "|"                                              '34 - Código do modelo do documento fiscal referenciado, conforme a Tabela 4.1.1
                                linha &= "|"                                              '35 - Código de autenticação digital do registro (Convênio 115/2003).
                                linha &= "|"                                              '36 - Série do documento fiscal referenciado.
                                linha &= "|"                                              '37 - Número do documento fiscal referenciado. 
                                linha &= "|"                                              '38 - Mês e ano da emissão do documento fiscal referenciado.
                                linha &= "|"                                              '39 - Energia injetada
                                linha &= "|"                                              '40 - Outras deduções 
                            End If

                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        RegistroC500 += 1
                        RegistroGeral += 1
                        'Registro C510 - Notas Fiscais X Itens
                        Emissao = "A"
                        If Emissao <> "A" Then
                            Sql = "SELECT NF.EntradaSaida_Id," & vbCrLf & _
                                  "       NF.Cliente_Id," & vbCrLf & _
                                  "       NF.Serie_Id," & vbCrLf & _
                                  "	      NF.Nota_Id," & vbCrLf & _
                                  "		  'C510' AS Registro," & vbCrLf & _
                                  "       nfi.Produto_Id AS Item," & vbCrLf & _
                                  "		  prd.Nome," & vbCrLf & _
                                  "		  NF.NossaEmissao," & vbCrLf & _
                                  "       '' AS ChaveNfe," & vbCrLf & _
                                  "       NF.DataDaNota," & vbCrLf & _
                                  "       NF.Movimento," & vbCrLf & _
                                  "       nfi.QuantidadeFiscal," & vbCrLf & _
                                  "       prd.Unidade," & vbCrLf & _
                                  "       enc.CodigoFiscal as CFOP," & vbCrLf & _
                                  "       NF.Operacao, " & vbCrLf & _
                                  "       NF.SubOperacao," & vbCrLf & _
                                  "       enc.SituacaoTributaria," & vbCrLf & _
                                  "       enc.ValorTotal," & vbCrLf & _
                                  "       enc.BaseIcms," & vbCrLf & _
                                  "       enc.Aliquota," & vbCrLf & _
                                  "       enc.ValorIcms," & vbCrLf & _
                                  "       enc.CST_IPI," & vbCrLf & _
                                  "       enc.VL_BC_IPI," & vbCrLf & _
                                  "       enc.ALI_IPI," & vbCrLf & _
                                  "       enc.VL_IPI," & vbCrLf & _
                                  "       enc.CST_PIS," & vbCrLf & _
                                  "       enc.VL_BC_PIS," & vbCrLf & _
                                  "       enc.ALI_PIS," & vbCrLf & _
                                  "       enc.VL_PIS," & vbCrLf & _
                                  "       enc.CST_COFINS," & vbCrLf & _
                                  "       enc.VL_BC_COFINS," & vbCrLf & _
                                  "       enc.ALI_COFINS," & vbCrLf & _
                                  "       enc.VL_COFINS," & vbCrLf & _
                                  "       '' AS COD_CTA," & vbCrLf & _
                                  "       SO.EstoqueFiscal" & vbCrLf & _
                                  "  FROM NotasFiscais NF" & vbCrLf & _
                                  "  INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
                                  "     ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                                  "	   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                                  "	   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                                  "	   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                                  "	   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                                  "	   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                                  "	   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                                  "  INNER JOIN Produtos prd" & vbCrLf & _
                                  "     ON nfi.Produto_Id = prd.Produto_Id" & vbCrLf & _
                                  "  INNER JOIN SubOperacoes SO" & vbCrLf & _
                                  "     ON NF.Operacao    = SO.Operacao_Id " & vbCrLf & _
                                  "    AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                                  "  Inner join (Select nfxi.Empresa_Id," & vbCrLf & _
                                  "	                    nfxi.EndEmpresa_Id," & vbCrLf & _
                                  "	                    nfxi.Cliente_Id," & vbCrLf & _
                                  "	                    nfxi.EndCliente_Id," & vbCrLf & _
                                  "	                    nfxi.EntradaSaida_Id," & vbCrLf & _
                                  "	                    nfxi.Serie_Id," & vbCrLf & _
                                  "	                    nfxi.Nota_Id," & vbCrLf & _
                                  "					    nfxi.Produto_Id," & vbCrLf & _
                                  "					    nfxi.Sequencia_Id," & vbCrLf & _
                                  "					    OE.CodigoFiscal," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'PRODUTO'" & vbCrLf & _
                                  "					    	    then nfxe.valor" & vbCrLf & _
                                  "					    	    else 0" & vbCrLf & _
                                  "					        end) as ValorTotal," & vbCrLf & _
                                  "                     --**************************************" & vbCrLf & _
                                  "                     --*************** ICMS *****************" & vbCrLf & _
                                  "					    --**************************************" & vbCrLf & _
                                  "					    OE.STICMS as SituacaoTributaria," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'ICMS'" & vbCrLf & _
                                  "					   		    then nfxe.Base" & vbCrLf & _
                                  "					   		    else 0" & vbCrLf & _
                                  "					        end) as BaseIcms," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'ICMS'" & vbCrLf & _
                                  "					   		    then nfxe.Percentual" & vbCrLf & _
                                  "					   		    else 0" & vbCrLf & _
                                  "					        end) as Aliquota," & vbCrLf & _
                                  "                     sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'ICMS'" & vbCrLf & _
                                  "					   		    then nfxe.Valor" & vbCrLf & _
                                  "					   		    else 0" & vbCrLf & _
                                  "					        end) as ValorIcms," & vbCrLf & _
                                  "                     --**************************************" & vbCrLf & _
                                  "                     --*************** IPI  *****************" & vbCrLf & _
                                  "					    --**************************************" & vbCrLf & _
                                  "					    oe.STIPI as CST_IPI," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'IPI'" & vbCrLf & _
                                  "					    		then nfxe.Base" & vbCrLf & _
                                  "					    		else 0" & vbCrLf & _
                                  "					        end) as VL_BC_IPI," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'IPI'" & vbCrLf & _
                                  "					    		then nfxe.Percentual" & vbCrLf & _
                                  "					    		else 0" & vbCrLf & _
                                  "					        end) as ALI_IPI," & vbCrLf & _
                                  "                     sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'IPI'" & vbCrLf & _
                                  "							    then nfxe.Valor" & vbCrLf & _
                                  "							    else 0" & vbCrLf & _
                                  "					     end) as VL_IPI," & vbCrLf & _
                                  "                     --**************************************" & vbCrLf & _
                                  "                     --*************** PIS  *****************" & vbCrLf & _
                                  "					    --**************************************" & vbCrLf & _
                                  "					    OE.STPISCOFINS as CST_PIS," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'PIS'" & vbCrLf & _
                                  "					   		    then nfxe.Base" & vbCrLf & _
                                  "					   		    else 0" & vbCrLf & _
                                  "					        end) as VL_BC_PIS," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'PIS'" & vbCrLf & _
                                  "							    then nfxe.Percentual" & vbCrLf & _
                                  "							    else 0" & vbCrLf & _
                                  "					        end) as ALI_PIS," & vbCrLf & _
                                  "                     sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'PIS'" & vbCrLf & _
                                  "							    then nfxe.Valor" & vbCrLf & _
                                  "							    else 0" & vbCrLf & _
                                  "					        end) as VL_PIS," & vbCrLf & _
                                  "                     --*****************************************" & vbCrLf & _
                                  "                     --*************** COFINS  *****************" & vbCrLf & _
                                  "					    --*****************************************" & vbCrLf & _
                                  "					    OE.STPISCOFINS as CST_COFINS," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'COFINS'" & vbCrLf & _
                                  "					   		    then nfxe.Base" & vbCrLf & _
                                  "					   		    else 0" & vbCrLf & _
                                  "					        end) as VL_BC_COFINS," & vbCrLf & _
                                  "					    sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'COFINS'" & vbCrLf & _
                                  "							    then nfxe.Percentual" & vbCrLf & _
                                  "							    else 0" & vbCrLf & _
                                  "					        end) as ALI_COFINS," & vbCrLf & _
                                  "                     sum(Case" & vbCrLf & _
                                  "					          When isnull(e.EncargoAgrupador,nfxe.Encargo_id) = 'COFINS'" & vbCrLf & _
                                  "							    then nfxe.Valor" & vbCrLf & _
                                  "							    else 0" & vbCrLf & _
                                  "					        end) as VL_COFINS" & vbCrLf & _
                                  "                From NotasFiscaisXItens nfxi" & vbCrLf & _
                                  "			      INNER JOIN NotasFiscaisXEncargos nfxe" & vbCrLf & _
                                  "                  ON nfxi.Empresa_Id      = nfxe.Empresa_Id" & vbCrLf & _
                                  "                 AND nfxi.EndEmpresa_Id   = nfxe.EndEmpresa_Id" & vbCrLf & _
                                  "                 AND nfxi.Cliente_Id      = nfxe.Cliente_Id" & vbCrLf & _
                                  "                 AND nfxi.EndCliente_Id   = nfxe.EndCliente_Id" & vbCrLf & _
                                  "                 AND nfxi.EntradaSaida_Id = nfxe.EntradaSaida_Id" & vbCrLf & _
                                  "                 AND nfxi.Serie_Id        = nfxe.Serie_Id" & vbCrLf & _
                                  "                 AND nfxi.Nota_Id         = nfxe.Nota_Id" & vbCrLf & _
                                  "                 AND nfxi.Produto_Id      = nfxe.Produto_Id" & vbCrLf & _
                                  "                 AND nfxi.Sequencia_Id    = nfxe.Sequencia_Id" & vbCrLf & _
                                  "			      Inner Join OperacaoXEstado OE" & vbCrLf & _
                                  "			         on OE.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf & _
                                  "                Left join Encargos E" & vbCrLf & _
                                  "			         on e.Encargo_id = nfxe.Encargo_Id" & vbCrLf & _
                                  "				    AND ISNULL(E.EncargoAgrupador,'') <> ''" & vbCrLf & _
                                  "			      group by nfxi.Empresa_Id," & vbCrLf & _
                                  "	                       nfxi.EndEmpresa_Id," & vbCrLf & _
                                  "	                       nfxi.Cliente_Id," & vbCrLf & _
                                  "	                       nfxi.EndCliente_Id," & vbCrLf & _
                                  "	                       nfxi.EntradaSaida_Id," & vbCrLf & _
                                  "	                       nfxi.Serie_Id," & vbCrLf & _
                                  "	                       nfxi.Nota_Id," & vbCrLf & _
                                  "					       nfxi.Produto_Id," & vbCrLf & _
                                  "					       nfxi.Sequencia_Id," & vbCrLf & _
                                  "					       OE.CodigoFiscal," & vbCrLf & _
                                  "					       OE.STICMS," & vbCrLf & _
                                  "					       oe.STIPI," & vbCrLf & _
                                  "					       OE.STPISCOFINS" & vbCrLf & _
                                  "               )enc" & vbCrLf & _
                                  "     ON enc.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                                  "    AND enc.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                                  "    AND enc.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                                  "    AND enc.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                                  "    AND enc.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                                  "    AND enc.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                                  "    AND enc.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                                  "    AND enc.Produto_Id      = nfi.Produto_Id" & vbCrLf & _
                                  "    AND enc.Sequencia_Id    = nfi.Sequencia_Id" & vbCrLf & _
                                  "  WHERE NF.Empresa_Id      ='" & Empresa(0) & "'" & vbCrLf & _
                                  "    AND NF.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                                  "    AND NF.EndCliente_Id   = " & EndCliente & vbCrLf & _
                                  "    And NF.Nota_Id         = " & Nota & vbCrLf & _
                                  "    AND NF.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                                  "    And NF.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                                  "    AND NF.Operacao        > 0" & vbCrLf & _
                                  "    AND NF.Situacao not in(9,10)" & vbCrLf

                            ds1 = Banco.ConsultaDataSet(Sql, "Consulta")
                            IntSequencia = 0

                            If ds1.Tables(0).Rows.Count > 0 Then
                                For Each dr1 As DataRow In ds1.Tables(0).Rows
                                    With dr1
                                        linha = "|C510"
                                        IntSequencia += 1

                                        linha &= "|" & IntSequencia
                                        linha &= "|" & Trim(.Item("Item"))
                                        linha &= "|00"      'Código de Classificação do Item

                                        linha &= "|" & .Item("QuantidadeFiscal")
                                        linha &= "|" & Trim(.Item("Unidade"))
                                        linha &= "|" & .Item("ValorTotal")
                                        linha &= "|"                            'Descontos
                                        linha &= "|" & Funcoes.AlinharDireita(.Item("SituacaoTributaria"), 3, "0")
                                        linha &= "|" & .Item("CFOP")
                                        linha &= "|" & .Item("BaseIcms")
                                        linha &= "|" & .Item("Aliquota")
                                        linha &= "|" & .Item("ValorIcms")
                                        linha &= "|"                            'Valor Base de Icms Substituição Tributária
                                        linha &= "|"                            'Aliquota de Icms por Substituição Tributária
                                        linha &= "|"                            'Valor de Icms por Substituição Tributária

                                        linha &= "|0"                           'Indicador do tipo de Receita
                                        linha &= "|"                            'Codigo Do participante Receptor da Receita de Terceiros

                                        linha &= "|" & .Item("VL_PIS")
                                        linha &= "|" & .Item("VL_COFINS")
                                        linha &= "|111010001"
                                        linha &= "|"
                                    End With

                                    strm.WriteLine(linha)
                                    RegistroC510 += 1
                                    RegistroGeral += 1

                                Next
                            End If
                        End If

                        'Registro C590 - Notas Fiscais X Itens
                        'Sql = "Select STICMS          AS CST_ICMS," & vbCrLf & _
                        '      "       CodigoFiscal    AS CFOP," & vbCrLf & _
                        '      "       Percentual      AS ALIQ_ICMS," & vbCrLf & _
                        '      "       Sum(Produto)    AS VL_OPR," & vbCrLf & _
                        '      "       Sum(Base)       AS VL_BC_ICMS," & vbCrLf & _
                        '      "       Sum(Icms)       AS VL_ICMS," & vbCrLf & _
                        '      "       Sum(Ipi)        AS VL_IPI" & vbCrLf & _
                        '      "  FROM (" & vbCrLf & _
                        '      "        Select nfi.Produto_Id," & vbCrLf & _
                        '      "		          oe.STICMS," & vbCrLf & _
                        '      "               oe.CodigoFiscal," & vbCrLf & _
                        '      "               (case" & vbCrLf & _
                        '      "			         when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                        '      "					   then nfe.Percentual" & vbCrLf & _
                        '      "					   else 0" & vbCrLf & _
                        '      "				   end) Percentual," & vbCrLf & _
                        '      "               sum(case" & vbCrLf & _
                        '      "			         when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                        '      "					   then nfe.valor" & vbCrLf & _
                        '      "					   else 0" & vbCrLf & _
                        '      "				   end) Produto," & vbCrLf & _
                        '      "               sum(case" & vbCrLf & _
                        '      "			         when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'IPI'" & vbCrLf & _
                        '      "					   then nfe.valor" & vbCrLf & _
                        '      "					   else 0" & vbCrLf & _
                        '      "				   end) IPI," & vbCrLf & _
                        '      "               sum(case" & vbCrLf & _
                        '      "			         when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                        '      "					   then nfe.Base" & vbCrLf & _
                        '      "					   else 0" & vbCrLf & _
                        '      "				   end) Base," & vbCrLf & _
                        '      "               sum(case" & vbCrLf & _
                        '      "			         when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                        '      "					   then nfe.valor" & vbCrLf & _
                        '      "					   else 0" & vbCrLf & _
                        '      "				   end) icms" & vbCrLf & _
                        '      "          FROM NotasFiscaisXEncargos nfe" & vbCrLf & _
                        '      "         Inner Join NotasFiscaisXItens nfi" & vbCrLf & _
                        '      "            on nfe.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                        '      "           and nfe.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                        '      "           and nfe.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                        '      "           and nfe.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                        '      "           and nfe.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                        '      "           and nfe.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                        '      "           and nfe.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                        '      "         Inner join OperacaoXEstado OE" & vbCrLf & _
                        '      "            on OE.Codigo_id = nfi.OperacaoXEstado" & vbCrLf & _
                        '      "         Left Join Encargos e" & vbCrLf & _
                        '      "		   on e.Encargo_id = nfe.Encargo_Id" & vbCrLf & _
                        '      "          and isnull(e.encargoagrupador,'') <> '' " & vbCrLf & _
                        '      "         WHERE nfe.Empresa_Id      ='" & Empresa(0) & "'" & vbCrLf & _
                        '      "           AND nfe.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        '      "           AND nfe.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        '      "           And nfe.Nota_Id         = " & Nota & vbCrLf & _
                        '      "           AND nfe.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        '      "           And nfe.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        '      "           and isnull(e.encargoagrupador, nfe.Encargo_Id) in ('PRODUTO','IPI','ICMS')" & vbCrLf & _
                        '      "         Group by nfi.Produto_Id, oe.STICMS, oe.CodigoFiscal," & vbCrLf & _
                        '      "                  case" & vbCrLf & _
                        '      "	    		       when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                        '      "	    			     then nfe.Percentual" & vbCrLf & _
                        '      "	    				 else 0" & vbCrLf & _
                        '      "	    			 end) as Consulta" & vbCrLf & _
                        '      "   Group by  STICMS, CodigoFiscal, Percentual" & vbCrLf


                        '----------------
                        'AJUSTADO SQL POIS ESTA VINDO 2 LINHAS QUANDO TEM ICMS - FURLAN - 16/02/2022
                        '----------------
                        Sql = "Select Consulta.STICMS AS CST_ICMS, Consulta.CodigoFiscal AS CFOP, isnull(ConsultaIcms.Percentual,0) AS ALIQ_ICMS, Consulta.Produto AS VL_OPR, " & vbCrLf & _
                        "		isnull(ConsultaIcms.BaseIcms,0) AS VL_BC_ICMS, isnull(ConsultaIcms.VlrIcms,0) AS VL_ICMS, isnull(ConsultaIpi.VlrIpi,0) AS VL_IPI " & vbCrLf & _
                        "From NotasFiscaisXEncargos ne " & vbCrLf & _
                        "	left join (Select nfe.Empresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.Nota_Id, nfe.Serie_Id, nfe.EntradaSaida_Id, " & vbCrLf & _
                        "						nfe.Produto_Id, oe.STICMS, oe.CodigoFiscal, " & vbCrLf & _
                        "					sum(nfe.valor) as Produto " & vbCrLf & _
                        "				FROM NotasFiscaisXEncargos nfe " & vbCrLf & _
                        "					Inner Join NotasFiscaisXItens nfi " & vbCrLf & _
                        "							on nfe.Empresa_Id      = nfi.Empresa_Id " & vbCrLf & _
                        "							and nfe.EndEmpresa_Id   = nfi.EndEmpresa_Id " & vbCrLf & _
                        "							and nfe.Cliente_Id      = nfi.Cliente_Id " & vbCrLf & _
                        "							and nfe.EndCliente_Id   = nfi.EndCliente_Id " & vbCrLf & _
                        "							and nfe.EntradaSaida_Id = nfi.EntradaSaida_Id " & vbCrLf & _
                        "							and nfe.Nota_Id         = nfi.Nota_Id " & vbCrLf & _
                        "							and nfe.Serie_Id        = nfi.Serie_Id " & vbCrLf & _
                        "					Inner join OperacaoXEstado OE " & vbCrLf & _
                        "							on OE.Codigo_id = nfi.OperacaoXEstado " & vbCrLf & _
                        "					Left Join Encargos e " & vbCrLf & _
                        "							on e.Encargo_id = nfe.Encargo_Id " & vbCrLf & _
                        "					and isnull(e.encargoagrupador,'') <> '' " & vbCrLf & _
                        "				WHERE nfe.Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                        "				AND nfe.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "				AND nfe.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "				And nfe.Nota_Id         = " & Nota & vbCrLf & _
                        "				AND nfe.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "				And nfe.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "				and nfe.Encargo_Id in ('PRODUTO')" & vbCrLf & _
                        "				GROUP BY nfe.Empresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.Nota_Id, nfe.Serie_Id, nfe.EntradaSaida_Id, nfe.Produto_Id, oe.STICMS, oe.CodigoFiscal) Consulta" & vbCrLf & _
                        "			ON Consulta.Empresa_Id       ='" & Empresa(0) & "'" & vbCrLf & _
                        "			AND Consulta.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "			AND Consulta.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "			And Consulta.Nota_Id         = " & Nota & vbCrLf & _
                        "			AND Consulta.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "			And Consulta.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "	left join (Select nfe.Empresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.Nota_Id, nfe.Serie_Id, nfe.EntradaSaida_Id, nfe.Produto_Id," & vbCrLf & _
                        "						nfe.Percentual, sum(nfe.Base) as BaseIcms, sum(nfe.valor) as VlrIcms" & vbCrLf & _
                        "				FROM NotasFiscaisXEncargos nfe" & vbCrLf & _
                        "					Inner Join NotasFiscaisXItens nfi" & vbCrLf & _
                        "							on nfe.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                        "							and nfe.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                        "							and nfe.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                        "							and nfe.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                        "							and nfe.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                        "							and nfe.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                        "							and nfe.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                        "					Inner join OperacaoXEstado OE" & vbCrLf & _
                        "							on OE.Codigo_id = nfi.OperacaoXEstado" & vbCrLf & _
                        "					Left Join Encargos e" & vbCrLf & _
                        "							on e.Encargo_id = nfe.Encargo_Id" & vbCrLf & _
                        "					and isnull(e.encargoagrupador,'') <> '' " & vbCrLf & _
                        "				WHERE nfe.Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                        "				AND nfe.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "				AND nfe.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "				And nfe.Nota_Id         = " & Nota & vbCrLf & _
                        "				AND nfe.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "				And nfe.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "				and nfe.Encargo_Id in ('ICMS')" & vbCrLf & _
                        "				GROUP BY nfe.Empresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.Nota_Id, nfe.Serie_Id, nfe.EntradaSaida_Id, nfe.Produto_Id, nfe.Percentual) ConsultaIcms" & vbCrLf & _
                        "			ON ConsultaIcms.Empresa_Id       ='" & Empresa(0) & "'" & vbCrLf & _
                        "			AND ConsultaIcms.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "			AND ConsultaIcms.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "			And ConsultaIcms.Nota_Id         = " & Nota & vbCrLf & _
                        "			AND ConsultaIcms.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "			And ConsultaIcms.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "	left join (Select nfe.Empresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.Nota_Id, nfe.Serie_Id, nfe.EntradaSaida_Id, nfe.Produto_Id," & vbCrLf & _
                        "						sum(nfe.valor) as VlrIpi" & vbCrLf & _
                        "				FROM NotasFiscaisXEncargos nfe" & vbCrLf & _
                        "					Inner Join NotasFiscaisXItens nfi" & vbCrLf & _
                        "							on nfe.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                        "							and nfe.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                        "							and nfe.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                        "							and nfe.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                        "							and nfe.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                        "							and nfe.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                        "							and nfe.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                        "					Inner join OperacaoXEstado OE" & vbCrLf & _
                        "							on OE.Codigo_id = nfi.OperacaoXEstado" & vbCrLf & _
                        "					Left Join Encargos e" & vbCrLf & _
                        "							on e.Encargo_id = nfe.Encargo_Id" & vbCrLf & _
                        "					and isnull(e.encargoagrupador,'') <> '' " & vbCrLf & _
                        "				WHERE nfe.Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                        "				AND nfe.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "				AND nfe.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "				And nfe.Nota_Id         = " & Nota & vbCrLf & _
                        "				AND nfe.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "				And nfe.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "				and nfe.Encargo_Id in ('IPI')" & vbCrLf & _
                        "				GROUP BY nfe.Empresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.Nota_Id, nfe.Serie_Id, nfe.EntradaSaida_Id, nfe.Produto_Id) ConsultaIpi" & vbCrLf & _
                        "			ON ConsultaIcms.Empresa_Id       ='" & Empresa(0) & "'" & vbCrLf & _
                        "			AND ConsultaIcms.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "			AND ConsultaIcms.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "			And ConsultaIcms.Nota_Id         = " & Nota & vbCrLf & _
                        "			AND ConsultaIcms.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "			And ConsultaIcms.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "WHERE ne.Empresa_Id      ='" & Empresa(0) & "'" & vbCrLf & _
                        "  AND ne.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                        "  AND ne.EndCliente_Id   = " & EndCliente & vbCrLf & _
                        "  And ne.Nota_Id         = " & Nota & vbCrLf & _
                        "  AND ne.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                        "  And ne.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                        "  AND ne.Encargo_Id in ('PRODUTO')"

                        ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                        If ds2.Tables(0).Rows.Count > 0 Then
                            For Each dr2 As DataRow In ds2.Tables(0).Rows
                                With dr2
                                    linha = "|C590"
                                    linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                                    linha &= "|" & .Item("CFOP")
                                    linha &= "|" & String.Format("{0:N2}", .Item("ALIQ_ICMS"))
                                    linha &= "|" & .Item("VL_OPR")
                                    linha &= "|" & .Item("VL_BC_ICMS")
                                    linha &= "|" & .Item("VL_ICMS")
                                    linha &= "|0" '& .Item("VL_BC_ICMS_ST")
                                    linha &= "|0" '& .Item("VL_ICMS_ST")

                                    If .Item("VL_BC_ICMS") <> 0 And (.Item("VL_OPR") > .Item("VL_BC_ICMS")) Then
                                        linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                                    Else
                                        linha &= "|0" '& .Item("VL_RED_BC")
                                    End If

                                    linha &= "|" ' Código de Observacao
                                    linha &= "|"
                                End With

                                strm.WriteLine(linha)
                                RegistroC590 += 1
                                RegistroGeral += 1
                            Next
                        End If
                    Next
                End If


                ' Registro C990  - Encerramento do Bloco C

                RegistroC990 += 1

                linha = "|C990"
                linha &= "|" & RegistroC001 + RegistroC100 + RegistroC101 + RegistroC170 + RegistroC172 + RegistroC190 + RegistroC500 + RegistroC510 + RegistroC590 + RegistroC990 'RegistroC195 + RegistroC197 + RegistroC500 + RegistroC510 + RegistroC590 + RegistroC990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1


                ' Registro D001  - Abertura do Bloco D
                linha = "|D001"
                linha &= "|" & BlocoD       'Indicador de Movimento - 0 - Bloco com Dados Informados
                '                                                     1 - Bloco sem dados informados
                linha &= "|"

                strm.WriteLine(linha)
                RegistroD001 += 1
                RegistroGeral += 1

                'Registro D100 - Nota Fiscal de Serviço de Transporte
                ds = ConsultaRegistroD100()

                Nota = String.Empty
                Serie = String.Empty
                EntradaSaida = String.Empty
                Cliente = String.Empty
                EndCliente = String.Empty
                Emissao = String.Empty

                Dim sqlWhereNotaD100 As String = ""
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr

                            sqlWhereNotaD100 &= IIf(sqlWhereNotaD100.Length = 0, "", " Or " & vbCrLf)
                            sqlWhereNotaD100 &= "("
                            sqlWhereNotaD100 &= "      Cliente_Id      = '" & .Item("Cliente_Id") & "'" & vbCrLf
                            sqlWhereNotaD100 &= "  AND EndCliente_Id   =  " & .Item("EndCliente_Id") & vbCrLf
                            sqlWhereNotaD100 &= "  AND Nota_id         =  " & .Item("Nota_Id") & vbCrLf
                            sqlWhereNotaD100 &= "  AND Serie_Id        = '" & .Item("Serie_Id") & "'" & vbCrLf
                            sqlWhereNotaD100 &= "  AND EntradaSaida_Id = '" & .Item("EntradaSaida_Id") & "'" & vbCrLf
                            sqlWhereNotaD100 &= ")"

                            'Nota &= .Item("Nota_Id") & ","
                            'Serie &= "'" & .Item("Serie_Id") & "'" & ","
                            'EntradaSaida &= "'" & .Item("EntradaSaida_Id") & "'" & ","
                            'Cliente &= "'" & .Item("Cliente_Id") & "'" & ","
                            'EndCliente &= .Item("EndCliente_Id") & ","
                            'Emissao &= "'" & .Item("NossaEmissao") & "'" & ","
                        End With
                    Next
                    'Nota = Nota.Substring(0, Len(Nota) - 1)
                    'Serie = Serie.Substring(0, Len(Serie) - 1)
                    'EntradaSaida = EntradaSaida.Substring(0, Len(EntradaSaida) - 1)
                    'Cliente = Cliente.Substring(0, Len(Cliente) - 1)
                    'EndCliente = EndCliente.Substring(0, Len(EndCliente) - 1)
                    'Emissao = Nota.Substring(0, Len(Emissao) - 1)

                    Dim D100NF As Integer = 0
                    Dim D100Serie As String = ""
                    Dim D100EntradaSaida As String = ""
                    Dim D100Cliente As String = ""
                    Dim D100EndCliente As Integer = 0

                    ds1 = ConsultaRegistroD190(sqlWhereNotaD100)

                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|D100"
                            If .Item("Cfop") < 5000 Then
                                linha &= "|0" 'Aquisicao
                            Else
                                linha &= "|1" 'Prestacao
                            End If

                            If .Item("NossaEmissao") = "S" Then
                                linha &= "|0" 'Emissao Propria
                            Else
                                linha &= "|1" 'Terceiros
                            End If

                            If .Item("Situacao") = "2" Then
                                linha &= "|"
                            Else
                                linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                            End If

                            'If .Item("Cfop") < 5000 Then
                            '    linha &= "|08"
                            'Else
                            '    linha &= "|57"                      'Código do modelo do documento fiscal, conforme a Tabela 4.1.1
                            'End If

                            'Ajustei para pegar o tipo de Documento
                            'If .Item("ChaveNFE").Equals("") Then
                            '    linha &= "|08"
                            'Else
                            '    linha &= "|57"                      'Código do modelo do documento fiscal, conforme a Tabela 4.1.1
                            'End If

                            If .Item("TipoDeDocumento") = 2 AndAlso .Item("ChaveNFE").Equals("") Then
                                linha &= "|08"
                            ElseIf (.Item("TipoDeDocumento") = 1 Or .Item("TipoDeDocumento") = 2 Or .Item("TipoDeDocumento") = 58) AndAlso .Item("ChaveNFE").ToString.Length > 0 Then
                                linha &= "|57"
                            Else
                                linha &= "|" & CInt(.Item("TipoDeDocumento")).ToString("00")
                            End If

                            If .Item("Situacao") = 9 Then       '05 NF-e ou CT-e - Numeração inutilizada, 
                                linha &= "|05"                  '06 Documento Fiscal Complementar, 07 Escrituração extemporânea de documento complementar
                            ElseIf .Item("Situacao") = 10 Then  '04 NF-e ou CT-e - denegado
                                linha &= "|04"
                            ElseIf .Item("Situacao") = 2 Then   '02 Documento cancelado, 03 Escrituração extemporânea de documento cancelado
                                linha &= "|02"                  '08 Documento Fiscal emitido com base em Regime Especial ou Norma Específica
                            Else                                '00 Documento regular, 01 Escrituração extemporânea de documento regular
                                linha &= "|00"
                            End If

                            linha &= "|" & Trim(.Item("Serie_Id"))
                            linha &= "|"                        'SubSérie do Documento Fiscal
                            linha &= "|" & .Item("Nota_Id")

                            'If .Item("NossaEmissao") = "S" And .Item("Eletronica") = "S" Then
                            '    linha &= "|" & .Item("ChaveNfe")    'Chave do Conhecimento de Transporte eletronico
                            'Else
                            '    linha &= "|"
                            'End If

                            If .Item("ChaveNFE").Equals("") Then
                                linha &= "|"
                            Else
                                linha &= "|" & .Item("ChaveNfe")    'Chave do Conhecimento de Transporte eletronico
                            End If

                            If .Item("Situacao") = "2" Then
                                linha &= "|"
                                linha &= "|"
                                linha &= "|"
                            Else
                                linha &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy")
                                linha &= "|" & Format(CDate(.Item("Movimento")), "ddMMyyyy")
                                linha &= "|0"                       'Tipo de CT-e quando o modelo do documento for 57
                            End If

                            linha &= "|"                        'Chave do CT-e Informar Campo vazio

                            If .Item("Situacao") = "2" Then
                                linha &= "|"
                                linha &= "|"
                                linha &= "|"
                                linha &= "|"
                                linha &= "|"
                                linha &= "|"
                            Else
                                linha &= "|" & .Item("ValorTotal")
                                linha &= "|"                        'Valor Total do Desconto
                                linha &= "|1"                       'Indicador do Tipo de Frete 0 - Por conta de Terceiros
                                '                                                               1 - Por Conta do Emitente
                                '                                                               2 - Por Conta do Destinatário
                                '                                                               9 - Sem Frete
                                linha &= "|" & .Item("ValorTotal")
                                linha &= "|" & .Item("BaseIcms")
                                linha &= "|" & .Item("ValorIcms")
                            End If

                            linha &= "|"                        'Valor não Tributado
                            linha &= "|"                        'Código da informação complementar do documento fiscal - campo02 da tabela 450

                            If .Item("Situacao") = "2" Then
                                linha &= "|"
                            Else
                                linha &= "|111010001"           'Codigo da Conta Analitica Contabil Debitada/Creditada
                            End If

                            If .Item("Situacao") = "2" Then
                                linha &= "|"
                                linha &= "|"
                            Else
                                If .Item("EntradaSaida_Id") = "E" Then
                                    linha &= "|" & .Item("CODMUNCLI")
                                    linha &= "|" & .Item("CODMUNEMP")
                                Else
                                    linha &= "|" & .Item("CODMUNEMP")
                                    linha &= "|" & .Item("CODMUNCLI")
                                End If
                            End If

                            linha &= "|"

                            strm.WriteLine(linha)
                            RegistroD100 += 1
                            RegistroGeral += 1

                            D100NF = .Item("Nota_Id")
                            D100Serie = .Item("Serie_Id")
                            D100EntradaSaida = .Item("EntradaSaida_Id")
                            D100Cliente = .Item("Cliente_Id")
                            D100EndCliente = .Item("EndCliente_Id")

                            'Registro D190 - Registro Analitico dos Documentos
                            If ds1.Tables(0).Rows.Count > 0 AndAlso Not .Item("Situacao") = "2" Then
                                CompoeRegistroD190(ds1, strm, linha, RegistroD190, RegistroGeral, D100NF, D100Cliente, D100EndCliente, D100Serie, D100EntradaSaida)
                            End If
                        End With
                    Next
                End If

                'Registro D500 - Nota Fiscal de Serviço de Comunicação
                Sql = "SELECT NF.EntradaSaida_Id," & vbCrLf & _
                      "       NF.Cliente_Id, NF.EndCliente_Id," & vbCrLf & _
                      "       NF.Serie_Id, NF.Nota_Id," & vbCrLf & _
                      "       ISNULL(NFERealizadas.ChaveNfe, '') AS ChaveNfe," & vbCrLf & _
                      "	      NF.DataDaNota, NF.Movimento," & vbCrLf & _
                      "       sum(enc.ValorTotal) as ValorTotal," & vbCrLf & _
                      "	      sum(enc.BaseIcms) as BaseIcms," & vbCrLf & _
                      "	      0 AS Aliquota," & vbCrLf & _
                      "       sum(enc.ValorIcms) as ValorIcms" & vbCrLf & _
                      "  FROM NotasFiscais NF" & vbCrLf & _
                      " Inner Join NotasFiscaisXItens nfi" & vbCrLf & _
                      "    on NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                      "   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                      "   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                      "   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                      "   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                      "   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                      "   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                      "   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                      " Inner Join OperacaoXEstado OE" & vbCrLf & _
                      "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf & _
                      " Inner Join ( SELECT nfe.Empresa_id, nfe.EndEmpresa_id," & vbCrLf & _
                      "                     nfe.Cliente_Id, nfe.EndCliente_Id," & vbCrLf & _
                      "	                    nfe.Nota_Id, nfe.Serie_Id,  nfe.EntradaSaida_Id," & vbCrLf & _
                      "				        nfe.Produto_Id, nfe.Sequencia_id," & vbCrLf & _
                      "				       sum(case" & vbCrLf & _
                      "				             when isnull(e.encargoagrupador,nfe.encargo_id) = 'PRODUTO'" & vbCrLf & _
                      "				     		   then nfe.Valor" & vbCrLf & _
                      "				     		   else 0" & vbCrLf & _
                      "				     	   end) as ValorTotal," & vbCrLf & _
                      "                     sum(case" & vbCrLf & _
                      "				             when isnull(e.encargoagrupador,nfe.encargo_id) = 'ICMS'" & vbCrLf & _
                      "						       then nfe.Base" & vbCrLf & _
                      "						       else 0" & vbCrLf & _
                      "					        end) as BaseIcms," & vbCrLf & _
                      "                     sum(case" & vbCrLf & _
                      "				             when isnull(e.encargoagrupador,nfe.encargo_id) = 'ICMS'" & vbCrLf & _
                      "						       then nfe.Valor" & vbCrLf & _
                      "						       else 0" & vbCrLf & _
                      "					        end) as ValorIcms" & vbCrLf & _
                      "                from NotasFiscaisXEncargos nfe" & vbCrLf & _
                      "                Left Join Encargos e" & vbCrLf & _
                      "			         on e.Encargo_id = nfe.Encargo_Id" & vbCrLf & _
                      "                 and isnull(e.encargoagrupador,'') <> ''" & vbCrLf & _
                      "               Where isnull(e.encargoagrupador,nfe.Encargo_Id) in ('PRODUTO','ICMS')" & vbCrLf & _
                      "               group by nfe.Empresa_id, nfe.EndEmpresa_id," & vbCrLf & _
                      "                        nfe.Cliente_Id, nfe.EndCliente_Id," & vbCrLf & _
                      "	                       nfe.Nota_Id, nfe.Serie_Id,  nfe.EntradaSaida_Id," & vbCrLf & _
                      "				           nfe.Produto_Id, nfe.Sequencia_id" & vbCrLf & _
                      "            ) enc" & vbCrLf & _
                      "    on enc.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                      "   AND enc.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                      "   AND enc.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                      "   AND enc.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                      "   AND enc.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                      "   AND enc.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                      "   AND enc.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                      "   AND enc.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                      "   And enc.Produto_id      = nfi.Produto_Id" & vbCrLf & _
                      "   And enc.Sequencia_id    = nfi.Sequencia_Id" & vbCrLf & _
                      "  LEFT JOIN NFERealizadas" & vbCrLf & _
                      "    ON NF.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                      "   AND NF.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                      "   AND NF.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                      "   AND NF.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                      "   AND NF.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                      "   AND NF.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                      "   AND NF.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                      " WHERE NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf & _
                      "   and NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf & _
                      "   AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                      "   And NF.Situacao not in(9,10)" & vbCrLf & _
                      "   and NF.operacao > 0" & vbCrLf & _
                      "   and ((OE.CodigoFiscal Between 1300 and 1350) Or" & vbCrLf & _
                      "        (OE.CodigoFiscal Between 2300 and 2350) Or" & vbCrLf & _
                      "        (OE.CodigoFiscal IN (1932,2932)) Or" & vbCrLf & _
                      "        (OE.CodigoFiscal Between 5300 and 5350) Or" & vbCrLf & _
                      "	       (OE.CodigoFiscal Between 6300 and 6350))" & vbCrLf & _
                      "group by NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf & _
                      "         ISNULL(NFERealizadas.ChaveNfe, ''), NF.DataDaNota, NF.Movimento "


                ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds1.Tables(0).Rows.Count > 0 Then
                    For Each dr1 As DataRow In ds1.Tables(0).Rows
                        With dr1
                            Nota = .Item("Nota_Id")
                            Serie = .Item("Serie_Id")
                            EntradaSaida = .Item("EntradaSaida_Id")
                            Cliente = .Item("Cliente_Id")
                            EndCliente = .Item("EndCliente_Id")

                            linha = "|D500"
                            linha &= "|0"
                            linha &= "|1"
                            linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                            linha &= "|22"
                            linha &= "|00"                      '00 - Documento Regular 02 - Cancelado
                            linha &= "|" & Trim(.Item("Serie_Id"))
                            linha &= "|"                        'SubSérie do Documento Fiscal
                            linha &= "|" & .Item("Nota_Id")
                            linha &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy")
                            linha &= "|" & Format(CDate(.Item("Movimento")), "ddMMyyyy")

                            linha &= "|" & .Item("ValorTotal")
                            linha &= "|"                        'Valor Total do Desconto
                            linha &= "|" & .Item("ValorTotal")

                            linha &= "|"                        'Valor Total dos Servicos nao Tributados Pelo ICMS
                            linha &= "|"                        'Valor Cobrado em Nome de Terceiros
                            linha &= "|"                        'Valor de Outras Despesas Indicadas no Documento Fiscal

                            linha &= "|" & .Item("BaseIcms")
                            linha &= "|" & .Item("ValorIcms")
                            linha &= "|"                        'Codigo da Informacao Complementar
                            linha &= "|"                        'Valor do Pis
                            linha &= "|"                        'Valor da Cofins
                            linha &= "|111010001"               'Codigo da Conta Analitica Contabil Debitada/Creditada
                            linha &= "|1"                       'Tipo de Assinante 1 - Comercial/Industrial
                            linha &= "|"

                        End With
                        strm.WriteLine(linha)
                        RegistroD500 += 1
                        RegistroGeral += 1

                        'Registro D590 - Registro Analitico dos Documentos
                        'Sql = "Select oe.STICMS       as CST_ICMS," & vbCrLf & _
                        '      "		  OE.CodigoFiscal as CFOP," & vbCrLf & _
                        '      "		  enc.Produto     as VL_OPR," & vbCrLf & _
                        '      "		  enc.Base        as VL_BC_ICMS," & vbCrLf & _
                        '      "		  enc.Percentual  as ALIQ_ICMS," & vbCrLf & _
                        '      "		  enc.Icms        as VL_ICMS," & vbCrLf & _
                        '      "		  enc.Ipi         as VL_IPI" & vbCrLf & _
                        Sql = "Select oe.STICMS       as CST_ICMS," & vbCrLf & _
                              "		  OE.CodigoFiscal as CFOP," & vbCrLf & _
                              "		  sum(enc.Produto) as VL_OPR," & vbCrLf & _
                              "		  sum(enc.Base)    as VL_BC_ICMS," & vbCrLf & _
                              "		  enc.Percentual   as ALIQ_ICMS," & vbCrLf & _
                              "		  sum(enc.Icms)    as VL_ICMS," & vbCrLf & _
                              "		  sum(enc.Ipi)     as VL_IPI" & vbCrLf & _
                              "  FROM NotasFiscaisXItens nfi" & vbCrLf & _
                              " inner join OperacaoXEstado OE" & vbCrLf & _
                              "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf & _
                              "  Inner Join ( Select nfe.Empresa_Id, nfe.EndEmpresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.EntradaSaida_Id, nfe.Serie_Id, nfe.Nota_Id, nfe.Produto_Id, nfe.Sequencia_id," & vbCrLf & _
                              "		                 (Case" & vbCrLf & _
                              "				        	when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                              "				        	 then nfe.Percentual" & vbCrLf & _
                              "				        	 else 0" & vbCrLf & _
                              "                      end) Percentual," & vbCrLf & _
                              "		                 sum(Case" & vbCrLf & _
                              "				        	   when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                              "				        	     then nfe.Valor" & vbCrLf & _
                              "				        	     else 0" & vbCrLf & _
                              "                          end) Produto," & vbCrLf & _
                              "		                 sum(Case" & vbCrLf & _
                              "				        	   when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'IPI'" & vbCrLf & _
                              "				        	     then nfe.Valor" & vbCrLf & _
                              "				        	     else 0" & vbCrLf & _
                              "                          end) IPI," & vbCrLf & _
                              "		                 sum(Case" & vbCrLf & _
                              "				        	   when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                              "				        	     then nfe.Base" & vbCrLf & _
                              "				        	     else 0" & vbCrLf & _
                              "                          end) Base," & vbCrLf & _
                              "		                 sum(Case" & vbCrLf & _
                              "				        	   when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                              "				        	     then nfe.Valor" & vbCrLf & _
                              "				        	     else 0" & vbCrLf & _
                              "                          end) Icms" & vbCrLf & _
                              "		            from NotasFiscaisXEncargos nfe" & vbCrLf & _
                              "                 Left Join Encargos e" & vbCrLf & _
                              "					  on e.Encargo_id = nfe.Encargo_Id" & vbCrLf & _
                              "                  and isnull(e.encargoagrupador,'')  <> ''" & vbCrLf & _
                              "                Where isnull(e.encargoagrupador,nfe.Encargo_Id) in ('PRODUTO','IPI','ICMS')" & vbCrLf & _
                              "                Group by nfe.Empresa_Id, nfe.EndEmpresa_Id, nfe.Cliente_Id, nfe.EndCliente_Id, nfe.EntradaSaida_Id, nfe.Serie_Id, nfe.Nota_Id, nfe.Produto_Id, nfe.Sequencia_id," & vbCrLf & _
                              "		                 (Case" & vbCrLf & _
                              "				        	when isnull(e.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
                              "				        	 then nfe.Percentual" & vbCrLf & _
                              "				        	 else 0" & vbCrLf & _
                              "                      end)" & vbCrLf & _
                              "			     ) enc" & vbCrLf & _
                              "     on enc.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                              "	   and enc.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                              "	   and enc.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
                              "	   and enc.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
                              "	   and enc.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
                              "	   and enc.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
                              "	   and enc.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
                              "	   and enc.Produto_Id      = nfi.Produto_Id" & vbCrLf & _
                              "	   and enc.Sequencia_id    = nfi.Sequencia_id" & vbCrLf & _
                              "  Where nfi.Empresa_Id      ='" & Empresa(0) & "'" & vbCrLf & _
                              "	   AND nfi.Cliente_Id      ='" & Cliente & "'" & vbCrLf & _
                              "	   AND nfi.EndCliente_Id   = " & EndCliente & vbCrLf & _
                              "	   And nfi.Nota_Id         = " & Nota & vbCrLf & _
                              "	   AND nfi.Serie_Id        ='" & Serie & "'" & vbCrLf & _
                              "	   And nfi.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf & _
                              "group by oe.STICMS, OE.CodigoFiscal, enc.Percentual "

                        ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                        If ds2.Tables(0).Rows.Count > 0 Then
                            For Each dr2 As DataRow In ds2.Tables(0).Rows
                                With dr2
                                    linha = "|D590"
                                    linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                                    linha &= "|" & .Item("CFOP")
                                    linha &= "|" & String.Format("{0:N2}", .Item("ALIQ_ICMS"))
                                    linha &= "|" & .Item("VL_OPR")
                                    linha &= "|" & .Item("VL_BC_ICMS")
                                    linha &= "|" & .Item("VL_ICMS")
                                    linha &= "|0"                        ' Base de Calculo ICms Substituição Tributária
                                    linha &= "|0"                        ' Valor Icms Substituição Tributária

                                    If .Item("VL_BC_ICMS") <> 0 Then
                                        linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                                    Else
                                        linha &= "|0"        '& .Item("VL_RED_BC")
                                    End If

                                    linha &= "|"                        ' Códifo da observação do lançamento fiscal
                                    linha &= "|"
                                End With
                                strm.WriteLine(linha)
                                RegistroD590 += 1
                                RegistroGeral += 1
                            Next
                        End If
                    Next
                End If

                ' Registro D990  - Encerramento do Bloco D
                RegistroD990 += 1

                linha = "|D990"
                linha &= "|" & RegistroD001 + RegistroD100 + RegistroD190 + RegistroD500 + RegistroD590 + RegistroD990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1

                ' Registro E001  - Abertura do Bloco E
                linha = "|E001"
                linha &= "|0"       'Indicador de Movimento - 0 - Bloco com Dados Informados
                '                                             1 - Bloco sem dados informados
                linha &= "|"

                strm.WriteLine(linha)
                RegistroE001 += 1
                RegistroGeral += 1

                ' Registro E100  - Periodo da Apuração do ICMS
                linha = "|E100"
                linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                                     'Data Inicial
                linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                                       'Data Final
                linha &= "|"

                strm.WriteLine(linha)
                RegistroE100 += 1
                RegistroGeral += 1

                ' Registro E110  - Apuração do ICms - Operações Próprias

                Sql = "  SELECT 'E110' AS Reg,     " & vbCrLf & _
                      " SUM(CASE WHEN Codigo_Id = 1 THEN Valor ELSE 0 END) AS VL_TOT_DEBITOS,    " & vbCrLf & _
                      " SUM(CASE WHEN Codigo_Id = 2 THEN 0 ELSE 0 END) AS VL_AJ_DEBITOS," & vbCrLf & _
                      " SUM(CASE WHEN Codigo_Id = 2 THEN Valor ELSE 0 END) AS VL_TOT_AJ_DEBITOS,    " & vbCrLf

                Sql &= " SUM(CASE WHEN Codigo_Id = 3 THEN Valor ELSE 0 END) AS VL_ESTORNOS_CRED,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 5 THEN Valor ELSE 0 END) AS VL_TOT_CREDITOS,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 6 THEN 0 ELSE 0 END) AS VL_AJ_CREDITOS,    " & vbCrLf

                Sql &= " SUM(CASE WHEN Codigo_Id = 6 THEN Valor ELSE 0 END) AS VL_TOT_AJ_CREDITOS,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 7 THEN Valor ELSE 0 END) AS VL_ESTORNO_DEB,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 9 THEN Valor ELSE 0 END) AS VL_SLD_CREDOR_ANT,    " & vbCrLf

                Sql &= " SUM(CASE WHEN Codigo_Id = 11 THEN Valor ELSE 0 END) AS VL_SLD_APURADO,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 12 THEN Valor ELSE 0 END) AS VL_TOT_DEB,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 13 THEN Valor ELSE 0 END) AS VL_ICMS_RECOLHER,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 14 THEN Valor ELSE 0 END) AS VL_SLD_CREDOR_TRANSPORTAR,    " & vbCrLf & _
                       " SUM(CASE WHEN Codigo_Id = 14 THEN 0 ELSE 0 END) AS DEB_ESP    " & vbCrLf & _
                       "  FROM ResumoRaIcms" & vbCrLf & _
                       " WHERE Empresa_Id = '" & Empresa(0) & "' AND (Processo_Id = " & DdlProcesso.SelectedValue & ")" & vbCrLf & _
                       " GROUP BY Processo_Id" & vbCrLf

                Dim ds24 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds24.Tables(0).Rows.Count > 0 Then
                    For Each dr24 As DataRow In ds24.Tables(0).Rows
                        With dr24
                            'linha = "|E110"
                            'linha &= "|" & .Item("VL_TOT_DEBITOS")
                            'linha &= "|" & .Item("VL_AJ_DEBITOS")
                            'linha &= "|0" '& .Item("VL_TOT_AJ_DEBITOS")

                            'linha &= "|" & .Item("VL_ESTORNOS_CRED")
                            'linha &= "|" & .Item("VL_TOT_CREDITOS")
                            'linha &= "|" & .Item("VL_AJ_CREDITOS")

                            'linha &= "|" & .Item("VL_TOT_AJ_CREDITOS")
                            'linha &= "|" & .Item("VL_ESTORNO_DEB")
                            'linha &= "|" & .Item("VL_SLD_CREDOR_ANT")

                            'linha &= "|" & .Item("VL_SLD_APURADO")
                            'linha &= "|0" '& .Item("VL_TOT_DEBITOS")  '& .Item("VL_TOT_DEB")
                            'linha &= "|" & .Item("VL_ICMS_RECOLHER")

                            'IcmsARecolher = .Item("VL_ICMS_RECOLHER")

                            'linha &= "|" & .Item("VL_SLD_CREDOR_TRANSPORTAR")
                            'linha &= "|" & .Item("DEB_ESP")

                            'linha &= "|"

                            linha = "|E110"
                            linha &= "|" & .Item("VL_TOT_DEBITOS")
                            linha &= "|" & .Item("VL_AJ_DEBITOS")
                            linha &= "|" & .Item("VL_TOT_AJ_DEBITOS")

                            linha &= "|" & .Item("VL_ESTORNOS_CRED")
                            linha &= "|" & .Item("VL_TOT_CREDITOS")
                            linha &= "|" & .Item("VL_AJ_CREDITOS")

                            linha &= "|" & .Item("VL_TOT_AJ_CREDITOS")
                            linha &= "|" & .Item("VL_ESTORNO_DEB")
                            linha &= "|" & .Item("VL_SLD_CREDOR_ANT")

                            linha &= "|" & .Item("VL_SLD_APURADO")
                            linha &= "|" & .Item("VL_TOT_DEB")
                            linha &= "|" & .Item("VL_ICMS_RECOLHER")

                            IcmsARecolher = .Item("VL_ICMS_RECOLHER")

                            linha &= "|" & .Item("VL_SLD_CREDOR_TRANSPORTAR")
                            linha &= "|" & .Item("DEB_ESP")

                            linha &= "|"

                            strm.WriteLine(linha)
                            RegistroE110 += 1
                            RegistroGeral += 1
                        End With
                    Next
                End If

                ' Registro E111  - Ajustes/Beneficio/incentivo - Apuracao do Icms
                Sql = " SELECT 'E111' AS Reg,  " & vbCrLf & _
                      "        ROW_NUMBER() OVER (PARTITION BY ResumoItensRAICMS.Processo_Id,ResumoItensRAICMS.Codigo_Id,ResumoItensRAICMS.Descricao order by ResumoItensRAICMS.Codigo_Id) AS Sequencia_Id, " & vbCrLf & _
                      "        CASE  " & vbCrLf & _
                      "          WHEN ResumoItensRAICMS.Codigo_Id = 2  THEN AjustesApuracaoIcms_Id     " & vbCrLf & _
                      "          WHEN ResumoItensRAICMS.Codigo_Id = 3  THEN AjustesApuracaoIcms_Id     " & vbCrLf & _
                      "          WHEN ResumoItensRAICMS.Codigo_Id = 6  THEN AjustesApuracaoIcms_Id     " & vbCrLf & _
                      "          WHEN ResumoItensRAICMS.Codigo_Id = 7  THEN AjustesApuracaoIcms_Id     " & vbCrLf & _
                      "          WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN AjustesApuracaoIcms_Id    " & vbCrLf & _
                      "        END AS COD_AJ_APUR,  " & vbCrLf & _
                      "        ISNULL(CASE " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 2  THEN AjustesDaApuracaoIcms.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 3  THEN AjustesDaApuracaoIcms.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 6  THEN AjustesDaApuracaoIcms.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 7  THEN AjustesDaApuracaoIcms.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN AjustesDaApuracaoIcms.Descricao " & vbCrLf & _
                      "               END,'') AS DESCR_AJ_APUR,    " & vbCrLf & _
                      "        ISNULL(CASE " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 2  THEN ResumoItensRAICMS.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 3  THEN ResumoItensRAICMS.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 6  THEN ResumoItensRAICMS.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 7  THEN ResumoItensRAICMS.Descricao  " & vbCrLf & _
                      "                 WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN ResumoItensRAICMS.Descricao " & vbCrLf & _
                      "               END,'') AS DESCR_COMPL_AJ,    " & vbCrLf & _
                      "        SUM(CASE  " & vbCrLf & _
                      "              WHEN ResumoItensRAICMS.Codigo_Id = 2 THEN Valor  " & vbCrLf & _
                      "              WHEN ResumoItensRAICMS.Codigo_Id = 3 THEN Valor  " & vbCrLf & _
                      "              WHEN ResumoItensRAICMS.Codigo_Id = 6 THEN Valor  " & vbCrLf & _
                      "              WHEN ResumoItensRAICMS.Codigo_Id = 7 THEN Valor  " & vbCrLf & _
                      "              WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN Valor " & vbCrLf & _
                      "              ELSE 0 " & vbCrLf & _
                      "            END) AS VL_AJ_APUR    " & vbCrLf & _
                      "   FROM ResumoItensRAICMS " & vbCrLf & _
                      "   LEFT JOIN  AjustesDaApuracaoIcms" & vbCrLf & _
                      "     ON ResumoItensRAICMS.AjustesApuracaoIcms_Id = AjustesDaApuracaoIcms.Codigo_Id " & vbCrLf & _
                      "  WHERE Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf & _
                      "    AND Processo_Id = " & DdlProcesso.SelectedValue & vbCrLf & _
                      " GROUP BY ResumoItensRAICMS.Processo_Id,ResumoItensRAICMS.Codigo_Id,ResumoItensRAICMS.Descricao, " & vbCrLf & _
                      "          ResumoItensRAICMS.AjustesApuracaoIcms_Id,AjustesDaApuracaoIcms.descricao " & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Consulta")
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows

                        Dim codAjuApur As String = String.Empty

                        With dr

                            codAjuApur = .Item("COD_AJ_APUR")

                            linha = "|E111"
                            'linha &= "|" & Estado & "0" & .Item("COD_AJ_APUR") & funcoes.AlinharDireita(.Item("Sequencia_Id"), 4, "0")
                            linha &= "|" & .Item("COD_AJ_APUR")
                            linha &= "|" & LTrim(RTrim(.Item("DESCR_AJ_APUR")))
                            linha &= "|" & .Item("VL_AJ_APUR")
                            linha &= "|"
                        End With
                        strm.WriteLine(linha)
                        RegistroE111 += 1
                        RegistroGeral += 1

                        If codAjuApur = "RS020120" Then
                            ' Registro E112  - INFORMAÇÕES ADICIONAIS DOS AJUSTES DA APURAÇÃO DO ICMS
                            'TEMOS QUE FAZER - 'SÓ EXISTE SE TIVER O E111
                            Sql = "select Processo_id as Processo" & vbCrLf & _
                                    "from ProcessoRAICMS" & vbCrLf & _
                                    "where Empresa_Id  = '" & Empresa(0) & "' " & vbCrLf & _
                                    "  AND Processo_Id = " & DdlProcesso.SelectedValue

                            Dim dsE112 As DataSet = Banco.ConsultaDataSet(Sql, "ConsultaE112")

                            If dsE112.Tables(0).Rows.Count > 0 Then
                                For Each drE112 As DataRow In dsE112.Tables(0).Rows
                                    linha = "|E112"
                                    linha &= "|"
                                    linha &= "|" & drE112("Processo")
                                    linha &= "|0"
                                    linha &= "|" & "VALOR REF. RECOLHIMENTO ANTECIPADO"
                                    linha &= "|" & "VALORES INFORMADOS NO ANEXO VIII DA GIA - ICMS PROPRIO" 'IE: " & drr("Inscricao") - VER DE ONDE VEM ESSA INSCRIÇÃO
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroE112 += 1
                                    RegistroGeral += 1
                                Next
                            End If

                            ' Registro E113  - INFORMAÇÕES ADICIONAIS DOS AJUSTES DA APURAÇÃO DO ICMS – IDENTIFICAÇÃO DOS DOCUMENTOS FISCAIS
                            Sql = "select (n.Cliente_Id + convert(varchar,n.EndEmpresa_Id)) as COD_PART," & vbCrLf & _
                                    "		case" & vbCrLf & _
                                    "			when n.Eletronica = 'S'" & vbCrLf & _
                                    "				then '55'" & vbCrLf & _
                                    "				else '01'" & vbCrLf & _
                                    "		end as COD_MOD," & vbCrLf & _
                                    "		n.Serie_Id as SER," & vbCrLf & _
                                    "		ni.Nota_id AS NUM_DOC," & vbCrLf & _
                                    "		case" & vbCrLf & _
                                    "			when n.NossaEmissao = 'S'" & vbCrLf & _
                                    "				then n.Movimento" & vbCrLf & _
                                    "				else n.DataDaNota" & vbCrLf & _
                                    "		end as DT_DOC," & vbCrLf & _
                                    "		ni.Produto_Id AS COD_ITEM," & vbCrLf & _
                                    "		nei.Valor as VL_AJ_ITEM," & vbCrLf & _
                                    "		isnull(nfe.ChaveNfe,'') AS CHV_DOCe" & vbCrLf & _
                                    "from NotasFiscaisXItens ni" & vbCrLf & _
                                    "		inner join NotasFiscaisXEncargos ne" & vbCrLf & _
                                    "				ON ne.Empresa_Id       = ni.Empresa_Id" & vbCrLf & _
                                    "		 		and ne.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf & _
                                    "				and ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf & _
                                    "				and ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf & _
                                    "				and ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf & _
                                    "				and ne.Serie_Id        = ni.Serie_Id" & vbCrLf & _
                                    "				and ne.Nota_Id         = ni.Nota_Id" & vbCrLf & _
                                    "				and ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf & _
                                    "				and ne.Produto_Id      = ni.Produto_Id" & vbCrLf & _
                                    "				and ne.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf & _
                                    "				and ne.Encargo_Id      = 'PRODUTO'" & vbCrLf & _
                                    "		inner join NotasFiscaisXEncargos nei" & vbCrLf & _
                                    "				ON nei.Empresa_Id       = ni.Empresa_Id" & vbCrLf & _
                                    "		 		and nei.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf & _
                                    "				and nei.Cliente_Id      = ni.Cliente_Id" & vbCrLf & _
                                    "				and nei.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf & _
                                    "				and nei.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf & _
                                    "				and nei.Serie_Id        = ni.Serie_Id" & vbCrLf & _
                                    "				and nei.Nota_Id         = ni.Nota_Id" & vbCrLf & _
                                    "				and nei.CFOP_Id         = ni.CFOP_Id" & vbCrLf & _
                                    "				and nei.Produto_Id      = ni.Produto_Id	" & vbCrLf & _
                                    "				and nei.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf & _
                                    "				and nei.Encargo_Id      = 'ICMS'" & vbCrLf & _
                                    "		inner join NotasFiscais n" & vbCrLf & _
                                    "				ON n.Empresa_Id       = ni.Empresa_Id" & vbCrLf & _
                                    "		 		and n.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf & _
                                    "				and n.Cliente_Id      = ni.Cliente_Id" & vbCrLf & _
                                    "				and n.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf & _
                                    "				and n.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf & _
                                    "				and n.Serie_Id        = ni.Serie_Id" & vbCrLf & _
                                    "				and n.Nota_Id         = ni.Nota_Id" & vbCrLf & _
                                    "		left join NFERealizadas nfe" & vbCrLf & _
                                    "				ON nfe.Empresa_Id       = n.Empresa_Id" & vbCrLf & _
                                    "		 		and nfe.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf & _
                                    "				and nfe.Cliente_Id      = n.Cliente_Id" & vbCrLf & _
                                    "				and nfe.EndCliente_Id   = n.EndCliente_Id" & vbCrLf & _
                                    "				and nfe.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf & _
                                    "				and nfe.Serie_Id        = n.Serie_Id" & vbCrLf & _
                                    "				and nfe.Nota_Id         = n.Nota_Id" & vbCrLf & _
                                    "		inner join OperacaoXEstado oe" & vbCrLf & _
                                    "				on oe.Codigo_Id = ni.OperacaoXEstado" & vbCrLf & _
                                    "		inner join SubOperacoes so" & vbCrLf & _
                                    "				ON so.Operacao_Id      = ni.Operacao" & vbCrLf & _
                                    "		 		and so.SubOperacoes_Id = ni.SubOperacao" & vbCrLf & _
                                    "where n.Situacao = 1" & vbCrLf & _
                                    "  and n.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                                    "  and n.Movimento Between '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                                    "  and ni.CFOP_Id in(6101,6152)" & vbCrLf & _
                                    "  and (oe.CodigoBeneficio = 'RS051609' or nei.Valor > 0)" & vbCrLf & _
                                    "ORDER BY n.Movimento, n.Nota_Id"

                            Dim dsE113 As DataSet = Banco.ConsultaDataSet(Sql, "ConsultaE113")

                            If dsE113.Tables(0).Rows.Count > 0 Then
                                For Each drE113 As DataRow In dsE113.Tables(0).Rows
                                    linha = "|E113"
                                    linha &= "|" & drE113("COD_PART")
                                    linha &= "|" & drE113("COD_MOD")
                                    linha &= "|" & drE113("SER")
                                    linha &= "|"
                                    linha &= "|" & drE113("NUM_DOC")
                                    linha &= "|" & CDate(drE113("DT_DOC")).ToString("ddMMyyyy")
                                    linha &= "|" & drE113("COD_ITEM")
                                    linha &= "|" & drE113("VL_AJ_ITEM")
                                    linha &= "|" & drE113("CHV_DOCe")
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroE113 += 1
                                    RegistroGeral += 1
                                Next
                            End If

                        End If

                    Next
                End If


                ' Registro E115  - INFORMAÇÕES ADICIONAIS DA APURAÇÃO – VALORES DECLARATÓRIOS. (Observações ICMS)
                Sql = "  SELECT CONVERT(DECIMAL(18,2), ISNULL(SUM(Valor), 0)) AS Valor, SpedInfAdicionaisDeApuracao " & vbCrLf & _
                      "    FROM DarDiferencialDeAliquota " & vbCrLf & _
                      "   WHERE YEAR(DataReferencia)= " & CDate(txtDataInicial.Text).Year & vbCrLf & _
                      "     AND MONTH(DataReferencia) = " & CDate(txtDataInicial.Text).Month & vbCrLf & _
                      "     AND Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf & _
                      "     AND EndEmpresa_Id  =" & Empresa(1) & vbCrLf & _
                      "   GROUP BY SpedInfAdicionaisDeApuracao" & vbCrLf

                ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds1.Tables(0).Rows.Count > 0 Then
                    For Each dr1 As DataRow In ds1.Tables(0).Rows
                        With dr1
                            linha = "|E115" 'REG
                            linha &= "|" & .Item("SpedInfAdicionaisDeApuracao") 'COD_INF_ADIC - Código da informação adicional conforme tabela a ser definida pelas SEFAZ, conforme tabela definida no item 5.2.
                            linha &= "|" & .Item("Valor") ' VL_INF_ADIC
                            linha &= "|" & "|" 'DESCR_COMPL_AJ'
                            strm.WriteLine(linha)
                            RegistroE115 += 1
                            RegistroGeral += 1
                        End With
                    Next
                End If


                ' Registro E116  - Obrigações do Icms A REcolher - Operações Próprias

                Sql = "  SELECT * " & vbCrLf & _
                      " FROM ProcessoRaIcms" & vbCrLf & _
                      " WHERE Empresa_Id = '" & Empresa(0) & "' AND (Processo_Id = " & DdlProcesso.SelectedValue & ")" & vbCrLf

                ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds1.Tables(0).Rows.Count > 0 Then
                    For Each dr1 As DataRow In ds1.Tables(0).Rows
                        With dr1
                            linha = "|E116"
                            linha &= "|000" 'Código da Obrigação a Recolher, conforme a Tabela 5.4
                            linha &= "|" & IcmsARecolher
                            linha &= "|" & Format(CDate(.Item("VencimentoDaObrigacao")), "ddMMyyyy")
                            linha &= "|" & .Item("CodigoDaReceita")
                            linha &= "|"                'Numero do Processo
                            linha &= "|"                'Indicador da Origem do Processo
                            linha &= "|"                'Descricao Resumida do Processo
                            linha &= "|"                'Descricao complementar das obrigações a recolher
                            linha &= IIf(PeriodoFinal.Year < 2011, "|", "|" & PeriodoFinal.ToString("MMyyyy") & "|")

                            strm.WriteLine(linha)
                            RegistroE116 += 1
                            RegistroGeral += 1
                        End With
                    Next
                End If

                If Left(Empresa(0), 8) = "40938762" Then

                    Sql = "SELECT n.EstadoDoCliente," & vbCrLf &
                            "	   isnull(sum(CASE WHEN n.EntradaSaida_Id = 'S' then ne.Valor else 0 end),0) as ValorSaida," & vbCrLf &
                            "	   isnull(sum(CASE WHEN n.EntradaSaida_Id = 'E' then ne.Valor else 0 end),0) as ValorDevolucao" & vbCrLf &
                            "from NotasFiscaisXEncargos ne" & vbCrLf &
                            "		inner join NotasFiscaisXItens ni" & vbCrLf &
                            "				ON ni.Empresa_Id       = ne.Empresa_Id" & vbCrLf &
                            "				and ni.EndEmpresa_iD   = ne.EndEmpresa_Id" & vbCrLf &
                            "				and ni.Cliente_Id      = ne.Cliente_Id" & vbCrLf &
                            "				and ni.EndCliente_Id   = ne.EndCliente_Id" & vbCrLf &
                            "				and ni.EntradaSaida_Id = ne.EntradaSaida_Id" & vbCrLf &
                            "				and ni.Serie_Id        = ne.Serie_Id" & vbCrLf &
                            "				and ni.Nota_Id         = ne.Nota_Id" & vbCrLf &
                            "				and ni.Produto_Id      = ne.Produto_Id" & vbCrLf &
                            "				and ni.CFOP_Id         = ne.CFOP_Id" & vbCrLf &
                            "				and ni.Sequencia_Id    = ne.Sequencia_id" & vbCrLf &
                            "		inner join NotasFiscais n" & vbCrLf &
                            "				ON ni.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                            "				and ni.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf &
                            "				and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                            "				and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                            "				and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                            "				and ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                            "				and ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                            "where n.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                            "  and n.Situacao = 1" & vbCrLf &
                            "  and n.Movimento Between '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                            "  and ne.Encargo_Id = 'ICMS-ST'" & vbCrLf &
                            "group by n.EstadoDoCliente" & vbCrLf &
                            "order by n.EstadoDoCliente"

                    ds1 = Banco.ConsultaDataSet(Sql, "ConsultaSTICMS")

                    If ds1.Tables(0).Rows.Count > 0 Then

                        For Each dr1 As DataRow In ds1.Tables(0).Rows

                            ' Registro E200  -  PERÍODO DA APURAÇÃO DO ICMS - SUBSTITUIÇÃO TRIBUTÁRIA

                            linha = "|E200"
                            linha &= "|" & dr1.Item("EstadoDoCliente")
                            linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                                     'Data Inicial
                            linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                                       'Data Final
                            linha &= "|"

                            strm.WriteLine(linha)
                            RegistroE200 += 1
                            RegistroGeral += 1


                            ' Registro E210  -  APURAÇÃO DO ICMS – SUBSTITUIÇÃO TRIBUTÁRIA

                            With dr1
                                linha = "|E210" '01 - Texto fixo contendo "E210"
                                If CDec(.Item("ValorSaida")) > 0 OrElse CDec(.Item("ValorDevolucao")) Then
                                    linha &= "|1"  '02 - Com operações de ST
                                Else
                                    linha &= "|0"  '02 - Sem operações com ST
                                End If

                                linha &= "|0,00"                       '03 - Valor do "Saldo credor de período anterior Substituição Tributária
                                linha &= "|" & .Item("ValorDevolucao") '04 - Valor total do ICMS ST de devolução de mercadorias
                                linha &= "|0,00" '05 - Valor total do ICMS ST de ressarcimentos
                                linha &= "|0,00" '06 - Valor total de Ajustes "Outros créditos ST" e “Estorno de débitos ST”
                                linha &= "|0,00" '07 -  Valor total dos ajustes a crédito de ICMS ST, provenientes de ajustes do documento fiscal.
                                linha &= "|" & .Item("ValorSaida") '08 - Valor Total do ICMS retido por Substituição Tributária
                                linha &= "|0,00" '09 - Valor Total dos ajustes "Outros débitos ST" " e “Estorno de créditos ST”
                                linha &= "|0,00" '10 - Valor total dos ajustes a débito de ICMS ST, provenientes de ajustes do documento fiscal.
                                linha &= "|" & (.Item("ValorSaida") - .Item("ValorDevolucao")) '11 - Valor total de Saldo devedor antes das deduções
                                linha &= "|0,00" '12 - Valor total dos ajustes "Deduções ST"
                                linha &= "|" & (.Item("ValorSaida") - .Item("ValorDevolucao")) '13 - Imposto a recolher ST (11-12)
                                linha &= "|0,00" '14 - Saldo credor de ST a transportar para o período seguinte [(03+04+05+06+07+12) – (08+09+10)]
                                linha &= "|0,00" '15 - Valores recolhidos ou a recolher, extra-apuração.
                                linha &= "|"

                                strm.WriteLine(linha)
                                RegistroE210 += 1
                                RegistroGeral += 1

                            End With



                            ' Registro E250  -  APURAÇÃO DO ICMS – SUBSTITUIÇÃO TRIBUTÁRIA

                            With dr1
                                linha = "|E250" '01 - Texto fixo contendo "E250"
                                linha &= "|999" '02 - Código da obrigação a recolher, conforme a Tabela 5.4
                                linha &= "|" & (.Item("ValorSaida") - .Item("ValorDevolucao")) '3 - Valor da obrigação ICMS ST a recolher
                                linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")    '4 - Data de vencimento da obrigação

                                If dr1.Item("EstadoDoCliente") = "SP" Then
                                    linha &= "|246-0" '05 - Código de receita referente à obrigação, próprio da unidade da federação do contribuinte substituído
                                ElseIf dr1.Item("EstadoDoCliente") = "PR" OrElse dr1.Item("EstadoDoCliente") = "RJ" Then
                                    linha &= "|100048" '05 - Código de receita referente à obrigação, próprio da unidade da federação do contribuinte substituído
                                Else
                                    linha &= "|100099" '05 - Código de receita referente à obrigação, próprio da unidade da federação do contribuinte substituído
                                End If

                                linha &= "|" '06 - Número do processo ou auto de infração ao qual a obrigação está vinculada, se houver
                                linha &= "|" '07 - Indicador da origem do processo: 0- SEFAZ; 1- Justiça Federal; 2- Justiça Estadual; 9- Outros
                                linha &= "|" '08 - Descrição resumida do processo que embasou o lançamento
                                linha &= "|" '09 - Descrição complementar das obrigações a recolher
                                linha &= "|" & Format(CDate(txtDataFinal.Text), "MMyyyy") '10 - Informe o mês de referência no formato “mmaaaa”
                                linha &= "|"

                                strm.WriteLine(linha)
                                RegistroE250 += 1
                                RegistroGeral += 1

                            End With

                        Next

                        ' Registro E300  -  PERÍODO DE APURAÇÃO DO FUNDO DE COMBATE À POBREZA E DO ICMS DIFERENCIAL DE ALÍQUOTA – UF ORIGEM/DESTINO EC 87/15

                        Sql = "SELECT n.EstadoDoCliente," & vbCrLf &
                            "	   SUM(CASE WHEN so.Devolucao = 'N' then ne.Valor else 0 end) AS Valor," & vbCrLf &
                            "	   SUM(CASE WHEN so.Devolucao = 'S' then ne.Valor else 0 end) AS ValorDevolucao" & vbCrLf &
                            "from NotasFiscaisXEncargos ne" & vbCrLf &
                            "		inner join NotasFiscaisXItens ni" & vbCrLf &
                            "				ON ni.Empresa_Id       = ne.Empresa_Id" & vbCrLf &
                            "				and ni.EndEmpresa_iD   = ne.EndEmpresa_Id" & vbCrLf &
                            "				and ni.Cliente_Id      = ne.Cliente_Id" & vbCrLf &
                            "				and ni.EndCliente_Id   = ne.EndCliente_Id" & vbCrLf &
                            "				and ni.EntradaSaida_Id = ne.EntradaSaida_Id" & vbCrLf &
                            "				and ni.Serie_Id        = ne.Serie_Id" & vbCrLf &
                            "				and ni.Nota_Id         = ne.Nota_Id" & vbCrLf &
                            "				and ni.Produto_Id      = ne.Produto_Id" & vbCrLf &
                            "				and ni.CFOP_Id         = ne.CFOP_Id" & vbCrLf &
                            "				and ni.Sequencia_Id    = ne.Sequencia_id" & vbCrLf &
                            "		inner join NotasFiscais n" & vbCrLf &
                            "				ON ni.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                            "				and ni.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf &
                            "				and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                            "				and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                            "				and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                            "				and ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                            "				and ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                            "		inner join SubOperacoes so" & vbCrLf &
                            "				ON so.Operacao_Id      = n.Operacao" & vbCrLf &
                            "				and so.SubOperacoes_Id = n.SubOperacao" & vbCrLf &
                            "where n.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                            "  and n.Situacao = 1" & vbCrLf &
                            "  and n.Movimento Between '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                            "  and ne.Encargo_Id = 'FUNDO FECP'" & vbCrLf &
                            "group by n.EstadoDoCliente" & vbCrLf &
                            "order by n.EstadoDoCliente"

                        ds = Banco.ConsultaDataSet(Sql, "ConsultaE300")

                        If ds.Tables(0).Rows.Count > 0 Then

                            For Each dr300 As DataRow In ds.Tables(0).Rows

                                If (dr300.Item("Valor") - dr300.Item("ValorDevolucao")) > 0 Then

                                    linha = "|E300"
                                    linha &= "|" & dr300.Item("EstadoDoCliente")
                                    linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy") 'Data Inicial
                                    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")   'Data Final
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroE300 += 1
                                    RegistroGeral += 1

                                    ' Registro E310  -  PERÍODO DE APURAÇÃO DO FUNDO DE COMBATE À POBREZA E DO ICMS DIFERENCIAL DE ALÍQUOTA – UF ORIGEM/DESTINO EC 87/15

                                    linha = "|E310"  '01 - Texto fixo contendo "E310"
                                    linha &= "|1"    '02 - Indicador de movimento: 0 – Sem operações 1 – Com operações
                                    linha &= "|0,00" '03 - Valor do "Saldo credor de período anterior – ICMS Diferencial de Alíquota da UF de Origem/Destino"
                                    linha &= "|0,00" '04 - Valor total dos débitos por "Saídas e prestações com débito do ICMS referente ao diferencial de alíquota devido à UF de Origem/Destino
                                    linha &= "|0,00" '05 - Valor total dos ajustes "Outros débitos ICMS Diferencial de Alíquota da UF de Origem/Destino" e “Estorno de créditos ICMS Diferencial de Alíquota da UF de Origem/Destino”
                                    linha &= "|0,00" '06 - Valor total dos créditos do ICMS referente ao diferencial de alíquota devido à UF de Origem/Destino
                                    linha &= "|0,00" '07 - Valor total de Ajustes "Outros créditos ICMS Diferencial de Alíquota da UF de Origem/Destino" e “Estorno de débitos ICMS Diferencial de Alíquota da UF de Origem/Destino”
                                    linha &= "|0,00" '08 - Valor total de “Saldo devedor ICMS Diferencial de Alíquota da UF de Origem/Destino antes das deduções”
                                    linha &= "|0,00" '09 - Valor total dos ajustes "Deduções ICMS Diferencial de Alíquota da UF de Origem/Destino"
                                    linha &= "|0,00" '10 - Valor recolhido ou a recolher referente ao ICMS Diferencial de Alíquota da UF de Origem/Destino (08-09)
                                    linha &= "|0,00" '11 - Saldo credor a transportar para o período seguinte referente ao ICMS Diferencial de Alíquota da UF de Origem/Destino
                                    linha &= "|0,00" '12 - Valores recolhidos ou a recolher, extra-apuração - ICMS Diferencial de Alíquota da UF de Origem/Destino.
                                    linha &= "|0,00" '13 - Valor do "Saldo credor de período anterior – FCP"
                                    linha &= "|" & dr300.Item("Valor") '14 - Valor total dos débitos FCP por "Saídas e prestações”
                                    linha &= "|0,00" '15 - Valor total dos ajustes "Outros débitos FCP" e “Estorno de créditos FCP”
                                    linha &= "|" & dr300.Item("ValorDevolucao") '16 - Valor total dos créditos FCP por Entradas
                                    linha &= "|0,00" '17 - Valor total de Ajustes "Outros créditos FCP" e “Estorno de débitos FCP”
                                    linha &= "|" & (dr300.Item("Valor") - dr300.Item("ValorDevolucao")) '18 - Valor total de Saldo devedor FCP antes das deduções
                                    linha &= "|0,00" '19 - Valor total das deduções "FCP"
                                    linha &= "|" & (dr300.Item("Valor") - dr300.Item("ValorDevolucao")) '20 - Valor recolhido ou a recolher referente ao FCP (18–19)
                                    linha &= "|0,00" '21 - Saldo credor a transportar para o período seguinte referente ao FCP
                                    linha &= "|0,00" '22 - Valores recolhidos ou a recolher, extra-apuração - FCP
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroE310 += 1
                                    RegistroGeral += 1

                                    ' Registro E316  -  PERÍODO DE APURAÇÃO DO FUNDO DE COMBATE À POBREZA E DO ICMS DIFERENCIAL DE ALÍQUOTA – UF ORIGEM/DESTINO EC 87/15

                                    linha = "|E316" '01 - Texto fixo contendo "E316"
                                    linha &= "|090" '02 - Código da obrigação recolhida ou a recolher, conforme a Tabela 5.4    
                                    linha &= "|" & (dr300.Item("Valor") - dr300.Item("ValorDevolucao")) '03 - Valor da obrigação recolhida ou a recolher
                                    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")  '04 - Data de vencimento da obrigação 

                                    If dr300.Item("EstadoDoCliente") = "RJ" Then
                                        linha &= "|100137" '05 - Código de receita referente à obrigação, próprio da unidade da federação da origem/destino, conforme legislação estadual.
                                    Else
                                        linha &= "|100129" '05 - Código de receita referente à obrigação, próprio da unidade da federação da origem/destino, conforme legislação estadual.
                                    End If

                                    linha &= "|" '06 - Número do processo ou auto de infração ao qual a obrigação está vinculada, se houver
                                    linha &= "|" '07 - Indicador da origem do processo: 0- SEFAZ; 1- Justiça Federal; 2- Justiça Estadual; 9- Outros
                                    linha &= "|" '08 - Descrição resumida do processo que embasou o lançamento
                                    linha &= "|" '09 - Descrição complementar das obrigações recolhidas ou a recolher
                                    linha &= "|" & Format(CDate(txtDataFinal.Text), "MMyyyy")  '10 - Informe o mês de referência no formato “mmaaaa”
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroE316 += 1
                                    RegistroGeral += 1

                                End If

                            Next

                        End If


                        ''''' Registro E220
                        ''''linha = "|E220" '01 - Texto fixo contendo "E220"
                        ''''linha &= "|00000002"   '2 - 
                        ''''linha &= "|OUTROS CREDITOS"
                        ''''linha &= "|"
                        ''''linha &= "|"

                        ''''strm.WriteLine(linha)
                        ''''RegistroE220 += 1
                        ''''RegistroGeral += 1

                        ''''' Registro E240  -  APURAÇÃO DO ICMS – INFORMAÇÕES ADICIONAIS DOS AJUSTES DA APURAÇÃO DO ICMS SUBSTITUIÇÃO TRIBUTÁRIA – IDENTIFICAÇÃO DOS DOCUMENTOS FISCAIS

                        ''''Sql = "SELECT n.Cliente_Id, n.EndCliente_Id," & vbCrLf &
                        ''''  "		case" & vbCrLf &
                        ''''  "			when n.Eletronica = 'S'" & vbCrLf &
                        ''''  "				then '55'" & vbCrLf &
                        ''''  "				else '01'" & vbCrLf &
                        ''''  "		end as COD_MOD," & vbCrLf &
                        ''''  "	   n.Serie_Id, '' as SUB, n.Nota_Id, n.Movimento, ni.Produto_Id, ni.Valor, nfe.ChaveNfe" & vbCrLf &
                        ''''  "from NotasFiscaisXEncargos ne" & vbCrLf &
                        ''''  "		inner join NotasFiscaisXItens ni" & vbCrLf &
                        ''''  "				ON ni.Empresa_Id       = ne.Empresa_Id" & vbCrLf &
                        ''''  "				and ni.EndEmpresa_iD   = ne.EndEmpresa_Id" & vbCrLf &
                        ''''  "				and ni.Cliente_Id      = ne.Cliente_Id" & vbCrLf &
                        ''''  "				and ni.EndCliente_Id   = ne.EndCliente_Id" & vbCrLf &
                        ''''  "				and ni.EntradaSaida_Id = ne.EntradaSaida_Id" & vbCrLf &
                        ''''  "				and ni.Serie_Id        = ne.Serie_Id" & vbCrLf &
                        ''''  "				and ni.Nota_Id         = ne.Nota_Id" & vbCrLf &
                        ''''  "				and ni.Produto_Id      = ne.Produto_Id" & vbCrLf &
                        ''''  "				and ni.CFOP_Id         = ne.CFOP_Id" & vbCrLf &
                        ''''  "				and ni.Sequencia_Id    = ne.Sequencia_id" & vbCrLf &
                        ''''  "		inner join NotasFiscais n" & vbCrLf &
                        ''''  "				ON ni.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                        ''''  "				and ni.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf &
                        ''''  "				and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                        ''''  "				and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                        ''''  "				and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                        ''''  "				and ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                        ''''  "				and ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                        ''''  "		left join NFERealizadas nfe" & vbCrLf &
                        ''''  "				ON nfe.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                        ''''  "				and nfe.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf &
                        ''''  "				and nfe.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                        ''''  "				and nfe.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                        ''''  "				and nfe.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                        ''''  "				and nfe.Serie_Id        = n.Serie_Id" & vbCrLf &
                        ''''  "				and nfe.Nota_Id         = n.Nota_Id" & vbCrLf &
                        ''''  "where n.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf &
                        ''''  "  and n.Situacao = 1" & vbCrLf &
                        ''''  "  and n.Movimento Between '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                        ''''  "  and n.EstadoDoCliente in('AC','AM','CE','DF','ED','MA','MG','MS','PA','RJ','SP')" & vbCrLf &
                        ''''  "  and ne.Encargo_Id = 'ICMS-ST'"

                        ''''ds1 = Banco.ConsultaDataSet(Sql, "ConsultaSTICMS")

                        ''''For Each dr1 As DataRow In ds1.Tables(0).Rows
                        ''''    With dr1
                        ''''        linha = "|E240" '01 - Texto fixo contendo "E240"
                        ''''        linha &= "|" & Trim(.Item("Cliente_Id")) & Trim(.Item("EndCliente_Id")) '02 - Código do participante (campo 02 do Registro 0150):- do emitente do documento ou do remetente das mercadorias,no caso de entradas;- do adquirente, no caso de saídas
                        ''''        linha &= "|" & .Item("COD_MOD") '03 - Código do modelo do documento fiscal, conforme a Tabela 4.1.1

                        ''''        linha &= "|" & .Item("Serie_Id") '04 - Série do documento fiscal
                        ''''        linha &= "|" & .Item("SUB")      '05 - Subsérie do documento fiscal
                        ''''        linha &= "|" & .Item("Nota_Id")  '06 - Número do documento fiscal
                        ''''        linha &= "|" & Format(CDate(.Item("Movimento")), "ddMMyyyy") '07 - Data da emissão do documento fiscal
                        ''''        linha &= "|" & .Item("Produto_Id")  '08 - Código do item (campo 02 do Registro 0200)
                        ''''        linha &= "|" & .Item("Valor")       '09 - Valor do ajuste para a operação/item
                        ''''        linha &= "|" & .Item("ChaveNfe")    '10 - Chave do Documento Eletrônico
                        ''''        linha &= "|"

                        ''''        strm.WriteLine(linha)
                        ''''        RegistroE240 += 1
                        ''''        RegistroGeral += 1

                        ''''    End With
                        ''''Next
                        ''''

                    End If



                End If


                ' Registro E500  - Periodo da Apuração do IPI
                'Ajustei para Sulgrain não colocar esse bloco, vamos ter q analisar e usar o Código da Atividade Econômica da Empresa, por hora fica assim - Furlan - 25-02-2014
                If Not Empresa(0) = "04376053000162" AndAlso Left(Empresa(0), 8) <> "44979506" AndAlso Left(Empresa(0), 8) <> "62747840" AndAlso Left(Empresa(0), 8) <> "62780383" AndAlso Left(Empresa(0), 8) <> "63358210" Then
                    linha = "|E500"
                    linha &= "|0"
                    linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                                     'Data Inicial
                    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                                       'Data Final
                    linha &= "|"

                    strm.WriteLine(linha)
                    RegistroE500 += 1
                    RegistroGeral += 1

                    ' Registro E510  - Apuração do IPI - Operações Próprias
                    'Temporaria Controla Notas de Serviço Conjugadas


                    Sql = "  Select nfxi.Empresa_Id," & vbCrLf &
                          "         nfxi.EndEmpresa_Id," & vbCrLf &
                          "         nfxi.Cliente_Id, nfxi.EndCliente_Id, nfxi.EntradaSaida_Id, nfxi.Serie_Id," & vbCrLf &
                          "         nfxi.Nota_Id," & vbCrLf &
                          "         sum(Case" & vbCrLf &
                          "		        when OEx.codigoFiscal in (1933,2933)" & vbCrLf &
                          "                 then 1" & vbCrLf &
                          "                 else 0" & vbCrLf &
                          "             end ) ItensCfop," & vbCrLf &
                          "         sum(1) nrItens" & vbCrLf &
                          "    into #NotasCompostas" & vbCrLf &
                          "    From notasfiscaisxitens nfxi" & vbCrLf &
                          "   Inner join OperacaoXEstado OEx" & vbCrLf &
                          "      on OEx.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                          "   Inner join (SELECT nf.Empresa_Id,      nf.EndEmpresa_Id," & vbCrLf &
                          "                      nf.Cliente_Id,      nf.EndCliente_Id," & vbCrLf &
                          "                      nf.EntradaSaida_Id, nf.Serie_Id,      nf.Nota_Id" & vbCrLf &
                          "                 FROM NotasFiscais nf" & vbCrLf &
                          "                INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                          "                   ON nf.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                          "                  AND nf.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                          "                  AND nf.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                          "                  AND nf.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                          "                  AND nf.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                          "                  AND nf.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                          "                  AND nf.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                          "				Inner join OperacaoXEstado OE" & vbCrLf &
                          "				   on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                          "                Inner Join Suboperacoes so" & vbCrLf &
                          "                   on so.Operacao_id     = nfi.Operacao" & vbCrLf &
                          "                  and so.SubOperacoes_id = nfi.suboperacao" & vbCrLf &
                          "                WHERE nf.EntradaSaida_Id = 'E'" & vbCrLf &
                          "                  And nf.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                          "                  AND nf.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                          "                  AND nf.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf &
                          "                  AND nf.TipoDeDocumento NOT IN (15)" & vbCrLf &
                          "                  AND nf.Finalidade NOT IN (30)" & vbCrLf &
                          "                  AND oe.CodigoFiscal  IN (1933,2933)" & vbCrLf &
                          "               )sb" & vbCrLf &
                          "      ON sb.Empresa_Id      = nfxi.Empresa_Id" & vbCrLf &
                          "     AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id" & vbCrLf &
                          "     AND sb.Cliente_Id      = nfxi.Cliente_Id" & vbCrLf &
                          "     AND sb.EndCliente_Id   = nfxi.EndCliente_Id" & vbCrLf &
                          "     AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id" & vbCrLf &
                          "     AND sb.Serie_Id        = nfxi.Serie_Id" & vbCrLf &
                          "     AND sb.Nota_Id         = nfxi.Nota_Id" & vbCrLf &
                          "   group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id," & vbCrLf &
                          "            nfxi.Cliente_Id, nfxi.EndCliente_Id," & vbCrLf &
                          "            nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf &
                          "  having sum(1)  <> sum(Case" & vbCrLf &
                          "                          when OEx.CodigoFiscal in (1933,2933)" & vbCrLf &
                          "                            then 1" & vbCrLf &
                          "                            else 0" & vbCrLf &
                          "                         end);" & vbCrLf &
                          "--****************************************************************************" & vbCrLf &
                          " Select oe.CodigoFiscal as CFOP," & vbCrLf &
                          "        oe.STIPI AS CST_IPI," & vbCrLf &
                          "        Sum(IsNull(case " & vbCrLf &
                          "                     when isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'PRODUTO'" & vbCrLf &
                          "                       then nfe.valor" & vbCrLf &
                          "                       else 0" & vbCrLf &
                          "                   end,0)) as VL_CONT_IPI," & vbCrLf &
                          "        Sum(IsNull(case" & vbCrLf &
                          "                     when isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'" & vbCrLf &
                          "                       then nfe.Base" & vbCrLf &
                          "                       else 0" & vbCrLf &
                          "                   end,0)) as VL_BC_IPI," & vbCrLf &
                          "        Sum(IsNull(case" & vbCrLf &
                          "                     when isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'IPI'" & vbCrLf &
                          "                       then nfe.valor" & vbCrLf &
                          "                       else 0" & vbCrLf &
                          "                   end,0)) as VL_IPI" & vbCrLf &
                          "   From NotasFiscais NF" & vbCrLf &
                          "  Inner Join NotasFiscaisXItens nfi" & vbCrLf &
                          "     ON nfi.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                          "    AND nfi.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                          "    AND nfi.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                          "    AND nfi.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                          "    AND nfi.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                          "    AND nfi.Serie_Id        = NF.Serie_Id" & vbCrLf &
                          "    AND nfi.Nota_Id         = NF.Nota_Id" & vbCrLf &
                          "  INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                          "     ON nfe.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                          "    AND nfe.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                          "    AND nfe.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                          "    AND nfe.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                          "    AND nfe.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                          "    AND nfe.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                          "    AND nfe.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                          "	   AND nfe.Produto_Id      = nfi.Produto_Id" & vbCrLf &
                          "	   AND nfe.Sequencia_id    = nfi.Sequencia_id" & vbCrLf &
                          "  Inner join OperacaoXEstado OE" & vbCrLf &
                          "     on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                          "   LEFT join Encargos e" & vbCrLf &
                          "     on e.Encargo_id = nfe.Encargo_Id" & vbCrLf &
                          "	   and isnull(e.encargoagrupador,'') <> ''" & vbCrLf &
                          "  WHERE (" & vbCrLf &
                          "			(    nfe.Empresa_Id      ='" & Empresa(0) & "'" & vbCrLf &
                          "			 AND nfe.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                          "			 AND isnull(NF.situacao,1) = 1" & vbCrLf &
                          "			 AND NF.Movimento  BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                          "			 AND OE.codigoFiscal NOT IN (1933,2933)" & vbCrLf &
                          "			)" & vbCrLf &
                          "			OR" & vbCrLf &
                          "			(" & vbCrLf &
                          "			 exists (Select  1" & vbCrLf &
                          "					  from #notascompostas" & vbCrLf &
                          "					 where #NotasCompostas.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                          "					   AND #NotasCompostas.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                          "					   AND #NotasCompostas.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                          "					   AND #NotasCompostas.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                          "					   AND #NotasCompostas.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                          "					   AND #NotasCompostas.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                          "					   AND #NotasCompostas.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                          "					)" & vbCrLf &
                          "			 )" & vbCrLf &
                          "         )" & vbCrLf &
                          "    AND isnull(e.EncargoAgrupador,nfe.Encargo_Id) in ('IPI','PRODUTO') " & vbCrLf &
                          "  Group by oe.CodigoFiscal," & vbCrLf &
                          "           oe.STIPI" & vbCrLf &
                          "  ORDER BY oe.CodigoFiscal," & vbCrLf &
                          "           oe.STIPI" & vbCrLf

                    ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds1.Tables(0).Rows.Count > 0 Then
                        For Each dr1 As DataRow In ds1.Tables(0).Rows
                            With dr1
                                linha = "|E510"
                                linha &= "|" & .Item("CFOP")
                                linha &= "|" & Funcoes.AlinharDireita(.Item("CST_IPI"), 2, "0")
                                linha &= "|" & .Item("VL_CONT_IPI")
                                linha &= "|" & .Item("VL_BC_IPI")
                                linha &= "|" & .Item("VL_IPI")
                                linha &= "|"

                            End With
                            strm.WriteLine(linha)
                            RegistroE510 += 1
                            RegistroGeral += 1
                        Next
                    End If


                    ' Registro E520  - Apuração do IPI - Operações Próprias

                    Sql = "  SELECT 'E520' AS Reg,     " & vbCrLf &
                          " SUM(CASE WHEN Codigo_Id = 9 THEN Valor ELSE 0 END) AS VL_SD_ANT_IPI,    " & vbCrLf &
                          " SUM(CASE WHEN Codigo_Id = 1 THEN Valor ELSE 0 END) AS VL_DEB_IPI,    " & vbCrLf &
                          " SUM(CASE WHEN Codigo_Id = 5 THEN Valor ELSE 0 END) AS VL_CRED_IPI,    " & vbCrLf
                    'Sql &= " SUM(CASE WHEN Codigo_Id = 2 or  THEN Valor ELSE 0 END) AS VL_OD_IPI,"
                    'Sql &= " SUM(CASE WHEN Codigo_Id = 6 THEN Valor ELSE 0 END) AS VL_OC_IPI,    "
                    Sql &= " SUM(CASE WHEN Codigo_Id = 2 or Codigo_Id= 3 THEN Valor ELSE 0 END) AS VL_OD_IPI, " & vbCrLf &
                           " SUM(CASE WHEN Codigo_Id = 6 or Codigo_Id= 7 THEN Valor ELSE 0 END) AS VL_OC_IPI, " & vbCrLf &
                           " SUM(CASE WHEN Codigo_Id = 14 THEN Valor ELSE 0 END) AS VL_SC_IPI,    " & vbCrLf &
                           " SUM(CASE WHEN Codigo_Id = 13 THEN Valor ELSE 0 END) AS VL_SD_IPI    " & vbCrLf &
                           " FROM ResumoRaIPI" & vbCrLf &
                           " WHERE Empresa_Id = '" & Empresa(0) & "' AND (Processo_Id = " & DdlProcessoIPI.SelectedValue & ")" & vbCrLf &
                           " GROUP BY Processo_Id" & vbCrLf

                    ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds1.Tables(0).Rows.Count > 0 Then
                        For Each dr1 As DataRow In ds1.Tables(0).Rows
                            With dr1
                                linha = "|E520"
                                linha &= "|" & .Item("VL_SD_ANT_IPI")
                                linha &= "|" & .Item("VL_DEB_IPI")
                                linha &= "|" & .Item("VL_CRED_IPI")
                                linha &= "|" & .Item("VL_OD_IPI")
                                linha &= "|" & .Item("VL_OC_IPI")
                                linha &= "|" & .Item("VL_SC_IPI")
                                linha &= "|" & .Item("VL_SD_IPI")
                                linha &= "|"

                            End With
                            strm.WriteLine(linha)
                            RegistroE520 += 1
                            RegistroGeral += 1
                        Next
                    End If


                    ' Registro E530  - AJUSTES DA APURAÇÃO DO IPI.

                    Sql = " SELECT 'E530' AS Reg,     " & vbCrLf &
                          " CASE WHEN ResumoItensRAIPI.Codigo_Id = 2 OR ResumoItensRAIPI.Codigo_Id = 3 THEN 0 WHEN ResumoItensRAIPI.Codigo_Id = 6 OR ResumoItensRAIPI.Codigo_Id = 7 THEN 1 END AS IND_AJ,    " & vbCrLf &
                          " ResumoItensRAIPI.Valor AS VL_AJ,    " & vbCrLf &
                          " AjustesDaApuracaoIPI.Codigo_Id AS COD_AJ,    " & vbCrLf &
                          " ResumoItensRAIPI.OrigemDocumentoAjusteIpi_Id AS IND_DOC, " & vbCrLf &
                          " '' AS NUM_DOC, " & vbCrLf &
                          " AjustesDaApuracaoIPI.Descricao AS DESCR_AJ    " & vbCrLf &
                          " FROM ResumoItensRAIPI " & vbCrLf &
                          " LEFT JOIN  AjustesDaApuracaoIPI " & vbCrLf &
                          " ON  ResumoItensRAIPI.AjustesApuracaoIPI_Id = AjustesDaApuracaoIPI.Codigo_Id " & vbCrLf &
                          " WHERE ResumoItensRAIPI.Empresa_Id = '" & Empresa(0) & "' AND (ResumoItensRAIPI.Processo_Id = " & DdlProcessoIPI.SelectedValue & ") AND ResumoItensRAIPI.Codigo_Id IN (2,3,6,7)" & vbCrLf

                    ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                    If Not ds1 Is Nothing AndAlso ds1.Tables(0).Rows.Count > 0 Then
                        For Each dr1 As DataRow In ds1.Tables(0).Rows
                            With dr1
                                linha = "|E530"
                                linha &= "|" & .Item("IND_AJ")
                                linha &= "|" & .Item("VL_AJ")
                                linha &= "|" & Funcoes.AlinharDireita(.Item("COD_AJ"), 3, "0")
                                linha &= "|" & .Item("IND_DOC")
                                linha &= "|" & .Item("NUM_DOC")
                                linha &= "|" & .Item("DESCR_AJ")
                                linha &= "|"

                            End With
                            strm.WriteLine(linha)
                            RegistroE530 += 1
                            RegistroGeral += 1
                        Next
                        'End If
                    End If
                End If

                ' Registro E990  - Encerramento do Bloco E
                RegistroE990 += 1

                linha = "|E990"

                If Left(Empresa(0), 8) = "40938762" Then
                    linha &= "|" & RegistroE001 + RegistroE100 + RegistroE110 + RegistroE111 + RegistroE112 + RegistroE113 + RegistroE115 + RegistroE116 + RegistroE200 + RegistroE210 + RegistroE250 + RegistroE300 + RegistroE310 + RegistroE316 + RegistroE500 + RegistroE510 + RegistroE520 + RegistroE530 + RegistroE990
                Else
                    linha &= "|" & RegistroE001 + RegistroE100 + RegistroE110 + RegistroE111 + RegistroE112 + RegistroE113 + RegistroE115 + RegistroE116 + RegistroE500 + RegistroE510 + RegistroE520 + RegistroE530 + RegistroE990
                End If


                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1

                ' Registro G001  - Abertura do Bloco G

                If PeriodoFinal.Year > 2010 Then
                    linha = "|G001"
                    linha &= "|1"       'Indicador de Movimento - 0 - Bloco com Dados Informados
                    '                                             1 - Bloco sem dados informados
                    linha &= "|"

                    strm.WriteLine(linha)
                    RegistroG001 += 1
                    RegistroGeral += 1
                End If

                ' Registro G990  - Encerramento do Bloco G
                If PeriodoFinal.Year > 2010 Then
                    RegistroG990 += 1

                    linha = "|G990"
                    linha &= "|" & RegistroG001 + RegistroG110 + RegistroG125 + RegistroG126 + RegistroG130 + RegistroG140 + RegistroG990
                    linha &= "|"

                    strm.WriteLine(linha)
                    RegistroGeral += 1
                End If


                ' Registro H001  - Abertura do Bloco H

                linha = "|H001"

                'Registro obrigatório à partir de Julho/2012 - furlan - 27/03/2013
                linha &= "|" & IIf(Format(CDate(txtDataFinal.Text), "MM") = 2, 0, 1)      'Indicador de Movimento - 0 - Bloco com Dados Informados
                '                                                                                                   1 - Bloco sem dados informados

                linha &= "|"

                strm.WriteLine(linha)
                RegistroH001 += 1
                RegistroGeral += 1

                ' Registro H005 - Totais do Inventário H
                If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
                    If dsNew.Tables("H005").Rows.Count > 0 Then
                        For Each dr As DataRow In dsNew.Tables("H005").Rows
                            With dr
                                linha = "|H005"
                                linha &= "|3112" & Format(CDate(txtDataFinal.Text), "yyyy") - 1
                                linha &= "|" & .Item("Valor")
                                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2012 Then
                                    linha &= "|01"
                                End If
                                linha &= "|"
                            End With

                            strm.WriteLine(linha)
                            RegistroH005 += 1
                            RegistroGeral += 1
                        Next
                    End If

                    ' Registro H010 - Inventário
                    If dsNew.Tables("H010").Rows.Count > 0 Then
                        For Each dr As DataRow In dsNew.Tables("H010").Rows
                            With dr
                                'If .Item("ValorDoProduto") > 0 Then
                                linha = "|H010"                                                     'REG (Texto fixo contendo "H010")
                                linha &= "|" & .Item("Produto")                                     'COD_ITEM (Código do item)
                                linha &= "|" & RTrim(.Item("Unidade"))                              'UNID (Unidade do item)
                                linha &= "|" & Math.Round(CDec(.Item("Quantidade")), 3).ToString    'QTD (Quantidade do item)
                                linha &= "|" & .Item("Unitario")                                    'VL_UNIT (Valor unitário do item)
                                linha &= "|" & .Item("ValorDoProduto")                              'VL_ITEM (Valor do item)
                                linha &= "|" & .Item("IND_PROP")                                    'IND_PROP (Indicador de propriedade/posse do item)

                                If .Item("IND_PROP") > 0 Then
                                    linha &= "|" & .Item("Cliente_Id") & .Item("Endereco_Id")       'COD_PART (Código do participante)
                                Else
                                    linha &= "|" ' & .Item("Cliente_Id") & .Item("Endereco_Id")  '& .Item("COD_PART")
                                End If

                                linha &= "|" ' & .Item("TXT_COMPL")                                 'TXT_COMPL Descrição complementar.    
                                If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "03189063" Then
                                    linha &= "|101060101" 'COD_CTA (Código da conta analítica contábil debitada/creditada)
                                ElseIf Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                                    linha &= "|10106" 'COD_CTA (Código da conta sintética contábil debitada/creditada) - AJUSTADO CONF. SOLICITAÇÃO JESSICA/LUIZ EM 08/03/2022 
                                Else
                                    linha &= "|" & IIf(IsDBNull(.Item("DebitoMercadoria")) OrElse Trim(.Item("DebitoMercadoria")) = "", .Item("CreditoMercadoria"), .Item("DebitoMercadoria")) 'COD_CTA (Código da conta analítica contábil debitada/creditada)
                                End If

                                linha &= "|" & .Item("ValorDoProduto")                            'VL_ITEM_IR (Valor do item para efeitos do Imposto de Renda)
                                linha &= "|"
                                'End If
                            End With
                            strm.WriteLine(linha)
                            RegistroH010 += 1
                            RegistroGeral += 1
                        Next
                    End If
                End If

                ' Registro H990  - Encerramento do Bloco H
                RegistroH990 += 1

                linha = "|H990"
                linha &= "|" & RegistroH001 + RegistroH005 + RegistroH010 + RegistroH990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1


                'Obrigatório à partir de Janeiro/2016 - Furlan - 11/02/2016
                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2015 Then
                    ' Registro K001  - Abertura do Bloco K
                    RegistroK001 += 1
                    linha = "|K001"

                    If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                        linha &= "|0" '0- Bloco com dados informados   
                    Else
                        linha &= "|1" '1- Bloco sem dados informados 
                    End If

                    linha &= "|"

                    strm.WriteLine(linha)
                    RegistroGeral += 1

                    'Informação sobre o tipo de Leiaute(0-Simplificado, 1-Completo, 2-Restrito aos saldos de Estoque) Obrigatório à partir de 2023. Furlan - 08/02/2023.
                    If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2022 Then
                        RegistroK010 += 1
                        linha = "|K010"

                        linha &= "|2" '0-Simplificado   

                        linha &= "|"

                        strm.WriteLine(linha)
                        RegistroGeral += 1
                    End If

                    If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                        ' Registro K100  - Período de Apuração do ICMS/IPI
                        RegistroK100 += 1
                        linha = "|K100"
                        linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy") 'Data Inicial
                        linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy") 'Data Final
                        linha &= "|"

                        strm.WriteLine(linha)
                        RegistroGeral += 1

                        Sql = "Declare @movimento nvarchar(10);" & vbCrLf & _
                              "Set @movimento = '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf & _
                              "Select sb.Produto_id  AS Produto," & vbCrLf & _
                              "       Prd.Nome as Nome," & vbCrLf & _
                              "       cast(sum(sb.qtdeRazao + sb.Qtdefiscal) as decimal(18,4)) as Quantidade," & vbCrLf & _
                              "       case" & vbCrLf & _
                              "       			when sb.Classificacao = 'NOSSO EM PODER TERCEIROS' then 1" & vbCrLf & _
                              "       			when sb.Classificacao = 'TERCEIROS EM NOSSO PODER' then 2" & vbCrLf & _
                              "                 else 9" & vbCrLf & _
                              "       end as IND_EST," & vbCrLf & _
                              "       sb.Cliente_id AS COD_PART," & vbCrLf & _
                              "       sb.EndCliente_Id AS EndDeposito" & vbCrLf & _
                               "     from (" & vbCrLf & _
                               "   		select nfi.Empresa_Id," & vbCrLf & _
                               "   		       nfi.EndEmpresa_id," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
                               "   				  when SO.ProdutoDeTerceiro =  1                                                                    then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
                               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
                               "   			   end as Classificacao," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                               "   				   then nf.destino" & vbCrLf & _
                               "   				   else nfi.cliente_id" & vbCrLf & _
                               "   			   end Cliente_id," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                               "   				   then nf.Enddestino" & vbCrLf & _
                               "   				   else nfi.Endcliente_id" & vbCrLf & _
                               "   			   end EndCliente_id," & vbCrLf & _
                               "   			   nfi.Produto_Id," & vbCrLf & _
                               "   			   isnull(nf.pedido,0) as Pedido," & vbCrLf & _
                               "                  convert(numeric(18,4),0) as QtdeRazao," & vbCrLf & _
                               "   			   SUM(Case" & vbCrLf & _
                               "   					 When So.Devolucao = 'S'" & vbCrLf & _
                               "   					   then nfi.QuantidadeFiscal * -1" & vbCrLf & _
                               "   					   else nfi.QuantidadeFiscal" & vbCrLf & _
                               "   				   end) QtdeFiscal," & vbCrLf & _
                               "               convert(numeric(18,4),0) as ValorRazao," & vbCrLf & _
                               "   			   SUM(Case" & vbCrLf & _
                               "   					 When So.Devolucao = 'S'" & vbCrLf & _
                               "   					   then nfe.valor * -1" & vbCrLf & _
                               "   					   else nfe.valor" & vbCrLf & _
                               "   				   end) as valorFiscal" & vbCrLf & _
                               "   		  from notasfiscais NF" & vbCrLf & _
                               "   		 inner join Notasfiscaisxitens NFI" & vbCrLf & _
                               "   			on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
                               "   		   and NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                               "   		   and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                               "   		   and NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                               "   		   and NF.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
                               "   		   and NF.Nota_id         = NFI.Nota_id" & vbCrLf & _
                               "   		   and NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                               "   		 Inner Join NotasFiscaisxEncargos Nfe" & vbCrLf & _
                               "   			on Nfe.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
                               "   		   and Nfe.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                               "   		   and Nfe.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                               "   		   and Nfe.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                               "   		   and Nfe.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
                               "   		   and Nfe.Nota_id         = NFI.Nota_id" & vbCrLf & _
                               "   		   and Nfe.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                               "   		   and Nfe.Produto_id      = NFI.Produto_id" & vbCrLf & _
                               "   		   and Nfe.CFOP_Id         = NFI.CFOP_Id" & vbCrLf & _
                               "   		   and Nfe.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
                               "   		   and Nfe.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
                               "   		 Inner join produtos p" & vbCrLf & _
                               "   			on NFI.Produto_id = P.Produto_id" & vbCrLf & _
                               "   		 Inner Join Suboperacoes SO" & vbCrLf & _
                               "   			on SO.Operacao_Id        = NFI.Operacao" & vbCrLf & _
                               "   		   and SO.Suboperacoes_Id    = NFI.Suboperacao" & vbCrLf & _
                               "   		 where (SO.produtodeterceiro = 1 or SO.deposito = 'S')" & vbCrLf & _
                               "   		   and (SO.EstoqueFisico = 'S' or SO.EstoqueFiscal = 'S' or SO.Devolucao = 'S')" & vbCrLf & _
                               "   		   and SO.Operacao_id <> 80" & vbCrLf & _
                               "   		   and nf.movimento  <= @movimento" & vbCrLf & _
                               "           and nf.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                               "   		  group by nfi.Empresa_Id," & vbCrLf & _
                               "   		           nfi.EndEmpresa_id," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
                               "   				  when SO.ProdutoDeTerceiro = 1                                                                     then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
                               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
                               "   			   end," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				 When left(nf.empresa_id, 8) = left(nfi.Cliente_id, 8)" & vbCrLf & _
                               "   				   then nf.destino" & vbCrLf & _
                               "   				   else nfi.cliente_id" & vbCrLf & _
                               "   			   end," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                               "   				   then nf.Enddestino" & vbCrLf & _
                               "   				   else nfi.Endcliente_id" & vbCrLf & _
                               "   			   end," & vbCrLf & _
                               "               nf.pedido," & vbCrLf & _
                               "               nfi.Produto_Id" & vbCrLf & _
                               "        union all" & vbCrLf & _
                               "   		Select R.Empresa_Id," & vbCrLf & _
                               "   			   R.Endempresa_id," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
                               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
                               "   			   end Classificacao," & vbCrLf & _
                               "   			   Case" & vbCrLf & _
                               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
                               "                   then R.Deposito" & vbCrLf & _
                               "                   else R.Cliente_id" & vbCrLf & _
                               "               end Cliente_id," & vbCrLf & _
                               "   			   Case" & vbCrLf & _
                               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
                               "                   then R.EndDeposito" & vbCrLf & _
                               "                   else R.EndCliente_id" & vbCrLf & _
                               "               end EndCliente_id," & vbCrLf & _
                               "   			   R.Produto," & vbCrLf & _
                               "   			   isnull(R.Pedido,0) as Pedido," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then Isnull(R.CreditoQuantidade, 0.0000) - Isnull(R.DebitoQuantidade, 0.0000)" & vbCrLf & _
                               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoQuantidade, 0.0000)  - Isnull(R.CreditoQuantidade, 0.0000)" & vbCrLf & _
                               "   			   end," & vbCrLf & _
                               "   			   convert(numeric(18,4),0)," & vbCrLf & _
                               "   			   case" & vbCrLf & _
                               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then isnull(R.CreditoOficial, 0.00) - isnull(R.DebitoOficial, 0.00)" & vbCrLf & _
                               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoOficial, 0.00)  - isnull(R.CreditoOficial, 0.00)" & vbCrLf & _
                               "   			   end," & vbCrLf & _
                               "               convert(numeric(18,4),0)" & vbCrLf & _
                               "   		  from Razao R" & vbCrLf & _
                               "   		 inner join ClientesxEmpresas CxE" & vbCrLf & _
                               "   			on R.Empresa_id    = CxE.Empresa_Id" & vbCrLf & _
                               "   		   and R.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
                               "   		   and (left(R.Conta_Id,7) = CxE.ContaEstoqueEmNossoPoder or Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) = Left(cxe.ContaestoqueemPoderdeTerceiros,5))" & vbCrLf & _
                               "   		 inner join Clientes C" & vbCrLf & _
                               "   			on R.Cliente_id    = C.Cliente_id" & vbCrLf & _
                               "   		   and R.EndCliente_id = C.Endereco_id" & vbCrLf & _
                               "   		 where lote_id not in (9,10,11)" & vbCrLf & _
                               "           and R.Movimento_Id <= @movimento" & vbCrLf & _
                               "           and year(R.Movimento_Id) > 2010 " & vbCrLf & _
                               "           and R.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                               "   		   and len(isnull(produto,'')) > 0" & vbCrLf & _
                               "   		) as Sb" & vbCrLf & _
                               "     inner join Clientes Emp" & vbCrLf & _
                               "        on Sb.Empresa_id    = Emp.Cliente_id" & vbCrLf & _
                               "       and Sb.EndEmpresa_id = Emp.Endereco_id" & vbCrLf & _
                               "     inner join Clientes Cli" & vbCrLf & _
                               "        on Sb.cliente_id    = Cli.Cliente_id" & vbCrLf & _
                               "       and Sb.EndCliente_id = Cli.Endereco_id" & vbCrLf & _
                               "     inner Join Produtos Prd" & vbCrLf & _
                               "        on Prd.Produto_id = sb.Produto_id" & vbCrLf & _
                               "     Where Prd.TipoDoItem Not in(7,8,9,10,99)" & vbCrLf & _
                               "    Group by sb.Produto_id," & vbCrLf & _
                               "          Prd.Nome," & vbCrLf & _
                               "          sb.Cliente_id," & vbCrLf & _
                               "          sb.EndCliente_Id," & vbCrLf & _
                               "          sb.Classificacao" & vbCrLf & _
                               "    having SUM(sb.QtdeRazao + sb.QtdeFiscal)   <> 0  Or Sum(sb.ValorRazao + sb.ValorFiscal) <> 0" & vbCrLf

                        Sql &= "union all" & vbCrLf

                        ' Registro K200  - Estoque Escriturado
                        Sql &= " SELECT ApuracaoDeCustos.Produto_Id AS Produto, Produtos.Nome, " & vbCrLf & _
                        "    ApuracaoDeCustos.Quantidade,  " & vbCrLf & _
                        "	CASE	 " & vbCrLf & _
                        "		WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 109  " & vbCrLf & _
                        "			THEN 2  " & vbCrLf & _
                        "			ELSE  " & vbCrLf & _
                        "				CASE " & vbCrLf & _
                        "					WHEN  ApuracaoDeCustos.empresa_id <> ApuracaoDeCustos.Deposito_Id  " & vbCrLf & _
                        "					THEN 1 " & vbCrLf & _
                        "					ELSE 0 " & vbCrLf & _
                        "				END " & vbCrLf & _
                        "	END AS IND_EST, " & vbCrLf & _
                        "	CASE " & vbCrLf & _
                        "		WHEN ApuracaoDeCustos.Deposito_Id = ApuracaoDeCustos.Empresa_Id " & vbCrLf & _
                        "			THEN '' " & vbCrLf & _
                        "			ELSE ApuracaoDeCustos.Deposito_Id" & vbCrLf & _
                        "	END AS COD_PART, " & vbCrLf & _
                        "   ISNULL(ApuracaoDeCustos.EndDeposito_Id, 0) AS EndDeposito" & vbCrLf & _
                        "FROM ApuracaoDeCustos   " & vbCrLf & _
                        "       INNER JOIN PlanoDeCustos  " & vbCrLf & _
                        "               ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf & _
                        "       INNER JOIN Produtos  " & vbCrLf & _
                        "               ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
                        "WHERE ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                        "  And ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)   " & vbCrLf & _
                        "  And ApuracaoDeCustos.Quantidade > 0   " & vbCrLf & _
                        "  And ApuracaoDeCustos.ValorDoProduto > 0   " & vbCrLf & _
                        "  And ApuracaoDeCustos.Mes_Id =  " & CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Month & vbCrLf & _
                        "  And ApuracaoDeCustos.Ano_Id = " & CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year & vbCrLf & _
                        "  And  ISNULL(Produtos.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf
                        '"ORDER BY ApuracaoDeCustos.Produto_Id"

                        ds = Banco.ConsultaDataSet(Sql, "ConsultaK200")
                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr As DataRow In ds.Tables(0).Rows
                                With dr
                                    linha = "|K200"
                                    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy") 'Data Final
                                    linha &= "|" & .Item("Produto")
                                    linha &= "|" & Math.Round(CDec(.Item("Quantidade")), 3).ToString
                                    linha &= "|" & .Item("IND_EST")
                                    If .Item("COD_PART").ToString.Length > 0 Then
                                        linha &= "|" & .Item("COD_PART") & .Item("EndDeposito")
                                    Else
                                        linha &= "|" & .Item("COD_PART")
                                    End If
                                    linha &= "|"
                                End With

                                strm.WriteLine(linha)
                                RegistroK200 += 1
                                RegistroGeral += 1
                            Next
                        End If
                    End If

                    'If Left(Empresa(0), 8) = "05366261" AndAlso Format(CDate(txtDataFinal.Text), "MM") = 9 Then
                    '    linha = "|K200"
                    '    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy") 'Data Final
                    '    linha &= "|101020001"
                    '    linha &= "|85135,000"
                    '    linha &= "|2"
                    '    linha &= "|620099409970"
                    '    linha &= "|"

                    '    strm.WriteLine(linha)
                    '    RegistroK200 += 1
                    '    RegistroGeral += 1

                    '    linha = "|K200"
                    '    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy") 'Data Final
                    '    linha &= "|101020001"
                    '    linha &= "|50677,000"
                    '    linha &= "|2"
                    '    linha &= "|007532199800"
                    '    linha &= "|"

                    '    strm.WriteLine(linha)
                    '    RegistroK200 += 1
                    '    RegistroGeral += 1

                    '    linha = "|K200"
                    '    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy") 'Data Final
                    '    linha &= "|101020001"
                    '    linha &= "|110373,000"
                    '    linha &= "|2"
                    '    linha &= "|003354609680"
                    '    linha &= "|"

                    '    strm.WriteLine(linha)
                    '    RegistroK200 += 1
                    '    RegistroGeral += 1
                    'End If


                    ' Registro K990  - Encerramento do Bloco K
                    RegistroK990 += 1
                    linha = "|K990"

                    If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2022 Then
                        linha &= "|" & RegistroK001 + RegistroK010 + RegistroK100 + RegistroK200 + RegistroK990
                    ElseIf CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                        linha &= "|" & RegistroK001 + RegistroK100 + RegistroK200 + RegistroK990
                    Else
                        linha &= "|" & RegistroK001 + RegistroK990
                    End If

                    linha &= "|"

                    strm.WriteLine(linha)
                    RegistroGeral += 1
                End If


                ' Registro 1001  - Abertura do Bloco 1
                Registro1001 += 1
                linha = "|1001"

                'Registro obrigatório à partir de Julho/2012 - furlan - 19/03/2013

                'If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2012 OrElse _
                '   (CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2011 And CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Month > 6) Then
                '    linha &= "|0"
                'Else
                '    linha &= "|1"
                'End If

                If PeriodoFinal.Year > 2012 OrElse (PeriodoFinal.Year > 2011 AndAlso PeriodoFinal.Month > 6) Then
                    linha &= "|0"
                Else
                    linha &= "|1"
                End If

                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1

                ' Registro 1010 - Registro de Informações sobre exportações

                Dim ds1100 As DataSet = ConsultaRegistro1100()
                Dim ds1400 As DataSet = ConsultaRegistro1400()

                If PeriodoFinal.Year > 2012 OrElse (PeriodoFinal.Year > 2011 AndAlso PeriodoFinal.Month > 6) Then
                    Registro1010 += 1
                    linha = "|1010"
                    linha &= "|" & IIf(ds1100.Tables(0).Rows.Count > 0, "S", "N") 'S - Com Dados  /  N - Sem Dados
                    linha &= "|N|N|N"
                    linha &= "|" & IIf(ds1400.Tables(0).Rows.Count > 0, "S", "N")  'S - Com Dados  /  N - Sem Dados

                    If PeriodoFinal.Year > 2019 Then
                        linha &= "|N|N|N|N|N|N|N|N|"
                    ElseIf PeriodoFinal.Year > 2018 Then
                        linha &= "|N|N|N|N|N|N|N|"
                    Else
                        linha &= "|N|N|N|N|"
                    End If

                    strm.WriteLine(linha)
                    RegistroGeral += 1

                    ' Registro 1100 - Registro de Informações sobre exportações

                    CompoeRegistro1100(ds1100, strm, linha, Registro1100, Registro1105, Registro1110, RegistroGeral)

                    '' Registro 1400 - Informações sobre Valores Agregados
                    CompoeRegistro1400(ds1400, strm, linha, Registro1400, RegistroGeral)
                End If

                ' Registro 1990  - Encerramento do Bloco 1
                Registro1990 += 1

                linha = "|1990"
                linha &= "|" & Registro1001 + Registro1010 + Registro1100 + Registro1105 + Registro1110 + Registro1400 + Registro1990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1

                ' Registro 9001  - Abertura do Bloco 9
                linha = "|9001"
                linha &= "|0"       'Indicador de Movimento - 0 - Bloco com Dados Informados
                '                                             1 - Bloco sem dados informados
                linha &= "|"

                strm.WriteLine(linha)
                Registro9001 += 1
                RegistroGeral += 1


                ' Registro 9900  - Registros Do Arquivo
                linha = "|9900|0000|" & Registro0000 & "|" & vbCrLf
                linha &= "|9900|0001|" & Registro0001 & "|" & vbCrLf

                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2019 Then
                    linha &= "|9900|0002|" & Registro0002 & "|" & vbCrLf
                    Registro9900 += 1
                    RegistroGeral += 1
                End If

                linha &= "|9900|0005|" & Registro0005 & "|" & vbCrLf

                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2023 Then
                    linha &= "|9900|0015|" & Registro0015 & "|" & vbCrLf
                    Registro9900 += 1
                    RegistroGeral += 1
                End If

                linha &= "|9900|0100|" & Registro0100 & "|" & vbCrLf
                linha &= "|9900|0150|" & Registro0150 & "|" & vbCrLf
                linha &= "|9900|0190|" & Registro0190 & "|" & vbCrLf
                linha &= "|9900|0200|" & Registro0200 & "|" & vbCrLf

                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                    linha &= "|9900|0220|" & Registro0220 & "|" & vbCrLf
                    Registro9900 += 1
                    RegistroGeral += 1
                End If

                linha &= "|9900|0400|" & Registro0400 & "|" & vbCrLf
                linha &= "|9900|0450|" & Registro0450 & "|" & vbCrLf
                linha &= "|9900|0460|" & Registro0460 & "|" & vbCrLf
                linha &= "|9900|0990|" & Registro0990 & "|" & vbCrLf

                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2018 Then
                    linha &= "|9900|B001|" & RegistroB001 & "|" & vbCrLf
                    linha &= "|9900|B990|" & RegistroB990 & "|" & vbCrLf
                    Registro9900 += 2
                    RegistroGeral += 2
                End If

                linha &= "|9900|C001|" & RegistroC001 & "|" & vbCrLf
                linha &= "|9900|C100|" & RegistroC100 & "|" & vbCrLf

                If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2024 Then
                    linha &= "|9900|C101|" & RegistroC101 & "|" & vbCrLf
                    Registro9900 += 1
                    RegistroGeral += 1
                End If

                linha &= "|9900|C170|" & RegistroC170 & "|" & vbCrLf
                linha &= "|9900|C172|" & RegistroC172 & "|" & vbCrLf
                linha &= "|9900|C190|" & RegistroC190 & "|" & vbCrLf
                'linha &= "|9900|C195|" & RegistroC195 & "|" & vbCrLf
                'linha &= "|9900|C197|" & RegistroC197 & "|" & vbCrLf
                linha &= "|9900|C500|" & RegistroC500 & "|" & vbCrLf
                linha &= "|9900|C510|" & RegistroC510 & "|" & vbCrLf
                linha &= "|9900|C590|" & RegistroC590 & "|" & vbCrLf
                linha &= "|9900|C990|" & RegistroC990 & "|" & vbCrLf
                linha &= "|9900|D001|" & RegistroD001 & "|" & vbCrLf
                linha &= "|9900|D100|" & RegistroD100 & "|" & vbCrLf
                linha &= "|9900|D190|" & RegistroD190 & "|" & vbCrLf
                linha &= "|9900|D500|" & RegistroD500 & "|" & vbCrLf
                linha &= "|9900|D590|" & RegistroD590 & "|" & vbCrLf
                linha &= "|9900|D990|" & RegistroD990 & "|" & vbCrLf
                linha &= "|9900|E001|" & RegistroE001 & "|" & vbCrLf
                linha &= "|9900|E100|" & RegistroE100 & "|" & vbCrLf
                linha &= "|9900|E110|" & RegistroE110 & "|" & vbCrLf
                linha &= "|9900|E111|" & RegistroE111 & "|" & vbCrLf
                linha &= "|9900|E112|" & RegistroE112 & "|" & vbCrLf
                linha &= "|9900|E113|" & RegistroE113 & "|" & vbCrLf
                linha &= "|9900|E115|" & RegistroE115 & "|" & vbCrLf
                linha &= "|9900|E116|" & RegistroE116 & "|" & vbCrLf

                If Left(Empresa(0), 8) = "40938762" Then
                    linha &= "|9900|E200|" & RegistroE200 & "|" & vbCrLf
                    linha &= "|9900|E210|" & RegistroE210 & "|" & vbCrLf
                    linha &= "|9900|E250|" & RegistroE250 & "|" & vbCrLf
                    'linha &= "|9900|E220|" & RegistroE220 & "|" & vbCrLf
                    'linha &= "|9900|E240|" & RegistroE240 & "|" & vbCrLf
                    linha &= "|9900|E300|" & RegistroE300 & "|" & vbCrLf
                    linha &= "|9900|E310|" & RegistroE310 & "|" & vbCrLf
                    linha &= "|9900|E316|" & RegistroE316 & "|" & vbCrLf
                    Registro9900 += 6
                    RegistroGeral += 6
                End If

                linha &= "|9900|E500|" & RegistroE500 & "|" & vbCrLf
                linha &= "|9900|E510|" & RegistroE510 & "|" & vbCrLf
                linha &= "|9900|E520|" & RegistroE520 & "|" & vbCrLf
                linha &= "|9900|E530|" & RegistroE530 & "|" & vbCrLf
                linha &= "|9900|E990|" & RegistroE990 & "|" & vbCrLf
                Registro9900 += 39
                RegistroGeral += 39

                If PeriodoFinal.Year > 2010 Then
                    linha &= "|9900|G001|" & RegistroG001 & "|" & vbCrLf
                    linha &= "|9900|G990|" & RegistroG990 & "|" & vbCrLf
                    Registro9900 += 2
                    RegistroGeral += 2
                End If

                linha &= "|9900|H001|" & RegistroH001 & "|" & vbCrLf
                linha &= "|9900|H005|" & RegistroH005 & "|" & vbCrLf
                linha &= "|9900|H010|" & RegistroH010 & "|" & vbCrLf
                linha &= "|9900|H990|" & RegistroH990 & "|" & vbCrLf

                If PeriodoFinal.Year > 2015 Then
                    If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2022 Then
                        linha &= "|9900|K001|" & RegistroK001 & "|" & vbCrLf
                        linha &= "|9900|K010|" & RegistroK010 & "|" & vbCrLf
                        linha &= "|9900|K100|" & RegistroK100 & "|" & vbCrLf
                        linha &= "|9900|K200|" & RegistroK200 & "|" & vbCrLf
                        linha &= "|9900|K990|" & RegistroK990 & "|" & vbCrLf
                        Registro9900 += 5
                        RegistroGeral += 5
                    ElseIf CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2017 Then
                        linha &= "|9900|K001|" & RegistroK001 & "|" & vbCrLf
                        linha &= "|9900|K100|" & RegistroK100 & "|" & vbCrLf
                        linha &= "|9900|K200|" & RegistroK200 & "|" & vbCrLf
                        linha &= "|9900|K990|" & RegistroK990 & "|" & vbCrLf
                        Registro9900 += 4
                        RegistroGeral += 4
                    Else
                        linha &= "|9900|K001|" & RegistroK001 & "|" & vbCrLf
                        linha &= "|9900|K990|" & RegistroK990 & "|" & vbCrLf
                        Registro9900 += 2
                        RegistroGeral += 2
                    End If
                End If

                linha &= "|9900|1001|" & Registro1001 & "|" & vbCrLf

                'If ds.Tables(0).Rows.Count > 0 Or Format(CDate(txtDataFinal.Text), "MMyyyy") > 62012 Or (ds1400.Tables(0).Rows.Count > 0 And Format(CDate(txtDataFinal.Text), "MMyyyy") > 62012) Then linha &= "|9900|1010|" & Registro1010 & "|" & vbCrLf
                'Registro obrigatório à partir de Julho/2012 - furlan - 19/03/2013
                'If CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2012 OrElse _
                '   (CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Year > 2011 And CDate(Format(PeriodoFinal, "yyyy/MM/dd")).Month > 6) Then
                If PeriodoFinal.Year > 2012 OrElse (PeriodoFinal.Year > 2012 AndAlso PeriodoFinal.Month > 6) Then
                    linha &= "|9900|1010|" & Registro1010 & "|" & vbCrLf
                End If

                linha &= "|9900|1100|" & Registro1100 & "|" & vbCrLf
                linha &= "|9900|1105|" & Registro1105 & "|" & vbCrLf
                linha &= "|9900|1110|" & Registro1110 & "|" & vbCrLf
                linha &= "|9900|1400|" & Registro1400 & "|" & vbCrLf
                linha &= "|9900|1990|" & Registro1990 & "|" & vbCrLf
                linha &= "|9900|9001|" & Registro9001 & "|" & vbCrLf
                linha &= "|9900|9900|" & Registro9900 + IIf(Registro1010 > 0, 15, 14) & "|" & vbCrLf
                linha &= "|9900|9990|1|" & vbCrLf
                linha &= "|9900|9999|1|"

                If Registro1010 > 0 Then
                    Registro9900 += 15
                    RegistroGeral += 15
                Else
                    Registro9900 += 14
                    RegistroGeral += 14
                End If

                strm.WriteLine(linha)

                ' Registro 9990  - Encerramento do Bloco 9
                Registro9990 += 1
                RegistroGeral += 1

                linha = "|9990"
                linha &= "|" & Registro9001 + Registro9900 + Registro9990 + 1
                linha &= "|"

                strm.WriteLine(linha)

                ' Registro 9999  - Encerramento do Arquivo Digital
                Registro9999 += 1
                RegistroGeral += 1
                linha = "|9999"
                linha &= "|" & RegistroGeral
                linha &= "|"

                strm.WriteLine(linha)
                imdDownload.Visible = True
                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                strm.Close()
                strm.Dispose()
            End Try
        End Using
    End Sub

    Private Function ConsultaRegistro1100() As DataSet
        Sql = " /*REGISTRO1100 - Registro de Informações sobre Exportação*/" & vbCrLf & _
              " SELECT NF.Nota_id AS Nota, ISNULL(ME.Memorando_Id, '') Memorando_Id," & vbCrLf & _
              "        ISNULL(ME.TipoDocumento,0) AS TipoDocumento," & vbCrLf & _
              "        ISNULL(NFEx.NrDespachoExp, ME.NrDespachoExp) AS NrDespachoExp," & vbCrLf & _
              "        ISNULL(ISNULL(NFEx.DataDespachoExp, ME.DataDespachoExp), '') as DataDespachoExp," & vbCrLf & _
              "        CASE " & vbCrLf & _
              "  	     WHEN ISNULL(Isnull(NF.EntradaSaida_Id, ME.EntradaSaida),'E') = 'S' " & vbCrLf & _
              "  		   THEN 0 " & vbCrLf & _
              "  		 ELSE 1 " & vbCrLf & _
              "        END AS NatExp," & vbCrLf & _
              "        ISNULL(MERe.RegistroDeExportacao_Id, NRE.RegistroDeExportacao_Id) AS RegistroDeExportacao_Id," & vbCrLf & _
              "  	   ISNULL(MERe.DataExportacao, NRE.DataRegistroDeExportacao) as DataExportacao," & vbCrLf & _
              "  	   ISNULL(ISNULL(MECe.ConhecimentoDeEmbarque_Id, NCE.ConhecimentoDeEmbarque_Id), '') AS ConhecimentoDeEmbarque_Id," & vbCrLf & _
              "        ISNULL(ISNULL(MECe.DataConhecimento, NCE.DataConhecimento), '') AS DataConhecimento," & vbCrLf & _
              "        ISNULL(ISNULL(ME.DataAverbacao, NFEx.DataAverbacao), '') AS DataAverbacao," & vbCrLf & _
              "        ISNULL(ISNULL(MECe.TipoConhecimento, NCE.TipoConhecimento),'') AS TipoConhecimento,              " & vbCrLf & _
              "        --ISNULL(ME.PaisDestino, NFEx.PaisDestino) AS PaisDestino" & vbCrLf & _
              "        CONVERT(INT, LEFT(REPLICATE('0', 4-LEN(ISNULL(ME.PaisDestino, NFEx.PaisDestino))) +CONVERT(VARCHAR(5), ISNULL(ME.PaisDestino, NFEx.PaisDestino)),3)) AS PaisDestino" & vbCrLf & _
              "   FROM NotasFiscais AS NF " & vbCrLf & _
              "   LEFT JOIN NotaFiscalXExportacao AS NFEx " & vbCrLf & _
              "     ON NF.Empresa_Id      = NFEx.Empresa_Id " & vbCrLf & _
              "    AND NF.EndEmpresa_Id   = NFEx.EndEmpresa_Id " & vbCrLf & _
              "    AND NF.Cliente_Id      = NFEx.Cliente_Id " & vbCrLf & _
              "    AND NF.EndCliente_Id   = NFEx.EndCliente_Id " & vbCrLf & _
              "    AND NF.EntradaSaida_Id = NFEx.EntradaSaida_Id " & vbCrLf & _
              "    AND NF.Serie_Id        = NFEx.Serie_Id " & vbCrLf & _
              "    AND NF.Nota_Id         = NFEx.Nota_Id " & vbCrLf & _
              "   LEFT JOIN NotaFiscalXRE NRE" & vbCrLf & _
              "     ON NF.Empresa_Id      = NRE.Empresa_Id" & vbCrLf & _
              "    AND NF.EndEmpresa_Id   = NRE.EndEmpresa_Id" & vbCrLf & _
              "    AND NF.Cliente_Id      = NRE.Cliente_Id" & vbCrLf & _
              "    AND NF.EndCliente_Id   = NRE.EndCliente_Id" & vbCrLf & _
              "    AND NF.EntradaSaida_Id = NRE.EntradaSaida_Id" & vbCrLf & _
              "    AND NF.Serie_Id        = NRE.Serie_Id" & vbCrLf & _
              "    AND NF.Nota_Id         = NRE.Nota_Id" & vbCrLf & _
              "   LEFT JOIN NotaFiscalXCE NCE" & vbCrLf & _
              "     ON NF.Empresa_Id      = NCE.Empresa_Id" & vbCrLf & _
              "    AND NF.EndEmpresa_Id   = NCE.EndEmpresa_Id" & vbCrLf & _
              "    AND NF.Cliente_Id      = NCE.Cliente_Id" & vbCrLf & _
              "    AND NF.EndCliente_Id   = NRE.EndCliente_Id" & vbCrLf & _
              "    AND NF.EntradaSaida_Id = NCE.EntradaSaida_Id" & vbCrLf & _
              "    AND NF.Serie_Id        = NCE.Serie_Id" & vbCrLf & _
              "    AND NF.Nota_Id         = NCE.Nota_Id" & vbCrLf & _
              "   LEFT JOIN MemorandoDeExportacao AS ME " & vbCrLf & _
              "     ON NF.Empresa_Id      = ME.Empresa " & vbCrLf & _
              "    AND NF.EndEmpresa_Id   = ME.EndEmpresa " & vbCrLf & _
              "    AND NF.Cliente_Id      = ME.Cliente " & vbCrLf & _
              "    AND NF.EndCliente_Id   = ME.EndCliente" & vbCrLf & _
              "    AND NF.EntradaSaida_Id = ME.EntradaSaida " & vbCrLf & _
              "    AND NF.Serie_Id        = ME.Serie " & vbCrLf & _
              "    AND NF.Nota_Id         = ME.Nota" & vbCrLf & _
              "   LEFT JOIN MemorandoDeExportacaoXConhecimentoDeEmbarque MECe" & vbCrLf & _
              "     ON ME.EmpresaMemorando_Id    = MECe.EmpresaMemorando_Id" & vbCrLf & _
              "    AND ME.EndEmpresaMemorando_Id = MECe.EndEmpresaMemorando_Id" & vbCrLf & _
              "    AND ME.Memorando_Id           = MECe.Memorando_Id" & vbCrLf & _
              "   LEFT JOIN MemorandoDeExportacaoXRegistroDeExportacao MERe" & vbCrLf & _
              "     ON ME.EmpresaMemorando_Id    = MERe.EmpresaMemorando_Id" & vbCrLf & _
              "    AND ME.EndEmpresaMemorando_Id = MERe.EndEmpresaMemorando_Id" & vbCrLf & _
              "    AND ME.Memorando_Id           = MERe.Memorando_Id" & vbCrLf & _
              "  WHERE ISNULL(ME.EmpresaMemorando_Id, NFEx.Empresa_Id) = '" & Empresa(0) & "'" & vbCrLf & _
              "    AND ISNULL(ME.DataAverbacao, NFEx.DataAverbacao) BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf & _
              "    AND NF.Situacao = 1 " & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Function CompoeRegistro1100(ByRef ds As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef Registro1100 As Integer, ByRef Registro1105 As Integer, ByRef Registro1110 As Integer, ByRef registroGeral As Integer) As String
        Dim retorno As String = "N"
        Dim Memorando As String = ""
        Dim NatExp As String = ""
        Dim Nota As Integer = 0
        Dim NotaAnterior As Integer = 0
        Dim msg As String = ""

        Try
            If ds.Tables(0).Rows.Count > 0 Then

                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        If IsDBNull(.Item("RegistroDeExportacao_Id")) OrElse IsDBNull(.Item("NrDespachoExp")) Then
                            msg &= .Item("Nota") & ", "
                        End If
                    End With
                Next

                If Not String.IsNullOrWhiteSpace(msg) Then
                    Throw New Exception("A(s) seguinte(s) NF(s): " & msg.Substring(0, msg.Length - 2) & " está(ão) sem informações de exportação.")
                End If

                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr

                        If IsDBNull(.Item("RegistroDeExportacao_Id")) OrElse IsDBNull(.Item("NrDespachoExp")) Then
                            Throw New Exception("A NF: " & .Item("Nota") & " está sem informações de exportação.")
                        End If

                        Memorando = .Item("Memorando_Id")
                        NatExp = .Item("NatExp")
                        Nota = .Item("Nota")

                        linha = "|1100"
                        linha &= "|" & .Item("TipoDocumento")                               'TipoDocumento
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(RTrim(.Item("NrDespachoExp"))).ToString.Replace("/", "") 'NrDespachoExp
                        linha &= "|" & Format(CDate(.Item("DataDespachoExp")), "ddMMyyyy")  'DataDespachoExp
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(RTrim(.Item("NatExp"))).Replace("-", "")  'NatExp
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(RTrim(.Item("RegistroDeExportacao_Id"))).Replace("-", "").Replace("/", "") 'RegistroDeExportacao
                        linha &= "|" & Format(CDate(.Item("DataExportacao")), "ddMMyyyy")   'DataExportacao
                        linha &= "|" & .Item("ConhecimentoDeEmbarque_Id")                   'ConhecimentoDeEmbarque
                        linha &= "|" & Format(CDate(.Item("DataConhecimento")), "ddMMyyyy") 'DataConhecimento
                        linha &= "|" & Format(CDate(.Item("DataAverbacao")), "ddMMyyyy")    'DataAverbacao
                        linha &= "|" & .Item("TipoConhecimento")                            'TipoConhecimento
                        linha &= "|" & Left(.Item("PaisDestino"), 3)                        'PaisDestino
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)

                    registroGeral += 1
                    Registro1100 += 1


                    ' Registro 1105 - Documentos Fiscais de Exportação

                    Sql = "/*Registro 1105 - Documentos Fiscais de Exportação*/" & vbCrLf & _
                          "  SELECT NF.Eletronica," & vbCrLf & _
                          "         OE.CodigoFiscal," & vbCrLf & _
                          "         NF.Serie_Id," & vbCrLf & _
                          "         NF.Nota_Id," & vbCrLf & _
                          "         NFERealizadas.ChaveNfe," & vbCrLf & _
                          "         isnull(NF.DataDaNota,'') as DataDaNota," & vbCrLf & _
                          "         P.Produto_Id" & vbCrLf & _
                          "    FROM NotasFiscais NF" & vbCrLf & _
                          "   INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                          "      ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                          "     AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                          "     AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                          "     AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                          "     AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                          "     AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                          "     AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                          "   Inner Join OperacaoXEstado OE" & vbCrLf & _
                          "      on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf & _
                          "    LEFT JOIN ProdutosAgrupados PA" & vbCrLf & _
                          "      ON PA.ProdutoAgrupado_Id = NFxI.Produto_Id" & vbCrLf & _
                          "   INNER JOIN Produtos P" & vbCrLf & _
                          "      ON P.Produto_Id = isnull(PA.Produto_id,NFxI.Produto_Id)" & vbCrLf & _
                          "   INNER JOIN NFERealizadas" & vbCrLf & _
                          "      ON NF.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                          "     AND NF.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                          "     AND NF.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                          "     AND NF.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                          "     AND NF.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                          "     AND NF.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                          "     AND NF.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                          "    LEFT JOIN MemorandoDeExportacao Me" & vbCrLf & _
                          "      ON Me.Empresa      = NF.Empresa_Id" & vbCrLf & _
                          "     AND Me.EndEmpresa   = NF.EndEmpresa_Id" & vbCrLf & _
                          "     AND Me.Cliente      = NF.Cliente_Id" & vbCrLf & _
                          "     AND Me.EndCliente   = NF.EndCliente_Id" & vbCrLf & _
                          "     AND Me.EntradaSaida = NF.EntradaSaida_Id" & vbCrLf & _
                          "     AND Me.Nota         = NF.Nota_Id" & vbCrLf & _
                          "     AND Me.Serie        = NF.Serie_Id" & vbCrLf & _
                          "   WHERE NFxI.QuantidadeFiscal > 0" & vbCrLf & _
                          "     AND NF.Operacao           > 0" & vbCrLf & _
                          "     AND NF.Situacao not in(9,10)" & vbCrLf & _
                          "     AND ISNULL(ME.DataMemorando, NF.Movimento) BETWEEN '" & PeriodoInicial.ToSqlDate() & "'" & vbCrLf & _
                          "     AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf & _
                          "     AND NF.Empresa_Id      = '" & Empresa(0) & "'" & vbCrLf & _
                          "     AND NF.EstadoDoCliente = 'EX'" & vbCrLf & _
                          "     AND ISNULL(Me.Memorando_Id, '') = '" & Memorando & "'" & vbCrLf & _
                          "     AND NF.Nota_id = " & Nota & vbCrLf
                    'If Not String.IsNullOrWhiteSpace(Memorando) Then
                    'Sql &= "     AND Me.Memorando_Id= '" & Memorando & "' "
                    'End If

                    Dim ds1105 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds1105.Tables(0).Rows.Count > 0 Then
                        For Each dr2 As DataRow In ds1105.Tables(0).Rows
                            With dr2
                                linha = "|1105"
                                If .Item("Eletronica") = "N" Then
                                    linha &= "|01"
                                Else
                                    linha &= "|55"                      'Código do modelo do documento fiscal, conforme a Tabela 4.1.1
                                End If
                                linha &= "|" & .Item("Serie_Id")        'Serie
                                linha &= "|" & .Item("Nota_Id")         'Num. nota
                                linha &= "|" & .Item("ChaveNfe")        'ChaveNfe
                                linha &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy") 'Data Emissao
                                linha &= "|" & .Item("Produto_Id")      'Cod. do Produto
                                linha &= "|"
                            End With

                            strm.WriteLine(linha)
                            Registro1105 += 1
                            registroGeral += 1
                        Next
                    End If


                    ' Registro 1110 - Operações De Exportação Indireta - Mercadorias De Terceiros
                    Sql = "/*Registro 1110 - Operações De Exportação Indireta - Mercadorias De Terceiros*/  " & vbCrLf & _
                          "SELECT NF.Cliente_Id," & vbCrLf & _
                          "       NF.EndCliente_Id," & vbCrLf & _
                          "       OE.CodigoFiscal," & vbCrLf & _
                          "       NF.Serie_Id," & vbCrLf & _
                          "       NF.Nota_Id," & vbCrLf & _
                          "       NfeR.ChaveNfe," & vbCrLf & _
                          "       isnull(NF.DataDaNota,'') as DataDaNota," & vbCrLf & _
                          "       MeXNf.Memorando_Id, " & vbCrLf & _
                          "       Sum(Quantidade) as  Quantidade," & vbCrLf & _
                          "       P.Unidade" & vbCrLf & _
                          "  FROM NotasFiscais NF " & vbCrLf & _
                          "  LEFT JOIN MemorandoDeExportacaoXNotaFiscal MeXNF" & vbCrLf & _
                          "    ON MeXNF.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
                          "   AND MeXNF.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
                          "   AND MeXNF.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
                          "   AND MeXNF.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
                          "   AND MeXNF.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
                          "   AND MeXNF.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                          "   AND MeXNF.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                          " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                          "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                          "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                          "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                          "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                          "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                          "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                          "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                          " Inner Join OperacaoXEstado OE" & vbCrLf & _
                          "    on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf & _
                          " INNER JOIN Produtos P" & vbCrLf & _
                          "    ON P.Produto_Id = NFxI.Produto_Id" & vbCrLf & _
                          " INNER JOIN NFERealizadas NfeR" & vbCrLf & _
                          "    ON NF.Empresa_Id      = NfeR.Empresa_Id" & vbCrLf & _
                          "   AND NF.EndEmpresa_Id   = NfeR.EndEmpresa_Id" & vbCrLf & _
                          "   AND NF.Cliente_Id      = NfeR.Cliente_Id" & vbCrLf & _
                          "   AND NF.EndCliente_Id   = NfeR.EndCliente_Id" & vbCrLf & _
                          "   AND NF.EntradaSaida_Id = NfeR.EntradaSaida_Id" & vbCrLf & _
                          "   AND NF.Serie_Id        = NfeR.Serie_Id" & vbCrLf & _
                          "   AND NF.Nota_Id         = NfeR.Nota_Id" & vbCrLf & _
                          "  LEFT JOIN MemorandoDeExportacao Me" & vbCrLf & _
                          "    ON MeXNF.EmpresaMemorando_Id    = Me.EmpresaMemorando_Id " & vbCrLf & _
                          "   AND MeXNF.EndEmpresaMemorando_Id = Me.EndEmpresaMemorando_Id " & vbCrLf & _
                          "   AND MeXNF.Memorando_Id           = Me.Memorando_Id " & vbCrLf & _
                          " WHERE NFxI.QuantidadeFiscal > 0" & vbCrLf & _
                          "   AND NF.Operacao           > 0" & vbCrLf & _
                          "   AND NF.Situacao not in(9,10) " & vbCrLf & _
                          "   AND MeXNF.Memorando_Id    ='" & Memorando & "'  " & vbCrLf & _
                          "   AND NF.Empresa_Id         ='" & Empresa(0) & "'     " & vbCrLf & _
                          "   AND MeXNF.EntradaSaida_Id ='E'" & vbCrLf & _
                          "   AND Me.Nota            = " & Nota & vbCrLf & _
                          " Group By NF.Cliente_Id," & vbCrLf & _
                          "          NF.EndCliente_Id," & vbCrLf & _
                          "          OE.CodigoFiscal," & vbCrLf & _
                          "          NF.Serie_Id," & vbCrLf & _
                          "          NF.Nota_Id," & vbCrLf & _
                          "          NfeR.ChaveNfe," & vbCrLf & _
                          "          isnull(NF.DataDaNota,'')," & vbCrLf & _
                          "          MeXNF.Memorando_Id," & vbCrLf & _
                          "          P.Unidade" & vbCrLf

                    Dim ds1110 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds1110.Tables(0).Rows.Count > 0 Then
                        For Each dr2 As DataRow In ds1110.Tables(0).Rows
                            With dr2
                                linha = "|1110"
                                linha &= "|" & Trim(.Item("Cliente_Id")) & Trim(.Item("EndCliente_Id")) 'Fonecedor Merc. exportacao
                                'Código do modelo do documento fiscal, conforme a Tabela 4.1.1 
                                ' 01 - Nota Fiscal
                                ' 04 - Nota fiscal de produtor
                                ' 1B - Nota fiscal avulsa
                                ' 55 - Nota fiscal Eletrônica
                                If .Item("CodigoFiscal") < 5000 Then
                                    linha &= "|55"
                                Else
                                    linha &= "|55"
                                End If

                                linha &= "|" & .Item("Serie_Id")        'Serie
                                linha &= "|" & .Item("Nota_Id")         'Num. nota
                                linha &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy") 'Data Emissao
                                linha &= "|" & .Item("ChaveNfe")        'ChaveNfe
                                linha &= "|" & .Item("Memorando_Id")    'Num Memorando
                                linha &= "|" & Math.Round(CDec(.Item("Quantidade")), 3).ToString 'Quantidade do Produto Exp.
                                linha &= "|" & Trim(.Item("Unidade"))   'Unidade Produto
                                linha &= "|"
                            End With

                            strm.WriteLine(linha)
                            Registro1110 += 1
                            registroGeral += 1
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return retorno
    End Function

    Private Function ConsultaRegistro1400() As DataSet
        Sql = " SELECT CodigoDoMunicipio,EstadoIbge,Produto_Id,Sum(ValorTotal)AS ValorTotal" & vbCrLf & _
              "   FROM ( " & vbCrLf & _
              "         SELECT Clientes.CodigoDoMunicipio,Municipios.municipio_id, Municipios.EstadoIbge, " & vbCrLf & _
              "                NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,  " & vbCrLf & _
              "                NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscais.Situacao," & vbCrLf & _
              "                NotasFiscais.LocalEmbarque, NotasFiscais.EndLocalEmbarque,NFxI.Produto_Id,  " & vbCrLf & _
              "                ISNULL ((SELECT SUM(Valor) AS Valor " & vbCrLf & _
              "                           FROM NotasFiscaisXEncargos " & vbCrLf & _
              "                          WHERE NotasFiscais.Nota_Id         = Nota_Id" & vbCrLf & _
              "                            AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id" & vbCrLf & _
              "                            AND NotasFiscais.Empresa_Id      = Empresa_Id" & vbCrLf & _
              "                            AND NotasFiscais.EndEmpresa_Id   = EndEmpresa_Id" & vbCrLf & _
              "                            AND NotasFiscais.Cliente_Id      = Cliente_Id" & vbCrLf & _
              "                            AND NotasFiscais.EndCliente_Id   = EndCliente_Id" & vbCrLf & _
              "                            AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id " & vbCrLf & _
              "                            AND NotasFiscais.Serie_Id        = Serie_Id" & vbCrLf & _
              "                            AND Encargo_Id                   = 'PRODUTO'), 0) AS ValorTotal  " & vbCrLf & _
              "           FROM NotasFiscais  " & vbCrLf & _
              "          INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "             ON NotasFiscais.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "            AND NotasFiscais.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "            AND NotasFiscais.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "            AND NotasFiscais.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "            AND NotasFiscais.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "            AND NotasFiscais.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "            AND NotasFiscais.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              "           Left JOIN Clientes " & vbCrLf & _
              "             ON Clientes.Cliente_Id  = NotasFiscais.LocalEmbarque " & vbCrLf & _
              "            AND Clientes.Endereco_Id = NotasFiscais.EndLocalEmbarque  " & vbCrLf & _
              "           Left JOIN Estados  " & vbCrLf & _
              "             ON Clientes.Estado = Estados.Estado_Id" & vbCrLf & _
              "           Left JOIN Municipios  " & vbCrLf & _
              "             ON Clientes.Estado            = Municipios.Estado_id" & vbCrLf & _
              "            AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id" & vbCrLf & _
              "            AND Clientes.cidade            = Municipios.municipio_Id" & vbCrLf & _
              "          WHERE NotasFiscais.Empresa_Id      = '" & Empresa(0) & "'" & vbCrLf & _
              "            AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf & _
              "            AND Notasfiscais.Situacao        = 1" & vbCrLf & _
              "            AND Notasfiscais.operacao        > 0" & vbCrLf & _
              "            AND NotasFiscais.TipoDeDocumento = 2" & vbCrLf & _
              "            AND NotasFiscais.NossaEmissao    = 'S'" & vbCrLf & _
              "         ) Consulta " & vbCrLf & _
              "    GROUP BY CodigoDoMunicipio,EstadoIbge,Produto_Id " & vbCrLf
        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Function CompoeRegistro1400(ByRef ds As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef Registro1400 As Integer, ByRef registroGeral As Integer) As String
        Dim retorno As String = "N"
        Try
            If ds.Tables(0).Rows.Count > 0 Then
                retorno = "S"
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|1400"
                        linha &= "|" & Trim(.Item("Produto_Id"))             'TipoDocumento
                        linha &= "|" & .Item("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0") 'CodMunicipio (UF+Municipio)
                        linha &= "|" & .Item("ValorTotal")                   'Valor Mensal Correspondente ao municipio
                        linha &= "|"
                    End With
                    strm.WriteLine(linha)
                    registroGeral += 1
                    Registro1400 += 1
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return retorno
    End Function

    Private Function ConsultaRegistroC100_C170_C172_C190() As DataSet
        Sql = GetSqlRegistroC100()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** C101 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        Sql &= GetSqlRegistroC101()


        '******************************************************************************************************************************************************************************
        '*********************************************************************** C170 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        Sql &= GetSqlRegistroC170()


        '******************************************************************************************************************************************************************************
        '*********************************************************************** C172 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        Sql &= GetSqlRegistroC172()


        '******************************************************************************************************************************************************************************
        '*********************************************************************** C190 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        Sql &= GetSqlRegistroC190()

        '******************************************************************************************************************************************************************************
        '****************************************************** RESULTADO C100 - C170 - C190 ******************************************************************************************
        '******************************************************************************************************************************************************************************

        Sql &= "Select * from #C100;" & vbCrLf &
               "Select * from #C101;" & vbCrLf &
               "Select * from #C170;" & vbCrLf &
               "Select * from #C172;" & vbCrLf &
               "Select * from #C190;" & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(Sql, "Consulta")
        ds.Tables(0).TableName = "C100"
        ds.Tables(1).TableName = "C101"
        ds.Tables(2).TableName = "C170"
        ds.Tables(3).TableName = "C172"
        ds.Tables(4).TableName = "C190"

        Return ds
    End Function

    Private Function GetSqlRegistroC100()
        Dim Query As String = String.Empty

        Query = "/* RegistroC100 - Nota Fiscal (Código 01), Nota Fiscal Avulsa (Código 1B), Nota Fiscal de Produtor (Código 04) e Nfe (código 55) */" & vbCrLf &
                "SELECT SB.Empresa_id," & vbCrLf &
                "       SB.EndEmpresa_Id," & vbCrLf &
                "       SB.Cliente_Id," & vbCrLf &
                "       SB.EndCliente_Id," & vbCrLf &
                "       SB.EntradaSaida_Id," & vbCrLf &
                "       SB.Serie_Id," & vbCrLf &
                "       SB.Nota_Id," & vbCrLf &
                "       CASE" & vbCrLf &
                "         WHEN ISNULL(DocXML.Tipo,'') = 'Mic'" & vbCrLf &
                "           THEN 'S'" & vbCrLf &
                "           ELSE SB.NossaEmissao" & vbCrLf &
                "       END AS NossaEmissao," & vbCrLf &
                "       ISNULL(NFe.ChaveNfe, '') AS ChaveNfe," & vbCrLf &
                "       SB.DataDaNota," & vbCrLf &
                "       SB.Movimento," & vbCrLf &
                "       SB.situacao," & vbCrLf &
                "       SB.CIFFobNf," & vbCrLf &
                "       SB.TipoDeDocumento," & vbCrLf &
                "       sum(SB.BaseIcms)         AS BaseIcms," & vbCrLf &
                "       Sum(SB.Aliquota)         AS Aliquota," & vbCrLf &
                "       Sum(SB.ValorIcms)        AS ValorIcms," & vbCrLf &
                "       sum(SB.BaseIcmsST)       AS BaseIcmsST," & vbCrLf &
                "       Sum(SB.ValorIcmsST)        AS ValorIcmsST," & vbCrLf &
                "       Sum(SB.ValorIPI)         AS ValorIPI," & vbCrLf &
                "       Sum(SB.ValorPIS)         AS ValorPIS," & vbCrLf &
                "       Sum(SB.ValorCofins)      AS ValorCofins," & vbCrLf &
                "       SUM(SB.QuantidadeFiscal) AS QuantidadeFiscal," & vbCrLf &
                "       Sum(SB.ValorDesconto) AS ValorDesconto," & vbCrLf &
                "       Sum(SB.ValorFrete)    AS ValorFrete," & vbCrLf &
                "       Sum(SB.ValorSeguro)   AS ValorSeguro," & vbCrLf &
                "       Sum(SB.ValorTotal+SB.ValorIPI+SB.ValorFrete+SB.OutrasDespesas) AS ValorBruto," & vbCrLf &
                "       Sum(CASE" & vbCrLf &
                "             WHEN SB.Situacao <> 1" & vbCrLf &
                "               THEN 0" & vbCrLf &
                "               ELSE ((SB.ValorTotal+SB.ValorIPI+SB.ValorFrete+SB.OutrasDespesas) - SB.ValorDesconto)" & vbCrLf &
                "           END) AS ValorTotal" & vbCrLf &
                "  Into #C100 " & vbCrLf &
                "  FROM (" & vbCrLf &
                "         SELECT NF.Empresa_id," & vbCrLf &
                "                NF.EndEmpresa_Id," & vbCrLf &
                "                NF.EntradaSaida_Id," & vbCrLf &
                "                NF.Cliente_Id," & vbCrLf &
                "                NF.EndCliente_Id," & vbCrLf &
                "                NF.Serie_Id," & vbCrLf &
                "                NF.Nota_Id," & vbCrLf &
                "                CASE" & vbCrLf &
                "                   WHEN (NF.EntradaSaida_Id = 'S'  AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf &
                "                     THEN 'S'" & vbCrLf &
                "		                ELSE NF.NossaEmissao" & vbCrLf &
                "	               END NossaEmissao," & vbCrLf &
                "                NF.DataDaNota," & vbCrLf &
                "                NF.Movimento," & vbCrLf &
                "                NF.CIFFOB AS CIFFobNf," & vbCrLf &
                "                NF.TipoDeDocumento," & vbCrLf &
                "                Enc.ValorTotal," & vbCrLf &
                "                Enc.BaseIcms," & vbCrLf &
                "                Enc.ValorIcms," & vbCrLf &
                "                Enc.BaseIcmsST," & vbCrLf &
                "                Enc.ValorIcmsST," & vbCrLf &
                "                Enc.ValorIPI," & vbCrLf &
                "                Enc.ValorPIS," & vbCrLf &
                "                Enc.ValorFrete," & vbCrLf &
                "                Enc.ValorSeguro," & vbCrLf &
                "                Enc.OutrasDespesas," & vbCrLf &
                "                Enc.ValorDesconto," & vbCrLf &
                "                Enc.ValorCOFINS," & vbCrLf &
                "                Enc.Aliquota," & vbCrLf &
                "                NF.situacao," & vbCrLf &
                "                SO.Classe," & vbCrLf &
                "                NFI.QuantidadeFiscal" & vbCrLf &
                "          FROM NotasFiscais NF" & vbCrLf &
                "         inner join NotasFiscaisxItens NFI" & vbCrLf &
                "            on NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
                "		      AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
                "		      AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                "		      AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
                "		      AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
                "		      AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf &
                "		      AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf &
                "            Inner Join OperacaoXEstado OE" & vbCrLf &
                "		       on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
                "            Inner Join Clientes cli" & vbCrLf &
                "		       on cli.Cliente_Id    = NF.Cliente_Id" & vbCrLf &
                "		       AND cli.Endereco_Id  = NF.EndCliente_Id" & vbCrLf &
                "         Inner Join (" & vbCrLf &
                "					    SELECT NFxE.Empresa_Id," & vbCrLf &
                "					           NFxE.EndEmpresa_Id," & vbCrLf &
                "					           NFxE.Cliente_Id," & vbCrLf &
                "					           NFxE.EndCliente_Id," & vbCrLf &
                "					           NFxE.EntradaSaida_Id," & vbCrLf &
                "					           NFxE.Nota_Id," & vbCrLf &
                "					           NFxE.Serie_Id," & vbCrLf &
                "					           NFxE.Produto_id," & vbCrLf &
                "					           NFxE.Sequencia_Id," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'PRODUTO'" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorTotal," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'ICMS'" & vbCrLf &
                "					    					   then case when ((OxE.CodigoFiscal IN (1551,2551)) or " & vbCrLf &
                "                                                             (OxE .CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 or " & vbCrLf &
                "                                                             so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') then 0 else NFxE.Base end " & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as BaseIcms," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'ICMS'" & vbCrLf &
                "					    					   then case when ((OxE.CodigoFiscal IN (1551,2551)) or  " & vbCrLf &
                "                                                             (OxE .CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 or " & vbCrLf &
                "                                                             so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') then 0 else NFxE.Percentual end " & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as Aliquota," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'ICMS'" & vbCrLf &
                "					    					   then case when ((OxE.CodigoFiscal IN (1551,2551)) or " & vbCrLf &
                "                                                             (OxE .CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 or " & vbCrLf &
                "                                                             so.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') then 0 else NFxE.Valor end " & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorIcms," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'ICMS-ST'" & vbCrLf &
                "					    					   then NFxE.Base" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as BaseIcmsST," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'ICMS-ST'" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorIcmsST," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'IPI'" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorIPI," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'PIS'" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorPIS," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'COFINS'" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorCOFINS," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) = 'DESCONTOS'" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorDesconto," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) IN ('FRETES')" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorFrete," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) IN ('SEGURO')" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as ValorSeguro," & vbCrLf &
                "					    	   isnull(sum(case" & vbCrLf &
                "					    					 when ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) IN ('DESP.ADUANEIRAS','ICMS SUBSTITUIC')" & vbCrLf &
                "					    					   then NFxE.Valor" & vbCrLf &
                "					    					   else 0" & vbCrLf &
                "					    				   end),0) as OutrasDespesas" & vbCrLf &
                "					      FROM NotasFiscaisXEncargos AS NFxE " & vbCrLf &
                "                         Inner Join NotasFiscaisXItens NFxi " & vbCrLf &
                "                            on NFxi.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                "		                   AND NFxi.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                "		                   AND NFxi.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
                "		                   AND NFxi.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                "		                   AND NFxi.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                "		                   AND NFxi.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                "		                   AND NFxi.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                "		                   AND NFxi.Produto_id      = NFxE.Produto_id" & vbCrLf &
                "		                   AND NFxi.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                "                         Inner Join OperacaoXEstado OxE" & vbCrLf &
                "						   on Oxe.Codigo_Id = nfxi.OperacaoXEstado " & vbCrLf &
                "                         INNER JOIN SubOperacoes so" & vbCrLf &
                "						   ON NFxi.Operacao     = so.Operacao_Id " & vbCrLf &
                "						   AND NFxi.SubOperacao = so.SubOperacoes_Id " & vbCrLf &
                "                         INNER JOIN Clientes cli" & vbCrLf &
                "						   ON Cli.Cliente_Id    = NFxE.Cliente_Id " & vbCrLf &
                "						   AND Cli.Endereco_Id  = NFxE.EndCliente_Id " & vbCrLf &
                "					      LEFT JOIN Encargos AS Enc " & vbCrLf &
                "					        ON NFxE.Encargo_Id = Enc.Encargo_id " & vbCrLf &
                "					       AND ISNULL(Enc.EncargoAgrupador,'') <> '' " & vbCrLf &
                "                        Where NFxE.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                "                          AND ISNULL(Enc.EncargoAgrupador,NFxE.Encargo_id) in ('PRODUTO','ICMS','ICMS-ST','IPI','PIS','COFINS','DESCONTOS','FRETES','SEGURO','DESP.ADUANEIRAS','ICMS SUBSTITUIC')" & vbCrLf &
                "                        Group by NFxE.Empresa_Id," & vbCrLf &
                "					              NFxE.EndEmpresa_Id," & vbCrLf &
                "					              NFxE.Cliente_Id," & vbCrLf &
                "					              NFxE.EndCliente_Id," & vbCrLf &
                "					              NFxE.EntradaSaida_Id," & vbCrLf &
                "					              NFxE.Nota_Id," & vbCrLf &
                "					              NFxE.Serie_Id," & vbCrLf &
                "					              NFxE.Produto_id," & vbCrLf &
                "					              NFxE.Sequencia_Id" & vbCrLf &
                "                    ) Enc" & vbCrLf &
                "            on NFI.Empresa_Id      = Enc.Empresa_Id" & vbCrLf &
                "		   AND NFI.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf &
                "		   AND NFI.Cliente_Id      = Enc.Cliente_Id" & vbCrLf &
                "		   AND NFI.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf &
                "		   AND NFI.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf &
                "		   AND NFI.Serie_Id        = Enc.Serie_Id" & vbCrLf &
                "		   AND NFI.Nota_Id         = Enc.Nota_Id" & vbCrLf &
                "		   AND NFI.Produto_id      = Enc.Produto_id" & vbCrLf &
                "		   AND NFI.Sequencia_Id    = Enc.Sequencia_Id" & vbCrLf &
                "         INNER JOIN SubOperacoes SO" & vbCrLf &
                "            ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf &
                "           AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                "         WHERE NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
                "           AND isnull(NF.Operacao,0)  > 0" & vbCrLf &
                "           AND NF.Situacao not in(9,10)" & vbCrLf

        'NÃO LEVAR TIPO DE DOCUMENTO 65, SOLICITAÇÃO JÉSSICA EM 19/12/2023 - FURLAN
        If Left(Empresa(0), 8) = "05366261" OrElse
            Left(Empresa(0), 8) = "38198213" OrElse
            Left(Empresa(0), 8) = "62747840" OrElse
            Left(Empresa(0), 8) = "62780383" OrElse
            Left(Empresa(0), 8) = "63358210" OrElse
            Left(Empresa(0), 8) = "40938762" OrElse
            Left(Empresa(0), 8) = "49673784" Then
            Query &= "           AND NF.TipoDeDocumento NOT IN(2,8,10,11,14,15,16,65)" & vbCrLf
        Else
            Query &= "           AND NF.TipoDeDocumento NOT IN(2,8,10,11,14,15,16)" & vbCrLf
        End If

        Query &= "           AND NF.Finalidade NOT IN (30)" & vbCrLf &
                "           AND NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
                "           AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT IN (1932,2932, 5932, 6932))" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                "           AND(OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                "           AND NOT(NF.EntradaSaida_Id = 'E'  AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf &
                "           AND NOT(SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                "           --AND NOT(NF.EntradaSaida_Id = 'S'  AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf &
                "           AND NOT(SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                "     ) AS SB" & vbCrLf &
                "  LEFT JOIN NFERealizadas NFe" & vbCrLf &
                "    ON NFe.Empresa_Id      = SB.Empresa_Id" & vbCrLf &
                "   AND NFe.EndEmpresa_Id   = SB.EndEmpresa_Id" & vbCrLf &
                "   AND NFe.Cliente_Id      = SB.Cliente_Id" & vbCrLf &
                "   AND NFe.EndCliente_Id   = SB.EndCliente_Id" & vbCrLf &
                "   AND NFe.EntradaSaida_Id = SB.EntradaSaida_Id" & vbCrLf &
                "   AND NFe.Serie_Id        = SB.Serie_Id" & vbCrLf &
                "   AND NFe.Nota_Id         = SB.Nota_Id" & vbCrLf &
                "  LEFT JOIN DocumentoXML DocXML" & vbCrLf &
                "    ON DocXML.Empresa_Id      = SB.Empresa_Id" & vbCrLf &
                "   AND DocXML.Cliente_Id      = SB.Cliente_Id" & vbCrLf &
                "   AND DocXML.Serie_Id        = SB.Serie_Id" & vbCrLf &
                "   AND DocXML.Numero_Id       = SB.Nota_Id" & vbCrLf &
                " Group by SB.Empresa_Id," & vbCrLf &
                "          SB.EndEmpresa_Id," & vbCrLf &
                "          SB.Cliente_Id," & vbCrLf &
                "          SB.EndCliente_Id," & vbCrLf &
                "          SB.EntradaSaida_Id," & vbCrLf &
                "          SB.Nota_Id," & vbCrLf &
                "          SB.Serie_Id," & vbCrLf &
                "          SB.NossaEmissao," & vbCrLf &
                "          ISNULL(DocXML.Tipo,'')," & vbCrLf &
                "          ISNULL(NFe.ChaveNfe, '')," & vbCrLf &
                "          SB.DataDaNota," & vbCrLf &
                "          SB.Movimento," & vbCrLf &
                "          SB.situacao," & vbCrLf &
                "          SB.CIFFobNf," & vbCrLf &
                "          SB.TipoDeDocumento;" & vbCrLf &
                "" & vbCrLf

        Return Query
    End Function

    Private Function GetSqlRegistroC101()
        Dim Query As String = String.Empty
        Query = "SELECT NF.Empresa_Id, NF.EndEmpresa_Id," & vbCrLf &
                "       NF.Cliente_id, NF.EndCliente_Id," & vbCrLf &
                "	    NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf &
                "	    SUM(NF.Valor) AS Valor" & vbCrLf &
                "  INTO #C101" & vbCrLf &
                "  FROM NotasFiscaisXEncargos NF" & vbCrLf &
                " Inner Join NotasFiscaisXItens NFxi" & vbCrLf &
                "    on NFxi.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                "   AND NFxi.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "   AND NFxi.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "   AND NFxi.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                "   AND NFxi.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                "   AND NFxi.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "   AND NFxi.Nota_Id         = NF.Nota_Id" & vbCrLf &
                "   AND NFxi.Produto_id      = NF.Produto_id" & vbCrLf &
                "   AND NFxi.Sequencia_Id    = NF.Sequencia_Id" & vbCrLf &
                " Inner Join OperacaoXEstado OxE" & vbCrLf &
                "	on Oxe.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                " Inner Join #C100 C" & vbCrLf &
                "    on C.Empresa_id      = nf.Empresa_id" & vbCrLf &
                "   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf &
                "   and c.Cliente_id      = nf.cliente_id" & vbCrLf &
                "   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf &
                "   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf &
                "   and c.Nota_id         = nf.Nota_id" & vbCrLf &
                "   and C.Serie_id        = nf.Serie_id" & vbCrLf &
                "  LEFT Join Encargos AS Enc" & vbCrLf &
                "    on NF.Encargo_Id = Enc.Encargo_id" & vbCrLf &
                "   and ISNULL(Enc.EncargoAgrupador,'') <> ''" & vbCrLf &
                " WHERE isnull(enc.EncargoAgrupador,NF.Encargo_Id) IN ('FUNDO FECP')" & vbCrLf &
                "  Group by NF.Empresa_Id, NF.EndEmpresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id;" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf

        Return Query
    End Function

    Private Function GetSqlRegistroC170()
        Dim Query As String = String.Empty
        Query = "SELECT NF.Empresa_Id, NF.EndEmpresa_Id," & vbCrLf &
                "       NF.Cliente_id, NF.EndCliente_Id," & vbCrLf &
                "	    NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf &
                "	    NF.Produto_Id, NF.Sequencia_Id, isnull(enc.EncargoAgrupador,NF.Encargo_Id) as Encargo_id," & vbCrLf &
                "       OxE.CodigoFiscal,  OxE.STICMS, OxE.STIPI, OxE.STPISCOFINS," & vbCrLf &
                "	    NF.Base, NF.Valor, NF.Percentual, cli.Estado" & vbCrLf &
                "  INTO #Encargos" & vbCrLf &
                "  FROM NotasFiscaisXEncargos NF" & vbCrLf &
                " Inner Join NotasFiscaisXItens NFxi" & vbCrLf &
                "    on NFxi.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                "   AND NFxi.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "   AND NFxi.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "   AND NFxi.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                "   AND NFxi.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                "   AND NFxi.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "   AND NFxi.Nota_Id         = NF.Nota_Id" & vbCrLf &
                "   AND NFxi.Produto_id      = NF.Produto_id" & vbCrLf &
                "   AND NFxi.Sequencia_Id    = NF.Sequencia_Id" & vbCrLf &
                " Inner Join OperacaoXEstado OxE" & vbCrLf &
                "	on Oxe.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf &
                " Inner Join Clientes cli" & vbCrLf &
                "	on cli.Cliente_Id    = NF.Cliente_Id" & vbCrLf &
                "	AND cli.Endereco_Id  = NF.EndCliente_Id" & vbCrLf &
                " Inner Join #C100 C" & vbCrLf &
                "    on C.Empresa_id      = nf.Empresa_id" & vbCrLf &
                "   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf &
                "   and c.Cliente_id      = nf.cliente_id" & vbCrLf &
                "   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf &
                "   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf &
                "   and c.Nota_id         = nf.Nota_id" & vbCrLf &
                "   and C.Serie_id        = nf.Serie_id" & vbCrLf &
                "  LEFT Join Encargos AS Enc" & vbCrLf &
                "    on NF.Encargo_Id = Enc.Encargo_id" & vbCrLf &
                "   and ISNULL(Enc.EncargoAgrupador,'') <> ''" & vbCrLf &
                " WHERE isnull(enc.EncargoAgrupador,NF.Encargo_Id) IN ('PRODUTO', 'DESCONTOS', 'IPI', 'PIS', 'COFINS', 'ISS', 'ICMS','ICMS-ST');" & vbCrLf &
                "                                         " & vbCrLf &
                "                                         " & vbCrLf &
                "SELECT NF.EntradaSaida_Id," & vbCrLf &
                "       NF.Cliente_Id," & vbCrLf &
                "       NF.EndCliente_Id," & vbCrLf &
                "       NF.Serie_Id," & vbCrLf &
                "       NF.Nota_Id," & vbCrLf &
                "       'C170' AS Registro," & vbCrLf &
                "       Produtos.Produto_Id AS Item," & vbCrLf &
                "       Produtos.Nome," & vbCrLf &
                "       NF.NossaEmissao," & vbCrLf &
                "       '' AS ChaveNfe," & vbCrLf &
                "       NF.DataDaNota," & vbCrLf &
                "       NF.Movimento," & vbCrLf &
                "       NFI.QuantidadeFiscal," & vbCrLf &
                "       Produtos.Unidade," & vbCrLf &
                "       OE.CodigoFiscal AS CFOP," & vbCrLf &
                "       NFI.Operacao," & vbCrLf &
                "       NFI.SubOperacao," & vbCrLf &
                "       Enc.SituacaoTributaria," & vbCrLf &
                "       Enc.ValorTotal," & vbCrLf &
                "       Enc.ValorDesconto," & vbCrLf &
                "       Enc.BaseICMS," & vbCrLf &
                "       Enc.Aliquota," & vbCrLf &
                "       Enc.ValorICMS," & vbCrLf &
                "       Enc.BaseIcmsST," & vbCrLf &
                "       Enc.AliquotaST," & vbCrLf &
                "       Enc.ValorIcmsST," & vbCrLf &
                "       Enc.ValorIPI," & vbCrLf &
                "       Enc.ValorPIS," & vbCrLf &
                "       Enc.ValorCOFINS," & vbCrLf &
                "       Enc.CST_IPI," & vbCrLf &
                "       Enc.COD_ENQ," & vbCrLf &
                "       Enc.VL_BC_IPI," & vbCrLf &
                "       Enc.ALI_IPI," & vbCrLf &
                "       Enc.VL_IPI," & vbCrLf &
                "       Enc.CST_PIS," & vbCrLf &
                "       Enc.VL_BC_PIS," & vbCrLf &
                "       Enc.ALI_PIS," & vbCrLf &
                "       Enc.VL_PIS," & vbCrLf &
                "       Enc.CST_COFINS," & vbCrLf &
                "       Enc.VL_BC_COFINS," & vbCrLf &
                "       Enc.ALI_COFINS," & vbCrLf &
                "       Enc.VL_COFINS," & vbCrLf &
                "       '' AS COD_CTA," & vbCrLf &
                "       SO.EstoqueFiscal," & vbCrLf &
                "       0.00 AS QUANT_BC_COFINS," & vbCrLf &
                "       0.00 AS ALI_COFINSR," & vbCrLf &
                "       0.00 AS QUANT_BC_PIS," & vbCrLf &
                "       0 AS ALI_PISR" & vbCrLf &
                "  Into #C170" & vbCrLf &
                "  FROM NotasFiscais NF" & vbCrLf &
                " INNER JOIN NotasFiscaisXItens NFI" & vbCrLf &
                "    ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
                "   AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
                "   AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                "   AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
                "   AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
                "   AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf &
                "   AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf &
                " Inner Join OperacaoXEstado OE" & vbCrLf &
                "    on OE.Codigo_id = NFI.OperacaoXEstado" & vbCrLf &
                "  LEFT Join ProdutosAgrupados PA" & vbCrLf &
                "    ON PA.ProdutoAgrupado_Id = NFI.Produto_Id" & vbCrLf &
                " INNER JOIN(" & vbCrLf &
                "			SELECT Empresa_Id," & vbCrLf &
                "			       EndEmpresa_Id," & vbCrLf &
                "			       Cliente_Id," & vbCrLf &
                "			       EndCliente_Id," & vbCrLf &
                "			       EntradaSaida_Id," & vbCrLf &
                "			       Nota_Id," & vbCrLf &
                "			       Serie_Id," & vbCrLf &
                "			       Produto_Id," & vbCrLf &
                "			       CodigoFiscal AS Cfop_Id," & vbCrLf &
                "			       Sequencia_Id," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'PRODUTO'" & vbCrLf &
                "								   then STICMS" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as SituacaoTributaria," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'PRODUTO'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ValorTotal," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'DESCONTOS'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ValorDesconto," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'ICMS'" & vbCrLf &
                "								   then CASE when ((CodigoFiscal IN (1551,2551)) or (CodigoFiscal IN (1653,2653) and STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 then 0 else isnull(Base,0)" & vbCrLf &
                "								   END else 0" & vbCrLf &
                "							   end),0) as BaseIcms," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id= 'ICMS'" & vbCrLf &
                "								   then CASE when ((CodigoFiscal IN (1551,2551)) or (CodigoFiscal IN (1653,2653) and STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 then 0 else Percentual" & vbCrLf &
                "								   END else 0" & vbCrLf &
                "							   end),0) as Aliquota," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'ICMS'" & vbCrLf &
                "								   then CASE when ((CodigoFiscal IN (1551,2551)) or (CodigoFiscal IN (1653,2653) and STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 then 0 else Valor" & vbCrLf &
                "								   END else 0" & vbCrLf &
                "							   end),0) as ValorIcms," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'ICMS-ST'" & vbCrLf &
                "								   then Base" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as BaseIcmsST," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'ICMS-ST'" & vbCrLf &
                "								   then Percentual" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as AliquotaST," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'ICMS-ST'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ValorIcmsST," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'IPI'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ValorIPI," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'PIS'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ValorPIS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'COFINS'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ValorCOFINS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'IPI'" & vbCrLf &
                "								   then STIPI" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as CST_IPI," & vbCrLf &
                "				   0 AS COD_ENQ," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'IPI'" & vbCrLf &
                "								   then Base" & vbCrLf &
                "								   else 0 " & vbCrLf &
                "							   end),0) as VL_BC_IPI," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'IPI'" & vbCrLf &
                "								   then Percentual" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ALI_IPI," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'IPI'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as VL_IPI," & vbCrLf &
                "				   STPISCOFINS as CST_PIS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'PIS'" & vbCrLf &
                "								   then Base" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as VL_BC_PIS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'PIS'" & vbCrLf &
                "								   then Percentual" & vbCrLf &
                "								   else 0 " & vbCrLf &
                "							   end),0) as ALI_PIS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'PIS'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as VL_PIS," & vbCrLf &
                "				   STPISCOFINS as CST_COFINS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'COFINS'" & vbCrLf &
                "								   then Base" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as VL_BC_COFINS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'COFINS'" & vbCrLf &
                "								   then Percentual" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as ALI_COFINS," & vbCrLf &
                "				   isnull(sum(case" & vbCrLf &
                "								 when Encargo_Id = 'COFINS'" & vbCrLf &
                "								   then Valor" & vbCrLf &
                "								   else 0" & vbCrLf &
                "							   end),0) as VL_COFINS" & vbCrLf &
                "			  FROM #Encargos" & vbCrLf &
                "			 Group By Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id, Produto_Id, CodigoFiscal, Sequencia_Id, STPISCOFINS" & vbCrLf &
                "          ) Enc" & vbCrLf &
                "    ON NFI.Empresa_Id      = Enc.Empresa_Id" & vbCrLf &
                "   AND NFI.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf &
                "   AND NFI.Cliente_Id      = Enc.Cliente_Id" & vbCrLf &
                "   AND NFI.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf &
                "   AND NFI.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf &
                "   AND NFI.Serie_Id        = Enc.Serie_Id" & vbCrLf &
                "   AND NFI.Nota_Id         = Enc.Nota_Id" & vbCrLf &
                "   AND NFI.Produto_Id      = Enc.Produto_Id" & vbCrLf &
                "   AND NFI.CFOP_Id         = Enc.Cfop_Id" & vbCrLf &
                "   AND NFI.Sequencia_id    = Enc.Sequencia_id" & vbCrLf &
                " INNER JOIN Produtos" & vbCrLf &
                "    ON Produtos.Produto_Id = isnull(PA.Produto_Id,NFI.Produto_Id)" & vbCrLf &
                " INNER JOIN SubOperacoes SO" & vbCrLf &
                "    ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf &
                "   AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                "  LEFT JOIN NFERealizadas " & vbCrLf &
                "    ON NF.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf &
                "   AND NF.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf &
                "   AND NF.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf &
                "   AND NF.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf &
                "   AND NF.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf &
                "   AND NF.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf &
                "   AND NF.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf &
                "  LEFT JOIN DocumentoXML DocXML" & vbCrLf &
                "    ON DocXML.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                "   AND DocXML.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "   AND DocXML.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "   AND DocXML.Numero_Id       = NF.Nota_Id" & vbCrLf &
                " Inner Join #C100 C" & vbCrLf &
                "    on C.Empresa_id      = nf.Empresa_id" & vbCrLf &
                "   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf &
                "   and c.Cliente_id      = nf.cliente_id" & vbCrLf &
                "   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf &
                "   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf &
                "   and c.Nota_id         = nf.Nota_id" & vbCrLf &
                "   and C.Serie_id        = nf.Serie_id" & vbCrLf &
                " WHERE isnull(NFI.Operacao,0) > 0" & vbCrLf &
                "   AND NF.Situacao = 1 --Colocado Agora " & vbCrLf &
                "   AND (NF.NossaEmissao = 'N' AND ISNULL(DocXML.Tipo,'') NOT IN ('Mic')) --Colocado Agora " & vbCrLf

        'NÃO LEVAR TIPO DE DOCUMENTO 65, SOLICITAÇÃO JÉSSICA EM 19/12/2023 - FURLAN
        If Left(Empresa(0), 8) = "05366261" OrElse
            Left(Empresa(0), 8) = "38198213" OrElse
            Left(Empresa(0), 8) = "62747840" OrElse
            Left(Empresa(0), 8) = "62780383" OrElse
            Left(Empresa(0), 8) = "63358210" OrElse
            Left(Empresa(0), 8) = "40938762" OrElse
            Left(Empresa(0), 8) = "49673784" Then

            Query &= "   AND NF.TipoDeDocumento NOT IN(2,8,10,11,14,15,16,57,65)" & vbCrLf

        Else
            Query &= "   AND NF.TipoDeDocumento NOT IN (15,57)" & vbCrLf
        End If

        Query &= "   AND NF.Finalidade NOT IN (30)" & vbCrLf &
                "   AND NOT(NF.EntradaSaida_Id = 'E'  AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf &
                "   AND NOT(SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                "   --AND NOT(NF.EntradaSaida_Id = 'S'  AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf &
                "   AND NOT(SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949)); " & vbCrLf &
                "" & vbCrLf &
                ""  & vbCrLf

        Return Query
    End Function

    Private Function GetSqlRegistroC172()
        Dim Query As String = String.Empty

        Query = "SELECT Resultado.Nota_Id," & vbCrLf & _
                "       Resultado.Cliente_Id," & vbCrLf & _
                "       Resultado.EndCliente_Id," & vbCrLf & _
                "       Resultado.Serie_Id," & vbCrLf & _
                "       Resultado.EntradaSaida_Id," & vbCrLf & _
                "       SUM(VL_BC_ISSQN) as VL_BC_ISSQN," & vbCrLf & _
                "       ALIQ_ISSQN AS ALIQ_ISSQN," & vbCrLf & _
                "       SUM(VL_ISSQN) as VL_ISSQN" & vbCrLf & _
                "  INTO #C172" & vbCrLf & _
                "  FROM (" & vbCrLf & _
                "                Select NF.Empresa_Id," & vbCrLf & _
                "                       NF.EndEmpresa_Id," & vbCrLf & _
                "                       NF.Cliente_Id," & vbCrLf & _
                "                       NF.EndCliente_Id," & vbCrLf & _
                "                       NF.Nota_Id," & vbCrLf & _
                "                       NF.Serie_Id," & vbCrLf & _
                "                       NF.EntradaSaida_Id," & vbCrLf & _
                "                       SUM(CASE" & vbCrLf & _
                "                              when ((OxE.CodigoFiscal IN (1551,2551)) or (OxE.CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011" & vbCrLf & _
                "                                then 0" & vbCrLf & _
                "                                else case" & vbCrLf & _
                "                                       when isnull(e.encargoagrupador,NFE.Encargo_Id) = 'ISS'" & vbCrLf & _
                "                                         then NFE.Base" & vbCrLf & _
                "                                         else 0" & vbCrLf & _
                "                                     end" & vbCrLf & _
                "                           end) as VL_BC_ISSQN," & vbCrLf & _
                "                        MAX(CASE" & vbCrLf & _
                "						      when ((OxE.CodigoFiscal IN (1551,2551)) or (OxE.CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011" & vbCrLf & _
                "						    	then 0" & vbCrLf & _
                "						     	else case" & vbCrLf & _
                "						     	       when isnull(e.encargoagrupador,NFE.Encargo_Id) ='ISS'" & vbCrLf & _
                "						     	         then Case" & vbCrLf & _
                "						    		            when isnull(OEE.AliquotaExibicao,0) = 0" & vbCrLf & _
                "						    		              then NFE.Percentual" & vbCrLf & _
                "						    			          else OEE.AliquotaExibicao" & vbCrLf & _
                "						    			      end" & vbCrLf & _
                "						    		   else 0" & vbCrLf & _
                "						    		 end" & vbCrLf & _
                "					         end) as ALIQ_ISSQN," & vbCrLf & _
                "                       SUM(CASE" & vbCrLf & _
                "                             when ((OxE.CodigoFiscal IN (1551,2551)) or (OxE.CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011" & vbCrLf & _
                "                               then 0" & vbCrLf & _
                "                               else case" & vbCrLf & _
                "                                       when isnull(e.encargoagrupador,NFE.Encargo_Id) = 'ISS'" & vbCrLf & _
                "                                         then NFE.valor" & vbCrLf & _
                "                                         else 0" & vbCrLf & _
                "                                     end" & vbCrLf & _
                "                           end) as VL_ISSQN" & vbCrLf & _
                "                  FROM NotasFiscais NF" & vbCrLf & _
                "				 INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
                "					ON NFI.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
                "				   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
                "				   AND NFI.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
                "				   AND NFI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
                "				   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
                "				   AND NFI.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                "				   AND NFI.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                "                 INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf & _
                "                    ON NFE.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
                "                   AND NFE.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                "                   AND NFE.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                "                   AND NFE.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                "                   AND NFE.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
                "                   AND NFE.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                "                   AND NFE.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
                "                   AND NFE.Produto_Id      = NFI.Produto_Id" & vbCrLf & _
                "                   AND NFE.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
                "                  Left Join Encargos e" & vbCrLf & _
                "                    on e.Encargo_id = nfe.encargo_id" & vbCrLf & _
                "                   and isnull(e.encargoagrupador,'') <> ''" & vbCrLf & _
                "                 Inner Join OperacaoXEstado OxE" & vbCrLf & _
                "				    on OxE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf & _
                "                 Inner join OperacaoXEstadoXEncargo OEE" & vbCrLf & _
                "				    on OEE.Codigo_Id  = OxE.Codigo_Id" & vbCrLf & _
                "				   and OEE.Encargo_Id = NFE.Encargo_Id" & vbCrLf & _
                "                 INNER JOIN SubOperacoes SO" & vbCrLf & _
                "                    ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf & _
                "                   AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                "                 Inner Join #C100 C" & vbCrLf & _
                "                    on C.Empresa_id      = nf.Empresa_id" & vbCrLf & _
                "                   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf & _
                "                   and c.Cliente_id      = nf.cliente_id" & vbCrLf & _
                "                   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf & _
                "                   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf & _
                "                   and c.Nota_id         = nf.Nota_id" & vbCrLf & _
                "                   and C.Serie_id        = nf.Serie_id" & vbCrLf & _
                "                 WHERE NF.Situacao = 1" & vbCrLf

        'NÃO LEVAR TIPO DE DOCUMENTO 65, SOLICITAÇÃO JÉSSICA EM 19/12/2023 - FURLAN
        If Left(Empresa(0), 8) = "05366261" OrElse
            Left(Empresa(0), 8) = "38198213" OrElse
            Left(Empresa(0), 8) = "62747840" OrElse
            Left(Empresa(0), 8) = "62780383" OrElse
            Left(Empresa(0), 8) = "63358210" OrElse
            Left(Empresa(0), 8) = "40938762" OrElse
            Left(Empresa(0), 8) = "49673784" Then

            Query &= "                   AND NF.TipoDeDocumento NOT IN(2,8,10,11,14,15,16,57,65)" & vbCrLf

        Else
            Query &= "                   AND NF.TipoDeDocumento NOT IN (15)" & vbCrLf
        End If

        Query &= "                   AND NF.Finalidade NOT IN (30)" & vbCrLf &
                "                   AND isnull(e.encargoagrupador,NFE.Encargo_Id) in ('PRODUTO','ISS')" & vbCrLf &
                "                   --AND NOT (NF.EntradaSaida_Id = 'E'  AND OxE.CodigoFiscal IN (1933,2933))" & vbCrLf &
                "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OxE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                "                   AND (NF.EntradaSaida_Id = 'S' AND OxE.CodigoFiscal IN (5933,6933))" & vbCrLf &
                "                   --AND NOT (NF.EntradaSaida_Id = 'S'  AND NFI.CFOP_Id IN (5933,6933))" & vbCrLf &
                "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OxE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                "                 GROUP BY NF.Empresa_id," & vbCrLf &
                "                          NF.EndEmpresa_Id," & vbCrLf &
                "                          NF.Cliente_id," & vbCrLf &
                "                          NF.EndCliente_id," & vbCrLf &
                "                          NF.EntradaSaida_id," & vbCrLf &
                "                          NF.Nota_id," & vbCrLf &
                "                          NF.Serie_id" & vbCrLf &
                "     ) as Resultado" & vbCrLf &
                " GROUP BY Resultado.Empresa_Id," & vbCrLf &
                "          Resultado.EndEmpresa_Id," & vbCrLf &
                "          Resultado.Cliente_Id," & vbCrLf &
                "          Resultado.EndCliente_Id," & vbCrLf &
                "          Resultado.Nota_Id," & vbCrLf &
                "          Resultado.Serie_Id," & vbCrLf &
                "          Resultado.EntradaSaida_Id," & vbCrLf &
                "          Resultado.ALIQ_ISSQN;" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf


        'Query = "SELECT Resultado.Nota_Id, " & vbCrLf & _
        '       "       Resultado.Cliente_Id, " & vbCrLf & _
        '       "       Resultado.EndCliente_Id, " & vbCrLf & _
        '       "       Resultado.Serie_Id, " & vbCrLf & _
        '       "       Resultado.EntradaSaida_Id, " & vbCrLf & _
        '       "       SUM(VL_BC_ISSQN) as VL_BC_ISSQN," & vbCrLf & _
        '       "       ALIQ_ISSQN AS ALIQ_ISSQN," & vbCrLf & _
        '       "       SUM(VL_ISSQN) as VL_ISSQN " & vbCrLf & _
        '       "  INTO #C172" & vbCrLf & _
        '       "  FROM (" & vbCrLf & _
        '       "                Select NF.Empresa_Id," & vbCrLf & _
        '       "                       NF.EndEmpresa_Id," & vbCrLf & _
        '       "                       NF.Cliente_Id, " & vbCrLf & _
        '       "                       NF.EndCliente_Id, " & vbCrLf & _
        '       "                       NF.Nota_Id, " & vbCrLf & _
        '       "                       NF.Serie_Id, " & vbCrLf & _
        '       "                       NF.EntradaSaida_Id, " & vbCrLf & _
        '       "                       SUM(CASE" & vbCrLf & _
        '       "                              when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '       "                                then 0" & vbCrLf & _
        '       "                                else case" & vbCrLf & _
        '       "                                       when NFE.Encargo_Id = 'ISS'" & vbCrLf & _
        '       "                                         then NFE.Base" & vbCrLf & _
        '       "                                         else 0" & vbCrLf & _
        '       "                                     end" & vbCrLf & _
        '       "                           end) as VL_BC_ISSQN," & vbCrLf
        'If Empresa(0).ToString.Contains("04440724") Then
        '    Query &= "                        MAX(CASE" & vbCrLf & _
        '           "						      when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '           "						    	then 0" & vbCrLf & _
        '           "						     	else case" & vbCrLf & _
        '           "						     	       when NFE.Encargo_Id ='ISS'" & vbCrLf & _
        '           "						     	         then NFE.Percentual" & vbCrLf & _
        '           "						    			 else 0" & vbCrLf & _
        '           "						    		 end" & vbCrLf & _
        '           "					        end) as ALIQ_ISSQN," & vbCrLf
        'Else
        '    Query &= "                        MAX(CASE" & vbCrLf & _
        '           "						      when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '           "						    	then 0" & vbCrLf & _
        '           "						     	else case" & vbCrLf & _
        '           "						     	       when NFE.Encargo_Id ='ISS'" & vbCrLf & _
        '           "						     	         then Case" & vbCrLf & _
        '           "						    		            when isnull(NFE.PercentualExibicao,0) = 0" & vbCrLf & _
        '           "						    		              then NFE.Percentual" & vbCrLf & _
        '           "						    			          else NFE.PercentualExibicao" & vbCrLf & _
        '           "						    			      end" & vbCrLf & _
        '           "						    		   else 0" & vbCrLf & _
        '           "						    		 end" & vbCrLf & _
        '           "					         end) as ALIQ_ISSQN," & vbCrLf
        'End If

        'Query &= "                       SUM(CASE" & vbCrLf & _
        '       "                             when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '       "                               then 0" & vbCrLf & _
        '       "                               else case" & vbCrLf & _
        '       "                                       when NFE.Encargo_Id = 'ISS'" & vbCrLf & _
        '       "                                         then NFE.valor" & vbCrLf & _
        '       "                                         else 0" & vbCrLf & _
        '       "                                     end" & vbCrLf & _
        '       "                           end) as VL_ISSQN" & vbCrLf & _
        '       "                  FROM NotasFiscais NF" & vbCrLf & _
        '       "				 INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
        '       "					ON NFI.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
        '       "				   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
        '       "				   AND NFI.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
        '       "				   AND NFI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
        '       "				   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
        '       "				   AND NFI.Serie_Id        = NF.Serie_Id" & vbCrLf & _
        '       "				   AND NFI.Nota_Id         = NF.Nota_Id" & vbCrLf & _
        '       "                 INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf & _
        '       "                    ON NFE.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
        '       "                   AND NFE.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
        '       "                   AND NFE.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
        '       "                   AND NFE.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
        '       "                   AND NFE.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
        '       "                   AND NFE.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
        '       "                   AND NFE.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
        '       "                   AND NFE.Produto_Id      = NFI.Produto_Id" & vbCrLf & _
        '       "                   AND NFE.CFOP_Id         = NFI.CFOP_Id" & vbCrLf & _
        '       "                   AND NFE.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
        '       "                 INNER JOIN SubOperacoes SO" & vbCrLf & _
        '       "                    ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf & _
        '       "                   AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
        '       "                 Inner Join #C100 C" & vbCrLf & _
        '       "                    on C.Empresa_id      = nf.Empresa_id" & vbCrLf & _
        '       "                   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf & _
        '       "                   and c.Cliente_id      = nf.cliente_id" & vbCrLf & _
        '       "                   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf & _
        '       "                   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf & _
        '       "                   and c.Nota_id         = nf.Nota_id" & vbCrLf & _
        '       "                   and C.Serie_id        = nf.Serie_id" & vbCrLf & _
        '       "                 WHERE NF.Situacao = 1" & vbCrLf & _
        '       "                   AND NFE.Encargo_Id in ('PRODUTO','ISS')" & vbCrLf & _
        '       "                   --AND NOT (NF.EntradaSaida_Id = 'E'  AND NFI.CFOP_Id IN (1933,2933))" & vbCrLf & _
        '       "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NFI.CFOP_Id IN (1949,2949))" & vbCrLf & _
        '       "                   AND NFE.Encargo_Id in ('PRODUTO','ISS') " & vbCrLf & _
        '       "                   AND (NF.EntradaSaida_Id = 'S' AND NFI.CFOP_Id IN (5933,6933))" & vbCrLf & _
        '       "                   --AND NOT (NF.EntradaSaida_Id = 'S'  AND NFI.CFOP_Id IN (5933,6933))" & vbCrLf & _
        '       "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NFI.CFOP_Id IN (5949,6949))" & vbCrLf & _
        '       "                 GROUP BY NF.Empresa_id," & vbCrLf & _
        '       "                          NF.EndEmpresa_Id," & vbCrLf & _
        '       "                          NF.Cliente_id," & vbCrLf & _
        '       "                          NF.EndCliente_id," & vbCrLf & _
        '       "                          NF.EntradaSaida_id," & vbCrLf & _
        '       "                          NF.Nota_id," & vbCrLf & _
        '       "                          NF.Serie_id" & vbCrLf & _
        '       "     ) as Resultado" & vbCrLf & _
        '       " GROUP BY Resultado.Empresa_Id," & vbCrLf & _
        '       "          Resultado.EndEmpresa_Id," & vbCrLf & _
        '       "          Resultado.Cliente_Id," & vbCrLf & _
        '       "          Resultado.EndCliente_Id," & vbCrLf & _
        '       "          Resultado.Nota_Id," & vbCrLf & _
        '       "          Resultado.Serie_Id," & vbCrLf & _
        '       "          Resultado.EntradaSaida_Id, " & vbCrLf & _
        '       "          Resultado.ALIQ_ISSQN " & vbCrLf

        Return Query
    End Function

    Private Function GetSqlRegistroC190()
        Dim Query As String = String.Empty
        Query = "Select Resultado.Nota_Id," & vbCrLf &
                "       Resultado.Cliente_Id," & vbCrLf &
                "       Resultado.EndCliente_Id," & vbCrLf &
                "       Resultado.Serie_Id," & vbCrLf &
                "       Resultado.EntradaSaida_Id," & vbCrLf &
                "       STICMS AS CST_ICMS," & vbCrLf &
                "       CodigoFiscal AS CFOP," & vbCrLf &
                "       Percentual AS ALIQ_ICMS," & vbCrLf &
                "       Sum(Produto+IPI+ValorFrete+OutrasDespesas-ValorDesconto) as VL_OPR," & vbCrLf &
                "       Sum(Base) as VL_BC_ICMS," & vbCrLf &
                "       Sum(Icms) as VL_ICMS," & vbCrLf &
                "       Sum(BaseST) as VL_BC_ICMSST," & vbCrLf &
                "       Sum(IcmsST) as VL_ICMSST," & vbCrLf &
                "       Sum(Ipi) as VL_IPI" & vbCrLf &
                "  into #C190" & vbCrLf &
                "  FROM (" & vbCrLf &
                "                Select NF.Empresa_Id," & vbCrLf &
                "                       NF.EndEmpresa_Id," & vbCrLf &
                "                       NF.Cliente_Id," & vbCrLf &
                "                       NF.EndCliente_Id," & vbCrLf &
                "                       NF.Nota_Id," & vbCrLf &
                "                       NF.Serie_Id," & vbCrLf &
                "                       NF.EntradaSaida_Id," & vbCrLf &
                "                       NFI.Produto_Id," & vbCrLf &
                "                       OxE.CodigoFiscal," & vbCrLf &
                "                       OxE.STICMS," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'PRODUTO'" & vbCrLf &
                "                                then NFE.valor" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end) as Produto," & vbCrLf &
                "                       SUM(CASE" & vbCrLf &
                "                              when ((OxE.CodigoFiscal IN (1551,2551)) or " & vbCrLf &
                "                                    (OxE.CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 or " & vbCrLf &
                "                                    SO.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') " & vbCrLf &
                "                                then 0" & vbCrLf &
                "                                else case" & vbCrLf &
                "                                       when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS'" & vbCrLf &
                "                                         then NFE.Base" & vbCrLf &
                "                                         else 0" & vbCrLf &
                "                                     end" & vbCrLf &
                "                           end) as Base," & vbCrLf &
                "                        MAX(CASE" & vbCrLf &
                "						      when ((OxE.CodigoFiscal IN (1551,2551)) or " & vbCrLf &
                "                                    (OxE.CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 or " & vbCrLf &
                "                                    SO.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') " & vbCrLf &
                "						    	then 0" & vbCrLf &
                "                               else case" & vbCrLf &
                "                                       when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS'" & vbCrLf &
                "                                         then NFE.Percentual" & vbCrLf &
                "                                         else 0" & vbCrLf &
                "                                     end" & vbCrLf &
                "					         end) as Percentual," & vbCrLf &
                "                       SUM(CASE" & vbCrLf &
                "                             when ((OxE.CodigoFiscal IN (1551,2551)) or " & vbCrLf &
                "                                    (OxE.CodigoFiscal IN (1653,2653) and OxE.STICMS in(10,60))) AND YEAR('" & PeriodoFinal.ToSqlDate() & "') > 2011 or " & vbCrLf &
                "                                    SO.Descricao IN('TRANSFERENCIA DE CREDITO DE ICMS ACUMULADO') " & vbCrLf &
                "                               then 0" & vbCrLf &
                "                               else case" & vbCrLf &
                "                                       when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS'" & vbCrLf &
                "                                         then NFE.valor" & vbCrLf &
                "                                         else 0" & vbCrLf &
                "                                     end" & vbCrLf &
                "                           end) as Icms," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS-ST'" & vbCrLf &
                "                                then NFE.Base" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end) as BaseST," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'ICMS-ST'" & vbCrLf &
                "                                then NFE.valor" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end) as IcmsST," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'IPI'" & vbCrLf &
                "                                then NFE.valor" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end) as IPI," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) IN ('FRETES','SEGURO')" & vbCrLf &
                "                                then NFE.valor" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end) as ValorFrete," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) = 'DESCONTOS'" & vbCrLf &
                "                                then NFE.valor" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end)as ValorDesconto," & vbCrLf &
                "                       sum(case" & vbCrLf &
                "                              when ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) in ('DESP.ADUANEIRAS','OUTRASDESPESAS','ICMS SUBSTITUIC')" & vbCrLf &
                "                                then NFE.valor" & vbCrLf &
                "                                else 0" & vbCrLf &
                "                           end) as OutrasDespesas" & vbCrLf &
                "                  FROM NotasFiscais NF" & vbCrLf &
                "				 INNER JOIN NotasFiscaisXItens NFI" & vbCrLf &
                "					ON NFI.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                "				   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "				   AND NFI.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "				   AND NFI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                "				   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                "				   AND NFI.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "				   AND NFI.Nota_Id         = NF.Nota_Id" & vbCrLf &
                "                 INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf &
                "                    ON NFE.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
                "                   AND NFE.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
                "                   AND NFE.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                "                   AND NFE.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
                "                   AND NFE.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
                "                   AND NFE.Serie_Id        = NFI.Serie_Id" & vbCrLf &
                "                   AND NFE.Nota_Id         = NFI.Nota_Id" & vbCrLf &
                "                   AND NFE.Produto_Id      = NFI.Produto_Id" & vbCrLf &
                "                   AND NFE.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf &
                "                 Inner Join OperacaoXEstado OxE" & vbCrLf &
                "				     on OxE.Codigo_Id = NFI.OperacaoXEstado" & vbCrLf &
                "                 Inner Join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                "				     on OEE.Codigo_Id  = OxE.Codigo_Id" & vbCrLf &
                "                   and OEE.Encargo_Id = NFE.Encargo_Id" & vbCrLf &
                "                 INNER JOIN SubOperacoes SO" & vbCrLf &
                "                    ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf &
                "                   AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                "                 INNER JOIN Clientes cli" & vbCrLf &
                "                    ON cli.Cliente_Id    = NF.Cliente_Id" & vbCrLf &
                "                   AND cli.Endereco_Id  = NF.EndCliente_Id" & vbCrLf &
                "                 Inner Join #C100 C" & vbCrLf &
                "                    on C.Empresa_id      = nf.Empresa_id" & vbCrLf &
                "                   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf &
                "                   and c.Cliente_id      = nf.cliente_id" & vbCrLf &
                "                   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf &
                "                   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf &
                "                   and c.Nota_id         = nf.Nota_id" & vbCrLf &
                "                   and C.Serie_id        = nf.Serie_id" & vbCrLf &
                "                  Left JOIN Encargos AS Enc" & vbCrLf &
                "                    ON NFE.Encargo_Id = Enc.Encargo_id" & vbCrLf &
                "                   AND ISNULL(Enc.EncargoAgrupador,'') <> '' " & vbCrLf &
                "                 WHERE NF.Situacao = 1" & vbCrLf

        'NÃO LEVAR TIPO DE DOCUMENTO 65, SOLICITAÇÃO JÉSSICA EM 19/12/2023 - FURLAN
        If Left(Empresa(0), 8) = "05366261" OrElse
            Left(Empresa(0), 8) = "38198213" OrElse
            Left(Empresa(0), 8) = "62747840" OrElse
            Left(Empresa(0), 8) = "62780383" OrElse
            Left(Empresa(0), 8) = "63358210" OrElse
            Left(Empresa(0), 8) = "40938762" OrElse
            Left(Empresa(0), 8) = "49673784" Then

            Query &= "                   AND NF.TipoDeDocumento NOT IN(2,8,10,11,14,15,16,57,65)" & vbCrLf

        Else
            Query &= "                   AND NF.TipoDeDocumento NOT IN (15,57)" & vbCrLf
        End If

        Query &= "                   AND NF.Finalidade NOT IN (30)" & vbCrLf &
                "                   AND ISNULL(Enc.EncargoAgrupador,NFE.Encargo_Id) in ('PRODUTO','IPI','DESCONTOS','FRETES','SEGURO','DESP.ADUANEIRAS', 'ICMS','ICMS SUBSTITUIC','ICMS-ST')" & vbCrLf &
                "                   AND NOT (NF.EntradaSaida_Id = 'E'  AND OxE.CodigoFiscal IN (1933,2933))" & vbCrLf &
                "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OxE.CodigoFiscal IN (1949,2949))" & vbCrLf &
                "                   --AND NOT (NF.EntradaSaida_Id = 'S'  AND OxE.CodigoFiscal IN (5933,6933))" & vbCrLf &
                "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OxE.CodigoFiscal IN (5949,6949))" & vbCrLf &
                "                 GROUP BY NF.Empresa_id," & vbCrLf &
                "                          NF.EndEmpresa_Id," & vbCrLf &
                "                          NF.Cliente_id," & vbCrLf &
                "                          NF.EndCliente_id," & vbCrLf &
                "                          NF.EntradaSaida_id," & vbCrLf &
                "                          NF.Nota_id," & vbCrLf &
                "                          NF.Serie_id," & vbCrLf &
                "                          NFI.Produto_Id," & vbCrLf &
                "                          OxE.CodigoFiscal," & vbCrLf &
                "                          OxE.STICMS" & vbCrLf &
                "     ) as Resultado" & vbCrLf &
                " GROUP BY STICMS, CodigoFiscal, Percentual," & vbCrLf &
                "          Resultado.Empresa_Id," & vbCrLf &
                "          resultado.EndEmpresa_Id," & vbCrLf &
                "          Resultado.Cliente_Id," & vbCrLf &
                "          Resultado.EndCliente_Id," & vbCrLf &
                "          Resultado.Nota_Id," & vbCrLf &
                "          Resultado.Serie_Id," & vbCrLf &
                "          Resultado.EntradaSaida_Id;" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf



        'Query = "Select Resultado.Nota_Id, " & vbCrLf & _
        '       "       Resultado.Cliente_Id, " & vbCrLf & _
        '       "       Resultado.EndCliente_Id, " & vbCrLf & _
        '       "       Resultado.Serie_Id, " & vbCrLf & _
        '       "       Resultado.EntradaSaida_Id, " & vbCrLf & _
        '       "       SituacaoTributaria AS CST_ICMS," & vbCrLf & _
        '       "       Cfop_Id AS CFOP," & vbCrLf & _
        '       "       Percentual AS ALIQ_ICMS," & vbCrLf & _
        '       "       Sum(Produto+IPI+ValorFrete+OutrasDespesas-ValorDesconto) as VL_OPR," & vbCrLf & _
        '       "       Sum(Base) as VL_BC_ICMS," & vbCrLf & _
        '       "       Sum(Icms) as VL_ICMS," & vbCrLf & _
        '       "       Sum(Ipi) as VL_IPI" & vbCrLf & _
        '       "  into #C190" & vbCrLf & _
        '       "  FROM (" & vbCrLf & _
        '       "                Select NF.Empresa_Id," & vbCrLf & _
        '       "                       NF.EndEmpresa_Id," & vbCrLf & _
        '       "                       NF.Cliente_Id, " & vbCrLf & _
        '       "                       NF.EndCliente_Id, " & vbCrLf & _
        '       "                       NF.Nota_Id, " & vbCrLf & _
        '       "                       NF.Serie_Id, " & vbCrLf & _
        '       "                       NF.EntradaSaida_Id, " & vbCrLf & _
        '       "                       NFI.Produto_Id, " & vbCrLf & _
        '       "                       NFI.Cfop_Id," & vbCrLf & _
        '       "                       NFE.Situacaotributaria," & vbCrLf & _
        '       "                       sum(case" & vbCrLf & _
        '       "                              when NFE.Encargo_Id = 'PRODUTO'" & vbCrLf & _
        '       "                                then NFE.valor" & vbCrLf & _
        '       "                                else 0" & vbCrLf & _
        '       "                           end) as Produto," & vbCrLf & _
        '       "                       SUM(CASE" & vbCrLf & _
        '       "                              when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '       "                                then 0" & vbCrLf & _
        '       "                                else case" & vbCrLf & _
        '       "                                       when NFE.Encargo_Id = 'ICMS'" & vbCrLf & _
        '       "                                         then NFE.Base" & vbCrLf & _
        '       "                                         else 0" & vbCrLf & _
        '       "                                     end" & vbCrLf & _
        '       "                           end) as Base," & vbCrLf

        'If Empresa(0).ToString.Contains("04440724") Then
        '    Query &= "                        MAX(CASE" & vbCrLf & _
        '           "						      when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '           "						    	then 0" & vbCrLf & _
        '           "						     	else case" & vbCrLf & _
        '           "						     	       when NFE.Encargo_Id = 'ICMS'" & vbCrLf & _
        '           "						     	         then NFE.Percentual" & vbCrLf & _
        '           "						    			 else 0" & vbCrLf & _
        '           "						    		 end" & vbCrLf & _
        '           "					        end) as Percentual," & vbCrLf
        'Else
        '    Query &= "                        MAX(CASE" & vbCrLf & _
        '           "						      when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '           "						    	then 0" & vbCrLf & _
        '           "						     	else case" & vbCrLf & _
        '           "						     	       when NFE.Encargo_Id = 'ICMS'" & vbCrLf & _
        '           "						     	         then Case" & vbCrLf & _
        '           "						    		            when isnull(NFE.PercentualExibicao,0) = 0" & vbCrLf & _
        '           "						    		              then NFE.Percentual" & vbCrLf & _
        '           "						    			          else NFE.PercentualExibicao" & vbCrLf & _
        '           "						    			      end" & vbCrLf & _
        '           "						    		   else 0" & vbCrLf & _
        '           "						    		 end" & vbCrLf & _
        '           "					         end) as Percentual," & vbCrLf
        'End If

        'Query &= "                       SUM(CASE" & vbCrLf & _
        '       "                             when ((NFI.CFOP_Id IN (1551,2551)) or (NFI.CFOP_Id IN (1653,2653) and NFE.SituacaoTributaria in(10,60))) AND YEAR('2014/03/31') > 2011" & vbCrLf & _
        '       "                               then 0" & vbCrLf & _
        '       "                               else case" & vbCrLf & _
        '       "                                       when NFE.Encargo_Id = 'ICMS'" & vbCrLf & _
        '       "                                         then NFE.valor" & vbCrLf & _
        '       "                                         else 0" & vbCrLf & _
        '       "                                     end" & vbCrLf & _
        '       "                           end) as Icms," & vbCrLf & _
        '       "                       sum(case" & vbCrLf & _
        '       "                              when NFE.Encargo_Id = 'IPI'" & vbCrLf & _
        '       "                                then NFE.valor" & vbCrLf & _
        '       "                                else 0" & vbCrLf & _
        '       "                           end) as IPI," & vbCrLf & _
        '       "                       sum(case" & vbCrLf & _
        '       "                              when NFE.Encargo_Id = 'FRETES'" & vbCrLf & _
        '       "                                then NFE.valor" & vbCrLf & _
        '       "                                else 0" & vbCrLf & _
        '       "                           end) as ValorFrete," & vbCrLf & _
        '       "                       sum(case" & vbCrLf & _
        '       "                              when NFE.Encargo_Id = 'DESCONTOS'" & vbCrLf & _
        '       "                                then NFE.valor" & vbCrLf & _
        '       "                                else 0" & vbCrLf & _
        '       "                           end)as ValorDesconto," & vbCrLf & _
        '       "                       sum(case" & vbCrLf & _
        '       "                              when NFE.Encargo_Id in ('DESP.ADUANEIRAS','OUTRASDESPESAS')" & vbCrLf & _
        '       "                                then NFE.valor" & vbCrLf & _
        '       "                                else 0" & vbCrLf & _
        '       "                           end) as OutrasDespesas" & vbCrLf & _
        '       "                  FROM NotasFiscais NF" & vbCrLf & _
        '       "				 INNER JOIN NotasFiscaisXItens NFI" & vbCrLf & _
        '       "					ON NFI.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
        '       "				   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
        '       "				   AND NFI.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
        '       "				   AND NFI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
        '       "				   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
        '       "				   AND NFI.Serie_Id        = NF.Serie_Id" & vbCrLf & _
        '       "				   AND NFI.Nota_Id         = NF.Nota_Id" & vbCrLf & _
        '       "                 INNER JOIN NotasFiscaisXEncargos NFE" & vbCrLf & _
        '       "                    ON NFE.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
        '       "                   AND NFE.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
        '       "                   AND NFE.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
        '       "                   AND NFE.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
        '       "                   AND NFE.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
        '       "                   AND NFE.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
        '       "                   AND NFE.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
        '       "                   AND NFE.Produto_Id      = NFI.Produto_Id" & vbCrLf & _
        '       "                   AND NFE.CFOP_Id         = NFI.CFOP_Id" & vbCrLf & _
        '       "                   AND NFE.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
        '       "                 INNER JOIN SubOperacoes SO" & vbCrLf & _
        '       "                    ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf & _
        '       "                   AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
        '       "                 Inner Join #C100 C" & vbCrLf & _
        '       "                    on C.Empresa_id      = nf.Empresa_id" & vbCrLf & _
        '       "                   and C.EndEmpresa_id   = nf.EndEmpresa_id" & vbCrLf & _
        '       "                   and c.Cliente_id      = nf.cliente_id" & vbCrLf & _
        '       "                   and c.endcliente_Id   = nf.endcliente_id" & vbCrLf & _
        '       "                   and c.entradasaida_id = nf.entradasaida_id" & vbCrLf & _
        '       "                   and c.Nota_id         = nf.Nota_id" & vbCrLf & _
        '       "                   and C.Serie_id        = nf.Serie_id" & vbCrLf & _
        '       "                INNER JOIN Encargos AS Enc " & vbCrLf & _
        '       "                   ON NFE.Encargo_Id = Enc.Encargo_id " & vbCrLf & _
        '       "                   AND ISNULL(Enc.Encargo_id,'') <> '' " & vbCrLf & _
        '       "                 WHERE NF.Situacao = 1" & vbCrLf & _
        '       "                   AND (NFE.Encargo_Id in ('PRODUTO','IPI','DESCONTOS','FRETES','DESP.ADUANEIRAS', 'ICMS'))" & vbCrLf & _
        '       "                   AND NOT (NF.EntradaSaida_Id = 'E'  AND NFI.CFOP_Id IN (1933,2933))" & vbCrLf & _
        '       "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NFI.CFOP_Id IN (1949,2949))" & vbCrLf & _
        '       "                   --AND NOT (NF.EntradaSaida_Id = 'S'  AND NFI.CFOP_Id IN (5933,6933))" & vbCrLf & _
        '       "                   AND NOT (SO.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND NFI.CFOP_Id IN (5949,6949))" & vbCrLf & _
        '       "                 GROUP BY NF.Empresa_id," & vbCrLf & _
        '       "                          NF.EndEmpresa_Id," & vbCrLf & _
        '       "                          NF.Cliente_id," & vbCrLf & _
        '       "                          NF.EndCliente_id," & vbCrLf & _
        '       "                          NF.EntradaSaida_id," & vbCrLf & _
        '       "                          NF.Nota_id," & vbCrLf & _
        '       "                          NF.Serie_id," & vbCrLf & _
        '       "                          NFI.Produto_Id, " & vbCrLf & _
        '       "                          NFI.Cfop_Id," & vbCrLf & _
        '       "                          NFE.Situacaotributaria" & vbCrLf & _
        '       "     ) as Resultado" & vbCrLf & _
        '       " GROUP BY SituacaoTributaria, Cfop_Id, Percentual, " & vbCrLf & _
        '       "          Resultado.Empresa_Id," & vbCrLf & _
        '       "          resultado.EndEmpresa_Id," & vbCrLf & _
        '       "          Resultado.Cliente_Id," & vbCrLf & _
        '       "          Resultado.EndCliente_Id," & vbCrLf & _
        '       "          Resultado.Nota_Id," & vbCrLf & _
        '       "          Resultado.Serie_Id," & vbCrLf & _
        '       "          Resultado.EntradaSaida_Id " & vbCrLf
        Return Query
    End Function

    Private Sub CompoeRegistroC101(ByRef ds As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef RegistroC101 As Integer, ByRef RegistroGeral As Integer,
                                  C100NF As Integer, C100Cliente As String, C100EndCliente As Integer, C100Serie As String, C100EntradaSaida As String)

        If ds.Tables("C101").Rows.Count > 0 Then
            'For Each dr1 As DataRow In ds1.Tables(0).Rows
            For Each dr1 As DataRow In ds.Tables("C101").Select("Nota_id = " & C100NF & " AND Cliente_id = '" & C100Cliente & "' AND EndCliente_id = " &
                                                                  C100EndCliente & " AND Serie_id = '" & C100Serie & "' AND EntradaSaida_id = '" & C100EntradaSaida & "'")
                With dr1
                    linha = "|C101"
                    linha &= "|" & .Item("Valor")
                    linha &= "|0,00"
                    linha &= "|0,00"
                    linha &= "|"

                    strm.WriteLine(linha)
                    RegistroC101 += 1
                    RegistroGeral += 1

                End With
            Next
        End If
    End Sub

    Private Sub CompoeRegistroC170(ByRef ds As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef RegistroC170 As Integer, ByRef RegistroGeral As Integer, _
                                  C100NF As Integer, C100Cliente As String, C100EndCliente As Integer, C100Serie As String, C100EntradaSaida As String)

        Dim intSequencia = 0

        If ds.Tables("C170").Rows.Count > 0 Then
            'For Each dr1 As DataRow In ds1.Tables(0).Rows
            For Each dr1 As DataRow In ds.Tables("C170").Select("Nota_id = " & C100NF & " AND Cliente_id = '" & C100Cliente & "' AND EndCliente_id = " & _
                                                                  C100EndCliente & " AND Serie_id = '" & C100Serie & "' AND EntradaSaida_id = '" & C100EntradaSaida & "'")
                With dr1
                    linha = "|C170"
                    intSequencia += 1
                    linha &= "|" & intSequencia
                    linha &= "|" & Trim(.Item("Item"))
                    linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                    linha &= "|" & .Item("QuantidadeFiscal")
                    linha &= "|" & Trim(.Item("Unidade"))
                    linha &= "|" & .Item("ValorTotal")
                    linha &= "|" & .Item("ValorDesconto")   'Descontos
                    If .Item("EstoqueFiscal") = "S" Then
                        linha &= "|0"
                    Else
                        linha &= "|1"
                    End If
                    linha &= "|" & Funcoes.AlinharDireita(.Item("SituacaoTributaria"), 3, "0")
                    linha &= "|" & .Item("CFOP")
                    linha &= "|" & .Item("Operacao") & .Item("SubOperacao")
                    linha &= "|" & .Item("BaseIcms")
                    linha &= "|" & String.Format("{0:N2}", .Item("Aliquota"))
                    linha &= "|" & .Item("ValorIcms")
                    linha &= "|" & .Item("BaseIcmsST")                           'Valor Base de Icms Substituição Tributária
                    linha &= "|" & String.Format("{0:N2}", .Item("AliquotaST"))  'Aliquota de Icms por Substituição Tributária
                    linha &= "|" & .Item("ValorIcmsST")                          'Valor de Icms por Substituição Tributária
                    linha &= "|0"                           'Indicador de período de apuração do IPI
                    linha &= "|"                            'Valor de Icms por Substituição Tributária
                    linha &= "|"                            'Código de enquadramento legal do IPI, conforme tabela indicada no item 4.5.3
                    linha &= "|" & .Item("VL_BC_IPI")
                    linha &= "|" & String.Format("{0:N2}", .Item("ALI_IPI"))
                    linha &= "|" & .Item("VL_IPI")

                    linha &= "|" & .Item("CST_PIS")          'Código Da situação tributaria referente ao PIS, conforme tabela indicada no item 4.3.4                            
                    linha &= "|" & .Item("VL_BC_PIS")

                    If PeriodoFinal.Year > 2011 Then
                        linha &= "|" & String.Format("{0:N4}", .Item("ALI_PIS"))
                    Else
                        linha &= "|" & String.Format("{0:N2}", .Item("ALI_PIS"))
                    End If

                    linha &= "|" & .Item("QUANT_BC_PIS")

                    If PeriodoFinal.Year > 2011 Then
                        linha &= "|" & String.Format("{0:N4}", .Item("ALI_PISR"))
                    Else
                        linha &= "|" & String.Format("{0:N2}", .Item("ALI_PISR"))
                    End If

                    linha &= "|" & .Item("VL_PIS")

                    linha &= "|" & Format(.Item("CST_PIS"), "00")       'Código Da situação tributaria referente ao COFINS, conforme tabela indicada no item 4.3.5
                    linha &= "|" & .Item("VL_BC_COFINS")

                    If PeriodoFinal.Year > 2011 Then
                        linha &= "|" & String.Format("{0:N4}", .Item("ALI_COFINS"))
                    Else
                        linha &= "|" & String.Format("{0:N2}", .Item("ALI_COFINS"))
                    End If

                    linha &= "|" & .Item("QUANT_BC_COFINS")

                    If PeriodoFinal.Year > 2011 Then
                        linha &= "|" & String.Format("{0:N4}", .Item("ALI_COFINSR"))
                    Else
                        linha &= "|" & String.Format("{0:N2}", .Item("ALI_COFINSR"))
                    End If

                    linha &= "|" & .Item("VL_COFINS")
                    linha &= "|"
                    linha &= "|0,00"
                    linha &= "|"

                    'If .Item("QuantidadeFiscal") > 0 Then
                    '    strm.WriteLine(linha)
                    '    RegistroC170 += 1
                    '    RegistroGeral += 1
                    'End If

                    strm.WriteLine(linha)
                    RegistroC170 += 1
                    RegistroGeral += 1

                End With
            Next
        End If
    End Sub

    Private Sub CompoeRegistroC172(ByRef ds As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef RegistroC172 As Integer, ByRef RegistroGeral As Integer, _
                                  C100NF As Integer, C100Cliente As String, C100EndCliente As Integer, C100Serie As String, C100EntradaSaida As String)

        If ds.Tables("C172").Rows.Count > 0 Then
            For Each dr1 As DataRow In ds.Tables("C172").Select("Nota_id = " & C100NF & " AND Cliente_id = '" & C100Cliente & "' AND EndCliente_id = " & _
                                                                  C100EndCliente & " AND Serie_id = '" & C100Serie & "' AND EntradaSaida_id = '" & C100EntradaSaida & "'")
                With dr1
                    linha = "|C172"
                    linha &= "|" & .Item("VL_BC_ISSQN") 'Base
                    linha &= "|" & String.Format("{0:N2}", .Item("ALIQ_ISSQN"))
                    linha &= "|" & .Item("VL_ISSQN")
                    linha &= "|"

                    'If .Item("VL_ISSQN") > 0 Then
                    '    strm.WriteLine(linha)
                    '    RegistroC172 += 1
                    '    RegistroGeral += 1
                    'End If

                    strm.WriteLine(linha)
                    RegistroC172 += 1
                    RegistroGeral += 1

                End With
            Next
        End If
    End Sub

    Private Sub CompoeRegistroC190(ByRef ds As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef RegistroC190 As Integer, ByRef RegistroGeral As Integer, _
                                  C100NF As Integer, C100Cliente As String, C100EndCliente As Integer, C100Serie As String, C100EntradaSaida As String)

        If ds.Tables("C190").Rows.Count > 0 Then
            'For Each dr2 As DataRow In ds2.Tables(0).Rows
            For Each dr2 As DataRow In ds.Tables("C190").Select("Nota_id = " & C100NF & " AND Cliente_id = '" & C100Cliente & "' AND EndCliente_id = " & _
                                                             C100EndCliente & " AND Serie_id = '" & C100Serie & "' AND EntradaSaida_id = '" & C100EntradaSaida & "'")
                With dr2
                    linha = "|C190"
                    linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                    linha &= "|" & .Item("CFOP")
                    linha &= "|" & String.Format("{0:N2}", .Item("ALIQ_ICMS"))
                    linha &= "|" & .Item("VL_OPR")
                    linha &= "|" & .Item("VL_BC_ICMS")
                    linha &= "|" & .Item("VL_ICMS")
                    linha &= "|" & .Item("VL_BC_ICMSST")
                    linha &= "|" & .Item("VL_ICMSST")

                    If CInt(.Item("CST_ICMS")) = 20 OrElse CInt(.Item("CST_ICMS")) = 70 OrElse CInt(.Item("CST_ICMS")) = 620 OrElse CInt(.Item("CST_ICMS")) = 670 Then
                        If .Item("VL_BC_ICMS") <> 0 And (.Item("VL_OPR") > .Item("VL_BC_ICMS")) Then
                            linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                        Else
                            linha &= "|" & .Item("VL_BC_ICMS")
                        End If
                    Else
                        linha &= "|0"
                    End If

                    linha &= "|0" & .Item("VL_IPI")
                    linha &= "|"
                    linha &= "|"
                End With

                strm.WriteLine(linha)
                RegistroC190 += 1
                RegistroGeral += 1
            Next
        End If
    End Sub

    Private Function ConsultaRegistroD100() As DataSet
        Sql = "/*Registro D100 - Nota Fiscal de Serviço de Transporte*/" & vbCrLf &
              "SELECT NF.Empresa_id, NF.EndEmpresa_Id, NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
              "	      NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id," & vbCrLf &
              "	      ISNULL(NFER.ChaveNfe, '') AS ChaveNfe," & vbCrLf &
              "	      NF.DataDaNota," & vbCrLf &
              "       NF.Movimento," & vbCrLf &
              "	      NF.Situacao," & vbCrLf &
              "	      NF.TipoDeDocumento," & vbCrLf &
              "       CASE" & vbCrLf &
              "         WHEN ISNULL(DocXML.Tipo,'') = 'Mic'" & vbCrLf &
              "           THEN 'S'" & vbCrLf &
              "           ELSE NF.NossaEmissao" & vbCrLf &
              "       END AS NossaEmissao," & vbCrLf &
              "	      NF.Eletronica," & vbCrLf &
              "	      Max(Oe.CodigoFiscal)           AS Cfop," & vbCrLf &
              "       Sum(ISNULL(enc.ValorTotal, 0)) AS ValorTotal," & vbCrLf &
              "       Sum(ISNULL(enc.BaseIcms, 0))   AS BaseIcms," & vbCrLf &
              "       0                              AS Aliquota," & vbCrLf &
              "       Sum(ISNULL (enc.ValorIcms, 0)) AS ValorIcms" & vbCrLf &
              "       INTO #NOTAS " & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              " Inner Join NotasFiscaisXItens nfi" & vbCrLf &
              "    on NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
              "   and NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
              "   and NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
              " Inner Join SubOperacoes SO" & vbCrLf &
              "    on SO.Operacao_Id     = NF.Operacao" & vbCrLf &
              "   and SO.SubOperacoes_Id = NF.SubOperacao" & vbCrLf &
              " Inner Join OperacaoXEstado OE" & vbCrLf &
              "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf &
              "  LEFT JOIN (" & vbCrLf &
              "             Select nfe.Empresa_Id," & vbCrLf &
              "			        nfe.EndEmpresa_Id," & vbCrLf &
              "					nfe.Cliente_Id," & vbCrLf &
              "					nfe.EndCliente_Id," & vbCrLf &
              "					nfe.EntradaSaida_Id," & vbCrLf &
              "					nfe.Serie_Id," & vbCrLf &
              "					nfe.Nota_Id," & vbCrLf &
              "					nfe.Produto_id," & vbCrLf &
              "					nfe.Sequencia_id," & vbCrLf &
              "                    SUM(Case" & vbCrLf &
              "                          when isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'PRODUTO'" & vbCrLf &
              "                            then nfe.Valor" & vbCrLf &
              "                            else 0" & vbCrLf &
              "                        end) as ValorTotal," & vbCrLf &
              "                    SUM(Case" & vbCrLf &
              "                          when isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf &
              "                            then nfe.BASE" & vbCrLf &
              "                            else 0" & vbCrLf &
              "                        end) as BaseIcms," & vbCrLf &
              "                    SUM(Case" & vbCrLf &
              "                          when isnull(e.EncargoAgrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf &
              "                            then nfe.Valor" & vbCrLf &
              "                            else 0" & vbCrLf &
              "                        end) as ValorIcms" & vbCrLf &
              "               from NotasFiscaisxEncargos nfe" & vbCrLf &
              "               Left Join Encargos e" & vbCrLf &
              "			        on e.Encargo_id = nfe.Encargo_Id" & vbCrLf &
              "                and isnull(e.EncargoAgrupador,'') <> ''" & vbCrLf &
              "              Where isnull(e.EncargoAgrupador,nfe.Encargo_Id) in ('PRODUTO','ICMS')" & vbCrLf &
              "              Group by nfe.Empresa_Id," & vbCrLf &
              "			           nfe.EndEmpresa_Id," & vbCrLf &
              "					   nfe.Cliente_Id," & vbCrLf &
              "					   nfe.EndCliente_Id," & vbCrLf &
              "					   nfe.EntradaSaida_Id," & vbCrLf &
              "					   nfe.Serie_Id," & vbCrLf &
              "					   nfe.Nota_Id," & vbCrLf &
              "					   nfe.Produto_id," & vbCrLf &
              "					   nfe.Sequencia_id" & vbCrLf &
              "             ) Enc" & vbCrLf &
              "    ON nfi.Empresa_Id      = Enc.Empresa_Id" & vbCrLf &
              "   AND nfi.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf &
              "   AND nfi.Cliente_Id      = Enc.Cliente_Id" & vbCrLf &
              "   AND nfi.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf &
              "   AND nfi.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf &
              "   AND nfi.Serie_Id        = Enc.Serie_Id" & vbCrLf &
              "   AND nfi.Nota_Id         = Enc.Nota_Id" & vbCrLf &
              "   and nfi.Produto_id      = Enc.Produto_id" & vbCrLf &
              "   and nfi.sequencia_id    = Enc.sequencia_id" & vbCrLf &
              "  LEFT JOIN NFERealizadas NFER" & vbCrLf &
              "    ON NF.Empresa_Id      = NFER.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFER.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = NFER.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = NFER.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFER.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = NFER.Serie_Id" & vbCrLf &
              "   AND NF.Nota_Id         = NFER.Nota_Id" & vbCrLf &
              "  LEFT JOIN DocumentoXML DocXML" & vbCrLf &
              "    ON DocXML.Empresa_Id  = NF.Empresa_Id" & vbCrLf &
              "    AND DocXML.Cliente_Id = NF.Cliente_Id" & vbCrLf &
              "    AND DocXML.Serie_Id   = NF.Serie_Id" & vbCrLf &
             "     AND DocXML.Numero_Id  = NF.Nota_Id" & vbCrLf &
              " WHERE NF.Serie_Id NOT IN ('D', 'F', 'REC')" & vbCrLf &
              "   AND NF.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
              "   AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf &
              "   AND NF.Situacao not in(9,10)" & vbCrLf &
              "   AND ((NF.TipoDeDocumento in (2,8,10,11,14,57,58,67)) OR (SO.Descricao = 'COMPLEMENTACAO DE ICMS '))" & vbCrLf &
              "   AND NF.operacao > 0" & vbCrLf &
              "   And ((OE.CodigoFiscal Between 1350 and 1360) OR" & vbCrLf &
              "        (OE.CodigoFiscal Between 2350 and 2360) OR" & vbCrLf &
              "        (OE.CodigoFiscal Between 3350 and 3360) OR" & vbCrLf &
              "        (OE.CodigoFiscal IN (1932,2932, 5932, 6932))        OR" & vbCrLf &
              "        (OE.CodigoFiscal Between 5350 and 5360) OR" & vbCrLf &
              "        (OE.CodigoFiscal Between 6350 and 6360) OR" & vbCrLf &
              "        (OE.CodigoFiscal Between 7350 and 7360))" & vbCrLf &
              " group by NF.Empresa_id, NF.EndEmpresa_Id, NF.Cliente_Id, NF.EndCliente_Id," & vbCrLf &
              "	      NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id," & vbCrLf &
              "	      ISNULL(NFER.ChaveNfe, ''), " & vbCrLf &
              "	      NF.DataDaNota," & vbCrLf &
              "       NF.Movimento," & vbCrLf &
              "	      NF.Situacao," & vbCrLf &
              "	      NF.TipoDeDocumento," & vbCrLf &
              "       ISNULL(DocXML.Tipo,'')," & vbCrLf &
              "	      NF.NossaEmissao," & vbCrLf &
              "	      NF.Eletronica" & vbCrLf

        Sql &= "SELECT #NOTAS.Cliente_Id, #NOTAS.EndCliente_Id, Nota_Id, Serie_Id, EntradaSaida_Id, ChaveNfe, DataDaNota, Movimento, " & vbCrLf & _
            "	#NOTAS.Situacao, TipoDeDocumento, NossaEmissao, Eletronica, Cfop, ValorTotal, BaseIcms, Aliquota, ValorIcms, " & vbCrLf & _
            "   (SELECT TOP 1 CODIGOIBGECOMPLETO FROM MUNICIPIOS MUNI" & vbCrLf & _
            "	WHERE MUNI.ESTADO_ID = CLIEMP.ESTADO" & vbCrLf & _
            "	  AND MUNI.CODIGO_ID = CLIEMP.CODIGODOMUNICIPIO) AS CODMUNEMP," & vbCrLf & _
            "   (SELECT TOP 1 CODIGOIBGECOMPLETO FROM MUNICIPIOS MUNI" & vbCrLf & _
            "	WHERE MUNI.ESTADO_ID = CLINOTA.ESTADO" & vbCrLf & _
            "	  AND MUNI.CODIGO_ID = CLINOTA.CODIGODOMUNICIPIO) AS CODMUNCLI" & vbCrLf & _
            "FROM #NOTAS" & vbCrLf & _
            "       INNER JOIN CLIENTES CLIEMP" & vbCrLf & _
            "               ON CLIEMP.CLIENTE_ID   = #NOTAS.Empresa_Id" & vbCrLf & _
            "               AND CLIEMP.ENDERECO_ID = #NOTAS.EndEmpresa_Id" & vbCrLf & _
            "       INNER JOIN CLIENTES CLINOTA" & vbCrLf & _
            "               ON CLINOTA.CLIENTE_ID   = #NOTAS.Cliente_Id" & vbCrLf & _
            "       AND CLINOTA.ENDERECO_ID = #NOTAS.EndCliente_Id "

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Function ConsultaRegistroD190(ByVal sqlWhereNotaD100 As String) As DataSet
        Sql = "/*Registro D190 - Registro Analitico dos Documentos*/" & vbCrLf & _
              " SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id" & vbCrLf & _
              "   Into #TNF" & vbCrLf & _
              "   FROM NotasFiscais" & vbCrLf & _
              "  Where Empresa_id ='" & Empresa(0) & "'" & vbCrLf & _
              "    AND ( " & sqlWhereNotaD100 & ")" & vbCrLf & _
              "                                   " & vbCrLf & _
              "SELECT nfi.Empresa_Id, nfi.EndEmpresa_Id," & vbCrLf & _
              "       nfi.Cliente_Id, nfi.EndCliente_Id," & vbCrLf & _
              "	      nfi.Nota_id, nfi.Serie_Id, nfi.EntradaSaida_Id," & vbCrLf & _
              "	      OE.codigoFiscal         as CFOP," & vbCrLf & _
              "	      OE.STICMS               as CST_ICMS," & vbCrLf & _
              "	      Sum(Enc.Produto)        as VL_OPR," & vbCrLf & _
              "	      Sum(Enc.Base)           as VL_BC_ICMS," & vbCrLf & _
              "	      sum(Enc.Percentual)     as ALIQ_ICMS," & vbCrLf & _
              "	      Sum(Enc.Icms)           as VL_ICMS," & vbCrLf & _
              "	      Sum(Enc.Ipi)            as VL_IPI" & vbCrLf & _
              "  FROM Notasfiscaisxitens nfi" & vbCrLf & _
              " Inner join #TNF nf" & vbCrLf & _
              "    on nfi.Empresa_Id      = nf.empresa_id" & vbCrLf & _
              "   and nfi.EndEmpresa_Id   = nf.endempresa_id" & vbCrLf & _
              "   and nfi.Cliente_Id      = nf.cliente_id" & vbCrLf & _
              "   and nfi.endcliente_id   = nf.endcliente_id" & vbCrLf & _
              "   and nfi.EntradaSaida_Id = nf.entradasaida_id" & vbCrLf & _
              "   and nfi.Nota_Id         = nf.nota_id" & vbCrLf & _
              "   and nfi.Serie_Id        = nf.serie_id" & vbCrLf & _
              " Inner join OperacaoxEstado OE" & vbCrLf & _
              "   on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf & _
              " Inner Join" & vbCrLf & _
              "        (SELECT nfe.Empresa_Id, nfe.EndEmpresa_Id," & vbCrLf & _
              "	               nfe.Cliente_Id, nfe.EndCliente_Id," & vbCrLf & _
              "			       nfe.Nota_id, nfe.Serie_Id, nfe.EntradaSaida_Id," & vbCrLf & _
              "			       nfe.Produto_Id, nfe.Sequencia_id," & vbCrLf & _
              "                sum(Case" & vbCrLf & _
              "			             when isnull(Enc.encargoagrupador,nfe.Encargo_Id) = 'PRODUTO'" & vbCrLf & _
              "				          then valor" & vbCrLf & _
              "				          else 0" & vbCrLf & _
              "			           end) as Produto," & vbCrLf & _
              "                sum(case" & vbCrLf & _
              "			             when isnull(Enc.encargoagrupador,nfe.Encargo_Id) = 'IPI'" & vbCrLf & _
              "				          then valor" & vbCrLf & _
              "				          else 0" & vbCrLf & _
              "			           end) as IPI," & vbCrLf & _
              "			       sum(case" & vbCrLf & _
              "			             when isnull(Enc.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
              "			              then Base" & vbCrLf & _
              "			              else 0" & vbCrLf & _
              "			           end) as Base," & vbCrLf & _
              "			       sum(case" & vbCrLf & _
              "			             when isnull(Enc.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
              "			              then Percentual" & vbCrLf & _
              "			              else 0" & vbCrLf & _
              "			           end) as Percentual," & vbCrLf & _
              "			       sum(case" & vbCrLf & _
              "			             when isnull(Enc.encargoagrupador,nfe.Encargo_Id) = 'ICMS'" & vbCrLf & _
              "			              then valor" & vbCrLf & _
              "			              else 0" & vbCrLf & _
              "			           end) as Icms" & vbCrLf & _
              "           FROM NotasFiscaisXEncargos nfe" & vbCrLf & _
              "		      Left JOIN Encargos AS Enc" & vbCrLf & _
              "             ON nfe.Encargo_Id = Enc.Encargo_id" & vbCrLf & _
              "            AND ISNULL(Enc.EncargoAgrupador,'') <> ''" & vbCrLf & _
              "          WHERE isnull(Enc.encargoagrupador,nfe.Encargo_Id) in ('PRODUTO','IPI','ICMS')" & vbCrLf & _
              "          GROUP BY nfe.Empresa_Id, nfe.EndEmpresa_Id," & vbCrLf & _
              "	               nfe.Cliente_Id, nfe.EndCliente_Id," & vbCrLf & _
              "			       nfe.Nota_id, nfe.Serie_Id, nfe.EntradaSaida_Id," & vbCrLf & _
              "			       nfe.Produto_Id, nfe.Sequencia_id" & vbCrLf & _
              "        ) as Enc" & vbCrLf & _
              "   	on Enc.Empresa_Id      = nfi.empresa_id" & vbCrLf & _
              "    and Enc.EndEmpresa_Id   = nfi.endempresa_id" & vbCrLf & _
              "    and Enc.Cliente_Id      = nfi.cliente_id" & vbCrLf & _
              "    and Enc.endcliente_id   = nfi.endcliente_id" & vbCrLf & _
              "    and Enc.EntradaSaida_Id = nfi.entradasaida_id" & vbCrLf & _
              "    and Enc.Nota_Id         = nfi.nota_id" & vbCrLf & _
              "    and Enc.Serie_Id        = nfi.serie_id" & vbCrLf & _
              "    and Enc.Produto_Id      = nfi.Produto_Id" & vbCrLf & _
              "    and Enc.Sequencia_id    = nfi.Sequencia_id" & vbCrLf & _
              "  Group by nfi.Empresa_Id, nfi.EndEmpresa_Id," & vbCrLf & _
              "           nfi.Cliente_Id, nfi.EndCliente_Id," & vbCrLf & _
              "	          nfi.Nota_id, nfi.Serie_Id, nfi.EntradaSaida_Id," & vbCrLf & _
              "	          OE.codigoFiscal," & vbCrLf & _
              "           OE.STICMS" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroD190(ByRef ds1 As DataSet, ByRef strm As StreamWriter, ByRef linha As String, ByRef RegistroD190 As Integer, ByRef RegistroGeral As Integer, _
                                   D100NF As Integer, D100Cliente As String, D100EndCliente As Integer, D100Serie As String, D100EntradaSaida As String)


        If ds1.Tables(0).Rows.Count > 0 Then
            'For Each dr1 As DataRow In ds1.Tables(0).Rows
            For Each dr1 As DataRow In ds1.Tables(0).Select("Nota_id = " & D100NF & " AND Cliente_id = '" & D100Cliente & "' AND EndCliente_id = " & _
                                                                 D100EndCliente & " AND Serie_id = '" & D100Serie & "' AND EntradaSaida_id = '" & D100EntradaSaida & "'")
                With dr1
                    linha = "|D190"
                    linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                    linha &= "|" & .Item("CFOP")
                    linha &= "|" & String.Format("{0:N2}", .Item("ALIQ_ICMS"))
                    linha &= "|" & .Item("VL_OPR")
                    linha &= "|" & .Item("VL_BC_ICMS")
                    linha &= "|" & .Item("VL_ICMS")

                    'If .Item("VL_BC_ICMS") <> 0 And (.Item("VL_OPR") > .Item("VL_BC_ICMS")) Then
                    '    linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                    'Else
                    '   linha &= "|0" '& .Item("VL_RED_BC")
                    'End If

                    If CInt(.Item("CST_ICMS")) = 20 OrElse CInt(.Item("CST_ICMS")) = 70 OrElse CInt(.Item("CST_ICMS")) = 620 OrElse CInt(.Item("CST_ICMS")) = 670 Then
                        linha &= "|" & .Item("VL_BC_ICMS")
                    Else
                        linha &= "|0" '& .Item("VL_RED_BC")
                    End If

                    linha &= "|"                        ' Códifo da observação do lançamento fiscal
                    linha &= "|"
                End With
                strm.WriteLine(linha)
                RegistroD190 += 1
                RegistroGeral += 1
            Next
        End If
    End Sub

    Private Function GetLeiaute(ByRef Ano As Integer, ByRef Mes As Integer) As String
        Dim Leiaute As String = String.Empty

        Select Case Ano
            Case Is < 2011
                Leiaute = "003"
            Case Is < 2012
                Leiaute = "004"
            Case 2012
                If Mes < 7 Then
                    Leiaute = "005"
                Else
                    Leiaute = "006"
                End If
            Case Is < 2014
                Leiaute = "007"
            Case Is = 2014
                Leiaute = "008"
            Case Is = 2015
                Leiaute = "009"
            Case Is = 2016
                Leiaute = "010"
            Case Is = 2017
                Leiaute = "011"
            Case Is = 2018
                Leiaute = "012"
            Case Is = 2019
                Leiaute = "013"
            Case Is = 2020
                Leiaute = "014"
            Case Is = 2021
                Leiaute = "015"
            Case Is = 2022
                Leiaute = "016"
            Case Is = 2023
                Leiaute = "017"
            Case Is = 2024
                Leiaute = "018"
            Case Is = 2025
                Leiaute = "019"
            Case Is = 2026
                Leiaute = "020"
            Case Else
                Leiaute = "019"
        End Select

        Return Leiaute

    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    Processar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SpedFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class