Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RegistrosDeIPI
    Inherits BasePage

    Dim Sql As String
    Dim Sql2 As String
    Dim Row As DataRow

    Dim ds As New DataSet
    Dim dss As New DataSet

    Dim SqlArray As New ArrayList
    Dim Empresa() As String

    Dim Descricao As String = ""
    Dim Mensagem As String = ""

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date

    Dim Processo As Integer = 0
    Dim Livro As Integer = 0
    Dim Folha As Integer = 0
    Dim Opcao As String = ""

    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Fiscal.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("RegistrosDeIPI", "ACESSAR") Then
                    CarregarUnidade()
                    CriaOpcoes()
                    CarregarProcesso()
                    HID.Value = Guid.NewGuid().ToString
                    ucRegistrosIpiAjustaResumo.SetarHID(HID.Value)
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
        ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidade.SelectedIndexChanged
        Try
            Limpar()
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If DdlEmpresa.SelectedIndex > 0 Then
                CarregarProcesso()
            Else
                Limpar()
                DdlUnidade.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlProcesso_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If DdlProcesso.SelectedIndex > 0 Then
                Empresa = DdlEmpresa.SelectedValue.Split("-")
                HttpContext.Current.Session("ProcessoIpi") = DdlProcesso.SelectedValue

                Sql = "SELECT * " & vbCrLf & _
                      "  FROM ProcessoRAIPI " & vbCrLf & _
                      " WHERE Empresa_Id  ='" & Empresa(0) & "'" & vbCrLf & _
                      "   And Processo_Id = " & DdlProcesso.SelectedValue & vbCrLf & _
                      " Order by PeriodoInicial DESC" & vbCrLf

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    txtDataInicial.Text = Dr("PeriodoInicial")
                    txtDataFinal.Text = Dr("PeriodoFinal")
                    txtLivro.Text = Dr("Livro")
                    txtFolha.Text = Dr("PaginaInicial")
                Next
            Else
                txtDataInicial.Text = ""
                txtDataFinal.Text = ""
                txtFolha.Text = ""
                txtLivro.Text = ""
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIpi") = Empresa(0)
        DdlProcesso.Items.Clear()

        Sql = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & vbCrLf & _
              "  FROM ProcessoRAIPI " & vbCrLf & _
              " WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf & _
              " Order By PeriodoInicial DESC"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  " & vbCrLf & _
                        Format(Dr("PeriodoInicial"), "dd/MM/yyyy") & "  à  " & vbCrLf & _
                        Format(Dr("PeriodoFinal"), "dd/MM/yyyy") & "  Livro  " & vbCrLf & _
                        Format(Dr("Livro"), "000") & "  Folha  " & vbCrLf & _
                        Format(Dr("PaginaInicial"), "000") & vbCrLf
            DdlProcesso.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcesso.Items.Insert(0, "")
        DdlProcesso.SelectedIndex = 0
    End Sub

    Private Sub CriaOpcoes()
        Dim i As Integer = 0
        Dim dtOpcoes As DataTable

        dtOpcoes = New DataTable("Opcoes")
        dtOpcoes.Columns.Add("Codigo", Type.GetType("System.String")).DefaultValue = ""
        dtOpcoes.Columns.Add("Descricao", Type.GetType("System.String")).DefaultValue = ""

        While i < 3
            Row = dtOpcoes.NewRow()
            Select Case i
                Case 0
                    Row("Codigo") = "01"
                    Row("Descricao") = "Livro Registro de Apuração do IPI - RAIPI - Modelo P8 - Ajustes Outros Débitos/Créditos"
                Case 1
                    Row("Codigo") = "02"
                    Row("Descricao") = "Livro Registro de Apuração do IPI - RAIPI - Modelo P8 - Resumo da Apuração do Imposto"
                Case 2
                    Row("Codigo") = "03"
                    Row("Descricao") = "Termos de Abertura e Encerramento"
            End Select

            dtOpcoes.Rows.Add(Row)
            i += 1
        End While

        GridOpcoes.DataSource = dtOpcoes
        GridOpcoes.DataBind()
    End Sub

    Protected Sub GridOpcoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Validar() Then
                Select Case GridOpcoes.SelectedRow.Cells(1).Text()
                    Case "01"
                        ucRegistrosIpiAjustaResumo.BindGridView()
                        Popup.ConsultaDeRegistrosIpiAjustaResumo(Me.Page, "objRegistrosIpiAjustaResumo" & HID.Value)
                    Case ("02")
                        RegistroDeApuracaoDeIcmsModeloP9()
                    Case "03"
                        Opcao = "10"
                        Termos()
                End Select
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "Registro De Apuração De IPI Modelo P9"

    Private Sub RegistroDeApuracaoDeIcmsModeloP9()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text

        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
            Dim crpt As New ReportDocument()

            Try
                Sql = " SELECT Cliente_Id AS Cliente, Nome, Inscricao " & vbCrLf & _
                      "     FROM Clientes " & vbCrLf & _
                      " WHERE Cliente_Id = '" & Empresa(0) & "'" & vbCrLf & _
                      "   and endereco_id = " & Empresa(1)

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                Next

                Dim Periodo As String = "Mês ou Período/Ano : " & Format(PeriodoInicial, "dd/MM/yyyy") & " a " & Format(PeriodoFinal, "dd/MM/yyyy")

                'Temporaria Controla Notas de Serviço Conjugadas
                If CkConsNotasCompServicos.Checked Then
                    Sql2 = "  Select nfxi.Empresa_Id, " & vbCrLf &
                           "         nfxi.EndEmpresa_Id, " & vbCrLf &
                           "         nfxi.Cliente_Id, nfxi.EndCliente_Id, nfxi.EntradaSaida_Id, nfxi.Serie_Id, " & vbCrLf &
                           "         nfxi.Nota_Id, " & vbCrLf &
                           "         sum(Case when nfxi.CFOP_Id in (1933,2933)" & vbCrLf &
                           "               then 1 " & vbCrLf &
                           "               else 0 " & vbCrLf &
                           "               end ) ItensCfop," & vbCrLf &
                           "         sum(1) nrItens" & vbCrLf &
                           "  into #NotasCompostas" & vbCrLf &
                           "  From notasfiscaisxitens nfxi" & vbCrLf &
                           "  Inner join (SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, " & vbCrLf &
                           "               NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf &
                           "               NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id    " & vbCrLf &
                           "               FROM NotasFiscais " & vbCrLf &
                           "               INNER JOIN NotasFiscaisXItens " & vbCrLf &
                           "               ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf &
                           "               AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                           "               AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
                           "               AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                           "               AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
                           "               AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                           "               AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id   " & vbCrLf &
                           "               Inner Join Suboperacoes " & vbCrLf &
                           "               on Suboperacoes.Operacao_id = NotasFiscaisXItens.Operacao  " & vbCrLf &
                           "               and suboperacoes.SubOperacoes_id = NotasFiscaisXItens.suboperacao  " & vbCrLf &
                           "               WHERE NotasFiscais.EntradaSaida_Id = 'E'" & vbCrLf &
                           "               And NotasFiscais.Situacao = 1 " & vbCrLf &
                           "               And NotasFiscais.Empresa_Id = '" & Empresa(0) & "' " & vbCrLf &
                           "               AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf &
                           "               AND NotasFiscais.Serie_Id  NOT IN ('D', 'F', 'REC')" & vbCrLf &
                           "               AND(NotasFiscaisXItens.CFOP_Id  IN (1933,2933))" & vbCrLf &
                           "               )sb" & vbCrLf &
                           "               ON sb.Empresa_Id      = nfxi.Empresa_Id " & vbCrLf &
                           "               AND sb.EndEmpresa_Id   = nfxi.EndEmpresa_Id " & vbCrLf &
                           "               AND sb.Cliente_Id      = nfxi.Cliente_Id " & vbCrLf &
                           "               AND sb.EndCliente_Id   = nfxi.EndCliente_Id " & vbCrLf &
                           "               AND sb.EntradaSaida_Id = nfxi.EntradaSaida_Id " & vbCrLf &
                           "               AND sb.Serie_Id        = nfxi.Serie_Id " & vbCrLf &
                           "               AND sb.Nota_Id         = nfxi.Nota_Id   " & vbCrLf &
                           "               group by nfxi.Empresa_Id, nfxi.EndEmpresa_Id, " & vbCrLf &
                           "               nfxi.Cliente_Id, nfxi.EndCliente_Id, " & vbCrLf &
                           "               nfxi.EntradaSaida_Id, nfxi.Serie_Id, nfxi.Nota_Id" & vbCrLf &
                           "               having sum(1)  <> sum(Case when nfxi.CFOP_Id in (1933,2933)" & vbCrLf &
                           "               then 1" & vbCrLf &
                           "               else 0" & vbCrLf &
                           "               end )" & vbCrLf
                End If

                Sql = "  SELECT     Consulta.CFOP, " & vbCrLf &
                      "               Cfop.Descricao, " & vbCrLf &
                      "               ((Consulta.Contabil+Consulta.OutrasDespesas+Consulta.ValorContabilFrete)- Consulta.ValorDesconto) AS Contabil, " & vbCrLf &
                      "               Consulta.BaseICMS, " & vbCrLf &
                      "               Consulta.ValorICMS, " & vbCrLf &
                      "               0 AS ValorIPI" & vbCrLf &
                      "       FROM  (SELECT NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Contabil, " & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS," & vbCrLf
                Else
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS," & vbCrLf
                End If

                Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'FRETES' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESP.ADUANEIRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS OutrasDespesas, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorDesconto " & vbCrLf &
                      "       FROM   NotasFiscais INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
                      "               NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'IPI A RECUP.', 'PRODUTO','FRETES','DESP.ADUANEIRAS','DESCONTOS')" & vbCrLf
                Else
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'PRODUTO','FRETES','DESP.ADUANEIRAS','DESCONTOS')" & vbCrLf
                End If

                Sql &= "   WHERE  (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf &
                      "       And    (NotasFiscais.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf &
                      "       And    (NotasFiscais.EntradaSaida_Id = 'E') " & vbCrLf &
                      "       And    (NotasFiscais.Situacao = 1) " & vbCrLf &
                      "       And    (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      "       AND (NotasFiscaisXItens.CFOP_Id NOT IN (1933,2933))" & vbCrLf

                If CkConsNotasCompServicos.Checked Then
                    Sql &= "    OR exists (Select  1 from #notascompostas where " & vbCrLf & _
                           "         #NotasCompostas.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                           "         AND #NotasCompostas.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                           "         AND #NotasCompostas.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                           "         AND #NotasCompostas.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                           "         AND #NotasCompostas.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                           "         AND #NotasCompostas.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                           "         AND #NotasCompostas.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                           "              )                                  " & vbCrLf
                End If

                Sql &= "    Group BY" & vbCrLf & _
                       "           NotasFiscaisXItens.CFOP_Id) AS Consulta INNER JOIN" & vbCrLf & _
                       "           Cfop ON Consulta.CFOP = Cfop.Cfop_Id" & vbCrLf & _
                       "       GROUP BY " & vbCrLf & _
                       "           Consulta.CFOP, Consulta.Contabil, Consulta.OutrasDespesas, Consulta.ValorContabilFrete, Consulta.ValorDesconto, Consulta.BaseICMS, Consulta.ValorICMS, Cfop.Descricao ORDER BY Consulta.CFOP" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql2 + Sql, "ApuracaoIcmsEntradas"))

                Sql = "  SELECT	(CFOP + '.000') as CFOP, CASE " & vbCrLf &
                      "                WHEN CFOP = 1 THEN 'Do Estado'" & vbCrLf &
                      "                WHEN CFOP = 2 THEN 'De Outros Estados'" & vbCrLf &
                      "              ELSE 'Do Exterior' " & vbCrLf &
                      "              END AS Descricao, " & vbCrLf &
                      "		((sum(Contabil)+Sum(OutrasDespesas)+Sum(ValorContabilFrete))- Sum(ValorDesconto)) as Contabil, " & vbCrLf &
                      "		Sum(BaseICMS) as BaseIcms, " & vbCrLf &
                      "	    Sum(ValorICMS) as ValorIcms, " & vbCrLf &
                      "	    0 AS ValorIPI" & vbCrLf &
                      "   FROM  (SELECT LEFT(NotasFiscaisXItens.CFOP_Id, 1) AS CFOP, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Contabil, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorDesconto, " & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS, " & vbCrLf
                Else
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS, " & vbCrLf
                End If

                Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'FRETES' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorContabilFrete," & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'DESP.ADUANEIRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS OutrasDespesas " & vbCrLf &
                      "   FROM          NotasFiscais INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
                      "               NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'IPI A RECUP.', 'PRODUTO','DESCONTOS','FRETES','DESP.ADUANEIRAS')" & vbCrLf
                Else
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'PRODUTO','DESCONTOS','FRETES','DESP.ADUANEIRAS')" & vbCrLf
                End If

                Sql &= "   WHERE        (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf &
                      "       And          (NotasFiscais.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf &
                      "       And          (NotasFiscais.EntradaSaida_Id = 'E') " & vbCrLf &
                      "       And          (NotasFiscais.Situacao = 1) " & vbCrLf &
                      "       And          (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      "       AND (NotasFiscaisXItens.CFOP_Id NOT IN (1933,2933))" & vbCrLf

                If CkConsNotasCompServicos.Checked Then
                    Sql &= "  OR exists (Select  1 from #notascompostas where " & vbCrLf & _
                           "     #NotasCompostas.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                           "     AND #NotasCompostas.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                           "     AND #NotasCompostas.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                           "     AND #NotasCompostas.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                           "     AND #NotasCompostas.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                           "     AND #NotasCompostas.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                           "     AND #NotasCompostas.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                           "              )                                  " & vbCrLf
                End If
                Sql &= "    GROUP        BY NotasFiscaisXItens.CFOP_Id) AS Consulta" & vbCrLf & _
                    "       GROUP        BY CFOP"

                ds.Merge(Banco.ConsultaDataSet(Sql2 + Sql, "ApuracaoIcmsEntradasTotais"))

                Sql = "  SELECT     Consulta.CFOP, " & vbCrLf &
                      "               Cfop.Descricao, " & vbCrLf &
                      "               Consulta.Contabil, " & vbCrLf &
                      "               Consulta.BaseICMS, " & vbCrLf &
                      "               Consulta.ValorICMS, " & vbCrLf &
                      "               0 AS ValorIPI" & vbCrLf &
                      "       FROM  (SELECT NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Contabil, " & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                Else
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                End If

                Sql &= "       FROM   NotasFiscais INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
                      "               NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND  NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'IPI A RECUP.', 'PRODUTO')" & vbCrLf
                Else
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND  NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'PRODUTO')" & vbCrLf
                End If

                Sql &= "       WHERE  (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf &
                      "           And    (NotasFiscais.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf &
                      "           And    (NotasFiscais.EntradaSaida_Id = 'S') " & vbCrLf &
                      "           And    (NotasFiscais.Situacao = 1) " & vbCrLf &
                      "           And    (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      "       Group BY" & vbCrLf &
                      "               NotasFiscaisXItens.CFOP_Id) AS Consulta INNER JOIN" & vbCrLf &
                      "               Cfop ON Consulta.CFOP = Cfop.Cfop_Id" & vbCrLf &
                      "       GROUP BY " & vbCrLf &
                      "               Consulta.CFOP, Consulta.Contabil, Consulta.BaseICMS, Consulta.ValorICMS, Cfop.Descricao" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsSaidas"))

                Sql = "  SELECT	(CFOP + '.000') as CFOP, CASE " & vbCrLf &
                      "                WHEN CFOP = 5 THEN 'P/ O Estado'" & vbCrLf &
                      "                WHEN CFOP = 6 THEN 'P/ Outros Estados'" & vbCrLf &
                      "              ELSE 'P/ O Exterior' " & vbCrLf &
                      "              END AS Descricao, " & vbCrLf &
                      "		Sum(Contabil) as Contabil, " & vbCrLf &
                      "		Sum(BaseICMS) as BaseIcms, " & vbCrLf &
                      "	    Sum(ValorICMS) as ValorIcms, " & vbCrLf &
                      "	    0 AS ValorIPI" & vbCrLf &
                      " FROM  (SELECT LEFT(NotasFiscaisXItens.CFOP_Id, 1) AS CFOP, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Contabil, " & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN('IPI','IPI A RECUP.') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                Else
                    Sql &= "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "               SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                End If

                Sql &= " FROM          NotasFiscais INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
                      "               NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "               NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
                      "               NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'IPI A RECUP.', 'PRODUTO')" & vbCrLf
                Else
                    Sql &= "               NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'PRODUTO')" & vbCrLf
                End If

                Sql &= " WHERE        (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf &
                      " And          (NotasFiscais.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf &
                      " And          (NotasFiscais.EntradaSaida_Id = 'S') " & vbCrLf &
                      " And          (NotasFiscais.Situacao = 1 )" & vbCrLf &
                      " And          (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      " GROUP        BY NotasFiscaisXItens.CFOP_Id) AS Consulta" & vbCrLf &
                      " GROUP        BY CFOP" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsSaidasTotais"))

                ' Atualiza Debito do Imposto 
                Sql = "  SELECT	isnull(SUM(ValorICMS), 0) AS ValorIcms" & vbCrLf &
                      " FROM   (SELECT  LEFT(NotasFiscaisXItens.CFOP_Id, 1) AS CFOP, " & vbCrLf &
                      "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Contabil, " & vbCrLf
                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN ('IPI', 'IPI A RECUP.') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN ('IPI', 'IPI A RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                Else
                    Sql &= "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                End If

                Sql &= "       FROM      NotasFiscais INNER JOIN" & vbCrLf &
                      "                 NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                      "                 NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf &
                      "                 NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
                      "                 NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
                      "                 NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "                 NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "                 NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'IPI A RECUP.', 'PRODUTO')" & vbCrLf
                Else
                    Sql &= "                 NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'PRODUTO')" & vbCrLf
                End If

                Sql &= " WHERE          (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf &
                      " And            (NotasFiscais.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf &
                      " And            (NotasFiscais.EntradaSaida_Id = 'S') " & vbCrLf &
                      " And            (NotasFiscais.Situacao = 1 ) " & vbCrLf &
                      " And            (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      " GROUP BY       NotasFiscaisXItens.CFOP_Id) AS Consulta" & vbCrLf
                dss = Banco.ConsultaDataSet(Sql, "RegistrosFiscais")

                If dss.Tables(0).Rows.Count > 0 Then
                    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

                        Sql = "UPDATE ResumoRAIPI " & vbCrLf & _
                        " SET Valor = " & Microsoft.VisualBasic.Replace(Dr("ValorIcms"), ",", ".") & vbCrLf & _
                        " WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf & _
                        " AND Processo_Id = " & Processo & " AND Codigo_Id = 001" & vbCrLf
                        SqlArray.Add(Sql)

                        If Banco.GravaBanco(SqlArray) = False Then
                            'Return HttpContext.Current.Session("ssMessage")
                        End If
                    Next
                End If

                ' Atualiza Credito do Imposto 
                Sql = "  SELECT	isnull(SUM(ValorICMS), 0) AS ValorIcms" & vbCrLf &
                      " FROM   (SELECT  LEFT(NotasFiscaisXItens.CFOP_Id, 1) AS CFOP, " & vbCrLf &
                      "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Contabil, " & vbCrLf
                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN ('IPI', 'IPI A RECUP.') THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id IN ('IPI', 'IPI A RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                Else
                    Sql &= "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf &
                      "                 SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS" & vbCrLf
                End If

                Sql &= "       FROM      NotasFiscais INNER JOIN" & vbCrLf &
                      "                 NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                      "                 NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf &
                      "                 NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
                      "                 NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
                      "                 NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "                 NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf &
                      "                 NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf

                If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "49673784" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                    Sql &= "                 NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'IPI A RECUP.', 'PRODUTO')" & vbCrLf
                Else
                    Sql &= "                 NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_Id AND NotasFiscaisXEncargos.Encargo_Id IN ('ICMS', 'IPI', 'PRODUTO')" & vbCrLf
                End If

                Sql &= " WHERE          (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf &
                      " And            (NotasFiscais.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf &
                      " And            (NotasFiscais.EntradaSaida_Id = 'E') " & vbCrLf &
                      " And            (NotasFiscais.Situacao = 1) " & vbCrLf &
                      " And            (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf &
                      " AND (NotasFiscaisXItens.CFOP_Id NOT IN (1933,2933))" & vbCrLf

                If CkConsNotasCompServicos.Checked Then
                    Sql &= "  OR exists (Select  1 from #notascompostas where " & vbCrLf & _
                           "     #NotasCompostas.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                           "     AND #NotasCompostas.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                           "     AND #NotasCompostas.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                           "     AND #NotasCompostas.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                           "     AND #NotasCompostas.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                           "     AND #NotasCompostas.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                           "     AND #NotasCompostas.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                           "              )                                  " & vbCrLf
                End If


                Sql &= " GROUP BY       NotasFiscaisXItens.CFOP_Id) AS Consulta"
                dss = Banco.ConsultaDataSet(Sql2 + Sql, "RegistrosFiscais")

                If dss.Tables(0).Rows.Count > 0 Then
                    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql2 + Sql, "Clientes").Tables(0).Rows

                        Sql = "UPDATE ResumoRAIPI " & vbCrLf & _
                        "           SET Valor = " & Microsoft.VisualBasic.Replace(Dr("ValorIcms"), ",", ".") & vbCrLf & _
                        "       WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf & _
                        "           AND Processo_Id = " & Processo & " AND Codigo_Id = 005" & vbCrLf
                        SqlArray.Add(Sql)

                        If Banco.GravaBanco(SqlArray) = False Then
                            'Return HttpContext.Current.Session("ssMessage")
                        End If
                    Next

                End If

                ' Pega Saldo Credor do Periodo Anterior 
                Sql = "SELECT  COALESCE(SUM(Valor), 0) AS Valor" & vbCrLf & _
                "           FROM ResumoRAIPI " & vbCrLf & _
                "       WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo - 1 & " And  Codigo_Id = 14 " & vbCrLf
                dss = Banco.ConsultaDataSet(Sql, "ResumoRAIPI")

                If dss.Tables(0).Rows.Count > 0 Then
                    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

                        Sql = "UPDATE ResumoRAIPI " & vbCrLf & _
                        "           SET Valor = " & Microsoft.VisualBasic.Replace(Dr("Valor"), ",", ".") & vbCrLf & _
                        "       WHERE Empresa_Id = '" & Empresa(0) & "' " & vbCrLf & _
                        "           AND Processo_Id = " & Processo & " AND Codigo_Id = 009" & vbCrLf
                        SqlArray.Add(Sql)

                        If Banco.GravaBanco(SqlArray) = False Then
                            'Return HttpContext.Current.Session("ssMessage")
                        End If
                    Next
                End If

                TotalizaResumo()

                Sql = " SELECT 1 AS Ordem, RR.Codigo_Id as Codigo, DR.Descricao, RR.Valor" & vbCrLf & _
                      "           FROM ResumoRAIPI RR" & vbCrLf & _
                      "               INNER JOIN DescricaoRAIPI DR" & vbCrLf & _
                      "                   ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
                      "           WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & " and RR.Codigo_Id < 5" & vbCrLf & _
                      "               UNION" & vbCrLf & _
                      "           SELECT 2 AS Ordem, RR.Codigo_Id as Codigo, RIR.Descricao, RIR.Valor" & vbCrLf & _
                      "               FROM ResumoItensRAIPI RIR" & vbCrLf & _
                      "                   INNER JOIN ResumoRAIPI RR" & vbCrLf & _
                      "                       ON RIR.Empresa_Id = RR.Empresa_Id" & vbCrLf & _
                      "                       AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf & _
                      "                       AND RIR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
                      "           WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & " and RR.Codigo_Id < 5" & vbCrLf & _
                      "           ORDER BY 2, 1" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsDebitos"))

                Sql = "     SELECT 1 AS Ordem, RR.Codigo_Id as Codigo, DR.Descricao, RR.Valor" & vbCrLf & _
                      "           FROM ResumoRAIPI RR" & vbCrLf & _
                      "               INNER JOIN DescricaoRAIPI DR" & vbCrLf & _
                      "                   ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
                      "           WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & " and RR.Codigo_Id between 5 and 10" & vbCrLf & _
                      "           UNION" & vbCrLf & _
                      "           SELECT 2 AS Ordem, RR.Codigo_Id as Codigo, RIR.Descricao, RIR.Valor" & vbCrLf & _
                      "               FROM ResumoItensRAIPI RIR" & vbCrLf & _
                      "                   INNER JOIN ResumoRAIPI RR" & vbCrLf & _
                      "                       ON RIR.Empresa_Id = RR.Empresa_Id" & vbCrLf & _
                      "                       AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf & _
                      "                       AND RIR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
                      "           WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & " and RR.Codigo_Id between 5 and 10" & vbCrLf & _
                      "               ORDER BY 2, 1" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsCreditos"))

                Sql = "     SELECT 1 AS Ordem, RR.Codigo_Id as Codigo, DR.Descricao, RR.Valor" & vbCrLf & _
                      "               FROM ResumoRAIPI RR" & vbCrLf & _
                      "                   INNER JOIN DescricaoRAIPI DR" & vbCrLf & _
                      "                       ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
                      "               WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & " and RR.Codigo_Id > 10" & vbCrLf & _
                      "               UNION" & vbCrLf & _
                      "               SELECT 2 AS Ordem, RR.Codigo_Id as Codigo, RIR.Descricao, RIR.Valor" & vbCrLf & _
                      "                   FROM ResumoItensRAIPI RIR" & vbCrLf & _
                      "                       INNER JOIN ResumoRAIPI RR" & vbCrLf & _
                      "                           ON RIR.Empresa_Id = RR.Empresa_Id" & vbCrLf & _
                      "                           AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf & _
                      "                           AND RIR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
                      "               WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & " and RR.Codigo_Id > 10" & vbCrLf & _
                      "           ORDER BY 2, 1" & vbCrLf
                ds.Merge(Banco.ConsultaDataSet(Sql, "ApuracaoIcmsSaldos"))

                '---Definição do Relatorio Crystal Report---------
                crpt.FileName = Server.MapPath("~/Reports/Cr_Apuracao_IPI.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                crpt.SetDataSource(ds)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As ParameterFieldDefinitions
                Dim crparameterfielddefinition As ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaNome")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = EmpresaNome
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaCNPJ")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = EmpresaCNPJ
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("EmpresaInscricao")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = EmpresaInscricao
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Periodo")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Periodo
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Livro")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = "Livro.: " & Format(CInt(Livro), "000")
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Pagina")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Folha
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

                If IO.File.Exists(arquivo) Then
                    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
                End If
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crpt.Close()
                crpt.Dispose()
            End Try
        End If
    End Sub

#End Region

#Region "Termos"

    Private Sub TermosDeTeste()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text
        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then

        End If
    End Sub

    Private Sub Termos()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
        Processo = DdlProcesso.SelectedValue
        Livro = txtLivro.Text
        Folha = txtFolha.Text

        If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
            Dim Ds As New DataSet

            Dim DIAINI As String
            Dim MESINI As String
            Dim ANOINI As String

            Dim DIAFIM As String
            Dim MESFIM As String
            Dim ANOFIM As String

            Dim LIVRO01 As String

            MESINI = LTrim(RTrim(Str(Month(PeriodoInicial))))
            If Val(MESINI) < 10 Then
                MESINI = "0" + MESINI
            End If

            'MESINI = Mes(MESINI)
            MESINI = MonthName(MESINI)
            DIAINI = LTrim(RTrim(Str(Day(PeriodoInicial))))
            If Val(DIAINI) < 10 Then
                DIAINI = "0" + DIAINI
            End If
            ANOINI = Str(Year(PeriodoInicial))

            MESFIM = LTrim(RTrim(Str(Month(PeriodoFinal))))
            If Val(MESFIM) < 10 Then
                MESFIM = "0" + MESFIM
            End If
            'MESFIM = Mes(MESFIM)
            MESFIM = MonthName(MESFIM)
            DIAFIM = Str(Day(PeriodoFinal))
            If Val(DIAFIM) < 10 Then
                DIAFIM = "0" + DIAFIM
            End If
            ANOFIM = Str(Year(PeriodoFinal))
            LIVRO01 = Format(Val(Livro), "000")

            Dim mystr As String
            mystr = Format(Val(Folha), "000")
            Folha = mystr

            Sql = " SELECT  Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf & _
                  "               ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf & _
                  "               ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf & _
                  "               Estados.Descricao as NomeDoEstado, 1 as Ordem" & vbCrLf & _
                  "       FROM    Clientes INNER JOIN" & vbCrLf & _
                  "               ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                  "               Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                  "               Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf & _
                  "       WHERE Clientes.Cliente_Id Like '" & Empresa(0) & "'" & vbCrLf & _
                  "       Union" & vbCrLf & _
                  "       SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf & _
                  "               ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf & _
                  "               ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf & _
                  "               Estados.Descricao as NomeDoEstado, 2 as Ordem" & vbCrLf & _
                  "       FROM Clientes INNER JOIN" & vbCrLf & _
                  "               ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                  "               Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                  "               Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf & _
                  "       WHERE Clientes.Cliente_Id Like '" & Empresa(0) & "'" & vbCrLf
            Ds.Merge(Banco.ConsultaDataSet(Sql, "Termos"))

            'Definição do Relatorio Crystal Report---------
            Dim crpt As New ReportDocument()

            Try
                crpt.FileName = Server.MapPath("~/Reports/Cr_TermosRegistrosFiscaisICMS.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                crpt.SetDataSource(Ds)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("Folha")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Folha
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("DIAINI")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = DIAINI
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("MESINI")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = MESINI
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("ANOINI")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = ANOINI
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("DIAFIM")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = DIAFIM
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("MESFIM")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = MESFIM
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("ANOFIM")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = ANOFIM
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("LIVRO")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = LIVRO01
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Opcao")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Opcao
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crpt.Close()
                crpt.Dispose()
            End Try
        End If
    End Sub

#End Region

#Region "Funções"

    Private Function GravarProcesso(ByVal Processo As String, ByVal Livro As String, ByVal Folha As String, ByVal PeriodoIni As String, ByVal PeriodoFim As String, ByVal Cliente As String) As Boolean
        Dim strSQL As String

        strSQL = "SELECT * FROM ProcessoRAIPI WHERE Empresa_Id = '" & Cliente & "' And Processo_Id = " & Processo

        Dim dsProcesso As DataSet = Banco.ConsultaDataSet(strSQL, "ProcessoRAIPI")

        If dsProcesso.Tables(0).Rows.Count = 0 Then
            strSQL = "INSERT INTO ProcessoRAIPI " & vbCrLf & _
                "               (Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado) " & vbCrLf & _
                "        VALUES ('" & Cliente & "', " & Processo & ", '" & PeriodoIni.ToSqlDate() & "', " & vbCrLf & _
                "'" & PeriodoFim.ToSqlDate() & "', " & Livro & ", " & Folha & ", 'N')" & vbCrLf
        Else
            strSQL = "  UPDATE ProcessoRAIPI " & vbCrLf & _
                "           SET PeriodoInicial = '" & PeriodoIni.ToSqlDate() & "', " & vbCrLf & _
                "               PeriodoFinal = '" & PeriodoFim.ToSqlDate() & "', " & vbCrLf & _
                "               Livro = " & Livro & ", " & vbCrLf & _
                "               PaginaInicial = " & Folha & " " & vbCrLf & _
                "       WHERE Empresa_Id = '" & Cliente & "' " & vbCrLf & _
                "           AND Processo_Id = " & Processo & vbCrLf
        End If

        Dim alSQL As New ArrayList
        alSQL.Add(strSQL)

        Return Banco.GravaBanco(alSQL)
    End Function

    Private Sub GravaResumoIpi()
        Dim ds As New DataSet
        Empresa = DdlEmpresa.SelectedValue.Split("-")

        Sql = "Select * from ResumoRAIPI Where Empresa_Id = '" & Empresa(0) & "' and Processo_Id = " & Processo
        ds = Banco.ConsultaDataSet(Sql, "ResumoRAIPI")

        If ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            If ds.Tables(0).Rows.Count = 0 Then
                Sql = " INSERT INTO ResumoRAIPI " & vbCrLf & _
                                    "              (Empresa_Id, Processo_Id, Codigo_Id, Valor) " & vbCrLf & _
                                    "      SELECT '" & Empresa(0) & "', " & Processo & ", " & vbCrLf & _
                                    "              Codigo_Id, 0 " & vbCrLf & _
                                    "              FROM DescricaoRAIPI" & vbCrLf
                SqlArray.Add(Sql)

                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "Erro na gravacao de registro ResumoRAIPI.")
                End If
            End If
        End If

    End Sub

    'Private Function Mes(ByVal strMes As String) As String
    '    Select Case strMes
    '        Case "01"
    '            Return "janeiro"
    '        Case "02"
    '            Return "fevereiro"
    '        Case "03"
    '            Return "marco"
    '        Case "04"
    '            Return "abril"
    '        Case "05"
    '            Return "maio"
    '        Case "06"
    '            Return "junho"
    '        Case "07"
    '            Return "julho"
    '        Case "08"
    '            Return "agosto"
    '        Case "09"
    '            Return "setembro"
    '        Case "10"
    '            Return "outubro"
    '        Case "11"
    '            Return "novembro"
    '        Case "12"
    '            Return "dezembro"
    '    End Select

    '    Return ""
    'End Function

    Private Function Validar()
        Dim ok As Boolean = True

        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a unidade de negócio.")
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
        If txtLivro.Text = "" Then
            MsgBox(Me.Page, "Informe o número do livro.")
            ok = False
        End If
        If txtFolha.Text = "" Then
            MsgBox(Me.Page, "Informe o número da folha.")
            ok = False
        End If

        Return ok
    End Function

    Private Sub PegarNumeroDoProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        Processo = 1

        Sql = "SELECT COALESCE(max(Processo_Id), 0) AS Processo  " & vbCrLf & _
            "       FROM ProcessoRAIPI " & vbCrLf & _
            "   WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Processo = Dr("Processo") + 1
        Next
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
        txtLivro.Text = ""
        txtFolha.Text = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Private Sub TotalizaResumo()
        Empresa = DdlEmpresa.SelectedValue.Split("-")

        Dim Total01 As Decimal = 0
        Dim Total02 As Decimal = 0
        Dim Total03 As Decimal = 0
        Dim Total04 As Decimal = 0
        Dim Total05 As Decimal = 0
        Dim Total06 As Decimal = 0
        Dim Total07 As Decimal = 0
        Dim Total08 As Decimal = 0
        Dim Total09 As Decimal = 0
        Dim Total10 As Decimal = 0
        Dim Total11 As Decimal = 0
        Dim Total12 As Decimal = 0
        Dim Total13 As Decimal = 0
        Dim Total14 As Decimal = 0

        Sql = " SELECT 1 AS Ordem, RR.Codigo_Id, DR.Descricao, RR.Valor" & vbCrLf & _
              "       FROM ResumoRAIPI RR" & vbCrLf & _
              "           INNER JOIN DescricaoRAIPI DR" & vbCrLf & _
              "               ON DR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
              "       WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & vbCrLf & _
              "           UNION" & vbCrLf & _
              "   SELECT 2 AS Ordem, RR.Codigo_Id, RIR.Descricao, RIR.Valor" & vbCrLf & _
              "       FROM ResumoItensRAIPI RIR" & vbCrLf & _
              "           INNER JOIN ResumoRAIPI RR" & vbCrLf & _
              "               ON RIR.Empresa_Id = RR.Empresa_Id" & vbCrLf & _
              "               AND RIR.Processo_Id = RR.Processo_Id" & vbCrLf & _
              "               AND RIR.Codigo_Id = RR.Codigo_Id" & vbCrLf & _
              "       WHERE RR.Empresa_Id = '" & Empresa(0) & "' AND RR.Processo_Id = " & Processo & vbCrLf & _
              "           ORDER BY 2, 1" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 2 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 3 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 6 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 7 Then
                Dr("Valor") = 0
            End If
            If Dr("Ordem") = 1 And Dr("Codigo_Id") = 12 Then
                Dr("Valor") = 0
            End If

            Select Case Dr("Codigo_Id")
                Case 1
                    Total01 += Dr("Valor")
                Case 2
                    Total02 += Dr("Valor")
                Case 3
                    Total03 += Dr("Valor")
                Case 4
                    Total04 = Total01 + Total02 + Total03
                    Dr("Valor") = Total04
                Case 5
                    Total05 += Dr("Valor")
                Case 6
                    Total06 += Dr("Valor")
                Case 7
                    Total07 += Dr("Valor")
                Case 8
                    Total08 = Total05 + Total06 + Total07
                    Dr("Valor") = Total08
                Case 9
                    Total09 += Dr("Valor")
                Case 10
                    Total10 = Total08 + Total09
                    Dr("Valor") = Total10
                Case 11
                    If Total04 > Total10 Then
                        Total11 = Total04 - Total10
                        Dr("Valor") = Total11
                    End If
                Case 12
                    Total12 += Dr("Valor")
                Case 13
                    If Total11 > 0 Then
                        Total13 = Total11 - Total12
                        Dr("Valor") = Total13
                    End If
                Case 14
                    If Total10 > Total04 Then
                        Total14 = Total10 - Total04
                        Dr("Valor") = Total14
                    End If
            End Select
        Next

        Dim valor() As Decimal = {Total02, Total03, Total04, Total05, Total06, Total07, Total08, Total09, Total10, Total11, Total12, Total13, Total14}

        '---------Laço para atualizar os valores do Resumo onde Codigo_Id = 002 a 014----------------------------------------
        For i = 0 To 12
            Sql = " UPDATE ResumoRAIPI " & vbCrLf & _
                "       SET Valor = " & Microsoft.VisualBasic.Replace(valor(i), ",", ".") & vbCrLf & _
                "   WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & Processo & " AND Codigo_Id = " & i + 2 & vbCrLf
            SqlArray.Add(Sql)
        Next

        If Banco.GravaBanco(SqlArray) = False Then
            Mensagem = "Erro na atualização do resumo."
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("RegistrosDeIPI", "GRAVAR") Then
                If Validar() Then
                    Empresa = DdlEmpresa.SelectedValue.Split("-")
                    PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
                    PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
                    Livro = txtLivro.Text
                    Folha = txtFolha.Text

                    If DdlProcesso.Text = "" Then
                        PegarNumeroDoProcesso()
                    Else
                        Processo = DdlProcesso.SelectedValue
                    End If

                    If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
                        GravaResumoIpi()
                        CarregarProcesso()
                        DdlProcesso.SelectedValue = Processo
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
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
            Funcoes.Ajuda(Me.Page, "RegistrosDeIPI")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class