Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class OperacoesXEncargos
    Inherits BasePage

#Region "Declaracao Local"
    Dim Sql As String

    Dim DS As DataSet
    Dim Mensagem As String

    Private ObjListVersoes As [Lib].Negocio.ListOperacaoXEstado
    Private ObjOxE As [Lib].Negocio.OperacaoXEstado
    Private objNotaFiscal As NotaFiscal


#End Region

#Region "LOAD"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("OperacoesXEncargos", "ACESSAR") Then
                    ddl.Carregar(DdlGruposDeEstoques, CarregarDDL.Tabela.NivelGrupoProduto, "(LEN(Grupo_Id) >= 5) ", True)
                    ddl.Carregar(DdlOperacoes, CarregarDDL.Tabela.Operacao, "", True)
                    ddl.Carregar(DdlEstadoOrigem, CarregarDDL.Tabela.EstadoERegiao, "")
                    ddl.Carregar(DdlEstadoDestino, CarregarDDL.Tabela.EstadoERegiao, "")
                    ddl.Carregar(ddlSituacaoTributariaICMS, CarregarDDL.Tabela.SituacaoTributariaICMS, "", True)
                    ddl.Carregar(ddlSituacaoTributariaIPI, CarregarDDL.Tabela.SituacaoTributariaIPI, "", True)
                    ddl.Carregar(ddlSituacaoTributariaPISCOFINS, CarregarDDL.Tabela.SituacaoTributariaPISCOFINS, "")
                    ddl.Carregar(ddlSituacaoTributariaIBSCBS, CarregarDDL.Tabela.SituacaoTributariaIBSCBS, "")

                    CarregarEmpresasConsolidadas()

                    CFOPTitulos()
                    CarregarEncargosAdd()
                    Limpar()

                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaObservacoes.SetarHID(HID.Value)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Sessao"
    Private Sub SessaoSalvarVersoes()
        Session("ObjListVersoes" & HID.Value) = ObjListVersoes
    End Sub

    Private Sub SessaoRecuperarVersoes()
        If Session("ObjListVersoes" & HID.Value) Is Nothing Then
            ObjListVersoes = New [Lib].Negocio.ListOperacaoXEstado
        Else
            ObjListVersoes = Session("ObjListVersoes" & HID.Value)
        End If
    End Sub

    Private Sub SessaoSalvarObjeto()
        Session("ObjOxE" & HID.Value) = ObjOxE
    End Sub

    Private Sub SessaoRecuperarObjeto()
        If Session("ObjOxE" & HID.Value) Is Nothing Then
            ObjOxE = New [Lib].Negocio.OperacaoXEstado(True)
        Else
            ObjOxE = Session("ObjOxE" & HID.Value)
        End If
    End Sub

#End Region

#Region "BOTOES"

    Protected Sub BtnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim add As Button = CType(sender, Button)
            Dim row As GridViewRow = CType(add.NamingContainer, GridViewRow)

            Dim Encargo As String = row.Cells(1).Text

            SessaoRecuperarObjeto()
            If ObjOxE.Encargos.Where(Function(s) s.CodigoEncargo = Encargo).Count > 0 Then
                MsgBox(Me.Page, "Encargo já Adicionado")
                Exit Sub
            End If

            ObjOxE.Encargos.Add(New [Lib].Negocio.OperacaoXEstadoXEncargo(ObjOxE, Encargo))

            ' Utilizado para que o encargo LIQUIDO fique sempre na última posição da lista ao incluir novos encargos
            Dim objLIquido = ObjOxE.Encargos.Find(Function(f) f.CodigoEncargo = "LIQUIDO")
            ObjOxE.Encargos.Remove(objLIquido)
            ObjOxE.Encargos.Add(objLIquido)

            SessaoSalvarObjeto()
            '*********************************
            GridEncargos.DataSource = ObjOxE.Encargos
            GridEncargos.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("OperacoesXEncargos", "GRAVAR") Then
                If ValidarGravar() Then
                    GravaRegistro()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
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

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            Consultar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub lnkPlanoDeContas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPlanoDeContas.Click
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(False)
            Popup.ConsultaDePlanoDeContas(Me, "objOperacoesxEncargos" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim strSQL As String
        strSQL = "SELECT OE.Codigo_Id AS Codigo,OE.GrupoProduto as GrupoProduto_Id," & vbCrLf &
                 "       Ge.Descricao AS DescGrupo," & vbCrLf &
                 "       OE.Produto as Produto_Id," & vbCrLf &
                 "       isnull(Produtos.Nome, '') AS DescProduto," & vbCrLf &
                 "       CONVERT(VarChar, OE.Operacao) AS Operacao_Id," & vbCrLf &
                 "       CONVERT(VarChar,OE.SubOperacao) AS SubOperacao_Id," & vbCrLf &
                 "       SO.Descricao AS DescSubOp," & vbCrLf &
                 "       OE.EstadoOrigem as EstadoOrigem_Id," & vbCrLf &
                 "       UFOrigem.Descricao AS DescEstadoOrigem," & vbCrLf &
                 "       OE.EstadoDestino as EstadoDestino_Id," & vbCrLf &
                 "       UFDestino.Descricao AS DescEstadoDestino," & vbCrLf &
                 "       OEE.Encargo_Id," & vbCrLf &
                 "       Enc.Descricao AS DescEncargo," & vbCrLf &
                 "       CONVERT(VarChar, OE.STICMS) AS SituacaoTributaria," & vbCrLf &
                 "       case" & vbCrLf &
                 "          when Enc.ValorOuPeso = 0" & vbCrLf &
                 "             then 'V'" & vbCrLf &
                 "             else" & vbCrLf &
                 "                case" & vbCrLf &
                 "                   when Enc.ValorOuPeso = 1" & vbCrLf &
                 "                      then 'P'" & vbCrLf &
                 "                      else ''" & vbCrLf &
                 "                   end" & vbCrLf &
                 "          end AS ValorOuPeso," & vbCrLf &
                 "       CONVERT(VarChar, OE.CodigoFiscal) AS CodigoFiscal," & vbCrLf &
                 "       CONVERT(VarChar, Cfop.GrupoCfop_Id) AS GrupoCfop," & vbCrLf &
                 "       ISNULL(OEE.DebitaConta, '') DebitaConta," & vbCrLf &
                 "       PlanoDebito.Titulo AS DebitaTitulo," & vbCrLf &
                 "       ISNULL(OEE.CreditaConta,'') CreditaConta," & vbCrLf &
                 "       PlanoCredito.Titulo AS CreditaTitulo," & vbCrLf &
                 "	     '' as DescricaoFiscal," & vbCrLf &
                 "	     '' as Calculado," & vbCrLf &
                 "       OEE.Sinal," & vbCrLf &
                 "       OE.ObsICMS as Observacao," & vbCrLf &
                 "       CONVERT(VarChar, OEE.AliquotaBase) AS PercentualBase," & vbCrLf &
                 "       CONVERT(VarChar, OEE.Aliquota) AS Aliquota," & vbCrLf &
                 "       CONVERT(VarChar, OEE.AliquotaExibicao) AS AliquotaExibicao," & vbCrLf &
                 "       CONVERT(VarChar, OEE.AliquotaExibicao) AS PercentualLimite," & vbCrLf &
                 "       isnull(OE.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento" & vbCrLf &
                 "  Into #Consulta" & vbCrLf &
                 "  FROM OperacaoXEstado OE" & vbCrLf &
                 " Inner Join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                 "    on OE.Codigo_Id = OEE.Codigo_Id" & vbCrLf &
                 " INNER JOIN GruposDeEstoques Ge" & vbCrLf &
                 "    ON OE.GrupoProduto = Ge.Grupo_Id" & vbCrLf &
                 " INNER JOIN SubOperacoes SO" & vbCrLf &
                 "    ON OE.Operacao    = SO.Operacao_Id" & vbCrLf &
                 "   AND OE.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                 " INNER JOIN Estados AS UFOrigem" & vbCrLf &
                 "    ON OE.EstadoOrigem = UFOrigem.Estado_Id" & vbCrLf &
                 " INNER JOIN Estados AS UFDestino" & vbCrLf &
                 "    ON OE.EstadoDestino = UFDestino.Estado_Id" & vbCrLf &
                 " INNER JOIN Encargos Enc" & vbCrLf &
                 "    ON OEE.Encargo_Id = Enc.Encargo_Id" & vbCrLf &
                 " INNER JOIN Cfop" & vbCrLf &
                 "    ON OE.CodigoFiscal = Cfop.Cfop_Id" & vbCrLf &
                 "  LEFT JOIN PlanoDeContas AS PlanoDebito" & vbCrLf &
                 "    ON OEE.DebitaConta = PlanoDebito.Conta_Id" & vbCrLf &
                 "   AND PlanoDebito.Empresa_Id = '999999999999999999'" & vbCrLf &
                 "   AND PlanoDebito.EndEmpresa_Id = 0" & vbCrLf &
                 "  LEFT JOIN PlanoDeContas AS PlanoCredito" & vbCrLf &
                 "    ON OEE.CreditaConta = PlanoCredito.Conta_Id" & vbCrLf &
                 "   AND PlanoCredito.Empresa_Id    = '999999999999999999'" & vbCrLf &
                 "   AND PlanoCredito.EndEmpresa_Id = 0" & vbCrLf &
                 "  LEFT JOIN  Produtos" & vbCrLf &
                 "    ON OE.Produto = Produtos.Produto_Id" & vbCrLf &
                 " Where 1=1" & vbCrLf &
                 "   AND OE.Empresa = '" & ddlEmpresa.SelectedValue & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(DdlGruposDeEstoques.Text) Then
            strSQL &= "AND OE.GrupoProduto = '" & DdlGruposDeEstoques.SelectedValue & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlProdutos.Text) Then
            strSQL &= "AND OE.Produto = '" & DdlProdutos.SelectedValue & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlOperacoes.Text) Then
            strSQL &= "AND OE.Operacao = " & DdlOperacoes.SelectedValue & " " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlSubOperacoes.Text) Then
            strSQL &= "AND OE.SubOperacao = " & DdlSubOperacoes.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlEstadoOrigem.Text) Then
            strSQL &= "AND OE.EstadoOrigem = '" & DdlEstadoOrigem.Text & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlEstadoDestino.Text) Then
            strSQL &= "AND OE.EstadoDestino = '" & DdlEstadoDestino.SelectedValue & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlNaturezaDeRendimento.Text) Then
            strSQL &= "AND OE.CodigoNaturezaDeRendimento = '" & ddlNaturezaDeRendimento.SelectedValue & "' " & vbCrLf
        End If

        strSQL &= "ORDER BY OE.GrupoProduto, OE.Codigo_Id, OE.Produto, OE.Operacao, OE.SubOperacao, OE.EstadoOrigem, OE.EstadoDestino, CodigoNaturezaDeRendimento;" & vbCrLf &
                  " Select *, (GrupoProduto_Id + '- ' + DescGrupo + '- ' + Produto_Id + ' - ' + DescProduto + ' - ' + Operacao_Id + '-' + SubOperacao_Id + ' - ' + DescSubOp + ' - ' + EstadoOrigem_Id + '-' + EstadoDestino_Id) as Grupo" & vbCrLf &
                  "   From #Consulta"

        Dim Ds_OperacoesXEncargos As DataSet = Banco.ConsultaDataSet(strSQL, "OperacoesXEncargos")

        Return Ds_OperacoesXEncargos
    End Function

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("OperacoesXEncargos", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()

                Funcoes.BindReport(Me.Page, ds, "Cr_OperacoesXEncargos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(RTrim(ex.Message)) & ". Entre em contato com o Suporte de TI")
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "OperacoesXEncargos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "DDL"
    'GRUPO DE PRODUTO
    Protected Sub DdlGruposDeEstoques_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If DdlGruposDeEstoques.SelectedIndex > 0 Then
                ObjOxE.CodigoGrupoProduto = DdlGruposDeEstoques.SelectedValue
                ObjOxE.CodigoProduto = ""

                Produtos()
            Else
                ObjOxE.CodigoGrupoProduto = ""
                ObjOxE.CodigoProduto = ""

                DdlProdutos.Items.Clear()
            End If
            ObjOxE.Codigo = 0
            ObjOxE.InicioVigencia = Date.Now
            txtCodigoConfig.Text = String.Empty
            txtVigencia.Text = ObjOxE.InicioVigencia.ToString("dd/MM/yyyy")

            ObjOxE.Empresa = ddlEmpresa.SelectedValue

            SessaoSalvarObjeto()
            VerificaECarregaSeExistir()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'PRODUTO
    Protected Sub DdlProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If DdlProdutos.SelectedIndex > 0 Then
                ObjOxE.CodigoProduto = DdlProdutos.SelectedValue
            Else
                ObjOxE.CodigoProduto = ""
            End If
            SessaoSalvarObjeto()
            VerificaECarregaSeExistir()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'OPERACAO
    Protected Sub DdlOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlOperacoes.SelectedIndexChanged
        Try
            SessaoRecuperarObjeto()
            If DdlOperacoes.SelectedIndex > 0 Then
                ObjOxE.CodigoOperacao = DdlOperacoes.SelectedValue
            Else
                ObjOxE.CodigoOperacao = 0
            End If
            SubOperacoes(DdlOperacoes.SelectedValue)

            'CarregarGrupoFiscal()

            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'SUBOPERACAO
    Protected Sub DdlSubOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlSubOperacoes.SelectedIndexChanged
        Try
            SessaoRecuperarObjeto()
            If DdlSubOperacoes.SelectedIndex > 0 Then
                Dim SubOp As String() = DdlSubOperacoes.SelectedValue.Split("-")
                ObjOxE.CodigoSubOperacao = SubOp(1)

                If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If ObjOxE.CodigoSTIPI > 50 Then
                        ObjOxE.CodigoSTIPI = 3
                        ddlSituacaoTributariaIPI.SelectedValue = 3
                    End If
                Else
                    If ObjOxE.CodigoSTIPI < 50 Then
                        ObjOxE.CodigoSTIPI = 53
                        ddlSituacaoTributariaIPI.SelectedValue = 53
                    End If
                End If
            Else
                ObjOxE.CodigoSubOperacao = 0
            End If
            SessaoSalvarObjeto()
            If Not VerificaECarregaSeExistir() Then CarregarGrupoFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'ESTADO ORIGEM
    Protected Sub DdlEstadoOrigem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlEstadoOrigem.SelectedIndexChanged
        Try
            SessaoRecuperarObjeto()

            ObjOxE.EstadoOrigem = DdlEstadoOrigem.SelectedValue

            Dim UF As New Estado(ObjOxE.EstadoOrigem)
            ddl.Carregar(ddlObsICMS, CarregarDDL.Tabela.ObservacaoTributariaGeral, "Encargo like '%ICMS%' and estado in ('','" & UF.Codigo & "','" & UF.Regiao & "')")

            ddlBeneficioICMS.Items.Clear()

            'If UF.Codigo = "PR" Then
            ddl.Carregar(ddlBeneficioICMS, CarregarDDL.Tabela.AjustesDaApuracaoIcms, "left(codigo_id,2) = '" & UF.Codigo & "'", True)
            'End If

            If ObjOxE.EstadoOrigem = DdlEstadoOrigem.SelectedValue Then
                If ObjOxE.CodigoBeneficio.Length > 0 Then ddlBeneficioICMS.SelectedValue = ObjOxE.CodigoBeneficio
            Else
                ObjOxE.CodigoBeneficio = ""
                If ddlBeneficioICMS.Items.Count > 0 Then ddlBeneficioICMS.SelectedIndex = 0
            End If

            SessaoSalvarObjeto()

            If Not VerificaECarregaSeExistir() Then CarregarGrupoFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'ESTADO DESTINO
    Protected Sub DdlEstadoDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlEstadoDestino.SelectedIndexChanged
        Try
            SessaoRecuperarObjeto()
            ObjOxE.EstadoDestino = DdlEstadoDestino.SelectedValue
            ObjOxE.CodigoObsICMS = 0
            ddl.Carregar(ddlObsICMS, CarregarDDL.Tabela.ObservacaoTributariaGeral, "Encargo like '%ICMS%' and estado ='" & DdlEstadoDestino.SelectedValue & "'")
            SessaoSalvarObjeto()

            If Not VerificaECarregaSeExistir() Then CarregarGrupoFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'GRUPO FISCAL CFOP
    Protected Sub DdlCFOPTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If DdlCFOPTitulos.SelectedIndex > 0 Then
                ObjOxE.CodigoGrupoFiscal = DdlCFOPTitulos.SelectedValue
            Else
                ObjOxE.CodigoGrupoFiscal = 0
                ObjOxE.CodigoFiscal = 0
            End If
            CFOP()
            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'CFOP
    Protected Sub DdlCFOP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If DdlCFOP.SelectedIndex > 0 Then
                ObjOxE.CodigoFiscal = DdlCFOP.SelectedValue
            Else
                ObjOxE.CodigoFiscal = 0
            End If
            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'SIT. TRIB. ICMS
    Protected Sub ddlSituacaoTributariaICMS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSituacaoTributariaICMS.SelectedIndexChanged
        Try
            If ddlSituacaoTributariaICMS.SelectedIndex > 0 Then
                SessaoRecuperarObjeto()
                ObjOxE.CodigoSTICMS = ddlSituacaoTributariaICMS.SelectedValue
                SessaoSalvarObjeto()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'SIT. TRIB. IPI
    Protected Sub ddlSituacaoTributariaIPI_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlSituacaoTributariaIPI.SelectedIndex <> 0 Then
                SessaoRecuperarObjeto()
                If ObjOxE.CodigoSubOperacao = 0 Then
                    MsgBox(Me.Page, "Selecione a Operacao antes de configurar a Situacao do IPI")
                    ObjOxE.CodigoSTIPI = 0
                    ddlSituacaoTributariaIPI.SelectedIndex = 0
                    Exit Sub
                End If

                If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If ddlSituacaoTributariaIPI.SelectedIndex = 0 Then
                        ObjOxE.CodigoSTIPI = 3
                        ddlSituacaoTributariaIPI.SelectedValue = 3
                    ElseIf ddlSituacaoTributariaIPI.SelectedValue > 50 Then
                        ObjOxE.CodigoSTIPI = ddlSituacaoTributariaIPI.SelectedValue - 50
                        ddlSituacaoTributariaIPI.SelectedValue = ddlSituacaoTributariaIPI.SelectedValue - 50
                    Else
                        ObjOxE.CodigoSTIPI = ddlSituacaoTributariaIPI.SelectedValue
                    End If
                Else
                    If ddlSituacaoTributariaIPI.SelectedIndex = 0 Then
                        ObjOxE.CodigoSTIPI = 53
                        ddlSituacaoTributariaIPI.SelectedValue = 53
                    ElseIf ddlSituacaoTributariaIPI.SelectedValue < 50 Then
                        ObjOxE.CodigoSTIPI = ddlSituacaoTributariaIPI.SelectedValue + 50
                        ddlSituacaoTributariaIPI.SelectedValue = ddlSituacaoTributariaIPI.SelectedValue + 50
                    Else
                        ObjOxE.CodigoSTIPI = ddlSituacaoTributariaIPI.SelectedValue
                    End If
                End If

                SessaoSalvarObjeto()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'SIT. TRIB. PIS/COFINS
    Protected Sub ddlSituacaoTributariaPISCOFINS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If ddlSituacaoTributariaPISCOFINS.SelectedIndex > 0 Then
                ddl.Carregar(ddlObsPISCOFINS, CarregarDDL.Tabela.SituacaoTributariaPISCOFINSObs, "SituacaoTributariaPISCOFINS_Id = " & ddlSituacaoTributariaPISCOFINS.SelectedValue, True)
                ObjOxE.CodigoSTPISCOFINS = ddlSituacaoTributariaPISCOFINS.SelectedValue
                ObjOxE.CodigoObsPISCOFINS = 0
            Else
                ObjOxE.CodigoSTPISCOFINS = 0
                ObjOxE.CodigoObsPISCOFINS = 0
                ddlObsPISCOFINS.Items.Clear()
            End If
            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'SIT. TRIB. PIS/COFINS OBS
    Protected Sub ddlObsPISCOFINS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If ddlObsPISCOFINS.SelectedIndex > 0 Then
                ObjOxE.CodigoObsPISCOFINS = ddlObsPISCOFINS.SelectedValue
            Else
                ObjOxE.CodigoObsPISCOFINS = 0
            End If
            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'SIT. TRIB. IBS/CBS
    Protected Sub ddlSituacaoTributariaIBSCBS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If ddlSituacaoTributariaIBSCBS.SelectedIndex > 0 Then
                ObjOxE.CodigoSTIBSCBS = ddlSituacaoTributariaIBSCBS.SelectedValue
            Else
                ObjOxE.CodigoSTIBSCBS = -1
            End If

            CarregarClassificacaoTributariaIBSCBS()

            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlClassificacaoIBSCBS_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            SessaoRecuperarObjeto()
            If ddlClassificacaoIBSCBS.SelectedIndex > 0 Then
                ObjOxE.CodigoClassificacaoIBSCBS = ddlClassificacaoIBSCBS.SelectedValue
            Else
                ObjOxE.CodigoClassificacaoIBSCBS = -1
            End If

            SessaoSalvarObjeto()

            CarregarReDUCAOIBSCBS()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'NATUREZA DE REMDIMENTOS
    Protected Sub chkNaturezaDeRendimento_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            ddlNaturezaDeRendimento.Items.Clear()

            If chkNaturezaDeRendimento.Checked Then
                divNaturezaDeRendimento.Visible = True

                BuscarNaturezaDeRendimentos(ObjOxE)
            Else
                divNaturezaDeRendimento.Visible = False
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'PRODUTO
    Protected Sub ddlNaturezaDeRendimento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            If ddlNaturezaDeRendimento.SelectedIndex > 0 Then
                ObjOxE.CodigoNaturezaDeRendimento = ddlNaturezaDeRendimento.SelectedValue
            End If
            SessaoSalvarObjeto()
            'VerificaECarregaSeExistir()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkCalculadora_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperarObjeto()
            ObjOxE.UsarCalculadoraDeImposto = chkCalculadora.Checked
            SessaoSalvarObjeto()
            'VerificaECarregaSeExistir()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Carrega DDL"
    Private Sub Produtos()
        ddl.Carregar(DdlProdutos, CarregarDDL.Tabela.Produto, " isnull(Situacao,1) = 1 AND Grupo = '" & DdlGruposDeEstoques.SelectedValue & "'", True)
    End Sub

    Private Sub SubOperacoes(ByVal Operacao As String)
        ddl.Carregar(DdlSubOperacoes, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & Operacao & " ", True)
    End Sub

    Private Sub CFOPTitulos()
        Dim ListGrupoCFOP As New [Lib].Negocio.ListGrupoCFOP(True)

        Dim j As Integer = 0
        While j < ListGrupoCFOP.Count
            DdlCFOPTitulos.Items.Add(New ListItem(ListGrupoCFOP(j).Codigo & "-" & ListGrupoCFOP(j).Descricao, ListGrupoCFOP(j).Codigo))
            j += 1
        End While

        Funcoes.InserirLinhaEmBranco(DdlCFOPTitulos)
    End Sub

    Private Sub CFOP()
        If String.IsNullOrWhiteSpace(DdlCFOPTitulos.SelectedValue) Then
            ddl.Carregar(DdlCFOP, CarregarDDL.Tabela.CFOP, "GrupoCfop_Id =0", True)
        Else
            ddl.Carregar(DdlCFOP, CarregarDDL.Tabela.CFOP, "GrupoCfop_Id =" & DdlCFOPTitulos.SelectedValue, True)
        End If
    End Sub

    Private Sub BuscarNaturezaDeRendimentos(ByVal ObjOxE As OperacaoXEstado)
        Dim lstNaturezaDeRendimentos As New [Lib].Negocio.NaturezaDeRendimentos("Where TipoDePessoa in (3,1)")

        ddlNaturezaDeRendimento.Items.Add(New ListItem("", 0))

        For Each item In lstNaturezaDeRendimentos
            ddlNaturezaDeRendimento.Items.Add(New ListItem(item.Codigo & " - " & item.Descricao, item.Codigo))
        Next
    End Sub
#End Region

#Region "Grid"
    Protected Sub GridEncargos_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridEncargos.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                SessaoRecuperarObjeto()

                Dim lblEncargo As String = CType(e.Row.FindControl("lblEncargo"), Label).Text
                Dim ddlSinal As DropDownList = CType(e.Row.FindControl("DdlSinal"), DropDownList)

                Dim encSel = ObjOxE.Encargos.Where(Function(s) s.CodigoEncargo.Equals(lblEncargo)).FirstOrDefault

                CType(e.Row.FindControl("txtDebito"), TextBox).Text = encSel.CodigoDebitaConta
                If encSel.CodigoDebitaConta.Length > 0 Then CType(e.Row.FindControl("txtDebito"), TextBox).ToolTip = encSel.DebitaConta.Titulo

                CType(e.Row.FindControl("txtCredito"), TextBox).Text = encSel.CodigoCreditaConta
                If encSel.CodigoCreditaConta.Length > 0 Then CType(e.Row.FindControl("txtCredito"), TextBox).ToolTip = encSel.CreditaConta.Titulo

                CType(e.Row.FindControl("txtBase"), TextBox).Text = encSel.AliquotaBase.ToString("N9")
                CType(e.Row.FindControl("TxtAliquota"), TextBox).Text = encSel.Aliquota.ToString("N9")
                CType(e.Row.FindControl("txtAliqExib"), TextBox).Text = encSel.AliquotaExibicao.ToString("N2")
                CType(e.Row.FindControl("TxtLimite"), TextBox).Text = encSel.AliquotaLimite.ToString("N2")

                'Pode-se incluir observação para os encargos, menos: LIQUIDO, ICMS, PIS, COFINS
                Dim ddlObservacao As DropDownList = CType(e.Row.FindControl("ddlObservacao"), DropDownList)
                If Not (lblEncargo.Equals("LIQUIDO") OrElse lblEncargo.Equals("ICMS") OrElse lblEncargo.Equals("PIS")) Then
                    ddlObservacao.Visible = True
                    ddl.Carregar(ddlObservacao, CarregarDDL.Tabela.ObservacaoTributariaGeral, "Encargo like '%" & CType(e.Row.FindControl("lblEncargo"), Label).Text & "%'")
                    If encSel.ObservacaoTributaria <> 0 Then ddlObservacao.SelectedValue = encSel.ObservacaoTributaria
                    If ddlObservacao.Items.Count <= 0 Then ddlObservacao.Visible = False
                Else
                    ddlObservacao.Visible = False
                End If

                'ddlObservacao.Visible = Not (lblEncargo.Equals("LIQUIDO") OrElse lblEncargo.Equals("ICMS") OrElse lblEncargo.Equals("PIS") OrElse lblEncargo.Equals("COFINS"))

                If lblEncargo.Equals("PRODUTO") Then
                    ddlSinal.SelectedValue = "+"
                    ddlSinal.Enabled = False
                    CType(e.Row.FindControl("imgExcluir"), Image).Visible = False

                ElseIf lblEncargo.Equals("AFIXAR") Then
                    ddlSinal.SelectedValue = "="
                    ddlSinal.Enabled = False

                ElseIf lblEncargo.Equals("LIQUIDO") Then
                    CType(e.Row.FindControl("imbContaDebito"), ImageButton).Visible = False
                    CType(e.Row.FindControl("imbContaCredito"), ImageButton).Visible = False
                    ddlSinal.SelectedValue = "="
                    ddlSinal.Enabled = False
                    CType(e.Row.FindControl("imgExcluir"), Image).Visible = False
                    CType(e.Row.FindControl("txtDebito"), TextBox).ReadOnly = True
                    CType(e.Row.FindControl("txtCredito"), TextBox).ReadOnly = True
                    CType(e.Row.FindControl("txtDebito"), TextBox).BackColor = Drawing.Color.LightYellow
                    CType(e.Row.FindControl("txtCredito"), TextBox).BackColor = Drawing.Color.LightYellow

                    If ObjOxE.SubOperacao IsNot Nothing Then
                        If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                            If Left(Session("ssEmpresa"), 8) = "24450490" Then
                                CType(e.Row.FindControl("imbContaDebito"), ImageButton).Visible = True
                                CType(e.Row.FindControl("txtDebito"), TextBox).ReadOnly = False
                            Else
                                CType(e.Row.FindControl("txtDebito"), TextBox).Text = ""
                            End If

                            CType(e.Row.FindControl("txtCredito"), TextBox).Text = ObjOxE.SubOperacao.CodigoGrupoContas
                            If ObjOxE.SubOperacao.CodigoGrupoContas.Length > 0 Then CType(e.Row.FindControl("txtCredito"), TextBox).ToolTip = ObjOxE.SubOperacao.GrupoDeConta.Titulo
                        Else
                            CType(e.Row.FindControl("txtDebito"), TextBox).Text = ObjOxE.SubOperacao.CodigoGrupoContas
                            If ObjOxE.SubOperacao.CodigoGrupoContas.Length > 0 Then CType(e.Row.FindControl("txtDebito"), TextBox).ToolTip = ObjOxE.SubOperacao.GrupoDeConta.Titulo

                            If Left(Session("ssEmpresa"), 8) = "24450490" Then
                                CType(e.Row.FindControl("imbContaCredito"), ImageButton).Visible = True
                                CType(e.Row.FindControl("txtCredito"), TextBox).ReadOnly = False
                            Else
                                CType(e.Row.FindControl("txtCredito"), TextBox).Text = ""
                            End If
                        End If
                    End If
                Else
                    CType(e.Row.FindControl("imgExcluir"), Image).Visible = True
                    ddlSinal.SelectedValue = encSel.Sinal
                End If

                If ddlSinal.SelectedIndex > 0 And Not lblEncargo.Equals("LIQUIDO") Then
                    Dim txtDebito As TextBox = CType(e.Row.FindControl("txtDebito"), TextBox)
                    Dim txtCredito As TextBox = CType(e.Row.FindControl("txtCredito"), TextBox)
                    Dim imgCredito As ImageButton = CType(e.Row.FindControl("imbContaCredito"), ImageButton)
                    Dim imgDebito As ImageButton = CType(e.Row.FindControl("imbContaDebito"), ImageButton)

                    txtCredito.Visible = True
                    imgCredito.Visible = True
                    txtDebito.Visible = True
                    imgDebito.Visible = True

                    If ObjOxE.SubOperacao IsNot Nothing Then
                        If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                            If ddlSinal.SelectedValue.Equals("+") Then
                                txtCredito.Text = String.Empty
                                txtCredito.Visible = False
                                imgCredito.Visible = False
                            ElseIf ddlSinal.SelectedValue.Equals("-") Then
                                txtDebito.Text = String.Empty
                                txtDebito.Visible = False
                                imgDebito.Visible = False
                            End If
                        Else
                            If ddlSinal.SelectedValue.Equals("+") Then
                                txtDebito.Text = String.Empty
                                txtDebito.Visible = False
                                imgDebito.Visible = False
                            ElseIf ddlSinal.SelectedValue.Equals("-") Then
                                txtCredito.Text = String.Empty
                                txtCredito.Visible = False
                                imgCredito.Visible = False
                            End If
                        End If
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridConsulta.SelectedIndexChanged
        Try
            Dim Parametros As New OperacaoXEstado
            Parametros.Codigo = GridConsulta.SelectedRow.Cells(1).Text()

            ObjListVersoes = New [Lib].Negocio.ListOperacaoXEstado(Parametros)
            ObjOxE = New [Lib].Negocio.OperacaoXEstado(Parametros)

            SessaoSalvarVersoes()
            SessaoSalvarObjeto()

            If Funcoes.VerificaPermissao("OperacoesXEncargos", "LIBERAR") Then
                chkAtivo.Enabled = True
            Else
                chkAtivo.Enabled = False
            End If

            txtCodigoConfig.ReadOnly = True
            CarregarFormComAClasse()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

    Private Function ValidarGravar() As Boolean
        SessaoRecuperarObjeto()
        If String.IsNullOrWhiteSpace(DdlSubOperacoes.SelectedValue) Then
            MsgBox(Me.Page, "Campo SubOperações é obrigatório", eTitulo.Info)
            Return False
        End If

        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Campo Empresa é obrigatório", eTitulo.Info)
            Return False
        End If

        If ((ObjOxE.SubOperacao.Classe = eClassesOperacoes.AFIXAR And ObjOxE.SubOperacao.Devolucao And Not ObjOxE.SubOperacao.QuantidadeFiscal) Or
            ObjOxE.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES) And
            Not ObjOxE.Encargos.Any(Function(s) s.CodigoEncargo = "AFIXAR") Then
            MsgBox(Me.Page, "Encargo AFIXAR não foi selecionado", eTitulo.Info)
            Return False
        End If

        For Each row In GridEncargos.Rows
            Dim encargo As Label = CType(row.FindControl("lblEncargo"), Label)

            If encargo.Text.Equals("IBS") AndAlso ddlSituacaoTributariaIBSCBS.SelectedIndex = 0 Then
                MsgBox(Me.Page, "É necessário informar a situação tributaria do encargo IBS", eTitulo.Info)
                Return False
            ElseIf encargo.Text.Equals("CBS") AndAlso ddlSituacaoTributariaIBSCBS.SelectedIndex = 0 Then
                MsgBox(Me.Page, "É necessário informar a situação tributaria do encargo CBS", eTitulo.Info)
                Return False
            End If

            If encargo.Text.Equals("IBS") AndAlso ddlClassificacaoIBSCBS.SelectedIndex = 0 Then
                MsgBox(Me.Page, "É necessário informar a classificação do encargo IBS", eTitulo.Info)
                Return False
            ElseIf encargo.Text.Equals("CBS") AndAlso ddlClassificacaoIBSCBS.SelectedIndex = 0 Then
                MsgBox(Me.Page, "É necessário informar a classificação do encargo CBS", eTitulo.Info)
                Return False
            End If

        Next

        'APEBAS BAXI FOODS - FURLAN - 21/05/25
        If Left(Session("ssEmpresa"), 8) = "40938762" Then
            Dim objGrupoProduto As GrupoProduto = New GrupoProduto(DdlGruposDeEstoques.SelectedValue)
            If objGrupoProduto.RelatorioEstoque Then
                For Each row In GridEncargos.Rows
                    Dim debito As TextBox = CType(row.FindControl("txtDebito"), TextBox)
                    Dim credito As TextBox = CType(row.FindControl("txtCredito"), TextBox)
                    Dim encargo As Label = CType(row.FindControl("lblEncargo"), Label)

                    'If encargo.Text.Equals("PRODUTO") Then

                    '    If Not String.IsNullOrWhiteSpace(debito.Text) Then
                    '        If debito.Text.Substring(0, 1).Equals("4") Then
                    '            MsgBox(Me.Page, "O grupo do produto quando controla estoque deve estar vinculado a uma conta de estoque. No encargo PRODUTO.")
                    '            Return False
                    '        End If
                    '    End If
                    '    If Not String.IsNullOrWhiteSpace(credito.Text) Then
                    '        If credito.Text.Substring(0, 1).Equals("4") Then
                    '            MsgBox(Me.Page, "O grupo do produto quando controla estoque deve estar vinculado a uma conta de estoque. No encargo PRODUTO.")
                    '            Return False
                    '        End If
                    '    End If
                    'End If
                Next

            End If
        End If

        Return True
    End Function

    Private Sub CarregarEncargosAdd()
        Sql = "SELECT Encargo_Id As Encargo" & vbCrLf &
              "  FROM Encargos" & vbCrLf &
              " Where OperacaoXEncargo = 'S'" & vbCrLf &
              "   And Encargo_Id not in ('PRODUTO','LIQUIDO')" & vbCrLf &
              " Order By Encargo_Id"

        DS = Banco.ConsultaDataSet(Sql, "Encargos")
        gridEncargosAdd.DataSource = DS
        gridEncargosAdd.DataBind()
    End Sub

    Private Sub CarregarClassificacaoTributariaIBSCBS()

        If ddlSituacaoTributariaIBSCBS.SelectedIndex = 0 Then
            Exit Sub
        End If

        Dim strSQL As String = "SELECT ClassificacaoTributaria_Id AS Codigo, Descricao " & vbCrLf &
                               "  FROM ClassificacaoTributariaIBSCBS " & vbCrLf

        strSQL &= "WHERE CST = " & ddlSituacaoTributariaIBSCBS.SelectedValue & " " & vbCrLf

        strSQL &= "ORDER BY ClassificacaoTributaria_Id"

        ddlClassificacaoIBSCBS.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "GeneroDoProdutoXSub").Tables(0).Rows
            ddlClassificacaoIBSCBS.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next

        If ddlClassificacaoIBSCBS IsNot Nothing Then
            For Each li As ListItem In ddlClassificacaoIBSCBS.Items
                li.Attributes("title") = li.Text
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddlClassificacaoIBSCBS)
    End Sub

    Private Sub CarregarReducaoIBSCBS()

        If ddlSituacaoTributariaIBSCBS.SelectedIndex = 0 Then
            Exit Sub
        End If

        If ddlClassificacaoIBSCBS.SelectedIndex = 0 Then
            Exit Sub
        End If

        Try

            SessaoRecuperarObjeto()

            Dim strSQL As String = "  SELECT ReducaoIBS_Perc, ReducaoCBS_Perc " & vbCrLf &
                                   "  FROM ClassificacaoTributariaIBSCBS " & vbCrLf

            strSQL &= "WHERE CST = " & ddlSituacaoTributariaIBSCBS.SelectedValue & " AND ClassificacaoTributaria_Id = " & ddlClassificacaoIBSCBS.SelectedValue & "" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "GeneroDoProdutoXSub").Tables(0).Rows

                ObjOxE.ReducaoIBS_Perc = Dr("ReducaoIBS_Perc")
                ObjOxE.ReducaoCBS_Perc = Dr("ReducaoCBS_Perc")

            Next

            SessaoSalvarObjeto()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub CarregarEmpresasConsolidadas()
        Dim sql As String
        Dim consolidaEmpresa As Boolean = False

        sql = "Select top 1 * from OperacaoXEstado where Empresa = '99999999'"

        DS = Banco.ConsultaDataSet(sql, "VerEmpresa")

        For Each Dr As DataRow In DS.Tables(0).Rows
            ddlEmpresa.Items.Add(New ListItem("99999999 - CONSOLIDADO", "99999999"))
            consolidaEmpresa = True
        Next

        If consolidaEmpresa Then
            ddlEmpresa.SelectedValue = "99999999"
        Else
            sql = "select distinct substring(ce.Empresa_Id,1,8) as Empresa_Id," & vbCrLf &
                  "       (Select top(1) Nome " & vbCrLf &
                  "          from Clientes" & vbCrLf &
                  "         Where substring(Cliente_Id,1,8) = substring(ce.Empresa_Id,1,8)" & vbCrLf &
                  "        ) as nome" & vbCrLf &
                  "  From clientesxempresas ce" & vbCrLf

            DS = Banco.ConsultaDataSet(sql, "Empresas")

            ddlEmpresa.Items.Clear()
            ddlEmpresa.DataTextField = "Descricao"
            ddlEmpresa.DataValueField = "Codigo"

            For Each Dr As DataRow In DS.Tables(0).Rows
                ddlEmpresa.Items.Add(New ListItem(Dr("Empresa_Id") & " - " & Dr("Nome"), Dr("Empresa_Id")))
            Next

            ddlEmpresa.SelectedValue = Left(Session("ssEmpresa"), 8)
        End If

    End Sub

    Function OperacaoXNotas() As Boolean
        Dim i As Integer = 0
        Dim strOR As String = ""
        Dim strEncargos As String = ""
        Dim Produto As String = ""
        Mensagem = ""
        Dim strOpe() As String = DdlSubOperacoes.SelectedValue.ToString.Split("-")

        If DdlProdutos.SelectedIndex > 0 Then
            Produto = DdlProdutos.SelectedValue
        End If

        While i < GridEncargos.Rows.Count
            If CType(GridEncargos.Rows(i).Cells(0).FindControl("CbkSelecionado"), CheckBox).Checked = True Then
                strEncargos &= strOR & "'" & GridEncargos.Rows(i).Cells(1).Text() & "'"
                If strOR.Length = 0 Then
                    strOR = ","
                End If
            End If
            i += 1
        End While

        Sql = " select  " & vbCrLf &
              "     case " & vbCrLf &
              "        when  exists(Select 'TEM' " & vbCrLf &
              "                       from NotasFiscaisXEncargos " & vbCrLf &
              "                      Where Grupo         = '" & DdlGruposDeEstoques.SelectedValue & "' " & vbCrLf &
              "                        and Produto       = '" & Produto & "' " & vbCrLf &
              "                        And Operacao      = " & strOpe(0) & " " & vbCrLf &
              "                        And SubOperacao   = " & strOpe(1) & " " & vbCrLf &
              "                        and EstadoOrigem  = '" & DdlEstadoOrigem.SelectedValue & "' " & vbCrLf &
              "                        And EstadoDestino = '" & DdlEstadoDestino.SelectedValue & "' " & vbCrLf &
              "                        And Encargo_Id    IN(" & strEncargos & ")" & vbCrLf &
              "                           ) " & vbCrLf &
              "           then 'TRUE' " & vbCrLf &
              "           else 'FALSE' " & vbCrLf &
              "     End Existe"
        Dim ds As New DataSet

        ds = Banco.ConsultaDataSet(Sql, "Encargos")

        If ds.Tables(0).Rows(0).Item("Existe") = "TRUE" Then
            Mensagem = "Encargo com Nota Fiscal emitida não pode ser Alterado/Excluído."
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub Limpar()
        txtCodigoConfig.ReadOnly = False
        lnkRecontabilizar.Parent.Visible = False

        ddlVersao.Items.Clear()

        ddlEmpresa.Enabled = False

        ddlBeneficioICMS.Items.Clear()
        ddlNaturezaDeRendimento.SelectedValue = 0

        ddlSituacaoTributariaIBSCBS.SelectedIndex = 0
        ddlClassificacaoIBSCBS.Items.Clear()
        chkCalculadora.Checked = False

        If Funcoes.VerificaPermissao("OperacoesXEncargos", "LIBERAR") Then
            chkAtivo.Enabled = True
        Else
            chkAtivo.Enabled = False
        End If

        chkAtivo.Checked = True

        chkNaturezaDeRendimento.Checked = False
        divNaturezaDeRendimento.Visible = False

        ObjListVersoes = Nothing
        SessaoSalvarVersoes()

        ObjOxE = New [Lib].Negocio.OperacaoXEstado(True)
        ObjOxE.Empresa = ddlEmpresa.SelectedValue

        SessaoSalvarObjeto()
        CarregarFormComAClasse(True)
        txtVigencia.Text = Now().ToString("dd/MM/yyyy")
    End Sub

    Public Sub CarregarFormComAClasse(Optional ByVal Limpar As Boolean = False)
        SessaoRecuperarObjeto()
        SessaoRecuperarVersoes()
        If ObjListVersoes IsNot Nothing Then
            ddlVersao.DataTextField = "DESCRICAO"
            ddlVersao.DataValueField = "CODIGO"
            If ObjListVersoes Is Nothing OrElse ObjListVersoes.Count = 0 Then
                ddlVersao.DataSource = Nothing
            Else
                ddlVersao.DataSource = (From x In ObjListVersoes
                                        Order By x.InicioVigencia, x.Codigo
                                        Select x.Codigo, Descricao = x.InicioVigencia.ToString("dd-MM-yyyy") & "   ID-" & x.Codigo & "   " & x.UsuarioInclusao & "   " & x.DataHoraInclusao.ToString("dd-MM-yyyy") & IIf(x.Ativo, "   Ativo", "   Cancelada"))
            End If
            ddlVersao.DataBind()
            If ObjOxE.Codigo > 0 Then ddlVersao.SelectedValue = ObjOxE.Codigo
        End If

        txtCodigoConfig.Text = ObjOxE.Codigo
        txtVigencia.Text = ObjOxE.InicioVigencia
        chkAtivo.Checked = CBool(ObjOxE.Ativo)

        ddlEmpresa.SelectedValue = ObjOxE.Empresa

        If ObjOxE.EstadoOrigem.Length = 0 Then
            DdlEstadoOrigem.SelectedIndex = 0
        Else
            DdlEstadoOrigem.SelectedValue = ObjOxE.EstadoOrigem
            Dim uf As New Estado(ObjOxE.EstadoOrigem)
            ddl.Carregar(ddlObsICMS, CarregarDDL.Tabela.ObservacaoTributariaGeral, "Encargo like '%ICMS%' and estado in ('','" & uf.Codigo & "','" & uf.Regiao & "')")

            'If uf.Codigo = "PR" Then
            ddlBeneficioICMS.Items.Clear()

            ddl.Carregar(ddlBeneficioICMS, CarregarDDL.Tabela.AjustesDaApuracaoIcms, "left(codigo_id,2) = '" & uf.Codigo & "'", True)
            'End If
        End If

        If ObjOxE.EstadoDestino.Length = 0 Then
            DdlEstadoDestino.SelectedIndex = 0
        Else
            DdlEstadoDestino.SelectedValue = ObjOxE.EstadoDestino
        End If

        If ObjOxE.CodigoGrupoProduto.Length = 0 Then
            DdlGruposDeEstoques.SelectedIndex = 0
            DdlProdutos.Items.Clear()
        Else
            DdlGruposDeEstoques.SelectedValue = ObjOxE.CodigoGrupoProduto
            Produtos()
            If ObjOxE.CodigoProduto.Length = 0 Then
                DdlProdutos.SelectedIndex = 0
            Else
                DdlProdutos.SelectedValue = ObjOxE.CodigoProduto
            End If
        End If

        If ObjOxE.CodigoOperacao = 0 Then
            DdlOperacoes.SelectedIndex = 0
            DdlSubOperacoes.Items.Clear()
        Else
            DdlOperacoes.SelectedValue = ObjOxE.CodigoOperacao
            SubOperacoes(ObjOxE.CodigoOperacao)
            If ObjOxE.CodigoSubOperacao = 0 Then
                DdlSubOperacoes.SelectedIndex = 0
            Else
                DdlSubOperacoes.SelectedValue = ObjOxE.CodigoOperacao & "-" & ObjOxE.CodigoSubOperacao
            End If
        End If

        CarregarGrupoFiscal()

        If ObjOxE.CodigoGrupoFiscal = 0 Then
            If DdlCFOPTitulos.Items.Count > 0 Then DdlCFOPTitulos.SelectedIndex = 0
            DdlCFOP.Items.Clear()
        Else
            DdlCFOPTitulos.SelectedValue = ObjOxE.CodigoGrupoFiscal
            CFOP()
            If ObjOxE.CodigoFiscal = 0 Then
                DdlCFOP.SelectedIndex = 0
            Else
                DdlCFOP.SelectedValue = ObjOxE.CodigoFiscal
            End If
        End If

        If ObjOxE.CodigoBeneficio.Length = 0 Then
            'If Not ObjOxE.EstadoOrigem = "PR" Then ddlBeneficioICMS.Items.Clear()

            If ddlBeneficioICMS.Items.Count > 0 Then ddlBeneficioICMS.SelectedIndex = 0
        Else
            ddlBeneficioICMS.SelectedValue = ObjOxE.CodigoBeneficio
        End If

        If ObjOxE.Encargos.Count > 0 Then
            'O Limpar é usado para que seja setado como "Vazia" a lista de Situações Tributária de IPI e ICMS ao clicar em limpar 
            'e quando for consultada a Conf. Operações x encargos não fique em "Vazia" caso seja consultada uma Conf. com SituacaoTributariaIPI = 0 ou SituacaoTributariaICMS  = 0
            If Limpar Then
                If ObjOxE.CodigoSTICMS <= 0 Then
                    ddlSituacaoTributariaICMS.SelectedIndex = 0
                Else
                    ddlSituacaoTributariaICMS.SelectedValue = ObjOxE.CodigoSTICMS
                End If

                If ObjOxE.CodigoObsICMS <= 0 Then
                    If ddlObsICMS.Items.Count > 0 Then ddlObsICMS.SelectedIndex = 0
                Else
                    With ddlObsICMS
                        .SelectedIndex = .Items.IndexOf(.Items.FindByValue(ObjOxE.CodigoObsICMS))
                    End With
                End If

                If ObjOxE.CodigoSTIPI < 0 Then
                    ddlSituacaoTributariaIPI.SelectedIndex = 0
                Else
                    ddlSituacaoTributariaIPI.SelectedValue = ObjOxE.CodigoSTIPI
                End If
            Else

                If ObjOxE.CodigoSTICMS = 999 OrElse ObjOxE.EstadoOrigem.Length = 0 Then
                    ddlSituacaoTributariaICMS.SelectedIndex = 0
                Else
                    ddlSituacaoTributariaICMS.SelectedValue = ObjOxE.CodigoSTICMS
                End If

                If ObjOxE.CodigoObsICMS <= 0 Then
                    ddlObsICMS.SelectedIndex = 0
                Else
                    With ddlObsICMS
                        .SelectedIndex = .Items.IndexOf(.Items.FindByValue(ObjOxE.CodigoObsICMS))
                    End With
                End If

                If ObjOxE.CodigoSTIPI < 0 Then
                    ddlSituacaoTributariaIPI.SelectedIndex = 0
                Else
                    ddlSituacaoTributariaIPI.SelectedValue = ObjOxE.CodigoSTIPI
                End If
            End If

            If ObjOxE.CodigoSTPISCOFINS = 0 Then
                ddlSituacaoTributariaPISCOFINS.SelectedIndex = 0
            Else
                ddlSituacaoTributariaPISCOFINS.SelectedValue = ObjOxE.CodigoSTPISCOFINS
            End If

            If ObjOxE.CodigoNaturezaDeRendimento = 0 Then
                ddlNaturezaDeRendimento.SelectedValue = 0
            Else
                ddlNaturezaDeRendimento.SelectedValue = ObjOxE.CodigoNaturezaDeRendimento
            End If

            If ObjOxE.CodigoSTPISCOFINS > 0 Then
                ddl.Carregar(ddlObsPISCOFINS, CarregarDDL.Tabela.SituacaoTributariaPISCOFINSObs, "SituacaoTributariaPISCOFINS_Id = " & ddlSituacaoTributariaPISCOFINS.SelectedValue)
                If ObjOxE.CodigoObsPISCOFINS > 0 Then
                    ddlObsPISCOFINS.SelectedValue = ObjOxE.CodigoObsPISCOFINS
                Else
                    ddlObsPISCOFINS.SelectedIndex = 0
                End If
            Else
                ddlObsPISCOFINS.Items.Clear()
            End If
        Else
            ddlSituacaoTributariaICMS.SelectedIndex = 0
            ddlSituacaoTributariaIPI.SelectedIndex = 0
            ddlSituacaoTributariaPISCOFINS.SelectedIndex = 0
            ddlObsPISCOFINS.Items.Clear()
        End If

        If ObjOxE.CodigoSTIBSCBS < 0 Then
            ddlSituacaoTributariaIBSCBS.SelectedIndex = 0
        Else
            ddlSituacaoTributariaIBSCBS.SelectedValue = ObjOxE.CodigoSTIBSCBS
        End If

        CarregarClassificacaoTributariaIBSCBS()

        If ObjOxE.CodigoClassificacaoIBSCBS > 0 Then
            ddlClassificacaoIBSCBS.SelectedValue = ObjOxE.CodigoClassificacaoIBSCBS
        End If

        chkCalculadora.Checked = ObjOxE.UsarCalculadoraDeImposto

        GridEncargos.DataSource = ObjOxE.Encargos
        GridEncargos.DataBind()

        TabContainer1.ActiveTabIndex = 0
    End Sub

    Private Sub Consultar()
        If Funcoes.VerificaPermissao("OperacoesXEncargos", "LEITURA") Then
            '      "  FROM " & IIf(IsNumeric(txtCodigoConfig.Text) AndAlso CInt(txtCodigoConfig.Text) > 0, "OperacaoxEstado OE", "VW_OperacoesVigentes OE") & vbCrLf & _

            Sql = "Select OE.Codigo_id, OE.InicioVigencia," & vbCrLf &
                  "       OE.Empresa       AS Empresa," & vbCrLf &
                  "       OE.GrupoProduto  AS Grupo," & vbCrLf &
                  "       GE.Descricao     AS NomeGrupo," & vbCrLf &
                  "       OE.Produto       AS Produto," & vbCrLf &
                  "       isnull(P.Nome, '') as NomeProduto," & vbCrLf &
                  "       OE.Operacao      AS Operacao," & vbCrLf &
                  "       OE.SubOperacao   AS SubOperacao," & vbCrLf &
                  "       SO.Descricao     AS NomeOperacao," & vbCrLf &
                  "       OE.GrupoFiscal," & vbCrLf &
                  "       OE.CodigoFiscal  AS CFOP," & vbCrLf &
                  "       OE.EstadoOrigem  AS Origem," & vbCrLf &
                  "       OE.EstadoDestino AS Destino," & vbCrLf &
                  "       isnull(OE.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento," & vbCrLf &
                  "       CASE" & vbCrLf &
                  "         WHEN ISNULL(OE.Ativo,0) = 1" & vbCrLf &
                  "             THEN 'S'" & vbCrLf &
                  "             ELSE 'N'" & vbCrLf &
                  "         END AS Ativo" & vbCrLf &
                  "  FROM OperacaoxEstado OE" & vbCrLf &
                  " INNER JOIN SubOperacoes SO" & vbCrLf &
                  "    ON OE.Operacao    = SO.Operacao_Id" & vbCrLf &
                  "   AND OE.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                  "  LEFT JOIN GruposDeEstoques GE" & vbCrLf &
                  "    ON OE.GrupoProduto = GE.Grupo_Id" & vbCrLf &
                  "  LEFT JOIN Produtos P" & vbCrLf &
                  "    ON OE.Produto = P.Produto_Id" & vbCrLf &
                  " Where OE.GrupoProduto <> 'xxxxx'" & vbCrLf &
                  "   and isnull(P.Situacao,1) = 1 " & vbCrLf &
                  "   and OE.Empresa = '" & ddlEmpresa.SelectedValue & "'" & vbCrLf

            If IsNumeric(txtCodigoConfig.Text) AndAlso CInt(txtCodigoConfig.Text) > 0 Then
                Sql &= "   and OE.Codigo_id =" & txtCodigoConfig.Text
            Else
                If chkAtivo.Checked Then
                    Sql &= "  and OE.Ativo = 1"
                Else
                    Sql &= "  and OE.Ativo = 0"
                End If

                If DdlGruposDeEstoques.SelectedIndex > 0 Then
                    Sql &= "  and OE.GrupoProduto = '" & DdlGruposDeEstoques.SelectedValue & "'"
                End If

                If DdlProdutos.SelectedIndex > 0 Then
                    Sql &= "  and OE.Produto = '" & DdlProdutos.SelectedValue & "'"
                End If

                If DdlOperacoes.SelectedIndex > 0 Then
                    Sql &= "  and OE.Operacao = " & DdlOperacoes.SelectedValue & ""
                End If

                If DdlSubOperacoes.SelectedIndex > 0 Then
                    Dim subs As String() = DdlSubOperacoes.SelectedValue.Split("-")
                    Sql &= "  and OE.SubOperacao = " & subs(1) & ""
                End If

                If DdlEstadoOrigem.SelectedIndex > 0 Then
                    Sql &= "  and OE.EstadoOrigem = '" & DdlEstadoOrigem.SelectedValue & "'"
                End If

                If DdlEstadoDestino.SelectedIndex > 0 Then
                    Sql &= "  and OE.EstadoDestino = '" & DdlEstadoDestino.SelectedValue & "'"
                End If
            End If

            Sql &= " Order by OE.GrupoProduto, OE.Produto, OE.Operacao, OE.SubOperacao"

            DS = Banco.ConsultaDataSet(Sql, "OP")

            If DS.Tables(0).Rows.Count > 0 Then
                GridConsulta.DataSource = DS
                GridConsulta.DataBind()

                If DS.Tables(0).Rows.Count = 1 Then
                    Dim Parametros As New OperacaoXEstado
                    Parametros.Codigo = DS.Tables(0).Rows(0)("Codigo_id")

                    ObjOxE = New OperacaoXEstado(Parametros)
                    ObjListVersoes = New ListOperacaoXEstado(Parametros)

                    If Funcoes.VerificaPermissao("OperacoesXEncargos", "LIBERAR") Then
                        chkAtivo.Enabled = True
                    Else
                        chkAtivo.Enabled = False
                    End If

                    SessaoSalvarObjeto()
                    SessaoSalvarVersoes()
                    CarregarFormComAClasse()

                    If ObjOxE.CodigoNaturezaDeRendimento > 0 Then
                        BuscarNaturezaDeRendimentos(ObjOxE)

                        chkNaturezaDeRendimento.Checked = True
                        divNaturezaDeRendimento.Visible = True

                        ddlNaturezaDeRendimento.SelectedValue = ObjOxE.CodigoNaturezaDeRendimento
                    End If

                Else
                    TabContainer1.ActiveTabIndex = 1
                End If
            Else
                MsgBox(Me.Page, "Não foram encontrados Registros para essa seleção")
            End If
        Else
            MsgBox(Me.Page, "Usuario sem permissão para Consultar Registro")
        End If
    End Sub

    Public Function VerificaECarregaSeExistir() As Boolean
        SessaoRecuperarObjeto()

        If ObjOxE.Empresa.Length > 0 _
         And ObjOxE.CodigoGrupoProduto.Length > 0 _
         And ObjOxE.CodigoOperacao > 0 _
         And ObjOxE.CodigoSubOperacao > 0 _
         And ObjOxE.EstadoOrigem.Length > 0 _
         And ObjOxE.EstadoDestino.Length > 0 Then

            Dim Parametros As New OperacaoXEstado
            Parametros.Empresa = ObjOxE.Empresa
            Parametros.CodigoGrupoProduto = ObjOxE.CodigoGrupoProduto
            Parametros.CodigoProduto = ObjOxE.CodigoProduto
            Parametros.CodigoOperacao = ObjOxE.CodigoOperacao
            Parametros.CodigoSubOperacao = ObjOxE.CodigoSubOperacao
            Parametros.EstadoOrigem = ObjOxE.EstadoOrigem
            Parametros.EstadoDestino = ObjOxE.EstadoDestino
            Parametros.CodigoNaturezaDeRendimento = ObjOxE.CodigoNaturezaDeRendimento
            Dim OXE As New OperacaoXEstado(Parametros, True)

            If OXE.Encargos.Count > 1 Then
                MsgBox(Me.Page, "Configuracao ja existe para estes parametros, a Operacao sera carregada na tela principal e sera feita uma consulta com os mesmos parametros.")
                ObjOxE = OXE
                txtCodigoConfig.Text = OXE.Codigo
                ObjListVersoes = New [Lib].Negocio.ListOperacaoXEstado(Parametros)
                'Consultar() 'Carregava a consulta na tab Consulta com os novos parametros informados

                SessaoSalvarObjeto()
                SessaoSalvarVersoes()

                CarregarFormComAClasse()
                Return True
            Else
                If IsNumeric(txtCodigoConfig.Text) AndAlso CInt(txtCodigoConfig.Text) > 0 Then
                    txtVigencia.Text = Now.ToString("dd/MM/yyyy")
                End If
                txtCodigoConfig.Text = ""
                txtCodigoConfig.ReadOnly = True
                ObjListVersoes = Nothing
                ddlVersao.Items.Clear()
                'Consultar()
                SessaoSalvarVersoes()
                SessaoSalvarObjeto()
                Return False
            End If
        Else
            Return False
        End If
    End Function

#Region "Metodos Gravacao"
    Private Sub GravaRegistro()

        Dim valida As String = ValidaCampos()

        If Not String.IsNullOrWhiteSpace(valida) Then
            MsgBox(Me.Page, valida)
            Exit Sub
        End If

        SessaoRecuperarObjeto()

        ObjOxE.IUD = "I"
        ObjOxE.Ativo = True

        Dim Sqls As New ArrayList
        Dim Sql As String
        Sql = "Update OperacaoXEstado set" & vbCrLf &
              "       ativo = 0 " & vbCrLf &
              " Where Empresa                     ='" & ObjOxE.Empresa & "'" & vbCrLf &
              "   And GrupoProduto                ='" & ObjOxE.CodigoGrupoProduto & "'" & vbCrLf &
              "   And Produto                     ='" & ObjOxE.CodigoProduto & "'" & vbCrLf &
              "   And Operacao                    = " & ObjOxE.CodigoOperacao & vbCrLf &
              "   And SubOperacao                 = " & ObjOxE.CodigoSubOperacao & vbCrLf &
              "   And EstadoOrigem                ='" & ObjOxE.EstadoOrigem & "'" & vbCrLf &
              "   And EstadoDestino               ='" & ObjOxE.EstadoDestino & "'" & vbCrLf &
              "   And InicioVigencia              >'" & ObjOxE.InicioVigencia.ToString("yyyy-MM-dd") & "'"
        Sqls.Add(Sql)

        ObjOxE.SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            txtCodigoConfig.Text = ""
            Consultar()

            txtIdVersao.Text = GridConsulta.Rows(0).Cells(1).Text
            lnkConsultarNotas_Click(Nothing, Nothing)
            ddlVersaoRec.SelectedValue = txtIdVersao.Text
            CarregarNotas()

            'Limpar()
            Mensagem = "Registro: " & ObjOxE.Codigo & vbCrLf &
                       ", Gravado/Atualizado com sucesso."
            MsgBox(Me.Page, Mensagem)
        Else
            MsgBox(Me.Page, Mensagem)
        End If

        TabContainer1.ActiveTabIndex = 2
    End Sub

#End Region

    Private Function ValidaCampos() As String
        SessaoRecuperarObjeto()

        If Not IsDate(txtVigencia.Text) Then
            Return "Data de Inicio de Vigencia invalida..."
        Else
            ObjOxE.InicioVigencia = txtVigencia.Text
        End If

        If ObjOxE.Empresa.Trim.Length = 0 Then Return "Empresa é Obrigatório."
        If ObjOxE.CodigoGrupoProduto.Trim.Length = 0 Then Return "Grupo Do Produto é Obrigatório."
        If ObjOxE.CodigoOperacao = 0 Then Return "Operação é Obrigatório."
        If ObjOxE.CodigoSubOperacao = 0 Then Return "SubOperação é Obrigatório."
        If ObjOxE.EstadoOrigem.Trim.Length = 0 Then Return "UF Origem é Obrigatório."
        If ObjOxE.EstadoDestino.Trim.Length = 0 Then Return "UF Destino é Obrigatório."
        If ObjOxE.CodigoFiscal = 0 Then Return "CFOP é Obrigatório."
        If ddlSituacaoTributariaIPI.SelectedItem.Text.Length = 0 Then Return "Sit. Tributaria IPI e Obrigatorio"
        If ddlSituacaoTributariaICMS.SelectedItem.Text.Length = 0 Then Return "Sit. Tributaria ICMS e Obrigatorio"
        If ObjOxE.CodigoSTPISCOFINS = 0 Then Return "Sit. Tributaria PIS/COFINS e Obrigatorio"

        If ObjOxE.CodigoBeneficio.Length > 0 AndAlso Not Left(ObjOxE.CodigoBeneficio, 2) = ObjOxE.EstadoOrigem Then
            Return "Código do Benefício diferente do Estado de Origem, verifique."
        End If

        'Somente para suboperações marcadas como contábil
        If ObjOxE.SubOperacao.Contabil Then
            '**************** VALIDAÇÃO DAS CONTAS *********************
            'New regra
            'Encargo PRODUTO  Se Entrada então deve ter Conta Débito somente
            '                 Se Saída   então deve ter Conta Crédito somente

            'Sinais + e -
            'Encargo AFIXAR  Deve ter Conta Débito e Conta Crédito
            'Outros Encargos Devem ter Somente Conta Débito ou Conta Crédito

            'Sinal de Igual (=)
            'Exceto encargo LIQUIDO e Operação 80 (Fretes)
            'Tanto Conta Débito quanto Conta Crédito devem estar vazios ou preenchidos

            For Each row As GridViewRow In GridEncargos.Rows
                Dim ddlSinal As DropDownList = CType(row.FindControl("DdlSinal"), DropDownList)
                Dim txtDebito As TextBox = CType(row.FindControl("txtDebito"), TextBox)
                Dim txtCredito As TextBox = CType(row.FindControl("txtCredito"), TextBox)
                Dim lblEncargo As Label = CType(row.FindControl("lblEncargo"), Label)

                If ddlSinal.SelectedIndex = 0 Then
                    Return "Informe o Sinal para o Encargo " & lblEncargo.Text
                End If

                If lblEncargo.Text.Equals("PRODUTO") Then
                    If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso (Not String.IsNullOrWhiteSpace(txtCredito.Text) OrElse String.IsNullOrWhiteSpace(txtDebito.Text)) Then
                        Return "Informe somente a conta de débito para o encargo PRODUTO."
                    ElseIf ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso (Not String.IsNullOrWhiteSpace(txtDebito.Text) OrElse String.IsNullOrWhiteSpace(txtCredito.Text)) Then
                        Return "Informe somente a conta de crédito para o encargo PRODUTO."
                    End If
                ElseIf ddlSinal.SelectedValue.Equals("+") OrElse ddlSinal.SelectedValue.Equals("-") Then
                    If lblEncargo.Text.Equals("AFIXAR") Then
                        If String.IsNullOrWhiteSpace(txtDebito.Text.Trim()) OrElse String.IsNullOrWhiteSpace(txtCredito.Text.Trim()) Then _
                            Return "Informe conta Débito e Conta Crédito para o Encargo AFIXAR."
                    Else
                        If (Not String.IsNullOrWhiteSpace(txtDebito.Text.Trim()) AndAlso Not String.IsNullOrWhiteSpace(txtCredito.Text.Trim())) OrElse
                            String.IsNullOrWhiteSpace(txtDebito.Text.Trim()) AndAlso String.IsNullOrWhiteSpace(txtCredito.Text.Trim()) Then _
                            Return "Informe somente a conta Débito ou Conta Crédito no encargo " & lblEncargo.Text
                    End If
                ElseIf Not lblEncargo.Text.Equals("LIQUIDO") AndAlso ddlSinal.SelectedValue.Equals("=") AndAlso ObjOxE.SubOperacao.CodigoOperacao <> 80 Then
                    If String.IsNullOrWhiteSpace(txtDebito.Text.Trim()) AndAlso Not String.IsNullOrWhiteSpace(txtCredito.Text.Trim()) OrElse
                       Not String.IsNullOrWhiteSpace(txtDebito.Text.Trim()) AndAlso String.IsNullOrWhiteSpace(txtCredito.Text.Trim()) Then
                        Return "Encargos com sinal de Igual podem ter os campos de Conta Débito e Conta Crédito ambos preenchidos ou vazios. Encargo: " & lblEncargo.Text
                    End If
                End If
            Next
            '****************************************************************************
        End If

        Dim i As Integer = 0
        For Each row In ObjOxE.Encargos

            Dim objPlaConta As [Lib].Negocio.PlanoDeConta

            'CONTA DEBITO
            If CType(GridEncargos.Rows(i).FindControl("txtDebito"), TextBox).Text.Length > 0 Then
                objPlaConta = New [Lib].Negocio.PlanoDeConta("", 0, CType(GridEncargos.Rows(i).FindControl("txtDebito"), TextBox).Text)
                If objPlaConta.Titulo.Length = 0 Then
                    Return "Conta Débito " & CType(GridEncargos.Rows(i).FindControl("txtDebito"), TextBox).Text & " informada no Encargo " & GridEncargos.Rows(i).Cells(1).Text() & " não existe no Plano de Contas, verifique."
                Else
                    row.CodigoDebitaConta = CType(GridEncargos.Rows(i).FindControl("txtDebito"), TextBox).Text
                End If
            Else
                row.CodigoDebitaConta = ""
            End If

            'CONTA DE CREDITO
            If CType(GridEncargos.Rows(i).FindControl("txtCredito"), TextBox).Text.Length > 0 Then
                objPlaConta = New [Lib].Negocio.PlanoDeConta("", 0, CType(GridEncargos.Rows(i).FindControl("txtCredito"), TextBox).Text)
                If objPlaConta.Titulo.Length = 0 Then
                    Return "Conta Crédito " & CType(GridEncargos.Rows(i).FindControl("txtCredito"), TextBox).Text & " informada no Encargo " & GridEncargos.Rows(i).Cells(1).Text() & " não existe no Plano de Contas, verifique."
                Else
                    row.CodigoCreditaConta = CType(GridEncargos.Rows(i).FindControl("txtCredito"), TextBox).Text
                End If
            Else
                row.CodigoCreditaConta = ""
            End If

            If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso ObjOxE.CodigoSTPISCOFINS < 50 Then
                Return "CST PIS/COFINS deve ser igual ou maior que 50"
            End If

            If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso ObjOxE.CodigoSTPISCOFINS > 49 Then
                Return "CST PIS/COFINS deve ser menor que 50"
            End If

            'BASE DE CALCULO
            If Not IsNumeric(CType(GridEncargos.Rows(i).FindControl("txtBase"), TextBox).Text) Then
                Return "Informe um Percentual Valido para a Base de Calculo"
            Else
                row.AliquotaBase = CType(GridEncargos.Rows(i).FindControl("txtBase"), TextBox).Text
            End If

            'ALIQUOTA
            If Not IsNumeric(CType(GridEncargos.Rows(i).FindControl("txtAliquota"), TextBox).Text) Then
                Return "Informe um Valor Valido para a Aliquota do Encargo"
            Else
                row.Aliquota = CType(GridEncargos.Rows(i).FindControl("txtAliquota"), TextBox).Text
            End If

            'ALIQUOTA EXIBICAO
            If Not IsNumeric(CType(GridEncargos.Rows(i).FindControl("txtAliqExib"), TextBox).Text) Then
                Return "Informe um Valor Valido para a Aliquota Exibicao do Encargo"
            Else
                row.AliquotaExibicao = CType(GridEncargos.Rows(i).FindControl("txtAliqExib"), TextBox).Text
            End If

            'PERCENTUAL LIMITE
            If Not IsNumeric(CType(GridEncargos.Rows(i).FindControl("txtLimite"), TextBox).Text) Then
                Return "Informe um Valor Valido para o Limite de variacao do valor do Encargo"
            Else
                row.AliquotaLimite = CType(GridEncargos.Rows(i).FindControl("txtLimite"), TextBox).Text
            End If

            'SINAL
            row.Sinal = CType(GridEncargos.Rows(i).FindControl("DdlSinal"), DropDownList).SelectedValue

            '*********************************************************************************************
            '*********************************************************************************************
            '*********************************************************************************************
            i += 1
        Next

        If ObjOxE.Encargos.Any(Function(s) s.CodigoEncargo = "ICMS") AndAlso
            Not (ObjOxE.CodigoSTICMS = 0 OrElse ObjOxE.CodigoSTICMS = 10 OrElse ObjOxE.CodigoSTICMS = 20 OrElse ObjOxE.CodigoSTICMS = 51 OrElse
                 ObjOxE.CodigoSTICMS = 100 OrElse ObjOxE.CodigoSTICMS = 151 OrElse
                 ObjOxE.CodigoSTICMS = 200 OrElse ObjOxE.CodigoSTICMS = 251 OrElse
                 ObjOxE.CodigoSTICMS = 300 OrElse ObjOxE.CodigoSTICMS = 351 OrElse
                 ObjOxE.CodigoSTICMS = 400 OrElse ObjOxE.CodigoSTICMS = 451 OrElse
                 ObjOxE.CodigoSTICMS = 500 OrElse ObjOxE.CodigoSTICMS = 551 OrElse
                 ObjOxE.CodigoSTICMS = 600 OrElse ObjOxE.CodigoSTICMS = 610 OrElse ObjOxE.CodigoSTICMS = 620 OrElse ObjOxE.CodigoSTICMS = 651 OrElse ObjOxE.CodigoSTICMS = 690) Then
            MsgBox(Me.Page, "Para o Encargo ICMS o ST ICMS deve ser 0, 10, 20 ou 51 - TRIBUTADO", eTitulo.Info)
            Return False
        End If

        If (ObjOxE.CodigoSTICMS = 0 OrElse ObjOxE.CodigoSTICMS = 10 OrElse ObjOxE.CodigoSTICMS = 20 OrElse
            ObjOxE.CodigoSTICMS = 100 OrElse ObjOxE.CodigoSTICMS = 200 OrElse ObjOxE.CodigoSTICMS = 500 OrElse
            ObjOxE.CodigoSTICMS = 600 OrElse ObjOxE.CodigoSTICMS = 610 OrElse ObjOxE.CodigoSTICMS = 620) AndAlso
        Not ObjOxE.Encargos.Any(Function(s) s.CodigoEncargo = "ICMS" Or s.CodigoEncargo = "ICMS A REC.") Then
            MsgBox(Me.Page, "Para ST ICMS TRIBUTADO INTEGRALMENTE deve existir o Encargo ICMS", eTitulo.Info)
            Return False
        End If

        If ObjOxE.Encargos.Any(Function(s) s.CodigoEncargo = "ICMS" And s.Aliquota = 0) Then
            MsgBox(Me.Page, "Selecione a Alíquota o Encargo ICMS", eTitulo.Info)
            Return False
        End If

        If ObjOxE.Encargos.Any(Function(s) s.CodigoEncargo = "PIS" And s.Aliquota = 0) Then
            MsgBox(Me.Page, "Selecione a Alíquota o Encargo PIS", eTitulo.Info)
            Return False
        End If

        If ObjOxE.Encargos.Any(Function(s) s.CodigoEncargo = "COFINS" And s.Aliquota = 0) Then
            MsgBox(Me.Page, "Selecione a Alíquota o Encargo COFINS", eTitulo.Info)
            Return False
        End If

        If ddlBeneficioICMS.SelectedValue.Length > 0 AndAlso Not Left(ddlBeneficioICMS.SelectedValue, 2) = ObjOxE.EstadoOrigem.Trim Then
            MsgBox(Me.Page, "Código de Benefício não pode ser usado para o Estado de Origem selecionado", eTitulo.Info)
            Return False
        End If

        SessaoSalvarObjeto()

        Return ""
    End Function

    Private Sub CarregarGrupoFiscal()
        If DdlOperacoes.SelectedIndex > 0 And DdlSubOperacoes.SelectedIndex > 0 And DdlEstadoOrigem.SelectedIndex > 0 And DdlEstadoDestino.SelectedIndex > 0 Then
            Dim Inicial As String

            SessaoRecuperarObjeto()

            'Verifica a inicial
            If ObjOxE.EstadoOrigem = ObjOxE.EstadoDestino Then
                If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    Inicial = "1"
                Else
                    Inicial = "5"
                End If
            Else
                If ObjOxE.EstadoDestino = "EX" Then
                    If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                        Inicial = "3"
                    Else
                        Inicial = "7"
                    End If
                Else
                    If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                        Inicial = "2"
                    Else
                        Inicial = "6"
                    End If
                End If
            End If

            ddl.Carregar(DdlCFOPTitulos, CarregarDDL.Tabela.CFOPGrupo, "left(convert(nvarchar,GrupoCfop_Id), 1) = '" & Inicial & "'", True)

            If Left(ObjOxE.CodigoFiscal, 1) <> Inicial Then
                ObjOxE.CodigoGrupoFiscal = 0
                ObjOxE.CodigoFiscal = 0
                DdlCFOP.Items.Clear()
            Else
                DdlCFOPTitulos.SelectedValue = ObjOxE.CodigoGrupoFiscal
                CFOP()
                If ObjOxE.CodigoFiscal > 0 Then DdlCFOP.SelectedValue = ObjOxE.CodigoFiscal
            End If

            SessaoSalvarObjeto()
        Else
            DdlCFOPTitulos.Items.Clear()
            DdlCFOP.Items.Clear()
        End If
    End Sub

    Protected Sub imgExcluir_Click(sender As Object, e As ImageClickEventArgs)
        Try
            SessaoRecuperarObjeto()
            Dim img As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

            Dim enc = CType(GridEncargos.Rows(row.RowIndex).FindControl("lblEncargo"), Label).Text
            Dim i = ObjOxE.Encargos.FindIndex(Function(s) s.CodigoEncargo = enc)
            ObjOxE.Encargos.RemoveAt(i)
            SessaoSalvarObjeto()
            GridEncargos.DataSource = ObjOxE.Encargos
            GridEncargos.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlVersao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlVersao.SelectedIndexChanged
        Try
            Dim Parametros As New OperacaoXEstado
            Parametros.Codigo = ddlVersao.SelectedValue
            ObjOxE = New OperacaoXEstado(Parametros)
            SessaoSalvarObjeto()
            CarregarFormComAClasse()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarNotas_Click(sender As Object, e As EventArgs) Handles lnkConsultarNotas.Click
        Try
            Dim Sql As String
            Sql = "Select 0 as Codigo, 'Selecione uma Versão Para ser Recontabilizada' Descricao " & vbCrLf &
                  " Union " & vbCrLf &
                  "Select Codigo_id as Codigo," & vbCrLf &
                  "       convert(varchar,Codigo_id) + ' | Op.'  + convert(varchar,Operacao) + ' - '  + convert(varchar,Suboperacao) + ' | Inic.Vigencia: '  + convert(varchar,InicioVigencia,103) + ' | Inclusao: '  + convert(varchar,UsuarioInclusaoData,103) as Descricao" & vbCrLf &
                  "  from OperacaoXEstado" & vbCrLf &
                  " Where Ativo = 1" & vbCrLf

            If IsNumeric(txtIdVersao.Text) AndAlso CInt(txtIdVersao.Text) > 0 Then
                Sql &= "  and codigo_id =" & txtIdVersao.Text & vbCrLf
            ElseIf IsDate(txtPeriodoInicial.Text) And IsDate(txtPeriodoFinal.Text) Then
                Sql &= "  and UsuarioInclusaoData between '" & CDate(txtPeriodoInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtPeriodoFinal.Text).ToString("yyyy-MM-dd") & "'"
            End If

            Dim ds As DataSet
            ds = Banco.ConsultaDataSet(Sql, "OxE")

            ddlVersaoRec.DataTextField = "DESCRICAO"
            ddlVersaoRec.DataValueField = "CODIGO"

            ddlVersaoRec.DataSource = ds.Tables(0)
            ddlVersaoRec.DataBind()
            If ds.Tables(0).Rows.Count = 1 Then CarregarNotas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparRec_Click(sender As Object, e As EventArgs) Handles lnkLimparRec.Click
        Try
            txtIdVersao.Text = String.Empty
            txtPeriodoInicial.Text = String.Empty
            txtPeriodoFinal.Text = String.Empty

            If ddlVersaoRec.Items.Count > 0 Then ddlVersaoRec.Items.Clear()

            gridNotas.DataSource = Nothing
            gridNotas.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlVersaoRec_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlVersaoRec.SelectedIndexChanged
        Try
            CarregarNotas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarNotas()
        If ddlVersaoRec.SelectedIndex = 0 Then Exit Sub
        Dim DataInicio As Date
        Dim DataFim As Date = Date.Now
        Dim strVersoes As String = ""
        Dim Posicao As Integer

        Dim Parametros As New OperacaoXEstado
        Parametros.Codigo = ddlVersaoRec.SelectedValue

        Dim listOp As New ListOperacaoXEstado(Parametros)
        Posicao = listOp.FindIndex(Function(s) s.Codigo = Parametros.Codigo)

        'Data Inicio é a data da Versao que vc esta tentando recontabilizar
        DataInicio = listOp.Where(Function(s) s.Codigo = Parametros.Codigo).FirstOrDefault.InicioVigencia
        'Para Descubrir as versoes que deverao se alteradas, vc primeiro de deve descubir a posicao dele na lista que ja esta ordenada por vigencia
        'e voltar ate chegar no inicio da lista ou em uma configuracao que esteja com a propriedade ativo marcada
        If Posicao > 0 Then
            For x As Integer = Posicao - 1 To 0 Step -1
                If listOp(x).Ativo Then
                    Dim Dt = listOp(x).InicioVigencia
                    For Each row In listOp.Where(Function(s) s.InicioVigencia = Dt)
                        strVersoes &= IIf(String.IsNullOrEmpty(strVersoes), "", ",") & row.Codigo
                    Next
                    Exit For
                Else
                    strVersoes &= IIf(String.IsNullOrEmpty(strVersoes), "", ",") & listOp(x).Codigo
                End If
            Next
        End If

        If Posicao <> listOp.Count - 1 Then
            For x As Integer = Posicao + 1 To listOp.Count - 1 Step 1
                If listOp(x).Ativo Then
                    DataFim = listOp(x).InicioVigencia
                    Exit For
                Else
                    strVersoes &= IIf(String.IsNullOrEmpty(strVersoes), "", ",") & listOp(x).Codigo
                End If
            Next
        End If

        Dim sql As String
        sql = "Select nfi.OperacaoXEstado," & vbCrLf &
              "       nf.Empresa_Id," & vbCrLf &
              "       nf.Cliente_id," & vbCrLf &
              "       nf.endCliente_id," & vbCrLf &
              "       c.nome," & vbCrLf &
              "       nf.Movimento," & vbCrLf &
              "       nf.Nota_id," & vbCrLf &
              "       nf.Serie_Id," & vbCrLf &
              "       nf.EntradaSaida_id," & vbCrLf &
              "       nfi.Produto_Id," & vbCrLf &
              "       P.nome NomeProduto" & vbCrLf &
              "  from notasfiscais nf" & vbCrLf &
              " inner join NotasFiscaisXItens nfi" & vbCrLf &
              "    on nf.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
              "   and nf.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
              "   and nf.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
              "   and nf.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
              "   and nf.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
              "   and nf.Nota_Id         = nfi.Nota_Id" & vbCrLf &
              "   and nf.Serie_Id        = nfi.Serie_Id" & vbCrLf &
              " inner join Clientes C" & vbCrLf &
              "    on C.Cliente_id  = nf.cliente_id" & vbCrLf &
              "   and C.Endereco_Id = nf.EndCliente_Id" & vbCrLf &
              " inner join Produtos P" & vbCrLf &
              "    on P.Produto_Id = nfi.Produto_Id" & vbCrLf &
              " where nf.Movimento >= '" & DataInicio.ToString("yyyy-MM-dd") & "'" & vbCrLf &
              "   and nfi.OperacaoXEstado in (" & IIf(strVersoes.Length = 0, "0", strVersoes) & ")" & vbCrLf &
              "   and nf.Movimento  <= '" & DataFim.ToString("yyyy-MM-dd") & "'" & vbCrLf &
              "   and nf.Situacao IN (1,4,7) " & vbCrLf &
              "  Order by nf.Movimento, nf.empresa_id, nf.cliente_id, nf.nota_id"
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Vrs")

        gridNotas.DataSource = ds.Tables(0)
        gridNotas.DataBind()
    End Sub

    Protected Sub chkNotaAll_CheckedChanged(sender As Object, e As EventArgs)
        Try
            If gridNotas.Rows.Count > 0 Then
                Dim chkTitulo As CheckBox

                For Each rowgrid As GridViewRow In gridNotas.Rows
                    chkTitulo = CType(rowgrid.FindControl("Chknota"), CheckBox)
                    chkTitulo.Checked = CType(sender, CheckBox).Checked
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Protected Sub lnkRecontabilizar_Click(sender As Object, e As EventArgs) Handles lnkRecontabilizar.Click
    '    Try
    '        Dim Parametros As New OperacaoXEstado
    '        Parametros.Codigo = ddlVersaoRec.SelectedValue
    '        Dim OXE As New OperacaoXEstado(Parametros)

    '        Dim sql As String = ""
    '        Dim sqls As New ArrayList

    '        For Each row As GridViewRow In gridNotas.Rows
    '            If CType(row.FindControl("Chknota"), CheckBox).Checked Then
    '                Dim nf As New [Lib].Negocio.NotaFiscal()
    '                nf.CodigoEmpresa = row.Cells(2).Text
    '                nf.EnderecoEmpresa = 0
    '                nf.CodigoCliente = row.Cells(3).Text
    '                nf.EnderecoCliente = row.Cells(4).Text
    '                nf.Codigo = row.Cells(7).Text
    '                nf.Serie = row.Cells(8).Text
    '                nf.EntradaSaida = IIf(row.Cells(9).Text = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
    '                nf = New [Lib].Negocio.NotaFiscal(nf)

    '                nf.Itens.ForEach(Function(s)
    '                                     If s.CodigoProduto = row.Cells(10).Text Then
    '                                         s.CodigoOperacaoEstado = ddlVersaoRec.SelectedValue
    '                                         s.Encargos = Nothing
    '                                         If s.CFOP <> OXE.CodigoFiscal Then
    '                                             s.CFOP = OXE.CodigoFiscal
    '                                         End If
    '                                     End If
    '                                     Return True
    '                                 End Function)

    '                For Each itens In nf.Itens.Where(Function(s) s.CodigoProduto = row.Cells(10).Text)
    '                    itens.IUD = "U"
    '                    itens.SalvarSql(sqls)
    '                Next

    '                If nf.NFG Then
    '                    Dim cLote As Integer = 0
    '                    If nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or _
    '                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ReciboDeFrete) Or _
    '                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Or _
    '                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or _
    '                        nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Anulacao) Then
    '                        cLote = 21
    '                    Else
    '                        cLote = 9
    '                    End If

    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 9)
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 10)
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 11)
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 21)
    '                    nf.Razao.NotaFiscalRazao = nf
    '                    nf.Razao.ContabilizarNotaSql(sqls, cLote)
    '                Else
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 9)
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 10)
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 11)
    '                    nf.Razao.ExcluiContabilizacaoNotaSQL(sqls, 21)
    '                    nf.Razao.NotaFiscalRazao = nf
    '                    nf.Razao.ContabilizarNotaSql(sqls, 10)
    '                End If
    '            End If
    '        Next

    '        If Banco.GravaBanco(sqls) Then
    '            MsgBox(Me.Page, "Processo realizado com Sucesso.", eTitulo.Sucess)
    '        Else
    '            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Private Sub updateAtivo(codigoID As Integer, ativo As Boolean)
        Dim sql As String

        sql = " UPDATE OperacaoXEstado SET Ativo = "

        If ativo Then sql &= "1" Else sql &= "0"

        sql &= " WHERE Codigo_Id = " & codigoID

        If Banco.GravaBanco(sql) Then
            If ativo Then
                MsgBox(Me.Page, "ID ativada com sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "ID desativada com sucesso.", eTitulo.Sucess)
            End If
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Protected Sub Ativo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If CInt(txtCodigoConfig.Text) <> 0 Then
            updateAtivo(txtCodigoConfig.Text, chkAtivo.Checked)
            Limpar()
            'Consultar()
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("objOperacoesXEncargosDebito" & HID.Value) IsNot Nothing Then
            CType(GridEncargos.Rows(ViewState("index")).FindControl("txtDebito"), TextBox).Text = CType(HttpContext.Current.Session("objOperacoesXEncargosDebito" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
            'txtNomeDepreciacaoDebito.Text = CType(HttpContext.Current.Session("objAtivosXContasDepreciacaoDebito" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
            Session.Remove("objOperacoesXEncargosDebito" & HID.Value)
        ElseIf Session("objOperacoesXEncargosCredito" & HID.Value) IsNot Nothing Then
            CType(GridEncargos.Rows(ViewState("index")).FindControl("txtCredito"), TextBox).Text = CType(HttpContext.Current.Session("objOperacoesXEncargosCredito" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
            'txtNomeDepreciacaoCredito.Text = CType(HttpContext.Current.Session("objAtivosXContasDepreciacaoCredito" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
            Session.Remove("objOperacoesXEncargosCredito" & HID.Value)
        End If

    End Sub

    Protected Sub imbContaDebito_Click(sender As Object, e As EventArgs)
        Try
            Dim imgLaudoDePesagem As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgLaudoDePesagem.NamingContainer, GridViewRow)
            ViewState("index") = row.RowIndex

            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objOperacoesXEncargosDebito" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbContaCredito_click(sender As Object, e As EventArgs)
        Try
            Dim imgLaudoDePesagem As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgLaudoDePesagem.NamingContainer, GridViewRow)
            ViewState("index") = row.RowIndex

            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objOperacoesXEncargosCredito" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function vereficaEncargos(txtconta As TextBox) As Boolean
        If txtconta.Text.Length <> 7 AndAlso txtconta.Text.Length <> 9 Then
            MsgBox(Me.Page, "Conta de grupo deve conter 7 ou 9 posições.")
            Return False
        End If

        Dim objConta As New PlanoDeConta("99999999999999", 0, txtconta.Text)
        If String.IsNullOrWhiteSpace(objConta.Conta) Then
            MsgBox(Me.Page, "Conta de grupo não encontrada.")
            Return False
        ElseIf txtconta.Text.Length = 9 AndAlso objConta.TemCliente Then
            MsgBox(Me.Page, "Conta de grupo não permitida se houver clientes.")
            Return False
        ElseIf txtconta.Text.Length = 7 AndAlso Not objConta.TemCliente Then
            MsgBox(Me.Page, "conta de grupo não permitida se não houver clientes.")
            Return False
        Else
            txtconta.ToolTip = objConta.Titulo
            Return True
        End If
    End Function

    Protected Sub txtDebito_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim txt As TextBox = CType(sender, TextBox)
            If Not String.IsNullOrWhiteSpace(txt.Text) AndAlso Not vereficaEncargos(txt) Then
                txt.Text = String.Empty
                txt.Focus()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtCredito_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim txt As TextBox = CType(sender, TextBox)
            If Not String.IsNullOrWhiteSpace(txt.Text) AndAlso Not vereficaEncargos(txt) Then
                txt.Text = String.Empty
                txt.Focus()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlSinal_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim row As GridViewRow = CType(sender.NamingContainer, GridViewRow)
        Dim lblEncargo As String = CType(row.FindControl("lblEncargo"), Label).Text
        Dim txtDebito As TextBox = CType(row.FindControl("txtDebito"), TextBox)
        Dim txtCredito As TextBox = CType(row.FindControl("txtCredito"), TextBox)

        Dim imgCredito As ImageButton = CType(row.FindControl("imbContaCredito"), ImageButton)
        Dim imgDebito As ImageButton = CType(row.FindControl("imbContaDebito"), ImageButton)

        Dim ddlSinal As DropDownList = sender

        SessaoRecuperarObjeto()

        Dim objEncargo = ObjOxE.Encargos.Find(Function(f) f.CodigoEncargo = lblEncargo)

        txtCredito.Text = objEncargo.CodigoCreditaConta
        txtDebito.Text = objEncargo.CodigoDebitaConta

        txtCredito.Visible = True
        imgCredito.Visible = True
        txtDebito.Visible = True
        imgDebito.Visible = True

        If ddlSinal.SelectedIndex > 0 Then
            objEncargo.Sinal = ddlSinal.SelectedValue
            If ObjOxE.SubOperacao IsNot Nothing Then
                If ObjOxE.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If ddlSinal.SelectedValue.Equals("+") Then
                        txtCredito.Text = String.Empty
                        txtCredito.Visible = False
                        imgCredito.Visible = False
                    ElseIf ddlSinal.SelectedValue.Equals("-") Then
                        txtDebito.Text = String.Empty
                        txtDebito.Visible = False
                        imgDebito.Visible = False
                    End If
                Else
                    If ddlSinal.SelectedValue.Equals("+") Then
                        txtDebito.Text = String.Empty
                        txtDebito.Visible = False
                        imgDebito.Visible = False
                    ElseIf ddlSinal.SelectedValue.Equals("-") Then
                        txtCredito.Text = String.Empty
                        txtCredito.Visible = False
                        imgCredito.Visible = False
                    End If
                End If
            End If
        End If
        SessaoSalvarObjeto()
    End Sub

    Protected Sub ddlObservacao_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            SessaoRecuperarObjeto()
            Dim ddl As DropDownList = CType(sender, DropDownList)
            Dim row As GridViewRow = CType(ddl.NamingContainer, GridViewRow)
            Dim lblEncargo As String = CType(row.FindControl("lblEncargo"), Label).Text

            Dim objEncargo = ObjOxE.Encargos.Find(Function(f) f.CodigoEncargo = lblEncargo)
            objEncargo.ObservacaoTributaria = IIf(String.IsNullOrWhiteSpace(ddl.SelectedValue), "0", ddl.SelectedValue)

            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlObsICMS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlObsICMS.SelectedIndexChanged
        Try
            SessaoRecuperarObjeto()
            If ddlObsICMS.SelectedIndex > 0 Then
                ObjOxE.CodigoObsICMS = ddlObsICMS.SelectedValue
            Else
                ObjOxE.CodigoObsICMS = 0
            End If
            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlBeneficioICMS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlBeneficioICMS.SelectedIndexChanged
        Try
            SessaoRecuperarObjeto()
            If ddlBeneficioICMS.SelectedIndex > 0 Then
                ObjOxE.CodigoBeneficio = ddlBeneficioICMS.SelectedValue
            Else
                ObjOxE.CodigoBeneficio = String.Empty
            End If

            If ObjOxE.Codigo > 0 AndAlso ObjOxE.CodigoBeneficio.Length > 0 Then
                If Funcoes.VerificaPermissao("OperacoesXEncargos", "LIBERAR") Then
                    updateBeneficioICMS(ObjOxE.Codigo, ObjOxE.CodigoBeneficio)
                End If
            End If

            SessaoSalvarObjeto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub updateBeneficioICMS(ByVal codigoID As Integer, ByVal BeneficioICMS As String)
        Dim sql As String

        sql = " UPDATE OperacaoXEstado SET CodigoBeneficio = '" & BeneficioICMS & "'" & vbCrLf
        sql &= " WHERE Codigo_Id = " & codigoID

        If Banco.GravaBanco(sql) Then
            MsgBox(Me.Page, "Código do Benefício atualizado com sucesso. Não CLICK em GRAVAR para não gerar outra versão.", eTitulo.Sucess)
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

End Class