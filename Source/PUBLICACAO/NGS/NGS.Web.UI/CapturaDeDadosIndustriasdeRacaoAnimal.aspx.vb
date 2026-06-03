Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CapturaDeDadosIndustriasdeRacaoAnimal
    Inherits BasePage

    Dim Sqls As New ArrayList
    Dim Sqll As String
    Dim SqlAux As String
    Dim Empresa() As String
    Dim Array As New ArrayList
    Dim pegaLadrao As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/ApuracaoDeCustos.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("CapturaDeDados", "ACESSAR") Then
                If (Left(Session("ssEmpresa"), 8) = "40938762") OrElse (Left(Session("ssEmpresa"), 8) = "49673784") Then

                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.EmpresasConsolidadas, "", False)
                    ddlEmpresa.SelectedValue = Left(HttpContext.Current.Session("ssEmpresa"), 8)

                    Dim AnoInicial As Integer = Date.Now.Year - 5
                    ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, AnoInicial.ToString() & ";10;C", False)

                    Limpar()
                    ddlEmpresa.Focus()
                Else
                    MsgBox(Me.Page, "Empresa sem permissão para acessar essa página!", "~/ApuracaoDeCustos.aspx")
                    Exit Sub
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/ApuracaoDeCustos.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub lnkContabilizar_Click(sender As Object, e As EventArgs) Handles lnkContabilizar.Click
        If Funcoes.VerificaPermissao("CapturaDeDados", "RELATORIO") Then
            If Validar() Then

                If Not Funcoes.verificaAcessoCusto(ddlEmpresa.SelectedValue, 0, DdlMesFinal.SelectedValue, DdlAno.SelectedValue, "CONTABIL") Then
                    MsgBox(Me.Page, "Movimento Contábil já fechado para esta data!")
                    Exit Sub
                End If

                Try
                    For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue

                        Dim Sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Etapa, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino, Reduzido FROM ApuracaoDeCustos" & vbCrLf
                        Sql &= "  Where ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'" & vbCrLf
                        Sql &= "    And ApuracaoDeCustos.Ano_Id = " & DdlAno.SelectedValue & vbCrLf
                        Sql &= "    And ApuracaoDeCustos.Mes_ID = " & Mes

                        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")

                        If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                            MsgBox(Me.Page, "Apuração de Custos do mês " & Mes & " não foi processado!")
                        Else
                            If Left(Empresa(0), 8) = "40938762" Then
                                ContabilizaCustos(Mes)
                                AjustesRazaoCustoIndireto(Mes)
                            Else
                                Contabiliza(Mes)
                            End If
                        End If
                    Next
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão executar essa ação!")
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("CapturaDeDados", "RELATORIO") Then
            If Validar() Then
                Try
                    If Not Funcoes.verificaAcessoCusto(ddlEmpresa.SelectedValue, 0, DdlMesFinal.SelectedValue, DdlAno.SelectedValue, "CONTABIL") Then
                        MsgBox(Me.Page, "Movimento já Fechado para esta data...")
                        Exit Sub
                    End If

                    Dim erroAux As Boolean
                    ChkResultado.Items.Clear()
                    For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
                        ChkResultado.Items.Add(New ListItem("*********************************************************************"))
                        ChkResultado.Items.Add(New ListItem("************  Mês: " & Mes & "  Ano: " & DdlAno.SelectedValue))
                        ChkResultado.Items.Add(New ListItem("*********************************************************************"))

                        '****************************************************************************************
                        '***************************  LIMPAR MOVIMENTO DO MES   *********************************
                        '****************************************************************************************
                        If LimpaMovimentoDoMes(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Movimento do mês limpo"))
                        Else
                            ChkResultado.Items.Add(New ListItem("Erro ao limpar o movimento do mês"))
                        End If

                        pegaLadrao = "Movimento do mês limpo"

                        '****************************************************************************************
                        '*****************************   APURAR SALDO INICIAL   *********************************
                        '****************************************************************************************
                        erroAux = ApuraSaldoInicial(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Saldos iniciais apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao de saldos iniciais"))
                        End If

                        '****************************************************************************************
                        '********************************   CAPTURAR RAZAO   ************************************
                        '****************************************************************************************
                        erroAux = CapturaRazao(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Dados do razão apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos dados do razao"))
                        End If

                        pegaLadrao = "Dados do razão apurados"

                        '****************************************************************************************
                        '********************************   CAPTURAR ESTOQUES  **********************************
                        '****************************************************************************************
                        erroAux = CapturaEstoques(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Dados do estoque apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos dados do estoque"))
                        End If


                        If Left(Empresa(0), 8) = "44005444" Then
                            '****************************************************************************************
                            '********************************   CAPTURAR CUSTOS INDIRETOS****************************
                            '****************************************************************************************
                            erroAux = CapturaCustosIndiretos(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Custos Indiretos apurados"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos Custos Indiretos"))
                            End If
                        End If

                        pegaLadrao = "Custos Indiretos apurados"


                        '****************************************************************************************
                        '*****************************   CAPTURAR NOTAS FISCAIS  ********************************
                        '****************************************************************************************
                        erroAux = CapturaNotasFiscais(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Dados das notas fiscais apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos dados das notas fiscais"))
                        End If

                        pegaLadrao = "Dados das notas fiscais apurados"

                        If Not Left(Empresa(0), 8) = "40938762" _
                             Then
                            '****************************************************************************************
                            '************************* CAPTURAR NOTAS FISCAIS CONTRAPARTIDA *************************
                            '****************************************************************************************
                            erroAux = CapturaNotasFiscaisContraPartida(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Dados das notas fiscais contrapartida apurados"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos dados das notas fiscais contrapartida"))
                            End If

                        End If

                        pegaLadrao = "Dados das notas fiscais contrapartida apurados"


                        If Left(Empresa(0), 8) = "40938762" Then
                            '****************************************************************************************
                            '*****************************   CAPTURAR NOTAS FISCAIS  DE TRANSFERENCIAS **************
                            '****************************************************************************************
                            If Transferencias(Mes) Then
                                ChkResultado.Items.Add(New ListItem("Dados das transferencias apurados"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & "Erro(s) na apuracao das notas de transferencias"))
                            End If

                            AjustaTotalizadores(Mes)

                            pegaLadrao = "Dados das transferencias apurados"

                            '****************************************************************************************
                            '***************************   AJUSTAR CUSTOS DE SAIDA   ********************************
                            '****************************************************************************************
                            erroAux = AjustaCustosDeSaidas(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Custos de saida ajustador"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nos custos de saída"))
                            End If

                            pegaLadrao = "Custos de saida ajustador"

                            '****************************************************************************************
                            '***************************   AJUSTAR PRODUTO DERIVADO   *******************************
                            '****************************************************************************************
                            erroAux = AjustaProdutoDerivado(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Produto derivado ajustado"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do produto derivado"))
                            End If

                            pegaLadrao = "Produto derivado ajustado"

                            erroAux = AjustaProdutoDerivadoNovo(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Limpa Valores Executado"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante Limpa Valores"))
                            End If

                            pegaLadrao = "Limpa Valores Executado"

                            AjustaTotalizadores(Mes)

                            pegaLadrao = "Ajusta Totalizadores"

                            '****************************************************************************************
                            '***************************   AJUSTAR CONSUMO X PRODUCAO   *****************************
                            '****************************************************************************************
                            erroAux = AjustaConsumoXProducao(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Consumo x producao ajustado"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do consumo x Producao"))
                            End If

                            pegaLadrao = "Consumo x producao ajustado"
                        End If


                        erroAux = AjustaTotalizadores(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Ajustado Totalizadores"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes dos totalizadores"))
                        End If

                        pegaLadrao = "Ajusta Totalizadores 2"

                        '****************************************************************************************
                        '********************************  CALCULO CIRCULAR   ***********************************
                        '****************************************************************************************
                        ChkResultado.Items.Add(New ListItem("Iniciando Calculo Circular"))
                        For ii = 0 To CInt(txtCiclos.Text)

                            If ii = 2 Then
                                CapturaCustosIndiretos(Mes)
                            End If

                            pegaLadrao = "CapturaCustosIndiretos " & ii

                            erroAux = AjustaFormacaoDeLotes(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Formação De Lotes"))

                            pegaLadrao = "AjustaFormacaoDeLotes " & ii

                            erroAux = AjustaCustosDeSaidas(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nos custos de saida"))

                            pegaLadrao = "AjustaCustosDeSaidas " & ii

                            erroAux = AjustaProdutoDerivado(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem("Erro(s) durante os ajustes de Produto derivado"))

                            pegaLadrao = "AjustaProdutoDerivado " & ii

                            erroAux = AjustaMutuos(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes dos mutuos"))

                            pegaLadrao = "AjustaMutuos " & ii

                            erroAux = AjustaTransferencias(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nas saidas por transferencias"))

                            pegaLadrao = "AjustaTransferencias " & ii

                            erroAux = AjustaConsumoXProducao(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do consumo x producao"))

                            pegaLadrao = "AjustaConsumoXProducao " & ii

                            erroAux = AjustaTotalizadores(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Totalizadores"))

                            pegaLadrao = "AjustaTotalizadores " & ii

                        Next

                        erroAux = AjustaCentavos(Mes)
                        If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Centavos"))

                        pegaLadrao = "AjustaCentavos "

                        erroAux = ContabilizaCentroCusto(Mes)
                        If erroAux = False Then
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante a contabilização dos custos"))
                        Else
                            ChkResultado.Items.Add(New ListItem(" contabilização dos custos"))
                        End If

                        pegaLadrao = "contabilização dos custos "

                        ChkResultado.Items.Add(New ListItem(" FIM - Custo Processado com Sucesso"))
                    Next
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão executar essa ação!")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Private Sub Limpar()
        DdlAno.SelectedValue = IIf(Now.Month = 1, Now.Year - 1, Now.Year)
        DdlMesInicial.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)
        DdlMesFinal.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)
        ChkResultado.Items.Clear()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function Validar()
        Empresa = ddlEmpresa.SelectedValue.Split("-")
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório!")
            Return False
        End If

        If DdlMesInicial.SelectedValue <> DdlMesFinal.SelectedValue Then
            MsgBox(Me.Page, "Mês Inicial e final devem ser iguais.")
            Return False
        End If

        Return True
    End Function

    Private Sub CapturarDados()
        Dim MesInicial As Integer
        Dim MesFinal As Integer

        MesInicial = Left(DdlMesInicial.SelectedValue, 2)
        MesFinal = Left(DdlMesFinal.SelectedValue, 2)

        For Mes = MesInicial To MesFinal
            LimpaMovimentoDoMes(Mes)
            ApuraSaldoInicial(Mes)
            CapturaRazao(Mes)
            CapturaEstoques(Mes)
            CapturaNotasFiscais(Mes)
            CapturaNotasFiscaisContraPartida(Mes)
            LimpaValores(Mes)
            CapturaSaidasPorTransferencias(Mes)
            'AjustaMovimento (Mes)
            AjustaTotalizadores(Mes)

            If Not String.IsNullOrWhiteSpace(txtCiclos.Text) Then
                For index = 0 To CInt(txtCiclos.Text)
                    'txtLoop.value = i
                    AjustaCustosDeSaidas(Mes)
                    AjustaMutuos(Mes)
                    AjustaTransferencias(Mes)
                    AjustaProdutoDerivado(Mes)
                    AjustaConsumoXProducao(Mes)
                    AjustaTotalizadores(Mes)
                Next
            End If

            AjustaCentavos(Mes)
            AjustaMovimento(Mes)
            Contabiliza(Mes)
        Next

        MsgBox(Me.Page, "Processo realizado com Sucesso.", eTitulo.Sucess)
    End Sub

    Private Sub AjusteComplementacaoDePreco(ByVal pMes As Integer)
        Dim sql As String = " SELECT  Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id,  Quantidade, ValorDoProduto"
        sql &= "   Into #Temp "
        sql &= "   from ApuracaoDeCustos "
        sql &= "   where CodigoDeCusto_id = 111"
        sql &= "   and Ano_Id           = " & DdlAno.SelectedValue
        sql &= "   and Mes_Id           = " & pMes

        sql = " UPDATE ApuracaoDeCustos set"
        sql &= "    Quantidade     = AC.Quantidade  - T.Quantidade"
        sql &= " from #Temp T"
        sql &= " Inner Join ApuracaoDeCustos AC"
        sql &= "    on T.Empresa_Id            = AC.Empresa_Id"
        sql &= "   and T.EndEmpresa_Id         = AC.EndEmpresa_Id"
        sql &= "   and T.Deposito_Id           = AC.Deposito_Id "
        sql &= "   and T.EndDeposito_Id        = AC.EndDeposito_Id"
        sql &= "   and T.Ano_Id                = AC.Ano_Id"
        sql &= "   and T.Mes_Id                = AC.Mes_Id "
        sql &= "   and T.Produto_Id            = AC.Produto_Id"
        sql &= "   and T.EmpresaDestino_Id     = AC.EmpresaDestino_Id"
        sql &= "   and T.EndEmpresaDestino_Id  = AC.EndEmpresaDestino_Id"
        sql &= "   and T.DepositoDestino_Id    = AC.DepositoDestino_Id "
        sql &= "   and T.EndDepositoDestino_Id = AC.EndDepositoDestino_Id"
        sql &= "   and T.ProdutoDerivado_Id    = AC.ProdutoDerivado_Id"
        sql &= "   and AC.CodigoDeCusto_Id     = 110"
        sql &= " where AC.Ano_Id               = " & DdlAno.SelectedValue
        sql &= "   and AC.Mes_Id               = " & pMes
    End Sub

    Private Sub AvaliarEstoqueAFixar(ByVal pMes As Integer)
        Dim sql As String
        '" ValorDoProduto = ((ac.Quantidade*P.BaseDeCalculo) * PM.ValorOficial) / PM.basedecalculo" & vbCrLf & _
        sql = " Update ApuracaoDeCustos set" & vbCrLf &
              " ValorDoProduto = ((ac.Quantidade * PM.ValorOficial) / PM.basedecalculo" & vbCrLf &
              "  from ApuracaoDeCustos AC" & vbCrLf &
              " inner join TabelaDePrecosDeMercado PM" & vbCrLf &
              "    on AC.Empresa_Id    = PM.Empresa_Id" & vbCrLf &
              "   And AC.EndEmpresa_Id = Pm.EndEmpresa_Id" & vbCrLf &
              "   And AC.Deposito_Id   = PM.Deposito_Id" & vbCrLf &
              "   And AC.EndEmpresa_Id = Pm.EndDeposito_id" & vbCrLf &
              "   And AC.Produto_Id    = PM.Produto_Id" & vbCrLf &
              "   And PM.Data_Id       = '" & New DateTime(DdlAno.SelectedValue, pMes, Date.DaysInMonth(DdlAno.SelectedValue, pMes)) & "'" & vbCrLf &
              " inner join Produtos P" & vbCrLf &
              "    on P.Produto_Id = AC.Produto_Id" & vbCrLf &
              " where AC.CodigoDeCusto_id = 110" & vbCrLf &
              "   and AC.Ano_Id           = " & DdlAno.SelectedValue & vbCrLf &
              "   and AC.Mes_Id           = " & pMes
        Sqls.Add(sql)

        'Reavaliar Quantidade Negativa
        sql = " Update ApuracaoDeCustos set"
        sql &= "  ValorDoProduto = ac.Quantidade * (MesAnterior.ValorDoProduto / MesAnterior.Quantidade)"
        sql &= "  from ApuracaoDeCustos AC"
        sql &= "  Inner Join (Select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto from ApuracaoDeCustos) MesAnterior"
        sql &= "   on AC.Empresa_Id            = MesAnterior.Empresa_Id"
        sql &= "   and AC.EndEmpresa_Id         = MesAnterior.EndEmpresa_Id"
        sql &= "   and AC.Deposito_Id           = MesAnterior.Deposito_Id "
        sql &= "   and AC.EndDeposito_Id        = MesAnterior.EndDeposito_Id"
        sql &= "   and case when AC.Mes_Id = 1 then AC.Ano_Id -1 else Ac.Ano_id end = MesAnterior.Ano_Id"
        sql &= "   and case when AC.Mes_Id = 1 then 12           else Ac.Mes_Id end = MesAnterior.Mes_Id "
        sql &= "   and AC.Produto_Id            = MesAnterior.Produto_Id"
        sql &= "   and AC.EmpresaDestino_Id     = MesAnterior.EmpresaDestino_Id"
        sql &= "   and AC.EndEmpresaDestino_Id  = MesAnterior.EndEmpresaDestino_Id"
        sql &= "   and AC.DepositoDestino_Id    = MesAnterior.DepositoDestino_Id "
        sql &= "   and AC.EndDepositoDestino_Id = MesAnterior.EndDepositoDestino_Id"
        sql &= "   and AC.ProdutoDerivado_Id    = MesAnterior.ProdutoDerivado_Id"
        sql &= "   and AC.CodigoDeCusto_Id      = MesAnterior.CodigoDeCusto_Id"
        sql &= "   where AC.CodigoDeCusto_id = 110"
        sql &= "   and AC.Ano_Id           = " & DdlAno.SelectedValue
        sql &= "   and AC.Mes_Id           = " & pMes
        sql &= "   AND ac.Quantidade       < 0"
        Sqls.Add(sql)
        Banco.GravaBanco(Sqls)
    End Sub

    Private Function LimpaMovimentoDoMes(ByVal pMes As Integer) As Boolean
        Dim aux As Boolean = True
        Dim sql As String = "DELETE FROM ApuracaoDeCustos "
        sql &= " WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= " AND Ano_Id = " & DdlAno.SelectedValue
        sql &= " AND Mes_Id = " & pMes
        aux = Banco.GravaBanco(sql)

        sql = "UPDATE PlanoDeCustos "
        sql &= " SET Classe = '' "
        sql &= " WHERE (Classe IS NULL)"
        aux = Banco.GravaBanco(sql)
        Return aux
    End Function

    Private Function LimpaValores(ByVal pMes As Integer) As Boolean
        Dim sql As String = "  UPDATE ApuracaoDeCustos "
        sql &= " SET ValorDoProduto = 0 "
        sql &= " WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "   AND Ano_Id = " & DdlAno.SelectedValue
        sql &= "   AND Mes_Id = " & pMes

        If Left(Empresa(0), 8) = "05272759" _
            OrElse Left(Empresa(0), 8) = "05366261" _
            OrElse Left(Empresa(0), 8) = "62747840" _
            OrElse Left(Empresa(0), 8) = "38198213" _
            OrElse Left(Empresa(0), 8) = "40938762" _
            OrElse Left(Empresa(0), 8) = "44005444" _
            OrElse Left(Empresa(0), 8) = "48984539" Then
            sql &= "   AND CodigoDeCusto_Id In (357) "
        Else
            sql &= "   AND CodigoDeCusto_Id In (357,364) "
        End If
        'Gerencial não tem o 364 apenas 357 - Furlan 04/05/2020

        If Not Banco.GravaBanco(sql) Then Return False

        sql = "         DELETE FROM ApuracaoDeCustos"
        sql &= "   FROM        ApuracaoDeCustos INNER JOIN"
        sql &= "               PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
        sql &= "   WHERE       ApuracaoDeCustos.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "           AND ApuracaoDeCustos.Ano_Id = " & DdlAno.SelectedValue
        sql &= "           AND ApuracaoDeCustos.Mes_Id = " & pMes
        sql &= "           AND ApuracaoDeCustos.CodigoDeCusto_Id < 500 "
        sql &= "           AND PlanoDeCustos.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'"
        If Not Banco.GravaBanco(sql) Then Return False

        Return True
    End Function

    Private Function AjustaCentavos(ByVal pMes As Integer) As Boolean
        Dim sql As String = "  Update ApuracaoDeCustos "
        sql &= " Set Quantidade = 0, ValorDoProduto = 0 "
        sql &= " WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "   AND Ano_Id = " & DdlAno.SelectedValue
        sql &= "   AND Mes_Id = " & pMes
        sql &= "   AND CodigoDeCusto_Id In (920, 940) And Quantidade < 0.001 and ValorDoProduto > 0"
        Return Banco.GravaBanco(sql)
    End Function

    Private Function AjustaFormacaoDeLotes(ByVal pMes As Integer) As Boolean
        Dim sql As String = "  Delete ApuracaoDeCustos "
        sql &= " WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "   AND Ano_Id = " & DdlAno.SelectedValue
        sql &= "   AND Mes_Id = " & pMes
        sql &= "   AND CodigoDeCusto_Id In (358)"
        Return Banco.GravaBanco(sql)
    End Function

    Private Function ApuraSaldoInicial(ByVal pMes As Integer) As Boolean
        Dim i = 0
        Dim Ano = DdlAno.SelectedValue
        Dim MesAnterior = pMes

        If pMes = 1 Then
            MesAnterior = 12
            Ano = Ano - 1
        Else
            MesAnterior = pMes - 1
        End If

        Dim sql As String = " Select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, "
        sql &= "  DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Etapa, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino, Reduzido"
        sql &= "  FROM ApuracaoDeCustos"
        sql &= "  WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'"
        sql &= "  AND Ano_Id = " & Ano
        sql &= "  AND Mes_Id = " & MesAnterior
        sql &= "  AND CodigoDeCusto_Id = 920"
        sql &= "  AND (Quantidade <> 0 OR ValorDoProduto <> 0 OR ValorDoFrete <> 0 OR ValorAuxiliar <> 0)" & vbCrLf

        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                sql = " set @Exist = (select case "
                sql &= "                        when exists ("
                sql &= "                                      select 1 "
                sql &= "                                        from ApuracaoDeCustos "
                sql &= "                                       Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                sql &= "                                         And EndEmpresa_Id         = " & rs("EndEmpresa_Id")
                sql &= "                                         And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                sql &= "                                         And EndDeposito_Id        = " & rs("EndDeposito_Id")
                sql &= "                                         And Ano_Id                = " & DdlAno.SelectedValue
                sql &= "                                         And Mes_Id                = " & pMes
                sql &= "                                         And Produto_Id            ='" & rs("Produto_Id") & "'"
                sql &= "                                         And CodigoDeCusto_Id      = 101"
                sql &= "                                         And EmpresaDestino_Id     ='" & rs("EmpresaDestino_Id") & "'"
                sql &= "                                         And EndEmpresaDestino_Id  = " & rs("EndEmpresaDestino_Id")
                sql &= "                                         And DepositoDestino_Id    ='" & rs("DepositoDestino_Id") & "'"
                sql &= "                                         And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                sql &= "                                         And ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'"
                sql &= "                                    )"
                sql &= "                           then 'S'"
                sql &= "                           else 'N'"
                sql &= "                       end) ;"
                sql &= " if @Exist = 'N' "
                sql &= " INSERT INTO ApuracaoDeCustos ("
                sql &= "                                Empresa_Id"
                sql &= "                               ,EndEmpresa_Id"
                sql &= "                               ,Deposito_Id"
                sql &= "                               ,EndDeposito_Id"
                sql &= "                               ,Ano_Id"
                sql &= "                               ,Mes_Id"
                sql &= "                               ,Produto_Id"
                sql &= "                               ,CodigoDeCusto_Id"
                sql &= "                               ,EmpresaDestino_Id"
                sql &= "                               ,EndEmpresaDestino_Id"
                sql &= "                               ,DepositoDestino_Id"
                sql &= "                               ,EndDepositoDestino_Id"
                sql &= "                               ,ProdutoDerivado_Id"
                sql &= "                               ,Quantidade"
                sql &= "                               ,ValorDoProduto"
                sql &= "                               ,ValorDoFrete"
                sql &= "                               ,ProdutoDestino"
                sql &= "                               ,CodigoDestino)"
                sql &= " VALUES( '" & rs("Empresa_Id") & "'"
                sql &= "        , " & rs("EndEmpresa_Id")
                sql &= "        ,'" & rs("Deposito_Id") & "'"
                sql &= "        , " & rs("EndDeposito_Id")
                sql &= "        , " & DdlAno.SelectedValue
                sql &= "        , " & pMes
                sql &= "        ,'" & rs("Produto_Id") & "'"
                sql &= "        , 101"
                sql &= "        ,'" & rs("EmpresaDestino_Id") & "'"
                sql &= "        , " & rs("EndEmpresaDestino_Id")
                sql &= "        ,'" & rs("DepositoDestino_Id") & "'"
                sql &= "        , " & rs("EndDepositoDestino_Id")
                sql &= "        ,'" & rs("ProdutoDerivado_Id") & "'"
                sql &= "        , " & Str(rs("Quantidade"))
                sql &= "        , " & Replace(rs("ValorDoProduto"), ",", ".")
                sql &= "        , " & Replace(rs("ValorDoFrete"), ",", ".")
                sql &= "        ,'" & rs("ProdutoDestino") & "'"
                sql &= "        ,0"
                sql &= "       )"
                sql &= " Else"
                sql &= "  Update ApuracaoDeCustos set "
                sql &= "     Quantidade     = " & Str(rs("Quantidade"))
                sql &= "    ,ValorDoProduto = " & Replace(rs("ValorDoProduto"), ",", ".")
                sql &= "    ,ValorDoFrete   = " & Replace(rs("ValorDoFrete"), ",", ".")
                sql &= "  Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                sql &= "    And EndEmpresa_Id         = " & rs("EndEmpresa_Id")
                sql &= "    And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                sql &= "    And EndDeposito_Id        = " & rs("EndDeposito_Id")
                sql &= "    And Ano_Id                = " & DdlAno.SelectedValue
                sql &= "    And Mes_Id                = " & pMes
                sql &= "    And Produto_Id            ='" & rs("Produto_Id") & "'"
                sql &= "    And CodigoDeCusto_Id      = 101"
                sql &= "    And EmpresaDestino_Id     ='" & rs("EmpresaDestino_Id") & "'"
                sql &= "    And EndEmpresaDestino_Id  = " & rs("EndEmpresaDestino_Id")
                sql &= "    And DepositoDestino_Id    ='" & rs("DepositoDestino_Id") & "'"
                sql &= "    And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                sql &= "    And ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'"
                Sqll = Sqll & sql
            Next
        End If
        Return Banco.GravaBanco(Sqll)
    End Function

    Private Function CapturaRazao(ByVal pMes As Integer) As Boolean
        Dim i = 0
        Dim Sql As String = "SELECT Ordem, Empresa_Id, EndEmpresa_Id, EmpresaSinal_Id, Conta_Id, ContaSinal_Id, Produto_Id, ProdutoSinal_Id, Lote_Id, LoteSinal_Id, Placus_Id, PlacusSinal_Id, Processo, Empresa, EndEmpresa, Conta, Produto, Lote, Placus, Observacoes, Sql from ApuracaoDeCustosXFiltroRazao"
        SqlAux = "; "

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                i = 0
                If rs("Processo") = "D" Then
                    SqlAux = SqlAux & "Delete #Razao where "
                End If

                If rs("Processo") = "U" Then
                    SqlAux = SqlAux & " Update #Razao "
                    SqlAux = SqlAux & " Set Produto = '" & rs("Produto") & "'"
                    SqlAux = SqlAux & " where "
                End If

                If rs("Empresa_Id") <> "" Then
                    If rs("EmpresaSinal_Id") = "LIKE" Then
                        SqlAux = SqlAux & " Empresa_Id Like ('" & rs("Empresa_Id") & "%')"
                        i = i + 1
                    Else
                        SqlAux = SqlAux & " Empresa_Id " & rs("EmpresaSinal_Id") & " '" & rs("Empresa_Id") & "'"
                        i = i + 1
                    End If
                End If

                If rs("Conta_Id") <> "" Then
                    If rs("ContaSinal_Id") = "LIKE" Then
                        SqlAux = SqlAux & IIf(i = 0, " Conta_Id LIKE " & "('" & rs("Conta_Id") & "%')", " And Conta_Id  LIKE " & "('" & rs("Conta_Id") & "%')")
                        i = i + 1
                    Else
                        SqlAux = SqlAux & IIf(i = 0, " Conta_Id " & rs("ContaSinal_Id") & " '" & rs("Conta_Id") & "'", " And Conta_Id " & rs("ContaSinal_Id") & " '" & rs("Conta_Id") & "'")
                        i = i + 1
                    End If
                End If


                If rs("Produto_Id") <> "" Then
                    If rs("ProdutoSinal_Id") = "LIKE" Then
                        SqlAux = SqlAux & IIf(i = 0, " Produto LIKE " & "('" & rs("Produto_Id") & "%')", " And Produto  LIKE " & "('" & rs("Produto_Id") & "%')")
                        i = i + 1
                    Else
                        SqlAux = SqlAux & IIf(i = 0, " Produto " & rs("ProdutoSinal_Id") & " '" & rs("Produto_Id") & "'", " And Produto " & rs("ProdutoSinal_Id") & " '" & rs("Produto_Id") & "'")
                        i = i + 1
                    End If
                End If


                If rs("Lote_Id") > 0 Then
                    SqlAux = SqlAux & IIf(i = 0, " Lote_Id " & rs("LoteSinal_Id") & " " & rs("Lote_Id"), " And Lote_Id " & rs("LoteSinal_Id") & " " & rs("Lote_Id"))
                    i = i + 1
                End If

                If rs("Placus_Id") > 0 Then
                    SqlAux = SqlAux & IIf(i = 0, " Codigo_Id " & rs("PlacusSinal_Id") & " " & rs("Placus_Id"), " And Codigo_Id " & rs("PlacusSinal_Id") & " " & rs("Placus_Id"))
                End If

                If rs("Sql") <> "" Then
                    SqlAux = SqlAux & IIf(i = 0, " " & rs("sql"), " And " & rs("sql"))
                End If

                SqlAux = SqlAux & "; "
            Next
        End If

        Sql = "          SELECT  Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Codigo_Id, Produto, SUM(DebitoOficial - CreditoOficial) AS Valor" & vbCrLf
        Sql &= "    FROM (" & vbCrLf

        Sql &= "    SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id,  Razao.Conta_Id, Razao.Lote_Id, PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto," & vbCrLf
        Sql &= "    Razao.Empresa_Id AS Deposito," & vbCrLf
        Sql &= "    Razao.EndEmpresa_Id AS EndDeposito," & vbCrLf
        Sql &= "    Sum(Razao.DebitoOficial) As DebitoOficial," & vbCrLf
        Sql &= "    Sum(Razao.CreditoOficial) As CreditoOficial" & vbCrLf
        Sql &= "    FROM   Razao INNER JOIN   PlanoDeCustosXOrigem ON Razao.Conta_Id LIKE PlanoDeCustosXOrigem.Conta_Id + '%' INNER JOIN Produtos" & vbCrLf
        Sql &= "    ON Razao.Produto = Produtos.Produto_Id INNER JOIN GruposDeEstoques    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf

        Sql &= " WHERE  Empresa_Id like '" & Left(Empresa(0), 8) & "%'" & vbCrLf
        Sql &= "        And MONTH(Razao.Movimento_Id) = " & pMes & vbCrLf
        Sql &= "        AND YEAR(Razao.Movimento_Id) = " & DdlAno.SelectedValue & vbCrLf
        Sql &= "        And GruposDeEstoques.custo = 1" & vbCrLf
        Sql &= "        And Razao.Produto <> '' " & vbCrLf
        Sql &= "        And  Razao.Lote_Id not in('7000') And Left(Razao.Conta_Id, 7) <> '1010601'" & vbCrLf
        Sql &= " GROUP  BY Razao.Empresa_Id, Razao.EndEmpresa_Id, " & vbCrLf
        Sql &= "        isnull(Razao.Deposito, Razao.Empresa_Id) , " & vbCrLf
        Sql &= "        isnull(Razao.EndDeposito, Razao.EndEmpresa_Id), " & vbCrLf
        Sql &= "        PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto, Razao.Conta_Id, Razao.Lote_Id " & vbCrLf

        Sql &= " UNION" & vbCrLf

        Sql &= "    SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id,  Razao.Conta_Id, Razao.Lote_Id, PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto," & vbCrLf
        Sql &= "    Razao.Empresa_Id AS Deposito," & vbCrLf
        Sql &= "    Razao.EndEmpresa_Id AS EndDeposito," & vbCrLf
        Sql &= "    Sum(Razao.DebitoOficial) As DebitoOficial," & vbCrLf
        Sql &= "    Sum(Razao.CreditoOficial) As CreditoOficial" & vbCrLf
        Sql &= "    FROM   Razao INNER JOIN   PlanoDeCustosXOrigem ON Razao.Conta_Id LIKE PlanoDeCustosXOrigem.Conta_Id + '%' INNER JOIN Produtos" & vbCrLf
        Sql &= "    ON Razao.Produto = Produtos.Produto_Id INNER JOIN GruposDeEstoques    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf

        Sql &= " WHERE  Empresa_Id like '" & Left(Empresa(0), 8) & "%'" & vbCrLf
        Sql &= "        And MONTH(Razao.Movimento_Id) = " & pMes & vbCrLf
        Sql &= "        AND YEAR(Razao.Movimento_Id) = " & DdlAno.SelectedValue & vbCrLf
        Sql &= "        And GruposDeEstoques.custo = 1" & vbCrLf
        Sql &= "        And Razao.Produto <> '' " & vbCrLf
        Sql &= "        And  Razao.Lote_Id not in(7000, 9, 10) And Left(Razao.Conta_Id, 7) IN ('1010601', '1010603')" & vbCrLf
        Sql &= " GROUP  BY Razao.Empresa_Id, Razao.EndEmpresa_Id, " & vbCrLf
        Sql &= "        isnull(Razao.Deposito, Razao.Empresa_Id) , " & vbCrLf
        Sql &= "        isnull(Razao.EndDeposito, Razao.EndEmpresa_Id), " & vbCrLf
        Sql &= "        PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto, Razao.Conta_Id, Razao.Lote_Id " & vbCrLf

        Sql &= " UNION" & vbCrLf

        Sql &= "    SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id,  Razao.Conta_Id, Razao.Lote_Id, 155, Razao.Produto," & vbCrLf
        Sql &= "    Razao.Empresa_Id AS Deposito," & vbCrLf
        Sql &= "    Razao.EndEmpresa_Id AS EndDeposito," & vbCrLf
        Sql &= "    Sum(Razao.DebitoOficial) As DebitoOficial," & vbCrLf
        Sql &= "    Sum(Razao.CreditoOficial) As CreditoOficial" & vbCrLf
        Sql &= "    FROM   Razao INNER JOIN Produtos" & vbCrLf
        Sql &= "    ON Razao.Produto = Produtos.Produto_Id INNER JOIN GruposDeEstoques    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf

        Sql &= " WHERE  Empresa_Id like '" & Left(Empresa(0), 8) & "%'" & vbCrLf
        Sql &= "        And MONTH(Razao.Movimento_Id) = " & pMes & vbCrLf
        Sql &= "        AND YEAR(Razao.Movimento_Id) = " & DdlAno.SelectedValue & vbCrLf
        Sql &= "        And GruposDeEstoques.custo = 1" & vbCrLf
        Sql &= "        And Razao.Produto <> '' " & vbCrLf
        Sql &= "        And  Razao.Lote_Id in (1) And Left(Razao.Conta_Id, 7) IN ('1010603')" & vbCrLf
        Sql &= " GROUP  BY Razao.Empresa_Id, Razao.EndEmpresa_Id, " & vbCrLf
        Sql &= "        isnull(Razao.Deposito, Razao.Empresa_Id) , " & vbCrLf
        Sql &= "        isnull(Razao.EndDeposito, Razao.EndEmpresa_Id), " & vbCrLf
        Sql &= "        Razao.Produto, Razao.Conta_Id, Razao.Lote_Id " & vbCrLf

        Sql &= " UNION " & vbCrLf

        Sql &= " SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id, Razao.Conta_Id, Razao.Lote_Id, PlanoDeCustosXOrigem.Codigo_Id, ISNULL(PlanoDeCustosXOrigem.Produto, N'')" & vbCrLf
        Sql &= "        AS Produto, Razao.Empresa_Id AS Deposito, Razao.EndEmpresa_Id AS EndDeposito," & vbCrLf
        Sql &= "        SUM(Razao.DebitoOficial) AS DebitoOficial, SUM(Razao.CreditoOficial) AS CreditoOficial" & vbCrLf
        Sql &= " FROM   Razao INNER JOIN" & vbCrLf
        Sql &= "        PlanoDeCustosXOrigem ON Razao.Conta_Id LIKE PlanoDeCustosXOrigem.Conta_Id + '%'" & vbCrLf
        Sql &= " WHERE  Razao.Empresa_Id like '" & Left(Empresa(0), 8) & "%'" & vbCrLf
        Sql &= "        And MONTH(Razao.Movimento_Id) = " & pMes & vbCrLf
        Sql &= "        AND YEAR(Razao.Movimento_Id) = " & DdlAno.SelectedValue & vbCrLf
        Sql &= "        AND (Len(PlanoDeCustosXOrigem.Produto) = 8)" & vbCrLf
        Sql &= "        And  Razao.Lote_Id not in('7000') And Left(Razao.Conta_Id, 7) <> '1010601'" & vbCrLf

        Sql &= " GROUP  BY Razao.Empresa_Id, Razao.EndEmpresa_Id, ISNULL(Razao.Deposito, Razao.Empresa_Id), ISNULL(Razao.EndDeposito, Razao.EndEmpresa_Id)," & vbCrLf
        Sql &= "        PlanoDeCustosXOrigem.Codigo_Id , PlanoDeCustosXOrigem.Produto, Razao.Conta_Id, Razao.Lote_Id" & vbCrLf
        'sql &=  Sqla & "; "

        Sql &= " ) as Consulta INNER JOIN Produtos ON Consulta.Produto LIKE Produtos.Produto_Id" & vbCrLf
        Sql &= " GROUP   BY Empresa_Id, EndEmpresa_Id, Codigo_Id, Produto, Deposito, EndDeposito" & vbCrLf

        '05 - CAPTURANDO DADOS DO RAZÃO - ATUALIZANDO TABELA DE APURAÇÃO"
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds2.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "       when exists ("
                Sql &= "                 select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino from ApuracaoDeCustos "
                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                Sql &= " And Deposito_Id = '" & rs("Deposito") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDeposito")
                Sql &= " And Ano_Id = " & DdlAno.SelectedValue
                Sql &= " And Mes_Id = " & pMes
                Sql &= " And Produto_Id = '" & rs("Produto") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("Codigo_Id")

                Sql &= ")"
                Sql &= "            then 'S'"
                Sql &= "             else 'N'"
                Sql &= "               end) ;"

                Sql &= "  if @Exist = 'N' "
                Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"
                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino)"
                Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito") & "'"
                Sql &= ", " & rs("EndDeposito")
                Sql &= ", " & DdlAno.SelectedValue
                Sql &= ", " & pMes
                Sql &= ", '" & rs("Produto") & "'"
                Sql &= ", " & rs("Codigo_Id")
                Sql &= ", ''" 'Empresa Destino
                Sql &= ", 0"  'End Empresa Destino
                Sql &= ", ''" 'Deposito Destino
                Sql &= ", 0"  'End DepositoDestino
                Sql &= ", ''" 'Produto Derivado

                Sql &= ", 0" 'Quantidade
                Sql &= ", " & Str(rs("Valor"))
                Sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                Sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                Sql &= ", ''" '& rs("ProdutoDestino") & "'"
                Sql &= ", 0" '& rs("CodigoDestino")
                Sql &= ")"

                Sql &= " Else"

                Sql &= "  Update ApuracaoDeCustos set "
                Sql &= "  ValorDoProduto = ValorDoProduto + " & Str(rs("Valor"))
                Sql &= " Where Empresa_Id = '" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                Sql &= "   And Deposito_Id = '" & rs("Deposito") & "'"
                Sql &= "   And EndDeposito_Id = " & rs("EndDeposito")
                Sql &= "   And Ano_Id = " & DdlAno.SelectedValue
                Sql &= "   And Mes_Id = " & pMes
                Sql &= "   And Produto_Id = '" & rs("Produto") & "'"
                Sql &= "   And CodigoDeCusto_Id = " & rs("Codigo_Id")
                Sqll = Sqll & Sql
            Next
        End If
        Return Banco.GravaBanco(Sqll)
    End Function

    'Private Function CapturaCustosIndiretos(ByVal pMes As Integer) As Boolean
    '    Dim sql As String = " EXEC SP_CAPTURA_CUSTOS_INDIRETOS " & DdlAno.SelectedValue & ", " & pMes & ", '" & Left(Empresa(0), 8) & "' "
    '    Return Banco.GravaBanco(sql)
    'End Function

    Private Function CapturaCustosIndiretos(ByVal pMes As Integer) As Boolean
        'txt_Processo.value = "Processando Custos Indiretos..."

        Dim QuantidadeTotal As Double
        Dim ValorTotal As Double
        Dim ValorTotalConsumido As Double
        Dim Valor As Double
        Dim sql As String


        '------------"  Capturando dados de Custos Indiretos

        sql = " SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id, SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS Valor"
        sql &= " FROM  Razao INNER JOIN"
        sql &= "       PlanoDeCustosXOrigem on LEFT(Razao.Conta_Id, 7)  =  PlanoDeCustosXOrigem.Conta_Id "
        sql &= " Where LEFT(Razao.Empresa_Id , 8) like '" & Left(Empresa(0), 8) & "'       "
        sql &= "   And Month(Movimento_Id) = " & pMes
        sql &= "   And YEAR(Movimento_Id)  =  " & DdlAno.SelectedValue
        sql &= "   And Lote_Id <> 7000"
        sql &= "   And isnull(PlanoDeCustosXOrigem.Produto, '0') = 0    "
        sql &= "   And  PlanoDeCustosXOrigem.Codigo_Id IN (402)        "
        sql &= "   Group By Razao.Empresa_Id, Razao.EndEmpresa_Id   "
        sql &= " UNION ALL "
        sql &= " SELECT   Empresa_Id "
        sql &= "                ,EndEmpresa_Id "
        sql &= "                ,SUM(ValorDoProduto) AS Valor "
        sql &= "            FROM ApuracaoDeCustos AC "
        sql &= "            WHERE AC.Ano_id = " & DdlAno.SelectedValue
        sql &= "                AND AC.Mes_id = " & pMes
        sql &= "                AND LEFT(AC.Empresa_Id, 8) LIKE '" & Left(Empresa(0), 8) & "' "
        sql &= "                and ISNULL(ac.ProdutoDerivado_Id,'') = '' "
        sql &= "                AND CodigoDeCusto_Id = 501 "
        sql &= "            GROUP BY Empresa_Id "
        sql &= "                ,EndEmpresa_Id "

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ApuracaoDeCustos")

        For Each rs As DataRow In ds.Tables(0).Rows

            ValorTotal = 0

            If Not IsDBNull(rs("Valor")) Then
                ValorTotal = ValorTotal + rs("Valor")
                ValorTotalConsumido = ValorTotal
            End If


            '--------------------------------------------------------------------
            'Se for Fabrica tem que buscar produção
            If rs("Empresa_Id") = "40938762000239" Then

                sql = "SELECT            p.Empresa_Id                                                   "
                sql &= "                ,p.EndEmpresa_Id                                                "
                sql &= "                ,p.Deposito_Id                                                  "
                sql &= "                ,p.EndDeposito_Id                                               "
                sql &= "                ,p.Produto_Id                                                   "
                sql &= "                ,SUM(p.Entradas) Quantidade                                     "
                sql &= "            FROM Producao p                                                     "
                sql &= "            INNER Join SubOperacoes so ON so.Operacao_Id = p.Operacao_Id        "
                sql &= "                AND so.SubOperacoes_Id = p.SubOperacao_Id                       "
                sql &= "            INNER Join Produtos pro ON pro.Produto_Id = p.Produto_Id            "
                sql &= "            WHERE YEAR(p.Movimento_Id) = " & DdlAno.SelectedValue
                sql &= "                And MONTH(p.Movimento_Id) = " & pMes
                sql &= "                AND p.Empresa_Id    = '40938762000239'                          "
                sql &= "                And so.EntradaSaida = 'E'                                       "
                sql &= "                And p.FisicoFiscal_Id = 2                                       "
                sql &= "                AND pro.CustoIndireto = 1                                       "
                sql &= "                And Not (                                                       "
                sql &= "                    p.Operacao_Id = 40                                          "
                sql &= "                    And p.SubOperacao_Id = 10                                   "
                sql &= "                    )                                                           "
                sql &= "            Group BY p.Empresa_Id                                               "
                sql &= "                ,p.EndEmpresa_Id                                                "
                sql &= "                ,p.Deposito_Id                                                  "
                sql &= "                ,p.EndDeposito_Id                                               "
                sql &= "                ,p.Produto_Id "
                sql &= "   Order by p.Produto_Id    "

            Else

                sql = "SELECT n.Empresa_Id, n.EndEmpresa_Id, n.Deposito AS Deposito_Id, n.EndDeposito AS EndDeposito_Id, NI.Produto_id," & vbCrLf &
                        "	Sum(Case" & vbCrLf &
                        "			When so.Devolucao = 'S'" & vbCrLf &
                        "				Then NI.PesoFiscal * -1" & vbCrLf &
                        "				else NI.PesoFiscal" & vbCrLf &
                        "		end) AS Quantidade" & vbCrLf &
                        "FROM NOTASFISCAISXITENS NI" & vbCrLf &
                        "	INNER JOIN NotasFiscais n" & vbCrLf &
                        "			ON n.Empresa_Id = NI.Empresa_Id" & vbCrLf &
                        "			and n.EndEmpresa_iD = NI.EndEmpresa_Id" & vbCrLf &
                        "			and n.EndEmpresa_iD = NI.EndEmpresa_Id" & vbCrLf &
                        "			and n.Cliente_Id    = NI.Cliente_Id" & vbCrLf &
                        "			and n.EndCliente_Id = NI.EndCliente_Id" & vbCrLf &
                        "			and n.EntradaSaida_Id = NI.EntradaSaida_Id" & vbCrLf &
                        "			and n.Serie_Id        = NI.Serie_Id" & vbCrLf &
                        "			and n.Nota_Id         = NI.Nota_Id" & vbCrLf &
                        "    INNER JOIN SubOperacoes so" & vbCrLf &
                        "        ON so.Operacao_Id      = n.Operacao" & vbCrLf &
                        "		AND so.SubOperacoes_Id = n.SubOperacao" & vbCrLf &
                        "	INNER JOIN Produtos p" & vbCrLf &
                        "			ON p.Produto_Id = NI.Produto_Id" & vbCrLf &
                        "WHERE n.Empresa_Id = '" & rs("Empresa_Id") & "'" & vbCrLf &
                        "  AND n.EndEmpresa_Id = " & rs("EndEmpresa_Id") & vbCrLf &
                        "  AND n.Situacao = 1" & vbCrLf &
                        "  AND n.TipoDeDocumento = 1" & vbCrLf &
                        "  AND MONTH(n.Movimento) =  " & pMes & vbCrLf &
                        "  AND YEAR(n.Movimento)  =  " & DdlAno.SelectedValue & vbCrLf &
                        "  AND so.Operacao_id in(1,7)" & vbCrLf &
                        "GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Deposito, n.EndDeposito, NI.produto_id;"

            End If

            Dim dsa As DataSet = Banco.ConsultaDataSet(sql, "ApuracaoDeCustos")
            If dsa IsNot Nothing Then
                If dsa.Tables(0).Rows.Count > 0 Then
                    QuantidadeTotal = dsa.Tables(0).Compute("sum(Quantidade)", "")
                    'For Each rsa As DataRow In dsa.Tables(0).Rows
                    '    QuantidadeTotal = QuantidadeTotal + rsa("Quantidade")
                    'Next


                    '-----------------" Atualizando Tabela de Apuração"))

                    Sqll = " Declare"
                    Sqll = Sqll & " @Exist as varchar(1)"


                    Dim dsb As DataSet = Banco.ConsultaDataSet(sql, "ApuracaoDeCustos")

                    Dim i As Integer = 1

                    For Each rsb As DataRow In dsb.Tables(0).Rows

                        If i = dsb.Tables(0).Rows.Count Then
                            Valor = ValorTotalConsumido
                        Else
                            Valor = Math.Round((ValorTotal / QuantidadeTotal) * rsb("Quantidade"), 2, MidpointRounding.AwayFromZero)

                            ValorTotalConsumido = Math.Round(ValorTotalConsumido - Valor, 2, MidpointRounding.AwayFromZero)
                        End If

                        'Valor = (ValorTotal / QuantidadeTotal) * rsb("Quantidade")

                        sql = " set @Exist = (select case "
                        sql &= "                        when exists ("
                        sql &= "                                     select 1"
                        sql &= "                                       from ApuracaoDeCustos "
                        sql &= "                                      Where Empresa_Id         ='" & rsb("Empresa_Id") & "'"
                        sql &= "                                        And EndEmpresa_Id      = " & rsb("EndEmpresa_Id")
                        sql &= "                                        And Deposito_Id        ='" & rsb("Deposito_Id") & "'"
                        sql &= "                                        And EndDeposito_Id     = " & rsb("EndDeposito_Id")
                        sql &= "                                        And Ano_Id             = " & DdlAno.SelectedValue
                        sql &= "                                        And Mes_Id             = " & pMes
                        sql &= "                                        And Produto_Id         ='" & rsb("Produto_Id") & "'"
                        sql &= "                                        And ProdutoDerivado_Id =''"
                        sql &= "                                        And CodigoDeCusto_Id   = 402"
                        sql &= "                                        And CodigoDestino      = 0"
                        sql &= "                                     )"
                        sql &= "                           then 'S'"
                        sql &= "                           else 'N'"
                        sql &= "               end) ;"
                        sql &= " if @Exist = 'N' "
                        sql &= "  INSERT INTO ApuracaoDeCustos ( "
                        sql &= "  Empresa_Id"
                        sql &= ", EndEmpresa_Id"
                        sql &= ", Deposito_Id"
                        sql &= ", EndDeposito_Id"
                        sql &= ", Ano_Id"
                        sql &= ", Mes_Id"
                        sql &= ", Produto_Id"
                        sql &= ", CodigoDeCusto_Id"
                        sql &= ", EmpresaDestino_Id"
                        sql &= ", EndEmpresaDestino_Id"
                        sql &= ", DepositoDestino_Id"
                        sql &= ", EndDepositoDestino_Id"
                        sql &= ", ProdutoDerivado_Id"
                        sql &= ", Quantidade"
                        sql &= ", ValorDoProduto"
                        sql &= ", ValorDoFrete"
                        sql &= ", ValorAuxiliar"
                        sql &= ", ProdutoDestino"
                        sql &= ", CodigoDestino"
                        sql &= ", Reduzido)"
                        sql &= " VALUES('" & rsb("Empresa_Id") & "'"
                        sql &= ", " & rsb("EndEmpresa_Id")
                        sql &= ", '" & rsb("Deposito_Id") & "'"
                        sql &= ", " & rsb("EndDeposito_Id")
                        sql &= ", " & DdlAno.SelectedValue
                        sql &= ", " & pMes
                        sql &= ", '" & rsb("Produto_Id") & "'"
                        sql &= ", 402"
                        sql &= ", ''" 'Empresa Destino
                        sql &= ", 0"  'End Empresa Destino
                        sql &= ", ''" 'Deposito Destino
                        sql &= ", 0"  'End DepositoDestino
                        sql &= ", ''" 'Produto Derivado
                        sql &= ", 0"  ' & Replace(rs("Quantidade"), ",", ".")
                        sql &= ", " & Replace(Valor, ",", ".")
                        sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                        sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                        sql &= ", ''" '& rs("ProdutoDestino") & "'"
                        sql &= ", ''" '& rs("PlacusContraPartida")
                        sql &= ", ''" '& rs("Reduzido") & "'"
                        sql &= ")"
                        sql &= " Else"
                        sql &= "  Update ApuracaoDeCustos set "
                        sql &= "  ValorDoProduto = ValorDoProduto + " & Replace(Valor, ",", ".")
                        sql &= " Where Empresa_Id         ='" & rsb("Empresa_Id") & "'"
                        sql &= "   And EndEmpresa_Id      = " & rsb("EndEmpresa_Id")
                        sql &= "   And Deposito_Id        ='" & rsb("Deposito_Id") & "'"
                        sql &= "   And EndDeposito_Id     = " & rsb("EndDeposito_Id")
                        sql &= "   And Ano_Id             = " & DdlAno.SelectedValue
                        sql &= "   And Mes_Id             = " & pMes
                        sql &= "   And Produto_Id         ='" & rsb("Produto_Id") & "'"
                        sql &= "   And ProdutoDerivado_Id =''" '& rs("ProdutoDerivado_Id") & "'"
                        sql &= "   And CodigoDeCusto_Id   = 402" '& rs("Placus_Id")
                        sql &= "   And CodigoDestino      = 0;" '& rs("PlacusContraPartida")

                        Sqll = Sqll & sql

                        i += 1

                    Next

                    Banco.GravaBanco(Sqll)

                End If
            Else

            End If
        Next

        Return True
    End Function

    Private Function AjustesRazaoCustoIndireto(ByVal pMes As Integer) As Boolean

        Dim AnoC As Integer
        AnoC = DdlAno.SelectedValue
        Dim Dia As String
        Dia = Format(CDate("01/" & pMes & "/" & AnoC).AddMonths(+1).AddDays(-1), "dd/MM/yyyy")
        Dia = Format(CDate(Dia), "yyyy-MM-dd")

        Dim strHistorico As String = Format(CDate("01/" & pMes & "/" & AnoC).AddMonths(+1).AddDays(-1), "MM/yyyy")

        Dim sql As String = ""
        sql &= "INSERT INTO Razao (                                                                                                     "
        sql &= "                             [Empresa_Id]                                                                               "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "	                        ,[Conta_Id]                                                                                 "
        sql &= "	                        ,[Cliente_Id]                                                                               "
        sql &= "	                        ,[EndCliente_Id]                                                                            "
        sql &= "	                        ,[Movimento_Id]                                                                             "
        sql &= "	                        ,[Lote_Id]                                                                                  "
        sql &= "	                        ,[Sequencia_Id]                                                                             "
        sql &= "	                        ,[Produto]                                                                                  "
        sql &= "	                        ,[Custo]                                                                                    "
        sql &= "	                        ,[Indexador]                                                                                "
        sql &= "	                        ,[DataMoeda]                                                                                "
        sql &= "	                        ,[DebitoOficial]                                                                            "
        sql &= "	                        ,[CreditoOficial]                                                                           "
        sql &= "	                        ,[DebitoMoeda]                                                                              "
        sql &= "	                        ,[CreditoMoeda]                                                                             "
        sql &= "	                        ,[Historico]                                                                                "
        sql &= "	                        ,[PrevistoRealizado]                                                                        "
        sql &= "	                        ,[Situacao]                                                                                 "
        sql &= "	                        )                                                                                           "
        sql &= "                        Select  [Empresa_Id]                                                                        "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "	                        ,'401020102' AS Conta_Id                                                                    "
        sql &= "	                        ,'' AS Cliente_Id                                                                           "
        sql &= "	                        ,0 AS EndCliente_Id                                                                         "
        sql &= "	                        ,[Movimento_Id]                                                                             "
        sql &= "	                        ,7000 AS Lote_Id                                                                            "
        sql &= "	                        ,ROW_NUMBER() OVER (                                                                        "
        sql &= "		                        Partition BY [Empresa_Id]                                                               "
        sql &= "		                        ,[EndEmpresa_Id]                                                                        "
        sql &= "		                        ,'401020102'                                                                            "
        sql &= "		                        ,''                                                                                     "
        sql &= "		                        ,0                                                                                      "
        sql &= "		                        ,7000                                                                                   "
        sql &= "		                        ,[Movimento_Id] ORDER BY [Movimento_Id]                                                 "
        sql &= "		                        ) AS [Sequencia_Id]                                                                     "
        sql &= "	                        ,'' AS [Produto]                                                                            "
        sql &= "	                        ,0 AS [Custo]                                                                               "
        sql &= "	                        ,3 AS [Indexador]                                                                           "
        sql &= "	                        ,[DataMoeda]                                                                                "
        sql &= "	                        ,0 AS [DebitoOficial]                                                                       "
        sql &= "	                        ,ABS(ISNULL(SUM(ISNULL(debitoOficial, 0) - ISNULL(CreditoOficial, 0)), 0)) AS [CreditoOficial] "
        sql &= "	                        ,0 [DebitoMoeda]                                                                            "
        sql &= "	                        ,0 [CreditoMoeda]                                                                           "
        sql &= "	                        ,CONCAT (                                                                                   "
        sql &= "		                        'CUSTOS PRODUCAO REF: '                                                                 "
        sql &= "		                        ,month(movimento_Id)                                                                    "
        sql &= "		                        ,'/'                                                                                    "
        sql &= "		                        ,year(movimento_Id)                                                                     "
        sql &= "		                        ) AS [Historico]                                                                        "
        sql &= "	                        ,'R' AS [PrevistoRealizado]                                                                 "
        sql &= "	                        ,1 AS [Situacao]                                                                            "
        sql &= "                        FROM razao                                                                                      "
        sql &= "                        WHERE Month(movimento_Id) = " & pMes
        sql &= "	                        AND year(movimento_Id) = " & DdlAno.SelectedValue
        sql &= "	                        And LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'                                     "
        sql &= "	                        And Conta_Id Like '4010201%'                                                                "
        sql &= "                        Group BY [Empresa_Id]                                                                           "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "	                        ,[Movimento_Id]                                                                             "
        sql &= "	                        ,[DataMoeda]                                                                                "
        sql &= "                                                                                                                     "
        sql &= "                        INSERT INTO Razao (                                                                             "
        sql &= "	                        [Empresa_Id]                                                                                "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "	                        ,[Conta_Id]                                                                                 "
        sql &= "	                        ,[Cliente_Id]                                                                               "
        sql &= "	                        ,[EndCliente_Id]                                                                            "
        sql &= "	                        ,[Movimento_Id]                                                                             "
        sql &= "	                        ,[Lote_Id]                                                                                  "
        sql &= "	                        ,[Sequencia_Id]                                                                             "
        sql &= "	                        ,[Produto]                                                                                  "
        sql &= "	                        ,[Custo]                                                                                    "
        sql &= "	                        ,[Indexador]                                                                                "
        sql &= "	                        ,[DataMoeda]                                                                                "
        sql &= "	                        ,[DebitoOficial]                                                                            "
        sql &= "	                        ,[CreditoOficial]                                                                           "
        sql &= "	                        ,[DebitoMoeda]                                                                              "
        sql &= "	                        ,[CreditoMoeda]                                                                             "
        sql &= "	                        ,[Historico]                                                                                "
        sql &= "	                        ,[PrevistoRealizado]                                                                        "
        sql &= "	                        ,[Situacao]                                                                                 "
        sql &= "	                        )                                                                                           "
        sql &= "                        Select [Empresa_Id]                                                                             "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "	                        ,'401020298' AS Conta_Id                                                                    "
        sql &= "	                        ,'' AS Cliente_Id                                                                           "
        sql &= "	                        ,0 AS EndCliente_Id                                                                         "
        sql &= "	                        ,'" & Dia.ToSqlDate() & "' [Movimento_Id]                      "
        sql &= "	                        ,7000 AS Lote_Id                                                                            "
        sql &= "	                        ,(                                                                                          "
        sql &= "		                        Select  ISNULL(MAX(Sequencia_Id), 1) As Sequencia_Id                                "
        sql &= "		                        FROM razao r                                                                            "
        sql &= "                                WHERE Month(movimento_Id) = " & pMes
        sql &= "			                        AND year(movimento_Id) = " & DdlAno.SelectedValue
        sql &= "			                        And LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'                             "
        sql &= "			                        And Conta_Id Like '401020298'                                                       "
        sql &= "		                        ) AS [Sequencia_Id]                                                                     "
        sql &= "	                        ,'' AS [Produto]                                                                            "
        sql &= "	                        ,0 AS [Custo]                                                                               "
        sql &= "	                        ,3 AS [Indexador]                                                                           "
        sql &= "	                        ,'" & Dia.ToSqlDate() & "' AS [DataMoeda]                                                   "
        sql &= "	                        ,ABS(ISNULL(SUM(ISNULL(debitoOficial, 0) - ISNULL(CreditoOficial, 0)), 0)) AS [DebitoOficial] "
        sql &= "	                        ,0 AS [CreditoOficial]                                                                      "
        sql &= "	                        ,0 [DebitoMoeda]                                                                            "
        sql &= "	                        ,0 [CreditoMoeda]                                                                           "
        sql &= "	                        ,'CUSTOS PRODUCAO REF.: " & strHistorico & "' AS [Historico]                                "
        sql &= "	                        ,'R' AS [PrevistoRealizado]                                                                 "
        sql &= "	                        ,1 AS [Situacao]                                                                            "
        sql &= "                        FROM razao                                                                                      "
        sql &= "                        WHERE Month(movimento_Id) = " & pMes
        sql &= "	                        AND year(movimento_Id) = " & DdlAno.SelectedValue
        sql &= "	                        And LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'                                     "
        sql &= "	                        And Conta_Id Like '4010202%'                                                                "
        sql &= "                        Group BY [Empresa_Id]                                                                           "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "                                                                                                                        "
        sql &= "                        INSERT INTO Razao (                                                                             "
        sql &= "	                        [Empresa_Id]                                                                                "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "
        sql &= "	                        ,[Conta_Id]                                                                                 "
        sql &= "	                        ,[Cliente_Id]                                                                               "
        sql &= "	                        ,[EndCliente_Id]                                                                            "
        sql &= "	                        ,[Movimento_Id]                                                                             "
        sql &= "	                        ,[Lote_Id]                                                                                  "
        sql &= "	                        ,[Sequencia_Id]                                                                             "
        sql &= "	                        ,[Produto]                                                                                  "
        sql &= "	                        ,[Custo]                                                                                    "
        sql &= "	                        ,[Indexador]                                                                                "
        sql &= "	                        ,[DataMoeda]                                                                                "
        sql &= "	                        ,[DebitoOficial]                                                                            "
        sql &= "	                        ,[CreditoOficial]                                                                           "
        sql &= "	                        ,[DebitoMoeda]                                                                              "
        sql &= "	                        ,[CreditoMoeda]                                                                             "
        sql &= "	                        ,[Historico]                                                                                "
        sql &= "	                        ,[PrevistoRealizado]                                                                        "
        sql &= "	                        ,[Situacao]                                                                                 "
        sql &= "	                        )                                                                                           "
        sql &= "                        Select [Empresa_Id]                                                            "
        sql &= "	                        ,[EndEmpresa_Id]                                                                        "
        sql &= "	                        ,'401020303' AS Conta_Id                                                                    "
        sql &= "	                        ,'' AS Cliente_Id                                                                           "
        sql &= "	                        ,0 AS EndCliente_Id                                                                         "
        sql &= "	                        ,'" & Dia.ToSqlDate() & "' [Movimento_Id]                                                   "
        sql &= "	                        ,7000 AS Lote_Id                                                                            "
        sql &= "	                        ,(                                                                                          "
        sql &= "		                        Select ISNULL(MAX(Sequencia_Id), 1) As Sequencia_Id                                     "
        sql &= "		                        FROM razao r                                                                            "
        sql &= "                                WHERE Month(movimento_Id) = " & pMes
        sql &= "			                        AND year(movimento_Id) = " & DdlAno.SelectedValue
        sql &= "			                        And LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'                             "
        sql &= "			                        And Conta_Id Like '401020303'                                                       "
        sql &= "		                        ) AS [Sequencia_Id]                                                                     "
        sql &= "	                        ,'' AS [Produto]                                                                            "
        sql &= "	                        ,0 AS [Custo]                                                                               "
        sql &= "	                        ,3 AS [Indexador]                                                                           "
        sql &= "	                        ,'" & Dia.ToSqlDate() & "' AS [DataMoeda]                                                   "
        sql &= "	                        ,ABS(ISNULL(SUM(ISNULL(debitoOficial, 0) - ISNULL(CreditoOficial, 0)), 0)) AS [DebitoOficial]  "
        sql &= "	                        ,0 AS [CreditoOficial]                                                                      "
        sql &= "	                        ,0 [DebitoMoeda]                                                                            "
        sql &= "	                        ,0 [CreditoMoeda]                                                                           "
        sql &= "	                        ,'CUSTOS PRODUCAO REF.: " & strHistorico & "' AS [Historico]                                "
        sql &= "	                        ,'R' AS [PrevistoRealizado]                                                                 "
        sql &= "	                        ,1 AS [Situacao]                                                                            "
        sql &= "                        FROM razao                                                                                      "
        sql &= "                        WHERE Month(movimento_Id)  = " & pMes
        sql &= "	                        AND year(movimento_Id) = " & DdlAno.SelectedValue
        sql &= "	                        And LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'                                     "
        sql &= "	                        And Conta_Id Like '4010203%'                                                                "
        sql &= "                        Group BY [Empresa_Id]                                                                           "
        sql &= "	                        ,[EndEmpresa_Id]                                                                            "

        Return Banco.GravaBanco(sql)


    End Function
    Private Function CapturaEstoques(ByVal pMes As Integer) As Boolean
        Dim i = 0
        Dim aux As Boolean = True

        '07 - CAPTURANDO DADOS DE ESTOQUES - CRIANDO FILTRO
        Dim Sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, ProdutoDerivado_Id, Operacao_Id, SubOperacao_Id from ApuracaoDeCustosXFiltroEstoques"
        SqlAux = "; "

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                i = 0
                SqlAux = SqlAux & "Delete #Estoques where "

                If rs("Empresa_Id") <> "" Then
                    SqlAux = SqlAux & " Empresa_Id like '" & rs("Empresa_Id") & "%'"
                    i = i + 1
                End If

                If rs("Deposito_Id") <> "" Then
                    SqlAux = SqlAux & IIf(i = 0, " Deposito_Id = '" & rs("Deposito_Id"), " And Deposito_Id = " & rs("Deposito_Id"))
                    i = i + 1
                End If

                If rs("Produto_Id") <> "" Then
                    SqlAux = SqlAux & IIf(i = 0, " Produto = " & rs("Produto_Id"), " And Produto = " & rs("Produto_Id"))
                    i = i + 1
                End If

                If rs("ProdutoDerivado_Id") <> "" Then
                    SqlAux = SqlAux & IIf(i = 0, " ProdutoDerivado_Id = " & rs("ProdutoDerivado_Id"), " And ProdutoDerivado_Id = " & rs("ProdutoDerivado_Id"))
                    i = i + 1
                End If

                If rs("Operacao_Id") > 0 Then
                    SqlAux = SqlAux & IIf(i = 0, " Operacao_Id = " & rs("Operacao_Id"), " And Operacao_Id = " & rs("Operacao_Id"))
                    i = i + 1
                End If

                If rs("SubOperacao_Id") > 0 Then
                    SqlAux = SqlAux & IIf(i = 0, " SubOperacao_Id = " & rs("SubOperacao_Id"), " And SubOperacao_Id = " & rs("SubOperacao_Id"))
                End If
                SqlAux = SqlAux & "; "

                aux = Banco.GravaBanco(Sql)
            Next
        End If

        Sql = "SELECT    Consulta.Empresa_Id, Consulta.EndEmpresa_Id, Consulta.Deposito_Id, Consulta.EndDeposito_Id, Consulta.Produto_Id, Consulta.ProdutoDerivado_Id,"
        Sql &= "              Consulta.Placus_Id, Consulta.PlacusContraPartida, SUM(Consulta.Quantidade) AS Quantidade"
        Sql &= "  FROM         (SELECT     Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Operacao_Id,"
        Sql &= "                                      Producao.SubOperacao_Id, Producao.ProdutoDerivado_Id, SubOperacoes.ApuracaoDeCustos AS Placus_Id,"
        Sql &= "                                      ISNULL(SubOperacoes.ApuracaodeCustosContraPartida, 0) AS PlacusContraPartida, "
        Sql &= "               Case when Producao.Entradas <> 0 then Producao.Entradas else Producao.Saidas end AS Quantidade"
        Sql &= "               FROM          Producao INNER JOIN"
        Sql &= "                                      SubOperacoes ON Producao.Operacao_Id = SubOperacoes.Operacao_Id AND Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id INNER JOIN"
        Sql &= "                                      Produtos ON Producao.Produto_Id = Produtos.Produto_Id INNER JOIN"
        Sql &= "                                      GruposDeEstoques ON Produtos.Grupo = GruposDeEstoques.Grupo_Id"

        Sql &= " WHERE Producao.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   AND MONTH(Producao.Movimento_Id) =  " & pMes
        Sql &= "   AND YEAR(Producao.Movimento_Id)  =  " & DdlAno.SelectedValue
        Sql &= "   AND Producao.FisicoFiscal_Id     = 2"
        Sql &= "   AND GruposDeEstoques.custo       = 1"

        'Acrescentado para o Curtume - 11/11/2019 - Furlan
        If Left(Empresa(0), 8) = "03189063" Then
            Sql &= "   AND SubOperacoes.ApuracaoDeCustos > 0"
        End If

        'sql &=  " GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.ProdutoDerivado_Id,"
        'sql &=  "          Producao.Operacao_Id, Producao.SubOperacao_Id, SubOperacoes.ApuracaoDeCustos, ISNULL(SubOperacoes.ApuracaodeCustosContraPartida, 0))"

        Sql &= "          ) AS Consulta INNER JOIN"
        Sql &= "          Produtos ON Consulta.Produto_Id = Produtos.Produto_Id"
        Sql &= " GROUP BY Consulta.Empresa_Id, Consulta.EndEmpresa_Id, Consulta.Deposito_Id, Consulta.EndDeposito_Id, Consulta.Produto_Id, Consulta.ProdutoDerivado_Id,"
        Sql &= "              Consulta.Placus_Id , Consulta.PlacusContraPartida"

        '08 - CAPTURANDO DADOS DE ESTOQUES - ATUALIZANDO TABELA DE APURAÇÃO
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds2.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "                        when exists ("
                Sql &= "                                     select 1"
                Sql &= "                                       from ApuracaoDeCustos "
                Sql &= "                                      Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                Sql &= "                                        And EndEmpresa_Id         = " & rs("EndEmpresa_Id")
                Sql &= "                                        And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                Sql &= "                                        And EndDeposito_Id        = " & rs("EndDeposito_Id")
                Sql &= "                                        And Ano_Id                = " & DdlAno.SelectedValue
                Sql &= "                                        And Mes_Id                = " & pMes
                Sql &= "                                        And Produto_Id            ='" & rs("Produto_Id") & "'"
                Sql &= "                                        And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                Sql &= "                                        And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                Sql &= "                                        And DepositoDestino_Id    ='" & rs("Empresa_Id") & "'"
                Sql &= "                                        And EndDepositoDestino_Id = " & rs("EndEmpresa_Id")
                Sql &= "                                        And ProdutoDerivado_Id ='" & rs("ProdutoDerivado_Id") & "'"
                Sql &= "                                        And CodigoDeCusto_Id   = " & rs("Placus_Id")
                Sql &= "                                        And CodigoDestino      = " & rs("PlacusContraPartida")
                Sql &= "                                     )"
                Sql &= "                           then 'S'"
                Sql &= "                           else 'N'"
                Sql &= "               end) ;"
                Sql &= " if @Exist = 'N' "
                Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"
                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino"
                Sql &= ", Reduzido)"
                Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito_Id") & "'"
                Sql &= ", " & rs("EndDeposito_Id")
                Sql &= ", " & DdlAno.SelectedValue
                Sql &= ", " & pMes
                Sql &= ", '" & rs("Produto_Id") & "'"
                Sql &= ", " & rs("Placus_Id")
                Sql &= ",'" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ",'" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("ProdutoDerivado_Id") & "'"


                If (Left(Empresa(0), 8) = "05366261" _
                    OrElse Left(Empresa(0), 8) = "38198213" _
                    OrElse Left(Empresa(0), 8) = "62747840" _
                    OrElse Left(Empresa(0), 8) = "40938762" _
                    OrElse Left(Empresa(0), 8) = "44005444" _
                    OrElse Left(Empresa(0), 8) = "48984539") AndAlso rs("Produto_Id") = "102010003" Then
                    Dim qdte As Decimal = rs("Quantidade") * 40
                    Sql &= ", " & Replace(qdte, ",", ".")
                Else
                    Sql &= ", " & Replace(rs("Quantidade"), ",", ".")
                End If

                Sql &= ", 0" '& rs("Valor").ToString.Replace(",", ".")
                Sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                Sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                Sql &= ", ''" '& rs("ProdutoDestino") & "'"
                Sql &= ", " & rs("PlacusContraPartida")
                Sql &= ", ''" '& rs("Reduzido") & "'"
                Sql &= ")"
                Sql &= " Else"
                Sql &= "  Update ApuracaoDeCustos set "
                Sql &= "  Quantidade = " & Replace(rs("Quantidade"), ",", ".")
                Sql &= " Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id         = " & rs("EndEmpresa_Id")
                Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                Sql &= "   And Ano_Id                = " & DdlAno.SelectedValue
                Sql &= "   And Mes_Id                = " & pMes
                Sql &= "   And Produto_Id            ='" & rs("Produto_Id") & "'"
                Sql &= "   And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                Sql &= "   And DepositoDestino_Id    ='" & rs("Empresa_Id") & "'"
                Sql &= "   And EndDepositoDestino_Id = " & rs("EndEmpresa_Id")
                Sql &= "   And ProdutoDerivado_Id ='" & rs("ProdutoDerivado_Id") & "'"
                Sql &= "   And CodigoDeCusto_Id   = " & rs("Placus_Id")
                Sql &= "   And CodigoDestino      = " & rs("PlacusContraPartida")
                Sqll = Sqll & Sql
            Next
        End If
        aux = Banco.GravaBanco(Sqll)
        Return aux
    End Function

    Private Function CapturaNotasFiscais(ByVal pMes As Integer) As Boolean
        Dim aux As Boolean = True
        Dim i = 0
        Dim sqlWhere As String

        Dim Sql As String = "SELECT Ordem, Empresa_Id, EndEmpresa_Id, EmpresaSinal_Id, Deposito_Id, EndDeposito_Id, DepositoSinal_Id, DepositoDestino_Id, EndDepositoDestino_Id, DepositoDestinoSinal_Id, GrupoDeEstoque_Id, GrupoDeEstoqueSinal_Id, Produto_Id, ProdutoSinal_Id, Operacao_Id, OperacaoSinal_Id, SubOperacao_Id, SubOperacaoSinal_Id, Placus_Id, PlacusSinal_Id, TipoDeDocumento_Id, TipoDeDocumentoSinal_Id, Processo, Empresa, EndEmpresa, Deposito, EndDeposito, DepositoDestino, EndDepositoDestino, GrupoDeEstoque, Produto, Operacao, SubOperacao, Placus from ApuracaoDeCustosXFiltroNotas order by ordem"
        SqlAux = "; "

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                sqlWhere = ""
                i = 0
                If rs("Processo") = "D" Then
                    SqlAux = SqlAux & "Delete #Notas where "
                End If

                If rs("Processo") = "U" Then
                    SqlAux = SqlAux & " Update #Notas Set "

                    If rs("Produto") <> "" Then
                        SqlAux = SqlAux & " Produto_Id = '" & rs("Produto") & "'"
                    End If

                    If rs("Empresa_Id") <> rs("Empresa") And Len(rs("Empresa")) = 14 Then
                        If rs("Produto") <> "" Then
                            SqlAux = SqlAux & ", "
                        End If

                        SqlAux = SqlAux & " Empresa_Id     ='" & rs("Empresa") & "',"
                        SqlAux = SqlAux & " EndEmpresa_Id  = " & rs("EndEmpresa") & ","

                        If rs("Deposito_Id") <> rs("Deposito") And Len(rs("Empresa")) = 14 Then
                            SqlAux = SqlAux & " Deposito_Id    ='" & rs("Deposito") & "',"
                            SqlAux = SqlAux & " EndDeposito_Id = " & rs("EndDeposito")
                        End If

                        If rs("Deposito_Id") <> "" Then
                            sqlWhere = sqlWhere & " and Deposito_id = '" & rs("Deposito_Id") & "'"
                            sqlWhere = sqlWhere & " and EndDeposito_id = '" & rs("EndDeposito_id") & "'"
                        End If
                    ElseIf rs("Deposito_Id") <> rs("Deposito") And Len(rs("Empresa")) = 14 Then
                        If rs("Produto") <> "" Then
                            SqlAux = SqlAux & ", "
                        End If

                        SqlAux = SqlAux & " Deposito_Id    = '" & rs("Deposito") & "',"
                        SqlAux = SqlAux & " EndDeposito_Id = " & rs("EndDeposito")

                        If rs("Deposito_Id") <> "" Then
                            sqlWhere = sqlWhere & " and Deposito_id = '" & rs("Deposito_Id") & "'"
                            sqlWhere = sqlWhere & " and EndDeposito_id = " & rs("EndDeposito_id")
                        End If
                    End If

                    SqlAux = SqlAux & " where "
                End If

                If rs("Empresa_Id") <> "" Then
                    If rs("EmpresaSinal_Id") = "LIKE" Then
                        SqlAux = SqlAux & " Empresa_Id Like ('" & rs("Empresa_Id") & "%')"
                        i = i + 1
                    Else
                        SqlAux = SqlAux & " Empresa_Id " & rs("EmpresaSinal_Id") & " '" & rs("Empresa_Id") & "'"
                        i = i + 1
                    End If
                End If


                If rs("Produto_Id") <> "" Then
                    If rs("ProdutoSinal_Id") = "LIKE" Then
                        SqlAux = SqlAux & IIf(i = 0, " Produto_Id LIKE " & "('" & rs("Produto_Id") & "%')", " And Produto_Id  LIKE " & "('" & rs("Produto_Id") & "%')")
                        i = i + 1
                    Else
                        SqlAux = SqlAux & IIf(i = 0, " Produto_Id " & rs("ProdutoSinal_Id") & " '" & rs("Produto_Id") & "'", " And Produto_Id " & rs("ProdutoSinal_Id") & " '" & rs("Produto_Id") & "'")
                        i = i + 1
                    End If
                End If

                If rs("Placus_Id") > 0 Then
                    SqlAux = SqlAux & IIf(i = 0, " Codigo_Id " & rs("PlacusSinal_Id") & " " & rs("Placus_Id"), " And Codigo_Id " & rs("PlacusSinal_Id") & " " & rs("Placus_Id"))
                End If

                SqlAux = SqlAux & sqlWhere & "; "
                aux = Banco.GravaBanco(Sql)
            Next
        End If

        Sql = " SELECT     Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, DepositoDestino_Id, EndDepositoDestino_Id, Placus_Id, ISNULL(ApuracaodeCustosContraPartida, 0)"
        Sql &= "              AS ApuracaodeCustosContrapartida, Produto_Id, SUM(Quantidade) AS Quantidade, ISNULL(SUM(CASE WHEN ((EntradaSaida_Id = 'E') OR"
        Sql &= "              (EntradaSaida_Id = 'S' AND Devolucao = 'S')) THEN (ValorDoProduto + DespesasAduaneiras + ValorDespAcessoria + ValorFretes) - (ValorCOFINS + ValorPIS + ValorICMS + ValorDescontos) ELSE ValorDoProduto END), 0) AS ValorDoProduto, "
        Sql &= "              Sum(ValorICMS) as ICMS"
        Sql &= " FROM         ( SELECT     NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, ISNULL(NFI.Deposito, NotasFiscais.Empresa_Id) AS Deposito_Id,"
        Sql &= "                                      ISNULL(NFI.EndDeposito, NotasFiscais.EndEmpresa_Id) AS EndDeposito_Id, YEAR(ISNULL(NotasFiscais.DataParaCusto,"
        Sql &= "                                      NotasFiscais.Movimento)) AS Ano_Id, MONTH(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)) AS Mes_Id, NFI.Produto_Id,"
        Sql &= "                                      SubOperacoes.ApuracaoDeCustos AS Placus_Id,"
        Sql &= "                                      CASE WHEN SubOperacoes.ApuracaodeCustosContraPartida <> 0 THEN IIF(Notasfiscais.Finalidade = 7, NotasFiscais.Destino, NotasFiscais.Cliente_Id)        ELSE '' END AS DepositoDestino_Id,"
        Sql &= "                                      CASE WHEN SubOperacoes.ApuracaodeCustosContraPartida <> 0 THEN IIF(Notasfiscais.Finalidade = 7, NotasFiscais.Enddestino, NotasFiscais.EndCliente_Id)  ELSE 0  END AS EndDepositoDestino_Id,"
        Sql &= "                                      SubOperacoes.ApuracaodeCustosContraPartida, SUM(IIF(Produtos.PesoQuantidade = 'Q', NFI.QuantidadeFiscal, NFI.PesoFiscal)) AS Quantidade, SUM(NFI.Valor) AS ValorDoProduto,"
        Sql &= "                                      NotasFiscais.EntradaSaida_Id, SUM(Enc.ValorIPI) AS ValorIPI, SUM(Enc.ValorDespAcessoria) AS  ValorDespAcessoria, SUM(Enc.ValorICMS) AS ValorICMS, SUM(Enc.ValorFretes) AS ValorFretes, SUM(Enc.ValorDescontos) AS ValorDescontos, SUM(Enc.ValorCOFINS) AS ValorCOFINS,"
        Sql &= "                                      SUM(Enc.ValorPIS) AS ValorPIS, SUM(Enc.DespesasAduaneiras) AS DespesasAduaneiras, "
        Sql &= "                                      SubOperacoes.Devolucao "
        Sql &= "                FROM          GruposDeEstoques LEFT OUTER JOIN "
        Sql &= "                                      Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf
        Sql &= "                RIGHT OUTER JOIN NotasFiscais " & vbCrLf
        Sql &= "                INNER JOIN (SELECT NFI_CONS.Empresa_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.EndEmpresa_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.Cliente_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.EndCliente_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.EntradaSaida_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.Serie_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.Nota_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.Produto_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.CFOP_Id, " & vbCrLf
        Sql &= "                                    NFI_CONS.Deposito, " & vbCrLf
        Sql &= "                                    NFI_CONS.EndDeposito, " & vbCrLf
        Sql &= "                                    NFI_CONS.DepositoTerceiro, " & vbCrLf
        Sql &= "                                    NFI_CONS.EndDepositoTerceiro, " & vbCrLf
        Sql &= "                                    SUM(NFI_CONS.QuantidadeFiscal) AS QuantidadeFiscal, " & vbCrLf
        Sql &= "                                    SUM(NFI_CONS.PesoFiscal) AS PesoFiscal, " & vbCrLf
        Sql &= "                                    SUM(NFI_CONS.Valor) AS Valor " & vbCrLf
        Sql &= "                            FROM NotasFiscaisXItens NFI_CONS " & vbCrLf
        Sql &= "                            GROUP BY NFI_CONS.Empresa_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.EndEmpresa_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.Cliente_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.EndCliente_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.EntradaSaida_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.Serie_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.Nota_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.Produto_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.CFOP_Id, " & vbCrLf
        Sql &= "                                        NFI_CONS.Deposito, " & vbCrLf
        Sql &= "                                        NFI_CONS.EndDeposito, " & vbCrLf
        Sql &= "                                        NFI_CONS.DepositoTerceiro, " & vbCrLf
        Sql &= "                                        NFI_CONS.EndDepositoTerceiro) AS NFI " & vbCrLf
        Sql &= "    ON NotasFiscais.Empresa_Id              = NFI.Empresa_Id " & vbCrLf
        Sql &= "    AND NotasFiscais.EndEmpresa_Id          = NFI.EndEmpresa_Id " & vbCrLf
        Sql &= "    AND NotasFiscais.Cliente_Id             = NFI.Cliente_Id " & vbCrLf
        Sql &= "    AND NotasFiscais.EndCliente_Id          = NFI.EndCliente_Id " & vbCrLf
        Sql &= "    AND NotasFiscais.EntradaSaida_Id        = NFI.EntradaSaida_Id " & vbCrLf
        Sql &= "    AND NotasFiscais.Serie_Id               = NFI.Serie_Id " & vbCrLf
        Sql &= "    AND NotasFiscais.Nota_Id                = NFI.Nota_Id " & vbCrLf
        Sql &= "    ON Produtos.Produto_Id                  = NFI.Produto_Id " & vbCrLf
        Sql &= " LEFT OUTER JOIN SubOperacoes "
        Sql &= "    ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id "
        Sql &= "    AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id "
        Sql &= " INNER JOIN PlanoDeCustos "
        Sql &= "    ON SubOperacoes.ApuracaoDeCustos = PlanoDeCustos.Codigo_Id "
        Sql &= " LEFT OUTER JOIN"
        Sql &= "     (Select     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, "
        Sql &= "                              ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorIPI,"
        Sql &= "                              ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'DESP. ACESSORIA' THEN NotasFiscaisXEncargos.Valor   ELSE 0 END), 0) AS ValorDespAcessoria,"
        Sql &= "                              ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'ICMS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorICMS,"
        Sql &= "                              ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'FRETES' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorFretes, "
        Sql &= "                              ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorDescontos, "

        If Left(Empresa(0), 8) = "05272759" _
            OrElse Left(Empresa(0), 8) = "05366261" _
            OrElse Left(Empresa(0), 8) = "38198213" _
            OrElse Left(Empresa(0), 8) = "62747840" _
            OrElse Left(Empresa(0), 8) = "40938762" _
            OrElse Left(Empresa(0), 8) = "44005444" _
            OrElse Left(Empresa(0), 8) = "48984539" Then
            Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('COFINS','COFINS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorCOFINS,"
            Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorPIS, "
        Else
            Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('COFINS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorCOFINS,"
            Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('PIS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorPIS, "
        End If

        Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('DESP.ADUANEIRAS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS DespesasAduaneiras "
        Sql &= "                                            FROM NotasFiscaisXEncargos"
        Sql &= "                                            GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id) AS Enc ON"
        Sql &= "                                      NFI.Empresa_Id = Enc.Empresa_Id AND NFI.EndEmpresa_Id = Enc.EndEmpresa_Id AND"
        Sql &= "                                      NFI.Cliente_Id = Enc.Cliente_Id AND NFI.EndCliente_Id = Enc.EndCliente_Id AND"
        Sql &= "                                      NFI.EntradaSaida_Id = Enc.EntradaSaida_Id AND NFI.Serie_Id = Enc.Serie_Id AND"
        Sql &= "                                      NFI.Nota_Id = Enc.Nota_Id And NFI.Produto_Id = Enc.Produto_Id And NFI.CFOP_Id = Enc.CFOP_Id"

        Sql &= " WHERE      (NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        Sql &= "        AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & pMes & ")"
        Sql &= "        AND (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))   = " & DdlAno.SelectedValue & ")"
        Sql &= "        AND (SubOperacoes.ApuracaoDeCustos > 0)"
        Sql &= "        And  NotasFiscais.Situacao = 1 "
        Sql &= "        And  GruposDeEstoques.custo = 1"
        Sql &= "        And PlanoDeCustos.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'"
        'sql &=  Sqla & "; "

        Sql &= " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NFI.Deposito, NFI.EndDeposito, NotasFiscais.Cliente_Id,"
        Sql &= "                                      NotasFiscais.EndCliente_Id, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,"
        Sql &= "                                      NotasFiscais.Movimento)), NFI.Produto_Id, SubOperacoes.ApuracaoDeCustos, NFI.DepositoTerceiro,"
        Sql &= "                                      NFI.EndDepositoTerceiro, SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id, SubOperacoes.Devolucao, Notasfiscais.Finalidade, Notasfiscais.Destino, NotasFiscais.EndDestino)"
        Sql &= "              AS Consulta"
        Sql &= " WHERE (Placus_Id <> 1)"
        Sql &= " GROUP BY Empresa_Id, EndEmpresa_Id, Placus_Id, ApuracaodeCustosContraPartida, Produto_Id, Deposito_Id, EndDeposito_Id, DepositoDestino_Id, EndDepositoDestino_Id"

        '11 - CAPTURANDO DADOS DAS NOTAS FISCAIS - ATUALIZANDO TABELA DE APURAÇÃO
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds2.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "       when exists ("
                Sql &= "                 select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino from ApuracaoDeCustos "
                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                Sql &= " And Deposito_Id = '" & rs("Deposito_Id") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDeposito_Id")
                Sql &= " And Ano_Id = " & DdlAno.SelectedValue
                Sql &= " And Mes_Id = " & pMes
                Sql &= " And Produto_Id = '" & rs("Produto_Id") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("Placus_Id")
                Sql &= " And EmpresaDestino_Id = ''"
                Sql &= " And EndEmpresaDestino_Id = 0"

                If rs("ApuracaodeCustosContrapartida") <> 0 Then
                    Sql &= "  And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= "  And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                    Sql &= "  And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                Else
                    Sql &= "  And DepositoDestino_Id = ''"  'Deposito Destino
                    Sql &= "  And EndDepositoDestino_Id = 0"   'End DepositoDestino
                    Sql &= "  And ProdutoDerivado_Id = ''"  'Produto Derivado
                End If

                Sql &= ")"
                Sql &= "            then 'S'"
                Sql &= "             else 'N'"
                Sql &= "               end) ;"

                Sql &= " if @Exist = 'N' "
                Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"
                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino)"

                Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito_Id") & "'"
                Sql &= ", " & rs("EndDeposito_Id")
                Sql &= ", " & DdlAno.SelectedValue
                Sql &= ", " & pMes
                Sql &= ", '" & rs("Produto_Id") & "'"
                Sql &= ", " & rs("Placus_Id")

                If rs("Placus_Id") = 355 Or rs("Placus_Id") = 751 Or rs("Placus_Id") = 750 Then
                    Sql &= ", '" & rs("Empresa_Id") & "'"
                    Sql &= ", " & rs("EndEmpresa_Id")
                Else
                    Sql &= ", ''"  'Empresa Destino
                    Sql &= ", 0"   'End Empresa Destino
                End If


                If rs("ApuracaodeCustosContrapartida") <> 0 Then
                    Sql &= ",  '" & rs("DepositoDestino_Id") & "'"
                    Sql &= ", " & rs("EndDeposito_Id")
                    Sql &= ", '" & rs("Produto_Id") & "'"
                Else
                    Sql &= ", ''"  'Deposito Destino
                    Sql &= ", 0"   'End DepositoDestino
                    Sql &= ", ''"  'Produto Derivado
                End If

                Sql &= ", " & Str(rs("Quantidade"))

                If rs("Placus_Id") = 803 Then
                    Sql &= ", " & Str(rs("ValorDoProduto"))
                    'Isolado em 12/01/2021 por estar descontando o ICMS duas vezes.
                    'Sql &= ", " & Str(rs("ValorDoProduto") - rs("ICMS"))
                Else
                    If rs("Placus_Id") = 355 Or rs("Placus_Id") = 751 Then
                        Sql &= ", 0"
                    Else
                        Sql &= ", " & Str(rs("ValorDoProduto"))
                    End If
                End If

                Sql &= ", 0"  '& rs("ValorDoFrete").ToString.Replace(",", ".")
                Sql &= ", 0"  '& rs("ValorAuxiliar").ToString.Replace(",", ".")

                If rs("ApuracaodeCustosContrapartida") <> 0 Then
                    Sql &= ", '" & rs("Produto_Id") & "'"
                    Sql &= ", '" & rs("ApuracaodeCustosContrapartida") & "'"
                Else
                    Sql &= ", ''"  '& rs("ProdutoDestino") & "'"
                    Sql &= ", 0"  '& rs("CodigoDestino")
                End If

                Sql &= ")"
                Sql &= " Else"

                Sql &= "   Update ApuracaoDeCustos set "
                Sql &= "   Quantidade = " & Str(rs("Quantidade"))

                If rs("Placus_Id") = 803 Then
                    Sql &= ",  ValorDoProduto = " & Str(rs("ValorDoProduto") - rs("ICMS"))
                Else
                    Sql &= ",  ValorDoProduto = " & Str(rs("ValorDoProduto"))
                End If

                Sql &= " Where Empresa_Id       = '" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                Sql &= "   And Deposito_Id      = '" & rs("Deposito_Id") & "'"
                Sql &= "   And EndDeposito_Id   = " & rs("EndDeposito_Id")
                Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
                Sql &= "   And Mes_Id           = " & pMes
                Sql &= "   And Produto_Id       = '" & rs("Produto_Id") & "'"
                Sql &= "   And CodigoDeCusto_Id = " & rs("Placus_Id")
                Sql &= "   And EmpresaDestino_Id = ''"
                Sql &= "   And EndEmpresaDestino_Id = 0"

                If rs("ApuracaodeCustosContrapartida") <> 0 Then
                    Sql &= "  And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= "  And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                    Sql &= "  And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                Else
                    Sql &= "  And DepositoDestino_Id = ''"  'Deposito Destino
                    Sql &= "  And EndDepositoDestino_Id = 0"   'End DepositoDestino
                    Sql &= "  And ProdutoDerivado_Id = ''"  'Produto Derivado
                End If

                Sqll = Sqll & Sql
            Next
        End If
        aux = Banco.GravaBanco(Sqll)
        Return aux
    End Function

    Private Function CapturaNotasFiscaisContraPartida(ByVal pMes As Integer) As Boolean
        Dim aux As Boolean = True
        Dim i = 0
        '10 - Capturando dados das Notas Fiscais Contra Partida- Criando Filtro
        'txt_Processo.value = "Processando dados das notas fiscais Contrapartida ..."

        Dim Sql As String = " SELECT     Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, DepositoDestino_Id, EndDepositoDestino_Id, Placus_Id as CodigoDestino, PlacusContra_Id, Produto_Id, SUM(Quantidade) AS Quantidade,"
        Sql &= "              0 AS ValorDoProduto"
        Sql &= " FROM         (SELECT     NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id AS Deposito_Id, NotasFiscais.EndCliente_Id AS EndDeposito_Id,"
        Sql &= "                          YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)) AS Ano_Id, MONTH(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento))"
        Sql &= "                          AS Mes_Id, NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaodeCustos AS Placus_Id, SubOperacoes.ApuracaodeCustosContraPartida AS PlacusContra_Id,"
        Sql &= "                          NotasFiscaisXItens.Deposito AS DepositoDestino_Id, NotasFiscaisXItens.EndDeposito AS EndDepositoDestino_Id, SubOperacoes.ApuracaoDeCustos,"
        Sql &= "                          SUM(IIF(Produtos.PesoQuantidade = 'Q', NotasFiscaisXItens.QuantidadeFiscal, NotasFiscaisXItens.PesoFiscal)) AS Quantidade, SUM(NotasFiscaisXItens.Valor) AS ValorDoProduto, NotasFiscais.EntradaSaida_Id, SUM(Enc.ValorIPI)"
        Sql &= "                          AS ValorIPI, SUM(Enc.ValorICMS) AS ValorICMS, SUM(Enc.ValorCOFINS) AS ValorCOFINS, SUM(Enc.ValorPIS) AS ValorPIS"
        Sql &= "             FROM         NotasFiscais INNER JOIN"
        Sql &= "                          NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND"
        Sql &= "                          NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND"
        Sql &= "                          NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND"
        Sql &= "                          NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN"
        Sql &= "                          SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN"
        Sql &= "                          Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id INNER JOIN"
        Sql &= "                          GruposDeEstoques ON Produtos.Grupo = GruposDeEstoques.Grupo_Id INNER JOIN"
        Sql &= "                          PlanoDeCustos ON SubOperacoes.ApuracaoDeCustos = PlanoDeCustos.Codigo_Id LEFT OUTER JOIN"
        Sql &= "                          (SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id,"
        Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorIPI,"
        Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'ICMS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorICMS,"

        If Left(Empresa(0), 8) = "05272759" _
            OrElse Left(Empresa(0), 8) = "05366261" _
            OrElse Left(Empresa(0), 8) = "38198213" _
            OrElse Left(Empresa(0), 8) = "62747840" _
            OrElse Left(Empresa(0), 8) = "40938762" _
            OrElse Left(Empresa(0), 8) = "44005444" _
            OrElse Left(Empresa(0), 8) = "48984539" Then
            Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('COFINS','COFINS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorCOFINS,"
            Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorPIS"
        Else
            Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('COFINS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorCOFINS,"
            Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('PIS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorPIS"
        End If

        Sql &= "                          FROM NotasFiscaisXEncargos"
        Sql &= "                          GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id) AS Enc ON"
        Sql &= "                          NotasFiscaisXItens.Empresa_Id = Enc.Empresa_Id AND NotasFiscaisXItens.EndEmpresa_Id = Enc.EndEmpresa_Id AND"
        Sql &= "                          NotasFiscaisXItens.Cliente_Id = Enc.Cliente_Id AND NotasFiscaisXItens.EndCliente_Id = Enc.EndCliente_Id AND"
        Sql &= "                          NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = Enc.Serie_Id AND"
        Sql &= "                          NotasFiscaisXItens.Nota_Id = Enc.Nota_Id And NotasFiscaisXItens.Produto_Id = Enc.Produto_Id And NotasFiscaisXItens.CFOP_Id = Enc.CFOP_Id"

        Sql &= " WHERE      (NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        Sql &= "        AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & pMes & ")"
        Sql &= "        AND (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))   = " & DdlAno.SelectedValue & ")"
        Sql &= "        AND (SubOperacoes.ApuracaoDeCustosContraPartida > 0)"
        Sql &= "        And  NotasFiscais.Situacao = 1 "
        Sql &= "        And  GruposDeEstoques.custo = 1"
        Sql &= "        And  PlanoDeCustos.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'"

        Sql &= " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito,"
        Sql &= "          NotasFiscaisXItens.EndDeposito, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,"
        Sql &= "          NotasFiscais.Movimento)), NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro,"
        Sql &= "          NotasFiscaisXItens.EndDepositoTerceiro, SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id) AS Consulta"
        Sql &= " WHERE (Placus_Id > 0)"
        Sql &= " GROUP BY Empresa_Id, EndEmpresa_Id, Placus_Id, PlacusContra_Id, Produto_Id, Deposito_Id, EndDeposito_Id, Produto_Id, DepositoDestino_Id, EndDepositoDestino_Id"

        '11 - Capturando dados das Notas Fiscais ContraPartida- Atualizando Tabela de Apuração
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "       when exists ("
                Sql &= "                 select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino from ApuracaoDeCustos "
                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                Sql &= " And Deposito_Id = '" & rs("Deposito_Id") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDeposito_Id")
                Sql &= " And Ano_Id = " & DdlAno.SelectedValue
                Sql &= " And Mes_Id = " & pMes
                Sql &= " And Produto_Id = '" & rs("Produto_Id") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("PlacusContra_Id")
                Sql &= " And EmpresaDestino_Id = ''"
                Sql &= " And EndEmpresaDestino_Id = 0"

                If rs("PlacusContra_Id") <> 0 Then
                    Sql &= "  And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= "  And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                    Sql &= "  And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                Else
                    Sql &= "  And DepositoDestino_Id = ''"  'Deposito Destino
                    Sql &= "  And EndDepositoDestino_Id = 0"   'End DepositoDestino
                    Sql &= "  And ProdutoDerivado_Id = ''"  'Produto Derivado
                End If

                Sql &= ")"
                Sql &= "            then 'S'"
                Sql &= "             else 'N'"
                Sql &= "               end) ;"

                Sql &= " if @Exist = 'N' "
                Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"

                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino)"

                Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito_Id") & "'"
                Sql &= ", " & rs("EndDeposito_Id")
                Sql &= ", " & DdlAno.SelectedValue
                Sql &= ", " & pMes
                Sql &= ", '" & rs("Produto_Id") & "'"
                Sql &= ", " & rs("PlacusContra_Id")
                Sql &= ", ''" 'Empresa Destino
                Sql &= ", 0"  'End Empresa Destino
                Sql &= ", '" & rs("DepositoDestino_Id") & "'"
                Sql &= ", " & rs("EndDepositoDestino_Id")
                Sql &= ", '" & rs("Produto_Id") & "'"

                Sql &= ", " & Str(rs("Quantidade"))
                Sql &= ", 0" '& Str(rs("ValorDoProduto"))
                Sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                Sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                Sql &= ", ''" '& rs("Produto_Id") & "'"
                Sql &= ", " & rs("CodigoDestino")
                Sql &= ")"

                Sql &= " Else"

                Sql &= "   Update ApuracaoDeCustos set "
                Sql &= "   Quantidade = " & Str(rs("Quantidade"))
                'sql &=  ",  ValorDoProduto = " & Str(rs("ValorDoProduto"))

                Sql &= " Where Empresa_Id       = '" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                Sql &= "   And Deposito_Id      = '" & rs("Deposito_Id") & "'"
                Sql &= "   And EndDeposito_Id   = " & rs("EndDeposito_Id")
                Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
                Sql &= "   And Mes_Id           = " & pMes
                Sql &= "   And Produto_Id       = '" & rs("Produto_Id") & "'"
                Sql &= "   And CodigoDeCusto_Id = " & rs("PlacusContra_Id")
                Sql &= "   And EmpresaDestino_Id = ''"
                Sql &= "   And EndEmpresaDestino_Id = 0"
                If rs("PlacusContra_Id") <> 0 Then
                    Sql &= "  And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= "  And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                    Sql &= "  And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                Else
                    Sql &= "  And DepositoDestino_Id = ''"  'Deposito Destino
                    Sql &= "  And EndDepositoDestino_Id = 0"   'End DepositoDestino
                    Sql &= "  And ProdutoDerivado_Id = ''"  'Produto Derivado
                End If

                Sqll = Sqll & Sql
            Next
        End If

        aux = Banco.GravaBanco(Sqll)
        Return aux
    End Function

    Private Function CapturaSaidasPorTransferencias(ByVal pMes As Integer) As Boolean
        Dim i = 0
        '10 - Capturando dados das Notas Fiscais Contra Partida - Criando Filtro
        'txt_Processo.value = "Processando dados das notas fiscais Contrapartida ..."

        Dim Sql As String = " SELECT          NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito AS Deposito_Id, NotasFiscaisXItens.EndDeposito AS EndDeposito_Id,"
        Sql &= "           NotasFiscais.Cliente_Id AS EmpresaDestino_Id, NotasFiscais.EndCliente_Id AS EndEmpresaDestino_Id, NotasFiscais.Cliente_Id AS DepositoDestino_Id,"
        Sql &= "           NotasFiscais.EndCliente_Id AS EndDepositoDestino_Id, NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaoDeCustos AS Placus_Id,"
        Sql &= "           SubOperacoes.ApuracaodeCustosContraPartida AS PlacusContra_Id, SUM(IIF(Produtos.PesoQuantidade = 'Q', NotasFiscaisXItens.QuantidadeFiscal, NotasFiscaisXItens.PesoFiscal)) AS Quantidade, 0 AS ValorDoProduto"
        Sql &= "  FROM     PlanoDeCustos INNER JOIN"
        Sql &= "           SubOperacoes ON PlanoDeCustos.Codigo_Id = SubOperacoes.ApuracaoDeCustos INNER JOIN"
        Sql &= "           NotasFiscais INNER JOIN"
        Sql &= "           NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND"
        Sql &= "           NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND"
        Sql &= "           NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND"
        Sql &= "           NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id ON SubOperacoes.Operacao_Id = NotasFiscaisXItens.Operacao AND"
        Sql &= "           SubOperacoes.SubOperacoes_Id = NotasFiscaisXItens.SubOperacao"

        Sql &= " WHERE      (NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        Sql &= "        AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & pMes & ")"
        Sql &= "        AND (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))   = " & DdlAno.SelectedValue & ")"
        Sql &= "        AND (SubOperacoes.ApuracaoDeCustosContraPartida > 0)"
        Sql &= "        And  NotasFiscais.Situacao = 1"
        Sql &= "        And  PlanoDeCustos.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'"

        Sql &= " GROUP BY  NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito,"
        Sql &= "           NotasFiscaisXItens.EndDeposito, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,"
        Sql &= "           NotasFiscais.Movimento)), NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro,"
        Sql &= "           NotasFiscaisXItens.EndDepositoTerceiro , SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id"

        '11 - Capturando dados das Notas Fiscais ContraPartida - Atualizando Tabela de Apuração
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "       when exists ("
                Sql &= "                 select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino from ApuracaoDeCustos "
                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                Sql &= " And Deposito_Id = '" & rs("Deposito_Id") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDeposito_Id")
                Sql &= " And Ano_Id = " & DdlAno.SelectedValue
                Sql &= " And Mes_Id = " & pMes
                Sql &= " And Produto_Id = '" & rs("Produto_Id") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("Placus_Id")

                If Left(Empresa(0), 8) = "05272759" _
                    OrElse Left(Empresa(0), 8) = "05366261" _
                    OrElse Left(Empresa(0), 8) = "38198213" _
                    OrElse Left(Empresa(0), 8) = "62747840" _
                    OrElse Left(Empresa(0), 8) = "40938762" _
                    OrElse Left(Empresa(0), 8) = "44005444" _
                    OrElse Left(Empresa(0), 8) = "48984539" Then
                    Sql &= " And EmpresaDestino_Id = '" & rs("EmpresaDestino_Id") & "'"
                    Sql &= " And EndEmpresaDestino_Id = " & rs("EndEmpresaDestino_Id")
                    Sql &= " And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= " And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                    Sql &= " And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                Else
                    Sql &= " And EmpresaDestino_Id = '" & rs("EmpresaDestino_Id") & "'"
                    Sql &= " And EndEmpresaDestino_Id = " & rs("EndEmpresaDestino_Id")
                    Sql &= " And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= " And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                End If
                'Gerencial tem essa opção sem o Produto Derivado - Furlan 04/05/2020

                Sql &= ")"
                Sql &= "            then 'S'"
                Sql &= "             else 'N'"
                Sql &= "               end) ;"

                Sql &= " if @Exist = 'N' "
                Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"

                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino)"

                Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito_Id") & "'"
                Sql &= ", " & rs("EndDeposito_Id")
                Sql &= ", " & DdlAno.SelectedValue
                Sql &= ", " & pMes
                Sql &= ", '" & rs("Produto_Id") & "'"
                Sql &= ", " & rs("Placus_Id")
                Sql &= ", '" & rs("EmpresaDestino_Id") & "'"
                Sql &= ", " & rs("EndEmpresaDestino_Id")
                Sql &= ", '" & rs("DepositoDestino_Id") & "'"
                Sql &= ", " & rs("EndDepositoDestino_Id")
                Sql &= ", '" & rs("Produto_Id") & "'"

                Sql &= ", " & Str(rs("Quantidade"))
                Sql &= ", 0" '& Str(rs("ValorDoProduto"))
                Sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                Sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                Sql &= ", ''" '& rs("Produto_Id") & "'"
                Sql &= ", " & rs("PlacusContra_Id")
                Sql &= ")"

                Sql &= " Else"

                Sql &= "   Update ApuracaoDeCustos set "
                Sql &= "   Quantidade = " & Str(rs("Quantidade"))
                'sql &=  ",  ValorDoProduto = " & Str(rs("ValorDoProduto"))

                Sql &= " Where Empresa_Id       = '" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                Sql &= "   And Deposito_Id      = '" & rs("Deposito_Id") & "'"
                Sql &= "   And EndDeposito_Id   = " & rs("EndDeposito_Id")
                Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
                Sql &= "   And Mes_Id = " & pMes
                Sql &= "   And Produto_Id = '" & rs("Produto_Id") & "'"
                Sql &= "   And CodigoDeCusto_Id = " & rs("Placus_Id")

                If Left(Empresa(0), 8) = "05272759" _
                    OrElse Left(Empresa(0), 8) = "05366261" _
                    OrElse Left(Empresa(0), 8) = "38198213" _
                    OrElse Left(Empresa(0), 8) = "62747840" _
                    OrElse Left(Empresa(0), 8) = "40938762" _
                    OrElse Left(Empresa(0), 8) = "44005444" _
                    OrElse Left(Empresa(0), 8) = "48984539" Then
                    Sql &= "   And EmpresaDestino_Id = '" & rs("EmpresaDestino_Id") & "'"
                    Sql &= "   And EndEmpresaDestino_Id = " & rs("EndEmpresaDestino_Id")
                    Sql &= "   And DepositoDestino_Id = '" & rs("DepositoDestino_Id") & "'"
                    Sql &= "   And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                    Sql &= "   And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                End If
                'Gerencial tem essa opção sem o Produto Derivado - Furlan - 04/05/2020

                Sqll = Sqll & Sql
            Next
        End If

        Return Banco.GravaBanco(Sqll)

    End Function

    Private Function AjustaTotalizadores(ByVal pMes As Integer) As Boolean
        Dim Fase As Integer
        '--------------" 13 - Ajustando Totalizadores"))
        'txt_Processo.value = "Ajustando totalizadores..."

        Dim Sql As String = " Update ApuracaoDeCustos SET"
        Sql &= "     Quantidade     = 0"
        Sql &= "    ,ValorDoProduto = 0"
        Sql &= "    ,ValorDoFrete   = 0"
        Sql &= "    ,ValorAuxiliar  = 0"
        Sql &= "  From PlanoDeCustos PC"
        Sql &= "  Inner join ApuracaoDeCustos"
        Sql &= "    ON ApuracaoDeCustos.CodigoDeCusto_Id = PC.Codigo_Id"
        Sql &= "  WHERE ApuracaoDeCustos.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   AND ApuracaoDeCustos.Ano_Id = " & DdlAno.SelectedValue
        Sql &= "   AND ApuracaoDeCustos.Mes_Id = " & pMes
        Sql &= "   AND PC.FaseDoTotalizador > 0" & vbCrLf

        If Not Banco.GravaBanco(Sql) Then Return False

        For Fase = 0 To 5
            Sql = " SELECT AP.Empresa_Id, AP.EndEmpresa_Id, AP.Deposito_Id, AP.EndDeposito_Id, AP.Ano_Id,"
            Sql &= "       AP.Mes_Id, AP.Produto_Id, AP.CodigoDeCusto_Id, AP.EmpresaDestino_Id,"
            Sql &= "       AP.EndEmpresaDestino_Id, AP.DepositoDestino_Id, AP.EndDepositoDestino_Id,"
            Sql &= "       AP.ProdutoDerivado_Id, AP.Etapa, AP.Quantidade, AP.ValorDoProduto,"
            Sql &= "       AP.ValorDoFrete, isnull(AP.ValorAuxiliar, 0) as ValorAuxiliar, AP.ProdutoDestino, AP.CodigoDestino,"
            Sql &= "       AP.Reduzido, PC.Totalizador,"
            Sql &= "       ISNULL(PC.SinalPeso, '') AS SinalPeso,"
            Sql &= "       ISNULL(PC.SinalValor, '') AS SinalValor"
            Sql &= "  FROM ApuracaoDeCustos AP"
            Sql &= "  INNER JOIN PlanoDeCustos PC"
            Sql &= "  ON AP.CodigoDeCusto_Id = PC.Codigo_Id"
            Sql &= "  WHERE (AP.Empresa_Id      LIKE '" & Left(Empresa(0), 8) & "%')"
            Sql &= "  AND (AP.Ano_Id          = " & DdlAno.SelectedValue & ")"
            Sql &= "  AND (AP.Mes_Id          = " & pMes & ")"
            Sql &= "  AND (PC.Totalizador        > 0)"
            Sql &= "  AND (AP.Quantidade + AP.ValorDoProduto + AP.ValorDoFrete <> 0)"
            Sql &= "  AND (PC.FaseDoTotalizador  = " & Fase & ") "
            Sql &= "  ORDER BY AP.Empresa_Id, AP.EndEmpresa_Id, AP.Deposito_Id, AP.EndDeposito_Id,"
            Sql &= "          AP.Ano_Id, AP.Mes_Id, AP.Produto_Id, AP.CodigoDeCusto_Id, AP.DepositoDestino_Id,"
            Sql &= "          AP.EndDepositoDestino_Id" & vbCrLf

            Sqll = " Declare"
            Sqll = Sqll & " @Exist as varchar(1)"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each rs As DataRow In ds.Tables(0).Rows

                    Sql = " set @Exist = (select case "
                    Sql &= "                        when exists ("
                    Sql &= "                                      select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino "
                    Sql &= "                                        from ApuracaoDeCustos "
                    Sql &= "                                       Where Empresa_Id       = '" & rs("Empresa_Id") & "'"
                    Sql &= "                                         And EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                    Sql &= "                                        And Deposito_Id      = '" & rs("Deposito_Id") & "'"
                    Sql &= "                                         And EndDeposito_Id   = " & rs("EndDeposito_Id")
                    Sql &= "                                         And Ano_Id           = " & DdlAno.SelectedValue
                    Sql &= "                                         And Mes_Id           = " & pMes
                    Sql &= "                                         And Produto_Id       = '" & rs("Produto_Id") & "'"
                    Sql &= "                                       And CodigoDeCusto_Id = " & rs("Totalizador")
                    Sql &= "                                    )"
                    Sql &= "                            then 'S'"
                    Sql &= "                            else 'N'"
                    Sql &= "                      end) ;"
                    Sql &= " if @Exist = 'N' "
                    Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                    Sql &= "                                Empresa_Id"
                    Sql &= "                               ,EndEmpresa_Id"
                    Sql &= "                               ,Deposito_Id"
                    Sql &= "                               ,EndDeposito_Id"
                    Sql &= "                               ,Ano_Id"
                    Sql &= "                               ,Mes_Id"
                    Sql &= "                               ,Produto_Id"
                    Sql &= "                               ,CodigoDeCusto_Id"
                    Sql &= "                              ,EmpresaDestino_Id"
                    Sql &= "                               ,EndEmpresaDestino_Id"
                    Sql &= "                              ,DepositoDestino_Id"
                    Sql &= "                              ,EndDepositoDestino_Id"
                    Sql &= "                               ,ProdutoDerivado_Id"
                    Sql &= "                               ,Quantidade"
                    Sql &= "                               ,ValorDoProduto"
                    Sql &= "                               ,ValorDoFrete"
                    Sql &= "                               ,ValorAuxiliar"
                    Sql &= "                              ,ProdutoDestino"
                    Sql &= "                               ,CodigoDestino)"
                    Sql &= " VALUES( '" & rs("Empresa_Id") & "'"
                    Sql &= "        , " & rs("EndEmpresa_Id")
                    Sql &= "        ,'" & rs("Deposito_Id") & "'"
                    Sql &= "       , " & rs("EndDeposito_Id")
                    Sql &= "        , " & DdlAno.SelectedValue
                    Sql &= "       , " & pMes
                    Sql &= "       ,'" & rs("Produto_Id") & "'"
                    Sql &= "        , " & rs("Totalizador")
                    Sql &= "        , ''"
                    Sql &= "        , 0"
                    Sql &= "        , ''"
                    Sql &= "        , 0"
                    Sql &= "        , ''"
                    Sql &= "       , 0"
                    Sql &= "        , 0"
                    Sql &= "        , 0"
                    Sql &= "       , 0"
                    Sql &= "        , '" & rs("ProdutoDestino") & "'"
                    Sql &= "       , " & rs("CodigoDestino")
                    Sql &= "       )"

                    If rs("SinalPeso") <> "" Or rs("SinalValor") <> "" Then
                        Sql &= "  Update ApuracaoDeCustos set "

                        If rs("SinalPeso") <> "" Then
                            Sql &= "  Quantidade = Quantidade " & rs("SinalPeso") & "(" & Str(rs("Quantidade")) & ")"
                        End If

                        If rs("SinalValor") <> "" Then
                            If rs("SinalPeso") <> "" Then Sql &= ", "

                            Sql &= " ValorDoProduto = ValorDoProduto " & rs("SinalValor") & " (" & Str(rs("ValorDoProduto")) & ")"
                            Sql &= ", ValorDoFrete  = ValorDoFrete   " & rs("SinalValor") & " (" & Str(rs("ValorDoFrete")) & ")"
                            Sql &= ", ValorAuxiliar = ValorAuxiliar  " & rs("SinalValor") & " (" & Str(rs("ValorAuxiliar")) & ")"
                        End If

                        Sql &= " Where Empresa_Id       ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                        Sql &= "   And Deposito_Id      ='" & rs("Deposito_Id") & "'"
                        Sql &= "   And EndDeposito_Id   = " & rs("EndDeposito_Id")
                        Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
                        Sql &= "   And Mes_Id           = " & pMes
                        Sql &= "   And Produto_Id       ='" & rs("Produto_Id") & "'"
                        Sql &= "   And CodigoDeCusto_Id = " & rs("Totalizador")
                    End If

                    Sqll = Sqll & Sql & vbCrLf

                Next

            End If

            If Not Banco.GravaBanco(Sqll) Then Return False
        Next

        Return True

    End Function

    Private Function AjustaCustosDeSaidas(ByVal pMes As Integer) As Boolean
        '17 - AJUSTANDO CUSTOS DE SAIDAS

        Dim Sql As String = " SELECT Empresa_Id, EndEmpresa_ID, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, "
        Sql &= "  CodigoDeCusto_Id, DepositoDestino_ID, EndDepositoDestino_Id, ProdutoDerivado_Id, "
        Sql &= "  Convert(Decimal(18,9),((ValorDoProduto + ValorDoFrete) / Quantidade)) as Medio,  Case when Empresa_Id = Deposito_Id then 1 else 2 end as Ordem  "
        Sql &= "  FROM ApuracaoDeCustos"
        Sql &= "  Where Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
        Sql &= "   And Mes_ID           = " & pMes
        Sql &= "   And CodigoDeCusto_Id = 495"
        Sql &= "   And Quantidade       > 0"
        'sql &= "   And Quantidade       > 0" No Custo Planilha Curtume estava Quantidade       <> 0 - Furlan - 11/11/2019

        Sql &= "   Order by Empresa_Id, EndEmpresa_ID, Case when Empresa_Id = Deposito_Id then 1 else 2 end, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, "
        Sql &= "   CodigoDeCusto_Id, DepositoDestino_ID, EndDepositoDestino_Id, ProdutoDerivado_Id"
        Sqll = ""

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = "  Update  ApuracaoDeCustos set"
                Sql &= "    ApuracaoDeCustos.ValorDoProduto = ApuracaoDeCustos.Quantidade * " & Replace(rs("Medio"), ",", ".")
                Sql &= " FROM  ApuracaoDeCustos INNER JOIN"
                Sql &= " PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
                Sql &= "    Where ApuracaoDeCustos.Empresa_Id       ='" & rs("Empresa_Id") & "'"
                Sql &= "    And ApuracaoDeCustos.EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                Sql &= "    And ApuracaoDeCustos.Deposito_Id      ='" & rs("Deposito_Id") & "'"
                Sql &= "    And ApuracaoDeCustos.EndDeposito_Id   = " & rs("EndDeposito_Id")
                Sql &= "    And ApuracaoDeCustos.Ano_Id           = " & rs("Ano_Id")
                Sql &= "    And ApuracaoDeCustos.Mes_Id           = " & rs("Mes_Id")
                Sql &= "    And ApuracaoDeCustos.Produto_Id       ='" & rs("Produto_Id") & "'"
                Sql &= "    And ApuracaoDeCustos.CodigoDeCusto_Id > 495 "
                Sql &= "    And Not (PlanoDeCustos.Classe In ('" & eClassesOperacoes.MUTUO.ToString & "'))"

                Sqll = Sqll & Sql

            Next
            Return Banco.GravaBanco(Sqll)
        Else
            Return False
        End If
    End Function

    Private Function AjustaMovimento(ByVal pMes As Integer) As Boolean
        '----------" 17 - Ajustando Custos sem Movimento "))
        'txt_Processo.value = "Ajustando Custos sem movimento. .."

        Dim Sql As String = "  SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino* FROM ApuracaoDeCustos"
        Sql &= "  Where ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   And ApuracaoDeCustos.Ano_Id           = " & DdlAno.SelectedValue
        Sql &= "   And ApuracaoDeCustos.Mes_ID           = " & pMes
        Sql &= "And CodigoDeCusto_Id in (402, 404)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                SqlAux = "        SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino "
                SqlAux = SqlAux & " FROM ApuracaoDeCustos"
                SqlAux = SqlAux & " WHERE Empresa_ID = '" & rs("Empresa_Id") & "'"
                SqlAux = SqlAux & " And Deposito_Id =  '" & rs("Deposito_Id") & "'"
                SqlAux = SqlAux & " And EndDeposito_ID = " & rs("EndDeposito_Id")
                SqlAux = SqlAux & " AND (Ano_Id = " & DdlAno.SelectedValue & ")"
                SqlAux = SqlAux & " AND (Mes_Id = " & pMes & ")"
                SqlAux = SqlAux & " And Produto_Id = '" & rs("Produto_Id") & "'"
                SqlAux = SqlAux & " AND (Quantidade > 0)"

                Dim i = 0
                Dim ds2 As DataSet = Banco.ConsultaDataSet(SqlAux, "ApuracaoDeCustos")
                If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
                    For Each rsa As DataRow In ds2.Tables(0).Rows
                        i = i + 1
                    Next
                End If

                If i = 0 Then
                    Sql = "  Delete ApuracaodeCustos"
                    Sql &= " WHERE Empresa_ID = '" & rs("Empresa_Id") & "'"
                    Sql &= " And Deposito_Id = '" & rs("Deposito_Id") & "'"
                    Sql &= " And EndDeposito_ID = " & rs("EndDeposito_Id")
                    Sql &= " AND (Ano_Id = " & DdlAno.SelectedValue & ")"
                    Sql &= " AND (Mes_Id = " & pMes & ")"
                    Sql &= " And Produto_Id = '" & rs("Produto_Id") & "'"
                    If Not Banco.GravaBanco(Sql) Then Return False
                End If
            Next
        End If

        Return True
    End Function

    Private Function AjustaMovimentoDevolucaoDeDeposito(ByVal pMes As Integer) As Boolean
        'Ajustando Custos sem do código 355 com quantidade zerada "))

        Dim Sql As String = "  Delete ApuracaodeCustos"
        Sql &= " WHERE Left(Empresa_Id, 8) in ('05272759','05366261','38198213','40938762','44005444','48984539','62747840')"
        Sql &= " AND (Ano_Id = " & DdlAno.SelectedValue & ")"
        Sql &= " AND (Mes_Id = " & pMes & ")"
        Sql &= "And CodigoDeCusto_Id = (355)  AND (Quantidade = 0)"
        If Not Banco.GravaBanco(Sql) Then Return False

        Return True
    End Function

    Private Function AjustaMutuos(ByVal pMes As Integer) As Boolean
        Dim Medio As Double
        Dim Valor As Double

        '17 - Ajustando Custos de Saidas
        'txt_Processo.value = "Ajustando Mutuos de Saida. .."

        Dim Sql As String = "DELETE  ApuracaoDeCustos"
        Sql &= " FROM    ApuracaoDeCustos INNER JOIN"
        Sql &= "         PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
        Sql &= "  Where ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   And ApuracaoDeCustos.Ano_Id           = " & DdlAno.SelectedValue
        Sql &= "   And ApuracaoDeCustos.Mes_ID           = " & pMes
        Sql &= "   And ApuracaoDeCustos.CodigoDeCusto_Id > 495"
        Sql &= "   And ApuracaoDeCustos.Quantidade       = 0"
        Sql &= "   AND (PlanoDeCustos.Classe = '" & eClassesOperacoes.MUTUO.ToString & "')"
        Banco.GravaBanco(Sql)

        Sql = "       SELECT  ApuracaoDeCustos.*"
        Sql &= " FROM    ApuracaoDeCustos INNER JOIN"
        Sql &= "         PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
        Sql &= "  Where ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   And ApuracaoDeCustos.Ano_Id           = " & DdlAno.SelectedValue
        Sql &= "   And ApuracaoDeCustos.Mes_ID           = " & pMes
        Sql &= "   And ApuracaoDeCustos.CodigoDeCusto_Id > 495"
        Sql &= "   And ApuracaoDeCustos.Quantidade       > 0"
        Sql &= "   AND (PlanoDeCustos.Classe = '" & eClassesOperacoes.MUTUO.ToString & "')"
        Sqll = ""

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"

                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino)"

                Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito_Id") & "'"
                Sql &= ", " & rs("EndDeposito_Id")
                Sql &= ", " & DdlAno.SelectedValue
                Sql &= ", " & pMes
                Sql &= ", '" & rs("Produto_Id") & "'"
                Sql &= ", " & rs("CodigoDeCusto_Id") + 1
                Sql &= ", ''" 'Empresa Destino
                Sql &= ", 0"  'End Empresa Destino
                Sql &= ", ''" '& rs("DepositoDestino_Id") & "'"
                Sql &= ", 0" '& rs("EndDepositoDestino_Id")
                Sql &= ", ''"

                SqlAux = "        SELECT    ValorDoProduto / Quantidade AS Medio"
                SqlAux = SqlAux & " FROM ApuracaoDeCustos"
                SqlAux = SqlAux & " WHERE Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
                SqlAux = SqlAux & " And Deposito_Id = '" & rs("Deposito_Id") & "'"
                SqlAux = SqlAux & " And EndDeposito_ID = " & rs("EndDeposito_Id")
                SqlAux = SqlAux & " AND (Ano_Id = " & DdlAno.SelectedValue & ")"
                SqlAux = SqlAux & " AND (Mes_Id = " & pMes & ")"
                SqlAux = SqlAux & " And Produto_Id = '" & rs("Produto_Id") & "'"
                SqlAux = SqlAux & " AND (CodigoDeCusto_Id = 495)"
                SqlAux = SqlAux & " AND (Quantidade > 0)"

                Medio = 0
                Dim ds2 As DataSet = Banco.ConsultaDataSet(SqlAux, "ApuracaoDeCustos")
                If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
                    For Each rsa As DataRow In ds2.Tables(0).Rows
                        Medio = rsa("Medio")
                    Next
                End If

                Medio = Medio - (rs("ValorDoProduto") / rs("Quantidade"))
                Valor = rs("Quantidade") * Medio

                Sql &= ", 0" '& str(rs("Quantidade"))
                Sql &= ", " & Replace(Valor, ",", ".")
                Sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                Sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                Sql &= ", ''" '& rs("Produto_Id") & "'"
                Sql &= ", " & rs("CodigoDestino")
                Sql &= ")"

                If Valor <> 0 Then
                    Sqll = Sqll & Sql
                End If
            Next
        End If

        If Sqll <> "" Then
            Return Banco.GravaBanco(Sqll)
        Else
            Return True
        End If
    End Function

    Private Function AjustaProdutoDerivado(ByVal pMes As Integer) As Boolean
        '15 - AJUSTANDO PRODUTOS DERIVADOS
        Dim Sql As String = "   Select case when len(EmpresaDestino_Id) > 0 then EmpresaDestino_Id else Empresa_Id end as Empresa_Id,"
        Sql &= "          case when len(EmpresaDestino_Id) > 0 then EndEmpresaDestino_Id else EndEmpresa_Id end as EndEmpresa_Id,"
        Sql &= "          case when len(DepositoDestino_Id) > 0 then DepositoDestino_Id else Deposito_Id end as Deposito_Id,"
        Sql &= "          case when len(DepositoDestino_Id) > 0 then EndDepositoDestino_Id else EndDeposito_Id end as EndDeposito_Id,"
        Sql &= "          Ano_Id, Mes_Id,Produto_Id as ProdutoOrigem, ProdutoDerivado_Id as Produto_Id, CodigoDestino as CodigoDeCusto_Id, '' as EmpresaDestino_Id, 0 as EndEmpresaDestino_Id, "
        Sql &= "          '' as DepositoDestino_Id, 0 as EndDepositoDestino_Id, '' as ProdutoDerivado_Id, isnull(Etapa,0) as Etapa, "
        Sql &= "            Case when CodigoDestino In (350, 351, 352, 358, 750, 751, 755, 757,758) then Quantidade else 0 end as Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, "
        Sql &= "          CodigoDeCusto_Id as CodigoDestino, Reduzido"
        Sql &= "   from ApuracaoDeCustos INNER JOIN"
        Sql &= "        PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
        Sql &= " WHERE Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   and Ano_id        = " & DdlAno.SelectedValue
        Sql &= "   And Mes_id        = " & pMes
        Sql &= "   And ProdutoDerivado_Id > 0"
        Sql &= "   And (CodigoDestino = 301)"

        Sql &= "   AND LEFT(ISNULL(ApuracaoDeCustos.Produto_Id,''), 5) NOT IN ('10101','10102')"

        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "                        when exists ("
                Sql &= "                                      select 1 "
                Sql &= "                                        from ApuracaoDeCustos "
                Sql &= "                                       Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                Sql &= "                                         And EndEmpresa_Id         = " & rs("EndEmpresa_Id")
                Sql &= "                                         And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                Sql &= "                                         And EndDeposito_Id        = " & rs("EndDeposito_Id")
                Sql &= "                                         And Ano_Id                = " & rs("Ano_Id")
                Sql &= "                                         And Mes_Id                = " & rs("Mes_Id")
                Sql &= "                                         And Produto_Id            ='" & rs("Produto_Id") & "'"
                Sql &= "                                         And CodigoDeCusto_Id      = " & rs("CodigoDeCusto_Id")
                If rs("CodigoDeCusto_Id") = 301 Then
                    Sql &= "   And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                    Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                    Sql &= "   And DepositoDestino_Id    ='" & rs("Empresa_Id") & "'"
                    Sql &= "   And EndDepositoDestino_Id = " & rs("EndEmpresa_Id")
                Else
                    Sql &= "   And EmpresaDestino_Id     ='" & rs("EmpresaDestino_Id") & "'"
                    Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresaDestino_Id")
                    Sql &= "   And DepositoDestino_Id    ='" & rs("DepositoDestino_Id") & "'"
                    Sql &= "   And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                End If

                'Sql &= "                                         And EmpresaDestino_Id     ='" & rs("EmpresaDestino_Id") & "'"
                'Sql &= "                                         And EndEmpresaDestino_Id  = " & rs("EndEmpresaDestino_Id")
                'Sql &= "                                         And DepositoDestino_Id    ='" & rs("DepositoDestino_Id") & "'"
                'Sql &= "                                         And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")

                Sql &= "                                         And ProdutoDerivado_Id    ='" & rs("ProdutoOrigem") & "'"
                Sql &= "                                    )"
                Sql &= "                          then 'S'"
                Sql &= "                          else 'N'"
                Sql &= "                     end) ;" & vbCrLf

                If rs("Quantidade") > 0 Then
                    Sql &= " if @Exist = 'N' "
                    Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                    Sql &= "  Empresa_Id"
                    Sql &= ", EndEmpresa_Id"
                    Sql &= ", Deposito_Id"
                    Sql &= ", EndDeposito_Id"
                    Sql &= ", Ano_Id"
                    Sql &= ", Mes_Id"
                    Sql &= ", Produto_Id"
                    Sql &= ", CodigoDeCusto_Id"
                    Sql &= ", EmpresaDestino_Id"
                    Sql &= ", EndEmpresaDestino_Id"
                    Sql &= ", DepositoDestino_Id"
                    Sql &= ", EndDepositoDestino_Id"
                    Sql &= ", ProdutoDerivado_Id"
                    Sql &= ", Etapa"
                    Sql &= ", Quantidade"
                    Sql &= ", ValorDoProduto"
                    Sql &= ", ValorDoFrete"
                    Sql &= ", ValorAuxiliar"
                    Sql &= ", ProdutoDestino"
                    Sql &= ", CodigoDestino"
                    Sql &= ", Reduzido)"
                    Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                    Sql &= ", " & rs("EndEmpresa_Id")
                    Sql &= ",'" & rs("Deposito_Id") & "'"
                    Sql &= ", " & rs("EndDeposito_Id")
                    Sql &= ", " & rs("Ano_Id")
                    Sql &= ", " & rs("Mes_Id")
                    Sql &= ",'" & rs("Produto_Id") & "'"
                    Sql &= ", " & rs("CodigoDeCusto_Id")

                    If rs("CodigoDeCusto_Id") = 301 Then
                        Sql &= ",'" & rs("Empresa_Id") & "'"
                        Sql &= ", " & rs("EndEmpresa_Id")
                        Sql &= ",'" & rs("Empresa_Id") & "'"
                        Sql &= ", " & rs("EndEmpresa_Id")
                    Else
                        Sql &= ",'" & rs("EmpresaDestino_Id") & "'"
                        Sql &= ", " & rs("EndEmpresaDestino_Id")
                        Sql &= ",'" & rs("DepositoDestino_Id") & "'"
                        Sql &= ", " & rs("EndDepositoDestino_Id")
                    End If

                    Sql &= ",'" & rs("ProdutoOrigem") & "'"
                    Sql &= ", " & rs("Etapa")
                    Sql &= ", " & Replace(rs("Quantidade"), ",", ".")
                    Sql &= ", " & Replace(rs("ValorDoProduto"), ",", ".")
                    Sql &= ", " & Replace(rs("ValorDoFrete"), ",", ".")
                    Sql &= ", " & Replace(rs("ValorAuxiliar"), ",", ".")
                    Sql &= ",'" & rs("ProdutoDestino") & "'"
                    Sql &= ", " & rs("CodigoDestino")
                    Sql &= ",'" & rs("Reduzido") & "'"
                    Sql &= ")"

                    Sql &= "  Update ApuracaoDeCustos set "
                    Sql &= "         Quantidade         =  " & Replace(rs("Quantidade"), ",", ".")
                    Sql &= ",        ValorDoProduto     =  " & Replace(rs("ValorDoProduto"), ",", ".")
                    Sql &= "        ,ValorDoFrete       =  " & Replace(rs("ValorDoFrete"), ",", ".")
                    Sql &= " Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                    Sql &= "   And EndEmpresa_Id         = " & rs("EndEmpresa_Id")

                    If rs("CodigoDestino") = "355" _
                        And (Left(Empresa(0), 8) = "05272759" _
                             OrElse Left(Empresa(0), 8) = "05366261" _
                             OrElse Left(Empresa(0), 8) = "38198213" _
                             OrElse Left(Empresa(0), 8) = "62747840" _
                             OrElse Left(Empresa(0), 8) = "40938762" _
                             OrElse Left(Empresa(0), 8) = "44005444" _
                             OrElse Left(Empresa(0), 8) = "48984539") Then
                        Sql &= "   And Deposito_Id           ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndDeposito_Id        = " & rs("EndEmpresa_Id")
                    Else
                        Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                        Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                    End If

                    'Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                    'Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                    Sql &= "   And Ano_Id                = " & rs("Ano_Id")
                    Sql &= "   And Mes_Id                = " & rs("Mes_Id")
                    Sql &= "   And Produto_Id            ='" & rs("Produto_Id") & "'"
                    Sql &= "   And CodigoDeCusto_Id      = " & rs("CodigoDeCusto_Id")

                    If rs("CodigoDeCusto_Id") = 301 Then
                        Sql &= "   And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                        Sql &= "   And DepositoDestino_Id    ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndDepositoDestino_Id = " & rs("EndEmpresa_Id")
                    Else
                        Sql &= "   And EmpresaDestino_Id     ='" & rs("EmpresaDestino_Id") & "'"
                        Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresaDestino_Id")
                        Sql &= "   And DepositoDestino_Id    ='" & rs("DepositoDestino_Id") & "'"
                        Sql &= "   And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                    End If

                    Sql &= "   And ProdutoDerivado_Id    ='" & rs("ProdutoOrigem") & "'"
                    Sqll = Sqll & Sql
                End If

                If rs("Quantidade") = 0 Then
                    Sql &= " if @Exist = 'N' "
                    Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                    Sql &= "  Empresa_Id"
                    Sql &= ", EndEmpresa_Id"
                    Sql &= ", Deposito_Id"
                    Sql &= ", EndDeposito_Id"
                    Sql &= ", Ano_Id"
                    Sql &= ", Mes_Id"
                    Sql &= ", Produto_Id"
                    Sql &= ", CodigoDeCusto_Id"
                    Sql &= ", EmpresaDestino_Id"
                    Sql &= ", EndEmpresaDestino_Id"
                    Sql &= ", DepositoDestino_Id"
                    Sql &= ", EndDepositoDestino_Id"
                    Sql &= ", ProdutoDerivado_Id"
                    Sql &= ", Etapa"
                    Sql &= ", Quantidade"
                    Sql &= ", ValorDoProduto"
                    Sql &= ", ValorDoFrete"
                    Sql &= ", ValorAuxiliar"
                    Sql &= ", ProdutoDestino"
                    Sql &= ", CodigoDestino"
                    Sql &= ", Reduzido)"
                    Sql &= " VALUES('" & rs("Empresa_Id") & "'"
                    Sql &= ", " & rs("EndEmpresa_Id")
                    Sql &= ",'" & rs("Deposito_Id") & "'"
                    Sql &= ", " & rs("EndDeposito_Id")
                    Sql &= ", " & rs("Ano_Id")
                    Sql &= ", " & rs("Mes_Id")
                    Sql &= ",'" & rs("Produto_Id") & "'"
                    Sql &= ", " & rs("CodigoDeCusto_Id")

                    If rs("CodigoDeCusto_Id") = 301 Then
                        Sql &= ",'" & rs("Empresa_Id") & "'"
                        Sql &= ", " & rs("EndEmpresa_Id")
                        Sql &= ",'" & rs("Empresa_Id") & "'"
                        Sql &= ", " & rs("EndEmpresa_Id")
                    Else
                        Sql &= ",'" & rs("EmpresaDestino_Id") & "'"
                        Sql &= ", " & rs("EndEmpresaDestino_Id")
                        Sql &= ",'" & rs("DepositoDestino_Id") & "'"
                        Sql &= ", " & rs("EndDepositoDestino_Id")
                    End If

                    Sql &= ",'" & rs("ProdutoOrigem") & "'"
                    Sql &= ", " & rs("Etapa")
                    Sql &= ", " & Replace(rs("Quantidade"), ",", ".")
                    Sql &= ", " & Replace(rs("ValorDoProduto"), ",", ".")
                    Sql &= ", " & Replace(rs("ValorDoFrete"), ",", ".")
                    Sql &= ", " & Replace(rs("ValorAuxiliar"), ",", ".")
                    Sql &= ",'" & rs("ProdutoDestino") & "'"
                    Sql &= ", " & rs("CodigoDestino")
                    Sql &= ",'" & rs("Reduzido") & "'"
                    Sql &= ")"

                    Sql &= "  Update ApuracaoDeCustos set "
                    Sql &= "         Quantidade         =  " & Replace(rs("Quantidade"), ",", ".")
                    Sql &= ",        ValorDoProduto     =  " & Replace(rs("ValorDoProduto"), ",", ".")
                    Sql &= "        ,ValorDoFrete       =  " & Replace(rs("ValorDoFrete"), ",", ".")
                    Sql &= " Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                    Sql &= "   And EndEmpresa_Id         = " & rs("EndEmpresa_Id")

                    If rs("CodigoDestino") = "355" _
                        And (Left(Empresa(0), 8) = "05272759" _
                             OrElse Left(Empresa(0), 8) = "05366261" _
                             OrElse Left(Empresa(0), 8) = "38198213" _
                             OrElse Left(Empresa(0), 8) = "62747840" _
                             OrElse Left(Empresa(0), 8) = "40938762" _
                             OrElse Left(Empresa(0), 8) = "44005444" _
                             OrElse Left(Empresa(0), 8) = "48984539") Then
                        Sql &= "   And Deposito_Id           ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndDeposito_Id        = " & rs("EndEmpresa_Id")
                    Else
                        Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                        Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                    End If

                    'Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                    'Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")


                    Sql &= "   And Ano_Id                = " & rs("Ano_Id")
                    Sql &= "   And Mes_Id                = " & rs("Mes_Id")
                    Sql &= "   And Produto_Id            ='" & rs("Produto_Id") & "'"
                    Sql &= "   And CodigoDeCusto_Id      = " & rs("CodigoDeCusto_Id")

                    If rs("CodigoDeCusto_Id") = 301 Then
                        Sql &= "   And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                        Sql &= "   And DepositoDestino_Id    ='" & rs("Empresa_Id") & "'"
                        Sql &= "   And EndDepositoDestino_Id = " & rs("EndEmpresa_Id")
                    Else
                        Sql &= "   And EmpresaDestino_Id     ='" & rs("EmpresaDestino_Id") & "'"
                        Sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresaDestino_Id")
                        Sql &= "   And DepositoDestino_Id    ='" & rs("DepositoDestino_Id") & "'"
                        Sql &= "   And EndDepositoDestino_Id = " & rs("EndDepositoDestino_Id")
                    End If

                    Sql &= "   And ProdutoDerivado_Id    ='" & rs("ProdutoOrigem") & "'"
                    Sqll = Sqll & Sql
                End If

            Next
        End If
        Return Banco.GravaBanco(Sqll)
    End Function

    Private Function AjustaProdutoDerivadoNovo(ByVal pMes As Integer) As Boolean

        Dim SqlArrayList As New ArrayList

        Dim Sql As String = " SELECT        Empresa_Id, "
        Sql &= "      EndEmpresa_Id, "
        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id In (350, 355, 358, 750, 751, 758)  then DepositoDestino_Id    else Empresa_Id    End as Deposito_Id, "
        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id in (350, 355, 358, 750, 751, 758) then EndDepositoDestino_Id else EndEmpresa_Id End as EndDeposito_Id, "
        Sql &= "	  ApuracaoDeCustos.Ano_Id, "
        Sql &= "	  ApuracaoDeCustos.Mes_Id, "
        Sql &= "	  ApuracaoDeCustos.ProdutoDerivado_Id AS Produto_Id, "
        Sql &= "	  ApuracaoDeCustos.CodigoDestino AS CodigoDeCusto_Id,"

        Sql &= "	  Empresa_Id     AS EmpresaDestino_Id, "
        Sql &= "      EndEmpresa_Id  AS EmpresaDestino_Id,"

        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id In (350, 355, 358, 750, 751, 758)  then Empresa_Id     else Empresa_Id    End as DepositoDestino_Id, "
        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id in (350, 355, 358, 750, 751, 758) then EndEmpresa_Id  else EndEmpresa_Id End as EndDepositoDestino_Id, "

        Sql &= "      ApuracaoDeCustos.Produto_Id AS ProdutoDerivado_Id,  "
        Sql &= "	  CASE WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 501  THEN 0 ELSE Quantidade END AS Quantidade,"
        Sql &= "	  0 as ValorDoProduto,"
        Sql &= "	  ApuracaoDeCustos.ValorDoFrete,"
        Sql &= "	  ApuracaoDeCustos.ValorAuxiliar,"
        Sql &= "	  ApuracaoDeCustos.ProdutoDestino, "
        Sql &= "      ApuracaoDeCustos.CodigoDeCusto_Id AS CodigoDestino,"
        Sql &= "      ApuracaoDeCustos.Reduzido"


        Sql &= "      FROM  ApuracaoDeCustos INNER JOIN"
        Sql &= "            PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"

        Sql &= "      WHERE Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "        and Ano_id        = " & DdlAno.SelectedValue
        Sql &= "        And Mes_id        = " & pMes
        Sql &= "       And ProdutoDerivado_Id > 0"
        Sql &= "       and isnull(codigodestino,0) > 0"
        Sql &= "        And (ApuracaoDeCustos.CodigoDeCusto_Id IN (350, 355, 358, 750, 751, 758)) "

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")

        For Each rs As DataRow In ds.Tables(0).Rows

            Try

                Dim checkSql As String = "  SELECT COUNT(*) COUNT
                                            FROM ApuracaoDeCustos
                                            WHERE Empresa_Id = '" & rs("Empresa_Id") & "'
                                              AND Ano_Id = " & rs("Ano_Id") & "
                                              AND Mes_Id = " & rs("Mes_Id") & "
                                              AND Produto_Id = '" & rs("Produto_Id") & "'
                                              AND CodigoDeCusto_Id = '" & rs("CodigoDeCusto_Id") & "' 
                                              AND ProdutoDestino = '" & rs("ProdutoDestino") & "'; "

                Dim existe As Integer
                Dim dsVerificar As DataSet = Banco.ConsultaDataSet(checkSql, "checkSql")

                For Each rsVerificar As DataRow In dsVerificar.Tables(0).Rows
                    existe = rsVerificar("Count")
                Next

                If existe = 0 Then
                    ' INSERT
                    Dim insertSql As String = "INSERT INTO ApuracaoDeCustos (Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, " &
                                      "DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino, Reduzido) " &
                                      "VALUES ('" & rs("Empresa_Id") & "', '" & rs("EndEmpresa_Id") & "', '" & rs("Deposito_Id") & "', '" & rs("EndDeposito_Id") & "', " &
                                      rs("Ano_Id") & ", " & rs("Mes_Id") & ", '" & rs("Produto_Id") & "', '" & rs("CodigoDeCusto_Id") & "', '" & rs("EmpresaDestino_Id") & "', '" & rs("EndEmpresa_Id") & "', " &
                                      "'" & rs("DepositoDestino_Id") & "', '" & rs("EndDepositoDestino_Id") & "', " & rs("ProdutoDerivado_Id") & ", " & Replace(rs("Quantidade"), ",", ".") & ", " &
                                      Replace(rs("ValorDoProduto"), ",", ".") & ", " & Replace(rs("ValorDoFrete"), ",", ".") & ", " & Replace(rs("ValorAuxiliar"), ",", ".") & ", '" & rs("ProdutoDestino") & "', " &
                                      rs("CodigoDestino") & ", '" & rs("Reduzido") & "')"
                    SqlArrayList.Add(insertSql)

                Else
                    ' UPDATE
                    Dim updateSql As String = "UPDATE ApuracaoDeCustos SET " &
                                                  "Quantidade = Quantidade + " & Replace(rs("Quantidade"), ",", ".") & ", " &
                                                  "ValorDoFrete = ValorDoFrete + " & Replace(rs("ValorDoFrete"), ",", ".") & ", " &
                                                  "ValorAuxiliar = ValorAuxiliar + " & Replace(rs("ValorAuxiliar"), ",", ".") & " " &
                                                  "WHERE Empresa_Id = '" & rs("Empresa_Id") & "' AND " &
                                                  "Ano_Id = " & rs("Ano_Id") & " AND " &
                                                  "Mes_Id = " & rs("Mes_Id") & " AND " &
                                                  "Produto_Id = '" & rs("Produto_Id") & "' AND " &
                                                  "CodigoDeCusto_Id = '" & rs("CodigoDeCusto_Id") & "' AND " &
                                                  "ProdutoDestino = '" & rs("ProdutoDestino") & "'"

                    SqlArrayList.Add(updateSql)

                End If

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        Next

        Return Banco.GravaBanco(SqlArrayList)

    End Function

    Private Function AjustaProdutoDerivadoNovo_bkp(ByVal pMes As Integer) As Boolean

        Dim Sql As String = "Insert into ApuracaoDeCustos (Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, "
        Sql &= "                 DepositoDestino_Id, EndDepositoDestino_Id, Produtoderivado_Id, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino,CodigoDestino, Reduzido) "

        Sql &= " SELECT        Empresa_Id, "
        Sql &= "      EndEmpresa_Id, "
        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id In (350, 355, 358, 750, 751, 758)  then DepositoDestino_Id    else Empresa_Id    End as Deposito_Id, "
        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id in (350, 355, 358, 750, 751, 758) then EndDepositoDestino_Id else EndEmpresa_Id End as EndDeposito_Id, "
        Sql &= "	  ApuracaoDeCustos.Ano_Id, "
        Sql &= "	  ApuracaoDeCustos.Mes_Id, "
        Sql &= "	  ApuracaoDeCustos.ProdutoDerivado_Id AS Produto_Id, "
        Sql &= "	  ApuracaoDeCustos.CodigoDestino AS CodigoDeCusto_Id,"

        Sql &= "	  Empresa_Id     AS EmpresaDestino_Id, "
        Sql &= "      EndEmpresa_Id  AS EmpresaDestino_Id,"

        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id In (350, 355, 358, 750, 751, 758)  then Empresa_Id     else Empresa_Id    End as DepositoDestino_Id, "
        Sql &= "	  Case when ApuracaoDeCustos.CodigoDeCusto_Id in (350, 355, 358, 750, 751, 758) then EndEmpresa_Id  else EndEmpresa_Id End as EndDepositoDestino_Id, "

        Sql &= "      ApuracaoDeCustos.Produto_Id AS ProdutoDerivado_Id,  "
        Sql &= "	  CASE WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 501  THEN 0 ELSE Quantidade END AS Quantidade,"
        Sql &= "	  0 as ValorDoProduto,"
        Sql &= "	  ApuracaoDeCustos.ValorDoFrete,"
        Sql &= "	  ApuracaoDeCustos.ValorAuxiliar,"
        Sql &= "	  ApuracaoDeCustos.ProdutoDestino, "
        Sql &= "      ApuracaoDeCustos.CodigoDeCusto_Id AS CodigoDestino,"
        Sql &= "      ApuracaoDeCustos.Reduzido"


        Sql &= "      FROM  ApuracaoDeCustos INNER JOIN"
        Sql &= "            PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"

        Sql &= "      WHERE Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "        and Ano_id        = " & DdlAno.SelectedValue
        Sql &= "        And Mes_id        = " & pMes
        Sql &= "       And ProdutoDerivado_Id > 0"
        Sql &= "       and isnull(codigodestino,0) > 0"
        Sql &= "        And (ApuracaoDeCustos.CodigoDeCusto_Id IN (350, 355, 358, 750, 751, 758)) "

        Return Banco.GravaBanco(Sql)

    End Function

    Private Function AjustaTransferencias(ByVal pMes As Integer) As Boolean
        '15 - AJUSTANDO PRODUTOS DERIVADOS (TRANSFERÊNCIAS)
        Dim Sql As String = " SELECT ApuracaoDeCustos.Empresa_Id," &
                        "    ApuracaoDeCustos.EndEmpresa_Id," &
                        "    ApuracaoDeCustos.Deposito_Id," &
                        "    ApuracaoDeCustos.EndDeposito_Id," &
                        "    ApuracaoDeCustos.Ano_Id," &
                        "    ApuracaoDeCustos.Mes_Id," &
                        "    ApuracaoDeCustos.Produto_Id," &
                        "    ApuracaoDeCustos.CodigoDeCusto_Id," &
                        "    ApuracaoDeCustos.EmpresaDestino_Id," &
                        "    ApuracaoDeCustos.EndEmpresaDestino_Id," &
                        "    ApuracaoDeCustos.DepositoDestino_Id," &
                        "    ApuracaoDeCustos.EndDepositoDestino_Id," &
                        "    ApuracaoDeCustos.ProdutoDerivado_Id," &
                        "    ApuracaoDeCustos.Quantidade," &
                        "    ApuracaoDeCustos.ValorDoProduto," &
                        "    ApuracaoDeCustos.ValorDoFrete," &
                        "    ApuracaoDeCustos.ValorAuxiliar," &
                        "    ApuracaoDeCustos.ProdutoDestino," &
                        "    ApuracaoDeCustos.CodigoDestino," &
                        "    ISNULL(ApuracaoDeCustos.Etapa,0) AS Etapa "

        Sql &= " FROM      ApuracaoDeCustos INNER JOIN"
        Sql &= "           PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"

        Sql &= " WHERE     ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   and     ApuracaoDeCustos.Ano_id        = " & DdlAno.SelectedValue
        Sql &= "   And     ApuracaoDeCustos.Mes_id        = " & pMes
        Sql &= "   And     (ApuracaoDeCustos.CodigoDeCusto_Id > 500)"
        Sql &= "   And     ((PlanoDeCustos.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "') Or  (ApuracaoDeCustos.CodigoDeCusto_Id in (750,751,755,758)))"
        Sql &= "   And     ISNULL(ApuracaoDeCustos.codigodestino,0) > 0"
        Sql &= "   And     ISNULL(ApuracaoDeCustos.Etapa,0) <> 99"  'NÃO REPROCESSA LINHA GERADA PELO AJUSTE

        '==========================================================
        ' ISOLAMENTO: NÃO PROCESSA GRUPOS 10101 e 10102 (Produto_Id)
        '==========================================================
        'Sql &= "   AND LEFT(ApuracaoDeCustos.Produto_Id, 5) NOT IN ('10101','10102')"
        '==========================================================

        Sqll = " Declare"
        Sqll &= " @Exist as varchar(1)" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

            For Each rs As DataRow In ds.Tables(0).Rows

                Sql = " set @Exist = (select case " &
                  "                        when exists (" &
                  "                                      select 1 " &
                  "                                        from ApuracaoDeCustos " &
                  "                                       Where Empresa_Id            ='" & rs("EmpresaDestino_Id") & "'" &
                  "                                         And EndEmpresa_Id         = " & rs("EndEmpresaDestino_Id") &
                  "                                         And Deposito_Id           ='" & rs("DepositoDestino_Id") & "'" &
                  "                                         And EndDeposito_Id        = " & rs("EndDepositoDestino_Id") &
                  "                                         And Ano_Id                = " & rs("Ano_Id") &
                  "                                         And Mes_Id                = " & rs("Mes_Id") &
                  "                                         And Produto_Id            ='" & rs("Produto_Id") & "'" &
                  "                                         And CodigoDeCusto_Id      = " & rs("CodigoDestino") &
                  "                                         And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'" &
                  "                                         And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id") &
                  "                                         And DepositoDestino_Id    ='" & rs("Deposito_Id") & "'" &
                  "                                         And EndDepositoDestino_Id = " & rs("EndDeposito_Id") &
                  "                                         And ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'" &
                  "                                    )" &
                  "                          then 'S'" &
                  "                          else 'N'" &
                  "                     end) ;" & vbCrLf

                Sql &= " if @Exist = 'N' " & vbCrLf
                Sql &= "  INSERT INTO ApuracaoDeCustos ( " &
                   "  Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id," &
                   "  EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id," &
                   "  Etapa, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino) " & vbCrLf

                Sql &= " VALUES('" & rs("EmpresaDestino_Id") & "'" &
                   ", " & rs("EndEmpresaDestino_Id") &
                   ",'" & rs("DepositoDestino_Id") & "'" &
                   ", " & rs("EndDepositoDestino_Id") &
                   ", " & rs("Ano_Id") &
                   ", " & rs("Mes_Id") &
                   ",'" & rs("Produto_Id") & "'" &
                   ", " & rs("CodigoDestino") &
                   ",'" & rs("Empresa_Id") & "'" &
                   ", " & rs("EndEmpresa_Id") &
                   ",'" & rs("Deposito_Id") & "'" &
                   ", " & rs("EndDeposito_Id") &
                   ",'" & rs("ProdutoDerivado_Id") & "'" &
                   ", 99" &
                   ", " & Replace(rs("Quantidade"), ",", ".") &
                   ", " & Replace(rs("ValorDoProduto"), ",", ".") &
                   ", " & Replace(rs("ValorDoFrete"), ",", ".") &
                   ", " & Replace(rs("ValorAuxiliar"), ",", ".") &
                   ",'" & rs("ProdutoDestino") & "'" &
                   ", " & rs("CodigoDeCusto_Id") &
                   ")" & vbCrLf

                Sql &= " Else " & vbCrLf
                Sql &= "  UPDATE ApuracaoDeCustos SET " & vbCrLf &
                   "         Etapa          = 99" & vbCrLf &
                   "        ,Quantidade     = " & Replace(rs("Quantidade"), ",", ".") & vbCrLf &
                   "        ,ValorDoProduto = " & Replace(rs("ValorDoProduto"), ",", ".") & vbCrLf &
                   "        ,ValorDoFrete   = " & Replace(rs("ValorDoFrete"), ",", ".") & vbCrLf &
                   "        ,ValorAuxiliar  = " & Replace(rs("ValorAuxiliar"), ",", ".") & vbCrLf &
                   "   WHERE Empresa_Id            ='" & rs("EmpresaDestino_Id") & "'" & vbCrLf &
                   "     AND EndEmpresa_Id         = " & rs("EndEmpresaDestino_Id") & vbCrLf &
                   "     AND Deposito_Id           ='" & rs("DepositoDestino_Id") & "'" & vbCrLf &
                   "     AND EndDeposito_Id        = " & rs("EndDepositoDestino_Id") & vbCrLf &
                   "     AND Ano_Id                = " & rs("Ano_Id") & vbCrLf &
                   "     AND Mes_Id                = " & rs("Mes_Id") & vbCrLf &
                   "     AND Produto_Id            ='" & rs("Produto_Id") & "'" & vbCrLf &
                   "     AND CodigoDeCusto_Id      = " & rs("CodigoDestino") & vbCrLf &
                   "     AND EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'" & vbCrLf &
                   "     AND EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id") & vbCrLf &
                   "     AND DepositoDestino_Id    ='" & rs("Deposito_Id") & "'" & vbCrLf &
                   "     AND EndDepositoDestino_Id = " & rs("EndDeposito_Id") & vbCrLf &
                   "     AND ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'" & vbCrLf

                Sqll &= Sql
            Next

        End If

        Return Banco.GravaBanco(Sqll)
    End Function

    Private Function AjustaDepositos(ByVal pMes As Integer) As Boolean
        '15 - AJUSTANDO Depositos
        Dim Sql As String = " SELECT ApuracaoDeCustos.Empresa_Id,"
        Sql &= "    ApuracaoDeCustos.EndEmpresa_Id,"
        Sql &= "    ApuracaoDeCustos.Deposito_Id,"
        Sql &= "    ApuracaoDeCustos.EndDeposito_Id,"
        Sql &= "    ApuracaoDeCustos.Ano_Id,"
        Sql &= "    ApuracaoDeCustos.Mes_Id,"
        Sql &= "    ApuracaoDeCustos.Produto_Id,"
        Sql &= "    ApuracaoDeCustos.CodigoDeCusto_Id,"
        Sql &= "    ApuracaoDeCustos.DepositoDestino_Id as EmpresaDestino_Id,"
        Sql &= "    ApuracaoDeCustos.EndDepositoDestino_Id as EndEmpresaDestino_Id,"
        Sql &= "    ApuracaoDeCustos.DepositoDestino_Id,"
        Sql &= "    ApuracaoDeCustos.EndDepositoDestino_Id,"
        Sql &= "    ApuracaoDeCustos.ProdutoDerivado_Id,"
        Sql &= "    ApuracaoDeCustos.Quantidade,"
        Sql &= "    ApuracaoDeCustos.ValorDoProduto,"
        Sql &= "    ApuracaoDeCustos.ValorDoFrete,"
        Sql &= "    ApuracaoDeCustos.ValorAuxiliar,"
        Sql &= "    ApuracaoDeCustos.ProdutoDestino,"
        Sql &= "    ApuracaoDeCustos.CodigoDestino "

        Sql &= " FROM      ApuracaoDeCustos INNER JOIN"
        Sql &= "           PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"

        Sql &= " WHERE     ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   and     ApuracaoDeCustos.Ano_id        = " & DdlAno.SelectedValue
        Sql &= "   And     ApuracaoDeCustos.Mes_id        = " & pMes
        Sql &= "   And     (ApuracaoDeCustos.CodigoDeCusto_Id > 500)"
        Sql &= "   And     (PlanoDeCustos.Classe = 'DEPOSITOS')"

        Sql &= "   and     isnull(ApuracaoDeCustos.codigodestino,0) > 0"

        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows

                Sql = "  Update ApuracaoDeCustos set "
                Sql &= "         ValorDoProduto     =  " & Replace(rs("ValorDoProduto"), ",", ".")
                Sql &= "        ,ValorDoFrete       =  " & Replace(rs("ValorDoFrete"), ",", ".")
                Sql &= " Where Empresa_Id            ='" & rs("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id         = " & rs("EndEmpresa_Id")
                Sql &= "   And Deposito_Id           ='" & rs("Empresa_Id") & "'"
                Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                Sql &= "   And Ano_Id                = " & rs("Ano_Id")
                Sql &= "   And Mes_Id                = " & rs("Mes_Id")
                Sql &= "   And Produto_Id            ='" & rs("Produto_Id") & "'"
                Sql &= "   And CodigoDeCusto_Id      = " & rs("CodigoDestino")
                Sql &= "   And EmpresaDestino_Id     = ''"
                Sql &= "   And EndEmpresaDestino_Id  = 0"
                Sql &= "   And DepositoDestino_Id    ='" & rs("Deposito_Id") & "'"
                Sql &= "   And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                Sql &= "   And ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'"
                Sqll = Sqll & Sql
            Next
        End If
        Return Banco.GravaBanco(Sqll)
    End Function

    Private Function AjustaConsumoXProducao(ByVal pMes As Integer) As Boolean
        '15.1 - AJUSTANDO CONSUMOXPRODUCAO
        Dim aux As Boolean = True
        Dim Sql As String = "SELECT DISTINCT ProdutoOrigem_Id, CodigoCustoOrigem_Id FROM ConsumoxProducao "
        Sqll = ""

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = " Declare"
                Sql &= "   @Soma as numeric(18,9)"
                Sql &= "   select ACO.Empresa_Id,         ACO.EndEmpresa_id,"
                Sql &= "           ACO.Deposito_Id,        ACO.EndDeposito_Id,"
                Sql &= "           ACO.Ano_Id,"
                Sql &= "           ACO.Mes_Id,"
                Sql &= "           ACO.produto_Id as ProdutoOrigem,"
                Sql &= "           ACO.Quantidade as QuantidadeOrigem,"
                Sql &= "           ACO.ValorDoProduto as VlrProdutoOrigem,"
                Sql &= "           ACD.produto_Id as ProdutoDestino,"
                Sql &= "           ACD.CodigoDeCusto_Id as CodigoDeCustoDestino,"
                Sql &= "           ACD.Quantidade as QuantidadeDestino,"
                Sql &= "           TPM.ValorOficial as ValorDeMercado,"
                Sql &= "           (ACD.Quantidade/TPM.BaseDeCalculo)*TPM.ValorOficial as VlrTotal,"
                Sql &= "           convert(numeric(18,9),0) as Soma"
                Sql &= "       into #Temp"
                Sql &= "       from ApuracaoDeCustos ACO"
                Sql &= "       Inner Join(select AC.Empresa_Id,"
                Sql &= "                      AC.EndEmpresa_Id,"
                Sql &= "                      AC.Ano_Id,"
                Sql &= "                      AC.Mes_Id,"
                Sql &= "                      AC.Produto_Id,"
                Sql &= "                      AC.CodigoDeCusto_Id,"
                Sql &= "                      AC.Quantidade"
                Sql &= "                 from ApuracaoDeCustos AC "
                Sql &= "                   Inner Join ConsumoxProducao CP "
                Sql &= "                   on CP.ProdutoOrigem_Id    ='" & rs("ProdutoOrigem_Id") & "'"
                Sql &= "                   and CP.CodigoCustoOrigem_Id  = " & rs("CodigoCustoOrigem_Id")
                Sql &= "                   and CP.ProdutoDestino_Id     = AC.Produto_Id"
                Sql &= "                   and CP.CodigoCustoDestino_Id = AC.CodigoDeCusto_Id"
                Sql &= "                   ) ACD "
                Sql &= "      on ACD.Empresa_Id    = ACO.Empresa_Id"
                Sql &= "      and ACD.EndEmpresa_Id = ACO.EndEmpresa_Id"
                Sql &= "      and ACD.Ano_Id        = ACO.Ano_Id"
                Sql &= "      and ACD.Mes_Id        = ACO.Mes_Id"
                Sql &= "    Inner Join (SELECT TPM1.Empresa_Id,    TPM1.EndEmpresa_Id,"
                Sql &= "                       TPM1.Deposito_Id,   TPM1.EndDeposito_Id,"
                Sql &= "                       TPM1.Produto_Id,    TPM1.Data_Id,"
                Sql &= "                       TPM1.ValorOficial,  TPM1.ValorMoeda,"
                Sql &= "                       TPM1.BaseDeCalculo"
                Sql &= "                 FROM TabelaDePrecosDeMercado TPM1"
                Sql &= "                Inner Join (SELECT Empresa_Id,"
                Sql &= "                                    EndEmpresa_Id,"
                Sql &= "                                    Deposito_Id,"
                Sql &= "                                    EndDeposito_Id,"
                Sql &= "                                    Produto_Id,"
                Sql &= "                                    max(Data_Id) as Data_Id"
                Sql &= "                                from TabelaDePrecosDeMercado"
                Sql &= "                               Group by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id"
                Sql &= "                             ) TPM2"
                Sql &= "                     on TPM1.Empresa_Id     = TPM2.Empresa_Id"
                Sql &= "                   and TPM1.EndEmpresa_Id  = TPM2.EndEmpresa_Id"
                Sql &= "                   and TPM1.Deposito_Id    = TPM2.Deposito_Id"
                Sql &= "                    and TPM1.EndDeposito_Id = TPM2.EndDeposito_Id"
                Sql &= "                    and TPM1.Produto_Id     = TPM2.Produto_Id"
                Sql &= "                    and TPM1.Data_Id        = TPM2.Data_Id"
                Sql &= "                ) TPM"
                Sql &= "       on ACO.Empresa_Id     = TPM.Empresa_Id"
                Sql &= "       and ACO.EndEmpresa_id  = TPM.EndEmpresa_id"
                Sql &= "       and ACO.Deposito_Id    = TPM.Deposito_Id"
                Sql &= "       And ACO.EndDeposito_Id = TPM.EndDeposito_Id"
                Sql &= "       And ACO.Ano_Id         = year(TPM.Data_Id)"
                Sql &= "       And ACO.Mes_Id         = Month(TPM.Data_Id)"
                Sql &= "       And ACD.Produto_Id     = TPM.Produto_Id "
                Sql &= "       where ACO.Empresa_Id     ='" & Empresa(0) & "'"
                If Left(Empresa(0), 8) = "05272759" _
                    OrElse Left(Empresa(0), 8) = "05366261" _
                    OrElse Left(Empresa(0), 8) = "38198213" _
                    OrElse Left(Empresa(0), 8) = "62747840" _
                    OrElse Left(Empresa(0), 8) = "40938762" _
                    OrElse Left(Empresa(0), 8) = "44005444" _
                    OrElse Left(Empresa(0), 8) = "48984539" Then
                    Sql &= "       AND ACO.EndEmpresa_id    = " & Empresa(1)
                Else
                    Sql &= "       AND ACO.EndEmpresa_id    = 0" 'Custo do Curtume está 0 - Furlan - 11/11/2019
                End If
                Sql &= "       AND ACO.Ano_Id           = " & DdlAno.SelectedValue
                Sql &= "       And ACO.Mes_Id           = " & pMes
                Sql &= "       And ACO.Produto_Id       ='" & rs("ProdutoOrigem_Id") & "'"
                Sql &= "       And ACO.CodigoDeCusto_Id = " & rs("CodigoCustoOrigem_Id")
                Sql &= "       set @Soma =(Select Sum(VlrTotal) from #Temp)"
                Sql &= "   Update #Temp set"
                Sql &= "       Soma = @Soma"
                Sql &= "   Select  Empresa_Id,"
                Sql &= "           EndEmpresa_id,"
                Sql &= "           Deposito_Id,"
                Sql &= "           EndDeposito_Id,"
                Sql &= "           Ano_Id,"
                Sql &= "           Mes_Id,"
                Sql &= "           ProdutoOrigem,"
                Sql &= "           QuantidadeOrigem,"
                Sql &= "           VlrProdutoOrigem,"
                Sql &= "           ProdutoDestino,"
                Sql &= "           CodigoDeCustoDestino,"
                Sql &= "           QuantidadeDestino,"
                Sql &= "           ValorDeMercado,"
                Sql &= "           VlrTotal,"
                Sql &= "           (VlrTotal * 100) / Soma as Percentual,"
                Sql &= "           (VlrProdutoOrigem *  ((VlrTotal * 100) / Soma)) / 100 as CustoDoProdutoDestino "
                Sql &= "      from #Temp"

                Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
                If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
                    For Each rsa As DataRow In ds2.Tables(0).Rows
                        Sql = "  Update ApuracaoDeCustos set "
                        Sql &= "         ValorDoProduto     =  " & Replace(rsa("CustoDoProdutoDestino"), ",", ".")
                        Sql &= " Where Empresa_Id            ='" & rsa("Empresa_Id") & "'"
                        Sql &= "   And EndEmpresa_Id         = " & rsa("EndEmpresa_Id")
                        Sql &= "   And Deposito_Id           ='" & rsa("Deposito_Id") & "'"
                        Sql &= "   And EndDeposito_Id        = " & rsa("EndDeposito_Id")
                        Sql &= "   And Ano_Id                = " & rsa("Ano_Id")
                        Sql &= "   And Mes_Id                = " & rsa("Mes_Id")
                        Sql &= "   And Produto_Id            ='" & rsa("ProdutoDestino") & "'"
                        Sql &= "   And CodigoDeCusto_Id      = " & rsa("CodigoDeCustoDestino")
                        Sql &= "   And EmpresaDestino_Id     =''"
                        Sql &= "   And EndEmpresaDestino_Id  = 0"
                        Sql &= "   And DepositoDestino_Id    =''"
                        Sql &= "   And EndDepositoDestino_Id = 0"
                        Sql &= "   And ProdutoDerivado_Id    =''"
                        aux = Banco.GravaBanco(Sql)
                    Next
                End If
                aux = Banco.GravaBanco(Sql)
            Next
        End If
        Return aux
    End Function

    Private Sub CalculoCircular(ByVal pMes As Integer)
        '19 - Calculo Circular - Transferindo Custo para o destino
        'txt_Processo.value = "Calculo Circular - Transferindo Custo para o destino..."

        Dim Sql As String = "  SELECT Origem.Empresa_Id, "
        Sql &= "        Origem.EndEmpresa_Id, "
        Sql &= "        Origem.Deposito_Id, "
        Sql &= "        Origem.EndDeposito_Id, "
        Sql &= "        Origem.Produto_Id,"
        Sql &= "        Origem.CodigoDeCusto_Id, "
        Sql &= "        Origem.Ano_Id, "
        Sql &= "        Origem.Mes_Id, "
        Sql &= "        Origem.EmpresaDestino_Id, "
        Sql &= "        Origem.EndEmpresaDestino_Id, "
        Sql &= "        Origem.DepositoDestino_Id, "
        Sql &= "        Origem.EndDepositoDestino_Id, "
        Sql &= "        CASE WHEN Origem.ProdutoDerivado_Id = '000000' THEN '' ELSE Origem.ProdutoDerivado_Id END AS ProdutoDerivado_Id,"
        Sql &= "        Origem.CodigoDestino,"
        Sql &= "        Origem.ValorDoProduto as ValorOrigem, "
        Sql &= "        Origem.ValorDoFrete as FreteOrigem, "
        Sql &= "        isnull(Destino.ValorDoProduto, 0) AS ValorDestino"
        Sql &= " FROM   ApuracaoDeCustos AS Origem LEFT OUTER JOIN ApuracaoDeCustos AS Destino ON"
        Sql &= "        Origem.CodigoDeCusto_Id = Destino.CodigoDestino"
        Sql &= "        AND Origem.Produto_Id = Destino.ProdutoDerivado_Id  "
        Sql &= "        AND Origem.EndDeposito_Id = Destino.EndDepositoDestino_Id "
        Sql &= "        AND Origem.Deposito_Id = Destino.DepositoDestino_Id "
        Sql &= "        AND Origem.EndEmpresa_Id = Destino.EndEmpresaDestino_Id "
        Sql &= "        AND Origem.Empresa_Id = Destino.EmpresaDestino_Id "
        Sql &= "        AND Origem.EmpresaDestino_Id = Destino.Empresa_Id "
        Sql &= "        AND Origem.EndEmpresaDestino_Id = Destino.EndEmpresa_Id "
        Sql &= "        AND Origem.DepositoDestino_Id = Destino.Deposito_Id "
        Sql &= "        AND Origem.EndDepositoDestino_Id = Destino.EndDeposito_Id "
        Sql &= "        AND Origem.ProdutoDerivado_Id = Destino.Produto_Id "
        Sql &= "        AND Origem.Ano_Id = Destino.Ano_Id "
        Sql &= "        AND Origem.Mes_Id = Destino.Mes_Id "
        Sql &= "        AND Origem.CodigoDestino = Destino.CodigoDeCusto_Id"
        Sql &= " WHERE  (Origem.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        Sql &= "        And   (Origem.Ano_Id = " & DdlAno.SelectedValue & ")"
        Sql &= "        And   (Origem.Mes_Id = " & pMes & ")"
        Sql &= "        And   (Origem.CodigoDeCusto_Id > 495)"
        Sql &= "        And   (Origem.CodigoDestino <> 0)"
        Sql &= "        And   (Origem.ValorDoProduto > 0)"
        Sql &= "        And   (Origem.EmpresaDestino_Id <> '')"
        Sql &= "        And   (Origem.DepositoDestino_ID <> '')"
        Sql &= "        And   (Origem.ValorDoProduto <> Destino.ValorDoProduto)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Dim Circular = "S"

                Sql = " Declare"
                Sql &= " @Exist as varchar(1)"
                Sql &= " set @Exist = (select case "
                Sql &= "       when exists ("
                Sql &= "                 select Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Etapa, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino, Reduzido from ApuracaoDeCustos "
                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("EmpresaDestino_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresaDestino_Id")
                Sql &= " And Deposito_Id = '" & rs("DepositoDestino_Id") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDepositoDestino_Id")
                Sql &= " And Ano_Id = " & rs("Ano_Id")
                Sql &= " And Mes_Id = " & rs("Mes_Id")
                Sql &= " And Produto_Id = '" & rs("ProdutoDerivado_Id") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("CodigoDestino")

                Sql &= " And EmpresaDestino_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresaDestino_Id = " & rs("EndEmpresa_Id")
                Sql &= " And DepositoDestino_Id = '" & rs("Deposito_Id") & "'"
                Sql &= " And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                Sql &= " And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"

                Sql &= ")"
                Sql &= "            then  'S'"
                Sql &= "             else 'N'"
                Sql &= "               end) ;"

                Sql &= " if @Exist = 'N' "

                Sql &= "  INSERT INTO ApuracaoDeCustos ( "
                Sql &= "  Empresa_Id"
                Sql &= ", EndEmpresa_Id"
                Sql &= ", Deposito_Id"
                Sql &= ", EndDeposito_Id"
                Sql &= ", Ano_Id"
                Sql &= ", Mes_Id"
                Sql &= ", Produto_Id"
                Sql &= ", CodigoDeCusto_Id"
                Sql &= ", EmpresaDestino_Id"
                Sql &= ", EndEmpresaDestino_Id"
                Sql &= ", DepositoDestino_Id"
                Sql &= ", EndDepositoDestino_Id"
                Sql &= ", ProdutoDerivado_Id"
                Sql &= ", Etapa"
                Sql &= ", Quantidade"
                Sql &= ", ValorDoProduto"
                Sql &= ", ValorDoFrete"
                Sql &= ", ValorAuxiliar"
                Sql &= ", ProdutoDestino"
                Sql &= ", CodigoDestino"
                Sql &= ", Reduzido)"

                Sql &= " VALUES('" & rs("EmpresaDestino_Id") & "'"
                Sql &= ", " & rs("EndEmpresaDestino_Id")
                Sql &= ", '" & rs("DepositoDestino_Id") & "'"
                Sql &= ", " & rs("EndDepositoDestino_Id")
                Sql &= ", " & rs("Ano_Id")
                Sql &= ", " & rs("Mes_Id")
                Sql &= ", '" & rs("ProdutoDerivado_Id") & "'"
                Sql &= ", " & rs("CodigoDestino")

                Sql &= ", '" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("Deposito_Id") & "'"
                Sql &= ", " & rs("EndDeposito_Id")
                Sql &= ", '" & rs("Produto_Id") & "'"

                Sql &= ", 1"    'Etapa
                Sql &= ", 0"    'Quantidade
                Sql &= ", " & Replace(rs("ValorOrigem"), ",", ".")
                Sql &= ", " & Replace(rs("FreteOrigem"), ",", ".")
                Sql &= ", 0"    'Valor Auxiliar

                Sql &= ", ''" '& rs("ProdutoDestino") & "'"
                Sql &= ", " & rs("CodigoDeCusto_Id")
                Sql &= ", ''"
                Sql &= ")"

                Sql &= "  Update ApuracaoDeCustos set "
                Sql &= "         ValorDoProduto = " & Replace(rs("ValorOrigem"), ",", ".")
                Sql &= ",        ValorDoFrete = " & Replace(rs("FreteOrigem"), ",", ".")

                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("EmpresaDestino_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresaDestino_Id")
                Sql &= " And Deposito_Id = '" & rs("DepositoDestino_Id") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDepositoDestino_Id")
                Sql &= " And Ano_Id = " & rs("Ano_Id")
                Sql &= " And Mes_Id = " & rs("Mes_Id")
                Sql &= " And Produto_Id = '" & rs("ProdutoDerivado_Id") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("CodigoDestino")

                Sql &= " And EmpresaDestino_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresaDestino_Id = " & rs("EndEmpresa_Id")
                Sql &= " And DepositoDestino_Id = '" & rs("Deposito_Id") & "'"
                Sql &= " And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                Sql &= " And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
                Banco.GravaBanco(Sql)
            Next
        End If
    End Sub

    Sub ContabilizaCustos(ByVal pMes As Integer)
        Dim AnoC As Integer
        Dim sql As String
        Dim i As Integer


        AnoC = DdlAno.SelectedValue


        Dim Dia As String


        Dia = Format(CDate("01/" & pMes & "/" & AnoC).AddMonths(+1).AddDays(-1), "dd/MM/yyyy")
        Dia = Format(CDate(Dia), "yyyy-MM-dd")


        sql = " Delete Razao WHERE LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'" & vbCrLf
        sql &= "                And Lote_Id = 7000 "
        sql &= "                And Year(Movimento_Id)  = " & AnoC & vbCrLf
        sql &= "                And Month(Movimento_Id)  = " & pMes & "; " & vbCrLf

        ' Captura indiretos para dividir por proporção.
        sql &= "  WITH TotaisPorEmpresa																																	            "
        sql &= "  AS (																																								"
        sql &= "  	SELECT Razao.Empresa_Id																																			"
        sql &= "  		,SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS TotalEmpresa																							"
        sql &= "  	FROM Razao																																						"
        sql &= "  	INNER JOIN PlanoDeCustosXOrigem ON LEFT(Razao.Conta_Id, 7) = PlanoDeCustosXOrigem.Conta_Id																		"
        sql &= "  	WHERE LEFT(Razao.Empresa_id, 8) = '" & Left(Empresa(0), 8) & "'                                                                                                                         "
        sql &= "  		AND Month(Razao.Movimento_Id) = " & pMes & "                                                                                                                "
        sql &= "  		AND YEAR(Razao.Movimento_Id) = " & AnoC & "                                                                                                                 "
        sql &= "  		AND Razao.Lote_Id <> 7000																																	"
        sql &= "  		AND ISNULL(PlanoDeCustosXOrigem.Produto, '0') = 0																											"
        sql &= "  		AND PlanoDeCustosXOrigem.Codigo_Id = 402																													"
        sql &= "  	GROUP BY Razao.Empresa_Id																																		"
        sql &= "  	)																																								"
        sql &= "  SELECT R.Empresa_Id																																				"
        sql &= "  	,R.EndEmpresa_Id																																				"
        sql &= "  	,CONCAT (																																						"
        sql &= "  		P.Conta_Id																																					"
        sql &= "  		,'99'																																						"
        sql &= "  		) AS Conta_Id																																				"
        sql &= "  	,(100 / T.TotalEmpresa) * SUM(R.DebitoOficial - R.CreditoOficial) AS Percentual																					"
        sql &= "  INTO #tmpIndireto																																					"
        sql &= "  FROM Razao R																																						"
        sql &= "  INNER JOIN PlanoDeCustosXOrigem P ON LEFT(R.Conta_Id, 7) = P.Conta_Id																								"
        sql &= "  INNER JOIN TotaisPorEmpresa T ON R.Empresa_Id = T.Empresa_Id																										"
        sql &= "  WHERE LEFT(R.Empresa_id, 8) = '" & Left(Empresa(0), 8) & "'																																"
        sql &= "  	AND Month(R.Movimento_Id) = " & pMes & "                                                                                                                        "
        sql &= "  	AND YEAR(R.Movimento_Id) = " & AnoC & "                                                                                                                         "
        sql &= "  	AND R.Lote_Id <> 7000																																			"
        sql &= "  	AND ISNULL(P.Produto, '0') = 0																																	"
        sql &= "  	AND P.Codigo_Id = 402																																			"
        sql &= "  GROUP BY CONCAT (																																					"
        sql &= "  		P.Conta_Id																																					"
        sql &= "  		,'99'																																						"
        sql &= "  		)																																							"
        sql &= "  	,R.Empresa_Id																																					"
        sql &= "  	,R.EndEmpresa_Id																																				"
        sql &= "  	,T.TotalEmpresa																																					"
        sql &= "  																																									"
        'Obtem os valores com base na proporção do custo indireto
        sql &= "  SELECT AP.Empresa_Id																																				"
        sql &= "  	,AP.EndEmpresa_Id																																				"
        sql &= "  	,AP.Deposito_Id																																					"
        sql &= "  	,AP.EndDeposito_Id																																				"
        sql &= "  	,AP.Ano_Id																																						"
        sql &= "  	,AP.Mes_Id																																						"
        sql &= "  	,AP.Produto_Id																																					"
        sql &= "  	,AP.CodigoDeCusto_Id																																			"
        sql &= "  	,PC.Descricao AS TituloDoCusto																																	"
        sql &= "  	,PC.DebitoMercadoria																																			"
        sql &= "  	,ti.Conta_Id AS CreditoMercadoria																																"
        sql &= "  	,CASE 																																							"
        sql &= "  		WHEN HM.Descricao <> ''																																		"
        sql &= "  			THEN HM.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)	"
        sql &= "  		ELSE PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)		"
        sql &= "  		END AS HistoricoMercadoria																																	"
        sql &= "  	,PC.DebitoFrete																																					"
        sql &= "  	,PC.CreditoFrete																																				"
        sql &= "  	,CASE 																																							"
        sql &= "  		WHEN HF.Descricao <> ''																																		"
        sql &= "  			THEN HF.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)	"
        sql &= "  		ELSE PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)		"
        sql &= "  		END AS HistoricoFrete																																		"
        sql &= "  	,AP.EmpresaDestino_Id																																			"
        sql &= "  	,AP.EndEmpresaDestino_Id																																		"
        sql &= "  	,AP.DepositoDestino_Id																																			"
        sql &= "  	,AP.EndDepositoDestino_Id																																		"
        sql &= "  	,AP.ProdutoDerivado_Id																																			"
        sql &= "  	,AP.Quantidade																																					"
        sql &= "  	,ROUND(AP.ValorDoProduto / 100 * ti.Percentual, 2) AS ValorDoProduto																							"
        sql &= "  	,AP.ValorDoFrete																																				"
        sql &= "  FROM ApuracaoDeCustos AS AP																																		"
        sql &= "  INNER JOIN PlanoDeCustos AS PC ON AP.CodigoDeCusto_Id = PC.Codigo_Id																								"
        sql &= "  LEFT OUTER JOIN Historicos AS HF ON PC.HistoricoFrete = HF.Historico_Id																							"
        sql &= "  LEFT OUTER JOIN Historicos AS HM ON PC.HistoricoMercadoria = HM.Historico_Id																						"
        sql &= "  LEFT JOIN #tmpIndireto ti ON ti.Empresa_Id = ap.Empresa_Id																										"
        sql &= "  	AND ti.EndEmpresa_Id = ap.EndEmpresa_Id																															"
        sql &= "  WHERE LEFT(AP.Empresa_id, 8) ='" & Left(Empresa(0), 8) & "'                                                                                                                              "
        sql &= "  	AND AP.Ano_Id =  " & AnoC & "																																	"
        sql &= "  	AND AP.Mes_Id = " & pMes & "                                                                                                                                    "
        sql &= "  	AND ap.CodigoDeCusto_Id = 402																																	"
        sql &= "  																																									"
        sql &= "  UNION																																								"
        sql &= "  																																									"
        '  Obtem os demais lançamentos
        sql &= "  SELECT AP.Empresa_Id																																				"
        sql &= "  	,AP.EndEmpresa_Id																																				"
        sql &= "  	,AP.Deposito_Id																																					"
        sql &= "  	,AP.EndDeposito_Id																																				"
        sql &= "  	,AP.Ano_Id																																						"
        sql &= "  	,AP.Mes_Id																																						"
        sql &= "  	,AP.Produto_Id																																					"
        sql &= "  	,AP.CodigoDeCusto_Id																																			"
        sql &= "  	,PC.Descricao AS TituloDoCusto																																	"
        sql &= "  	,PC.DebitoMercadoria																																			"
        sql &= "  	,PC.CreditoMercadoria																																			"
        sql &= "  	,CASE 																																							"
        sql &= "  		WHEN HM.Descricao <> ''																																		"
        sql &= "  			THEN HM.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)	"
        sql &= "  		ELSE PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)		"
        sql &= "  		END AS HistoricoMercadoria																																	"
        sql &= "  	,PC.DebitoFrete																																					"
        sql &= "  	,PC.CreditoFrete																																				"
        sql &= "  	,CASE 																																							"
        sql &= "  		WHEN HF.Descricao <> ''																																		"
        sql &= "  			THEN HF.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)	"
        sql &= "  		ELSE PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)		"
        sql &= "  		END AS HistoricoFrete																																		"
        sql &= "  	,AP.EmpresaDestino_Id																																			"
        sql &= "  	,AP.EndEmpresaDestino_Id																																		"
        sql &= "  	,AP.DepositoDestino_Id																																			"
        sql &= "  	,AP.EndDepositoDestino_Id																																		"
        sql &= "  	,AP.ProdutoDerivado_Id																																			"
        sql &= "  	,AP.Quantidade																																					"
        sql &= "  	,AP.ValorDoProduto AS ValorDoProduto																															"
        sql &= "  	,AP.ValorDoFrete																																				"
        sql &= "  FROM ApuracaoDeCustos AS AP																																		"
        sql &= "  INNER JOIN PlanoDeCustos AS PC ON AP.CodigoDeCusto_Id = PC.Codigo_Id																								"
        sql &= "  LEFT OUTER JOIN Historicos AS HF ON PC.HistoricoFrete = HF.Historico_Id																							"
        sql &= "  LEFT OUTER JOIN Historicos AS HM ON PC.HistoricoMercadoria = HM.Historico_Id																						"
        sql &= "  WHERE LEFT(AP.Empresa_id, 8) ='" & Left(Empresa(0), 8) & "'                                                                                                                              "
        sql &= "  	AND AP.Ano_Id = " & AnoC & "                                                                                                                                    "
        sql &= "  	AND AP.Mes_Id = " & pMes & "                                                                                                                                    "
        sql &= "  	AND (																																							"
        sql &= "  		PC.DebitoMercadoria <> ''																																	"
        sql &= "  		OR PC.CreditoMercadoria <> ''																																"
        sql &= "  		)																																							"
        sql &= "  	AND ap.CodigoDeCusto_Id NOT IN (402, 507)																														"

        i = 0
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows

            If Dr("ValorDoProduto") > 0 Then
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, Dr("ValorDoProduto"), 0, 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, Dr("ValorDoProduto"), 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
            End If

            If Dr("ValorDoProduto") < 0 Then
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, (Dr("ValorDoProduto") * -1), 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, (Dr("ValorDoProduto") * -1), 0, 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
            End If



            If Dr("ValorDoFrete") > 0 Then
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, Dr("ValorDoFrete"), 0, Dr("MoedaValorDoFrete"), 0, Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, Dr("ValorDoFrete"), 0, Dr("MoedaValorDoFrete"), Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
            End If

            If Dr("ValorDoFrete") < 0 Then
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, (Dr("ValorDoFrete") * -1), 0, (Dr("MoedaValorDoFrete") * -1), Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
                i += 1
                sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, (Dr("ValorDoFrete") * -1), 0, (Dr("MoedaValorDoFrete") * -1), 0, Dr("HistoricoMercadoria"), 0, "R")
                If sql.Length > 0 Then
                    Array.Add(sql)
                End If
            End If



        Next
        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
            ChkResultado.Items.Add(New ListItem(" FIM - Contabilização Automática"))
        Else
            ChkResultado.Items.Add(New ListItem(" FIM - Erro no Processo de Contabilização Automática"))
        End If


    End Sub


    Function ContabilizaCentroCusto(ByVal pMes As Integer, Optional pLote As Integer = 7100) As Boolean
        Dim AnoC As Integer
        Dim sql As String
        Dim i As Integer


        AnoC = DdlAno.SelectedValue


        Dim Dia As String


        Dia = Format(CDate("01/" & pMes & "/" & AnoC).AddMonths(+1).AddDays(-1), "dd/MM/yyyy")
        Dia = Format(CDate(Dia), "yyyy-MM-dd")


        sql = " Delete Razao WHERE LEFT(Empresa_Id, 8) = '" & Left(Empresa(0), 8) & "'" & vbCrLf
        sql &= "                And Lote_Id = " & pLote & " "
        sql &= "                And Year(Movimento_Id)  = " & AnoC & vbCrLf
        sql &= "                And Month(Movimento_Id)  = " & pMes & "; " & vbCrLf

        ' Capturar os consumos materiais de produção para contabilização.
        sql &= " WITH CTE_ContabilizarCC																																											  "
        sql &= " AS (																																															  "
        sql &= " 	SELECT AP.Empresa_Id																																										  "
        sql &= " 		,AP.EndEmpresa_Id																																										  "
        sql &= " 		,AP.Deposito_Id																																											  "
        sql &= " 		,AP.EndDeposito_Id																																										  "
        sql &= " 		,AP.Ano_Id																																												  "
        sql &= " 		,AP.Mes_Id																																												  "
        sql &= " 		,AP.Produto_Id																																											  "
        sql &= " 		,AP.CodigoDeCusto_Id																																									  "
        sql &= " 		,PC.Descricao AS TituloDoCusto																																							  "
        sql &= " 		,ISNULL(PXCC.Conta_Id, '') AS DebitoMercadoria																																						  "
        sql &= " 		,ISNULL(PC.CreditoMercadoria, '') AS CreditoMercadoria																																				  "
        sql &= " 		,CASE 																																													  "
        sql &= " 			WHEN HM.Descricao <> ''																																								  "
        sql &= " 				THEN HM.Descricao + ' - ' + prod.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)	  "
        sql &= " 			ELSE PC.Descricao + ' - ' + prod.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)		  "
        sql &= " 			END AS HistoricoMercadoria																																							  "
        sql &= " 		,PC.DebitoFrete																																											  "
        sql &= " 		,PC.CreditoFrete																																										  "
        sql &= " 		,CASE 																																													  "
        sql &= " 			WHEN HF.Descricao <> ''																																								  "
        sql &= " 				THEN HF.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)							  "
        sql &= " 			ELSE PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(VARCHAR, AP.Mes_Id))) + CONVERT(VARCHAR, AP.Mes_Id) + '/' + convert(VARCHAR, AP.Ano_Id)								  "
        sql &= " 			END AS HistoricoFrete																																								  "
        sql &= " 		,AP.EmpresaDestino_Id																																									  "
        sql &= " 		,AP.EndEmpresaDestino_Id																																								  "
        sql &= " 		,AP.DepositoDestino_Id																																									  "
        sql &= " 		,AP.EndDepositoDestino_Id																																								  "
        sql &= " 		,AP.ProdutoDerivado_Id																																									  "
        sql &= " 		,prod.Saidas Quantidade																																									  "
        sql &= " 		,ROUND((AP.ValorDoProduto / AP.Quantidade) * prod.Saidas, 2) AS ValorDoProduto																											  "
        sql &= " 		,AP.ValorDoFrete																																										  "
        sql &= " 		,prod.CentroDeCusto_Id AS Custo																																							  "
        sql &= " 	FROM ApuracaoDeCustos AS AP																																									  "
        sql &= " 	INNER JOIN PlanoDeCustos AS PC ON AP.CodigoDeCusto_Id = PC.Codigo_Id																														  "
        sql &= " 	LEFT OUTER JOIN Historicos AS HF ON PC.HistoricoFrete = HF.Historico_Id																														  "
        sql &= " 	LEFT OUTER JOIN Historicos AS HM ON PC.HistoricoMercadoria = HM.Historico_Id																												  "
        sql &= " 	CROSS APPLY (																																												  "
        sql &= " 		SELECT DISTINCT pca.Conta_Id																																							  "
        sql &= " 			,Pro.Entradas																																										  "
        sql &= " 			,pro.Saidas																																											  "
        sql &= " 			,cc.CentroDeCusto_Id																																								  "
        sql &= " 			,cc.Descricao																																										  "
        sql &= " 		FROM Producao Pro																																										  "
        sql &= " 		INNER JOIN SubOperacoes SO ON SO.Operacao_Id = pro.Operacao_Id																															  "
        sql &= " 			AND SO.SubOperacoes_Id = pro.SubOperacao_Id																																			  "
        sql &= " 		INNER JOIN CentrosDeCustos CC ON CC.CentroDeCusto_Id = pro.CentroDeCusto																												  "
        sql &= " 		INNER JOIN PlanoDeContas PCA ON PCA.Conta_Id = SO.GrupoDeContas																															  "
        sql &= " 			AND PCA.CentroDeCusto = 'S'																																							  "
        sql &= " 		WHERE Pro.Empresa_Id = ap.Empresa_Id																																					  "
        sql &= " 			AND pro.EndEmpresa_Id = ap.EndEmpresa_Id																																			  "
        sql &= " 			AND pro.Produto_Id = ap.Produto_Id																																					  "
        sql &= " 			AND ap.Ano_Id = YEAR(pro.Movimento_Id)																																				  "
        sql &= " 			AND ap.Mes_Id = MONTH(pro.Movimento_Id)																																				  "
        sql &= " 			AND pro.FisicoFiscal_Id = 1																																							  "
        sql &= " 			AND Pro.CentroDeCusto > 0																																							  "
        sql &= " 			AND cc.Ativo = 1																																									  "
        sql &= " 		) AS prod																																												  "
        sql &= " 	OUTER APPLY (																																												  "
        sql &= " 		SELECT DISTINCT pxcc.conta_Id																																							  "
        sql &= " 		FROM ProdutoXCentrosDeCustos PXCC																																						  "
        sql &= " 		WHERE PXCC.Produto_Id = AP.Produto_Id																																					  "
        sql &= " 			AND PXCC.CentroDeCusto_Id = prod.CentroDeCusto_Id																																	  "
        sql &= " 		) AS pxcc																																												  "
        sql &= " 	WHERE LEFT(AP.Empresa_id, 8)  = '" & Left(Empresa(0), 8) & "'																																						  "
        sql &= " 		AND AP.Ano_Id = " & AnoC & "																																						      "
        sql &= " 		AND AP.Mes_Id = " & pMes & "																																							  "
        sql &= " 		AND AP.CodigoDeCusto_Id = 507																																							  "
        sql &= " 	)																																															  "
        sql &= " /*débito*/																																														  "
        sql &= " SELECT Empresa_Id																																												  "
        sql &= " 	,EndEmpresa_Id																																												  "
        sql &= " 	,Deposito_Id																																												  "
        sql &= " 	,EndDeposito_Id																																												  "
        sql &= " 	,Ano_Id																																														  "
        sql &= " 	,Mes_Id																																														  "
        sql &= " 	,Produto_Id																																													  "
        sql &= " 	,CodigoDeCusto_Id																																											  "
        sql &= " 	,TituloDoCusto																																												  "
        sql &= " 	,DebitoMercadoria																																											  "
        sql &= " 	,'' AS CreditoMercadoria																																									  "
        sql &= " 	,HistoricoMercadoria																																										  "
        sql &= " 	,DebitoFrete																																												  "
        sql &= " 	,CreditoFrete																																												  "
        sql &= " 	,HistoricoFrete																																												  "
        sql &= " 	,EmpresaDestino_Id																																											  "
        sql &= " 	,EndEmpresaDestino_Id																																										  "
        sql &= " 	,DepositoDestino_Id																																											  "
        sql &= " 	,EndDepositoDestino_Id																																										  "
        sql &= " 	,ProdutoDerivado_Id																																											  "
        sql &= " 	,Quantidade																																													  "
        sql &= " 	,ValorDoProduto																																												  "
        sql &= " 	,ValorDoFrete																																												  "
        sql &= " 	,Custo																																														  "
        sql &= " FROM CTE_ContabilizarCC																																										  "
        sql &= " 																																																  "
        sql &= " UNION																																															  "
        sql &= " 																																																  "
        sql &= " /*Crédito*/																																													  "
        sql &= " SELECT Empresa_Id																																												  "
        sql &= " 	,EndEmpresa_Id																																												  "
        sql &= " 	,Deposito_Id																																												  "
        sql &= " 	,EndDeposito_Id																																												  "
        sql &= " 	,Ano_Id																																														  "
        sql &= " 	,Mes_Id																																														  "
        sql &= " 	,Produto_Id																																													  "
        sql &= " 	,CodigoDeCusto_Id																																											  "
        sql &= " 	,TituloDoCusto																																												  "
        sql &= " 	,'' AS DebitoMercadoria																																										  "
        sql &= " 	,CreditoMercadoria																																											  "
        sql &= " 	,HistoricoMercadoria																																										  "
        sql &= " 	,DebitoFrete																																												  "
        sql &= " 	,CreditoFrete																																												  "
        sql &= " 	,HistoricoFrete																																												  "
        sql &= " 	,EmpresaDestino_Id																																											  "
        sql &= " 	,EndEmpresaDestino_Id																																										  "
        sql &= " 	,DepositoDestino_Id																																											  "
        sql &= " 	,EndDepositoDestino_Id																																										  "
        sql &= " 	,ProdutoDerivado_Id																																											  "
        sql &= " 	,Quantidade																																													  "
        sql &= " 	,ValorDoProduto																																												  "
        sql &= " 	,ValorDoFrete																																												  "
        sql &= " 	,Custo																																														  "
        sql &= " FROM CTE_ContabilizarCC																																									      "

        i = 0
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows

            Try

                If Dr("ValorDoProduto") > 0 Then
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("DebitoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, Dr("ValorDoProduto"), 0, 0, 0, Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("CreditoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, Dr("ValorDoProduto"), 0, 0, Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                End If

                If Dr("ValorDoProduto") < 0 Then
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("DebitoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, (Dr("ValorDoProduto") * -1), 0, 0, Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("CreditoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, (Dr("ValorDoProduto") * -1), 0, 0, 0, Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                End If

                If Dr("ValorDoFrete") > 0 Then
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("DebitoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, Dr("ValorDoFrete"), 0, Dr("MoedaValorDoFrete"), 0, Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("CreditoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, Dr("ValorDoFrete"), 0, Dr("MoedaValorDoFrete"), Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                End If

                If Dr("ValorDoFrete") < 0 Then
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("DebitoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, (Dr("ValorDoFrete") * -1), 0, (Dr("MoedaValorDoFrete") * -1), Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                    i += 1
                    sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, pLote, i, Dr("CreditoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, (Dr("ValorDoFrete") * -1), 0, (Dr("MoedaValorDoFrete") * -1), 0, Dr("HistoricoMercadoria"), 0, "R", Dr("Custo"))
                    If sql.Length > 0 Then
                        Array.Add(sql)
                    End If
                End If

            Catch ex As Exception
                ChkResultado.Items.Add(New ListItem(String.Format(" Erro no Processo de Contabilização centro de custos. ERRO: {0}", ex.Message)))
                Return False
            End Try

        Next

        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
            Return True
        Else
            ChkResultado.Items.Add(New ListItem(" FIM - Erro no Processo de Contabilização centro de custos"))
            Return False
        End If

    End Function

    Function SqlRazao(ByVal CnpjEmpresa As String, ByVal EndEmpresa As Integer, ByVal DataMovimento As String, ByVal Lote As String, ByVal Sequencia As String, ByVal Conta As String, ByVal Cliente As String, ByVal EndCliente As Integer, ByVal Produto As String, ByVal Indexador As String, ByVal DataMoeda As String, ByVal DebitoOficial As Double, ByVal CreditoOficial As Double, ByVal DebitoMoeda As Double, ByVal CreditoMoeda As Double, ByVal Historico As String, ByVal Moeda As String, ByVal PrevistoRealizado As String, Optional cCusto As Integer = 0) As String
        Dim TemCliente As String = "N"
        Dim TemProduto As String = "N"
        Dim sqll As String
        Dim sql As String
        sql = ""

        sqll = "Select Top 1 * from PlanoDeContas where Left(Conta_Id, 7) = '" & Microsoft.VisualBasic.Left(Conta, 7) & "'"

        For Each dr As DataRow In Banco.ConsultaDataSet(sqll, "Consulta").Tables(0).Rows
            TemCliente = dr("Cliente")
            TemProduto = dr("Produto")


            sql = " INSERT INTO Razao (Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Produto,"
            sql &= " Movimento_Id,Lote_Id,Sequencia_Id,Indexador, DataMoeda,"
            sql &= " DebitoOficial,CreditoOficial,DebitoMoeda,CreditoMoeda,Historico,PrevistoRealizado,Custo) Values ("
            sql &= "  '" & CnpjEmpresa & "'"
            sql &= ",  " & EndEmpresa
            sql &= ", '" & Conta & "'"

            If TemCliente = "S" Then
                sql &= ", '" & Cliente & "'"
                sql &= ", " & CInt(EndCliente)
            Else
                sql &= ", ''"
                sql &= ", 0"
            End If

            If TemProduto = "S" Then
                sql &= ", '" & Produto & "'"
            Else
                sql &= ", ''"
            End If

            sql &= ", '" & DataMovimento.ToSqlDate() & "'"
            sql &= ", " & CInt(Lote)
            sql &= ", " & CInt(Sequencia)
            sql &= ", 3"
            sql &= ", '" & DataMoeda.ToSqlDate() & "'"
            sql &= ", " & DebitoOficial.ToString.Replace(",", ".")
            sql &= ", " & CreditoOficial.ToString.Replace(",", ".")
            sql &= ", " & DebitoMoeda.ToString.Replace(",", ".")
            sql &= ", " & CreditoMoeda.ToString.Replace(",", ".")
            sql &= ", '" & UCase(RTrim(Historico)) & "'"
            sql &= ", '" & PrevistoRealizado & "'"
            sql &= ", '" & cCusto.ToString() & "')"
        Next

        Return sql

    End Function

    Private Sub Contabiliza(ByVal pMes As Integer)
        Dim AnoC As Integer
        AnoC = DdlAno.SelectedValue
        Dim Array As New ArrayList
        Dim Sql As String

        Sql = " if(object_id('tempdb..#Temp') is not null) begin drop table #Temp end; " & vbCrLf &
              " Delete Razao where lote_id='7000' and movimento_id='" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & " ' AND LEFT(Empresa_id,8)='" & Left(Empresa(0), 8) & "'; " & vbCrLf &
              " SELECT * " & vbCrLf &
              "   into #temp" & vbCrLf &
              "   from (  " & vbCrLf &
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf &
              "                 PlanoDeCustos.DebitoMercadoria as Conta_Id, " & vbCrLf &
              "                 case " & vbCrLf &
              "                   when len(DebitoMercadoria) = 7    " & vbCrLf &
              "                     then EmpresaDestino_Id " & vbCrLf &
              "                     else ''  " & vbCrLf &
              "                 end Cliente_Id,  " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(DebitoMercadoria) = 7  " & vbCrLf &
              "                     then EndEmpresaDestino_Id " & vbCrLf &
              "                     else 0 " & vbCrLf &
              "                 end EndCliente_Id,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf &
              "                 '7000' as  Lote, " & vbCrLf &
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf &
              "                 0 as Titulo, " & vbCrLf &
              "                 3 as Indexador,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoProduto AS DebitoOficial,  " & vbCrLf &
              "                 0.00 as CreditoOficial, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoProduto AS DebitoMoeda,  " & vbCrLf &
              "                 0.00 as CreditoMoeda, " & vbCrLf &
              "                 convert(nvarchar,PlanoDeCustos.HistoricoMercadoria) +' '+ HMercadoria.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf &
              "                 0 as CentroDeCustos, " & vbCrLf &
              "                 'P' as PrevistoRealizado, " & vbCrLf &
              "                 0 as Serie_Nf, " & vbCrLf &
              "                 0 as Numero_Nf, " & vbCrLf &
              "                 NULL as Pedido, " & vbCrLf &
              "                 'E' as EntradaSaida_Nf " & vbCrLf &
              "            FROM ApuracaoDeCustos " & vbCrLf &
              "  		  INNER JOIN Clientes AS Empresa " & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf &
              " 		  INNER JOIN Clientes AS Deposito " & vbCrLf &
              "              ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id " & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id   " & vbCrLf &
              "            LEFT JOIN Clientes AS Destino " & vbCrLf &
              "              ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id    " & vbCrLf &
              "  		  INNER JOIN Produtos " & vbCrLf &
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf &
              "           INNER JOIN PlanoDeCustos" & vbCrLf &
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf &
              "            LEFT JOIN TabelaDePrecosDeMercado" & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id  " & vbCrLf &
              "             AND month(Data_Id)                  = '" & pMes & "' AND year(Data_Id)= '" & AnoC & "'" & vbCrLf &
              "            LEFT JOIN Historicos AS HMercadoria" & vbCrLf &
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf &
              "            LEFT JOIN Historicos AS HFrete " & vbCrLf &
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf &
              "           WHERE ApuracaoDeCustos.Ano_Id             = '" & AnoC & "'" & vbCrLf &
              "             AND ApuracaoDeCustos.Mes_Id             = '" & pMes & "'" & vbCrLf &
              "             And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf &
              "    UNION ALL" & vbCrLf &
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf &
              "                 PlanoDeCustos.CreditoMercadoria as Conta_Id, " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(CreditoMercadoria) = 7 " & vbCrLf &
              "                     then EmpresaDestino_Id " & vbCrLf &
              "                     else ''  " & vbCrLf &
              "                 end Cliente_Id,  " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(CreditoMercadoria) = 7  " & vbCrLf &
              "                     then EndEmpresaDestino_Id " & vbCrLf &
              "                     else 0 " & vbCrLf &
              "                 end EndCliente_Id,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf &
              "                 '7000' as  Lote, " & vbCrLf &
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf &
              "                 0 as Titulo, " & vbCrLf &
              "                 3 as Indexador,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf &
              "                 0.00 as DebitoOficial, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoProduto AS CreditoOficial,  " & vbCrLf &
              "                 0.00 as DebitoMoeda, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoProduto AS CreditoMoeda,  " & vbCrLf &
              "                 convert(nvarchar,PlanoDeCustos.HistoricoMercadoria) +' '+ HMercadoria.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf &
              "                 0 as CentroDeCustos, " & vbCrLf &
              "                 'P' as PrevistoRealizado, " & vbCrLf &
              "                 0 as Serie_Nf, " & vbCrLf &
              "                 0 as Numero_Nf, " & vbCrLf &
              "                 NULL as Pedido, " & vbCrLf &
              "                 'S' as EntradaSaida_Nf " & vbCrLf &
              "            FROM ApuracaoDeCustos   " & vbCrLf &
              "  		  INNER JOIN Clientes AS Empresa" & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id   " & vbCrLf &
              "  		  INNER JOIN Clientes AS Deposito" & vbCrLf &
              "              ON ApuracaoDeCustos.Deposito_Id    = Deposito.Cliente_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id   " & vbCrLf &
              "            LEFT JOIN Clientes AS Destino" & vbCrLf &
              "              ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id    " & vbCrLf &
              "  		  INNER JOIN Produtos " & vbCrLf &
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf &
              "           INNER JOIN PlanoDeCustos" & vbCrLf &
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf &
              "            Left JOIN TabelaDePrecosDeMercado" & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id  " & vbCrLf &
              "             AND month(Data_Id)                  = " & pMes & " AND year(Data_Id)= '" & AnoC & "'  " & vbCrLf &
              "            LEFT JOIN Historicos AS HMercadoria" & vbCrLf &
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf &
              "            LEFT JOIN Historicos AS HFrete" & vbCrLf &
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf &
              "           WHERE ApuracaoDeCustos.Ano_Id             = '" & AnoC & "'" & vbCrLf &
              "             AND ApuracaoDeCustos.Mes_Id             = '" & pMes & "'" & vbCrLf &
              "             And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf &
              "           Union ALL" & vbCrLf &
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf &
              "                 PlanoDeCustos.DebitoFrete as Conta_Id, " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(DebitoFrete) = 7    " & vbCrLf &
              "                     then EmpresaDestino_Id " & vbCrLf &
              "                     else ''  " & vbCrLf &
              "                 end Cliente_Id,  " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(DebitoFrete) = 7  " & vbCrLf &
              "                     then EndEmpresaDestino_Id " & vbCrLf &
              "                     else 0 " & vbCrLf &
              "                 end EndCliente_Id,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf &
              "                 '7000' as  Lote, " & vbCrLf &
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf &
              "                 0 as Titulo, " & vbCrLf &
              "                 3 as Indexador,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoFrete AS DebitoOficial,  " & vbCrLf &
              "                 0.00 as CreditoOficial, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoFrete AS DebitoMoeda,  " & vbCrLf &
              "                 0.00 as CreditoMoeda, " & vbCrLf &
              "                 convert(nvarchar,PlanoDeCustos.HistoricoFrete) +' '+ HFrete.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf &
              "                 0 as CentroDeCustos, " & vbCrLf &
              "                 'P' as PrevistoRealizado, " & vbCrLf &
              "                 0 as Serie_Nf, " & vbCrLf &
              "                 0 as Numero_Nf, " & vbCrLf &
              "                 NULL as Pedido, " & vbCrLf &
              "                 'E' as EntradaSaida_Nf " & vbCrLf &
              "            FROM ApuracaoDeCustos   " & vbCrLf &
              "  		  INNER JOIN Clientes AS Empresa " & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf &
              "  	  	  INNER JOIN Clientes AS Deposito " & vbCrLf &
              "              ON ApuracaoDeCustos.Deposito_Id    = Deposito.Cliente_Id" & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id" & vbCrLf &
              "            LEFT JOIN Clientes AS Destino " & vbCrLf &
              "              ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id" & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id" & vbCrLf &
              "  	  	  INNER JOIN Produtos" & vbCrLf &
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf &
              "           INNER JOIN PlanoDeCustos" & vbCrLf &
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf &
              "            Left JOIN TabelaDePrecosDeMercado " & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id  " & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id   " & vbCrLf &
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id  " & vbCrLf &
              "             AND month(Data_Id)                  = '" & pMes & "' AND year(Data_Id)= '" & AnoC & "'  " & vbCrLf &
              "            LEFT JOIN Historicos AS HMercadoria " & vbCrLf &
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf &
              "            LEFT JOIN Historicos AS HFrete" & vbCrLf &
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf &
              "           WHERE ApuracaoDeCustos.Ano_Id = '" & AnoC & "'" & vbCrLf &
              "             AND ApuracaoDeCustos.Mes_Id = '" & pMes & "'" & vbCrLf &
              "             AND left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf &
              "           UNION ALL " & vbCrLf &
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf &
              "                 PlanoDeCustos.CreditoFrete as Conta_Id, " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(CreditoFrete) = 7    " & vbCrLf &
              "                     then EmpresaDestino_Id " & vbCrLf &
              "                     else ''  " & vbCrLf &
              "                 end Cliente_Id,  " & vbCrLf &
              "                 case  " & vbCrLf &
              "                   when len(CreditoFrete) = 7  " & vbCrLf &
              "                     then EndEmpresaDestino_Id " & vbCrLf &
              "                     else 0 " & vbCrLf &
              "                 end EndCliente_Id,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf &
              "                 '7000' as  Lote, " & vbCrLf &
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf &
              "                 0 as Titulo, " & vbCrLf &
              "                 3 as Indexador,  " & vbCrLf &
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf &
              "                 0.00 as DebitoOficial, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoFrete AS CreditoOficial, " & vbCrLf &
              "                 0.00 as DebitoMoeda, " & vbCrLf &
              "                 ApuracaoDeCustos.ValorDoFrete AS CreditoMoeda,  " & vbCrLf &
              "                 convert(nvarchar,PlanoDeCustos.HistoricoFrete) +' '+ HFrete.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf &
              "                 0 as CentroDeCustos, " & vbCrLf &
              "                 'P' as PrevistoRealizado, " & vbCrLf &
              "                 0 as Serie_Nf, " & vbCrLf &
              "                 0 as Numero_Nf, " & vbCrLf &
              "                 NULL as Pedido, " & vbCrLf &
              "                 'S' as EntradaSaida_Nf " & vbCrLf &
              "            FROM ApuracaoDeCustos " & vbCrLf &
              "  	 	  INNER JOIN Clientes AS Empresa " & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id = Empresa.Cliente_Id " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf &
              "  		  INNER JOIN Clientes AS Deposito " & vbCrLf &
              "              ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id " & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id " & vbCrLf &
              "            LEFT JOIN Clientes AS Destino " & vbCrLf &
              "              ON ApuracaoDeCustos.EmpresaDestino_Id = Destino.Cliente_Id " & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id " & vbCrLf &
              "  		  INNER JOIN Produtos  " & vbCrLf &
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf &
              "           INNER JOIN PlanoDeCustos " & vbCrLf &
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf &
              "            Left JOIN TabelaDePrecosDeMercado " & vbCrLf &
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id" & vbCrLf &
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id" & vbCrLf &
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id" & vbCrLf &
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id" & vbCrLf &
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id" & vbCrLf &
              "             AND month(Data_Id)                  = '" & pMes & "' AND year(Data_Id)= '" & AnoC & "'" & vbCrLf &
              "            LEFT JOIN Historicos AS HMercadoria " & vbCrLf &
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf &
              "            LEFT JOIN Historicos AS HFrete " & vbCrLf &
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf &
              "           WHERE ApuracaoDeCustos.Ano_Id              = '" & AnoC & "' AND ApuracaoDeCustos.Mes_Id = '" & pMes & "'" & vbCrLf &
              "              And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf &
              "        ) as consulta; " & vbCrLf &
              "  Delete #temp " & vbCrLf &
              "  Where Conta_Id  = ''" & vbCrLf &
              "    OR (    DebitoOficial  = 0" & vbCrLf &
              "        AND CreditoOficial = 0" & vbCrLf &
              "        AND DebitoMoeda    = 0" & vbCrLf &
              "        AND CreditoMoeda   = 0" & vbCrLf &
              "       ) ; " & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              " Update #Temp  " & vbCrLf &
              "    Set Produto_Id = ''" & vbCrLf &
              "   From #temp " & vbCrLf &
              "   Left join PlanoDeContas PC  " & vbCrLf &
              "     on #Temp.Empresa_Id    = PC.Empresa_Id  " & vbCrLf &
              "    And #Temp.EndEmpresa_Id = PC.EndEmpresa_Id  " & vbCrLf &
              "    AND #Temp.Conta_Id      = PC.Conta_Id " & vbCrLf &
              "  Where PC.Produto = 'N'; " & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "Update #Temp" & vbCrLf &
              "    Set debitooficial = creditooficial * -1," & vbCrLf &
              "        debitomoeda   = creditomoeda * -1" & vbCrLf &
              "where creditooficial < 0" & vbCrLf &
              "" & vbCrLf &
              "Update #Temp" & vbCrLf &
              "    Set creditooficial = 0," & vbCrLf &
              "        creditomoeda   = 0" & vbCrLf &
              "where creditooficial < 0" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "Update #Temp" & vbCrLf &
              "    Set creditooficial = debitooficial * -1," & vbCrLf &
              "        creditomoeda   = debitomoeda * -1" & vbCrLf &
              "where debitooficial < 0" & vbCrLf &
              "" & vbCrLf &
              "Update #Temp" & vbCrLf &
              "    Set debitooficial = 0," & vbCrLf &
              "        debitomoeda   = 0" & vbCrLf &
              "where debitooficial < 0" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "  INSERT INTO Razao " & vbCrLf &
              "          (Empresa_Id,  " & vbCrLf &
              "           EndEmpresa_Id,  " & vbCrLf &
              "           Conta_Id, " & vbCrLf &
              "           Cliente_Id,  " & vbCrLf &
              "           EndCliente_Id,  " & vbCrLf &
              "           Movimento_Id,  " & vbCrLf &
              "           Lote_Id,  " & vbCrLf &
              "           Sequencia_Id,  " & vbCrLf &
              "           Produto,  " & vbCrLf &
              "           Titulo,  " & vbCrLf &
              "           Indexador, " & vbCrLf &
              "           DataMoeda,  " & vbCrLf &
              "           DebitoOficial,  " & vbCrLf &
              "           CreditoOficial,  " & vbCrLf &
              "           DebitoMoeda,  " & vbCrLf &
              "           CreditoMoeda, " & vbCrLf &
              "           Historico,  " & vbCrLf &
              "           custo,  " & vbCrLf &
              "           PrevistoRealizado,Serie_Nf,Numero_Nf,Pedido,EntradaSaida_Nf)  " & vbCrLf &
              "           (  " & vbCrLf &
              "            SELECT " & vbCrLf &
              "            Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Movimento_Id,Lote,  " & vbCrLf &
              "            ROW_NUMBER() OVER (PARTITION BY Movimento_Id,Lote order by Movimento_Id) AS Sequencia_Id,  " & vbCrLf &
              "            Produto_Id, Titulo,Indexador,DataMoeda,abs(DebitoOficial),abs(CreditoOficial), " & vbCrLf &
              "            abs(DebitoMoeda),abs(CreditoMoeda),  " & vbCrLf &
              "            Historico,CentroDeCustos,PrevistoRealizado, Serie_Nf, Numero_Nf, Pedido, EntradaSaida_Nf  " & vbCrLf &
              "            from  #temp )" & vbCrLf

        Array.Add(Sql)

        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
            ChkResultado.Items.Add(New ListItem(" FIM - Contabilização Automática"))
        Else
            ChkResultado.Items.Add(New ListItem(" FIM - Erro no Processo de Contabilização Automática"))
        End If

    End Sub

    Public Function Transferencias(ByVal pMes As Integer) As Boolean
        Dim erros As Integer = 0
        Dim i = 0

        Dim Sql As String = "SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito AS Deposito_Id, NotasFiscaisXItens.EndDeposito AS EndDeposito_Id," & vbCrLf &
              "       NotasFiscais.Cliente_Id AS EmpresaDestino_Id, NotasFiscais.EndCliente_Id AS EndEmpresaDestino_Id, NotasFiscais.Cliente_Id AS DepositoDestino_Id," & vbCrLf &
              "       NotasFiscais.EndCliente_Id AS EndDepositoDestino_Id," & vbCrLf &
              "       NotasFiscaisXItens.Produto_Id, " & vbCrLf &
              "       SubOperacoes.ApuracaoDeCustos AS Placus_Id," & vbCrLf &
              "       SubOperacoes.ApuracaodeCustosContraPartida AS PlacusContra_Id," & vbCrLf &
              "       SUM(IIF(Produtos.PesoQuantidade = 'Q', NotasFiscaisXItens.QuantidadeFiscal, NotasFiscaisXItens.PesoFiscal)) AS Quantidade," & vbCrLf &
              "       convert(numeric(18,2),0) AS ValorDoProduto" & vbCrLf &
              "  FROM NotasFiscais" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf &
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              " INNER JOIN SubOperacoes" & vbCrLf &
              "    ON SubOperacoes.Operacao_Id = NotasFiscaisXItens.Operacao  " & vbCrLf &
              "   AND SubOperacoes.SubOperacoes_Id = NotasFiscaisXItens.SubOperacao" & vbCrLf &
              " INNER JOIN PlanoDeCustos " & vbCrLf &
              "    ON PlanoDeCustos.Codigo_Id = SubOperacoes.ApuracaoDeCustos " & vbCrLf &
              " INNER JOIN Produtos " & vbCrLf &
              "    ON Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id " & vbCrLf &
              " INNER JOIN GruposDeEstoques" & vbCrLf &
              "    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf &
              " WHERE LEFT(NotasFiscais.Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'"

        Sql &= "   AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue & vbCrLf &
               "   AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes & vbCrLf

        Sql &= "   AND (SubOperacoes.ApuracaoDeCustosContraPartida > 0)" & vbCrLf &
               "   And  NotasFiscais.Situacao  = 1" & vbCrLf &
               "   And  GruposDeEstoques.custo = 1" & vbCrLf &
               "   And  PlanoDeCustos.Classe   ='" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf &
               " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito," & vbCrLf &
               "          NotasFiscaisXItens.EndDeposito, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), " & vbCrLf &
               "          NotasFiscaisXItens.Produto_Id," & vbCrLf &
               "         SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro," & vbCrLf &
               "         NotasFiscaisXItens.EndDepositoTerceiro , SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare" & vbCrLf &
                  " @Exist as varchar(1)" & vbCrLf &
                  " set @Exist = (select case " & vbCrLf &
                  "                        when exists (" & vbCrLf &
                  "                                      select 1 " & vbCrLf &
                  "                                        from ApuracaoDeCustos " & vbCrLf &
                  "                                       Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf &
                  "                                         And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf &
                  "                                         And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf &
                  "                                         And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf &
                  "                                         And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf &
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf &
                  "                                         And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf &
                  "                                         And Mes_Id                = " & pMes & vbCrLf &
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf &
                  "                                         And CodigoDeCusto_Id      = " & Dr("Placus_Id") & vbCrLf &
                  "                                    )" & vbCrLf &
                  "                           then 'S'" & vbCrLf &
                  "                           else 'N'" & vbCrLf &
                  "                       end) ;" & vbCrLf &
                  " if @Exist = 'N' " & vbCrLf &
                  "  INSERT INTO ApuracaoDeCustos ( " & vbCrLf &
                  "  Empresa_Id" & vbCrLf &
                  ", EndEmpresa_Id" & vbCrLf &
                  ", Deposito_Id" & vbCrLf &
                  ", EndDeposito_Id" & vbCrLf &
                  ", Ano_Id" & vbCrLf &
                  ", Mes_Id" & vbCrLf &
                  ", Produto_Id" & vbCrLf &
                  ", CodigoDeCusto_Id" & vbCrLf &
                  ", EmpresaDestino_Id" & vbCrLf &
                  ", EndEmpresaDestino_Id" & vbCrLf &
                  ", DepositoDestino_Id" & vbCrLf &
                  ", EndDepositoDestino_Id" & vbCrLf &
                  ", ProdutoDerivado_Id" & vbCrLf &
                  ", Quantidade" & vbCrLf &
                  ", ValorDoProduto" & vbCrLf &
                  ", ValorDoFrete" & vbCrLf &
                  ", ValorAuxiliar" & vbCrLf &
                  ", ProdutoDestino" & vbCrLf &
                  ", CodigoDestino)" & vbCrLf &
                  " VALUES('" & Dr("Empresa_Id") & "'" & vbCrLf &
                  ", " & Dr("EndEmpresa_Id") & vbCrLf &
                  ",'" & Dr("Deposito_Id") & "'" & vbCrLf &
                  ", " & Dr("EndDeposito_Id") & vbCrLf &
                  ", " & DdlAno.SelectedValue & vbCrLf &
                  ", " & pMes & vbCrLf &
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf &
                  ", " & Dr("Placus_Id") & vbCrLf &
                  ",'" & Dr("EmpresaDestino_Id") & "'" & vbCrLf &
                  ", " & Dr("EndEmpresaDestino_Id") & vbCrLf &
                  ",'" & Dr("DepositoDestino_Id") & "'" & vbCrLf &
                  ", " & Dr("EndDepositoDestino_Id") & vbCrLf &
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf &
                  ", " & Str(Dr("Quantidade")) & vbCrLf &
                  ", 0" & vbCrLf &
                  ", 0" & vbCrLf &
                  ", 0" & vbCrLf &
                  ",''" & vbCrLf &
                  ", " & Dr("PlacusContra_Id") & vbCrLf &
                  ")" & vbCrLf &
                  " Else" & vbCrLf &
                  "   Update ApuracaoDeCustos set " & vbCrLf &
                  "   Quantidade = " & Str(Dr("Quantidade")) & vbCrLf &
                  " Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf &
                  "   And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf &
                  "   And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf &
                  "   And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf &
                  "   And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf &
                  "   And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf &
                  "   And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf &
                  "   And Mes_Id                = " & pMes & vbCrLf &
                  "   And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf &
                  "   And CodigoDeCusto_Id      = " & Dr("Placus_Id") & vbCrLf

            Sqls.Add(Sql)
            If Not Banco.GravaBanco(Sqls) Then Return False
        Next
        Return True
    End Function

    'aLTERADO 14/08/2025 - Erro na apuração de custo Baxi - transferencia
    Public Function Transferencias_bkp(ByVal pMes As Integer) As Boolean
        Dim erros As Integer = 0
        Dim i = 0

        Dim Sql As String = "SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito AS Deposito_Id, NotasFiscaisXItens.EndDeposito AS EndDeposito_Id," & vbCrLf &
              "       NotasFiscais.Cliente_Id AS EmpresaDestino_Id, NotasFiscais.EndCliente_Id AS EndEmpresaDestino_Id, NotasFiscais.Cliente_Id AS DepositoDestino_Id," & vbCrLf &
              "       NotasFiscais.EndCliente_Id AS EndDepositoDestino_Id," & vbCrLf &
              "       case" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '101010001' then '101010007'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '101010003' then '101010007'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '101080002' then '101080001'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '101060002' then '101060001'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '701010004' then '701010001'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '404010002' then '404010001'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '404010003' then '404010001'" & vbCrLf &
              "         when NotasFiscaisXItens.Produto_Id = '404010004' then '404010001'" & vbCrLf &
              "         else NotasFiscaisXItens.Produto_Id" & vbCrLf &
              "       end Produto_id," & vbCrLf &
              "       SubOperacoes.ApuracaoDeCustos AS Placus_Id," & vbCrLf &
              "       SubOperacoes.ApuracaodeCustosContraPartida AS PlacusContra_Id," & vbCrLf &
              "       SUM(IIF(Produtos.PesoQuantidade = 'Q', NotasFiscaisXItens.QuantidadeFiscal, NotasFiscaisXItens.PesoFiscal)) AS Quantidade," & vbCrLf &
              "       convert(numeric(18,2),0) AS ValorDoProduto" & vbCrLf &
              "  FROM NotasFiscais" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf &
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              " INNER JOIN SubOperacoes" & vbCrLf &
              "    ON SubOperacoes.Operacao_Id = NotasFiscaisXItens.Operacao  " & vbCrLf &
              "   AND SubOperacoes.SubOperacoes_Id = NotasFiscaisXItens.SubOperacao" & vbCrLf &
              " INNER JOIN PlanoDeCustos " & vbCrLf &
              "    ON PlanoDeCustos.Codigo_Id = SubOperacoes.ApuracaoDeCustos " & vbCrLf &
              " INNER JOIN Produtos " & vbCrLf &
              "    ON Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id " & vbCrLf &
              " INNER JOIN GruposDeEstoques" & vbCrLf &
              "    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf &
              " WHERE LEFT(NotasFiscais.Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'"

        Sql &= "   AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue & vbCrLf &
               "   AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes & vbCrLf

        Sql &= "   AND (SubOperacoes.ApuracaoDeCustosContraPartida > 0)" & vbCrLf &
               "   And  NotasFiscais.Situacao  = 1" & vbCrLf &
               "   And  GruposDeEstoques.custo = 1" & vbCrLf &
               "   And  PlanoDeCustos.Classe   ='" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf &
               " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito," & vbCrLf &
               "          NotasFiscaisXItens.EndDeposito, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), " & vbCrLf &
               "          case" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '101010001' then '101010007'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '101010003' then '101010007'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '101080002' then '101080001'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '101060002' then '101060001'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '701010004' then '701010001'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '404010002' then '404010001'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '404010003' then '404010001'" & vbCrLf &
               "            when NotasFiscaisXItens.Produto_Id = '404010004' then '404010001'" & vbCrLf &
               "            else NotasFiscaisXItens.Produto_Id" & vbCrLf &
               "         end," & vbCrLf &
               "         SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro," & vbCrLf &
               "         NotasFiscaisXItens.EndDepositoTerceiro , SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare" & vbCrLf &
                  " @Exist as varchar(1)" & vbCrLf &
                  " set @Exist = (select case " & vbCrLf &
                  "                        when exists (" & vbCrLf &
                  "                                      select 1 " & vbCrLf &
                  "                                        from ApuracaoDeCustos " & vbCrLf &
                  "                                       Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf &
                  "                                         And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf &
                  "                                         And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf &
                  "                                         And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf &
                  "                                         And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf &
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf &
                  "                                         And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf &
                  "                                         And Mes_Id                = " & pMes & vbCrLf &
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf &
                  "                                         And CodigoDeCusto_Id      = " & Dr("Placus_Id") & vbCrLf &
                  "                                    )" & vbCrLf &
                  "                           then 'S'" & vbCrLf &
                  "                           else 'N'" & vbCrLf &
                  "                       end) ;" & vbCrLf &
                  " if @Exist = 'N' " & vbCrLf &
                  "  INSERT INTO ApuracaoDeCustos ( " & vbCrLf &
                  "  Empresa_Id" & vbCrLf &
                  ", EndEmpresa_Id" & vbCrLf &
                  ", Deposito_Id" & vbCrLf &
                  ", EndDeposito_Id" & vbCrLf &
                  ", Ano_Id" & vbCrLf &
                  ", Mes_Id" & vbCrLf &
                  ", Produto_Id" & vbCrLf &
                  ", CodigoDeCusto_Id" & vbCrLf &
                  ", EmpresaDestino_Id" & vbCrLf &
                  ", EndEmpresaDestino_Id" & vbCrLf &
                  ", DepositoDestino_Id" & vbCrLf &
                  ", EndDepositoDestino_Id" & vbCrLf &
                  ", ProdutoDerivado_Id" & vbCrLf &
                  ", Quantidade" & vbCrLf &
                  ", ValorDoProduto" & vbCrLf &
                  ", ValorDoFrete" & vbCrLf &
                  ", ValorAuxiliar" & vbCrLf &
                  ", ProdutoDestino" & vbCrLf &
                  ", CodigoDestino)" & vbCrLf &
                  " VALUES('" & Dr("Empresa_Id") & "'" & vbCrLf &
                  ", " & Dr("EndEmpresa_Id") & vbCrLf &
                  ",'" & Dr("Deposito_Id") & "'" & vbCrLf &
                  ", " & Dr("EndDeposito_Id") & vbCrLf &
                  ", " & DdlAno.SelectedValue & vbCrLf &
                  ", " & pMes & vbCrLf &
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf &
                  ", " & Dr("Placus_Id") & vbCrLf &
                  ",'" & Dr("EmpresaDestino_Id") & "'" & vbCrLf &
                  ", " & Dr("EndEmpresaDestino_Id") & vbCrLf &
                  ",'" & Dr("DepositoDestino_Id") & "'" & vbCrLf &
                  ", " & Dr("EndDepositoDestino_Id") & vbCrLf &
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf &
                  ", " & Str(Dr("Quantidade")) & vbCrLf &
                  ", 0" & vbCrLf &
                  ", 0" & vbCrLf &
                  ", 0" & vbCrLf &
                  ",''" & vbCrLf &
                  ", " & Dr("PlacusContra_Id") & vbCrLf &
                  ")" & vbCrLf &
                  " Else" & vbCrLf &
                  "   Update ApuracaoDeCustos set " & vbCrLf &
                  "   Quantidade = " & Str(Dr("Quantidade")) & vbCrLf &
                  " Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf &
                  "   And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf &
                  "   And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf &
                  "   And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf &
                  "   And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf &
                  "   And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf &
                  "   And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf &
                  "   And Mes_Id                = " & pMes & vbCrLf &
                  "   And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf &
                  "   And CodigoDeCusto_Id      = " & Dr("Placus_Id") & vbCrLf

            Sqls.Add(Sql)
            If Not Banco.GravaBanco(Sqls) Then Return False
        Next
        Return True
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CapturaDeDados")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class
