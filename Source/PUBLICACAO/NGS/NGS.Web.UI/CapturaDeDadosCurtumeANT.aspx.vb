Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CapturaDeDadosCurtumeANT
    Inherits BasePage

    Dim Sqls As New ArrayList
    Dim Sqll As String
    Dim SqlAux As String
    Dim Empresa() As String
    Dim Array As New ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CapturaDeDados", "ACESSAR") Then
                If (Left(Session("ssEmpresa"), 8) = "03189063" Or
                    Left(Session("ssEmpresa"), 8) = "05272759") Then

                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.EmpresasConsolidadas, "", False)
                    ddlEmpresa.SelectedValue = Left(HttpContext.Current.Session("ssEmpresa"), 8)
                    ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
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
                    MsgBox(Me.Page, "Movimento já Fechado para esta data...")
                    Exit Sub
                End If
                Try
                    For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue

                        Dim Sql As String = "SELECT * FROM ApuracaoDeCustos" & vbCrLf
                        Sql &= "  Where ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'" & vbCrLf
                        Sql &= "    And ApuracaoDeCustos.Ano_Id = " & DdlAno.SelectedValue & vbCrLf
                        Sql &= "    And ApuracaoDeCustos.Mes_ID = " & Mes

                        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")

                        If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                            MsgBox(Me.Page, "Apuração de Custos do mês " & Mes & " não foi processado!")
                        Else
                            If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                                ContabilizaCustos(Mes)
                            Else
                                If Left(Empresa(0), 8) = "05272759" Then
                                    Contabiliza(Mes)
                                Else
                                    ContabilizaCurtume(Mes)
                                End If
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


                        '****************************************************************************************
                        '********************************   CAPTURAR ESTOQUES  **********************************
                        '****************************************************************************************
                        erroAux = CapturaEstoques(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Dados do estoque apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos dados do estoque"))
                        End If

                        '****************************************************************************************
                        '*****************************   CAPTURAR NOTAS FISCAIS  ********************************
                        '****************************************************************************************
                        erroAux = CapturaNotasFiscais(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Dados das notas fiscais apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos dados das notas fiscais"))
                        End If


                        If Not Left(Empresa(0), 8) = "05272759" AndAlso Not Left(Empresa(0), 8) = "05366261" AndAlso Not Left(Empresa(0), 8) = "38198213" AndAlso Not Left(Empresa(0), 8) = "40938762" AndAlso Not Left(Empresa(0), 8) = "44005444" Then
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


                        If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
                            '****************************************************************************************
                            '*****************************   CAPTURAR NOTAS FISCAIS  DE TRANSFERENCIAS **************
                            '****************************************************************************************
                            If Transferencias(Mes) Then
                                ChkResultado.Items.Add(New ListItem("Dados das transferencias apurados"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & "Erro(s) na apuracao das notas de transferencias"))
                            End If
                            AjustaTotalizadores(Mes)


                            '****************************************************************************************
                            '***************************   AJUSTAR CUSTOS DE SAIDA   ********************************
                            '****************************************************************************************
                            erroAux = AjustaCustosDeSaidas(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Custos de saida ajustador"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nos custos de saída"))
                            End If

                            '****************************************************************************************
                            '*************************** AJUSTAR SAIDAS POR TRANSFERENCIAS  *************************
                            '****************************************************************************************
                            erroAux = AjustaTransferencias(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Saidas por transferencias ajustador"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes das saidas por transferencias"))
                            End If

                            AjustaTotalizadores(Mes)


                            '****************************************************************************************
                            '***************************   AJUSTAR PRODUTO DERIVADO   *******************************
                            '****************************************************************************************
                            erroAux = AjustaProdutoDerivado(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Produto derivado ajustado"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do produto derivado"))
                            End If

                            AjustaTotalizadores(Mes)


                            '****************************************************************************************
                            '***************************   AJUSTAR CONSUMO X PRODUCAO   *****************************
                            '****************************************************************************************
                            erroAux = AjustaConsumoXProducao(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Consumo x producao ajustado"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do consumo x Producao"))
                            End If
                        End If


                        'Acrescentado para o Curtume - Furlan - 11/11/2019
                        If Not Left(Empresa(0), 8) = "05272759" AndAlso Not Left(Empresa(0), 8) = "05366261" AndAlso Not Left(Empresa(0), 8) = "38198213" AndAlso Not Left(Empresa(0), 8) = "40938762" AndAlso Not Left(Empresa(0), 8) = "44005444" Then
                            erroAux = LimpaValores(Mes)

                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Limpa Valores Executado"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante Limpa Valores"))
                            End If

                            erroAux = CapturaSaidasPorTransferencias(Mes)

                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Capturado Saidas Por Transferencias"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante Captura Saidas PorTransferencias"))
                            End If
                        End If

                        erroAux = AjustaTotalizadores(Mes)
                        If erroAux Then
                            ChkResultado.Items.Add(New ListItem("Ajustado Totalizadores"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes dos totalizadores"))
                        End If
                        '****************************************************************************************
                        '********************************   CAPTURAR CUSTOS INDIRETOS****************************
                        '****************************************************************************************


                        If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" OrElse Left(Empresa(0), 8) = "03189063" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
                            '    '****************************************************************************************
                            '    '********************************   CAPTURAR CUSTOS INDIRETOS****************************
                            '    '****************************************************************************************
                            erroAux = CapturaCustosIndiretos(Mes)
                            If erroAux Then
                                ChkResultado.Items.Add(New ListItem("Custos Indiretos apurados"))
                            Else
                                ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) na apuracao dos Custos Indiretos"))
                            End If
                        End If
                        '********************************  CALCULO CIRCULAR   ***********************************
                        '****************************************************************************************
                        ChkResultado.Items.Add(New ListItem("Iniciando Calculo Circular"))
                        For ii = 0 To CInt(txtCiclos.Text)

                            erroAux = AjustaCustosDeSaidas(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nos custos de saida"))

                            erroAux = AjustaMutuos(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes dos mutuos"))

                            erroAux = AjustaTransferencias(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nas saidas por transferencias"))

                            erroAux = AjustaProdutoDerivado(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do produto derivado"))

                            erroAux = AjustaDepositos(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes nas Saidas Para Depósitos"))

                            erroAux = AjustaConsumoXProducao(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante os ajustes do consumo x producao"))

                            erroAux = AjustaTotalizadores(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Totalizadores"))

                            If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
                                erroAux = AjustaTotalizadores(Mes)
                                If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Totalizadores"))
                            End If

                            If Not Left(Empresa(0), 8) = "05272759" AndAlso Not Left(Empresa(0), 8) = "05366261" AndAlso Not Left(Empresa(0), 8) = "38198213" AndAlso Not Left(Empresa(0), 8) = "40938762" AndAlso Not Left(Empresa(0), 8) = "44005444" Then
                                erroAux = AjustaMovimento(Mes)
                                If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Movimento"))
                            End If

                        Next

                        erroAux = AjustaCentavos(Mes)
                        If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Centavos"))

                        If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
                            erroAux = AjustaMovimento(Mes)
                            If erroAux > 0 Then ChkResultado.Items.Add(New ListItem(erroAux & " - Erro(s) durante o ajusta Movimento"))
                            'Contabiliza(Mes)

                            ''''''ContabilizaCustos(Mes)
                            ChkResultado.Items.Add(New ListItem(" FIM - Custo Processado com Sucesso"))
                        Else
                            ''''''Contabiliza(Mes)
                            ChkResultado.Items.Add(New ListItem(" FIM - Custo Processado com Sucesso"))
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

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Private Sub Limpar()
        'ddlEmpresa.SelectedIndex = 0
        DdlAno.SelectedValue = IIf(Now.Month = 1, Now.Year - 1, Now.Year)
        DdlMesInicial.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)
        DdlMesFinal.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)
        ChkResultado.Items.Clear()
    End Sub

    Private Function Validar()
        Empresa = ddlEmpresa.SelectedValue.Split("-")
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório!")
            Return False
        End If
        Return True
    End Function

    'Private Sub CapturarDados()
    '    Dim MesInicial As Integer
    '    Dim MesFinal As Integer

    '    MesInicial = Left(DdlMesInicial.SelectedValue, 2)
    '    MesFinal = Left(DdlMesFinal.SelectedValue, 2)

    '    For Mes = MesInicial To MesFinal
    '        LimpaMovimentoDoMes(Mes)
    '        ApuraSaldoInicial(Mes)
    '        CapturaRazao(Mes)
    '        CapturaEstoques(Mes)
    '        CapturaNotasFiscais(Mes)
    '        CapturaNotasFiscaisContraPartida(Mes)
    '        LimpaValores(Mes)
    '        CapturaSaidasPorTransferencias(Mes)
    '        'AjustaMovimento (Mes)
    '        AjustaTotalizadores(Mes)

    '        If Not String.IsNullOrWhiteSpace(txtCiclos.Text) Then
    '            For index = 0 To CInt(txtCiclos.Text)
    '                'txtLoop.value = i
    '                AjustaCustosDeSaidas(Mes)
    '                AjustaMutuos(Mes)
    '                AjustaTransferencias(Mes)
    '                AjustaProdutoDerivado(Mes)
    '                AjustaConsumoXProducao(Mes)
    '                AjustaTotalizadores(Mes)
    '            Next
    '        End If

    '        AjustaCentavos(Mes)
    '        AjustaMovimento(Mes)
    '        Contabiliza(Mes)
    '    Next

    '    MsgBox(Me.Page, "Processo realizado com Sucesso.", eTitulo.Sucess)
    'End Sub

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
        sql = " Update ApuracaoDeCustos set" & vbCrLf & _
              " ValorDoProduto = ((ac.Quantidade * PM.ValorOficial) / PM.basedecalculo" & vbCrLf & _
              "  from ApuracaoDeCustos AC" & vbCrLf & _
              " inner join TabelaDePrecosDeMercado PM" & vbCrLf & _
              "    on AC.Empresa_Id    = PM.Empresa_Id" & vbCrLf & _
              "   And AC.EndEmpresa_Id = Pm.EndEmpresa_Id" & vbCrLf & _
              "   And AC.Deposito_Id   = PM.Deposito_Id" & vbCrLf & _
              "   And AC.EndEmpresa_Id = Pm.EndDeposito_id" & vbCrLf & _
              "   And AC.Produto_Id    = PM.Produto_Id" & vbCrLf & _
              "   And PM.Data_Id       = '" & New DateTime(DdlAno.SelectedValue, pMes, Date.DaysInMonth(DdlAno.SelectedValue, pMes)) & "'" & vbCrLf & _
              " inner join Produtos P" & vbCrLf & _
              "    on P.Produto_Id = AC.Produto_Id" & vbCrLf & _
              " where AC.CodigoDeCusto_id = 110" & vbCrLf & _
              "   and AC.Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
              "   and AC.Mes_Id           = " & pMes
        Sqls.Add(sql)

        'Reavaliar Quantidade Negativa
        sql = " Update ApuracaoDeCustos set"
        sql &= "  ValorDoProduto = ac.Quantidade * (MesAnterior.ValorDoProduto / MesAnterior.Quantidade)"
        sql &= "  from ApuracaoDeCustos AC"
        sql &= "  Inner Join ( Select * from ApuracaoDeCustos) MesAnterior"
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

        If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
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
        Dim Sql As String = "SELECT * from ApuracaoDeCustosXFiltroRazao"
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
        Sql &= "        And  Razao.Lote_Id not in(7000, 9, 10) And Left(Razao.Conta_Id, 7) = '1010601'" & vbCrLf
        Sql &= " GROUP  BY Razao.Empresa_Id, Razao.EndEmpresa_Id, " & vbCrLf
        Sql &= "        isnull(Razao.Deposito, Razao.Empresa_Id) , " & vbCrLf
        Sql &= "        isnull(Razao.EndDeposito, Razao.EndEmpresa_Id), " & vbCrLf
        Sql &= "        PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto, Razao.Conta_Id, Razao.Lote_Id " & vbCrLf

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
                sql &= "       when exists ("
                sql &= "                 select * from ApuracaoDeCustos "
                sql &= " Where "
                sql &= "     Empresa_Id = '" & rs("Empresa_Id") & "'"
                sql &= " And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                sql &= " And Deposito_Id = '" & rs("Deposito") & "'"
                sql &= " And EndDeposito_Id = " & rs("EndDeposito")
                sql &= " And Ano_Id = " & DdlAno.SelectedValue
                sql &= " And Mes_Id = " & pMes
                sql &= " And Produto_Id = '" & rs("Produto") & "'"
                sql &= " And CodigoDeCusto_Id = " & rs("Codigo_Id")

                sql &= ")"
                sql &= "            then 'S'"
                sql &= "             else 'N'"
                sql &= "               end) ;"

                sql &= "  if @Exist = 'N' "
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
                sql &= ", CodigoDestino)"
                sql &= " VALUES('" & rs("Empresa_Id") & "'"
                sql &= ", " & rs("EndEmpresa_Id")
                sql &= ", '" & rs("Deposito") & "'"
                sql &= ", " & rs("EndDeposito")
                sql &= ", " & DdlAno.SelectedValue
                sql &= ", " & pMes
                sql &= ", '" & rs("Produto") & "'"
                sql &= ", " & rs("Codigo_Id")
                sql &= ", ''" 'Empresa Destino
                sql &= ", 0"  'End Empresa Destino
                sql &= ", ''" 'Deposito Destino
                sql &= ", 0"  'End DepositoDestino
                sql &= ", ''" 'Produto Derivado

                sql &= ", 0" 'Quantidade
                sql &= ", " & Str(rs("Valor"))
                sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                sql &= ", ''" '& rs("ProdutoDestino") & "'"
                sql &= ", 0" '& rs("CodigoDestino")
                sql &= ")"

                sql &= " Else"

                sql &= "  Update ApuracaoDeCustos set "
                Sql &= "  ValorDoProduto = ValorDoProduto + " & Str(rs("Valor"))
                sql &= " Where Empresa_Id = '" & rs("Empresa_Id") & "'"
                sql &= "   And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                sql &= "   And Deposito_Id = '" & rs("Deposito") & "'"
                sql &= "   And EndDeposito_Id = " & rs("EndDeposito")
                sql &= "   And Ano_Id = " & DdlAno.SelectedValue
                sql &= "   And Mes_Id = " & pMes
                sql &= "   And Produto_Id = '" & rs("Produto") & "'"
                sql &= "   And CodigoDeCusto_Id = " & rs("Codigo_Id")
                Sqll = Sqll & Sql
            Next
        End If
        Return Banco.GravaBanco(Sqll)
    End Function

    Private Function CapturaCustosIndiretos(ByVal pMes As Integer) As Boolean
        'txt_Processo.value = "Processando Custos Indiretos..."

        Dim QuantidadeTotal As Double
        Dim ValorTotal As Double
        Dim ValorTotalConsumido As Double
        Dim Valor As Double
        Dim sql As String
        

        '------------"  Capturando dados de Custos Indiretos

        sql = " SELECT SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS Valor"
        sql &= " FROM  Razao INNER JOIN"
        sql &= "       PlanoDeCustosXOrigem on LEFT(Razao.Conta_Id, 7)  =  PlanoDeCustosXOrigem.Conta_Id "
        sql &= " Where Razao.Empresa_Id like '" & Left(Empresa(0), 8) & "%'"
        sql &= "   And Month(Movimento_Id) = " & pMes
        sql &= "   And YEAR(Movimento_Id)  =  " & DdlAno.SelectedValue
        sql &= "   And Lote_Id <> 7000"
        sql &= "   And isnull(PlanoDeCustosXOrigem.Produto, '0') = 0 "
        sql &= "   And  PlanoDeCustosXOrigem.Codigo_Id = 402"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ApuracaoDeCustos")
        For Each rs As DataRow In ds.Tables(0).Rows
            If Not IsDBNull(rs("Valor")) Then
                ValorTotal = ValorTotal + rs("Valor")
                ValorTotalConsumido = ValorTotal
            End If
        Next

        '--------------------------------------------------------------------

        If Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then 'Nutri/NutriLog/Baxi/Fronteira
            sql = "SELECT n.Empresa_Id, n.EndEmpresa_Id, n.Deposito AS Deposito_Id, n.EndDeposito AS EndDeposito_Id, NI.Produto_id," & vbCrLf & _
                    "	Sum(Case" & vbCrLf & _
                    "			When so.Devolucao = 'S'" & vbCrLf & _
                    "				Then NI.PesoFiscal * -1" & vbCrLf & _
                    "				else NI.PesoFiscal" & vbCrLf & _
                    "		end) AS Quantidade" & vbCrLf & _
                    "FROM NOTASFISCAISXITENS NI" & vbCrLf & _
                    "	INNER JOIN NotasFiscais n" & vbCrLf & _
                    "			ON n.Empresa_Id = NI.Empresa_Id" & vbCrLf & _
                    "			and n.EndEmpresa_iD = NI.EndEmpresa_Id" & vbCrLf & _
                    "			and n.EndEmpresa_iD = NI.EndEmpresa_Id" & vbCrLf & _
                    "			and n.Cliente_Id    = NI.Cliente_Id" & vbCrLf & _
                    "			and n.EndCliente_Id = NI.EndCliente_Id" & vbCrLf & _
                    "			and n.EntradaSaida_Id = NI.EntradaSaida_Id" & vbCrLf & _
                    "			and n.Serie_Id        = NI.Serie_Id" & vbCrLf & _
                    "			and n.Nota_Id         = NI.Nota_Id" & vbCrLf & _
                    "    INNER JOIN SubOperacoes so" & vbCrLf & _
                    "        ON so.Operacao_Id      = n.Operacao" & vbCrLf & _
                    "		AND so.SubOperacoes_Id = n.SubOperacao" & vbCrLf & _
                    "	INNER JOIN Produtos p" & vbCrLf & _
                    "			ON p.Produto_Id = NI.Produto_Id" & vbCrLf & _
                    "WHERE n.Empresa_Id like '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
                    "  AND n.Situacao = 1" & vbCrLf & _
                    "  AND n.TipoDeDocumento = 1" & vbCrLf & _
                    "  AND MONTH(n.Movimento) =  " & pMes & vbCrLf & _
                    "  AND YEAR(n.Movimento)  =  " & DdlAno.SelectedValue & vbCrLf & _
                    "  AND so.Operacao_id in(1,7)" & vbCrLf & _
                    "GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Deposito, n.EndDeposito, NI.produto_id"

        Else 'Quimica
            sql = " Select  Empresa_ID, EndEmpresa_ID, Deposito_ID, EndDeposito_Id, "
            sql &= "   Produto_Id, Operacao_Id, SubOperacao_Id, ProdutoDerivado_Id, "
            sql &= "   Sum(Entradas) as Quantidade from producao"

            sql &= "   WHERE Producao.Empresa_Id like '" & Left(Empresa(0), 8) & "%'"
            sql &= "   AND MONTH(Producao.Movimento_Id) =  " & pMes
            sql &= "   AND YEAR(Producao.Movimento_Id)  =  " & DdlAno.SelectedValue
            sql &= "   AND Producao.FisicoFiscal_Id     = 2"
            sql &= "   And Operacao_Id = 40 and SubOPeracao_Id = 02 "
            sql &= "   And Left(Produto_Id, 5) = '30101'"

            sql &= "   Group by  Empresa_ID, EndEmpresa_ID, Deposito_ID, EndDeposito_Id, "
            sql &= "             Produto_Id, Operacao_Id, SubOperacao_Id,  ProdutoDerivado_Id"
            sql &= "   Order by Produto_ID"
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
                    sql &= "  ValorDoProduto = " & Replace(Valor, ",", ".")
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
                Return Banco.GravaBanco(Sqll)

            End If
        Else

        End If



    End Function


    Private Function CapturaEstoques(ByVal pMes As Integer) As Boolean
        Dim i = 0
        Dim aux As Boolean = True

        '07 - CAPTURANDO DADOS DE ESTOQUES - CRIANDO FILTRO
        Dim Sql As String = "SELECT * from ApuracaoDeCustosXFiltroEstoques"
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
        sql &= "              Consulta.Placus_Id, Consulta.PlacusContraPartida, SUM(Consulta.Quantidade) AS Quantidade"
        sql &= "  FROM         (SELECT     Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Operacao_Id,"
        sql &= "                                      Producao.SubOperacao_Id, Producao.ProdutoDerivado_Id, SubOperacoes.ApuracaoDeCustos AS Placus_Id,"
        sql &= "                                      ISNULL(SubOperacoes.ApuracaodeCustosContraPartida, 0) AS PlacusContraPartida, "
        sql &= "               Case when Producao.Entradas <> 0 then Producao.Entradas else Producao.Saidas end AS Quantidade"
        sql &= "               FROM          Producao INNER JOIN"
        sql &= "                                      SubOperacoes ON Producao.Operacao_Id = SubOperacoes.Operacao_Id AND Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id INNER JOIN"
        sql &= "                                      Produtos ON Producao.Produto_Id = Produtos.Produto_Id INNER JOIN"
        sql &= "                                      GruposDeEstoques ON Produtos.Grupo = GruposDeEstoques.Grupo_Id"

        sql &= " WHERE Producao.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'"
        sql &= "   AND MONTH(Producao.Movimento_Id) =  " & pMes
        sql &= "   AND YEAR(Producao.Movimento_Id)  =  " & DdlAno.SelectedValue
        sql &= "   AND Producao.FisicoFiscal_Id     = 2"
        sql &= "   AND GruposDeEstoques.custo       = 1"

        'Acrescentado para o Curtume - 11/11/2019 - Furlan
        If Left(Empresa(0), 8) = "03189063" Then
            Sql &= "   AND SubOperacoes.ApuracaoDeCustos > 0"
        End If

        'sql &=  " GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.ProdutoDerivado_Id,"
        'sql &=  "          Producao.Operacao_Id, Producao.SubOperacao_Id, SubOperacoes.ApuracaoDeCustos, ISNULL(SubOperacoes.ApuracaodeCustosContraPartida, 0))"

        sql &= "          ) AS Consulta INNER JOIN"
        sql &= "          Produtos ON Consulta.Produto_Id = Produtos.Produto_Id"
        sql &= " GROUP BY Consulta.Empresa_Id, Consulta.EndEmpresa_Id, Consulta.Deposito_Id, Consulta.EndDeposito_Id, Consulta.Produto_Id, Consulta.ProdutoDerivado_Id,"
        sql &= "              Consulta.Placus_Id , Consulta.PlacusContraPartida"

        '08 - CAPTURANDO DADOS DE ESTOQUES - ATUALIZANDO TABELA DE APURAÇÃO
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds2.Tables(0).Rows
                Sql = " set @Exist = (select case "
                sql &= "                        when exists ("
                sql &= "                                     select 1"
                sql &= "                                       from ApuracaoDeCustos "
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
                sql &= "                                        And CodigoDeCusto_Id   = " & rs("Placus_Id")
                sql &= "                                        And CodigoDestino      = " & rs("PlacusContraPartida")
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
                sql &= " VALUES('" & rs("Empresa_Id") & "'"
                sql &= ", " & rs("EndEmpresa_Id")
                sql &= ", '" & rs("Deposito_Id") & "'"
                sql &= ", " & rs("EndDeposito_Id")
                sql &= ", " & DdlAno.SelectedValue
                sql &= ", " & pMes
                sql &= ", '" & rs("Produto_Id") & "'"
                Sql &= ", " & rs("Placus_Id")
                Sql &= ",'" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ",'" & rs("Empresa_Id") & "'"
                Sql &= ", " & rs("EndEmpresa_Id")
                Sql &= ", '" & rs("ProdutoDerivado_Id") & "'"


                If (Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444") AndAlso rs("Produto_Id") = "102010003" Then
                    Dim qdte As Decimal = rs("Quantidade") * 40
                    Sql &= ", " & Replace(qdte, ",", ".")
                Else
                    Sql &= ", " & Replace(rs("Quantidade"), ",", ".")
                End If

                Sql &= ", 0" '& rs("Valor").ToString.Replace(",", ".")
                sql &= ", 0" '& rs("ValorDoFrete").ToString.Replace(",", ".")
                sql &= ", 0" '& rs("ValorAuxiliar").ToString.Replace(",", ".")
                sql &= ", ''" '& rs("ProdutoDestino") & "'"
                sql &= ", " & rs("PlacusContraPartida")
                sql &= ", ''" '& rs("Reduzido") & "'"
                sql &= ")"
                sql &= " Else"
                sql &= "  Update ApuracaoDeCustos set "
                sql &= "  Quantidade = " & Replace(rs("Quantidade"), ",", ".")
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
                sql &= "   And CodigoDeCusto_Id   = " & rs("Placus_Id")
                sql &= "   And CodigoDestino      = " & rs("PlacusContraPartida")
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

        Dim Sql As String = "SELECT * from ApuracaoDeCustosXFiltroNotas order by ordem"
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
        sql &= "              AS ApuracaodeCustosContrapartida, Produto_Id, SUM(Quantidade) AS Quantidade, ISNULL(SUM(CASE WHEN ((EntradaSaida_Id = 'E') OR"
        Sql &= "              (EntradaSaida_Id = 'S' AND Devolucao = 'S')) THEN (ValorDoProduto + DespesasAduaneiras + ValorFrete) - (ValorCOFINS + ValorPIS + ValorICMS + ValorDESCONTO) ELSE ValorDoProduto END), 0) AS ValorDoProduto, "
        Sql &= "              Sum(ValorICMS) as ICMS"
        sql &= " FROM         (SELECT     NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, ISNULL(NotasFiscaisXItens.Deposito, NotasFiscais.Empresa_Id) AS Deposito_Id,"
        sql &= "                                      ISNULL(NotasFiscaisXItens.EndDeposito, NotasFiscais.EndEmpresa_Id) AS EndDeposito_Id, YEAR(ISNULL(NotasFiscais.DataParaCusto,"
        sql &= "                                      NotasFiscais.Movimento)) AS Ano_Id, MONTH(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)) AS Mes_Id, NotasFiscaisXItens.Produto_Id,"
        sql &= "                                      SubOperacoes.ApuracaoDeCustos AS Placus_Id,"
        sql &= "                                      CASE WHEN SubOperacoes.ApuracaodeCustosContraPartida <> 0 THEN NotasFiscais.Cliente_Id ELSE '' END AS DepositoDestino_Id,"
        sql &= "                                      CASE WHEN SubOperacoes.ApuracaodeCustosContraPartida <> 0 THEN NotasFiscais.EndCliente_Id ELSE 0 END AS EndDepositoDestino_Id,"
        Sql &= "                                      SubOperacoes.ApuracaodeCustosContraPartida, SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade, SUM(NotasFiscaisXItens.Valor) AS ValorDoProduto,"
        sql &= "                                      NotasFiscais.EntradaSaida_Id, SUM(Enc.ValorIPI) AS ValorIPI, SUM(Enc.ValorICMS) AS ValorICMS, SUM(Enc.ValorCOFINS) AS ValorCOFINS,"
        Sql &= "                                      SUM(Enc.ValorPIS) AS ValorPIS, SUM(Enc.DespesasAduaneiras) AS DespesasAduaneiras, SUM(Enc.ValorDESCONTO) AS ValorDesconto ,SUM(Enc.ValorFrete) AS ValorFrete, "
        Sql &= "                                      SubOperacoes.Devolucao "
        sql &= "               FROM          GruposDeEstoques LEFT OUTER JOIN"
        sql &= "                                      Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo RIGHT OUTER JOIN"
        sql &= "                                      NotasFiscais INNER JOIN"
        sql &= "                                      NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND"
        sql &= "                                      NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND"
        sql &= "                                      NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND"
        sql &= "                                      NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id ON"
        sql &= "                                      Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id LEFT OUTER JOIN"
        sql &= "                                      SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND"
        sql &= "                                      NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN"
        sql &= "                                      PlanoDeCustos ON SubOperacoes.ApuracaoDeCustos = PlanoDeCustos.Codigo_Id "
        sql &= "                                      LEFT OUTER JOIN"
        sql &= "                                          (SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id,"
        Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN('IPI') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorIPI,"
        Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = 'ICMS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorICMS,"

        If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
            Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('COFINS','COFINS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorCOFINS,"
            Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorPIS, "
        Else
            Sql &= "                                                                  ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) = ('COFINS') Then Case When ValorNovo > 0 Then ValorNovo Else Valor End Else 0 	End), 0) As ValorCOFINS, "
            Sql &= "                                                                  ISNULL(SUM(Case When UPPER(Encargo_Id) = ('PIS')    THEN CASE WHEN ValorNovo > 0 THEN ValorNovo ELSE Valor END Else 0 	End), 0) As ValorPIS, "

            'Sql &= "                                                                   ISNULL(SUM(Case When UPPER(Encargo_Id) In ('COFINS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorCOFINS,"
            'Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('PIS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorPIS, "
        End If

        Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('DESPESAS', 'DESP.ADUANEIRAS', 'ICMS SUBSTITUIC') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS DespesasAduaneiras,"
        Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('DESCONTOS') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorDesconto, "
        Sql &= "                                                                   ISNULL(SUM(CASE WHEN UPPER(Encargo_Id) IN ('FRETES') THEN NotasFiscaisXEncargos.Valor ELSE 0 END), 0) AS ValorFrete "
        Sql &= "                                            FROM NotasFiscaisXEncargos"
        sql &= "                                            GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id) AS Enc ON"
        sql &= "                                      NotasFiscaisXItens.Empresa_Id = Enc.Empresa_Id AND NotasFiscaisXItens.EndEmpresa_Id = Enc.EndEmpresa_Id AND"
        sql &= "                                      NotasFiscaisXItens.Cliente_Id = Enc.Cliente_Id AND NotasFiscaisXItens.EndCliente_Id = Enc.EndCliente_Id AND"
        sql &= "                                      NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = Enc.Serie_Id AND"
        sql &= "                                      NotasFiscaisXItens.Nota_Id = Enc.Nota_Id And NotasFiscaisXItens.Produto_Id = Enc.Produto_Id And NotasFiscaisXItens.CFOP_Id = Enc.CFOP_Id"

        sql &= " WHERE      (NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        sql &= "        AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & pMes & ")"
        sql &= "        AND (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))   = " & DdlAno.SelectedValue & ")"
        sql &= "        AND (SubOperacoes.ApuracaoDeCustos > 0)"
        sql &= "        And  NotasFiscais.Situacao = 1"
        sql &= "        And  GruposDeEstoques.custo = 1"
        Sql &= "        And PlanoDeCustos.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'"
        'sql &=  Sqla & "; "

        sql &= " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito, NotasFiscaisXItens.EndDeposito, NotasFiscais.Cliente_Id,"
        sql &= "                                      NotasFiscais.EndCliente_Id, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,"
        sql &= "                                      NotasFiscais.Movimento)), NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro,"
        sql &= "                                      NotasFiscaisXItens.EndDepositoTerceiro, SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id, SubOperacoes.Devolucao)"
        sql &= "              AS Consulta"
        sql &= " WHERE (Placus_Id <> 1)"
        sql &= " GROUP BY Empresa_Id, EndEmpresa_Id, Placus_Id, ApuracaodeCustosContraPartida, Produto_Id, Deposito_Id, EndDeposito_Id, DepositoDestino_Id, EndDepositoDestino_Id"

        '11 - CAPTURANDO DADOS DAS NOTAS FISCAIS - ATUALIZANDO TABELA DE APURAÇÃO
        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds2.Tables(0).Rows
                Sql = " set @Exist = (select case "
                Sql &= "       when exists ("
                Sql &= "                 select * from ApuracaoDeCustos "
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

                Sql &= ", ''"  'Empresa Destino
                Sql &= ", 0"   'End Empresa Destino

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
                    Sql &= ", " & Str(rs("ValorDoProduto"))
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
        Sql &= "                          SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade, SUM(NotasFiscaisXItens.Valor) AS ValorDoProduto, NotasFiscais.EntradaSaida_Id, SUM(Enc.ValorIPI)"
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

        If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" OrElse Left(Empresa(0), 8) = "62747840" OrElse Left(Empresa(0), 8) = "62780383" OrElse Left(Empresa(0), 8) = "63358210" Then
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
                Sql &= "                 select * from ApuracaoDeCustos "
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
        Sql &= "           SubOperacoes.ApuracaodeCustosContraPartida AS PlacusContra_Id, SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade, 0 AS ValorDoProduto"
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
                Sql &= "                 select * from ApuracaoDeCustos "
                Sql &= " Where "
                Sql &= "     Empresa_Id = '" & rs("Empresa_Id") & "'"
                Sql &= " And EndEmpresa_Id = " & rs("EndEmpresa_Id")
                Sql &= " And Deposito_Id = '" & rs("Deposito_Id") & "'"
                Sql &= " And EndDeposito_Id = " & rs("EndDeposito_Id")
                Sql &= " And Ano_Id = " & DdlAno.SelectedValue
                Sql &= " And Mes_Id = " & pMes
                Sql &= " And Produto_Id = '" & rs("Produto_Id") & "'"
                Sql &= " And CodigoDeCusto_Id = " & rs("Placus_Id")

                If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
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

                If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
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
            Sql &= "  AND (PC.FaseDoTotalizador  = " & Fase & ")"
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
                    Sql &= "                                      select *"
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
        Dim a As String = "A"
        Dim b As String = "B"

        Dim Sql As String = " SELECT Empresa_Id, EndEmpresa_ID, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, "
        sql &= "  CodigoDeCusto_Id, DepositoDestino_ID, EndDepositoDestino_Id, ProdutoDerivado_Id, "
        Sql &= "  Convert(Decimal(18,9),((ValorDoProduto + ValorDoFrete) / Quantidade)) as Medio,  Case when Empresa_Id = Deposito_Id then 1 else 2 end as Ordem  "
        sql &= "  FROM ApuracaoDeCustos"
        sql &= "  Where Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
        sql &= "   And Mes_ID           = " & pMes
        sql &= "   And CodigoDeCusto_Id = 495"
        Sql &= "   And Quantidade       > 0"
        'sql &= "   And Quantidade       > 0" No Custo Planilha Curtume estava Quantidade       <> 0 - Furlan - 11/11/2019

        Sql &= "   Order by Empresa_Id, EndEmpresa_ID, Case when Empresa_Id = Deposito_Id then 1 else 2 end, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, "
        sql &= "   CodigoDeCusto_Id, DepositoDestino_ID, EndDepositoDestino_Id, ProdutoDerivado_Id"
        Sqll = ""

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = "  Update  ApuracaoDeCustos set"
                sql &= "    ApuracaoDeCustos.ValorDoProduto = ApuracaoDeCustos.Quantidade * " & Replace(rs("Medio"), ",", ".")
                sql &= " FROM  ApuracaoDeCustos INNER JOIN"
                sql &= " PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
                sql &= "    Where ApuracaoDeCustos.Empresa_Id       ='" & rs("Empresa_Id") & "'"
                sql &= "    And ApuracaoDeCustos.EndEmpresa_Id    = " & rs("EndEmpresa_Id")
                sql &= "    And ApuracaoDeCustos.Deposito_Id      ='" & rs("Deposito_Id") & "'"
                sql &= "    And ApuracaoDeCustos.EndDeposito_Id   = " & rs("EndDeposito_Id")
                sql &= "    And ApuracaoDeCustos.Ano_Id           = " & rs("Ano_Id")
                sql &= "    And ApuracaoDeCustos.Mes_Id           = " & rs("Mes_Id")
                sql &= "    And ApuracaoDeCustos.Produto_Id       ='" & rs("Produto_Id") & "'"
                sql &= "    And ApuracaoDeCustos.CodigoDeCusto_Id > 495 "
                Sql &= "    And Not (PlanoDeCustos.Classe In ('" & eClassesOperacoes.MUTUO.ToString & "'))"

                If Left(Empresa(0), 8) = "05272759" Then
                    If rs("Produto_Id") = "30101014" Then
                        a = b
                    End If
                End If

                Sqll = Sqll & Sql
            Next
        End If
        Return Banco.GravaBanco(Sqll)
    End Function

    Private Function AjustaMovimento(ByVal pMes As Integer) As Boolean
        '----------" 17 - Ajustando Custos sem Movimento "))
        'txt_Processo.value = "Ajustando Custos sem movimento. .."

        Dim Sql As String = "  SELECT * FROM ApuracaoDeCustos"
        Sql &= "  Where ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   And ApuracaoDeCustos.Ano_Id           = " & DdlAno.SelectedValue
        Sql &= "   And ApuracaoDeCustos.Mes_ID           = " & pMes
        Sql &= "And CodigoDeCusto_Id in (402, 404)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                SqlAux = "        SELECT * "
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
        sql &= "          case when len(EmpresaDestino_Id) > 0 then EndEmpresaDestino_Id else EndEmpresa_Id end as EndEmpresa_Id,"
        sql &= "          case when len(DepositoDestino_Id) > 0 then DepositoDestino_Id else Deposito_Id end as Deposito_Id,"
        sql &= "          case when len(DepositoDestino_Id) > 0 then EndDepositoDestino_Id else EndDeposito_Id end as EndDeposito_Id,"
        sql &= "          Ano_Id, Mes_Id,Produto_Id as ProdutoOrigem, ProdutoDerivado_Id as Produto_Id, CodigoDestino as CodigoDeCusto_Id, '' as EmpresaDestino_Id, 0 as EndEmpresaDestino_Id, "
        Sql &= "          '' as DepositoDestino_Id, 0 as EndDepositoDestino_Id, '' as ProdutoDerivado_Id, isnull(Etapa,0) as Etapa, "
        If Left(Empresa(0), 8) = "03189063" Then
            Sql &= "            Case when CodigoDestino In (755, 757) then Quantidade else 0 end as Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, "
        Else
            Sql &= "            Case when CodigoDestino In (351, 352, 755, 757) then Quantidade else 0 end as Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, "
        End If
        Sql &= "          CodigoDeCusto_Id as CodigoDestino, Reduzido"
        Sql &= "   from ApuracaoDeCustos INNER JOIN"
        Sql &= "        PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"
        Sql &= " WHERE Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        Sql &= "   and Ano_id        = " & DdlAno.SelectedValue
        Sql &= "   And Mes_id        = " & pMes
        Sql &= "   And ProdutoDerivado_Id > 0"
        Sql &= "   And ((CodigoDestino Between 001 and  500) or (CodigoDestino in (755,757)))"
        Sql &= "   And PlanoDeCustos.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' "
        Sql &= "   and isnull(codigodestino,0) > 0"

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

                    'If rs("CodigoDestino") = "355" Then
                    '    Sql &= "   And Deposito_Id           ='" & rs("Empresa_Id") & "'"
                    '    Sql &= "   And EndDeposito_Id        = " & rs("EndEmpresa_Id")
                    'Else
                    '    Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                    '    Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                    'End If

                    Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                    Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
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

                    'If rs("CodigoDestino") = "355" Then
                    '    Sql &= "   And Deposito_Id           ='" & rs("Empresa_Id") & "'"
                    '    Sql &= "   And EndDeposito_Id        = " & rs("EndEmpresa_Id")
                    'Else
                    '    Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                    '    Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
                    'End If

                    Sql &= "   And Deposito_Id           ='" & rs("Deposito_Id") & "'"
                    Sql &= "   And EndDeposito_Id        = " & rs("EndDeposito_Id")
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

    Private Function AjustaTransferencias(ByVal pMes As Integer) As Boolean
        '15 - AJUSTANDO PRODUTOS DERIVADOS
        Dim Sql As String = " SELECT ApuracaoDeCustos.Empresa_Id,"
        sql &= "    ApuracaoDeCustos.EndEmpresa_Id,"
        sql &= "    ApuracaoDeCustos.Deposito_Id,"
        sql &= "    ApuracaoDeCustos.EndDeposito_Id,"
        sql &= "    ApuracaoDeCustos.Ano_Id,"
        sql &= "    ApuracaoDeCustos.Mes_Id,"
        sql &= "    ApuracaoDeCustos.Produto_Id,"
        sql &= "    ApuracaoDeCustos.CodigoDeCusto_Id,"
        sql &= "    ApuracaoDeCustos.DepositoDestino_Id as EmpresaDestino_Id,"
        sql &= "    ApuracaoDeCustos.EndDepositoDestino_Id as EndEmpresaDestino_Id,"
        sql &= "    ApuracaoDeCustos.DepositoDestino_Id,"
        sql &= "    ApuracaoDeCustos.EndDepositoDestino_Id,"
        sql &= "    ApuracaoDeCustos.ProdutoDerivado_Id,"
        sql &= "    ApuracaoDeCustos.Quantidade,"
        sql &= "    ApuracaoDeCustos.ValorDoProduto,"
        sql &= "    ApuracaoDeCustos.ValorDoFrete,"
        sql &= "    ApuracaoDeCustos.ValorAuxiliar,"
        sql &= "    ApuracaoDeCustos.ProdutoDestino,"
        sql &= "    ApuracaoDeCustos.CodigoDestino "

        sql &= " FROM      ApuracaoDeCustos INNER JOIN"
        sql &= "           PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id"

        sql &= " WHERE     ApuracaoDeCustos.Empresa_ID like '" & Left(Empresa(0), 8) & "%'"
        sql &= "   and     ApuracaoDeCustos.Ano_id        = " & DdlAno.SelectedValue
        sql &= "   And     ApuracaoDeCustos.Mes_id        = " & pMes
        sql &= "   And     (ApuracaoDeCustos.CodigoDeCusto_Id > 500)"
        Sql &= "   And     (PlanoDeCustos.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "')"

        sql &= "   and     isnull(ApuracaoDeCustos.codigodestino,0) > 0"

        Sqll = " Declare"
        Sqll = Sqll & " @Exist as varchar(1)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Sql = " set @Exist = (select case "
                sql &= "                        when exists ("
                sql &= "                                      select 1 "
                sql &= "                                        from ApuracaoDeCustos "
                sql &= "                                       Where Empresa_Id            ='" & rs("EmpresaDestino_Id") & "'"
                sql &= "                                         And EndEmpresa_Id         = " & rs("EndEmpresaDestino_Id")
                sql &= "                                         And Deposito_Id           ='" & rs("DepositoDestino_Id") & "'"
                sql &= "                                         And EndDeposito_Id        = " & rs("EndDepositoDestino_Id")
                sql &= "                                         And Ano_Id                = " & rs("Ano_Id")
                sql &= "                                         And Mes_Id                = " & rs("Mes_Id")
                sql &= "                                         And Produto_Id            ='" & rs("Produto_Id") & "'"
                sql &= "                                         And CodigoDeCusto_Id      = " & rs("CodigoDestino")
                sql &= "                                         And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                sql &= "                                         And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                sql &= "                                         And DepositoDestino_Id    ='" & rs("Deposito_Id") & "'"
                sql &= "                                         And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                sql &= "                                         And ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'"
                sql &= "                                    )"
                sql &= "                          then 'S'"
                sql &= "                          else 'N'"
                sql &= "                     end) ;" & vbCrLf

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
                sql &= ", CodigoDestino)"

                sql &= " VALUES('" & rs("EmpresaDestino_Id") & "'"
                sql &= ", " & rs("EndEmpresaDestino_Id")
                sql &= ",'" & rs("DepositoDestino_Id") & "'"
                sql &= ", " & rs("EndDepositoDestino_Id")
                sql &= ", " & rs("Ano_Id")
                sql &= ", " & rs("Mes_Id")
                sql &= ",'" & rs("Produto_Id") & "'"
                sql &= ", " & rs("CodigoDestino")
                sql &= ",'" & rs("Empresa_Id") & "'"
                sql &= ", " & rs("EndEmpresa_Id")
                sql &= ",'" & rs("Deposito_Id") & "'"
                sql &= ", " & rs("EndDeposito_Id")
                sql &= ",'" & rs("ProdutoDerivado_Id") & "'"
                sql &= ", " & Replace(rs("Quantidade"), ",", ".")
                sql &= ", " & Replace(rs("ValorDoProduto"), ",", ".")
                sql &= ", " & Replace(rs("ValorDoFrete"), ",", ".")
                sql &= ", " & Replace(rs("ValorAuxiliar"), ",", ".")
                sql &= ",'" & rs("ProdutoDestino") & "'"
                sql &= ", " & rs("CodigoDeCusto_Id")
                sql &= ")"

                sql &= "  Update ApuracaoDeCustos set "
                sql &= "         ValorDoProduto     =  " & Replace(rs("ValorDoProduto"), ",", ".")
                sql &= "        ,ValorDoFrete       =  " & Replace(rs("ValorDoFrete"), ",", ".")
                sql &= " Where Empresa_Id            ='" & rs("EmpresaDestino_Id") & "'"
                sql &= "   And EndEmpresa_Id         = " & rs("EndEmpresaDestino_Id")
                sql &= "   And Deposito_Id           ='" & rs("DepositoDestino_Id") & "'"
                sql &= "   And EndDeposito_Id        = " & rs("EndDepositoDestino_Id")
                sql &= "   And Ano_Id                = " & rs("Ano_Id")
                sql &= "   And Mes_Id                = " & rs("Mes_Id")
                sql &= "   And Produto_Id            ='" & rs("Produto_Id") & "'"
                sql &= "   And CodigoDeCusto_Id      = " & rs("CodigoDestino")
                sql &= "   And EmpresaDestino_Id     ='" & rs("Empresa_Id") & "'"
                sql &= "   And EndEmpresaDestino_Id  = " & rs("EndEmpresa_Id")
                sql &= "   And DepositoDestino_Id    ='" & rs("Deposito_Id") & "'"
                sql &= "   And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                sql &= "   And ProdutoDerivado_Id    ='" & rs("ProdutoDerivado_Id") & "'"
                Sqll = Sqll & Sql
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
                sql &= "   @Soma as numeric(18,9)"
                sql &= "   select ACO.Empresa_Id,         ACO.EndEmpresa_id,"
                sql &= "           ACO.Deposito_Id,        ACO.EndDeposito_Id,"
                sql &= "           ACO.Ano_Id,"
                sql &= "           ACO.Mes_Id,"
                sql &= "           ACO.produto_Id as ProdutoOrigem,"
                sql &= "           ACO.Quantidade as QuantidadeOrigem,"
                sql &= "           ACO.ValorDoProduto as VlrProdutoOrigem,"
                sql &= "           ACD.produto_Id as ProdutoDestino,"
                sql &= "           ACD.CodigoDeCusto_Id as CodigoDeCustoDestino,"
                sql &= "           ACD.Quantidade as QuantidadeDestino,"
                sql &= "           TPM.ValorOficial as ValorDeMercado,"
                sql &= "           (ACD.Quantidade/TPM.BaseDeCalculo)*TPM.ValorOficial as VlrTotal,"
                sql &= "           convert(numeric(18,9),0) as Soma"
                sql &= "       into #Temp"
                sql &= "       from ApuracaoDeCustos ACO"
                sql &= "       Inner Join(select AC.Empresa_Id,"
                sql &= "                      AC.EndEmpresa_Id,"
                sql &= "                      AC.Ano_Id,"
                sql &= "                      AC.Mes_Id,"
                sql &= "                      AC.Produto_Id,"
                sql &= "                      AC.CodigoDeCusto_Id,"
                sql &= "                      AC.Quantidade"
                sql &= "                 from ApuracaoDeCustos AC "
                sql &= "                   Inner Join ConsumoxProducao CP "
                sql &= "                   on CP.ProdutoOrigem_Id    ='" & rs("ProdutoOrigem_Id") & "'"
                sql &= "                   and CP.CodigoCustoOrigem_Id  = " & rs("CodigoCustoOrigem_Id")
                sql &= "                   and CP.ProdutoDestino_Id     = AC.Produto_Id"
                sql &= "                   and CP.CodigoCustoDestino_Id = AC.CodigoDeCusto_Id"
                sql &= "                   ) ACD "
                sql &= "      on ACD.Empresa_Id    = ACO.Empresa_Id"
                sql &= "      and ACD.EndEmpresa_Id = ACO.EndEmpresa_Id"
                sql &= "      and ACD.Ano_Id        = ACO.Ano_Id"
                sql &= "      and ACD.Mes_Id        = ACO.Mes_Id"
                sql &= "    Inner Join (SELECT TPM1.Empresa_Id,    TPM1.EndEmpresa_Id,"
                sql &= "                       TPM1.Deposito_Id,   TPM1.EndDeposito_Id,"
                sql &= "                       TPM1.Produto_Id,    TPM1.Data_Id,"
                sql &= "                       TPM1.ValorOficial,  TPM1.ValorMoeda,"
                sql &= "                       TPM1.BaseDeCalculo"
                sql &= "                 FROM TabelaDePrecosDeMercado TPM1"
                sql &= "                Inner Join (SELECT Empresa_Id,"
                sql &= "                                    EndEmpresa_Id,"
                sql &= "                                    Deposito_Id,"
                sql &= "                                    EndDeposito_Id,"
                sql &= "                                    Produto_Id,"
                sql &= "                                    max(Data_Id) as Data_Id"
                sql &= "                                from TabelaDePrecosDeMercado"
                sql &= "                               Group by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id"
                sql &= "                             ) TPM2"
                sql &= "                     on TPM1.Empresa_Id     = TPM2.Empresa_Id"
                sql &= "                   and TPM1.EndEmpresa_Id  = TPM2.EndEmpresa_Id"
                sql &= "                   and TPM1.Deposito_Id    = TPM2.Deposito_Id"
                sql &= "                    and TPM1.EndDeposito_Id = TPM2.EndDeposito_Id"
                sql &= "                    and TPM1.Produto_Id     = TPM2.Produto_Id"
                sql &= "                    and TPM1.Data_Id        = TPM2.Data_Id"
                sql &= "                ) TPM"
                sql &= "       on ACO.Empresa_Id     = TPM.Empresa_Id"
                sql &= "       and ACO.EndEmpresa_id  = TPM.EndEmpresa_id"
                sql &= "       and ACO.Deposito_Id    = TPM.Deposito_Id"
                Sql &= "       And ACO.EndDeposito_Id = TPM.EndDeposito_Id"
                sql &= "       And ACO.Ano_Id         = year(TPM.Data_Id)"
                sql &= "       And ACO.Mes_Id         = Month(TPM.Data_Id)"
                sql &= "       And ACD.Produto_Id     = TPM.Produto_Id "
                Sql &= "       where ACO.Empresa_Id     ='" & Empresa(0) & "'"
                If Left(Empresa(0), 8) = "05272759" OrElse Left(Empresa(0), 8) = "05366261" OrElse Left(Empresa(0), 8) = "38198213" OrElse Left(Empresa(0), 8) = "40938762" OrElse Left(Empresa(0), 8) = "44005444" Then
                    Sql &= "       AND ACO.EndEmpresa_id    = " & Empresa(1)
                Else
                    Sql &= "       AND ACO.EndEmpresa_id    = 0" 'Custo do Curtume está 0 - Furlan - 11/11/2019
                End If
                Sql &= "       AND ACO.Ano_Id           = " & DdlAno.SelectedValue
                sql &= "       And ACO.Mes_Id           = " & pMes
                sql &= "       And ACO.Produto_Id       ='" & rs("ProdutoOrigem_Id") & "'"
                sql &= "       And ACO.CodigoDeCusto_Id = " & rs("CodigoCustoOrigem_Id")
                sql &= "       set @Soma =(Select Sum(VlrTotal) from #Temp)"
                sql &= "   Update #Temp set"
                sql &= "       Soma = @Soma"
                sql &= "   Select  Empresa_Id,"
                sql &= "           EndEmpresa_id,"
                sql &= "           Deposito_Id,"
                sql &= "           EndDeposito_Id,"
                sql &= "           Ano_Id,"
                sql &= "           Mes_Id,"
                sql &= "           ProdutoOrigem,"
                sql &= "           QuantidadeOrigem,"
                sql &= "           VlrProdutoOrigem,"
                sql &= "           ProdutoDestino,"
                sql &= "           CodigoDeCustoDestino,"
                sql &= "           QuantidadeDestino,"
                sql &= "           ValorDeMercado,"
                sql &= "           VlrTotal,"
                sql &= "           (VlrTotal * 100) / Soma as Percentual,"
                sql &= "           (VlrProdutoOrigem *  ((VlrTotal * 100) / Soma)) / 100 as CustoDoProdutoDestino "
                sql &= "      from #Temp"

                Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
                If ds2 IsNot Nothing AndAlso ds2.Tables IsNot Nothing AndAlso ds2.Tables.Count > 0 AndAlso ds2.Tables(0).Rows.Count > 0 Then
                    For Each rsa As DataRow In ds2.Tables(0).Rows
                        Sql = "  Update ApuracaoDeCustos set "
                        sql &= "         ValorDoProduto     =  " & Replace(rsa("CustoDoProdutoDestino"), ",", ".")
                        sql &= " Where Empresa_Id            ='" & rsa("Empresa_Id") & "'"
                        sql &= "   And EndEmpresa_Id         = " & rsa("EndEmpresa_Id")
                        sql &= "   And Deposito_Id           ='" & rsa("Deposito_Id") & "'"
                        sql &= "   And EndDeposito_Id        = " & rsa("EndDeposito_Id")
                        sql &= "   And Ano_Id                = " & rsa("Ano_Id")
                        sql &= "   And Mes_Id                = " & rsa("Mes_Id")
                        sql &= "   And Produto_Id            ='" & rsa("ProdutoDestino") & "'"
                        sql &= "   And CodigoDeCusto_Id      = " & rsa("CodigoDeCustoDestino")
                        sql &= "   And EmpresaDestino_Id     =''"
                        sql &= "   And EndEmpresaDestino_Id  = 0"
                        sql &= "   And DepositoDestino_Id    =''"
                        sql &= "   And EndDepositoDestino_Id = 0"
                        sql &= "   And ProdutoDerivado_Id    =''"
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
        sql &= "        Origem.EndEmpresa_Id, "
        sql &= "        Origem.Deposito_Id, "
        sql &= "        Origem.EndDeposito_Id, "
        sql &= "        Origem.Produto_Id,"
        sql &= "        Origem.CodigoDeCusto_Id, "
        sql &= "        Origem.Ano_Id, "
        sql &= "        Origem.Mes_Id, "
        sql &= "        Origem.EmpresaDestino_Id, "
        sql &= "        Origem.EndEmpresaDestino_Id, "
        sql &= "        Origem.DepositoDestino_Id, "
        sql &= "        Origem.EndDepositoDestino_Id, "
        sql &= "        CASE WHEN Origem.ProdutoDerivado_Id = '000000' THEN '' ELSE Origem.ProdutoDerivado_Id END AS ProdutoDerivado_Id,"
        sql &= "        Origem.CodigoDestino,"
        sql &= "        Origem.ValorDoProduto as ValorOrigem, "
        sql &= "        Origem.ValorDoFrete as FreteOrigem, "
        sql &= "        isnull(Destino.ValorDoProduto, 0) AS ValorDestino"
        sql &= " FROM   ApuracaoDeCustos AS Origem LEFT OUTER JOIN ApuracaoDeCustos AS Destino ON"
        sql &= "        Origem.CodigoDeCusto_Id = Destino.CodigoDestino"
        sql &= "        AND Origem.Produto_Id = Destino.ProdutoDerivado_Id  "
        sql &= "        AND Origem.EndDeposito_Id = Destino.EndDepositoDestino_Id "
        sql &= "        AND Origem.Deposito_Id = Destino.DepositoDestino_Id "
        sql &= "        AND Origem.EndEmpresa_Id = Destino.EndEmpresaDestino_Id "
        sql &= "        AND Origem.Empresa_Id = Destino.EmpresaDestino_Id "
        sql &= "        AND Origem.EmpresaDestino_Id = Destino.Empresa_Id "
        sql &= "        AND Origem.EndEmpresaDestino_Id = Destino.EndEmpresa_Id "
        sql &= "        AND Origem.DepositoDestino_Id = Destino.Deposito_Id "
        sql &= "        AND Origem.EndDepositoDestino_Id = Destino.EndDeposito_Id "
        sql &= "        AND Origem.ProdutoDerivado_Id = Destino.Produto_Id "
        sql &= "        AND Origem.Ano_Id = Destino.Ano_Id "
        sql &= "        AND Origem.Mes_Id = Destino.Mes_Id "
        sql &= "        AND Origem.CodigoDestino = Destino.CodigoDeCusto_Id"
        sql &= " WHERE  (Origem.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        sql &= "        And   (Origem.Ano_Id = " & DdlAno.SelectedValue & ")"
        sql &= "        And   (Origem.Mes_Id = " & pMes & ")"
        sql &= "        And   (Origem.CodigoDeCusto_Id > 495)"
        sql &= "        And   (Origem.CodigoDestino <> 0)"
        sql &= "        And   (Origem.ValorDoProduto > 0)"
        sql &= "        And   (Origem.EmpresaDestino_Id <> '')"
        sql &= "        And   (Origem.DepositoDestino_ID <> '')"
        sql &= "        And   (Origem.ValorDoProduto <> Destino.ValorDoProduto)"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ApuracaoDeCustos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each rs As DataRow In ds.Tables(0).Rows
                Dim Circular = "S"

                Sql = " Declare"
                sql &= " @Exist as varchar(1)"
                sql &= " set @Exist = (select case "
                sql &= "       when exists ("
                sql &= "                 select * from ApuracaoDeCustos "
                sql &= " Where "
                sql &= "     Empresa_Id = '" & rs("EmpresaDestino_Id") & "'"
                sql &= " And EndEmpresa_Id = " & rs("EndEmpresaDestino_Id")
                sql &= " And Deposito_Id = '" & rs("DepositoDestino_Id") & "'"
                sql &= " And EndDeposito_Id = " & rs("EndDepositoDestino_Id")
                sql &= " And Ano_Id = " & rs("Ano_Id")
                sql &= " And Mes_Id = " & rs("Mes_Id")
                sql &= " And Produto_Id = '" & rs("ProdutoDerivado_Id") & "'"
                sql &= " And CodigoDeCusto_Id = " & rs("CodigoDestino")

                sql &= " And EmpresaDestino_Id = '" & rs("Empresa_Id") & "'"
                sql &= " And EndEmpresaDestino_Id = " & rs("EndEmpresa_Id")
                sql &= " And DepositoDestino_Id = '" & rs("Deposito_Id") & "'"
                sql &= " And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                sql &= " And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"

                sql &= ")"
                sql &= "            then  'S'"
                sql &= "             else 'N'"
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

                sql &= ", Etapa"
                sql &= ", Quantidade"
                sql &= ", ValorDoProduto"
                sql &= ", ValorDoFrete"
                sql &= ", ValorAuxiliar"
                sql &= ", ProdutoDestino"
                sql &= ", CodigoDestino"
                sql &= ", Reduzido)"

                sql &= " VALUES('" & rs("EmpresaDestino_Id") & "'"
                sql &= ", " & rs("EndEmpresaDestino_Id")
                sql &= ", '" & rs("DepositoDestino_Id") & "'"
                sql &= ", " & rs("EndDepositoDestino_Id")
                sql &= ", " & rs("Ano_Id")
                sql &= ", " & rs("Mes_Id")
                sql &= ", '" & rs("ProdutoDerivado_Id") & "'"
                sql &= ", " & rs("CodigoDestino")

                sql &= ", '" & rs("Empresa_Id") & "'"
                sql &= ", " & rs("EndEmpresa_Id")
                sql &= ", '" & rs("Deposito_Id") & "'"
                sql &= ", " & rs("EndDeposito_Id")
                sql &= ", '" & rs("Produto_Id") & "'"

                sql &= ", 1"    'Etapa
                sql &= ", 0"    'Quantidade
                sql &= ", " & Replace(rs("ValorOrigem"), ",", ".")
                sql &= ", " & Replace(rs("FreteOrigem"), ",", ".")
                sql &= ", 0"    'Valor Auxiliar

                sql &= ", ''" '& rs("ProdutoDestino") & "'"
                sql &= ", " & rs("CodigoDeCusto_Id")
                sql &= ", ''"
                sql &= ")"

                sql &= "  Update ApuracaoDeCustos set "
                sql &= "         ValorDoProduto = " & Replace(rs("ValorOrigem"), ",", ".")
                sql &= ",        ValorDoFrete = " & Replace(rs("FreteOrigem"), ",", ".")

                sql &= " Where "
                sql &= "     Empresa_Id = '" & rs("EmpresaDestino_Id") & "'"
                sql &= " And EndEmpresa_Id = " & rs("EndEmpresaDestino_Id")
                sql &= " And Deposito_Id = '" & rs("DepositoDestino_Id") & "'"
                sql &= " And EndDeposito_Id = " & rs("EndDepositoDestino_Id")
                sql &= " And Ano_Id = " & rs("Ano_Id")
                sql &= " And Mes_Id = " & rs("Mes_Id")
                sql &= " And Produto_Id = '" & rs("ProdutoDerivado_Id") & "'"
                sql &= " And CodigoDeCusto_Id = " & rs("CodigoDestino")

                sql &= " And EmpresaDestino_Id = '" & rs("Empresa_Id") & "'"
                sql &= " And EndEmpresaDestino_Id = " & rs("EndEmpresa_Id")
                sql &= " And DepositoDestino_Id = '" & rs("Deposito_Id") & "'"
                sql &= " And EndDepositoDestino_Id = " & rs("EndDeposito_Id")
                sql &= " And ProdutoDerivado_Id = '" & rs("Produto_Id") & "'"
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


        Sql = " Delete Razao WHERE left(Empresa_id, 8)='" & Left(Empresa(0), 8) & "'" & vbCrLf
        Sql &= "                And Lote_Id = 7000 "
        Sql &= "                And Year(Movimento_Id)  = " & AnoC & vbCrLf
        Sql &= "                And Month(Movimento_Id)  = " & pMes & vbCrLf

        Sql &= "  SELECT  AP.Empresa_Id, AP.EndEmpresa_Id, "
        Sql &= "          AP.Deposito_Id, AP.EndDeposito_Id, "
        Sql &= "          AP.Ano_Id,  AP.Mes_Id, "
        Sql &= "          AP.Produto_Id, "
        Sql &= "          AP.CodigoDeCusto_Id, PC.Descricao AS TituloDoCusto, "
        Sql &= "          PC.DebitoMercadoria, PC.CreditoMercadoria, "
        Sql &= "          Case when HM.Descricao <> '' then HM.Descricao + ' REF: ' +  REPLICATE('0', 2 - LEN(convert(varchar, AP.Mes_Id))) +  CONVERT(varchar, AP.Mes_Id) + '/' + convert(varchar, AP.Ano_Id) else PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(varchar, AP.Mes_Id))) +  CONVERT(varchar, AP.Mes_Id) + '/' + convert(varchar, AP.Ano_Id) end as HistoricoMercadoria, "
        Sql &= "          PC.DebitoFrete, PC.CreditoFrete, "
        Sql &= "          Case when HF.Descricao <> '' then HF.Descricao + ' REF: ' +  REPLICATE('0', 2 - LEN(convert(varchar, AP.Mes_Id))) +  CONVERT(varchar, AP.Mes_Id) + '/' + convert(varchar, AP.Ano_Id) else PC.Descricao + ' REF: ' + REPLICATE('0', 2 - LEN(convert(varchar, AP.Mes_Id))) +  CONVERT(varchar, AP.Mes_Id) + '/' + convert(varchar, AP.Ano_Id) end as HistoricoFrete, "
        Sql &= "          AP.EmpresaDestino_Id,   AP.EndEmpresaDestino_Id, "
        Sql &= "          AP.DepositoDestino_Id,  AP.EndDepositoDestino_Id, "
        Sql &= "          AP.ProdutoDerivado_Id,  AP.Quantidade, "
        Sql &= "          AP.ValorDoProduto,      AP.ValorDoFrete "

        Sql &= " FROM   ApuracaoDeCustos AS AP INNER JOIN"
        Sql &= "        PlanoDeCustos AS PC ON AP.CodigoDeCusto_Id = PC.Codigo_Id LEFT OUTER JOIN"
        Sql &= "        Historicos AS HF ON PC.HistoricoFrete = HF.Historico_Id LEFT OUTER JOIN"
        Sql &= "        Historicos AS HM ON PC.HistoricoMercadoria = HM.Historico_Id"
        Sql &= " WHERE	left(Empresa_id, 8)='" & Left(Empresa(0), 8) & "'"
        Sql &= "         And AP.Ano_Id = " & AnoC & " "
        Sql &= "         And AP.Mes_Id = " & pMes & " "
        Sql &= "         And (PC.DebitoMercadoria <> '' OR PC.CreditoMercadoria <> '')"

        i = 0
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows

            If Dr("ValorDoProduto") > 0 Then
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, Dr("ValorDoProduto"), 0, 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, Dr("ValorDoProduto"), 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
            End If

            If Dr("ValorDoProduto") < 0 Then
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, (Dr("ValorDoProduto") * -1), 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoMercadoria"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, (Dr("ValorDoProduto") * -1), 0, 0, 0, Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
            End If



            If Dr("ValorDoFrete") > 0 Then
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, Dr("ValorDoFrete"), 0, Dr("MoedaValorDoFrete"), 0, Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, Dr("ValorDoFrete"), 0, Dr("MoedaValorDoFrete"), Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
            End If

            If Dr("ValorDoFrete") < 0 Then
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("DebitoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, 0, (Dr("ValorDoFrete") * -1), 0, (Dr("MoedaValorDoFrete") * -1), Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
                i += 1
                Sql = SqlRazao(Dr("Empresa_Id"), Dr("EndEmpresa_Id"), Dia, 7000, i, Dr("CreditoFrete"), Dr("EmpresaDestino_Id"), Dr("EndEmpresaDestino_Id"), Dr("Produto_Id"), 3, Dia, (Dr("ValorDoFrete") * -1), 0, (Dr("MoedaValorDoFrete") * -1), 0, Dr("HistoricoMercadoria"), 0, "R")
                If Sql.Length > 0 Then
                    Array.Add(Sql)
                End If
            End If

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
                'ChkResultado.Items.Add(New ListItem(" FIM - Contabilização Automática"))
            Else
                ChkResultado.Items.Add(New ListItem(" FIM - Erro no Processo de Contabilização Automática"))
            End If

        Next

        ChkResultado.Items.Add(New ListItem(" FIM - Contabilização Automática"))
        
    End Sub

    Function SqlRazao(ByVal CnpjEmpresa As String, ByVal EndEmpresa As Integer, ByVal DataMovimento As String, ByVal Lote As String, ByVal Sequencia As String, ByVal Conta As String, ByVal Cliente As String, ByVal EndCliente As Integer, ByVal Produto As String, ByVal Indexador As String, ByVal DataMoeda As String, ByVal DebitoOficial As Double, ByVal CreditoOficial As Double, ByVal DebitoMoeda As Double, ByVal CreditoMoeda As Double, ByVal Historico As String, ByVal Moeda As String, ByVal PrevistoRealizado As String) As String
        Dim TemCliente As String = "N"
        Dim TemProduto As String = "N"
        Dim sqll As String
        Dim sql As String
        sql = ""

        sqll = "Select Top 1 * from PlanoDeContas where Left(Conta_Id, 7) = '" & Microsoft.VisualBasic.Left(Conta, 7) & "'"

        For Each dr As DataRow In banco.ConsultaDataSet(sqll, "Consulta").Tables(0).Rows
            TemCliente = dr("Cliente")
            TemProduto = dr("Produto")


            Sql = " INSERT INTO Razao (Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Produto,"
            sql &= " Movimento_Id,Lote_Id,Sequencia_Id,Indexador, DataMoeda,"
            sql &= " DebitoOficial,CreditoOficial,DebitoMoeda,CreditoMoeda,Historico,PrevistoRealizado) Values ("
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
            sql &= ", '" & PrevistoRealizado & "')"
        Next

        Return sql

    End Function

    Private Sub Contabiliza(ByVal pMes As Integer)
        Dim AnoC As Integer
        AnoC = DdlAno.SelectedValue
        Dim Array As New ArrayList
        Dim Sql As String

        Sql = " if(object_id('tempdb..#Temp') is not null) begin drop table #Temp end; " & vbCrLf & _
              " Delete Razao where lote_id='7000' and movimento_id='" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & " ' and left(Empresa_id,8)='" & Left(Empresa(0), 8) & "'; " & vbCrLf & _
              " SELECT * " & vbCrLf & _
              "   into #temp" & vbCrLf & _
              "   from (  " & vbCrLf & _
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf & _
              "                 PlanoDeCustos.DebitoMercadoria as Conta_Id, " & vbCrLf & _
              "                 case " & vbCrLf & _
              "                   when len(DebitoMercadoria) = 7    " & vbCrLf & _
              "                     then EmpresaDestino_Id " & vbCrLf & _
              "                     else ''  " & vbCrLf & _
              "                 end Cliente_Id,  " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(DebitoMercadoria) = 7  " & vbCrLf & _
              "                     then EndEmpresaDestino_Id " & vbCrLf & _
              "                     else 0 " & vbCrLf & _
              "                 end EndCliente_Id,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf & _
              "                 '7000' as  Lote, " & vbCrLf & _
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf & _
              "                 0 as Titulo, " & vbCrLf & _
              "                 3 as Indexador,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoProduto AS DebitoOficial,  " & vbCrLf & _
              "                 0.00 as CreditoOficial, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoProduto AS DebitoMoeda,  " & vbCrLf & _
              "                 0.00 as CreditoMoeda, " & vbCrLf & _
              "                 convert(nvarchar,PlanoDeCustos.HistoricoMercadoria) +' '+ HMercadoria.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf & _
              "                 0 as CentroDeCustos, " & vbCrLf & _
              "                 'P' as PrevistoRealizado, " & vbCrLf & _
              "                 0 as Serie_Nf, " & vbCrLf & _
              "                 0 as Numero_Nf, " & vbCrLf & _
              "                 NULL as Pedido, " & vbCrLf & _
              "                 'E' as EntradaSaida_Nf " & vbCrLf & _
              "            FROM ApuracaoDeCustos " & vbCrLf & _
              "  		  INNER JOIN Clientes AS Empresa " & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf & _
              " 		  INNER JOIN Clientes AS Deposito " & vbCrLf & _
              "              ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id   " & vbCrLf & _
              "            LEFT JOIN Clientes AS Destino " & vbCrLf & _
              "              ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id    " & vbCrLf & _
              "  		  INNER JOIN Produtos " & vbCrLf & _
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
              "           INNER JOIN PlanoDeCustos" & vbCrLf & _
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf & _
              "            LEFT JOIN TabelaDePrecosDeMercado" & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id  " & vbCrLf & _
              "             AND month(Data_Id)                  = '" & pMes & "' AND year(Data_Id)= '" & AnoC & "'" & vbCrLf & _
              "            LEFT JOIN Historicos AS HMercadoria" & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf & _
              "            LEFT JOIN Historicos AS HFrete " & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf & _
              "           WHERE ApuracaoDeCustos.Ano_Id             = '" & AnoC & "'" & vbCrLf & _
              "             AND ApuracaoDeCustos.Mes_Id             = '" & pMes & "'" & vbCrLf & _
              "             And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf & _
              "    UNION ALL" & vbCrLf & _
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf & _
              "                 PlanoDeCustos.CreditoMercadoria as Conta_Id, " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(CreditoMercadoria) = 7 " & vbCrLf & _
              "                     then EmpresaDestino_Id " & vbCrLf & _
              "                     else ''  " & vbCrLf & _
              "                 end Cliente_Id,  " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(CreditoMercadoria) = 7  " & vbCrLf & _
              "                     then EndEmpresaDestino_Id " & vbCrLf & _
              "                     else 0 " & vbCrLf & _
              "                 end EndCliente_Id,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf & _
              "                 '7000' as  Lote, " & vbCrLf & _
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf & _
              "                 0 as Titulo, " & vbCrLf & _
              "                 3 as Indexador,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf & _
              "                 0.00 as DebitoOficial, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoProduto AS CreditoOficial,  " & vbCrLf & _
              "                 0.00 as DebitoMoeda, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoProduto AS CreditoMoeda,  " & vbCrLf & _
              "                 convert(nvarchar,PlanoDeCustos.HistoricoMercadoria) +' '+ HMercadoria.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf & _
              "                 0 as CentroDeCustos, " & vbCrLf & _
              "                 'P' as PrevistoRealizado, " & vbCrLf & _
              "                 0 as Serie_Nf, " & vbCrLf & _
              "                 0 as Numero_Nf, " & vbCrLf & _
              "                 NULL as Pedido, " & vbCrLf & _
              "                 'S' as EntradaSaida_Nf " & vbCrLf & _
              "            FROM ApuracaoDeCustos   " & vbCrLf & _
              "  		  INNER JOIN Clientes AS Empresa" & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id   " & vbCrLf & _
              "  		  INNER JOIN Clientes AS Deposito" & vbCrLf & _
              "              ON ApuracaoDeCustos.Deposito_Id    = Deposito.Cliente_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id   " & vbCrLf & _
              "            LEFT JOIN Clientes AS Destino" & vbCrLf & _
              "              ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id    " & vbCrLf & _
              "  		  INNER JOIN Produtos " & vbCrLf & _
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
              "           INNER JOIN PlanoDeCustos" & vbCrLf & _
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf & _
              "            Left JOIN TabelaDePrecosDeMercado" & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id  " & vbCrLf & _
              "             AND month(Data_Id)                  = " & pMes & " AND year(Data_Id)= '" & AnoC & "'  " & vbCrLf & _
              "            LEFT JOIN Historicos AS HMercadoria" & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf & _
              "            LEFT JOIN Historicos AS HFrete" & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf & _
              "           WHERE ApuracaoDeCustos.Ano_Id             = '" & AnoC & "'" & vbCrLf & _
              "             AND ApuracaoDeCustos.Mes_Id             = '" & pMes & "'" & vbCrLf & _
              "             And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf & _
              "           Union ALL" & vbCrLf & _
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf & _
              "                 PlanoDeCustos.DebitoFrete as Conta_Id, " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(DebitoFrete) = 7    " & vbCrLf & _
              "                     then EmpresaDestino_Id " & vbCrLf & _
              "                     else ''  " & vbCrLf & _
              "                 end Cliente_Id,  " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(DebitoFrete) = 7  " & vbCrLf & _
              "                     then EndEmpresaDestino_Id " & vbCrLf & _
              "                     else 0 " & vbCrLf & _
              "                 end EndCliente_Id,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf & _
              "                 '7000' as  Lote, " & vbCrLf & _
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf & _
              "                 0 as Titulo, " & vbCrLf & _
              "                 3 as Indexador,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoFrete AS DebitoOficial,  " & vbCrLf & _
              "                 0.00 as CreditoOficial, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoFrete AS DebitoMoeda,  " & vbCrLf & _
              "                 0.00 as CreditoMoeda, " & vbCrLf & _
              "                 convert(nvarchar,PlanoDeCustos.HistoricoFrete) +' '+ HFrete.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf & _
              "                 0 as CentroDeCustos, " & vbCrLf & _
              "                 'P' as PrevistoRealizado, " & vbCrLf & _
              "                 0 as Serie_Nf, " & vbCrLf & _
              "                 0 as Numero_Nf, " & vbCrLf & _
              "                 NULL as Pedido, " & vbCrLf & _
              "                 'E' as EntradaSaida_Nf " & vbCrLf & _
              "            FROM ApuracaoDeCustos   " & vbCrLf & _
              "  		  INNER JOIN Clientes AS Empresa " & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf & _
              "  	  	  INNER JOIN Clientes AS Deposito " & vbCrLf & _
              "              ON ApuracaoDeCustos.Deposito_Id    = Deposito.Cliente_Id" & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id" & vbCrLf & _
              "            LEFT JOIN Clientes AS Destino " & vbCrLf & _
              "              ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id" & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id" & vbCrLf & _
              "  	  	  INNER JOIN Produtos" & vbCrLf & _
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
              "           INNER JOIN PlanoDeCustos" & vbCrLf & _
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf & _
              "            Left JOIN TabelaDePrecosDeMercado " & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id  " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id   " & vbCrLf & _
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id  " & vbCrLf & _
              "             AND month(Data_Id)                  = '" & pMes & "' AND year(Data_Id)= '" & AnoC & "'  " & vbCrLf & _
              "            LEFT JOIN Historicos AS HMercadoria " & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf & _
              "            LEFT JOIN Historicos AS HFrete" & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf & _
              "           WHERE ApuracaoDeCustos.Ano_Id = '" & AnoC & "'" & vbCrLf & _
              "             AND ApuracaoDeCustos.Mes_Id = '" & pMes & "'" & vbCrLf & _
              "             AND left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf & _
              "           UNION ALL " & vbCrLf & _
              "          SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id,  " & vbCrLf & _
              "                 PlanoDeCustos.CreditoFrete as Conta_Id, " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(CreditoFrete) = 7    " & vbCrLf & _
              "                     then EmpresaDestino_Id " & vbCrLf & _
              "                     else ''  " & vbCrLf & _
              "                 end Cliente_Id,  " & vbCrLf & _
              "                 case  " & vbCrLf & _
              "                   when len(CreditoFrete) = 7  " & vbCrLf & _
              "                     then EndEmpresaDestino_Id " & vbCrLf & _
              "                     else 0 " & vbCrLf & _
              "                 end EndCliente_Id,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as Movimento_Id, " & vbCrLf & _
              "                 '7000' as  Lote, " & vbCrLf & _
              "                 ApuracaoDeCustos.Produto_Id,  " & vbCrLf & _
              "                 0 as Titulo, " & vbCrLf & _
              "                 3 as Indexador,  " & vbCrLf & _
              "                 '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' as DataMoeda, " & vbCrLf & _
              "                 0.00 as DebitoOficial, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoFrete AS CreditoOficial, " & vbCrLf & _
              "                 0.00 as DebitoMoeda, " & vbCrLf & _
              "                 ApuracaoDeCustos.ValorDoFrete AS CreditoMoeda,  " & vbCrLf & _
              "                 convert(nvarchar,PlanoDeCustos.HistoricoFrete) +' '+ HFrete.Descricao +' '+ '" & pMes & "/" & AnoC & " - ' + convert(nvarchar,PlanoDeCustos.Codigo_id) + ' ' + PlanoDeCustos.Descricao As Historico, " & vbCrLf & _
              "                 0 as CentroDeCustos, " & vbCrLf & _
              "                 'P' as PrevistoRealizado, " & vbCrLf & _
              "                 0 as Serie_Nf, " & vbCrLf & _
              "                 0 as Numero_Nf, " & vbCrLf & _
              "                 NULL as Pedido, " & vbCrLf & _
              "                 'S' as EntradaSaida_Nf " & vbCrLf & _
              "            FROM ApuracaoDeCustos " & vbCrLf & _
              "  	 	  INNER JOIN Clientes AS Empresa " & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id = Empresa.Cliente_Id " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf & _
              "  		  INNER JOIN Clientes AS Deposito " & vbCrLf & _
              "              ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id " & vbCrLf & _
              "            LEFT JOIN Clientes AS Destino " & vbCrLf & _
              "              ON ApuracaoDeCustos.EmpresaDestino_Id = Destino.Cliente_Id " & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id " & vbCrLf & _
              "  		  INNER JOIN Produtos  " & vbCrLf & _
              "              ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
              "           INNER JOIN PlanoDeCustos " & vbCrLf & _
              "              ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id  " & vbCrLf & _
              "            Left JOIN TabelaDePrecosDeMercado " & vbCrLf & _
              "              ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id" & vbCrLf & _
              "             AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id" & vbCrLf & _
              "             AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id" & vbCrLf & _
              "             AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id" & vbCrLf & _
              "             AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id" & vbCrLf & _
              "             AND month(Data_Id)                  = '" & pMes & "' AND year(Data_Id)= '" & AnoC & "'" & vbCrLf & _
              "            LEFT JOIN Historicos AS HMercadoria " & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id " & vbCrLf & _
              "            LEFT JOIN Historicos AS HFrete " & vbCrLf & _
              "              ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id " & vbCrLf & _
              "           WHERE ApuracaoDeCustos.Ano_Id              = '" & AnoC & "' AND ApuracaoDeCustos.Mes_Id = '" & pMes & "'" & vbCrLf & _
              "              And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'   " & vbCrLf & _
              "        ) as consulta; " & vbCrLf & _
              "  Delete #temp " & vbCrLf & _
              "  Where Conta_Id  = ''" & vbCrLf & _
              "    OR (    DebitoOficial  = 0" & vbCrLf & _
              "        AND CreditoOficial = 0" & vbCrLf & _
              "        AND DebitoMoeda    = 0" & vbCrLf & _
              "        AND CreditoMoeda   = 0" & vbCrLf & _
              "       ) ; " & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              " Update #Temp  " & vbCrLf & _
              "    Set Produto_Id = ''" & vbCrLf & _
              "   From #temp " & vbCrLf & _
              "   Left join PlanoDeContas PC  " & vbCrLf & _
              "     on #Temp.Empresa_Id    = PC.Empresa_Id  " & vbCrLf & _
              "    And #Temp.EndEmpresa_Id = PC.EndEmpresa_Id  " & vbCrLf & _
              "    AND #Temp.Conta_Id      = PC.Conta_Id " & vbCrLf & _
              "  Where PC.Produto = 'N'; " & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "Update #Temp" & vbCrLf & _
              "    Set debitooficial = creditooficial * -1," & vbCrLf & _
              "        debitomoeda   = creditomoeda * -1" & vbCrLf & _
              "where creditooficial < 0" & vbCrLf & _
              "" & vbCrLf & _
              "Update #Temp" & vbCrLf & _
              "    Set creditooficial = 0," & vbCrLf & _
              "        creditomoeda   = 0" & vbCrLf & _
              "where creditooficial < 0" & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "Update #Temp" & vbCrLf & _
              "    Set creditooficial = debitooficial * -1," & vbCrLf & _
              "        creditomoeda   = debitomoeda * -1" & vbCrLf & _
              "where debitooficial < 0" & vbCrLf & _
              "" & vbCrLf & _
              "Update #Temp" & vbCrLf & _
              "    Set debitooficial = 0," & vbCrLf & _
              "        debitomoeda   = 0" & vbCrLf & _
              "where debitooficial < 0" & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "" & vbCrLf & _
              "  INSERT INTO Razao " & vbCrLf & _
              "          (Empresa_Id,  " & vbCrLf & _
              "           EndEmpresa_Id,  " & vbCrLf & _
              "           Conta_Id, " & vbCrLf & _
              "           Cliente_Id,  " & vbCrLf & _
              "           EndCliente_Id,  " & vbCrLf & _
              "           Movimento_Id,  " & vbCrLf & _
              "           Lote_Id,  " & vbCrLf & _
              "           Sequencia_Id,  " & vbCrLf & _
              "           Produto,  " & vbCrLf & _
              "           Titulo,  " & vbCrLf & _
              "           Indexador, " & vbCrLf & _
              "           DataMoeda,  " & vbCrLf & _
              "           DebitoOficial,  " & vbCrLf & _
              "           CreditoOficial,  " & vbCrLf & _
              "           DebitoMoeda,  " & vbCrLf & _
              "           CreditoMoeda, " & vbCrLf & _
              "           Historico,  " & vbCrLf & _
              "           custo,  " & vbCrLf & _
              "           PrevistoRealizado,Serie_Nf,Numero_Nf,Pedido,EntradaSaida_Nf)  " & vbCrLf & _
              "           (  " & vbCrLf & _
              "            SELECT " & vbCrLf & _
              "            Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Movimento_Id,Lote,  " & vbCrLf & _
              "            ROW_NUMBER() OVER (PARTITION BY Movimento_Id,Lote order by Movimento_Id) AS Sequencia_Id,  " & vbCrLf & _
              "            Produto_Id, Titulo,Indexador,DataMoeda,abs(DebitoOficial),abs(CreditoOficial), " & vbCrLf & _
              "            abs(DebitoMoeda),abs(CreditoMoeda),  " & vbCrLf & _
              "            Historico,CentroDeCustos,PrevistoRealizado, Serie_Nf, Numero_Nf, Pedido, EntradaSaida_Nf  " & vbCrLf & _
              "            from  #temp )" & vbCrLf

        Array.Add(Sql)

        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
            ChkResultado.Items.Add(New ListItem(" FIM - Contabilização Automática"))
        Else
            ChkResultado.Items.Add(New ListItem(" FIM - Erro no Processo de Contabilização Automática"))
        End If

    End Sub
	
    Private Sub ContabilizaCurtume(ByVal pMes As Integer)
        Dim AnoC As Integer
        AnoC = DdlAno.SelectedValue
        Dim Array As New ArrayList
        Dim Sql As String

        Sql = " if(object_id('tempdb..#Temp') is not null) begin drop table #Temp end; " & vbCrLf &
              " Delete Razao where lote_id='7000' and movimento_id='" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & " ' and left(Empresa_id,8)='" & Left(Empresa(0), 8) & "'; " & vbCrLf &
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
              "             And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'     " & vbCrLf &
              "             AND ApuracaoDeCustos.CodigoDeCusto_Id <> 402    " & vbCrLf &
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
              "             And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "'    " & vbCrLf &
              "             AND ApuracaoDeCustos.CodigoDeCusto_Id <> 402    " & vbCrLf &
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
              "             AND ApuracaoDeCustos.CodigoDeCusto_Id <> 402   " & vbCrLf &
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
              "              And left(ApuracaoDeCustos.Empresa_ID,8) = '" & Left(Empresa(0), 8) & "' " & vbCrLf &
              "              AND ApuracaoDeCustos.CodigoDeCusto_Id <> 402  " & vbCrLf &
              " UNION ALL                                                                                                        " & vbCrLf &
              "             SELECT                                                                                               " & vbCrLf &
              "               ApuracaoDeCustos.Empresa_Id,                                                                       " & vbCrLf &
              "               ApuracaoDeCustos.EndEmpresa_Id,                                                                    " & vbCrLf &
              "               '101060105' AS Conta_Id,                                                                           " & vbCrLf &
              "               '' Cliente_Id,                                                                                     " & vbCrLf &
              "              '' EndCliente_Id,                                                                                   " & vbCrLf &
              "               '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' AS Movimento_Id,              " & vbCrLf &
              "               '7000' AS Lote,                                                                                    " & vbCrLf &
              "               ApuracaoDeCustos.Produto,                                                                          " & vbCrLf &
              "               0 AS Titulo,                                                                                       " & vbCrLf &
              "               3 AS Indexador,                                                                                    " & vbCrLf &
              "               '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' AS DataMoeda,			     " & vbCrLf &
              "			      ApuracaoDeCustos.ValorProporcional AS CreditoOficial,                                                                                 " & vbCrLf &
              "			      0.00 AS DebitoOficial,                                                  " & vbCrLf &
              "			      ApuracaoDeCustos.ValorProporcional AS CreditoMoeda,                                                                                   " & vbCrLf &
              "			      0.00 AS DebitoMoeda,                                                    " & vbCrLf &
              "               CONVERT(nvarchar, PlanoDeCustos.HistoricoMercadoria) + ' ' + HMercadoria.Descricao + ' ' + '" & pMes & " / " & AnoC & " - ' + CONVERT(nvarchar, PlanoDeCustos.Codigo_id) + ' ' +   PlanoDeCustos.Descricao  AS Historico, " & vbCrLf &
              "               0 AS CentroDeCustos,                                                                               " & vbCrLf &
              "               'P' AS PrevistoRealizado,                                                                          " & vbCrLf &
              "               0 AS Serie_Nf,                                                                                     " & vbCrLf &
              "               0 AS Numero_Nf,                                                                                    " & vbCrLf &
              "               NULL AS Pedido,                                                                                    " & vbCrLf &
              "               'E' AS EntradaSaida_Nf                                                                             " & vbCrLf &
              "             FROM FN_CalcularCustosIndiretos('" & Left(Empresa(0), 8) & "'," & AnoC & "," & pMes & ") as ApuracaoDeCustos           " & vbCrLf &
              "             INNER JOIN Clientes AS Empresa                                                                       " & vbCrLf &
              "               ON ApuracaoDeCustos.Empresa_Id = Empresa.Cliente_Id                                                " & vbCrLf &
              "               AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id                                           " & vbCrLf &
              "             INNER JOIN Clientes AS Deposito                                                                      " & vbCrLf &
              "               ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id                                              " & vbCrLf &
              "               AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id                                         " & vbCrLf &
              "             INNER JOIN Produtos                                                                                  " & vbCrLf &
              "               ON ApuracaoDeCustos.Produto = Produtos.Produto_Id                                                  " & vbCrLf &
              "             CROSS APPLY (                                                                                        " & vbCrLf &
              "                             SELECT TOP 1 * FROM PlanoDeCustos                                                    " & vbCrLf &
              "                                 WHERE                                                                            " & vbCrLf &
              "                                 ApuracaoDeCustos.Codigo_Id = PlanoDeCustos.Codigo_Id                             " & vbCrLf &
              "                         ) AS PlanoDeCustos                                                                       " & vbCrLf &
              "             LEFT JOIN Historicos AS HMercadoria                                                                  " & vbCrLf &
              "               ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id                                    " & vbCrLf &
              "             LEFT JOIN Historicos AS HFrete                                                                       " & vbCrLf &
              "               ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id                                              " & vbCrLf &
              " UNION ALL                                                                                                        " & vbCrLf &
              "             SELECT                                                                                               " & vbCrLf &
              "               ApuracaoDeCustos.Empresa_Id,                                                                       " & vbCrLf &
              "               ApuracaoDeCustos.EndEmpresa_Id,                                                                    " & vbCrLf &
              "               CONCAT(ApuracaoDeCustos.Conta,'99') AS Conta_Id,                                                   " & vbCrLf &
              "               '' Cliente_Id,                                                                                     " & vbCrLf &
              "              '' EndCliente_Id,                                                                                   " & vbCrLf &
              "               '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' AS Movimento_Id,              " & vbCrLf &
              "               '7000' AS Lote,                                                                                    " & vbCrLf &
              "               ApuracaoDeCustos.Produto,                                                                          " & vbCrLf &
              "               0 AS Titulo,                                                                                       " & vbCrLf &
              "               3 AS Indexador,                                                                                    " & vbCrLf &
              "               '" & AnoC & "/" & pMes & "/" & DateTime.DaysInMonth(AnoC, pMes) & "' AS DataMoeda,			     " & vbCrLf &
              "			  0.00 AS DebitoOficial,                                                                                 " & vbCrLf &
              "			  ApuracaoDeCustos.ValorProporcional AS CreditoOficial,                                                  " & vbCrLf &
              "			  0.00 AS DebitoMoeda,                                                                                   " & vbCrLf &
              "			  ApuracaoDeCustos.ValorProporcional AS CreditoMoeda,                                                    " & vbCrLf &
              "               CONVERT(nvarchar, PlanoDeCustos.HistoricoMercadoria) + ' ' + HMercadoria.Descricao + ' ' + '" & pMes & " / " & AnoC & " - ' + CONVERT(nvarchar, PlanoDeCustos.Codigo_id) + ' ' +   PlanoDeCustos.Descricao  AS Historico, " & vbCrLf &
              "               0 AS CentroDeCustos,                                                                               " & vbCrLf &
              "               'P' AS PrevistoRealizado,                                                                          " & vbCrLf &
              "               0 AS Serie_Nf,                                                                                     " & vbCrLf &
              "               0 AS Numero_Nf,                                                                                    " & vbCrLf &
              "               NULL AS Pedido,                                                                                    " & vbCrLf &
              "               'S' AS EntradaSaida_Nf                                                                             " & vbCrLf &
              "             FROM FN_CalcularCustosIndiretos('" & Left(Empresa(0), 8) & "'," & AnoC & "," & pMes & ") as ApuracaoDeCustos           " & vbCrLf &
              "             INNER JOIN Clientes AS Empresa                                                                       " & vbCrLf &
              "               ON ApuracaoDeCustos.Empresa_Id = Empresa.Cliente_Id                                                " & vbCrLf &
              "               AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id                                           " & vbCrLf &
              "             INNER JOIN Clientes AS Deposito                                                                      " & vbCrLf &
              "               ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id                                              " & vbCrLf &
              "               AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id                                         " & vbCrLf &
              "             INNER JOIN Produtos                                                                                  " & vbCrLf &
              "               ON ApuracaoDeCustos.Produto = Produtos.Produto_Id                                                  " & vbCrLf &
              "             CROSS APPLY (                                                                                        " & vbCrLf &
              "                             SELECT TOP 1 * FROM PlanoDeCustos                                                    " & vbCrLf &
              "                                 WHERE                                                                            " & vbCrLf &
              "                                 ApuracaoDeCustos.Codigo_Id = PlanoDeCustos.Codigo_Id                             " & vbCrLf &
              "                         ) AS PlanoDeCustos                                                                       " & vbCrLf &
              "             LEFT JOIN Historicos AS HMercadoria                                                                  " & vbCrLf &
              "               ON PlanoDeCustos.HistoricoMercadoria = HMercadoria.Historico_Id                                    " & vbCrLf &
              "             LEFT JOIN Historicos AS HFrete                                                                       " & vbCrLf &
              "               ON PlanoDeCustos.HistoricoFrete = HFrete.Historico_Id                                              " & vbCrLf &
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

        Dim Sql As String = "SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito AS Deposito_Id, NotasFiscaisXItens.EndDeposito AS EndDeposito_Id," & vbCrLf & _
              "       NotasFiscais.Cliente_Id AS EmpresaDestino_Id, NotasFiscais.EndCliente_Id AS EndEmpresaDestino_Id, NotasFiscais.Cliente_Id AS DepositoDestino_Id," & vbCrLf & _
              "       NotasFiscais.EndCliente_Id AS EndDepositoDestino_Id," & vbCrLf & _
              "       case" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '101010001' then '101010007'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '101010003' then '101010007'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '101080002' then '101080001'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '101060002' then '101060001'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '701010004' then '701010001'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '404010002' then '404010001'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '404010003' then '404010001'" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '404010004' then '404010001'" & vbCrLf & _
              "         else NotasFiscaisXItens.Produto_Id" & vbCrLf & _
              "       end Produto_id," & vbCrLf & _
              "       SubOperacoes.ApuracaoDeCustos AS Placus_Id," & vbCrLf & _
              "       SubOperacoes.ApuracaodeCustosContraPartida AS PlacusContra_Id," & vbCrLf & _
              "       SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade," & vbCrLf & _
              "       convert(numeric(18,2),0) AS ValorDoProduto" & vbCrLf & _
              "  FROM NotasFiscais" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf & _
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf & _
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes" & vbCrLf & _
              "    ON SubOperacoes.Operacao_Id = NotasFiscaisXItens.Operacao  " & vbCrLf & _
              "   AND SubOperacoes.SubOperacoes_Id = NotasFiscaisXItens.SubOperacao" & vbCrLf & _
              " INNER JOIN PlanoDeCustos " & vbCrLf & _
              "    ON PlanoDeCustos.Codigo_Id = SubOperacoes.ApuracaoDeCustos " & vbCrLf & _
              " INNER JOIN Produtos " & vbCrLf & _
              "    ON Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id " & vbCrLf & _
              " INNER JOIN GruposDeEstoques" & vbCrLf & _
              "    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf & _
              " WHERE NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%'"

        Sql &= "   AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue & vbCrLf & _
               "   AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes & vbCrLf

        Sql &= "   AND (SubOperacoes.ApuracaoDeCustosContraPartida > 0)" & vbCrLf & _
               "   And  NotasFiscais.Situacao  = 1" & vbCrLf & _
               "   And  GruposDeEstoques.custo = 1" & vbCrLf & _
               "   And  PlanoDeCustos.Classe   ='" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf & _
               " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito," & vbCrLf & _
               "          NotasFiscaisXItens.EndDeposito, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), " & vbCrLf & _
               "          case" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '101010001' then '101010007'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '101010003' then '101010007'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '101080002' then '101080001'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '101060002' then '101060001'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '701010004' then '701010001'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '404010002' then '404010001'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '404010003' then '404010001'" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '404010004' then '404010001'" & vbCrLf & _
               "            else NotasFiscaisXItens.Produto_Id" & vbCrLf & _
               "         end," & vbCrLf & _
               "         SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro," & vbCrLf & _
               "         NotasFiscaisXItens.EndDepositoTerceiro , SubOperacoes.ApuracaodeCustosContraPartida, NotasFiscais.EntradaSaida_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare" & vbCrLf & _
                  " @Exist as varchar(1)" & vbCrLf & _
                  " set @Exist = (select case " & vbCrLf & _
                  "                        when exists (" & vbCrLf & _
                  "                                      select 1 " & vbCrLf & _
                  "                                        from ApuracaoDeCustos " & vbCrLf & _
                  "                                       Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "                                         And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "                                         And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "                                         And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "                                         And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "                                         And Mes_Id                = " & pMes & vbCrLf & _
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                         And CodigoDeCusto_Id      = " & Dr("Placus_Id") & vbCrLf & _
                  "                                    )" & vbCrLf & _
                  "                           then 'S'" & vbCrLf & _
                  "                           else 'N'" & vbCrLf & _
                  "                       end) ;" & vbCrLf & _
                  " if @Exist = 'N' " & vbCrLf & _
                  "  INSERT INTO ApuracaoDeCustos ( " & vbCrLf & _
                  "  Empresa_Id" & vbCrLf & _
                  ", EndEmpresa_Id" & vbCrLf & _
                  ", Deposito_Id" & vbCrLf & _
                  ", EndDeposito_Id" & vbCrLf & _
                  ", Ano_Id" & vbCrLf & _
                  ", Mes_Id" & vbCrLf & _
                  ", Produto_Id" & vbCrLf & _
                  ", CodigoDeCusto_Id" & vbCrLf & _
                  ", EmpresaDestino_Id" & vbCrLf & _
                  ", EndEmpresaDestino_Id" & vbCrLf & _
                  ", DepositoDestino_Id" & vbCrLf & _
                  ", EndDepositoDestino_Id" & vbCrLf & _
                  ", ProdutoDerivado_Id" & vbCrLf & _
                  ", Quantidade" & vbCrLf & _
                  ", ValorDoProduto" & vbCrLf & _
                  ", ValorDoFrete" & vbCrLf & _
                  ", ValorAuxiliar" & vbCrLf & _
                  ", ProdutoDestino" & vbCrLf & _
                  ", CodigoDestino)" & vbCrLf & _
                  " VALUES('" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndEmpresa_Id") & vbCrLf & _
                  ",'" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndDeposito_Id") & vbCrLf & _
                  ", " & DdlAno.SelectedValue & vbCrLf & _
                  ", " & pMes & vbCrLf & _
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf & _
                  ", " & Dr("Placus_Id") & vbCrLf & _
                  ",'" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  ",'" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf & _
                  ", " & Str(Dr("Quantidade")) & vbCrLf & _
                  ", 0" & vbCrLf & _
                  ", 0" & vbCrLf & _
                  ", 0" & vbCrLf & _
                  ",''" & vbCrLf & _
                  ", " & Dr("PlacusContra_Id") & vbCrLf & _
                  ")" & vbCrLf & _
                  " Else" & vbCrLf & _
                  "   Update ApuracaoDeCustos set " & vbCrLf & _
                  "   Quantidade = " & Str(Dr("Quantidade")) & vbCrLf & _
                  " Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "   And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "   And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "   And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "   And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "   And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "   And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "   And Mes_Id                = " & pMes & vbCrLf & _
                  "   And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
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