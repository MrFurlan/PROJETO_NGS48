Imports System.Data
Imports System.Data.SqlClient
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CalculoDosJurosSobreAdiantamentos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim ds As DataSet
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Adiantamento)
        If Not IsPostBack AndAlso IsConnect Then
            If Funcoes.VerificaPermissao("CalculoDosJurosSobreAdiantamentos", "ACESSAR") Then
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                CalData.SelectedDate = Format(Today, "dd/MM/yyyy")
                CargaUnidade()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Adiantamento.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Nome, Clientes.Cidade, Clientes.Estado  "
        Sql &= " FROM Clientes INNER JOIN"
        Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
        Sql &= " WHERE ClientesXTipos.Tipo_Id = 050 Order by Nome"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Nome"), Dr("Codigo")))
        Next

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
    End Sub

    Private Sub CargaEmpresa()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
        Sql &= " FROM   GruposXEmpresas INNER JOIN"
        Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
        Sql &= " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0

    End Sub

    Private Sub Limpar()
        DdlEmpresa.SelectedIndex = 0
        txtCliente.Text = ""
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresa()
    End Sub

    Private Function Validar()
        Mensagem = ""
        If DdlUnidade.Text = "" Then
            Mensagem = "Unidade de negócio é obrigatório..."
            Return Mensagem
        End If
        If DdlEmpresa.Text = "" Then
            Mensagem = "Empresa é obrigatório..."
            Return Mensagem
        End If
        Return Mensagem
    End Function

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        HttpContext.Current.Session("ssCampo") = "Adiantamentos"
        ucConsultaClientes.Limpar()
        Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeClientes(Me.Page, "objCalcJuros" & HID.Value, txtNome.ClientID)
    End Sub

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Try
            If Funcoes.VerificaPermissao("CalculoDosJurosSobreAdiantamentos", "GRAVAR") Then

                Dim AteDia As String = Format(CDate(CalData.SelectedDate), "yyyy/MM/dd")
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")

                Try
                    Dim sql As String
                    Dim ArraySql As New ArrayList
                    Dim ds_Adiantamentos As New DataSet
                    Dim ds_Cotacoes As New DataSet
                    'Dim dr As DataRow
                    Dim drCotacao As DataRow
                    Dim drVariacao As DataRow
                    Dim NumDias As Integer
                    Dim juro As Double
                    Dim cotacao As Double
                    'Dim ultimacotacao As Double
                    'Dim variacao As Double
                    'Dim ultimavaricao As Double
                    Dim ds_sequencia As New DataSet
                    Dim dr_sequencia As DataRow
                    Dim sequencia As Integer
                    'Dim Registro As String
                    Dim Lote As Int16 = 24
                    Dim dt As New DateTime

                    dt = DateTime.Now()
                    dt = dt.AddSeconds(1)

                    'Busca Adiantamentos menores que a data de atualização informada
                    sql = "SELECT Adiantamentos.Empresa_ID,Adiantamentos.EndEmpresa_ID,Empresa.Reduzido as ReduzidoEmpresa, Adiantamentos.Cliente_ID, Adiantamentos.EndCliente_ID,Adiantamentos.Adiantamento_Id, Adiantamentos.Movimento, Adiantamentos.Vencimento, Adiantamentos.UltimaAtualizacao, "
                    sql &= " Adiantamentos.ValorMoeda AS CapitalAdiantamento,Adiantamentos.ValorOficial AS CapitalAdiantamentoOficial, SUM(ISNULL(AdiantamentosXBaixas.ValorMoeda, 0)) AS ValorBaixa, "
                    sql &= " SUM(ISNULL(AdiantamentosXBaixas.VariacaoMoeda, 0)) AS ValorJuros, Adiantamentos.Taxa, convert(Decimal,0) as Juro , Adiantamentos.ValorMoeda + SUM(ISNULL(AdiantamentosXBaixas.VariacaoMoeda, 0)) - SUM(ISNULL(AdiantamentosXBaixas.ValorMoeda, 0)) AS Saldo, Clientes.Nome, SUM(ISNULL(AdiantamentosXBaixas.VariacaoOficial, 0)) AS ValorVariacao, SUM(ISNULL(AdiantamentosXBaixas.ValorOficial, 0)) AS ValorBaixaOficial"
                    sql &= " FROM Adiantamentos INNER JOIN"
                    sql &= " Clientes AS Empresa ON  Adiantamentos.Empresa_ID = Empresa.Cliente_Id AND "
                    sql &= " Adiantamentos.EndEmpresa_ID = Empresa.Endereco_Id INNER JOIN Clientes ON Adiantamentos.Cliente_ID = Clientes.Cliente_Id AND "
                    sql &= " Adiantamentos.EndCliente_ID = Clientes.Endereco_Id LEFT OUTER JOIN"
                    sql &= " AdiantamentosXBaixas ON "
                    sql &= " Adiantamentos.Empresa_ID = AdiantamentosXBaixas.Empresa_Id AND Adiantamentos.EndEmpresa_ID = AdiantamentosXBaixas.EndEmpresa_Id AND "
                    sql &= " Adiantamentos.Cliente_ID = AdiantamentosXBaixas.Cliente_Id AND Adiantamentos.EndCliente_ID = AdiantamentosXBaixas.EndCliente_Id AND "
                    sql &= " Adiantamentos.Adiantamento_Id = AdiantamentosXBaixas.Adiantamento_Id"
                    sql &= " WHERE Adiantamentos.Movimento < '" & Format(CDate(AteDia), "yyyy/MM/dd") & "'"
                    sql &= "And Adiantamentos.Empresa_ID='" & campo(0) & "' And Adiantamentos.EndEmpresa_Id=" & campo(1) & ""
                    If txtCliente.Text <> Nothing Then
                        sql &= "And Adiantamentos.Cliente_ID='" & HttpContext.Current.Session("Cliente") & "' And Adiantamentos.EndCliente_ID=" & HttpContext.Current.Session("ClienteEnd") & ""
                    End If
                    sql &= " GROUP BY Adiantamentos.Adiantamento_Id, Adiantamentos.Movimento, Adiantamentos.Vencimento, Adiantamentos.UltimaAtualizacao, "
                    sql &= " Adiantamentos.ValorMoeda,Adiantamentos.ValorOficial, Adiantamentos.Taxa, Adiantamentos.Empresa_ID, Adiantamentos.EndEmpresa_ID,Empresa.Reduzido, Adiantamentos.Cliente_ID, "
                    sql &= " Adiantamentos.EndCliente_ID,Clientes.Nome"
                    sql &= " having(Adiantamentos.ValorMoeda + SUM(ISNULL(AdiantamentosXBaixas.VariacaoMoeda, 0)) - SUM(ISNULL(AdiantamentosXBaixas.ValorMoeda, 0)) > 0)"

                    ds_Adiantamentos = Banco.ConsultaDataSet(sql, "Adiantamentos")

                    'Busca Valor da Contação da Moeda
                    sql = "Select Data_Id,Realizado,Indice From Cotacoes "
                    sql &= " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & Format(CDate(AteDia), "yyyy/MM/dd") & "')"
                    ds_Cotacoes = Banco.ConsultaDataSet(sql, "Cotacoes")
                    For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                        cotacao = drCotacao("Indice")
                    Next

                    sql = "Select isnull(max(Sequencia_Id),0) as Sequencia_Id from AdiantamentosXBaixas where Empresa_ID='" & campo(0) & "' and EndEmpresa_ID=" & campo(1) & " "
                    ds_sequencia = Banco.ConsultaDataSet(sql, "AdiantamentosXBaixas")
                    For Each dr_sequencia In ds_sequencia.Tables(0).Rows
                        sequencia = dr_sequencia("Sequencia_Id")
                    Next

                    Dim drSeqRazao As DataRow
                    Dim dsSeqRazao As New DataSet
                    Dim seqLoteRazao As Integer
                    '''''''''''''''''''''''''
                    'GERA SEQUENCIA DE LOTES RAZAO 
                    '''''''''''''''''''''''''
                    Try

                        sql = "Select isnull(max(Sequencia_Id),0) as Sequencia_Id from Razao"
                        sql &= " WHERE (Empresa_Id = '" & campo(0) & "' And EndEmpresa_Id=" & campo(1) & " and Movimento_Id ='" & Format(CDate(AteDia), "yyyy/MM/dd") & "' and Lote_Id=" & Lote & ")"
                        dsSeqRazao = Banco.ConsultaDataSet(sql, "AdiantamentosXBaixas")
                        For Each drSeqRazao In dsSeqRazao.Tables(0).Rows
                            seqLoteRazao = drSeqRazao("Sequencia_Id")
                        Next
                    Catch ex As SqlException
                        MsgBox(Me.Page, ex.Message)
                        Exit Sub
                    End Try


                    '''''''''''''''''''''
                    ''Calcula a variação 
                    '''''''''''''''''''''
                    For Each drVariacao In ds_Adiantamentos.Tables(0).Rows


                        '    If IsDBNull(drVariacao("UltimaAtualizacao")) = False Then
                        '        'Busca Valor da Contação da Moeda
                        '        sql = "Select Data_Id,Realizado,Cotacao From Cotacoes "
                        '        sql &= " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & Format(CDate(drVariacao("UltimaAtualizacao")), "yyyy/MM/dd") & "')"
                        '        ds_Cotacoes = banco.ConsultaDataSet(sql, "Cotacoes")
                        '        For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                        '            ultimacotacao = drCotacao("Cotacao")
                        '        Next
                        '    Else
                        '        'Busca Valor da Contação da Moeda
                        '        sql = "Select Data_Id,Realizado,Cotacao From Cotacoes "
                        '        sql &= " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & Format(CDate(drVariacao("Movimento")), "yyyy/MM/dd") & "')"
                        '        ds_Cotacoes = banco.ConsultaDataSet(sql, "Cotacoes")
                        '        For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                        '            ultimacotacao = drCotacao("Cotacao")
                        '        Next
                        '    End If
                        '    If IsDBNull(drVariacao("UltimaAtualizacao")) = True Then
                        '        NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("Movimento"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                        '    Else
                        '        NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("UltimaAtualizacao"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                        '    End If
                        '    If NumDias > 0 Then
                        '        'ultimavaricao = ((drVariacao("CapitalAdiantamento") + drVariacao("ValorJuros")) * ultimacotacao)
                        '        ultimavaricao = ((drVariacao("CapitalAdiantamentoOficial") + drVariacao("ValorVariacao")) - drVariacao("ValorBaixaOficial"))
                        '        variacao = (((drVariacao("CapitalAdiantamento") + drVariacao("ValorJuros") - drVariacao("ValorBaixa")) * cotacao) - ultimavaricao)
                        '    End If
                        '    If variacao <> 0 Then
                        '        Registro = ""
                        '        Registro = Format(dt, "yy")
                        '        Registro &= Format(dt.Month(), "00")
                        '        Registro &= Format(dt.Day(), "00")
                        '        Registro &= Format(dt.Hour, "00")
                        '        Registro &= Format(dt.Minute(), "00")
                        '        Registro &= Format(dt.Second(), "00")

                        '        sql = "INSERT INTO AdiantamentosXBaixas"
                        '        sql &= "(Pais_Id, Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id,Sequencia_Id,RegistroPedido,Titulo,VariacaoMoeda,VariacaoOficial, DataBaixa)"
                        '        sql &= " VALUES('" & Utilitarios.Global.GPais & "','" & drVariacao("Empresa_ID") & "' ," & drVariacao("EndEmpresa_ID") & ",'" & drVariacao("Cliente_ID") & "'," & drVariacao("EndCliente_ID") & ",'" & drVariacao("Adiantamento_Id") & "','" & drVariacao("ReduzidoEmpresa") + Registro & "'," & "0" & " ," & 0 & "," & Replace(0, ",", ".") & "," & Replace(variacao, ",", ".") & " ,'" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "')"
                        '        ArraySql.Add(sql)
                        '        'Variação Positiva
                        '        If variacao > 0 Then
                        '            'Lançamento de Debito
                        '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "11301", drVariacao("Cliente_Id"), drVariacao("EndCliente_ID"), "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), variacao, 0, 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id"), "4", "R"))
                        '            'Lançamento de Credito
                        '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "422020003", "", 0, "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), 0, variacao, 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id") & " - " & drVariacao("Nome"), "4", "R"))
                        '            'Variação Negativa
                        '        ElseIf variacao < 0 Then
                        '            'Lançamento de Debito
                        '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "422010003", "", 0, "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), variacao * (-1), 0, 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id") & " - " & drVariacao("Nome"), "4", "R"))
                        '            'Lançamento de Credito
                        '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "11301", drVariacao("Cliente_Id"), drVariacao("EndCliente_ID"), "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), 0, variacao * (-1), 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id"), "4", "R"))
                        '        End If
                        '        dt = dt.AddSeconds(1)
                        '    End If

                        '''''''''''''''
                        'Calcula Juros
                        '''''''''''''''
                        If CDate(drVariacao("Vencimento")) < CDate(AteDia) Then
                            Dim TotCapital As Decimal
                            TotCapital = drVariacao("CapitalAdiantamentoOficial") + drVariacao("ValorVariacao") - drVariacao("ValorBaixaOficial")

                            If IsDBNull(drVariacao("UltimaAtualizacao")) = True Then
                                NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("Vencimento"), "yyyy/MM/dd")), CDate(AteDia))
                            Else
                                If CDate(Format(drVariacao("Vencimento"), "yyyy/MM/dd")) > CDate(Format(drVariacao("UltimaAtualizacao"), "yyyy/MM/dd")) Then
                                    NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("Vencimento"), "yyyy/MM/dd")), CDate(AteDia))
                                Else
                                    NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("UltimaAtualizacao"), "yyyy/MM/dd")), CDate(AteDia))
                                End If
                            End If
                            If NumDias > 0 Then
                                juro = ((((TotCapital * CDec(drVariacao("Taxa")) / 100) / 30) * NumDias))
                            End If
                        End If

                        If NumDias > 0 Then

                            sequencia = sequencia + 1

                            seqLoteRazao = seqLoteRazao + 1

                            If juro > 0 Then

                                sql = "INSERT INTO AdiantamentosXBaixas"
                                sql &= "(Empresa_Id"
                                sql &= ", EndEmpresa_Id"
                                sql &= ", Cliente_Id"
                                sql &= ", EndCliente_Id "
                                sql &= ", Adiantamento_Id"
                                sql &= ", Sequencia_Id"
                                sql &= ", RegistroPedido"
                                sql &= ", Titulo"
                                sql &= ", ValorOficial"
                                sql &= ", ValorMoeda"
                                sql &= ", VariacaoOficial"
                                sql &= ", VariacaoMoeda"
                                sql &= ", DataBaixa)"

                                sql &= " VALUES("
                                sql &= "'" & drVariacao("Empresa_ID") & "'"                             'Empresa
                                sql &= ", " & drVariacao("EndEmpresa_ID")                               'End Empresa
                                sql &= ", '" & drVariacao("Cliente_Id") & "'"                           'Cliente
                                sql &= ", " & drVariacao("EndCliente_ID")                               'EndCliente
                                sql &= ", " & drVariacao("Adiantamento_Id")                             'Adiantamento
                                sql &= ", " & sequencia                                                 'Sequencia
                                sql &= ", 0"                                                            'Pedido
                                sql &= ", 0"                                                            'Titulo
                                sql &= ", 0"                                                            'Valor Oficial
                                sql &= ", 0"                                                            'Valor Moeda
                                sql &= ", " & Replace(CDec(FormatNumber(juro, 2)), ",", ".")            'Variação Oficial
                                sql &= ", " & Replace(CDec(FormatNumber(juro / cotacao, 2)), ",", ".")  'Variação Moeda
                                sql &= ", '" & Format(CDate(AteDia), "yyyy/MM/dd") & "')"               'Data da Baixa

                                ArraySql.Add(sql)


                                'Lançamento de Debito

                                sql = "INSERT INTO Razao ("
                                sql &= "Empresa_Id"
                                sql &= ", EndEmpresa_Id"
                                sql &= ", Conta_Id"
                                sql &= ", Cliente_Id"
                                sql &= ", EndCliente_Id"
                                sql &= ", Movimento_Id"
                                sql &= ", Lote_Id"
                                sql &= ", Sequencia_Id"
                                sql &= ", Indexador"
                                sql &= ", DataMoeda"
                                sql &= ", DebitoOficial"
                                sql &= ", CreditoOficial"
                                sql &= ", DebitoMoeda"
                                sql &= ", CreditoMoeda"
                                sql &= ", Historico"
                                sql &= ", PrevistoRealizado)"
                                sql &= " Values ("

                                sql &= "'" & drVariacao("Empresa_ID") & "'"                             'Empresa
                                sql &= ", " & drVariacao("EndEmpresa_ID")                               'End Empresa
                                sql &= ", '1010905'"                                                    'Conta
                                sql &= ", '" & drVariacao("Cliente_Id") & "'"                           'Cliente
                                sql &= ", " & drVariacao("EndCliente_ID")                               'EndCliente
                                sql &= ", '" & Format(CDate(AteDia), "yyyy/MM/dd") & "'"                'Movto
                                sql &= ", " & CInt(Lote)                                                'Lote
                                sql &= "," & CInt(seqLoteRazao)                                         'Sequencia de Lote
                                sql &= ", 3"                                                            'Indexador
                                sql &= ", '" & Format(CDate(AteDia), "yyyy/MM/dd") & "'"                'Data da Moeda
                                sql &= ", " & Replace(CDec(FormatNumber(juro, 2)), ",", ".")            'Débito Oficial
                                sql &= ", 0"                                                            'Crédito Oficial
                                sql &= ", " & Replace(CDec(FormatNumber(juro / cotacao, 2)), ",", ".")  'Debito Moeda
                                sql &= ", 0"                                                            'Crédito Moeda
                                sql &= ", '" & UCase(RTrim("Juros sobre adiantamento N. " & drVariacao("Adiantamento_Id"))) & "'"
                                sql &= ", 'R')"

                                ArraySql.Add(sql)

                                seqLoteRazao = seqLoteRazao + 1
                                'Lançamento de Credito

                                sql = "INSERT INTO Razao ("
                                sql &= "Empresa_Id"
                                sql &= ", EndEmpresa_Id"
                                sql &= ", Conta_Id"
                                sql &= ", Cliente_Id"
                                sql &= ", EndCliente_Id"
                                sql &= ", Movimento_Id"
                                sql &= ", Lote_Id"
                                sql &= ", Sequencia_Id"
                                sql &= ", Indexador"
                                sql &= ", DataMoeda"
                                sql &= ", DebitoOficial"
                                sql &= ", CreditoOficial"
                                sql &= ", DebitoMoeda"
                                sql &= ", CreditoMoeda"
                                sql &= ", Historico"
                                sql &= ", PrevistoRealizado)"
                                sql &= " Values ("

                                sql &= "'" & drVariacao("Empresa_ID") & "'"                             'Empresa
                                sql &= ", " & drVariacao("EndEmpresa_ID")                               'End Empresa
                                sql &= ", '303030301'"                                                  'Conta
                                sql &= ", ''"                                                           'Cliente
                                sql &= ", 0"                                                            'EndCliente
                                sql &= ", '" & Format(CDate(AteDia), "yyyy/MM/dd") & "'"                'Movto
                                sql &= ", " & CInt(Lote)                                                'Lote
                                sql &= "," & CInt(seqLoteRazao)                                         'Sequencia de Lote
                                sql &= ", 3"                                                            'Indexador
                                sql &= ", '" & Format(CDate(AteDia), "yyyy/MM/dd") & "'"                'Data da Moeda

                                sql &= ", 0"                                                            'Débito Oficial
                                sql &= ", " & Replace(CDec(FormatNumber(juro, 2)), ",", ".")            'Crédito Oficial

                                sql &= ", 0"                                                            'Débito Moeda
                                sql &= ", " & Replace(CDec(FormatNumber(juro / cotacao, 2)), ",", ".")  'Crédito Moeda

                                sql &= ", '" & UCase(RTrim("Juros sobre adiantamento N. " & drVariacao("Adiantamento_Id"))) & "'"
                                sql &= ", 'R')"

                                ArraySql.Add(sql)

                            End If

                            dt = dt.AddSeconds(1)
                        End If


                        If NumDias > 0 Then
                            'Atualiza a data da ultima atualização de jurus realizadas 
                            sql = "Update Adiantamentos"
                            sql &= " SET UltimaAtualizacao ='" & Format(CDate(AteDia), "yyyy/MM/dd") & "'"
                            sql &= " WHERE (Empresa_ID = '" & drVariacao("Empresa_ID") & "') AND (EndEmpresa_ID = " & drVariacao("EndEmpresa_ID") & ") AND (Cliente_ID = '" & drVariacao("Cliente_ID") & "') AND (EndCliente_ID = " & drVariacao("EndCliente_ID") & ") AND (Adiantamento_Id = '" & drVariacao("Adiantamento_Id") & "')"
                            ArraySql.Add(sql)
                        End If
                    Next


                    If Banco.GravaBanco(ArraySql) = True Then
                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "alert('Processo concluido !!!');", True)
                    Else
                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "alert('Erro no Processo !!!');", True)
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message)
                    ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "alert('Erro no Processo !!!');", True)
                End Try
            Else
                MsgBox(Me.Page, "Usuário sem permissão para Gravar Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Dim NomeArquivo As String = "Manual/PosicaoDeEstoques.mht"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
    End Sub
End Class