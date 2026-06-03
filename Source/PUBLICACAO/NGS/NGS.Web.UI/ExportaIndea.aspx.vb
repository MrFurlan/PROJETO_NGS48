Imports System.Data
Imports System.IO
Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ExportaIndea
    Inherits BasePage

    Private Mensagem As String
    Private sql As String
    Private NomeArquivo As String
    Private NomeArquivoVerificacao As String
    Private NomeArquivoVerificacao2 As String
    Private NomeArquivo2 As String
    Private linha As String
    Private linha2 As String
    Private strm As StreamWriter
    Private strm2 As StreamWriter
    Private k As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ExportaIndea", "ACESSAR") Then
                If Not Directory.Exists(Server.MapPath("Indea")) Then
                    Directory.CreateDirectory(Server.MapPath("Indea"))
                End If
                If txtEmpresa.Text = "" Then
                    ListarEmpresas()
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaINDEA" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaINDEA" & HID.Value), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            Session.Remove("objEmpresaINDEA" & HID.Value)
        End If
    End Sub

    Private Sub ListarEmpresas()
        Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    Private Sub Limpar()
        lnkExportar.Enabled = True
        lnkLimpar.Enabled = True
        btnEmpresa.Enabled = True
        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        txtDataInicial.Text = String.Format("01/{0:00}/{1}", DateTime.Now.Month, DateTime.Now.Year) 'Funcoes.PrimeiroDiaDoMes
        txtDataFinal.Text = Funcoes.UltimoDiaDoMes

        Session("ssFornecedores") = ""
        Session("ssClientes") = ""
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaEmpresas.SetarHID(HID.Value)
    End Sub

    Private Function ValidarCampos() As Boolean
        Mensagem = ""
        If txtEmpresa.Text.Length = 0 Then
            Mensagem = "Empresa não foi selecionada."
            Return False
        ElseIf CDate(txtDataInicial.Text).Month <> CDate(txtDataFinal.Text).Month Then
            Mensagem = "Mês inicial e final não podem ser diferentes."
            Return False
        ElseIf CDate(txtDataFinal.Text).Day <> DateTime.DaysInMonth(CDate(txtDataFinal.Text).Year, CDate(txtDataFinal.Text).Month) Then
            Mensagem = "Data final deve ser o último dia do Mês de " & MonthName(CDate(txtDataFinal.Text).Month)
            Return False
        ElseIf CDate(txtDataInicial.Text).Day <> 1 Then
            Mensagem = "Data inicial deve ser o primeiro dia do Mês de " & MonthName(CDate(txtDataInicial.Text).Month)
            Return False
        Else
            Return True
        End If
    End Function

    Function ArqEmpresa() As Boolean
        Mensagem = ""

        Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
        NomeArquivo2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/EMP_REF.TXT"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Try
            strm = New StreamWriter(NomeArquivo, True)
            linha = Funcoes.FormatarCpfCnpj(Empresa(0)) & ";" & CDate(txtDataInicial.Text).Month.ToString("00") & CDate(txtDataInicial.Text).Year.ToString
            strm.WriteLine(linha)
            strm.Close()
            Return True
        Catch ex As Exception
            strm.Close()
            Mensagem = Funcoes.EliminarCaracteresEspeciais("EMP_REF - " & ex.Message.ToString)
            Return False
        End Try
    End Function

    Function ArqCompras() As Boolean
        Dim Empresa As String() = txtCodigoEmpresa.Value.Split("-")
        Mensagem = ""

        sql = "SELECT NotasFiscais.Empresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento," & vbCrLf & _
              "       NotasFiscaisxItens.QtdeDeEmbalagem as Quantidade, Produtos.ProdutoIndea, " & vbCrLf & _
              "       Embalagens.EmbalagemIndea, ProdutoXEmbalagem.CapacidadeEmbalagem_Id," & vbCrLf & _
              "       UnidadeDeMedida.UnidadeIndea, ProdutoXEmbalagem.TipoDeEmbalagem_Id, " & vbCrLf & _
              "       Produtos.Produto_Id, Produtos.Nome" & vbCrLf & _
              "  FROM NotasFiscais" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
              " INNER JOIN Produtos" & vbCrLf & _
              "    ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
              " INNER JOIN UnidadeDeMedida" & vbCrLf & _
              "    ON Produtos.Unidade = UnidadeDeMedida.Unidade_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes" & vbCrLf & _
              "    ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
              " INNER JOIN ProdutoXEmbalagem" & vbCrLf & _
              "    ON NotasFiscaisXItens.Produto_Id          = ProdutoXEmbalagem.Produto_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.Embalagem           = ProdutoXEmbalagem.Embalagem_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.TipoDeEmbalagem     = ProdutoXEmbalagem.TipoDeEmbalagem_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.CapacidadeEmbalagem = ProdutoXEmbalagem.CapacidadeEmbalagem_Id" & vbCrLf & _
              " INNER JOIN Embalagens" & vbCrLf & _
              "    ON ProdutoXEmbalagem.Embalagem_Id = Embalagens.Embalagem_Id" & vbCrLf & _
              " WHERE Produtos.Fitossanitario = 'S'" & vbCrLf & _
              "   AND NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
              "   AND NotasFiscais.EntradaSaida_Id = 'E'" & vbCrLf & _
              "   AND SubOperacoes.Classe          = '" & eClassesOperacoes.COMPRAS.ToString & "'" & vbCrLf & _
              "   AND NotasFiscais.Situacao        = 1" & vbCrLf & _
              "   AND NotasFiscais.Empresa_id      ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND NotasFiscais.EndEmpresa_id   = " & Empresa(1) & vbCrLf & _
              "ORDER BY NotasFiscais.Cliente_Id" & vbCrLf

        Dim dsCompras As New DataSet
        dsCompras = Banco.ConsultaDataSet(sql, "Compras")

        NomeArquivo2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/MOV_COMP.TXT"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        strm = New StreamWriter(NomeArquivo, True)

        If dsCompras.Tables(0).Rows.Count > 0 Then
            Try
                linha = ""
                k = 0

                For Each drCompras As DataRow In dsCompras.Tables(0).Rows
                    If drCompras("ProdutoIndea").ToString.Length = 0 OrElse drCompras("ProdutoIndea") = "0" Then
                        Mensagem = "Produto " & drCompras("Produto_Id") & "-" & drCompras("Nome") & " sem o código do Produto do Indea. Informe o código e rode o Processo novamente."
                        Return False
                    End If
                    If k > 0 Then
                        linha &= vbCrLf
                    End If
                    linha &= Funcoes.FormatarCpfCnpj(drCompras("Empresa_Id")) & ";" & vbCrLf & _
                             Funcoes.FormatarCpfCnpj(drCompras("Cliente_Id")) & ";" & vbCrLf & _
                             drCompras("Nota_Id") & ";" & vbCrLf & _
                             drCompras("ProdutoIndea") & "-" & drCompras("EmbalagemIndea") & "-" & CDec(drCompras("CapacidadeEmbalagem_Id")).ToString("N3") & "-" & drCompras("UnidadeIndea") & "-" & drCompras("TipoDeEmbalagem_Id") & ";" & vbCrLf & _
                             drCompras("Movimento").tosqldate() & ";" & vbCrLf & _
                             drCompras("Quantidade") & ";" & vbCrLf & _
                             CDate(txtDataInicial.Text).Month.ToString("00") & CDate(txtDataInicial.Text).Year.ToString & vbCrLf
                    k += 1
                    If Not Session("ssFornecedores").Contains(drCompras("Cliente_Id")) Then
                        Session("ssFornecedores") += drCompras("Cliente_Id") & ";" & drCompras("EndCliente_Id") & "]"
                    End If
                Next

                strm.WriteLine(linha)
                strm.Close()
                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais("MOV_COMP - " & ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Function ArqFornecedores() As Boolean
        Mensagem = ""

        NomeArquivo2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/MOV_FORN.TXT"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        strm = New StreamWriter(NomeArquivo, True)

        If Session("ssFornecedores").ToString.Length > 0 Then
            Try
                linha = ""
                k = 0

                Dim ListaDeFornecedores() As String = Session("ssFornecedores").ToString.Split("]")
                Dim Fornecedor() As String
                For j As Integer = 0 To ListaDeFornecedores.GetUpperBound(0)
                    Fornecedor = ListaDeFornecedores(j).ToString.Split(";")
                    If Fornecedor(0).Length > 0 Then
                        Dim objCliente As New [Lib].Negocio.Cliente(Fornecedor(0), Fornecedor(1))

                        If k > 0 Then
                            linha &= vbCrLf
                        End If

                        linha &= Funcoes.FormatarCpfCnpj(objCliente.Codigo) & ";" _
                                 & Left(objCliente.Nome.ToUpper, 40) & ";" _
                                 & Left(objCliente.Fantasia.ToUpper, 30) & ";" _
                                 & Left(objCliente.InscricaoEstadual, 16) & ";" _
                                 & Left(objCliente.Endereco.ToUpper, 40) & ";" _
                                 & Left(objCliente.Bairro.ToUpper, 25) & ";" _
                                 & objCliente.CodigoEstado & ";" _
                                 & objCliente.Municipio.EstadoIbge _
                                 & Funcoes.AlinharDireita(objCliente.CodigoMunicipio.ToString, 5, "0").Substring(0, 4) & ";" _
                                 & objCliente.CEP & ";" _
                                 & Left(objCliente.Telefone, 14) & ";" _
                                 & Left(objCliente.Fax, 14) & ";" _
                                 & Left(objCliente.Email, 40) & ";;"

                        k += 1
                    End If
                Next

                strm.Write(linha)
                strm.Close()
                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais("MOV_FORN - " & ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Function ArqVendas() As Boolean
        Dim Empresa As String() = txtCodigoEmpresa.Value.Split("-")
        Mensagem = ""
        sql = "SELECT NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Nota_Id, NF.Movimento," & vbCrLf & _
              "       Nfi.QtdeDeEmbalagem as Quantidade, P.ProdutoIndea," & vbCrLf & _
              "       EM.EmbalagemIndea, Nfi.CapacidadeEmbalagem," & vbCrLf & _
              "       UN.UnidadeIndea, Nfi.TipoDeEmbalagem," & vbCrLf & _
              "       P.Produto_Id, P.Nome," & vbCrLf & _
              "       isnull(R.ART,0) AS ART, isnull(R.ARTReceita,0) as NumeroReceita," & vbCrLf & _
              "       isnull(RT.OrgaoRegCategoria,'') as Crea," & vbCrLf & _
              "       isnull(RP.DosagemRecomendada,0) as Dosagem," & vbCrLf & _
              "       isnull(RP.AreaTratada,0) as AreaTratada," & vbCrLf & _
              "       '0005' as PostoColeta," & vbCrLf & _
              "       isnull(prg.CodigoIndea,'9999') as Praga," & vbCrLf & _
              "       NF.Situacao, " & vbCrLf & _
              "       isnull(C.CodigoIndeaMT,0) as Cultura," & vbCrLf & _
              "       isnull(FA.CodigoIndea,0) as CodigoAplicacao" & vbCrLf & _
              "  FROM NotasFiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens Nfi" & vbCrLf & _
              "    ON NF.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = Nfi.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = Nfi.Nota_Id" & vbCrLf & _
              "  LEFT JOIN ReceitaXProduto RP" & vbCrLf & _
              "    ON RP.Receita_id = Nfi.Receita" & vbCrLf & _
              "   And RP.Produto_Id = Nfi.Produto_Id" & vbCrLf & _
              "  LEFT JOIN Receita R" & vbCrLf & _
              "    on R.Receita_Id = RP.Receita_Id" & vbCrLf & _
              "  LEFT JOIN Clientes RT" & vbCrLf & _
              "    on R.RespTecnico    = RT.Cliente_Id" & vbCrLf & _
              "   and R.EndRespTecnico = RT.Endereco_Id" & vbCrLf & _
              " INNER JOIN Produtos P" & vbCrLf & _
              "    ON Nfi.Produto_Id  = P.Produto_Id" & vbCrLf & _
              " INNER JOIN UnidadeDeMedida UN" & vbCrLf & _
              "    ON P.Unidade       = UN.Unidade_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes SO" & vbCrLf & _
              "    ON Nfi.Operacao    = SO.Operacao_Id" & vbCrLf & _
              "   AND Nfi.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
              "  LEFT JOIN CulturaXPragaXFito CPF" & vbCrLf & _
              "    ON CPF.CulturaPragaFito_Id = RP.CulturaPragaFito" & vbCrLf & _
              "  LEFT JOIN Cultura C" & vbCrLf & _
              "    ON C.Cultura_id = CPF.Cultura" & vbCrLf & _
              "  LEFT JOIN Praga Prg" & vbCrLf & _
              "    ON Prg.Praga_id = CPF.Praga" & vbCrLf & _
              " INNER JOIN Embalagens EM" & vbCrLf & _
              "    ON EM.Embalagem_Id = NFi.Embalagem" & vbCrLf & _
              "  LEFT JOIN FormaDeAplicacao FA" & vbCrLf & _
              "    ON FA.FormaDeAplicacao_Id = RP.FormaDeAplicacao" & vbCrLf & _
              " WHERE P.Fitossanitario   = 'S'" & vbCrLf & _
              "   AND NF.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = 'S'" & vbCrLf & _
              "   AND ((SO.Classe <> '" & eClassesOperacoes.COMPRAS.ToString & "' AND SO.Classe <> '" & eClassesOperacoes.CONTAEORDEM.ToString & "') OR SO.Devolucao = 'S')" & vbCrLf & _
              "   AND NF.Situacao        = 1" & vbCrLf & _
              "   AND NF.Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND NF.EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "   AND isnull(R.ART,0) > 0 " & vbCrLf & _
              "ORDER BY NF.Nota_Id" & vbCrLf

        Dim dsVendas As New DataSet
        dsVendas = Banco.ConsultaDataSet(sql, "Vendas")

        NomeArquivoVerificacao = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/MOV_VEND.TXT"
        NomeArquivo = Server.MapPath(NomeArquivoVerificacao)
        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        strm = New StreamWriter(NomeArquivo, True)

        If dsVendas.Tables(0).Rows.Count > 0 Then
            Try
                linha = ""
                k = 0
                For Each drVendas As DataRow In dsVendas.Tables(0).Rows

                    If drVendas("ProdutoIndea").ToString.Length = 0 OrElse drVendas("ProdutoIndea") = "0" Then
                        Mensagem = "Produto " & drVendas("Produto_Id") & "-" & drVendas("Nome") & " sem o código do Produto do Indea. Informe o código e rode o Processo novamente."
                        Return False
                    End If

                    If k > 0 Then
                        linha &= vbCrLf
                    End If

                    linha &= Funcoes.FormatarCpfCnpj(drVendas("Empresa_Id")) & ";" _
                            & drVendas("Nota_Id") & ";" _
                            & drVendas("ProdutoIndea") & "-" & drVendas("EmbalagemIndea") & "-" & Math.Round(drVendas("CapacidadeEmbalagem"), 3) & "-" & drVendas("UnidadeIndea") & "-" & drVendas("TipoDeEmbalagem") & ";" _
                            & drVendas("Movimento".ToStrDate() & ";" _
                            & drVendas("Quantidade") & ";" _
                            & drVendas("Cultura") & ";" _
                            & drVendas("CodigoAplicacao") & ";" _
                            & Funcoes.FormatarCpfCnpj(drVendas("Cliente_Id")) & ";" _
                            & drVendas("ART") & ";" _
                            & drVendas("NumeroReceita") & ";" _
                            & drVendas("Crea") & ";" _
                            & Math.Round(drVendas("Dosagem"), 3) & ";" _
                            & Math.Round(drVendas("AreaTratada"), 3) & ";" _
                            & CDate(txtDataInicial.Text).Month.ToString("00") & CDate(txtDataInicial.Text).Year.ToString & ";" _
                            & drVendas("PostoColeta") & ";" _
                            & Funcoes.AlinharDireita(drVendas("Praga"), 4, "0"))

                    k += 1

                    If Not Session("ssClientes").Contains(drVendas("Cliente_Id")) Then
                        Session("ssClientes") += drVendas("Cliente_Id") & ";" & drVendas("EndCliente_Id") & "]"
                    End If
                Next

                strm.WriteLine(linha)
                strm.Close()
                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais("MOV_VEND - " & ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Function ArqCancelamento() As Boolean
        Dim Empresa As String() = txtCodigoEmpresa.Value.Split("-")
        Mensagem = ""
        sql = "SELECT NF.Empresa_Id,  NF.Nota_Id, NF.Movimento" & vbCrLf & _
              "  FROM NotasFiscais NF" & vbCrLf & _
              " WHERE NF.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = 'S'" & vbCrLf & _
              "   AND NF.Situacao        in (2,9)" & vbCrLf & _
              "   AND NF.Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND NF.EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "ORDER BY NF.Nota_Id" & vbCrLf

        Dim dsCancelamento As New DataSet
        dsCancelamento = Banco.ConsultaDataSet(sql, "Cancelamento")

        NomeArquivoVerificacao2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/CAN_VDA.TXT"
        NomeArquivo2 = Server.MapPath(NomeArquivoVerificacao2)
        If Dir(NomeArquivo2).Length > 0 Then Kill(NomeArquivo2)

        strm2 = New StreamWriter(NomeArquivo2, True)

        If dsCancelamento.Tables(0).Rows.Count > 0 Then
            Try
                linha2 = ""
                k = 0

                For Each drCanc As DataRow In dsCancelamento.Tables(0).Rows
                    If k > 0 Then
                        linha2 &= vbCrLf
                    End If


                    linha2 &= Funcoes.FormatarCpfCnpj(drCanc("Empresa_Id")) & ";" _
                           & drCanc("Nota_Id") & ";" _
                           & drCanc("Movimento").tostrdate() & ";" _
                           & CDate(txtDataInicial.Text).Month.ToString("00") & CDate(txtDataInicial.Text).Year.ToString
                    k += 1
                Next

                strm2.WriteLine(linha2)
                strm2.Close()

                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais("MOV_CANC - " & ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Function ArqDevolucaoVendas() As Boolean
        Dim Empresa As String() = txtCodigoEmpresa.Value.Split("-")
        Mensagem = ""

        sql = "SELECT NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Nota_Id, NF.Movimento," & vbCrLf & _
              "       Nfi.QtdeDeEmbalagem as Quantidade," & vbCrLf & _
              "       P.ProdutoIndea,P.Produto_id, P.Nome," & vbCrLf & _
              "       EM.EmbalagemIndea, PE.CapacidadeEmbalagem_Id," & vbCrLf & _
              "       UN.UnidadeIndea, PE.TipoDeEmbalagem_Id" & vbCrLf & _
              "  FROM NotasFiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens Nfi" & vbCrLf & _
              "    ON NF.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = Nfi.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = Nfi.Nota_Id" & vbCrLf & _
              "  INNER JOIN Produtos P" & vbCrLf & _
              "    ON Nfi.Produto_Id  = P.Produto_Id" & vbCrLf & _
              " INNER JOIN UnidadeDeMedida UN" & vbCrLf & _
              "    ON P.Unidade       = UN.Unidade_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes SO" & vbCrLf & _
              "    ON Nfi.Operacao    = SO.Operacao_Id" & vbCrLf & _
              "   AND Nfi.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
              " INNER JOIN ProdutoXEmbalagem PE" & vbCrLf & _
              "    ON Nfi.Produto_Id          = PE.Produto_Id" & vbCrLf & _
              "   AND Nfi.Embalagem           = PE.Embalagem_Id" & vbCrLf & _
              "   AND Nfi.TipoDeEmbalagem     = PE.TipoDeEmbalagem_Id" & vbCrLf & _
              "   AND Nfi.CapacidadeEmbalagem = PE.CapacidadeEmbalagem_Id" & vbCrLf & _
              " INNER JOIN Embalagens EM" & vbCrLf & _
              "    ON PE.Embalagem_Id = EM.Embalagem_Id" & vbCrLf & _
              " WHERE P.Fitossanitario   = 'S'" & vbCrLf & _
              "   AND NF.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = 'E'" & vbCrLf & _
              "   AND SO.Devolucao       = 'S'" & vbCrLf & _
              "   AND SO.Classe          <> '" & eClassesOperacoes.COMPRAS.ToString & "'" & vbCrLf & _
              "   AND NF.Situacao        = 1" & vbCrLf & _
              "   AND NF.Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND NF.EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              "ORDER BY NF.Nota_Id" & vbCrLf

        Dim dsDev As New DataSet
        dsDev = Banco.ConsultaDataSet(sql, "Vendas")

        NomeArquivo2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/DEV_VDA.TXT"
        NomeArquivo = Server.MapPath(NomeArquivo2)
        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        strm = New StreamWriter(NomeArquivo, True)

        If dsDev.Tables(0).Rows.Count > 0 Then
            Try
                linha = ""
                k = 0

                For Each drDev As DataRow In dsDev.Tables(0).Rows
                    If drDev("ProdutoIndea").ToString.Length = 0 OrElse drDev("ProdutoIndea") = "0" Then
                        Mensagem = "Produto " & drDev("Produto_Id") & "-" & drDev("Nome") & " sem o código do Produto do Indea. Informe o código e rode o Processo novamente."
                        Return False
                    End If

                    If k > 0 Then
                        linha &= vbCrLf
                    End If

                    linha &= Funcoes.FormatarCpfCnpj(drDev("Empresa_Id")) & ";" _
                          & drDev("Nota_Id") & ";" _
                          & drDev("ProdutoIndea") & "-" & drDev("EmbalagemIndea") & "-" & Math.Round(drDev("CapacidadeEmbalagem_Id"), 3) & "-" & drDev("UnidadeIndea") & "-" & drDev("TipoDeEmbalagem_Id") & ";" _
                          & drDev("Quantidade") & ";" _
                          & Funcoes.FormatarCpfCnpj(drDev("Cliente_Id")) & ";" _
                          & drDev("Movimento").tostrdate() & ";" _
                          & CDate(txtDataInicial.Text).Month.ToString & CDate(txtDataInicial.Text).Year.ToString

                    k += 1

                    If Not Session("ssClientes").Contains(drDev("Cliente_Id")) Then
                        Session("ssClientes") += drDev("Cliente_Id") & ";" & drDev("EndCliente_Id") & "]"
                    End If
                Next

                strm.Write(linha)
                strm.Close()
                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais("DEV_VDA - " & ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Function ArqClientes() As Boolean
        Mensagem = ""

        NomeArquivo2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/MOV_CLIE.TXT"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        strm = New StreamWriter(NomeArquivo, True)

        If Session("ssClientes").ToString.Length > 0 Then
            Try
                linha = ""
                k = 0

                Dim ListaDeClientes() As String = Session("ssClientes").ToString.Split("]")
                Dim Cliente() As String
                For j As Integer = 0 To ListaDeClientes.GetUpperBound(0)
                    Cliente = ListaDeClientes(j).ToString.Split(";")
                    If Cliente(0).Length > 0 Then
                        Dim objCliente As New [Lib].Negocio.Cliente(Cliente(0), Cliente(1))

                        If k > 0 Then
                            linha &= vbCrLf
                        End If

                        linha &= IIf(objCliente.Codigo.Length = 14, "2", "0") & ";" _
                              & Funcoes.FormatarCpfCnpj(objCliente.Codigo) & ";" _
                              & Left(objCliente.Nome.ToUpper, 45) & ";" _
                              & Left(objCliente.Endereco.ToUpper, 40) & ";" _
                              & Left(objCliente.Bairro.ToUpper, 25) & ";" _
                              & objCliente.CodigoEstado & ";" _
                              & objCliente.Municipio.EstadoIbge _
                              & Funcoes.AlinharDireita(objCliente.CodigoMunicipio.ToString, 5, "0").Substring(0, 4) & ";" _
                              & objCliente.CEP & ";" _
                              & Left(objCliente.Telefone, 14) & ";" _
                              & Left(objCliente.Nome.ToUpper, 45) & ";" _
                              & Left(objCliente.Email, 40) & ";"

                        Dim revenda As Boolean = False
                        For Each t As [Lib].Negocio.ClientexTipo In objCliente.Tipos
                            If t.CodigoTipo = 17 Then
                                revenda = True
                            End If
                        Next

                        If revenda Then
                            linha &= "1"
                        Else
                            linha &= "0"
                        End If

                        'q += 1

                        k += 1
                    End If
                Next

                strm.Write(linha)
                strm.Close()
                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais("MOV_CLIE - " & ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Function ArqProdutos() As Boolean
        Dim Empresa As String() = txtCodigoEmpresa.Value.Split("-")
        Mensagem = ""

        sql = "SELECT distinct NotasFiscais.Empresa_Id, Produtos.ProdutoIndea, " & vbCrLf & _
              "       Embalagens.EmbalagemIndea, ProdutoXEmbalagem.CapacidadeEmbalagem_Id," & vbCrLf & _
              "       UnidadeDeMedida.UnidadeIndea, ProdutoXEmbalagem.TipoDeEmbalagem_Id," & vbCrLf & _
              "       Produtos.Produto_Id, Produtos.Nome" & vbCrLf & _
              "  FROM NotasFiscais" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
              " INNER JOIN Produtos" & vbCrLf & _
              "    ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
              " INNER JOIN UnidadeDeMedida" & vbCrLf & _
              "    ON Produtos.Unidade = UnidadeDeMedida.Unidade_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes" & vbCrLf & _
              "    ON NotasFiscaisXItens.Operacao = SubOperacoes.Operacao_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
              " INNER JOIN ProdutoXEmbalagem" & vbCrLf & _
              "    ON NotasFiscaisXItens.Produto_Id          = ProdutoXEmbalagem.Produto_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.Embalagem           = ProdutoXEmbalagem.Embalagem_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.TipoDeEmbalagem     = ProdutoXEmbalagem.TipoDeEmbalagem_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.CapacidadeEmbalagem = ProdutoXEmbalagem.CapacidadeEmbalagem_Id" & vbCrLf & _
              " INNER JOIN Embalagens" & vbCrLf & _
              "    ON ProdutoXEmbalagem.Embalagem_Id = Embalagens.Embalagem_Id" & vbCrLf & _
              " WHERE Produtos.Fitossanitario = 'S'" & vbCrLf & _
              "   AND NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
              "   AND SubOperacoes.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "'" & vbCrLf & _
              "   AND NotasFiscais.Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
              "   AND NotasFiscais.EndEmpresa_id = " & Empresa(1) & vbCrLf & _
              " ORDER BY NotasFiscais.Empresa_Id, Produtos.ProdutoIndea"


        Dim dsProd As New DataSet
        dsProd = Banco.ConsultaDataSet(sql, "Produtos")

        NomeArquivo2 = "Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString & "/Mov_Prod.TXT"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        strm = New StreamWriter(NomeArquivo, True)

        If dsProd.Tables(0).Rows.Count > 0 Then
            Try
                linha = ""
                k = 0

                For Each drProd As DataRow In dsProd.Tables(0).Rows
                    If drProd("ProdutoIndea").ToString.Length = 0 OrElse drProd("ProdutoIndea") = "0" Then
                        Mensagem = "Produto " & drProd("Produto_Id") & "-" & drProd("Nome") & " sem o código do Produto do Indea. Informe o código e rode o Processo novamente."
                        Return False
                    End If

                    If k > 0 Then
                        linha &= vbCrLf
                    End If

                    linha &= Funcoes.FormatarCpfCnpj(drProd("Empresa_Id")) & ";" _
                          & drProd("ProdutoIndea") & ";" _
                          & drProd("Nome") & ";" _
                          & drProd("EmbalagemIndea") & ";" _
                          & drProd("TipoDeEmbalagem_Id") & ";" _
                          & drProd("CapacidadeEmbalagem_Id") & ";" _
                          & drProd("UnidadeIndea") & ";" _
                          & drProd("ProdutoIndea") & "-" & drProd("EmbalagemIndea") & "-" & drProd("CapacidadeEmbalagem_Id") & "-" & drProd("UnidadeIndea") & "-" & drProd("TipoDeEmbalagem_Id") & ";" _
                          & "01/01/2008" & ";" _
                          & "0;0"

                    k += 1
                Next

                strm.WriteLine(linha)
                strm.Close()
                Return True
            Catch ex As Exception
                strm.Close()
                Mensagem = Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
                Return False
            End Try
        Else
            strm.Close()
            Return True
        End If
    End Function

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Try
            Session("ssCampo") = "Livre"
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaINDEA" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExportar_Click(sender As Object, e As EventArgs) Handles lnkExportar.Click
        Try
            If ValidarCampos() Then
                If Not Directory.Exists(Server.MapPath("Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString)) Then
                    Directory.CreateDirectory(Server.MapPath("Indea/" & CDate(txtDataInicial.Text).Year.ToString & "/" & MonthName(CDate(txtDataInicial.Text).Month).ToString))
                End If

                If ArqEmpresa() Then
                    If ArqCompras() Then
                        If ArqFornecedores() Then
                            If ArqVendas() Then
                                If ArqDevolucaoVendas() Then
                                    If ArqCancelamento() Then
                                        If ArqClientes() Then
                                            If ArqProdutos() Then
                                                MsgBox(Me.Page, "Processo concluído.")
                                            Else
                                                MsgBox(Me.Page, Mensagem)
                                            End If
                                        Else
                                            MsgBox(Me.Page, Mensagem)
                                        End If
                                    Else
                                        MsgBox(Me.Page, Mensagem)
                                    End If
                                Else
                                    MsgBox(Me.Page, Mensagem)
                                End If
                            Else
                                MsgBox(Me.Page, Mensagem)
                            End If
                        Else
                            MsgBox(Me.Page, Mensagem)
                        End If
                    Else
                        MsgBox(Me.Page, Mensagem)
                    End If
                Else
                    MsgBox(Me.Page, Mensagem)
                End If
            Else
                MsgBox(Me.Page, Mensagem)
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
            Funcoes.Ajuda(Me.Page, "ExportaIndea")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class