Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PosicaoDeTransferencias
    Inherits BasePage

    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Expedicao)
                If Funcoes.VerificaPermissao("PosicaoDeTransferencias", "ACESSAR") Then
                    CargaUnidadeDeNegocio()
                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        Dim Codigo As String
        Dim Descricao As String

        sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Nome, Clientes.Cidade, Clientes.Estado  " & vbCrLf & _
              " FROM Clientes INNER JOIN" & vbCrLf & _
              " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id" & vbCrLf & _
              " WHERE ClientesXTipos.Tipo_Id = 050" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "UnidadeDeNegocio").Tables(0).Rows
            Codigo = Dr("Codigo")
            Descricao = Dr("Nome")

            DdlUnidadeDeNegocioOrigem.Items.Add(New ListItem(Descricao, Codigo))
            DdlUnidadeDeNegocioDestino.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlUnidadeDeNegocioOrigem.Items.Insert(0, "")
        DdlUnidadeDeNegocioOrigem.SelectedIndex = 0

        DdlUnidadeDeNegocioDestino.Items.Insert(0, "")
        DdlUnidadeDeNegocioDestino.SelectedIndex = 0
    End Sub

    Private Sub CargaEmpresaOrigem()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaClienteOrigem.Items.Clear()

        sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidadeDeNegocioOrigem.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlEmpresaClienteOrigem.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlEmpresaClienteOrigem.Items.Insert(0, "")
        DdlEmpresaClienteOrigem.SelectedIndex = 0
    End Sub

    Private Sub CargaEmpresaDestino()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaClienteDestino.Items.Clear()

        sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidadeDeNegocioDestino.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlEmpresaClienteDestino.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlEmpresaClienteDestino.Items.Insert(0, "")
        DdlEmpresaClienteDestino.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        DdlUnidadeDeNegocioOrigem.SelectedIndex = 0
        DdlEmpresaClienteOrigem.SelectedIndex = 0
        DdlUnidadeDeNegocioDestino.SelectedIndex = 0
        DdlEmpresaClienteDestino.SelectedIndex = 0
        ucSelecaoProduto.Limpar()
        txtPeriodoInicialConsultaTitulos.Text = Format(Today, "01/01/" & "yyyy")
        txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocioOrigem.Enabled = False
            DdlEmpresaClienteOrigem.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidadeDeNegocioOrigem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresaOrigem()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidadeDeNegocioDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresaDestino()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("PosicaoDeTransferencias", "RELATORIO") Then
                Dim linha As String
                Dim sql As String
                Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
                Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
                Dim ds As New DataSet
                Dim dr As DataRow

                Dim strm As StreamWriter = Nothing
                If Dir(arquivo).Length > 0 Then Kill(arquivo)

                Dim UnidadeOrigem As String
                Dim EmpresaOrigem As String
                Dim EndEmpresaOrigem As String
                Dim UnidadeDestino As String
                Dim EmpresaDestino As String
                Dim EndEmpresaDestino As String
                Dim cliente As String
                Dim campo() As String

                If DdlUnidadeDeNegocioOrigem.Text <> "" Then
                    UnidadeOrigem = DdlUnidadeDeNegocioOrigem.SelectedValue
                Else
                    UnidadeOrigem = ""                   'UnidadeDeNegocio
                End If

                If DdlUnidadeDeNegocioDestino.Text <> "" Then
                    UnidadeDestino = DdlUnidadeDeNegocioDestino.SelectedValue
                Else
                    UnidadeDestino = ""                   'UnidadeDeNegocio
                End If

                If DdlEmpresaClienteOrigem.Text <> "" Then
                    cliente = DdlEmpresaClienteOrigem.SelectedValue
                    campo = cliente.Split("-")
                    EmpresaOrigem = campo(0)                      'EmpresaCliente
                    EndEmpresaOrigem = campo(1)                   'Endereco Empresa Cliente
                Else
                    EmpresaOrigem = ""                            'Empresa Cliente
                    EndEmpresaOrigem = 0                          'Endereco Empresa Cliente
                End If

                If DdlEmpresaClienteDestino.Text <> "" Then
                    cliente = DdlEmpresaClienteDestino.SelectedValue
                    campo = cliente.Split("-")
                    EmpresaDestino = campo(0)                      'EmpresaCliente
                    EndEmpresaDestino = campo(1)                   'Endereco Empresa Cliente
                Else
                    EmpresaDestino = ""                            'Empresa Cliente
                    EndEmpresaDestino = 0                          'Endereco Empresa Cliente
                End If

                Dim DataInicial As String = Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "yyyy/MM/dd")
                Dim DataFinal As String = Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "yyyy/MM/dd")


                sql = " SELECT     ClientesOrigem.Cliente_Id, ClientesOrigem.Endereco_Id, ClientesOrigem.Reduzido, ClientesOrigem.Nome, ClientesOrigem.Cidade, " & vbCrLf & _
                      "           ClientesOrigem.Estado, Produtos.Produto_Id, Produtos.Grupo, Produtos.Nome AS NomeDoProduto, ISNULL(ClientesDestino.Cliente_Id, '-') " & vbCrLf & _
                      "                 AS EmpresaDestino, ISNULL(ClientesDestino.Endereco_Id, 0) AS EndEmpresaDestino, ISNULL(ClientesDestino.Reduzido, '-') AS ReduzidoEmpDestino, " & vbCrLf & _
                      "        ISNULL(ClientesDestino.Nome, '-') AS NomeEmpresaDestino, ISNULL(ClientesDestino.Cidade, '-') AS CidadeEmpresaDestino, " & vbCrLf & _
                      "                   ISNULL(ClientesDestino.Estado, '-') AS EstadoEmpresaDestino, NotasFiscais.Movimento, Pesagem.Movimento AS DataDaChegada, " & vbCrLf & _
                      "          ISNULL(NotasFiscaisXItens.PesoFiscal, 0) AS PesoSaida, ISNULL(Pesagem.BrutoBalanca, 0) AS PesoChegada, ISNULL(Pesagem.Placa, '-') AS Placa, " & vbCrLf & _
                      "                    NotasFiscais.Nota_Id, ISNULL(Pesagem.Pesagem_Id, 0) AS Pesagem_Id, ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0) AS Fisica, " & vbCrLf & _
                      "           ISNULL(Pesagem.Liquido, 0) AS Liquido, COALESCE (Umidade.Percentual, 0) AS PercUmidade, COALESCE (Umidade.Desconto, 0) AS DescUmidade, " & vbCrLf & _
                      "                  COALESCE (Impureza.Percentual, 0) AS PercImpureza, COALESCE (Impureza.Desconto, 0) AS DescImpureza, COALESCE (Avariados.Percentual, 0) " & vbCrLf & _
                      "         AS PercAvariados, COALESCE (Avariados.Desconto, 0) AS DescAvariados, ISNULL(Deposito.Reduzido, '-') AS DepositoReduzido, " & vbCrLf & _
                      "                  ISNULL(Deposito.Cliente_Id, '-') AS DepositoCodigo, ISNULL(Deposito.Nome, '-') AS DepositoNome, ISNULL(Deposito.Cidade, '-') AS DepositoCidade, " & vbCrLf & _
                      "         ISNULL(Deposito.Estado, '-') AS DepositoEstado" & vbCrLf & _
                      "  FROM         Clientes AS ClientesDestino INNER JOIN" & vbCrLf & _
                      "             Pesagem ON ClientesDestino.Cliente_Id = Pesagem.Empresa_Id AND ClientesDestino.Endereco_Id = Pesagem.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                      "                 Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id RIGHT OUTER JOIN" & vbCrLf & _
                      "        NotasFiscaisXItens INNER JOIN" & vbCrLf & _
                      "                 NotasFiscais ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id AND " & vbCrLf & _
                      "               NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id AND " & vbCrLf & _
                      "      NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id AND " & vbCrLf & _
                      "                NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id INNER JOIN" & vbCrLf & _
                      "                Clientes AS ClientesOrigem ON NotasFiscaisXItens.Empresa_Id = ClientesOrigem.Cliente_Id AND " & vbCrLf & _
                      "               NotasFiscaisXItens.EndEmpresa_Id = ClientesOrigem.Endereco_Id INNER JOIN" & vbCrLf & _
                      "                 Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id ON Pesagem.Cliente = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                      "        Pesagem.EndCliente = NotasFiscaisXItens.EndCliente_Id AND Pesagem.NumeroDaNota = NotasFiscaisXItens.Nota_Id LEFT OUTER JOIN" & vbCrLf & _
                      "                 vw_pesagemxanalises AS Umidade ON Umidade.Pesagem_Id = Pesagem.Pesagem_Id AND Umidade.Empresa_Id = Pesagem.Empresa_Id AND " & vbCrLf & _
                      "                Umidade.EndEmpresa_Id = Pesagem.EndEmpresa_Id AND Umidade.Analise_Id = 1 LEFT OUTER JOIN" & vbCrLf & _
                      "                vw_pesagemxanalises AS Impureza ON Impureza.Pesagem_Id = Pesagem.Pesagem_Id AND Impureza.Empresa_Id = Pesagem.Empresa_Id AND " & vbCrLf & _
                      "               Impureza.EndEmpresa_Id = Pesagem.EndEmpresa_Id AND Impureza.Analise_Id = 2 LEFT OUTER JOIN" & vbCrLf & _
                      "      vw_pesagemxanalises AS Avariados ON Avariados.Pesagem_Id = Pesagem.Pesagem_Id AND Avariados.Empresa_Id = Pesagem.Empresa_Id AND " & vbCrLf & _
                      "  Avariados.EndEmpresa_Id = Pesagem.EndEmpresa_Id And Avariados.Analise_Id = 3" & vbCrLf

                sql &= " WHERE (NotasFiscais.EntradaSaida_Id = 'S')"
                If EmpresaOrigem <> "" Then
                    sql &= " and (NotasFiscais.Empresa_Id = '" & EmpresaOrigem & "')  " & vbCrLf & _
                           " and (NotasFiscais.EndEmpresa_Id = " & EndEmpresaOrigem & ")  " & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NotasFiscaisXItens.Produto_Id", "")
                    sql &= " AND " & RetornoProdutos(0)
                End If

                If EmpresaDestino <> "" Then
                    sql &= " and (Pesagem.Empresa_Id = '" & EmpresaDestino & "')" & vbCrLf & _
                           " and (Pesagem.EndEmpresa_Id = " & EndEmpresaDestino & ")" & vbCrLf
                End If
               
                If DataInicial <> "" And DataFinal <> "" Then
                    sql &= " and  NotasFiscais.Movimento between '" & DataInicial & "' and '" & DataFinal & "'"
                End If

                ds = Banco.ConsultaDataSet(sql, "Clientes")

                linha = "<HTML>" & vbCrLf
                '<HEAD>
                linha &= "<HEAD>" & vbCrLf
                linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
                linha &= "<TITLE>Posicao de transferencia</TITLE>" & vbCrLf
                linha &= "</HEAD>" & vbCrLf

                '<BODY>
                linha &= "<BODY>" & vbCrLf

                'Cabeçalho Padrao
                linha &= "<table width= '5000' cellpadding='0' cellspacing='0' Border=1>"

                linha &= "<TR>"
                linha &= "<TD>Produto</TD>"
                linha &= "<TD>NomeProduto</TD>"
                linha &= "<TD>DataSaida</TD>"
                linha &= "<TD>DataChegada</TD>"
                linha &= "<TD >Placa</TD>"

                linha &= "<TD >Nota</TD>"

                linha &= "<TD>CodigoOrigem</TD>"
                linha &= "<TD>EmpresaOrigem</TD>"
                linha &= "<TD>CidadeOrigem</TD>"
                linha &= "<TD>UFOrigem</TD>"

                linha &= "<TD>CodigoDestino</TD>"
                linha &= "<TD>EmpresaDestino</TD>"
                linha &= "<TD>CidadeDestino</TD>"
                linha &= "<TD>UFDestino</TD>"

                linha &= "<TD>DepositoReduzido</TD>"
                linha &= "<TD>DepositoCodigo</TD>"
                linha &= "<TD>DepositoNome</TD>"
                linha &= "<TD>DepositoCidade</TD>"
                linha &= "<TD>DepositoEstado</TD>"

                linha &= "<TD>Laudo</TD>"

                linha &= "<TD >LiquidoOrigem</TD>"
                linha &= "<TD >LiquidoDestinoBalanca</TD>"
                linha &= "<TD >FiscalNota </TD>"
                linha &= "<TD >LiquidoDestinoClassificacao</TD>"
                linha &= "<TD >VariacaoDePeso</TD>"
                linha &= "<TD >VariacaoDeClassificacao</TD>"
                linha &= "<TD>PercUmidade</TD>"
                linha &= "<TD>DescUmidade</TD>"

                linha &= "<TD>PercImpureza</TD>"
                linha &= "<TD>DescImpureza</TD>"
                linha &= "<TD>PercAvariados</TD>"
                linha &= "<TD>DescAvariados</TD>"
                linha &= "</TR>"

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        linha &= "<TR><TD>" & dr("Produto_Id") & "</TD>"
                        linha &= "<TD>" & Funcoes.EliminarCaracteresEspeciais(dr("NomeDoProduto")) & "</TD>"
                        linha &= "<TD>" & Format(dr("Movimento"), "dd/MM/yyyy") & "</TD>"
                        If IsDBNull(dr("DataDaChegada")) = False Then
                            linha &= "<TD>" & Format(dr("DataDaChegada"), "dd/MM/yyyy") & "</TD>"
                        Else
                            linha &= "<TD>" & "-" & "</TD>"
                        End If
                        linha &= "<TD>" & dr("Placa") & "</TD>"

                        linha &= "<TD>" & dr("Nota_Id") & "</TD>"

                        linha &= "<TD>" & dr("Reduzido") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Nome") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Cidade") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Estado") & "</TD>" & vbCrLf

                        linha &= "<TD>" & dr("ReduzidoEmpDestino") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("NomeEmpresaDestino") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("CidadeEmpresaDestino") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("EstadoEmpresaDestino") & "</TD>" & vbCrLf

                        linha &= "<TD>" & dr("DepositoReduzido") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DepositoCodigo") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DepositoNome") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DepositoCidade") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DepositoEstado") & "</TD>" & vbCrLf

                        linha &= "<TD>" & dr("Pesagem_Id") & "</TD>"

                        linha &= "<TD>" & dr("PesoSaida") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("PesoChegada") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Fisica") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Liquido") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("PesoChegada") - dr("PesoSaida") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Liquido") - dr("Fisica") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("PercUmidade") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DescUmidade") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("PercImpureza") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DescImpureza") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("PercAvariados") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("DescAvariados") & "</TD>" & vbCrLf & _
                                 "</TR>" & vbCrLf

                    Next
                End If
                Try
                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
                Finally
                    strm.Close()
                End Try
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
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
            Funcoes.Ajuda(Me.Page, "PosicaoDeTransferencias")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class