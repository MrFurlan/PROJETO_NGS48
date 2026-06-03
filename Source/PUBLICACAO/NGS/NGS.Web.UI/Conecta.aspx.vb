Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Conecta
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Conecta", "ACESSAR") Then
                If String.IsNullOrWhiteSpace(txtEmpresa.Text) Then
                    If Not Directory.Exists(Server.MapPath("Conecta")) Then
                        Directory.CreateDirectory(Server.MapPath("Conecta"))
                    End If
                    txtData.Text = DateTime.Now.ToShortDateString()
                    ListarEmpresas()
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaEmpresas.SetarHID(HID.Value)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaCON" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaCON" & HID.Value), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            Session.Remove("objEmpresaCON" & HID.Value)
        End If
    End Sub

    Private Sub ListarEmpresas()
        Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Session("ssCampo") = "Livre"
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCON" & HID.Value)
    End Sub

    Protected Sub lnkEnviar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkEnviar.Click
        GerarArquivo("ENVIO")
    End Sub

    Protected Sub lnkSubstituir_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkSubstituir.Click
        GerarArquivo("SUBSTITUIR")
    End Sub

    Private Sub GerarArquivo(ByVal Tipo As String)
        Dim PeriodoInicial As Date = Format(DateValue(txtData.Text), "yyyy-MM-dd")
        Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
        Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))

        Dim NomeArquivo As String
        Dim linha As String
        Dim strm As StreamWriter
        Dim dr As DataRow
        Dim row As DataRow
        Dim row2 As DataRow

        Dim RegistroGeral As Integer = 0     'Registro Geral
        Dim QtdeNotaFiscal As Integer = 0    'Numero de Notas Contidas no Arquivo
        Dim ItensNota As Integer = 0         'Registro linhas por nota
        Dim itensEstoque As Integer = 0      'Registro itens movimentados Estoque dia 
        Dim QtdeCliente As Integer = 0       'Numero de Clientes
        Dim NumLinhaAplicacao As Integer = 0 'Registro NumLinha


        Dim NomeArquivo2 As String = "Conecta/" & Format(DateValue(txtData.Text), "yyyy-MM-dd") & "-" & Tipo & ".txt"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)
        Dim Sql As String = ""
        Dim ds As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        Dim Array As New ArrayList

        '***********************************************************************************************
        ' Dados do Distribuidor
        '---------------------------------------------------------------------------------------------
        ' ID 00
        '***********************************************************************************************

        Sql = "Select '" & PeriodoInicial & "' as Data," & vbCrLf & _
              "       'UNIFIC' as Tipo," & vbCrLf & _
              "       '01' as Doc," & vbCrLf & _
              "       cliente_id," & vbCrLf & _
              "       Nome " & vbCrLf & _
              "  from clientes" & vbCrLf & _
              " where cliente_id  = '" & objEmpresa.Codigo & "'" & vbCrLf & _
              "   and Endereco_id = " & objEmpresa.CodigoEndereco & "" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "T00")

        If ds.Tables("T00").Rows.Count > 0 Then
            For Each dr In ds.Tables("T00").Rows
                linha = "00"                                         'ID
                linha &= Format(CDate(txtData.Text), "yyyyMMddHHmm") 'Data
                linha &= "UNIFIC"                                    'Tipo De Mensagem 

                Select Case Tipo
                    Case "ENVIO" : linha &= "01"
                    Case "SUBSTITUIR" : linha &= "02"
                End Select 'Tipo Documento

                linha &= Funcoes.AlinharEsquerda(dr("Cliente_Id"), 14, " ")  'Identificacao do Distribuidor 
                linha &= Funcoes.AlinharEsquerda(RTrim(dr("Nome")), 40, " ")  'Razão Social do Distribuidor 

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                RegistroGeral += 1
            Next
        End If

        '***********************************************************************************************
        '*** Dados da Nota Fiscal
        '***--------------------------------------------------------------------------------------------
        '*** ID 10
        '***********************************************************************************************

        Sql = "Select NF.Empresa_Id, NF.EndEmpresa_Id, NF.Cliente_id, NF.EndCliente_Id, EntradaSaida_Id," & vbCrLf & _
              "       NF.Nota_Id, " & vbCrLf & _
              "       NF.Serie_Id," & vbCrLf & _
              "       convert(nvarchar,nf.operacao) + convert(nvarchar,nf.suboperacao) as Tipo," & vbCrLf & _
              "       sop.Descricao," & vbCrLf & _
              "       Pe.CondicaoPagamento," & vbCrLf & _
              "       isnull(pa.Descricao,'SEM PAGAMENTO') as DescPagamento," & vbCrLf & _
              "       nf.datadanota as DataVencimento," & vbCrLf & _
              "       nf.datadanota," & vbCrLf & _
              "       isnull(Cli.Cliente_Id,'" & objEmpresa.Codigo & "') as Vendedor," & vbCrLf & _
              "       isnull(Cli.Nome,'" & objEmpresa.Nome.ToUpper & "') as NomeVendedor," & vbCrLf & _
              "       CNF.Nome as ClienteNF," & vbCrLf & _
              "       CNF.Cep" & vbCrLf & _
              "  from NotasFiscais NF" & vbCrLf & _
              " inner Join Clientes CNF" & vbCrLf & _
              "    on CNF.Cliente_id  = NF.Cliente_id" & vbCrLf & _
              "   and CNF.Endereco_id = NF.EndCliente_Id" & vbCrLf & _
              " inner join suboperacoes sop" & vbCrLf & _
              "    on nf.operacao = sop.operacao_id" & vbCrLf & _
              "   and nf.suboperacao = sop.suboperacoes_id" & vbCrLf & _
              " Inner Join Pedidos Pe" & vbCrLf & _
              "    on Pe.Pedido_id = NF.Pedido" & vbCrLf & _
              "  Left join Pagamentos Pa" & vbCrLf & _
              "    on Pe.CondicaoPagamento = Pa.Pagamento_Id" & vbCrLf & _
              "  left Join Comissoes C" & vbCrLf & _
              "    on C.principal = 'S'" & vbCrLf & _
              "   And C.Pedido_Id = Pe.Pedido_Id" & vbCrLf & _
              "  left Join Clientes Cli" & vbCrLf & _
              "    on Cli.Cliente_id  = C.Representante_Id" & vbCrLf & _
              "   and Cli.Endereco_id = C.EndRepresentante_Id" & vbCrLf & _
              " where NF.empresa_Id   = '" & objEmpresa.Codigo & "'" & vbCrLf & _
              "   And NF.Movimento    = '" & PeriodoInicial.ToSqlDate() & "'" & vbCrLf & _
              "   and Exists(select *" & vbCrLf & _
              "                from notasfiscaisXitens nfi" & vbCrLf & _
              "               inner join Produtos P" & vbCrLf & _
              "                  on P.Produto_Id = nfi.Produto_Id" & vbCrLf & _
              "               where NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
              "                 AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
              "                 AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
              "                 AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
              "                 AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
              "                 AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
              "                 AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
              "                 AND P.Marca            = 1)" & vbCrLf
        '"   And NF.Movimento  between '2010-12-01' and '2010-12-13'" & vbCrLf & _
        '"   And NF.Movimento    = '" & Format(PeriodoInicial, "yyyy-MM-dd") & "'" & vbCrLf & _

        ds.Merge(Banco.ConsultaDataSet(Sql, "T10"))


        If ds.Tables("T10").Rows.Count > 0 Then
            For Each dr In ds.Tables("T10").Rows
                linha = "10"                                                 'ID
                linha &= Funcoes.AlinharDireita(dr("Nota_Id"), 9, "0")               'Numero Da Nota
                linha &= Funcoes.AlinharEsquerda(dr("Serie_Id"), 3, " ")             'Serie Da Nota
                linha &= Funcoes.AlinharEsquerda(dr("Tipo"), 25, " ")                'Tipo Da Nota
                linha &= Funcoes.AlinharEsquerda(dr("Descricao"), 40, " ")           'Descricao Do Tipo Da Nota
                linha &= Funcoes.AlinharEsquerda(dr("CondicaoPagamento"), 25, " ")   'Condicao De Pagamento
                linha &= Funcoes.AlinharEsquerda(dr("DescPagamento"), 40, " ")       'Descricao Da Forma De Pagamento
                linha &= Format(CDate(dr("DataVencimento")), "yyyyMMdd")     'Data Vencimento
                linha &= Format(CDate(dr("DataDaNota")), "yyyyMMdd")         'Data Nota
                linha &= Funcoes.AlinharEsquerda(dr("Vendedor"), 14, " ")            'CPF Vendedor
                Dim Vendedor() As String = dr("NomeVendedor").ToString.Trim.Split(" ")
                linha &= Funcoes.AlinharEsquerda(dr("NomeVendedor").ToString.Replace(Vendedor(Vendedor.Length - 1), ""), 30, " ") 'Nome Do Vendedor
                linha &= Funcoes.AlinharEsquerda(Vendedor(Vendedor.Length - 1), 20, " ")                                          'SobreNome Do Vendedor

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                RegistroGeral += 1

                '***********************************************************************************************
                '*** Dados da Nota Fiscal
                '***--------------------------------------------------------------------------------------------
                '*** ID 11
                '***********************************************************************************************
                linha = "11"
                linha &= Funcoes.AlinharEsquerda(dr("Cliente_Id"), 14, " ")
                linha &= Funcoes.AlinharEsquerda(" ", 8, " ")
                linha &= Funcoes.AlinharEsquerda(dr("ClienteNF"), 40, " ")
                linha &= Funcoes.AlinharEsquerda(dr("CEP").ToString.Replace("-", ""), 9, " ")

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                RegistroGeral += 1

                '***********************************************************************************************
                '*** Dados do Item Da Nota Fiscal
                '***--------------------------------------------------------------------------------------------
                '*** ID 12
                '***********************************************************************************************
                Sql = "Select '12' as Tipo," & vbCrLf & _
                      "       NFI.Produto_Id," & vbCrLf & _
                      "       P.Nome," & vbCrLf & _
                      "       NFI.QuantidadeFiscal," & vbCrLf & _
                      "       NFI.Embalagem, " & vbCrLf & _
                      "       isnull(Emb.Descricao,'NAO INFORMADA') as DescEmbalagem, " & vbCrLf & _
                      "       NFI.unitario," & vbCrLf & _
                      "       'S' as IdentificadorDoFabricante," & vbCrLf & _
                      "       'BASF' as Fabricante," & vbCrLf & _
                      "       'BASF' as DescricaoFabricante" & vbCrLf & _
                      "  From NotasFiscaisxItens NFI" & vbCrLf & _
                      " inner join Produtos P" & vbCrLf & _
                      "    on P.Produto_Id = nfi.Produto_Id" & vbCrLf & _
                      "   And P.Marca      = 1" & vbCrLf & _
                      "  left join Embalagens Emb" & vbCrLf & _
                      "    on Emb.Embalagem_Id = NFI.Embalagem" & vbCrLf & _
                      " where NFI.Empresa_Id      ='" & dr("Empresa_Id") & "'" & vbCrLf & _
                      "   And NFI.EndEmpresa_Id   = " & dr("EndEmpresa_Id") & vbCrLf & _
                      "   And NFI.Cliente_Id      ='" & dr("Cliente_Id") & "'" & vbCrLf & _
                      "   And NFI.EndCliente_Id   = " & dr("EndCliente_Id") & vbCrLf & _
                      "   And NFI.EntradaSaida_Id ='" & dr("EntradaSaida_Id") & "'" & vbCrLf & _
                      "   And NFI.Serie_Id        ='" & dr("Serie_Id") & "'" & vbCrLf & _
                      "   And NFI.Nota_Id         = " & dr("Nota_Id") & vbCrLf

                ds2 = Banco.ConsultaDataSet(Sql, "T12")

                If ds2.Tables("T12").Rows.Count > 0 Then
                    For Each row In ds2.Tables("T12").Rows
                        linha = "12"                                         'ID
                        linha &= Funcoes.AlinharEsquerda(row("Produto_Id"), 40, " ") 'Produto
                        linha &= Funcoes.AlinharEsquerda(row("Nome"), 40, " ")       'Nome Do Produto
                        linha &= Funcoes.AlinharDireita(CInt(Math.Round(CDec(row("QuantidadeFiscal")), 3) * 1000).ToString, 18, "0")
                        linha &= Funcoes.AlinharEsquerda(row("DescEmbalagem"), 15, " ")
                        linha &= Funcoes.AlinharDireita(CInt(Math.Round(CDec(row("Unitario")), 2) * 100).ToString, 17, "0")
                        linha &= Funcoes.AlinharDireita("0", 17, "0")
                        linha &= Funcoes.AlinharEsquerda(" ", 10, " ")
                        linha &= "S"
                        linha &= "BASF"
                        linha &= "BASF"
                        linha &= Funcoes.AlinharEsquerda(" ", 15, " ")
                        linha &= Funcoes.AlinharEsquerda(" ", 40, " ")
                        linha &= Funcoes.AlinharEsquerda(" ", 15, " ")
                        linha &= Funcoes.AlinharEsquerda(" ", 50, " ")
                        linha &= Funcoes.AlinharDireita("0", 18, "0")

                        strm = New StreamWriter(NomeArquivo, True)
                        strm.WriteLine(linha)
                        strm.Close()
                        RegistroGeral += 1
                        ItensNota += 1
                    Next
                End If
            Next

            '***********************************************************************************************
            '*** Qtde De Itens Da Nota Fiscal
            '***--------------------------------------------------------------------------------------------
            '*** ID 13
            '***********************************************************************************************
            linha = "13"
            linha &= Funcoes.AlinharDireita(ItensNota.ToString, 5, "0")
            strm = New StreamWriter(NomeArquivo, True)
            strm.WriteLine(linha)
            strm.Close()
            QtdeNotaFiscal += 1
            RegistroGeral += 1
        End If

        '***********************************************************************************************
        '*** Registros Inventory - Estoque
        '***--------------------------------------------------------------------------------------------
        '*** ID 20
        '***********************************************************************************************

        Sql = "SELECT NotasFiscaisXItens.Produto_Id," & vbCrLf & _
              "       P.Nome," & vbCrLf & _
              "       isnull(NotasFiscaisXItens.Lote,'') as Lote," & vbCrLf & _
              " 	    isnull(NotasFiscaisXItens.Classificacao,'') as Classificacao," & vbCrLf & _
              "       isnull(emb.descricao,'') as Embalagem," & vbCrLf & _
              "	   sum(case" & vbCrLf & _
              "			 when NotasFiscaisXItens.EntradaSaida_Id = 'E'" & vbCrLf & _
              "			  then NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf & _
              "			  else 0" & vbCrLf & _
              "		   end) as Entradas," & vbCrLf & _
              "	   sum(case" & vbCrLf & _
              "			 when NotasFiscaisXItens.EntradaSaida_Id = 'S'" & vbCrLf & _
              "			  then NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf & _
              "			  else 0" & vbCrLf & _
              "		   end) as Saidas," & vbCrLf & _
              "	   sum(case" & vbCrLf & _
              "			 when NotasFiscaisXItens.EntradaSaida_Id = 'E'" & vbCrLf & _
              "			  then NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf & _
              "			  else NotasFiscaisXItens.QuantidadeFiscal * - 1" & vbCrLf & _
              "		   end) as Total" & vbCrLf & _
              "	  FROM NotasFiscais" & vbCrLf & _
              "	 INNER JOIN NotasFiscaisXItens " & vbCrLf & _
              "		ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
              "	   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
              "	   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
              "	   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
              "	   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
              "	   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
              "	   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
              "    Inner Join Produtos P" & vbCrLf & _
              "       on P.Produto_id = NotasFiscaisXItens.Produto_Id" & vbCrLf & _
              "      and P.Marca = 1" & vbCrLf & _
              "     Left join Embalagens Emb" & vbCrLf & _
              "       on Emb.Embalagem_Id = NotasFiscaisXItens.Embalagem" & vbCrLf & _
              "    where NotasFiscais.empresa_Id = '" & objEmpresa.Codigo & "'" & vbCrLf & _
              "      And NotasFiscais.Movimento  =  '" & PeriodoInicial.ToSqlDate() & "' " & vbCrLf & _
              " Group by NotasFiscaisXItens.Produto_Id," & vbCrLf & _
              "          P.Nome," & vbCrLf & _
              "		     isnull(NotasFiscaisXItens.Lote,'')," & vbCrLf & _
              "		     isnull(NotasFiscaisXItens.Classificacao,'')," & vbCrLf & _
              "          isnull(emb.descricao,'')" & vbCrLf
        '"      And NotasFiscais.Movimento between '2010-12-01' and '2010-12-13'" & vbCrLf & _
        '"      And NotasFiscais.Movimento  =  '" & Format(PeriodoInicial, "yyyy-MM-dd") & "' " & vbCrLf & _

        ds.Merge(Banco.ConsultaDataSet(Sql, "T20"))

        itensEstoque = 0
        If ds.Tables("T20").Rows.Count > 0 Then
            For Each dr In ds.Tables("T20").Rows
                linha = "20"                                        'ID
                linha &= Funcoes.AlinharEsquerda(dr("Produto_Id"), 40, " ") 'Produto
                linha &= Funcoes.AlinharEsquerda(dr("Nome"), 40, " ")       'Nome Do Produto
                linha &= Funcoes.AlinharEsquerda(dr("Embalagem"), 15, " ")  'Embalagem
                linha &= Funcoes.AlinharEsquerda(dr("Lote"), 10, " ")       'Lote
                linha &= Funcoes.AlinharEsquerda("1", 15, " ")
                linha &= Funcoes.AlinharEsquerda("BASF", 40, " ")
                linha &= Funcoes.AlinharDireita(CInt(CDec(dr("Entradas")) * 1000), 18, "0")
                linha &= Funcoes.AlinharDireita(CInt(CDec(dr("Saidas")) * 1000), 18, "0")

                If CDec(dr("Saidas")) > CDec(dr("Entradas")) Then
                    linha &= Funcoes.AlinharDireita("0", 18, "0")
                Else
                    linha &= Funcoes.AlinharDireita(CInt(CDec(dr("Total")) * 1000), 18, "0")
                End If
                linha &= Funcoes.AlinharDireita("0", 18, "0")

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                itensEstoque += 1
                RegistroGeral += 1
            Next

            '***********************************************************************************************
            '*** Qtde De Itens Movimentados no Estoque 
            '***--------------------------------------------------------------------------------------------
            '*** ID 21
            '***********************************************************************************************
            linha = "21"
            linha &= Funcoes.AlinharDireita(itensEstoque.ToString, 5, "0")
            strm = New StreamWriter(NomeArquivo, True)
            strm.WriteLine(linha)
            strm.Close()
            RegistroGeral += 1
        End If

        '***********************************************************************************************
        '*** Registros Final Custumer - "Cliente Final"
        '***--------------------------------------------------------------------------------------------
        '*** ID 30
        '***********************************************************************************************

        Sql = "Select Distinct NF.Cliente_Id, sb.Nome" & vbCrLf & _
              "  from NotasFiscais NF" & vbCrLf & _
              " inner join Clientes C" & vbCrLf & _
              "    on NF.Cliente_Id = C.Cliente_Id" & vbCrLf & _
              " inner join (select Cl.Cliente_id, Cl.Nome" & vbCrLf & _
              "               from Clientes Cl" & vbCrLf & _
              "              where Endereco_id = (select min(endereco_id)" & vbCrLf & _
              "                                     from clientes" & vbCrLf & _
              "                                    where cliente_id = Cl.Cliente_Id" & vbCrLf & _
              "                                   )" & vbCrLf & _
              "              ) sb" & vbCrLf & _
              "    on C.Cliente_Id = Sb.Cliente_id" & vbCrLf & _
              " Where NF.empresa_Id      = '" & objEmpresa.Codigo & "'" & vbCrLf & _
              "   and NF.EntradaSaida_Id = 'S'" & vbCrLf & _
              "   and (C.UsuarioAlteracaoData = '" & Date.Now.ToSqlDate() & "' or C.UsuarioInclusaoData = '" & Date.Now.ToSqlDate() & "')" & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(Sql, "T30"))

        If ds.Tables("T30").Rows.Count > 0 Then
            For Each dr In ds.Tables("T30").Rows
                linha = "30"                                        'ID
                linha &= Funcoes.AlinharEsquerda(dr("Cliente_Id"), 14, " ")
                linha &= Funcoes.AlinharEsquerda(dr("Nome"), 40, " ")
                linha &= "OUTROS"

                strm = New StreamWriter(NomeArquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                RegistroGeral += 1

                '***********************************************************************************************
                '*** Registros Final Custumer - "ENDEREÇOS"
                '***--------------------------------------------------------------------------------------------
                '*** ID 31
                '***********************************************************************************************
                Sql = "Select Distinct Endereco, numero, complemento, bairro, CEP" & vbCrLf & _
                      "  from Clientes " & vbCrLf & _
                      " Where Cliente_Id = '" & dr("Cliente_Id") & "'"

                ds2 = Banco.ConsultaDataSet(Sql, "T31")
                If ds2.Tables("T31").Rows.Count > 0 Then
                    For Each row In ds2.Tables("T31").Rows
                        linha = "31"
                        linha &= Funcoes.AlinharEsquerda(row("Endereco"), 40, " ")
                        linha &= Funcoes.AlinharEsquerda(row("numero"), 6, " ")
                        linha &= Funcoes.AlinharEsquerda(row("complemento"), 10, " ")
                        linha &= Funcoes.AlinharEsquerda(row("bairro"), 20, " ")
                        linha &= Funcoes.AlinharEsquerda(row("cep").ToString.Replace("-", ""), 9, " ")

                        strm = New StreamWriter(NomeArquivo, True)
                        strm.WriteLine(linha)
                        strm.Close()
                        RegistroGeral += 1
                    Next
                End If

                '***********************************************************************************************
                '*** Registros "ANO APLICACAO"
                '***--------------------------------------------------------------------------------------------
                '*** ID 32
                '***********************************************************************************************
                Sql = "Select Distinct CC.Ano_Id" & vbCrLf & _
                      "  from Clientes C" & vbCrLf & _
                      " inner join ClientexCultura CC" & vbCrLf & _
                      "    on CC.Cliente_Id     =  C.Cliente_Id" & vbCrLf & _
                      "   and CC.EndCliente_id  =  C.endereco_Id" & vbCrLf & _
                      " Where CC.Cliente_Id     =  '" & dr("Cliente_Id") & "'" & vbCrLf & _
                      "   and len(C.Cliente_Id) <  14" & vbCrLf

                ds2 = Banco.ConsultaDataSet(Sql, "T32")
                If ds2.Tables("T32").Rows.Count > 0 Then
                    For Each row In ds2.Tables("T32").Rows
                        linha = "32"
                        linha &= Funcoes.AlinharDireita(row("Ano_Id"), 4, "0")

                        strm = New StreamWriter(NomeArquivo, True)
                        strm.WriteLine(linha)
                        strm.Close()
                        RegistroGeral += 1

                        '***********************************************************************************************
                        '*** Registros "APLICACAO"
                        '***--------------------------------------------------------------------------------------------
                        '*** ID 33
                        '***********************************************************************************************
                        Sql = "Select 99 as Cultura_Id, 'outros' as Descricao, sum(AreaPlantada) as AreaPlantada" & vbCrLf & _
                              "  from Clientes C" & vbCrLf & _
                              " inner join ClientexCultura CC" & vbCrLf & _
                              "    on CC.Cliente_Id    = C.Cliente_Id" & vbCrLf & _
                              "   and CC.EndCliente_id = C.endereco_Id" & vbCrLf & _
                              " inner join Cultura Cu" & vbCrLf & _
                              "    on Cu.Cultura_id = CC.Cultura_id" & vbCrLf & _
                              " Where C.Cliente_id      = '" & dr("Cliente_Id") & "' " & vbCrLf & _
                              "   and len(C.Cliente_Id) < 14" & vbCrLf & _
                              "   And CC.Ano_Id         = " & row("ANO_ID") & vbCrLf

                        'Sql = "Select CC.Cultura_Id, Cu.Descricao, sum(AreaPlantada) as AreaPlantada" & vbCrLf & _
                        '      "  from Clientes C" & vbCrLf & _
                        '      " inner join ClientexCultura CC" & vbCrLf & _
                        '      "    on CC.Cliente_Id    = C.Cliente_Id" & vbCrLf & _
                        '      "   and CC.EndCliente_id = C.endereco_Id" & vbCrLf & _
                        '      " inner join Cultura Cu" & vbCrLf & _
                        '      "    on Cu.Cultura_id = CC.Cultura_id" & vbCrLf & _
                        '      " Where C.Cliente_id      = '" & dr("Cliente_Id") & "' " & vbCrLf & _
                        '      "   and len(C.Cliente_Id) < 14" & vbCrLf & _
                        '      "   And CC.Ano_Id         = " & row("ANO_ID") & vbCrLf & _
                        '      " group by CC.Cultura_Id, Cu.Descricao" & vbCrLf

                        ds3 = Banco.ConsultaDataSet(Sql, "T33")
                        NumLinhaAplicacao = 0

                        If ds3.Tables("T33").Rows.Count > 0 Then
                            For Each row2 In ds3.Tables("T33").Rows
                                linha = "33"
                                linha &= Funcoes.AlinharEsquerda(row2("Cultura_Id"), 15, " ")
                                linha &= Funcoes.AlinharEsquerda(row2("Descricao"), 40, " ")
                                linha &= Funcoes.AlinharDireita(CInt(row2("AreaPlantada")), 10, "0")
                                linha &= Funcoes.AlinharEsquerda(" ", 15, " ")

                                strm = New StreamWriter(NomeArquivo, True)
                                strm.WriteLine(linha)
                                strm.Close()
                                NumLinhaAplicacao += 1
                                RegistroGeral += 1
                            Next
                        End If
                        '***********************************************************************************************
                        '*** Registros "NUMERO DE APLICACOES"
                        '***--------------------------------------------------------------------------------------------
                        '*** ID 34
                        '***********************************************************************************************
                        If NumLinhaAplicacao > 0 Then
                            linha = "34"
                            linha &= Funcoes.AlinharDireita(NumLinhaAplicacao.ToString, 5, "0")

                            strm = New StreamWriter(NomeArquivo, True)
                            strm.WriteLine(linha)
                            strm.Close()
                            RegistroGeral += 1
                        End If
                    Next
                Else
                    '************************************  32   **************************************************
                    linha = "32"
                    linha &= Funcoes.AlinharDireita(Date.Now.Year.ToString, 4, "0")

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    RegistroGeral += 1

                    '************************************  33   **************************************************
                    linha = "33"
                    linha &= Funcoes.AlinharEsquerda("99", 15, " ")
                    linha &= Funcoes.AlinharEsquerda("outros", 40, " ")
                    linha &= Funcoes.AlinharDireita(0, 10, "0")
                    linha &= Funcoes.AlinharEsquerda(" ", 15, " ")
                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    RegistroGeral += 1

                    '************************************  34   **************************************************
                    linha = "34"
                    linha &= Funcoes.AlinharDireita(1, 5, "0")

                    strm = New StreamWriter(NomeArquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    RegistroGeral += 1

                End If
                QtdeCliente += 1
            Next
        End If

        '***********************************************************************************************
        '*** Registros "TOTAIS DE REGISTRO"
        '***--------------------------------------------------------------------------------------------
        '*** ID 40
        '***********************************************************************************************
        linha = "40"
        linha &= Funcoes.AlinharDireita(QtdeNotaFiscal.ToString, 5, "0")
        linha &= Funcoes.AlinharDireita(itensEstoque.ToString, 5, "0")
        linha &= Funcoes.AlinharDireita(QtdeCliente.ToString, 5, "0")

        strm = New StreamWriter(NomeArquivo, True)
        strm.WriteLine(linha)
        strm.Close()

        MsgBox(Me.Page, "Processo concluído.")
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Conecta")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class