Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaRomaneios
    Inherits BaseUserControl

    Dim sql As String = String.Empty
    Dim Unitario As String = String.Empty
    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
        updConsultaRomaneios.Update()
    End Sub

    Public Sub CargaRomaneios()
        SessaoRecuperaNotaFiscal()
        Dim ds As New DataSet
        sql = "SELECT R.Romaneio_Id AS Romaneio, R.EntradaSaida AS ES, R.Pedido, R.Produto, " & vbCrLf & _
              "       Pr.Nome AS NomeDoProduto, R.Operacao, R.SubOperacao, R.Movimento, " & vbCrLf & _
              "		  CASE R.Processo " & vbCrLf & _
              "           WHEN 'AGRUPAMENTO' " & vbCrLf & _
              "			      THEN P.BrutoBalanca " & vbCrLf & _
              "			      ELSE R.PesoBruto " & vbCrLf & _
              "		  END AS Bruto, " & vbCrLf & _
              "		  CASE R.Processo " & vbCrLf & _
              "           WHEN 'AGRUPAMENTO' " & vbCrLf & _
              "				  THEN (P.BrutoBalanca - P.Liquido) " & vbCrLf & _
              "				  ELSE R.Desconto " & vbCrLf & _
              "		  END AS Desconto, " & vbCrLf & _
              "		  CASE R.Processo " & vbCrLf & _
              "           WHEN 'AGRUPAMENTO' " & vbCrLf & _
              "				  THEN P.Liquido " & vbCrLf & _
              "				  ELSE R.PesoLiquido " & vbCrLf & _
              "		   END AS Liquido, " & vbCrLf & _
              "        RxP.Pesagem_Id AS Laudo " & vbCrLf & _
              "   FROM Romaneios R " & vbCrLf & _
              "  INNER JOIN Produtos Pr" & vbCrLf & _
              "     ON R.Produto = Pr.Produto_Id " & vbCrLf & _
              "  INNER JOIN RomaneiosXPesagens RxP " & vbCrLf & _
              "     ON R.Empresa_Id    = RxP.Empresa_Id " & vbCrLf & _
              "    AND R.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf & _
              "    AND R.Romaneio_Id   = RxP.Romaneio_Id " & vbCrLf & _
              "  INNER JOIN Pesagem P" & vbCrLf & _
              "     ON RxP.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
              "    AND RxP.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
              "    AND RxP.Pesagem_Id    = P.Pesagem_Id " & vbCrLf & _
              "    AND RxP.Sequencia_Id  = P.Sequencia_Id " & vbCrLf & _
              "   FULL OUTER JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf & _
              "     ON R.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
              "    AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
              "    AND R.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
              "   LEFT JOIN NotasFiscais NF" & vbCrLf & _
              "     ON nfxr.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
              "    AND nfxr.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
              "    AND nfxr.Cliente_Id      = NF.Cliente_id" & vbCrLf & _
              "    AND nfxr.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
              "    AND nfxr.Nota_Id         = NF.Nota_Id" & vbCrLf & _
              "    AND nfxr.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
              "    AND nfxr.Serie_Id        = NF.Serie_Id " & vbCrLf

        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objRomaneioNXI") OrElse Session("ssTipoRetorno").ToString.Contains("RomaneioNFxP")) Then
            sql &= "  WHERE R.Empresa_id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                   "    AND R.EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                   "    AND R.Pedido        ='" & objNotaFiscal.CodigoPedido & "'" & vbCrLf & _
                   "    AND R.Operacao      = " & Session("Op" & HID.Value.ToString) & vbCrLf & _
                   "    AND R.SubOperacao   = " & Session("SOp" & HID.Value.ToString) & vbCrLf & _
                   "    AND P.Situacao        = 1 " & vbCrLf & _
                   "    AND (nfxr.Nota_Id IS NULL OR NF.situacao NOT IN (1,4,7))  " & vbCrLf
            Unitario = Session("Unitario" & HID.Value.ToString)
            HttpContext.Current.Session("txtPedido") = objNotaFiscal.CodigoPedido
            HttpContext.Current.Session("ssCnpjDaEmpresa") = objNotaFiscal.CodigoEmpresa
            HttpContext.Current.Session("ssEndDaEmpresa") = objNotaFiscal.EnderecoEmpresa
        Else
            sql &= "  WHERE (R.Pedido = " & IIf(objNotaFiscal.CodigoPedido = 0, HttpContext.Current.Session("txtPedido"), objNotaFiscal.CodigoPedido) & ") " & vbCrLf &
                   "    AND (P.Situacao = 1) " & vbCrLf &
                   "    AND (nfxr.Nota_Id IS NULL OR NF.situacao NOT IN (1,4,7))  " & vbCrLf
            Unitario = HttpContext.Current.Session("txtUnitarioOficial")
        End If

        sql &= "  ORDER BY R.Movimento "

        ds = Banco.ConsultaDataSet(sql, "Romaneios")
        GridRomaneios.DataSource = ds
        GridRomaneios.DataBind()
        updConsultaRomaneios.Update()
    End Sub

    Protected Sub GridRomaneios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridRomaneios.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objRomaneioNXI")) Then
            Session("objRomaneioNXI" & HID.Value.ToString) = New [Lib].Negocio.Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, GridRomaneios.SelectedRow.Cells(1).Text)
            Popup.CloseDialog(Me.Page, "divConsultaRomaneios")
            CType(Me.Page, NotaFiscalXItens).CarregarRomaneio()
            Exit Sub
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objRomaneioALT")) Then
            Session("objRomaneioALT" & HID.Value.ToString) = New [Lib].Negocio.Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, GridRomaneios.SelectedRow.Cells(1).Text)
            Popup.CloseDialog(Me.Page, "divConsultaRomaneios")
            CType(Me.Page, AlterarNotaFiscal).CarregarRomaneio()
            Exit Sub
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("RomaneioNFxP")) Then
            Session("RomaneioNFxP" & HID.Value.ToString) = New [Lib].Negocio.Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, GridRomaneios.SelectedRow.Cells(1).Text)
            Popup.CloseDialog(Me.Page, "divConsultaRomaneios")
            'CType(Me.Page, NotaFiscalXPesagem).AtribuirValoresCampos()
            Exit Sub
        End If

        Dim Romaneio As String = ""
        Dim Produto As String = ""
        Dim NomeDoProduto As String = ""
        Dim DescricaoDoProduto As String = ""
        Dim NcmDoProduto As String = ""
        Dim UnidadeDoProduto As String = ""
        'Dim PisCofinsIntegral As String = ""
        'Dim PisCofinsPresumido As String = ""

        Dim PesoBruto As String = ""
        Dim PesoLiquido As String = ""
        Dim ValorTotal As String = ""

        Dim NomeDoTransportador As String = ""
        Dim Placas As String = ""
        Dim Placa As String = ""
        Dim Grupo As String = ""

        Dim Operacao As String = ""
        Dim SubOperacao As String = ""
        Dim DescricaoDaOperacao As String = ""
        Dim Devolucao As String = ""

        Dim EstadoDaPlaca As String = ""
        Dim EnderecoDoTransportador As String = ""
        Dim CidadeDoTransportador As String = ""
        Dim EstadoDoTransportador As String = ""
        Dim CnpjDoTransportador As String = ""
        Dim EndCnpjDoTransportador As String = ""
        Dim InscricaoDoTransportador As String = ""

        Dim EntradaSaida As String = "E"
        Dim Registro As String = ""
        Dim Serie As String = ""
        Dim SerieDaNota As String = ""
        Dim NumeroDaNota As Integer = 0
        Dim PesoFiscal As String = ""

        sql = " SELECT R.Empresa_Id, R.EndEmpresa_Id, R.Romaneio_Id, R.EntradaSaida, R.Pedido, R.Deposito, " & vbCrLf & _
              "        R.EndDeposito, R.Destino, R.EndDestino, R.Produto, R.Operacao, R.SubOperacao, SOP.Descricao as DescricaoOperacao," & vbCrLf & _
              "        R.Movimento, R.PesoBruto, R.Desconto, R.PesoLiquido, R.Observacoes, Pr.Grupo, " & vbCrLf & _
              "        Pr.Nome AS NomeDoProduto, Pr.Descricao, Pr.NCM, Pr.Unidade, P.Transportador, P.EndTransportador, " & vbCrLf & _
              "        ISNULL(Transp.Nome, '') AS NomeDoTransportador, ISNULL(Transp.Endereco, '') AS EnderecoDoTransportador, " & vbCrLf & _
              "        ISNULL(Transp.Cidade, '') AS CidadeDoTransportador, ISNULL(Transp.Estado, '') AS EstadoDoTransportador, " & vbCrLf & _
              "        ISNULL(Transp.Inscricao, '') AS InscricaoDoTransportador, Pl.Placa_Id, Pl.Placa01, Pl.Placa02, Pl.Placa03, " & vbCrLf & _
              "        Pl.EstadoPlaca AS EstadoDaPlaca, P.NumeroDaNota, P.SerieDaNota, P.PesoFiscal, SOP.Devolucao " & vbCrLf & _
              "   FROM Romaneios R " & vbCrLf & _
              "  INNER JOIN Produtos Pr " & vbCrLf & _
              "     ON R.Produto = Pr.Produto_Id " & vbCrLf & _
              "  INNER JOIN RomaneiosXPesagens RxP " & vbCrLf & _
              "     ON R.Empresa_Id    = RxP.Empresa_Id  " & vbCrLf & _
              "    AND R.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf & _
              "    AND R.Romaneio_Id   = RxP.Romaneio_Id " & vbCrLf & _
              "  INNER JOIN Pesagem P" & vbCrLf & _
              "     ON RxP.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
              "    AND RxP.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
              "    AND RxP.Pesagem_Id    = P.Pesagem_Id " & vbCrLf & _
              "  INNER JOIN Placas Pl " & vbCrLf & _
              "     ON P.Placa = Pl.Placa_Id " & vbCrLf & _
              "  INNER JOIN SubOperacoes SOP" & vbCrLf & _
              "     ON R.Operacao    = SOP.Operacao_Id " & vbCrLf & _
              "    AND R.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf & _
              "   LEFT OUTER JOIN Clientes AS Trasp " & vbCrLf & _
              "     ON P.Transportador    = Transp.Cliente_Id  " & vbCrLf & _
              "    AND P.EndTransportador = Transp.Endereco_Id" & vbCrLf & _
              "  WHERE R.Empresa_Id    = '" & HttpContext.Current.Session("ssCnpjDaEmpresa") & "'" & vbCrLf & _
              "    AND R.EndEmpresa_Id = " & HttpContext.Current.Session("ssEndDaEmpresa") & vbCrLf & _
              "    AND R.Romaneio_Id   = " & GridRomaneios.SelectedRow.Cells(1).Text() & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Romaneios").Tables(0).Rows
            Romaneio = Dr("Romaneio_Id")
            Produto = Dr("Produto")
            NomeDoProduto = UCase(Dr("NomeDoProduto"))
            DescricaoDoProduto = UCase(Dr("Descricao"))
            NcmDoProduto = Dr("NCM")
            UnidadeDoProduto = UCase(Dr("Unidade"))
            PesoBruto = Dr("PesoBruto")
            PesoLiquido = Dr("PesoLiquido")
            NumeroDaNota = Dr("NumeroDaNota")
            SerieDaNota = Dr("SerieDaNota")
            PesoFiscal = Dr("PesoFiscal")
            Grupo = Dr("Grupo")
            NomeDoTransportador = UCase(Dr("NomeDoTransportador"))
            Placa = UCase(RTrim(Dr("Placa_Id")))
            Placas = UCase(Dr("Placa_Id")) & " " & UCase(Dr("Placa01")) & " " & UCase(Dr("Placa02")) & " " & UCase(Dr("Placa03"))
            EstadoDaPlaca = UCase(Dr("EstadoDaPlaca"))
            EnderecoDoTransportador = UCase(Dr("EnderecoDoTransportador"))
            CidadeDoTransportador = UCase(Dr("CidadeDoTransportador"))
            EstadoDoTransportador = UCase(Dr("EstadoDoTransportador"))
            CnpjDoTransportador = Dr("Transportador")
            EndCnpjDoTransportador = Dr("EndTransportador")
            InscricaoDoTransportador = Dr("InscricaoDoTransportador")
            EntradaSaida = Dr("EntradaSaida")
            Operacao = Dr("Operacao")
            SubOperacao = Dr("SubOperacao")
            DescricaoDaOperacao = UCase(Dr("DescricaoOperacao"))
            Devolucao = UCase(Dr("Devolucao"))
        Next

        '''''''''''''''''''''''''''''''
        'Consulta Sequencia de Notas'''
        ''''''''''''''''''''''''''''''

        If EntradaSaida = "E" Then
            If NumeroDaNota = 0 Then
                sql = " SELECT SUM(Sequencia + 1) as Sequencia " & vbCrLf & _
                      "   FROM Numerador " & vbCrLf & _
                      "  WHERE Empresa_Id = '" & HttpContext.Current.Session("ssCnpjDaEmpresa") & "'" & vbCrLf & _
                      "    AND EndEmpresa_Id = " & HttpContext.Current.Session("ssEndDaEmpresa") & vbCrLf
                'Preciso desse Salsicho porque na Agricola Entrada e Saída usam a mesma sequência - Furlan - 17/02/2010
                If HttpContext.Current.Session("ssCnpjDaEmpresa") = "04854422000185" Then
                    sql &= " AND Numerador_Id = 30"
                Else
                    sql &= " AND Numerador_Id = 20"
                End If

                For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Numerador").Tables(0).Rows
                    Registro = Dr("Sequencia")
                    'Preciso desse Salsicho porque na Agricola Entrada e Saída usam a mesma sequência - Furlan - 17/02/2010
                    If HttpContext.Current.Session("ssCnpjDaEmpresa") = "04854422000185" Then
                        Serie = "1"
                    Else
                        Serie = "2"
                    End If
                    HttpContext.Current.Session("txtNossaEmissao") = "S"
                Next
            Else
                Registro = NumeroDaNota
                Serie = SerieDaNota
                HttpContext.Current.Session("txtNossaEmissao") = "N"
            End If
        Else
            HttpContext.Current.Session("txtNossaEmissao") = "S"

            sql = " SELECT SUM(Sequencia + 1) AS Sequencia " & vbCrLf & _
                  "   FROM Numerador " & vbCrLf & _
                  "  WHERE Empresa_Id = '" & HttpContext.Current.Session("ssCnpjDaEmpresa") & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id = " & HttpContext.Current.Session("ssEndDaEmpresa") & vbCrLf & _
                  "    AND Numerador_Id = 30 " & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Numerador").Tables(0).Rows
                Registro = Dr("Sequencia")
                Serie = "1"
            Next
        End If

        If PesoFiscal = "0" Then
            PesoFiscal = PesoLiquido
        End If

        'Se o Saldo Fisal Pedido menor que o Romaneio vale o Pedido - Furlan - 29/01/2009
        If CInt(HttpContext.Current.Session("txtSaldoPedido")) > 0 Then
            If CInt(HttpContext.Current.Session("txtSaldoPedido")) < CInt(PesoFiscal) And EntradaSaida = "S" Then
                PesoFiscal = CInt(HttpContext.Current.Session("txtSaldoPedido"))
            End If
        End If

        ValorTotal = CInt(PesoFiscal) * CDec(Unitario)
        HttpContext.Current.Session("txtRomaneio") = Romaneio
        HttpContext.Current.Session("txtNomeDoTransportador") = NomeDoTransportador
        HttpContext.Current.Session("txtPlaca") = Placa
        HttpContext.Current.Session("txtPlacas") = Placas
        HttpContext.Current.Session("txtEstadoDaPlaca") = EstadoDaPlaca
        HttpContext.Current.Session("txtEnderecoDoTransportador") = EnderecoDoTransportador
        HttpContext.Current.Session("txtCidadeDoTransportador") = CidadeDoTransportador
        HttpContext.Current.Session("txtEstadoDoTransportador") = EstadoDoTransportador
        HttpContext.Current.Session("txtCnpjDoTransportador") = CnpjDoTransportador
        HttpContext.Current.Session("txtEndCnpjDoTransportador") = EndCnpjDoTransportador
        HttpContext.Current.Session("txtInscricaoDoTransportador") = InscricaoDoTransportador
        HttpContext.Current.Session("txtProduto") = Produto
        HttpContext.Current.Session("txtGrupoDoProduto") = Grupo
        HttpContext.Current.Session("txtNomeDoProduto") = NomeDoProduto
        HttpContext.Current.Session("txtDescricaoDoProduto") = DescricaoDoProduto
        HttpContext.Current.Session("txtNcmDoProduto") = NcmDoProduto
        HttpContext.Current.Session("txtUnidadeDoProduto") = UnidadeDoProduto
        HttpContext.Current.Session("txtPesoBruto") = PesoBruto
        HttpContext.Current.Session("txtPesoLiquido") = PesoLiquido
        HttpContext.Current.Session("txtPesoFiscal") = PesoFiscal
        HttpContext.Current.Session("txtValorTotal") = ValorTotal
        HttpContext.Current.Session("txtOperacao") = Operacao
        HttpContext.Current.Session("txtSubOperacao") = SubOperacao
        HttpContext.Current.Session("txtDescricaoDaOperacao") = DescricaoDaOperacao
        HttpContext.Current.Session("txtES") = EntradaSaida
        HttpContext.Current.Session("txtNumeredor") = Registro

        Popup.CloseDialog(Me.Page, "divConsultaRomaneios")
        CType(Me.Page, NotaFiscalXItens).CarregarRomaneio()
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Sub LimparCampos()
        GridRomaneios.DataSource = New List(Of Object)()
        GridRomaneios.DataBind()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaRomaneios")
        If TypeOf Me.Page Is NotaFiscalXItens Then
            CType(Me.Page, NotaFiscalXItens).CarregarRomaneio()
        End If
    End Sub

End Class