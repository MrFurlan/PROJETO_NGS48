Imports System.Data
Imports System.Data.SqlClient
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioEntradaSaidaNotasFiscalHTML
    Inherits BasePage

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Dim dsRelatorio As New DataSet
        Dim strSQL As String = BuscaSQLNotasFiscal()
        Dim drRelatorio As DataRow
        Dim intPos As Integer = 0

        'If (EnderecoConexao.SQlConnectionLocal Is Nothing) OrElse EnderecoConexao.SQlConnectionLocal.State = ConnectionState.Closed Then EnderecoConexao.ConectaBanco()
        'Dim daRelatorio As New SqlDataAdapter(strSQL, EnderecoConexao.SQlConnectionLocal)
        'daRelatorio.SelectCommand.CommandTimeout = 360
        'daRelatorio.Fill(dsRelatorio, "NotasFiscais")
        'EnderecoConexao.FecheConexaoBanco()
        dsRelatorio = Banco.ConsultaDataSet(strSQL, "NotasFiscais")

        writer.RenderBeginTag(HtmlTextWriterTag.Html)

        writer.RenderBeginTag(HtmlTextWriterTag.Head)
        writer.AddAttribute("http-equiv", "Content-Type")
        writer.AddAttribute(HtmlTextWriterAttribute.Content, "text/html; charset=iso-8859-1")
        writer.RenderBeginTag(HtmlTextWriterTag.Meta)
        writer.RenderEndTag() '</meta>
        writer.RenderBeginTag(HtmlTextWriterTag.Title)
        writer.Write("Posi&ccedil;&atilde;o de Cargas e Descargas")
        writer.RenderEndTag() '</title>
        writer.RenderEndTag() '</head>

        writer.RenderBeginTag(HtmlTextWriterTag.Body)

        '-----------------
        'Cabeçalho Padrao
        '-----------------
        writer.AddAttribute(HtmlTextWriterAttribute.Cellpadding, "0")
        writer.AddAttribute(HtmlTextWriterAttribute.Cellspacing, "0")
        writer.AddAttribute(HtmlTextWriterAttribute.Border, "1")
        writer.RenderBeginTag(HtmlTextWriterTag.Table)

        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Empresa_Id")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("EndEmpresa_Id")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("EmpresaReduzido")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Movimento")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("DataNota")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Nota_Id")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Serie_Id")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("EntradaSaida_Id")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Produto")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Operacao")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("SubOperacao")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Finalidade")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Cliente_Id")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Deposito")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("DepositoDestino")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("PesoBruto")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Desconto")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("PesoLiquido")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Variacao")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("Valor")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("ValorICMS")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("ValorFunrural")
        writer.RenderEndTag() '</th>
        writer.RenderBeginTag(HtmlTextWriterTag.Th)
        writer.Write("ValorFethab")
        writer.RenderEndTag() '</th>
        writer.RenderEndTag() '</tr>

        Dim strEmpresaAnt As String = ""

        For Each drRelatorio In dsRelatorio.Tables(0).Rows
            writer.RenderBeginTag(HtmlTextWriterTag.Tr)
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Empresa_Id").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("EndEmpresa_Id").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("EmpresaReduzido").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Movimento").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("DataNota").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Nota_Id").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Serie_Id").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("EntradaSaida_Id").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Produto").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Operacao").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("SubOperacao").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Finalidade").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Cliente_Id").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Deposito").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("DepositoDestino").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("PesoBruto").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Desconto").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("PesoLiquido").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Variacao").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("Valor").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("ValorIcms").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("ValorFunrunral").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write(drRelatorio("ValorFethab").ToString())
            writer.RenderEndTag() '</td>
            writer.RenderEndTag() '</tr>
        Next

        writer.RenderEndTag() '</table>
        writer.RenderEndTag() '</body>
        writer.RenderEndTag() '</html>
    End Sub

    Private Function BuscaSQLNotasFiscal() As String
        Dim dsRelatorio As New DataSet

        Dim strSQL As String = "SELECT NF.Empresa_Id, NF.EndEmpresa_Id, E.Nome AS EmpresaNome, E.Reduzido AS EmpresaReduzido, " & _
                                       "E.Cidade AS EmpresaCidade, E.Estado AS EmpresaEstado, NF.Movimento, NF.DataDaNota AS DataNota, " & _
                                       "NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id, NFI.Produto_Id AS Produto, NF.Operacao, " & _
                                       "NF.SubOperacao, NF.Finalidade, NF.Cliente_Id, NF.EndCliente_Id, COALESCE(DO.Reduzido, '') AS Deposito, " & _
                                       "COALESCE(DD.Reduzido, '') AS DepositoDestino, R.PesoBruto, R.Desconto, " & _
                                       "R.PesoLiquido, 0 AS Variacao, NFI.Valor, COALESCE (ICMS.Valor, 0) AS ValorIcms, " & _
                                       "COALESCE (FUNRURAL.Valor, 0) AS ValorFunrunral, " & _
                                       "COALESCE (FETHAB.Valor, 0) AS ValorFethab, P.Grupo " & _
                                       "FROM NotasFiscais NF " & _
                                       "INNER JOIN Clientes AS E " & _
                                       "ON NF.Empresa_Id = E.Cliente_Id " & _
                                       "AND NF.EndEmpresa_Id = E.Endereco_Id " & _
                                       "INNER JOIN NotasFiscaisXItens AS NFI " & _
                                       "ON NFI.Empresa_Id = NF.Empresa_Id " & _
                                       "AND NFI.EndEmpresa_Id = NF.EndEmpresa_Id " & _
                                       "AND NFI.Cliente_Id = NF.Cliente_Id " & _
                                       "AND NFI.EndCliente_Id = NF.EndCliente_Id " & _
                                       "AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & _
                                       "AND NFI.Serie_Id = NF.Serie_Id " & _
                                       "AND NFI.Nota_Id = NF.Nota_Id " & _
                                       "INNER JOIN NotasFiscaisXRomaneios NFR " & _
                                       "ON NFR.Empresa_Id = NFI.Empresa_Id " & _
                                       "AND NFR.EndEmpresa_Id = NFI.EndEmpresa_Id " & _
                                       "AND NFR.Cliente_Id = NFI.Cliente_Id " & _
                                       "AND NFR.EndCliente_Id = NFI.EndCliente_Id " & _
                                       "AND NFR.EntradaSaida_Id = NFI.EntradaSaida_Id " & _
                                       "AND NFR.Serie_Id = NFI.Serie_Id " & _
                                       "AND NFR.Nota_Id = NFI.Nota_Id " & _
                                       "INNER JOIN Romaneios R " & _
                                       "ON NFR.Empresa_Id = R.Empresa_Id " & _
                                       "AND NFR.EndEmpresa_Id = R.EndEmpresa_Id " & _
                                       "AND NFR.Romaneio_Id = R.Romaneio_Id " & _
                                       "LEFT JOIN Clientes DO " & _
                                       "ON R.Deposito = DO.Cliente_Id " & _
                                       "AND R.EndDeposito = DO.Endereco_Id " & _
                                       "LEFT JOIN Clientes DD " & _
                                       "ON R.Destino = DD.Cliente_Id " & _
                                       "AND R.EndDestino = DD.Endereco_Id " & _
                                       "LEFT JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, " & _
                                                         "Nota_Id, Produto_Id, Cfop_Id, Encargo_Id, COALESCE(Valor, 0) AS Valor " & _
                                                  "FROM NotasFiscaisXEncargos) ICMS " & _
                                       "ON ICMS.Empresa_Id = NFI.Empresa_Id " & _
                                       "AND ICMS.EndEmpresa_Id = NFI.EndEmpresa_Id " & _
                                       "AND ICMS.Cliente_Id = NFI.Cliente_Id " & _
                                       "AND ICMS.EndCliente_Id = NFI.EndCliente_Id " & _
                                       "AND ICMS.EntradaSaida_Id = NFI.EntradaSaida_Id " & _
                                       "AND ICMS.Serie_Id = NFI.Serie_Id " & _
                                       "AND ICMS.Nota_Id = NFI.Nota_Id " & _
                                       "AND ICMS.Produto_Id = NFI.Produto_Id " & _
                                       "AND ICMS.Encargo_Id = 'ICMS' " & _
                                       "LEFT JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id," & _
                                                         "Nota_Id, Produto_Id, Cfop_Id, Encargo_Id, COALESCE(Valor, 0) AS Valor " & _
                                                  "FROM NotasFiscaisXEncargos) FUNRURAL " & _
                                       "ON FUNRURAL.Empresa_Id = NFI.Empresa_Id " & _
                                       "AND FUNRURAL.EndEmpresa_Id = NFI.EndEmpresa_Id " & _
                                       "AND FUNRURAL.Cliente_Id = NFI.Cliente_Id " & _
                                       "AND FUNRURAL.EndCliente_Id = NFI.EndCliente_Id " & _
                                       "AND FUNRURAL.EntradaSaida_Id = NFI.EntradaSaida_Id " & _
                                       "AND FUNRURAL.Serie_Id = NFI.Serie_Id " & _
                                       "AND FUNRURAL.Nota_Id = NFI.Nota_Id " & _
                                       "AND FUNRURAL.Produto_Id = NFI.Produto_Id " & _
                                       "AND (FUNRURAL.Encargo_Id = 'FUNRURAL' or FUNRURAL.Encargo_Id = 'FUNRURAL JUDICIAL' or FUNRURAL.Encargo_Id = 'SENAR') " & _
                                       "LEFT JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, " & _
                                                         "Nota_Id, Produto_Id, Cfop_Id, Encargo_Id, COALESCE(Valor, 0) AS Valor " & _
                                                  "FROM NotasFiscaisXEncargos) FETHAB " & _
                                       "ON FETHAB.Empresa_Id = NFI.Empresa_Id " & _
                                       "AND FETHAB.EndEmpresa_Id = NFI.EndEmpresa_Id " & _
                                       "AND FETHAB.Cliente_Id = NFI.Cliente_Id " & _
                                       "AND FETHAB.EndCliente_Id = NFI.EndCliente_Id " & _
                                       "AND FETHAB.EntradaSaida_Id = NFI.EntradaSaida_Id " & _
                                       "AND FETHAB.Serie_Id = NFI.Serie_Id " & _
                                       "AND FETHAB.Nota_Id = NFI.Nota_Id " & _
                                       "AND FETHAB.Produto_Id = NFI.Produto_Id " & _
                                       "AND FETHAB.Encargo_Id = 'FETHAB' " & _
                                       "INNER JOIN Produtos P " & _
                                       "ON P.Produto_Id = NFI.Produto_Id " & _
                                       "WHERE NF.Movimento BETWEEN '" & Request.QueryString("de") & "' " & _
                                       "AND '" & Request.QueryString("ate") & "' " & _
                                       "AND NF.Empresa_Id = '" & Request.QueryString("empresa") & "' " & _
                                       "AND NF.EndEmpresa_Id = " & Request.QueryString("endempresa") & " "

        If Not Request.QueryString("es") Is Nothing Then strSQL &= "AND NF.EntradaSaida_Id = '" & Request.QueryString("es") & "' "

        If Not Request.QueryString("cliente") Is Nothing Then
            strSQL &= "AND NF.Cliente_Id = '" & Request.QueryString("cliente") & "' " & _
                      "AND NF.EndCliente_Id = " & Request.QueryString("endcliente") & " "
        End If

        If Not Request.QueryString("grupo") Is Nothing Then strSQL &= "AND P.Grupo = '" & Request.QueryString("grupo") & "' "
        If Not Request.QueryString("produto") Is Nothing Then strSQL &= "AND P.Produto = '" & Request.QueryString("produto") & "' "
        If Not Request.QueryString("operacao") Is Nothing Then strSQL &= "AND NF.Operacao = " & Request.QueryString("operacao") & " "
        If Not Request.QueryString("suboperacao") Is Nothing Then strSQL &= "AND NF.SubOperacao = " & Request.QueryString("suboperacao") & " "
        If Not Request.QueryString("finalidade") Is Nothing Then strSQL &= "AND NF.Finalidade = " & Request.QueryString("finalidade") & " "

        strSQL &= "ORDER BY NF.Empresa_Id"

        Return strSQL
    End Function

End Class