Imports System
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class EncargosDaNota
    Inherits BasePage

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Dim strHTML As String = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf & _
                                "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf & _
                                "   <head>" & vbCrLf & _
                                "       <title>Encargos</title>" & vbCrLf & _
                                "       <link href=""menucentral.css"" type=""text/css"" rel=""stylesheet"" />" & _
                                "       <script language=""javascript"" src=""js/funcoes.js"" type=""text/javascript""></script>" & _
                                "       <style type=""text/css"">" & vbCrLf & _
                                "           body" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               font-size: 7pt;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .titulos" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               width: 100%;" & vbCrLf & _
                                "               border-top: solid 1px #000000;" & vbCrLf & _
                                "               border-bottom: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .titulosbaixo" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               width: 100%;" & vbCrLf & _
                                "               border-bottom: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .linhastitulobaixo" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               border-bottom: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .linhastitulocima" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               border-top: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .quebrapagina" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               page-break-after: always;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "       </style>" & vbCrLf & _
                                "   </head>" & vbCrLf & _
                                "   <body>"

        writer.Write(strHTML)

        If Request.QueryString("tipo") = "TNOTA" Then
            ImprimirTotalEncargos(writer)
        ElseIf Request.QueryString("tipo") = "OPNOTA" Then
            ImprimirTotalEncargosPorOperacao(writer)
        Else
            ImprimirEncargos(writer)
        End If

        strHTML = "   </body>" & _
                  "</html>"

        writer.Write(strHTML)

    End Sub

    Private Sub ImprimirEncargos(ByRef writer As HtmlTextWriter)
        Dim sql As String = ""
        Dim strHTML As String = "<table width=""100%"" class=""borda"">"

        Dim pNota As New [Lib].Negocio.NotaFiscal()
        pNota.CodigoEmpresa = Request.QueryString("emp")
        pNota.EnderecoEmpresa = Request.QueryString("endemp")
        pNota.CodigoCliente = Request.QueryString("cli")
        pNota.EnderecoCliente = Request.QueryString("endcli")
        If Request.QueryString("entsai") = "E" Then
            pNota.EntradaSaida = eEntradaSaida.Entrada
        Else
            pNota.EntradaSaida = eEntradaSaida.Saida
        End If
        pNota.Serie = Request.QueryString("serie")
        pNota.Codigo = Request.QueryString("nota")

        Dim NotaFiscal As New [Lib].Negocio.NotaFiscal(pNota)

        strHTML &= "<tr>" & _
                  "<td align=""right"">" & _
                  "<b>EMPRESA:</b>" & _
                  "</td>" & _
                  "<td>" & NotaFiscal.Empresa.Nome & _
                  "</td>" & _
                  "</tr>" & _
                  "<tr>" & _
                  "<td align=""right"">" & _
                  "<b>CIDADE/UF:</b>" & _
                  "</td>" & _
                  "<td>" & NotaFiscal.Empresa.Cidade & "/" & NotaFiscal.Empresa.CodigoEstado & _
                  "</td>" & _
                  "</tr>" & _
                  "<tr>" & _
                  "<td align=""right"">" & _
                  "<b>CLIENTE:</b>" & _
                  "</td>" & _
                  "<td>" & NotaFiscal.Cliente.Nome & _
                  "</td>" & _
                  "</tr>" & _
                  "<tr>" & _
                  "<td align=""right"">" & _
                  "<b>NOTA FISCAL:</b>" & _
                  "</td>" & _
                  "<td>" & NotaFiscal.Serie & "-" & NotaFiscal.Codigo & "-" & Request.QueryString("entsai") & _
                  "</td>" & _
                  "</tr>"

        For Each item As [Lib].Negocio.NotaFiscalXItem In NotaFiscal.Itens
            If item.CodigoProduto = Request.QueryString("produto") Then
                strHTML &= "<tr>" & _
                          "<td align=""right"">" & _
                          "<b>PRODUTO:</b>" & _
                          "</td>" & _
                          "<td>" & item.CodigoProduto & " - " & item.Produto.Descricao & _
                          "</td>" & _
                          "</tr>" & _
                          "</table>" & _
                          "<hr>" & _
                          "<table width=""100%"" class=""borda"">" & _
                          "<tr>" & _
                          "<td class=""linhastitulobaixo"">" & _
                          "<b>ENCARGO</b>" & _
                          "</td>" & _
                          "<td class=""linhastitulobaixo"">" & _
                          "<b>BASE</b>" & _
                          "</td>" & _
                          "<td class=""linhastitulobaixo"">" & _
                          "<b>%</b>" & _
                          "</td>" & _
                          "<td class=""linhastitulobaixo"" align=""right"">" & _
                          "<b>VALOR</b>" & _
                          "</td>" & _
                          "</tr>"
                For Each encargo As [Lib].Negocio.NotaFiscalXItemXEncargo In item.Encargos
                    strHTML &= "<tr>" & _
                              "<td>" & encargo.Codigo & _
                              "</td>" & _
                              "<td>" & encargo.Base & _
                              "</td>" & _
                              "<td>" & encargo.Percentual & _
                              "</td>" & _
                              "<td align=""right"">" & encargo.Valor & "&nbsp;" & encargo.Sinal & _
                              "</td>" & _
                              "<tr>"
                Next
            End If
        Next

        strHTML &= "</table>" & _
                   "<hr>" & _
                   "<table width=""100%"" class=""borda"">" & _
                   "<tr>" & _
                   "<td align=""center""><img border=""0"" id=""img1"" onclick=""imprimePagina();"" style=""cursor: pointer"" src=""images/button11.jpg"" height=""20"" width=""100"" alt=""Imprimir"" onmouseover=FP_swapImg(1,0,/*id*/""img1"",/*url*/""images/button9.jpg"") onmouseout=FP_swapImg(0,0,/*id*/""img1"",/*url*/""images/button11.jpg"") onmousedown=FP_swapImg(1,0,/*id*/""img1"",/*url*/""images/button10.jpg"") onmouseup=FP_swapImg(0,0,/*id*/""img1"",/*url*/""images/button9.jpg"") title=""Imprimir"">" & _
                   "</td>" & _
                   "<td align=""center""><img border=""0"" id=""img2"" onclick=""window.close();"" style=""cursor: pointer"" src=""images/button3.jpg"" height=""20"" width=""100"" alt=""Fechar"" onmouseover=FP_swapImg(1,0,/*id*/""img2"",/*url*/""images/button7.jpg"") onmouseout=FP_swapImg(0,0,/*id*/""img2"",/*url*/""images/button3.jpg"") onmousedown=FP_swapImg(1,0,/*id*/""img2"",/*url*/""images/button8.jpg"") onmouseup=FP_swapImg(0,0,/*id*/""img2"",/*url*/""images/button7.jpg"") title=""Fechar"">" & _
                   "</td>" & _
                   "</tr>" & _
                   "</table>"

        writer.Write(strHTML)
    End Sub

    Private Sub ImprimirTotalEncargos(ByRef writer As HtmlTextWriter)
        Dim strHTML As String = "<table width=""100%"" class=""borda"">"

        Dim objPedido As New [Lib].Negocio.Pedido(Request.QueryString("emp"), Request.QueryString("endemp"), Request.QueryString("pedido"))

        Dim sql As String = "SELECT     NFXE.Encargo_Id, " & vbCrLf & _
                            "           NFXE.Percentual, " & vbCrLf & _
                            "           case when NFXE.Encargo_Id = 'PRODUTO' then 1 when NFXE.Encargo_Id = 'LIQUIDO' then 3 else 2 end as Ordem, " & vbCrLf & _
                            "           sum(case when OP.Devolucao = 'S' then NFXE.Base * -1 else NFXE.Base end) as Base, " & vbCrLf & _
                            "           sum(case when OP.Devolucao = 'S' then NFXE.Valor * -1 else NFXE.Valor end) as Valor " & vbCrLf & _
                            "FROM         NotasFiscaisXEncargos AS NFXE " & vbCrLf & _
                            "           INNER JOIN NotasFiscais AS NF " & vbCrLf & _
                            "                   ON NFXE.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                            "                  AND NFXE.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                            "                  AND NFXE.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                            "                  AND NFXE.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                            "                  AND NFXE.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                            "                  AND NFXE.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                            "                  AND NFXE.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                            "           INNER JOIN SubOperacoes AS OP " & vbCrLf & _
                            "                   ON OP.Operacao_Id     = NF.Operacao " & vbCrLf & _
                            "                  AND OP.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                            "WHERE (NF.Empresa_Id = '" & Request.QueryString("emp") & "') " & vbCrLf & _
                            "  AND (NF.EndEmpresa_Id = " & Request.QueryString("endemp") & ") " & vbCrLf & _
                            "  AND (NF.Pedido = " & objPedido.Codigo & ") " & vbCrLf & _
                            "  AND (NF.Situacao = 1) " & vbCrLf & _
                            "GROUP BY NFXE.Encargo_Id, NFXE.Percentual " & vbCrLf & _
                            "Order By Ordem"

        Dim dsNotasXEncargos As DataSet = Banco.ConsultaDataSet(sql, "NotasFiscais")

        strHTML &= "<tr>" & _
          "<td align=""right"">" & _
          "<b>EMPRESA:</b>" & _
          "</td>" & _
          "<td>" & objPedido.Empresa.Nome & _
          "</td>" & _
          "</tr>" & _
          "<tr>" & _
          "<td align=""right"">" & _
          "<b>CIDADE/UF:</b>" & _
          "</td>" & _
          "<td>" & objPedido.Empresa.Cidade & "/" & objPedido.Empresa.CodigoEstado & _
          "</td>" & _
          "</tr>" & _
          "<tr>" & _
          "<td align=""right"">" & _
          "<b>CLIENTE:</b>" & _
          "</td>" & _
          "<td>" & objPedido.Cliente.Nome & _
          "</td>" & _
          "</tr>" & _
          "</table>" & _
          "<hr>" & _
          "<table width=""100%"" class=""borda"">" & _
          "<tr>" & _
          "<td class=""linhastitulobaixo"">" & _
          "<b>ENCARGO</b>" & _
          "</td>" & _
          "<td class=""linhastitulobaixo"">" & _
          "<b>BASE</b>" & _
          "</td>" & _
          "<td class=""linhastitulobaixo"">" & _
          "<b>%</b>" & _
          "</td>" & _
          "<td class=""linhastitulobaixo"" align=""right"">" & _
          "<b>VALOR</b>" & _
          "</td>" & _
          "</tr>"

        For Each drEncargo As DataRow In dsNotasXEncargos.Tables(0).Rows
            strHTML &= "<tr>" & _
                      "<td>" & drEncargo("Encargo_Id") & _
                      "</td>" & _
                      "<td>" & drEncargo("Base") & _
                      "</td>" & _
                      "<td>" & drEncargo("Percentual") & _
                      "</td>" & _
                      "<td align=""right"">" & drEncargo("Valor") & _
                      "</td>" & _
                      "<tr>"
        Next

        strHTML &= "</table>" & _
                   "<hr>" & _
                   "<table width=""100%"" class=""borda"">" & _
                   "<tr>" & _
                   "<td align=""center""><img border=""0"" id=""img1"" onclick=""imprimePagina();"" style=""cursor: pointer"" src=""images/button11.jpg"" height=""20"" width=""100"" alt=""Imprimir"" onmouseover=FP_swapImg(1,0,/*id*/""img1"",/*url*/""images/button9.jpg"") onmouseout=FP_swapImg(0,0,/*id*/""img1"",/*url*/""images/button11.jpg"") onmousedown=FP_swapImg(1,0,/*id*/""img1"",/*url*/""images/button10.jpg"") onmouseup=FP_swapImg(0,0,/*id*/""img1"",/*url*/""images/button9.jpg"") title=""Imprimir"">" & _
                   "</td>" & _
                   "<td align=""center""><img border=""0"" id=""img2"" onclick=""window.close();"" style=""cursor: pointer"" src=""images/button3.jpg"" height=""20"" width=""100"" alt=""Fechar"" onmouseover=FP_swapImg(1,0,/*id*/""img2"",/*url*/""images/button7.jpg"") onmouseout=FP_swapImg(0,0,/*id*/""img2"",/*url*/""images/button3.jpg"") onmousedown=FP_swapImg(1,0,/*id*/""img2"",/*url*/""images/button8.jpg"") onmouseup=FP_swapImg(0,0,/*id*/""img2"",/*url*/""images/button7.jpg"") title=""Fechar"">" & _
                   "</td>" & _
                   "</tr>" & _
                   "</table>"

        writer.Write(strHTML)
    End Sub

    Private Sub ImprimirTotalEncargosPorOperacao(ByRef writer As HtmlTextWriter)
        Dim strHTML As String = "<table width=""100%"" class=""borda"">"

        Dim objPedido As New [Lib].Negocio.Pedido(Request.QueryString("emp"), Request.QueryString("endemp"), Request.QueryString("pedido"))
        Dim operacao As New [Lib].Negocio.SubOperacao(Request.QueryString("ope"), Request.QueryString("sop"))

        Dim sql As String = "SELECT     NFXE.Encargo_Id, " & vbCrLf & _
                            "           NFXE.Percentual, " & vbCrLf & _
                            "           case when NFXE.Encargo_Id = 'PRODUTO' then 1 when NFXE.Encargo_Id = 'LIQUIDO' then 3 else 2 end as Ordem, " & vbCrLf & _
                            "           sum(NFXE.Base) as Base, " & vbCrLf & _
                            "           sum(NFXE.Valor) as Valor " & vbCrLf & _
                            "FROM         NotasFiscaisXEncargos AS NFXE " & vbCrLf & _
                            "           INNER JOIN NotasFiscais AS NF " & vbCrLf & _
                            "                   ON NFXE.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                            "                  AND NFXE.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                            "                  AND NFXE.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                            "                  AND NFXE.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                            "                  AND NFXE.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                            "                  AND NFXE.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                            "                  AND NFXE.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                            "           INNER JOIN SubOperacoes AS OP " & vbCrLf & _
                            "                   ON OP.Operacao_Id     = NF.Operacao " & vbCrLf & _
                            "                  AND OP.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                            "WHERE (NF.Empresa_Id = '" & Request.QueryString("emp") & "') " & vbCrLf & _
                            "  AND (NF.EndEmpresa_Id = " & Request.QueryString("endemp") & ") " & vbCrLf & _
                            "  AND (NF.Pedido      = " & objPedido.Codigo & ") " & vbCrLf & _
                            "  AND (NF.Operacao    = " & operacao.CodigoOperacao & ") " & vbCrLf & _
                            "  AND (NF.SubOperacao = " & operacao.Codigo & ") " & vbCrLf & _
                            "  AND (NF.Situacao = 1) " & vbCrLf & _
                            "GROUP BY NFXE.Encargo_Id, NFXE.Percentual " & vbCrLf & _
                            "Order By Ordem"

        Dim dsNotasXEncargos As DataSet = Banco.ConsultaDataSet(sql, "NotasFiscais")

        strHTML &= "<tr>" & _
          "<td align=""right"">" & _
          "<b>EMPRESA:</b>" & _
          "</td>" & _
          "<td>" & objPedido.Empresa.Nome & _
          "</td>" & _
          "</tr>" & _
          "<tr>" & _
          "<td align=""right"">" & _
          "<b>CIDADE/UF:</b>" & _
          "</td>" & _
          "<td>" & objPedido.Empresa.Cidade & "/" & objPedido.Empresa.CodigoEstado & _
          "</td>" & _
          "</tr>" & _
          "<tr>" & _
          "<td align=""right"">" & _
          "<b>CLIENTE:</b>" & _
          "</td>" & _
          "<td>" & objPedido.Cliente.Nome & _
          "</td>" & _
          "</tr>" & _
          "<tr>" & _
          "<td align=""right"">" & _
          "<b>OPERAÇÃO:</b>" & _
          "</td>" & _
          "<td>" & operacao.CodigoOperacao.ToString("00") & "-" & operacao.Codigo.ToString("00") & " - " & operacao.Descricao & _
          "</td>" & _
          "</tr>" & _
          "</table>" & _
          "<hr>" & _
          "<table width=""100%"" class=""borda"">" & _
          "<tr>" & _
          "<td class=""linhastitulobaixo"">" & _
          "<b>ENCARGO</b>" & _
          "</td>" & _
          "<td class=""linhastitulobaixo"">" & _
          "<b>BASE</b>" & _
          "</td>" & _
          "<td class=""linhastitulobaixo"">" & _
          "<b>%</b>" & _
          "</td>" & _
          "<td class=""linhastitulobaixo"" align=""right"">" & _
          "<b>VALOR</b>" & _
          "</td>" & _
          "</tr>"

        For Each drEncargo As DataRow In dsNotasXEncargos.Tables(0).Rows
            strHTML &= "<tr>" & _
                      "<td>" & drEncargo("Encargo_Id") & _
                      "</td>" & _
                      "<td>" & drEncargo("Base") & _
                      "</td>" & _
                      "<td>" & drEncargo("Percentual") & _
                      "</td>" & _
                      "<td align=""right"">" & drEncargo("Valor") & _
                      "</td>" & _
                      "<tr>"
        Next

        strHTML &= "</table>" & _
                   "<hr>" & _
                   "<table width=""100%"" class=""borda"">" & _
                   "<tr>" & _
                   "<td align=""center""><img border=""0"" id=""img1"" onclick=""imprimePagina();"" style=""cursor: pointer"" src=""images/button11.jpg"" height=""20"" width=""100"" alt=""Imprimir"" onmouseover=FP_swapImg(1,0,/*id*/""img1"",/*url*/""images/button9.jpg"") onmouseout=FP_swapImg(0,0,/*id*/""img1"",/*url*/""images/button11.jpg"") onmousedown=FP_swapImg(1,0,/*id*/""img1"",/*url*/""images/button10.jpg"") onmouseup=FP_swapImg(0,0,/*id*/""img1"",/*url*/""images/button9.jpg"") title=""Imprimir"">" & _
                   "</td>" & _
                   "<td align=""center""><img border=""0"" id=""img2"" onclick=""window.close();"" style=""cursor: pointer"" src=""images/button3.jpg"" height=""20"" width=""100"" alt=""Fechar"" onmouseover=FP_swapImg(1,0,/*id*/""img2"",/*url*/""images/button7.jpg"") onmouseout=FP_swapImg(0,0,/*id*/""img2"",/*url*/""images/button3.jpg"") onmousedown=FP_swapImg(1,0,/*id*/""img2"",/*url*/""images/button8.jpg"") onmouseup=FP_swapImg(0,0,/*id*/""img2"",/*url*/""images/button7.jpg"") title=""Fechar"">" & _
                   "</td>" & _
                   "</tr>" & _
                   "</table>"

        writer.Write(strHTML)
    End Sub

End Class