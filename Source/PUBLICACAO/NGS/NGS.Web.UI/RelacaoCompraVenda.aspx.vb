Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelacaoCompraVenda
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RELACAOCOMPRAVENDA", "ACESSAR") Then
                limpar()
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Function Validar() As Boolean
        'If Not ddlEmpresa.SelectedIndex > 0 Then
        '    MsgBox(Me.Page, "Informe a empresa")
        '    Return False
        'End If

        If Not IsDate(txtDataDe.Text) Then
            MsgBox(Me.Page, "Informe a Data do Relatorio")
            Return False
        End If

        If Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Selecione o produto")
            Return False
        End If

        Return True
    End Function

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelacaoCompraVenda", "RELATORIO") Then
                If Not Validar() Then Exit Sub
                Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")

                Dim SelProdutoPedido As ArrayList
                SelProdutoPedido = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.grupo", "Prd.Produto_Id")
                'Dim SelProduto As ArrayList
                'SelProduto = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.grupo", "PedidosXItens.Produto_Id")

                Dim str As String
                str = " Select sb.*, MinVenc.Vencimento " & vbCrLf & _
                      "   Into #Temp " & vbCrLf & _
                      "   from (" & vbCrLf & _
                      "         SELECT Pedidos.Empresa_id," & vbCrLf & _
                      "                Pedidos.EndEmpresa_id," & vbCrLf & _
                      "                Empresa.Fantasia + ' ' + Empresa.complemento + ' ' + Empresa.Cidade + ' ' + Empresa.Estado as NomeEmpresa," & vbCrLf & _
                      "                SO.Classe, " & vbCrLf & _
                      "                0 AS Fixacao, " & vbCrLf & _
                      "                case pxi.TipoDeLancamento" & vbCrLf & _
                      "                   when 'N' then 'Normal'" & vbCrLf & _
                      "                   when 'C' then 'Complemento'" & vbCrLf & _
                      "                   when 'E' then 'Estorno'" & vbCrLf & _
                      "                end as TipoDeLancamento," & vbCrLf & _
                      "                Pedidos.DataEntrega," & vbCrLf & _
                      "                pxi.Movimento," & vbCrLf & _
                      "                Pedidos.Safra," & vbCrLf & _
                      "                Pedidos.Moeda," & vbCrLf & _
                      "                M.Descricao as DescMoeda," & vbCrLf & _
                      "                Pedidos.Operacao, " & vbCrLf & _
                      "                Pedidos.subOperacao," & vbCrLf & _
                      "                SO.Descricao, " & vbCrLf & _
                      "                pxi.Produto_Id as Produto," & vbCrLf & _
                      "                Prd.nome as NomeProduto, " & vbCrLf & _
                      "                Pedidos.pedido_id as Pedido," & vbCrLf & _
                      "                Clientes.Cliente_Id as Cliente, " & vbCrLf & _
                      "                Clientes.Endereco_Id as Endereco, " & vbCrLf & _
                      "                Clientes.Nome, " & vbCrLf & _
                      "                Clientes.Cidade, " & vbCrLf & _
                      "                Clientes.Estado, " & vbCrLf & _
                      "                case " & vbCrLf & _
                      "                  when pxi.TipoDeLancamento = 'E'" & vbCrLf & _
                      "                    then pxi.Quantidade * -1" & vbCrLf & _
                      "                    Else pxi.Quantidade" & vbCrLf & _
                      "                end Quantidade," & vbCrLf & _
                      "                case" & vbCrLf & _
                      "                   when M.Classificacao = 'O'" & vbCrLf & _
                      "                     then pxi.UnitarioOficial" & vbCrLf & _
                      "                     else pxi.UnitarioMoeda" & vbCrLf & _
                      "                end Unitario," & vbCrLf & _
                      "                case" & vbCrLf & _
                      "                   when M.Classificacao = 'O'" & vbCrLf & _
                      "                     then pxi.UnitarioOficial * 60" & vbCrLf & _
                      "                     else pxi.UnitarioMoeda * 60" & vbCrLf & _
                      "                end UnitarioSaco," & vbCrLf & _
                      "                case" & vbCrLf & _
                      "                   when M.Classificacao = 'O'" & vbCrLf & _
                      "                     then pxi.UnitarioOficial * 1000" & vbCrLf & _
                      "                     else pxi.UnitarioMoeda * 1000" & vbCrLf & _
                      "                end UnitarioTon," & vbCrLf & _
                      "                case" & vbCrLf & _
                      "                   when M.Classificacao = 'O' and pxi.TipoDeLancamento = 'E' " & vbCrLf & _
                      "                     then pxi.TotalOficial * - 1" & vbCrLf & _
                      "                   when M.Classificacao = 'O' and pxi.TipoDeLancamento <> 'E' " & vbCrLf & _
                      "                     then pxi.TotalOficial" & vbCrLf & _
                      "                   when M.Classificacao <> 'O' and pxi.TipoDeLancamento = 'E' " & vbCrLf & _
                      "                     then pxi.TotalMoeda * - 1" & vbCrLf & _
                      "                   when M.Classificacao <> 'O' and pxi.TipoDeLancamento <> 'E' " & vbCrLf & _
                      "                     then pxi.TotalMoeda" & vbCrLf & _
                      "                end Total," & vbCrLf & _
                      "                Pedidos.freteCifFob," & vbCrLf & _
                      "				   case" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'E' and Pedidos.freteCifFob = 'FOB' and len(isnull(Pedidos.LocalEmbarque,'')) = 0" & vbCrLf & _
                      "					    then Clientes.Cidade" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'E' and Pedidos.freteCifFob = 'FOB' and len(isnull(Pedidos.LocalEmbarque,'')) > 0" & vbCrLf & _
                      "					    then LEmb.Cidade" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'S' and Pedidos.freteCifFob = 'CIF' " & vbCrLf & _
                      "					    then isnull(Pedidos.CidadeEntrega,'')" & vbCrLf & _
                      "				      Else ''" & vbCrLf & _
                      "				   End as CidadeEntrega," & vbCrLf & _
                      "				   case" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'E' and Pedidos.freteCifFob = 'FOB' and len(isnull(Pedidos.LocalEmbarque,'')) = 0" & vbCrLf & _
                      "					    then Clientes.Estado" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'E' and Pedidos.freteCifFob = 'FOB' and len(isnull(Pedidos.LocalEmbarque,'')) > 0" & vbCrLf & _
                      "					    then LEmb.Estado" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'S' and Pedidos.freteCifFob = 'CIF' " & vbCrLf & _
                      "					    then isnull(Pedidos.EstadoEntrega,'')" & vbCrLf & _
                      "				      Else ''" & vbCrLf & _
                      "				   End as EstadoEntrega," & vbCrLf & _
                      "				   case" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'E' and Pedidos.freteCifFob = 'FOB' and len(isnull(Pedidos.LocalEmbarque,'')) = 0" & vbCrLf & _
                      "					    then Clientes.Complemento" & vbCrLf & _
                      "				      When SO.EntradaSaida = 'E' and Pedidos.freteCifFob = 'FOB' and len(isnull(Pedidos.LocalEmbarque,'')) > 0" & vbCrLf & _
                      "					    then LEmb.Complemento" & vbCrLf & _
                      "					    Else ''" & vbCrLf & _
                      "				   End as Complemento" & vbCrLf & _
                      "           FROM Pedidos " & vbCrLf & _
                      "          inner JOIN PedidoXItemxLancamento pxi " & vbCrLf & _
                      "             on Pedidos.Empresa_Id    = pxi.Empresa_Id " & vbCrLf & _
                      "            and Pedidos.EndEmpresa_Id = pxi.EndEmpresa_Id " & vbCrLf & _
                      "            and Pedidos.Pedido_Id     = pxi.Pedido_Id " & vbCrLf & _
                      "          inner JOIN Clientes Empresa " & vbCrLf & _
                      "             on Pedidos.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf & _
                      "            and Pedidos.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
                      "          inner JOIN Clientes " & vbCrLf & _
                      "             on Pedidos.Cliente    = Clientes.Cliente_Id " & vbCrLf & _
                      "            and Pedidos.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                      "          Inner Join Operacoes OP" & vbCrLf & _
                      "             on OP.Operacao_Id = Pedidos.Operacao" & vbCrLf & _
                      "          Inner Join SubOperacoes SO" & vbCrLf & _
                      "             on SO.Operacao_id     = Pedidos.Operacao" & vbCrLf & _
                      "            and SO.SubOperacoes_Id = Pedidos.subOperacao" & vbCrLf & _
                      "          Inner Join Moedas M" & vbCrLf & _
                      "             on M.Moeda_id = Pedidos.Moeda" & vbCrLf & _
                      "          Inner Join Produtos Prd" & vbCrLf & _
                      "             on Prd.Produto_Id = pxi.Produto_Id" & vbCrLf & _
                      "		      left join Clientes LEmb" & vbCrLf & _
                      "			    on Pedidos.LocalEmbarque    = LEmb.Cliente_id" & vbCrLf & _
                      "			   and Pedidos.EndLocalEmbarque = LEmb.Endereco_id" & vbCrLf & _
                      "          where Pedidos.situacao = 1" & vbCrLf & _
                      "            and " & SelProdutoPedido(0) & vbCrLf

                If chkMovimento.Checked Then
                    str &= "            and pxi.Movimento between convert(datetime,'" & CDate(txtDataDe.Text).ToString("yyyy-MM-dd") & "',101) and convert(datetime,'" & CDate(txtDataAte.Text).ToString("yyyy-MM-dd") & "',101)" & vbCrLf
                End If

                If ddlSafra.SelectedIndex > 0 Then
                    str &= "            and Pedidos.Safra ='" & ddlSafra.SelectedValue & "'"
                End If

                If ddlEmpresa.SelectedIndex > 0 Then
                    If chkConsolidarEmpresa.Checked Then
                        str &= "            and left(Pedidos.Empresa_id,8) ='" & Left(Emp(0), 8) & "'" & vbCrLf
                    Else
                        str &= "            and Pedidos.Empresa_id    ='" & Emp(0) & "'" & vbCrLf & _
                               "            and Pedidos.EndEmpresa_id = " & Emp(1)
                    End If
                End If

                If Not chkNormal.Checked Or Not chkEstorno.Checked Or Not chkComplemento.Checked Then
                    Dim TL As String = String.Empty

                    If chkNormal.Checked Then
                        TL = "'N'"
                    End If

                    If chkEstorno.Checked Then
                        TL &= IIf(TL.Length > 0, ",'E'", "'E'")
                    End If

                    If chkComplemento.Checked Then
                        TL &= IIf(TL.Length > 0, ",'C'", "'C'")
                    End If

                    str &= "            and pxi.TipoDeLancamento in (" & TL & ")"
                End If

                str &= "            and OP.Classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "')" & vbCrLf

                If (chkPrecoFixo.Checked And Not chkPrecoAFixar.Checked) Or (Not chkPrecoFixo.Checked And chkPrecoAFixar.Checked) Then
                    str &= "            and SO.PrecoFixo ='" & IIf(chkPrecoFixo.Checked, "S", "N") & "'" & vbCrLf
                End If

                If chkFixacao.Checked Then
                    str &= "  union " & vbCrLf & _
                           " Select P.Empresa_Id," & vbCrLf & _
                           "        P.EndEmpresa_id," & vbCrLf & _
                           "        Empresa.Fantasia + ' ' + Empresa.complemento + ' ' + Empresa.Cidade + ' ' + Empresa.Estado as NomeEmpresa," & vbCrLf & _
                           "        SOP.Classe, " & vbCrLf & _
                           "        PIF.Fixacao_Id, " & vbCrLf & _
                           "        'Fixacao' as TL," & vbCrLf & _
                           "        P.DataEntrega," & vbCrLf & _
                           "        PIF.Movimento, " & vbCrLf & _
                           "        P.Safra," & vbCrLf & _
                           "        P.Moeda, " & vbCrLf & _
                           "        M.Descricao, " & vbCrLf & _
                           "        PIF.Operacao," & vbCrLf & _
                           "        PIF.SubOperacao," & vbCrLf & _
                           "        SO.Descricao," & vbCrLf & _
                           "        PIF.Produto_Id," & vbCrLf & _
                           "        Prd.Nome," & vbCrLf & _
                           "        PIF.Pedido_id, " & vbCrLf & _
                           "        Clientes.Cliente_Id, " & vbCrLf & _
                           "        Clientes.Endereco_Id, " & vbCrLf & _
                           "        Clientes.Nome, " & vbCrLf & _
                           "        Clientes.Cidade, " & vbCrLf & _
                           "        Clientes.Estado, " & vbCrLf & _
                           "        PIF.Quantidade," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           when M.Classificacao = 'O'" & vbCrLf & _
                           "             then PIF.UnitarioOficial" & vbCrLf & _
                           "             else PIF.UnitarioMoeda" & vbCrLf & _
                           "        end Unitario," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           when M.Classificacao = 'O'" & vbCrLf & _
                           "             then PIF.UnitarioOficial * 60" & vbCrLf & _
                           "             else PIF.UnitarioMoeda * 60" & vbCrLf & _
                           "        end UnitarioSaco," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           when M.Classificacao = 'O'" & vbCrLf & _
                           "             then PIF.UnitarioOficial * 1000" & vbCrLf & _
                           "             else PIF.UnitarioMoeda * 1000" & vbCrLf & _
                           "        end UnitarioTon," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           when M.Classificacao = 'O'" & vbCrLf & _
                           "             then PIF.TotalOficial" & vbCrLf & _
                           "             else PIF.TotalMoeda" & vbCrLf & _
                           "        end Total," & vbCrLf & _
                           "        P.freteCifFob," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           When SO.EntradaSaida = 'E' and P.freteCifFob = 'FOB' and len(isnull(P.LocalEmbarque,'')) = 0" & vbCrLf & _
                           "             then Clientes.Cidade" & vbCrLf & _
                           "           When SO.EntradaSaida = 'E' and P.freteCifFob = 'FOB' and len(isnull(P.LocalEmbarque,'')) > 0" & vbCrLf & _
                           "             then LEmb.Cidade" & vbCrLf & _
                           "           When SO.EntradaSaida = 'S' and P.freteCifFob = 'CIF'" & vbCrLf & _
                           "             then isnull(P.CidadeEntrega,'')" & vbCrLf & _
                           "           Else ''" & vbCrLf & _
                           "        End as CidadeEntrega," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           When SO.EntradaSaida = 'E' and P.freteCifFob = 'FOB' and len(isnull(P.LocalEmbarque,'')) = 0" & vbCrLf & _
                           "             then Clientes.Estado" & vbCrLf & _
                           "           When SO.EntradaSaida = 'E' and P.freteCifFob = 'FOB' and len(isnull(P.LocalEmbarque,'')) > 0" & vbCrLf & _
                           "             then LEmb.Estado" & vbCrLf & _
                           "           When SO.EntradaSaida = 'S' and P.freteCifFob = 'CIF'" & vbCrLf & _
                           "             then isnull(P.EstadoEntrega,'')" & vbCrLf & _
                           "           Else ''" & vbCrLf & _
                           "        End as EstadoEntrega," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "           When SO.EntradaSaida = 'E' and P.freteCifFob = 'FOB' and len(isnull(P.LocalEmbarque,'')) = 0" & vbCrLf & _
                           "             then Clientes.Complemento" & vbCrLf & _
                           "           When SO.EntradaSaida = 'E' and P.freteCifFob = 'FOB' and len(isnull(P.LocalEmbarque,'')) > 0" & vbCrLf & _
                           "             then LEmb.Complemento" & vbCrLf & _
                           "             Else ''" & vbCrLf & _
                           "        End as Complemento" & vbCrLf & _
                           "   from VW_PedidosXItensXFixacoes PIF" & vbCrLf & _
                           "  Inner Join Pedidos P" & vbCrLf & _
                           "     on PIF.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
                           "    and PIF.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                           "    and PIF.Pedido_id     = P.Pedido_Id" & vbCrLf & _
                           "  INNER JOIN Clientes " & vbCrLf & _
                           "     ON P.Cliente    = Clientes.Cliente_Id " & vbCrLf & _
                           "    AND P.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                           "  inner JOIN Clientes Empresa " & vbCrLf & _
                           "     on P.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf & _
                           "    and P.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
                           "  Inner Join SubOperacoes SO" & vbCrLf & _
                           "     on SO.Operacao_id     = PIF.Operacao" & vbCrLf & _
                           "    and SO.SubOperacoes_Id = PIF.subOperacao" & vbCrLf & _
                           "  Inner Join SubOperacoes SOP" & vbCrLf & _
                           "     on SOP.Operacao_id     = P.Operacao" & vbCrLf & _
                           "    and SOP.SubOperacoes_Id = P.subOperacao" & vbCrLf & _
                           "  Inner Join Moedas M" & vbCrLf & _
                           "     on M.Moeda_id = P.Moeda" & vbCrLf & _
                           "  Inner Join Produtos Prd" & vbCrLf & _
                           "     on Prd.Produto_Id = PIF.Produto_Id" & vbCrLf & _
                           "   left join Clientes LEmb" & vbCrLf & _
                           "     on P.LocalEmbarque    = LEmb.Cliente_id" & vbCrLf & _
                           "    and P.EndLocalEmbarque = LEmb.Endereco_id" & vbCrLf & _
                           "  where P.situacao = 1" & vbCrLf & _
                           "    And " & SelProdutoPedido(0) & vbCrLf & _
                           "    And SO.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf

                    If ddlSafra.SelectedIndex > 0 Then
                        str &= "    And P.Safra ='" & ddlSafra.SelectedValue & "'"
                    End If

                    If chkMovimento.Checked Then
                        str &= "    And PIF.Movimento between convert(datetime,'" & CDate(txtDataDe.Text).ToString("yyyy-MM-dd") & "',101) and convert(datetime,'" & CDate(txtDataAte.Text).ToString("yyyy-MM-dd") & "',101)" & vbCrLf
                    End If

                    If ddlEmpresa.SelectedIndex > 0 Then
                        If chkConsolidarEmpresa.Checked Then
                            str &= "  and left(P.Empresa_id,8) ='" & Left(Emp(0), 8) & "'" & vbCrLf
                        Else
                            str &= "  and P.Empresa_id    ='" & Emp(0) & "'" & vbCrLf & _
                                   "  and P.EndEmpresa_id = " & Emp(1)
                        End If
                    End If
                End If

                str &= "  )sb" & vbCrLf & _
                       " Left Join (Select Pedido, Movimento, Venc.PedidoFixacao, " & vbCrLf
                If (FinanceiroNovo) Then
                    str &=
                      "                   Min(Reprogramacao) as Vencimento" & vbCrLf & _
                      "              from (" & vbCrLf & _
                      "                     Select Pedido, Reprogramacao, Movimento from Titulos" & vbCrLf & _
                      "                    ) Venc" & vbCrLf
                Else
                    str &=
                      "                   Min(Prorrogacao) as Vencimento" & vbCrLf & _
                      "              from (" & vbCrLf & _
                      "                     SELECT Pedido, Prorrogacao, Movimento, ISNULL(PedidoFixacao,0) PedidoFixacao " & vbCrLf & _
                      "                       FROM ContasApagar" & vbCrLf & _
                      "                      WHERE Provisao <> 1" & vbCrLf & _
                      "                      UNION " & vbCrLf & _
                      "                     SELECT Pedido, Prorrogacao, Movimento, ISNULL(PedidoFixacao,0) PedidoFixacao " & vbCrLf & _
                      "                       FROM ContasaReceber" & vbCrLf & _
                      "                      WHERE Provisao <> 1" & vbCrLf & _
                      "                    ) Venc" & vbCrLf
                End If

                str &= "              group by Pedido, Movimento, Venc.PedidoFixacao" & vbCrLf & _
                       "            ) MinVenc" & vbCrLf & _
                       "    on Sb.Pedido = MinVenc.Pedido" & vbCrLf & _
                       "   AND Sb.Movimento = MinVenc.Movimento  " & vbCrLf & _
                       "   AND Sb.Fixacao = MinVenc.PedidoFixacao " & vbCrLf & _
                       " Order by Movimento, Safra, Produto, Classe, Moeda, operacao, suboperacao;" & vbCrLf & _
                       " Select * from #Temp;" & vbCrLf & _
                       " Select T.Empresa_Id, T.EndEmpresa_Id, T.Movimento, T.Safra," & vbCrLf & _
                       "        T.Produto," & vbCrLf & _
                       "        T.NomeProduto," & vbCrLf & _
                       "        T.Classe," & vbCrLf & _
                       "        T.Moeda," & vbCrLf & _
                       "        T.DescMoeda," & vbCrLf & _
                       "        T.TipoDeLancamento," & vbCrLf & _
                       "        case when Sum(T.Quantidade) = 0 " & vbCrLf & _
                       "              then  0" & vbCrLf & _
                       "              else sum(T.Total) / Sum(T.Quantidade)" & vbCrLf & _
                       "        end as Unitario," & vbCrLf & _
                       "        case when Sum(T.Quantidade) = 0" & vbCrLf & _
                       "              then 0" & vbCrLf & _
                       "              else (sum(T.Total) / Sum(T.Quantidade)) * 60" & vbCrLf & _
                       "        end as UnitarioSC," & vbCrLf & _
                       "        case when Sum(T.Quantidade) = 0" & vbCrLf & _
                       "              then 0" & vbCrLf & _
                       "              else (sum(T.Total) / Sum(T.Quantidade))* 1000" & vbCrLf & _
                       "        end as UnitarioTON," & vbCrLf & _
                       "        sum(T.Quantidade) as Quantidade," & vbCrLf & _
                       "        sum(T.Total) as Total " & vbCrLf & _
                       "   from #Temp T" & vbCrLf & _
                       "  Inner Join Produtos Prd" & vbCrLf & _
                       "     on T.Produto = Prd.Produto_Id" & vbCrLf & _
                       "  Group by T.Safra, T.Empresa_Id, T.EndEmpresa_Id, T.Movimento,  T.Produto, T.NomeProduto, T.classe, T.Moeda, T.DescMoeda, T.TipoDeLancamento" & vbCrLf & _
                       "  order by T.Safra, T.Empresa_Id, T.EndEmpresa_Id, T.Movimento,T.Produto, T.Classe, T.Moeda," & vbCrLf & _
                       "           case T.TipoDeLancamento" & vbCrLf & _
                       "              when 'Normal' then 1" & vbCrLf & _
                       "              When 'Complemento' then 2" & vbCrLf & _
                       "              when 'Estorno' then 3" & vbCrLf & _
                       "              else 4" & vbCrLf & _
                       "          end;" & vbCrLf


                str &= "select T.Safra,T.Movimento," & vbCrLf & _
                       "       T.Produto," & vbCrLf & _
                       "       T.NomeProduto," & vbCrLf & _
                       "       T.Classe," & vbCrLf & _
                       "       T.Moeda," & vbCrLf & _
                       "       T.DescMoeda," & vbCrLf & _
                       "       T.TipoDeLancamento," & vbCrLf & _
                       "       case when Sum(T.Quantidade) = 0 " & vbCrLf & _
                       "             then  0" & vbCrLf & _
                       "             else sum(T.Total) / Sum(T.Quantidade)" & vbCrLf & _
                       "       end as Unitario," & vbCrLf & _
                       "       case when Sum(T.Quantidade) = 0" & vbCrLf & _
                       "             then 0" & vbCrLf & _
                       "             else (sum(T.Total) / Sum(T.Quantidade))* 60" & vbCrLf & _
                       "       end as UnitarioSC," & vbCrLf & _
                       "       case when Sum(T.Quantidade) = 0" & vbCrLf & _
                       "             then 0" & vbCrLf & _
                       "             else (sum(T.Total) / Sum(T.Quantidade))* 1000" & vbCrLf & _
                       "       end as UnitarioTON," & vbCrLf & _
                       "       sum(T.Quantidade) as Quantidade," & vbCrLf & _
                       "       sum(T.Total) as Total " & vbCrLf & _
                       "  from #Temp T" & vbCrLf & _
                       " Inner Join Produtos Prd" & vbCrLf & _
                       "    on T.Produto = Prd.Produto_Id" & vbCrLf & _
                       " Group by T.Safra,T.Movimento, T.Produto, T.NomeProduto, T.classe, T.Moeda, T.DescMoeda, T.TipoDeLancamento" & vbCrLf & _
                       " order by T.Safra,T.Movimento,T.Produto, T.Classe, T.Moeda," & vbCrLf & _
                       "          case T.TipoDeLancamento" & vbCrLf & _
                       "             when 'Normal' then 1" & vbCrLf & _
                       "             When 'Complemento' then 2" & vbCrLf & _
                       "             when 'Estorno' then 3" & vbCrLf & _
                       "             else 4" & vbCrLf & _
                       "          end;" & vbCrLf


                str &= "select T.Safra,T.Produto," & vbCrLf & _
                       "       T.NomeProduto," & vbCrLf & _
                       "       T.classe," & vbCrLf & _
                       "       T.Moeda," & vbCrLf & _
                       "       T.DescMoeda," & vbCrLf & _
                       "       T.TipoDeLancamento," & vbCrLf & _
                       "       case when Sum(T.Quantidade) = 0 " & vbCrLf & _
                       "             then  0" & vbCrLf & _
                       "             else sum(T.Total) / Sum(T.Quantidade)" & vbCrLf & _
                       "       end as Unitario," & vbCrLf & _
                       "       case when Sum(T.Quantidade) = 0" & vbCrLf & _
                       "             then 0" & vbCrLf & _
                       "             else (sum(T.Total) / Sum(T.Quantidade)) * 60" & vbCrLf & _
                       "       end as UnitarioSC," & vbCrLf & _
                       "       case when Sum(T.Quantidade) = 0" & vbCrLf & _
                       "             then 0" & vbCrLf & _
                       "             else (sum(T.Total) / Sum(T.Quantidade)) * 1000" & vbCrLf & _
                       "       end as UnitarioTON," & vbCrLf & _
                       "       sum(T.Quantidade) as Quantidade," & vbCrLf & _
                       "       sum(T.Total) as Total " & vbCrLf & _
                       "  from #Temp T" & vbCrLf & _
                       " Inner Join Produtos Prd" & vbCrLf & _
                       "    on T.Produto = Prd.Produto_Id" & vbCrLf & _
                       "  Group by T.Safra,T.Produto, T.NomeProduto, T.classe, T.Moeda, T.DescMoeda, T.TipoDeLancamento" & vbCrLf & _
                       " Order by T.Safra,T.Produto, T.Classe, T.Moeda, case T.TipoDeLancamento" & vbCrLf & _
                       "                                          when 'Normal' then 1" & vbCrLf & _
                       "                                          When 'Complemento' then 2" & vbCrLf & _
                       "                                          when 'Estorno' then 3" & vbCrLf & _
                       "                                          else 4" & vbCrLf & _
                       "                                       end" & vbCrLf

                Dim ds As DataSet
                ds = Banco.ConsultaDataSet(str, "Consulta")
                ds.Tables(0).TableName = "RelacaoCompraVenda"
                ds.Tables(1).TableName = "ResumoEmpresa"
                ds.Tables(2).TableName = "ResumoDia"
                ds.Tables(3).TableName = "ResumoRelatorio"

                Dim NomeArquivo2 As String = ""
                If Pdf = True Then
                    NomeArquivo2 = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Else
                    NomeArquivo2 = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
                End If
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                Dim DescEmpresa As String = ""
                Dim DescCidadeEstado As String = ""
                Dim pEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
                If ddlEmpresa.SelectedIndex > 0 Then
                    Dim Empresa As New [Lib].Negocio.Cliente(pEmpresa(0), pEmpresa(1))
                    DescEmpresa = Empresa.Nome & " - " & Funcoes.FormatarCpfCnpj(Empresa.Codigo)
                    DescCidadeEstado = Empresa.Cidade & " - " & Empresa.CodigoEstado
                End If

                Dim Parametros As String = ""
                If chkMovimento.Checked Then Parametros = "Relacao do dia de " & txtDataDe.Text & " ate " & txtDataAte.Text & vbCrLf
                If ddlSafra.SelectedIndex > 0 Then Parametros = "Safra: " & ddlSafra.SelectedValue

                If ddlEmpresa.SelectedIndex > 0 Then
                    Parametros &= "Empresa: '" & ddlEmpresa.SelectedItem.Text & vbCrLf
                End If

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Nome", DescEmpresa)
                parameters.Add("Cidade", DescCidadeEstado)
                parameters.Add("Parametros", Parametros & SelProdutoPedido(1))

                Funcoes.BindReport(Me.Page, ds, "Cr_RelacaoCompraVenda", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)

                'crparameterfielddefinition = crparameterfielddefinitions.Item("Estoque")
                'crparametervalues = crparameterfielddefinition.CurrentValues
                'crparameterdiscretevalue = New ParameterDiscreteValue
                'crparameterdiscretevalue.Value = 0
                'crparametervalues.Add(crparameterdiscretevalue)
                'crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If chkPrecoFixo.Checked Or chkPrecoAFixar.Checked Then
                    Parametros &= "Pedido com Preço: "
                    If chkPrecoFixo.Checked Then Parametros &= "Fixo"
                    If chkPrecoAFixar.Checked Then
                        Parametros &= IIf(chkPrecoFixo.Checked, " e ", "") & " A Fixar"
                    End If
                    Parametros &= vbCrLf
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

    Private Sub limpar()
        chkConsolidarEmpresa.Checked = False
        ddlEmpresa.SelectedIndex = 0
        chkPrecoFixo.Checked = True
        chkPrecoAFixar.Checked = False
        chkNormal.Checked = True
        chkComplemento.Checked = True
        chkEstorno.Checked = True
        chkFixacao.Checked = True
        ddlSafra.SelectedIndex = 0

        chkMovimento.Checked = True
        txtDataDe.Parent.Visible = chkMovimento.Checked

        txtDataDe.Text = Date.Now.AddDays(-1).ToString("dd/MM/yyyy")
        txtDataAte.Text = txtDataDe.Text

        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa, True)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelacaoCompraVenda")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkMovimento_CheckedChanged(sender As Object, e As EventArgs) Handles chkMovimento.CheckedChanged
        txtDataDe.Parent.Visible = chkMovimento.Checked
    End Sub
End Class