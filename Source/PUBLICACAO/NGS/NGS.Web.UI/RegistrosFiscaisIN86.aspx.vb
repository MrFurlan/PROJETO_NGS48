Imports System.IO
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RegistrosFiscaisIN86
    Inherits BasePage

    Dim sql As String
    Dim cl As Integer = 0
    Dim ds As New DataSet
    Dim dr As DataRow

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim PaginaInicial As Integer = 0

    Dim ClienteNome As String
    Dim ClienteInscricao As String
    Dim ClienteEndereco As String
    Dim ClienteMunicipal As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RegistrosFiscaisIN86", "ACESSAR") Then
                    Dim ddl As New CarregarDDL
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Function Arquivo431()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String

        Dim strSQL As String
        'Dim NomeArquivo As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)
        Dim Registros As Integer

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ431.TXT"
            Dim strm As StreamWriter = Nothing
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            strSQL = "  SELECT      consulta1.Movimento, consulta1.DataDaNota, consulta1.Empresa_Id, consulta1.EndEmpresa_Id, consulta1.Cliente_Id, consulta1.EndCliente_Id, " & vbCrLf & _
                     "                      consulta1.EntradaSaida_Id, consulta1.Serie_Id, consulta1.Nota_Id, consulta1.Pedido, consulta1.TotalMercadorias, consulta1.TotalDescontos, consulta1.TotalFrete, " & vbCrLf & _
                     "                      consulta1.TotalSeguro, consulta1.TotalOutras, consulta1.TotalIpi, consulta1.TotalIcmsST, consulta1.TotalDaNota, isnull(consulta1.Proprietario,'') as Proprietario, consulta1.EndProprietario, " & vbCrLf & _
                     "                      isnull(consulta1.Quantidade,0) as quantidade, isnull(consulta1.Especie,0) as Especie, isnull(consulta1.PesoBruto,0) as PesoBruto, isnull(consulta1.PesoLiquido,0) as PesoLiquido, consulta1.FreteCifFob, consulta1.Inscricao, isnull(NotasFiscais_2.Observacoes,'') as Observacoes ,isnull(NotasFiscais_2.Situacao,1) as Situacao " & vbCrLf & _
                     " FROM         (" & vbCrLf & _
                     "							SELECT      consulta.Movimento, consulta.DataDaNota, consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, " & vbCrLf & _
                     "                                              consulta.EntradaSaida_Id, consulta.Serie_Id, consulta.Nota_Id, consulta.Pedido, consulta.TotalMercadorias, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalDescontos, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'FRETE' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalFrete, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'SEGURO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalSeguro, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'OUTRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalOutras, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalIpi, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalIcmsST, " & vbCrLf & _
                     "                                              consulta.TotalMercadorias + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END)" & vbCrLf & _
                     "                                               + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'FRETE' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'SEGURO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'OUTRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalDaNota, " & vbCrLf & _
                     "                                              NotasFiscaisXTransportadores.Proprietario, NotasFiscaisXTransportadores.EndProprietario, NotasXEmbalagens.Quantidade, NotasXEmbalagens.Especie, " & vbCrLf & _
                     "                                              NotasXEmbalagens.PesoBruto, NotasXEmbalagens.PesoLiquido, isnull(Pedidos.FreteCifFob,'') as FreteCifFob, Clientes.Inscricao" & vbCrLf & _
                     "						FROM          (" & vbCrLf & _
                     "											SELECT      NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                                                      NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, SUM(NotasFiscaisXItens.Valor) AS TotalMercadorias, " & vbCrLf & _
                     "                                                                      NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscais.Pedido, NotasFiscais.Situacao" & vbCrLf & _
                     "                                               FROM          NotasFiscais INNER JOIN" & vbCrLf & _
                     "                                                                      NotasFiscaisXItens ON  " & vbCrLf & _
                     "                                                                      NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf & _
                     "                                                                      NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf & _
                     "                                                                      NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf & _
                     "                                                                      NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                     "                                               WHERE      (year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") " & vbCrLf & _
                     "                                                                      AND (NotasFiscais.EntradaSaida_Id = 'S')   and (NotasFiscais.NossaEmissao = 'S') " & vbCrLf & _
                     "                                                                      And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "                                               GROUP BY  NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                                                      NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscais.DataDaNota, " & vbCrLf & _
                     "                                                                      NotasFiscais.Pedido, NotasFiscais.Situacao) AS consulta INNER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXEncargos ON  consulta.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                     "                                             consulta.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND consulta.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                     "                                              consulta.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              consulta.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND consulta.Nota_Id = NotasFiscaisXEncargos.Nota_Id LEFT JOIN" & vbCrLf & _
                     "                                              Pedidos ON  consulta.Empresa_Id = Pedidos.Empresa_Id AND " & vbCrLf & _
                     "                                              consulta.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND consulta.Pedido = Pedidos.Pedido_Id INNER JOIN" & vbCrLf & _
                     "                                              Clientes ON consulta.Cliente_Id = Clientes.Cliente_Id AND " & vbCrLf & _
                     "                                              consulta.EndCliente_Id = Clientes.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                     "                                              NotasXEmbalagens ON  consulta.Empresa_Id = NotasXEmbalagens.Empresa_Id AND " & vbCrLf & _
                     "                                              consulta.EndEmpresa_Id = NotasXEmbalagens.EndEmpresa_Id AND consulta.Cliente_Id = NotasXEmbalagens.Cliente_Id AND " & vbCrLf & _
                     "                                              consulta.EndCliente_Id = NotasXEmbalagens.EndCliente_Id AND consulta.EntradaSaida_Id = NotasXEmbalagens.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              consulta.Serie_Id = NotasXEmbalagens.Serie_Id AND consulta.Nota_Id = NotasXEmbalagens.Nota_Id LEFT OUTER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXTransportadores ON  " & vbCrLf & _
                     "                                              consulta.Empresa_Id = NotasFiscaisXTransportadores.Empresa_Id AND consulta.EndEmpresa_Id = NotasFiscaisXTransportadores.EndEmpresa_Id AND " & vbCrLf & _
                     "                                              consulta.Cliente_Id = NotasFiscaisXTransportadores.Cliente_Id AND consulta.EndCliente_Id = NotasFiscaisXTransportadores.EndCliente_Id AND " & vbCrLf & _
                     "                                              consulta.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id AND consulta.Serie_Id = NotasFiscaisXTransportadores.Serie_Id AND " & vbCrLf & _
                     "                                              consulta.Nota_Id = NotasFiscaisXTransportadores.Nota_Id " & vbCrLf & _
                     "                       GROUP BY  consulta.Movimento, consulta.DataDaNota, consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, " & vbCrLf & _
                     "                                              consulta.EntradaSaida_Id, consulta.Serie_Id, consulta.Nota_Id, consulta.Pedido, consulta.TotalMercadorias, NotasFiscaisXTransportadores.Proprietario, " & vbCrLf & _
                     "                                              NotasFiscaisXTransportadores.EndProprietario, NotasXEmbalagens.Quantidade, NotasXEmbalagens.Especie, NotasXEmbalagens.PesoBruto, " & vbCrLf & _
                     "                                              NotasXEmbalagens.PesoLiquido, Pedidos.FreteCifFob, Clientes.Inscricao" & vbCrLf & _
                     " UNION " & vbCrLf & _
                     "				SELECT     consulta_1.Movimento, consulta_1.DataDaNota, consulta_1.Empresa_Id, consulta_1.EndEmpresa_Id, consulta_1.Cliente_Id, " & vbCrLf & _
                     "                      consulta_1.EndCliente_Id, consulta_1.EntradaSaida_Id, consulta_1.Serie_Id, consulta_1.Nota_Id, consulta_1.Pedido, consulta_1.TotalMercadorias, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalDescontos, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'FRETE' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalFrete, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'SEGURO' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalSeguro, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'OUTRAS' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalOutras, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalIpi, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalIcmsST, " & vbCrLf & _
                     "                      consulta_1.TotalMercadorias + SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) " & vbCrLf & _
                     "                      + SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'FRETE' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) " & vbCrLf & _
                     "                      + SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'SEGURO' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) " & vbCrLf & _
                     "                      + SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'OUTRAS' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) " & vbCrLf & _
                     "                      + SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'IPI' THEN CASE WHEN Produtos.ipi > 0 THEN 0 ELSE NotasFiscaisXEncargos_1.Valor END ELSE 0 END) " & vbCrLf & _
                     "                       + SUM(CASE WHEN NotasFiscaisXEncargos_1.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos_1.Valor ELSE 0 END) AS TotalDaNota, " & vbCrLf & _
                     "                      NotasFiscaisXTransportadores_1.Proprietario, NotasFiscaisXTransportadores_1.EndProprietario, NotasXEmbalagens_1.Quantidade, NotasXEmbalagens_1.Especie, " & vbCrLf & _
                     "                      NotasXEmbalagens_1.PesoBruto, NotasXEmbalagens_1.PesoLiquido, isnull(Pedidos_1.FreteCifFob,'') as FreteCifFob, Clientes_1.Inscricao " & vbCrLf & _
                     "					FROM         ( " & vbCrLf & _
                     "									SELECT     NotasFiscais_1.Empresa_Id, NotasFiscais_1.EndEmpresa_Id, NotasFiscais_1.Cliente_Id, NotasFiscais_1.EndCliente_Id, " & vbCrLf & _
                     "                                              NotasFiscais_1.EntradaSaida_Id, NotasFiscais_1.Serie_Id, NotasFiscais_1.Nota_Id, SUM(NotasFiscaisXItens_1.Valor) AS TotalMercadorias, " & vbCrLf & _
                     "                                              NotasFiscais_1.Movimento, NotasFiscais_1.DataDaNota, NotasFiscais_1.Pedido, NotasFiscais_1.Situacao " & vbCrLf & _
                     "									FROM          NotasFiscais AS NotasFiscais_1 INNER JOIN " & vbCrLf & _
                     "                                              NotasFiscaisXItens AS NotasFiscaisXItens_1 ON  " & vbCrLf & _
                     "                                              NotasFiscais_1.Empresa_Id = NotasFiscaisXItens_1.Empresa_Id AND NotasFiscais_1.EndEmpresa_Id = NotasFiscaisXItens_1.EndEmpresa_Id AND " & vbCrLf & _
                     "                                              NotasFiscais_1.Cliente_Id = NotasFiscaisXItens_1.Cliente_Id AND NotasFiscais_1.EndCliente_Id = NotasFiscaisXItens_1.EndCliente_Id AND " & vbCrLf & _
                     "                                              NotasFiscais_1.EntradaSaida_Id = NotasFiscaisXItens_1.EntradaSaida_Id AND NotasFiscais_1.Serie_Id = NotasFiscaisXItens_1.Serie_Id AND " & vbCrLf & _
                     "                                              NotasFiscais_1.Nota_Id = NotasFiscaisXItens_1.Nota_Id " & vbCrLf & _
                     "									WHERE      (year(NotasFiscais_1.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais_1.Movimento) = " & Mes & ") AND " & vbCrLf & _
                     "                                              (NotasFiscais_1.EntradaSaida_Id = 'E') and (NotasFiscais_1.NossaEmissao = 'S') " & vbCrLf & _
                     "                                                                      And NotasFiscais_1.Empresa_Id = '" & empresa(0) & "' and NotasFiscais_1.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "									GROUP BY  NotasFiscais_1.Empresa_Id, NotasFiscais_1.EndEmpresa_Id, NotasFiscais_1.Cliente_Id, NotasFiscais_1.EndCliente_Id, " & vbCrLf & _
                     "                                              NotasFiscais_1.EntradaSaida_Id, NotasFiscais_1.Serie_Id, NotasFiscais_1.Nota_Id, NotasFiscais_1.Movimento, NotasFiscais_1.DataDaNota, " & vbCrLf & _
                     "                                              NotasFiscais_1.Pedido, NotasFiscais_1.Situacao) AS consulta_1 INNER JOIN " & vbCrLf & _
                     "						  NotasFiscaisXEncargos AS NotasFiscaisXEncargos_1 ON  " & vbCrLf & _
                     "						  consulta_1.Empresa_Id = NotasFiscaisXEncargos_1.Empresa_Id AND consulta_1.EndEmpresa_Id = NotasFiscaisXEncargos_1.EndEmpresa_Id AND " & vbCrLf & _
                     "						  consulta_1.Cliente_Id = NotasFiscaisXEncargos_1.Cliente_Id AND consulta_1.EndCliente_Id = NotasFiscaisXEncargos_1.EndCliente_Id AND " & vbCrLf & _
                     "						  consulta_1.EntradaSaida_Id = NotasFiscaisXEncargos_1.EntradaSaida_Id AND consulta_1.Serie_Id = NotasFiscaisXEncargos_1.Serie_Id AND " & vbCrLf & _
                     "						  consulta_1.Nota_Id = NotasFiscaisXEncargos_1.Nota_Id LEFT JOIN " & vbCrLf & _
                     "						  Pedidos AS Pedidos_1 ON consulta_1.Empresa_Id = Pedidos_1.Empresa_Id AND " & vbCrLf & _
                     "						  consulta_1.EndEmpresa_Id = Pedidos_1.EndEmpresa_Id AND consulta_1.Pedido = Pedidos_1.Pedido_Id INNER JOIN " & vbCrLf & _
                     "						  Clientes AS Clientes_1 ON consulta_1.Cliente_Id = Clientes_1.Cliente_Id AND " & vbCrLf & _
                     "						  consulta_1.EndCliente_Id = Clientes_1.Endereco_Id INNER JOIN " & vbCrLf & _
                     "						  Produtos ON  NotasFiscaisXEncargos_1.Produto_Id = Produtos.Produto_Id LEFT OUTER JOIN " & vbCrLf & _
                     "						  NotasXEmbalagens AS NotasXEmbalagens_1 ON  " & vbCrLf & _
                     "						  consulta_1.Empresa_Id = NotasXEmbalagens_1.Empresa_Id AND consulta_1.EndEmpresa_Id = NotasXEmbalagens_1.EndEmpresa_Id AND " & vbCrLf & _
                     "						  consulta_1.Cliente_Id = NotasXEmbalagens_1.Cliente_Id AND consulta_1.EndCliente_Id = NotasXEmbalagens_1.EndCliente_Id AND " & vbCrLf & _
                     "						  consulta_1.EntradaSaida_Id = NotasXEmbalagens_1.EntradaSaida_Id AND consulta_1.Serie_Id = NotasXEmbalagens_1.Serie_Id AND " & vbCrLf & _
                     "						  consulta_1.Nota_Id = NotasXEmbalagens_1.Nota_Id LEFT OUTER JOIN" & vbCrLf & _
                     "						  NotasFiscaisXTransportadores AS NotasFiscaisXTransportadores_1 ON " & vbCrLf & _
                     "						  consulta_1.Empresa_Id = NotasFiscaisXTransportadores_1.Empresa_Id AND consulta_1.EndEmpresa_Id = NotasFiscaisXTransportadores_1.EndEmpresa_Id AND " & vbCrLf & _
                     "						  consulta_1.Cliente_Id = NotasFiscaisXTransportadores_1.Cliente_Id AND consulta_1.EndCliente_Id = NotasFiscaisXTransportadores_1.EndCliente_Id AND " & vbCrLf & _
                     "						  consulta_1.EntradaSaida_Id = NotasFiscaisXTransportadores_1.EntradaSaida_Id AND consulta_1.Serie_Id = NotasFiscaisXTransportadores_1.Serie_Id AND " & vbCrLf & _
                     "						  consulta_1.Nota_Id = NotasFiscaisXTransportadores_1.Nota_Id " & vbCrLf & _
                     " GROUP BY  consulta_1.Movimento, consulta_1.DataDaNota, consulta_1.Empresa_Id, consulta_1.EndEmpresa_Id, consulta_1.Cliente_Id, " & vbCrLf & _
                     "                      consulta_1.EndCliente_Id, consulta_1.EntradaSaida_Id, consulta_1.Serie_Id, consulta_1.Nota_Id, consulta_1.Pedido, consulta_1.TotalMercadorias, " & vbCrLf & _
                     "                      NotasFiscaisXTransportadores_1.Proprietario, NotasFiscaisXTransportadores_1.EndProprietario, NotasXEmbalagens_1.Quantidade, NotasXEmbalagens_1.Especie, " & vbCrLf & _
                     "                      NotasXEmbalagens_1.PesoBruto, NotasXEmbalagens_1.PesoLiquido, Pedidos_1.FreteCifFob, Clientes_1.Inscricao) AS consulta1 INNER JOIN" & vbCrLf & _
                     "                      NotasFiscais AS NotasFiscais_2 ON  consulta1.Empresa_Id = NotasFiscais_2.Empresa_Id AND " & vbCrLf & _
                     "                      consulta1.EndEmpresa_Id = NotasFiscais_2.EndEmpresa_Id AND consulta1.Cliente_Id = NotasFiscais_2.Cliente_Id AND " & vbCrLf & _
                     "                      consulta1.EndCliente_Id = NotasFiscais_2.EndCliente_Id AND consulta1.EntradaSaida_Id = NotasFiscais_2.EntradaSaida_Id AND " & vbCrLf & _
                     "                      consulta1.Serie_Id = NotasFiscais_2.Serie_Id AND consulta1.Nota_Id = NotasFiscais_2.Nota_Id" & vbCrLf & _
                     "                     AND      (consulta1.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (consulta1.Serie_Id <> '101' ) AND (consulta1.TotalDaNota <> 0) " & vbCrLf & _
                     " ORDER BY consulta1.Empresa_Id, consulta1.EndEmpresa_Id, consulta1.Cliente_Id, consulta1.EndCliente_Id, consulta1.EntradaSaida_Id, consulta1.Serie_Id, consulta1.Nota_Id" & vbCrLf


            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr

                    strLinha = Funcoes.AlinharEsquerda(.Item("EntradaSaida_Id"), 1, " ") 'Indicador do movimento
                    strLinha &= Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    strLinha &= Funcoes.AlinharEsquerda(CDate(.Item("DataDaNota")).ToStrDate(), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Cliente_Id"), 14, " ") 'codigo participante
                    strLinha &= Funcoes.AlinharEsquerda(CDate(.Item("Movimento")).ToStrDate(), 8, " ")  'data saida/entrada
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalMercadorias")).Replace(",", ""), 17, "0")     'total mercadorias
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalDescontos")).Replace(",", ""), 17, "0")     'total descontos
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalFrete")).Replace(",", ""), 17, "0")      'total frete
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalSeguro")).Replace(",", ""), 17, "0")     'total seguro
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalOutras")).Replace(",", ""), 17, "0")     'total outras
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalIPI")).Replace(",", ""), 17, "0")  ' total ipi
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalIcmsST")).Replace(",", ""), 17, "0")   'total icms st
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("TotalDaNota")).Replace(",", ""), 17, "0") ' total da nota
                    strLinha &= IIf(.Item("TotalIcmsST") > 0, Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Inscricao")), 14, " "), Funcoes.AlinharEsquerda(" ", 14, " ")) ' insricao substituido 
                    strLinha &= Funcoes.AlinharEsquerda("Rodoviario", 15, " ") 'via de transporte
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Proprietario"), 14, " ") ' transportador
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("Quantidade")).Replace(",", ""), 17, " ") 'quantidade volumes
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Especie"), 10, " ")  'especie de volumes
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("PesoBruto") * 1000, 0).Replace(".", ""), 17, " ")   'peso bruto
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("PesoLiquido") * 1000, 0).Replace(".", ""), 17, " ") 'peso liquido
                    strLinha &= IIf(.Item("FreteCifFob") = "FOB", Funcoes.AlinharEsquerda("FOB", 3, " "), Funcoes.AlinharEsquerda("SIF", 3, " "))    'Modalidade frete
                    strLinha &= Funcoes.AlinharEsquerda(" ", 15, " ")  'identificação do veiculo
                    strLinha &= IIf(.Item("Situacao") = 1, Funcoes.AlinharEsquerda("N", 1, " "), Funcoes.AlinharEsquerda("S", 1, " "))   'Indicador de Situação
                    strLinha &= Funcoes.AlinharEsquerda("2", 1, " ")   'TipoFatura
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.RemoveAllEnterKey(Funcoes.EliminarCaracteresEspeciais(.Item("Observacoes"))), 45, " ")      'observação
                    strLinha &= Funcoes.AlinharEsquerda(" ", 50, " ")      'Ato declaratorio
                    strLinha &= Funcoes.AlinharEsquerda(" ", 2, " ")      'modelo documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda(" ", 5, " ")      'serie documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda("0", 9, " ")      'numero documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda(" ", 8, " ")      'data documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda(" ", 14, " ")      'codigo do participante referenciado


                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next

        Next


        'strm.Close()
        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo432()

        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String

        Dim strSQL As String
        'Dim NomeArquivo As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)
        Dim Registros As Integer

        Dim NomeArquivo2 As String = "Sva/432.html"
        Dim NomeArquivo3 As String = Server.MapPath(NomeArquivo2)
        Dim arquivo1 As String = NomeArquivo3
        Dim Sql As String = ""
        Dim strmHtml As StreamWriter = Nothing

        Sql = "<html>" & vbCrLf & _
              "<head>" & vbCrLf & _
              "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
              "<title>NGS - Anotação de Responsabilidade Técnica</title>" & vbCrLf & _
              "<style type='text/css'>" & vbCrLf & _
              "H6 {page-break-after:always}" & vbCrLf & _
              "A.A1 {font-family: Ms serif;font-size: 5pt;line height: 15px}" & vbCrLf & _
              "A.A2 {font-family: Ms serif;font-size: 6pt;line height: 15px}" & vbCrLf & _
              "A.A3 {font-family: Ms serif;font-size: 7pt;line height: 15px}" & vbCrLf & _
              "A.A4 {font-family: Ms serif;font-size: 8pt;line height: 15px}" & vbCrLf & _
              "A.A5 {font-family: Ms serif;font-size: 9pt;line height: 15px}" & vbCrLf & _
              "A.A6 {font-family: Ms serif;font-size: 10pt;line height: 15px}" & vbCrLf & _
              "A.A7 {font-family: Ms serif bold;font-size: 8pt;line height: 15px}" & vbCrLf & _
              "A.A8 {font-family: Ms serif bold;font-size: 9pt;line height: 15px}" & vbCrLf & _
              "A.A9 {font-family: Ms serif bold;font-size: 10pt;line height: 15px}" & vbCrLf & _
              "A.A10 {font-family: Ms serif bold;font-size: 11pt;line height: 15px}" & vbCrLf & _
              "A.A11 {font-family: Ms serif bold;font-size: 12pt;line height: 15px}" & vbCrLf & _
              "A.A12 {font-family: Ms serif bold;font-size: 13pt;line height: 15px}" & vbCrLf & _
              "A.A13 {font-family: Ms serif bold;font-size: 14pt;line height: 15px}" & vbCrLf & _
              "</style>" & vbCrLf & _
              "</head>" & vbCrLf & _
              "<body text=#000000 bgcolor=#FFFFFF>" & vbCrLf & _
              "<table border='0' width='100%' cellpadding='0' cellspacing='0'>" & vbCrLf & _
              "<tr>" & vbCrLf & _
              "<td>Empresa</td>" & vbCrLf & _
              "<td>Cliente</td>" & vbCrLf & _
              "<td>Nota</td>" & vbCrLf & _
              "<td>Serie</td>" & vbCrLf & _
              "<td>E/S</td>" & vbCrLf & _
              "<td>EmissaoNf</td>" & vbCrLf & _
              "<td>Operacao</td>" & vbCrLf & _
              "<td>CFOP</td>" & vbCrLf & _
              "<td>QuntFiscal</td>" & vbCrLf & _
              "<td>Valor</td>" & vbCrLf & _
              "</tr>" & vbCrLf

        '        If Dir(arquivo1).Length > 0 Then Kill(arquivo1)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ432.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            'strSQL = "SELECT DISTINCT IndicadorMovimento, ModeloDocumento, SerieSubserieDocumento, NumeroDocumento, DataEmissaoDocumento, "
            'strSQL &= "NumeroItem, CodigoMercadoriaServico, DescricaoComplementar, CFOP, CodigoNaturezaOperacao, "
            'strSQL &= "COALESCE(ClassificacaoFiscalMercadoria, '" & StrDup(8, " ") & "') AS ClassificacaoFiscalMercadoria, Quantidade, Unidade, ValorUnitario, ValorTotalItem, "
            'strSQL &= "COALESCE(ValorDescontoItem, 0) AS ValorDescontoItem, IndicadorTributacaoIPI, COALESCE(AliquotaIPI, 0) AS AliquotaIPI, "
            'strSQL &= "COALESCE(BaseCalculoIPI, 0) AS BaseCalculoIPI, COALESCE(ValorIPI, 0) AS ValorIPI, SituacaoTributariaEstadual, "
            'strSQL &= "IndicadorTributacaoICMS, AliquotaICMS, BaseCalculoICMSProprio, ValorICMSProprio, "
            'strSQL &= "COALESCE(BaseCalculoICMSSubsTrib, 0) AS BaseCalculoICMSSubsTrib, "
            'strSQL &= "COALESCE(ValorICMSSubsTrib, 0) AS ValorICMSSubsTrib, IndicadorMovFisMercadoria "
            'strSQL &= "FROM ArquivoItensMercadoriasServicos432"

            strSQL = "   SELECT consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "          consulta.Nota_Id, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, isnull(consulta.Operacao_Id,1) as Operacao_Id, " & vbCrLf & _
                     "          isnull(consulta.SubOperacoes_Id,1) as SubOperacoes_Id, consulta.NCM,consulta.Unidade, " & vbCrLf & _
                     "          isnull(consulta.QuantidadeFiscal,1) as QuantidadeFiscal, " & vbCrLf & _
                     "          consulta.Valor / (isnull(consulta.QuantidadeFiscal,1)) as Unitario," & vbCrLf & _
                     "          consulta.Valor, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseIPI, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaIPI, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorIPI," & vbCrLf & _
                     "			SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaICMS, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS," & vbCrLf & _
                     "			SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMSST, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaICMSST, " & vbCrLf & _
                     "          SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMSST" & vbCrLf & _
                     "   FROM  (SELECT  NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                  NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.DataDaNota, NotasFiscaisXItens.Produto_Id, Produtos.Nome, " & vbCrLf & _
                     "                  NotasFiscaisXItens.Cfop_Id, SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, Produtos.NCM,Produtos.Unidade," & vbCrLf & _
                     "                  NotasFiscaisXItens.Valor, " & vbCrLf & _
                     "                  case " & vbCrLf & _
                     "                    when NotasFiscaisXItens.QuantidadeFiscal is null  " & vbCrLf & _
                     "                      then 1" & vbCrLf & _
                     "                    when NotasFiscaisXItens.QuantidadeFiscal = 0" & vbCrLf & _
                     "                      then 1" & vbCrLf & _
                     "                    else NotasFiscaisXItens.QuantidadeFiscal" & vbCrLf & _
                     "                  end as QuantidadeFiscal " & vbCrLf & _
                     "             FROM NotasFiscais " & vbCrLf & _
                     "            INNER JOIN NotasFiscaisXItens " & vbCrLf & _
                     "               ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                     "              AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                     "              AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id  " & vbCrLf & _
                     "              AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                     "              AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                     "              AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                     "              AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                     "            INNER JOIN Produtos" & vbCrLf & _
                     "               ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
                     "             LEFT JOIN SubOperacoes " & vbCrLf & _
                     "               ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id  " & vbCrLf & _
                     "              AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                     "            WHERE (year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ")" & vbCrLf & _
                     "              and NotasFiscais.NossaEmissao = 'S' " & vbCrLf & _
                     "              And NotasFiscais.Empresa_Id = '" & empresa(0) & "' " & vbCrLf & _
                     "              and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "            ) AS consulta " & vbCrLf & _
                     "    INNER JOIN NotasFiscaisXEncargos " & vbCrLf & _
                     "       ON consulta.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id  " & vbCrLf & _
                     "      AND consulta.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf & _
                     "      AND consulta.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf & _
                     "      AND consulta.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf & _
                     "      AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id " & vbCrLf & _
                     "      AND consulta.Serie_Id        = NotasFiscaisXEncargos.Serie_Id " & vbCrLf & _
                     "      AND consulta.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf & _
                     "      AND consulta.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                     "      AND (consulta.DataDaNota     <>  '" & ddlAno.SelectedValue & "/01/01') " & vbCrLf & _
                     "      AND consulta.Serie_Id        <> '101'" & vbCrLf & _
                     "      AND consulta.Valor           <> 0" & vbCrLf & _
                     "   GROUP BY  consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, consulta.Operacao_Id, consulta.SubOperacoes_Id, consulta.NCM,consulta.Unidade, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Valor "

            Dim nota As String = ""
            Dim item As Integer = 1
            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr

                    If nota = CStr(.Item("Nota_Id")) Then
                        item = item + 1
                    Else
                        item = 1
                    End If
                    strLinha = Funcoes.AlinharEsquerda(.Item("EntradaSaida_Id"), 1, " ") 'Indicador do movimento
                    strLinha &= Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    nota = .Item("Nota_Id")
                    strLinha &= Funcoes.AlinharEsquerda(CDate(.Item("DataDaNota")).ToStrDate(), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(item, 3, " ") 'numero do item
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Produto_Id"), 20, " ") 'codigo produto
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.RemoveAllEnterKey(Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))), 45, " ")      'Nome produto
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Cfop_Id"), 4, " ")     'total mercadorias
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Operacao_Id") & .Item("SubOperacoes_Id"), 6, " ")     'Operacao e subOperacao
                    strLinha &= Funcoes.AlinharEsquerda(.Item("NCM"), 8, " ")   'total descontos
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(IIf(.Item("QuantidadeFiscal") = 0, 1, .Item("QuantidadeFiscal")) * 1000, 0).Replace(".", ""), 17, "0")   'Quantidade
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Unidade"), 3, " ")     ' Unidade
                    'If .Item("BaseCalculo") > 1 Then
                    '    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(FormatNumber((.Item("ValorOficial") / 1), 4).Replace(".", "") * 10000, 0).Replace(".", ""), 17, " ")   'Valor unitario
                    'Else
                    '    strLinha &= Funcoes.AlinharDireita(FormatNumber(FormatNumber((.Item("ValorOficial") / .Item("QuantidadeFiscal")), 4).Replace(".", "") * 10000, 0).Replace(".", ""), 17, "0")   'Valor unitario
                    'End If
                    'strLinha &= Funcoes.AlinharDireita(CStr(.Item("Unitario")).Replace(",", ""), 17, "0")   'Valor unitario
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("Unitario") * 10000, 0).Replace(".", ""), 17, " ")
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("Valor")).Replace(",", ""), 17, "0")   'Valor
                    strLinha &= Funcoes.AlinharDireita(0, 17, " ")  ' valor desconto item 
                    If .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(1, 1, " ")  ' Indicador tributação ipi
                    ElseIf .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(2, 1, " ")  ' Indicador tributação ipi
                    ElseIf .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(1, 1, " ")  ' Indicador tributação ipi
                    ElseIf .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(2, 1, " ")  ' Indicador tributação ipi
                    End If
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("AliquotaIPI")).Replace(",", ""), 5, "0")   'Aliquota IPI
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("BaseIPI")).Replace(",", ""), 17, "0")   'Base IPI
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("ValorIPI")).Replace(",", ""), 17, "0")  ' total ipi
                    strLinha &= Funcoes.AlinharDireita("0", 3, " ")   'Situacao tributaria estadual
                    If .Item("ValorICMS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(1, 1, " ")
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(2, 1, " ")
                    End If
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("AliquotaICMS")).Replace(",", ""), 5, "0")   'Aliquota ICMS
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("BaseICMS")).Replace(",", ""), 17, "0")   'Base ICMS
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("ValorICMS")).Replace(",", ""), 17, "0")   'Valor ICMS
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("BaseICMSST")).Replace(",", ""), 17, "0")   'Base ICMS-ST
                    strLinha &= Funcoes.AlinharDireita(CStr(.Item("ValorICMSST")).Replace(",", ""), 17, "0")   'Valor ICMS-ST
                    strLinha &= Funcoes.AlinharEsquerda("N", 1, " ") ' Indicador movimentacao

                    If .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("00", 2, " ")   'Situacao tributaria IPI
                    ElseIf .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("49", 2, " ")   'Situacao tributaria IPI
                    ElseIf .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("50", 2, " ")   'Situacao tributaria IPI
                    ElseIf .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("99", 2, " ")   'Situacao tributaria IPI
                    End If

                    Sql &= "<tr>" & vbCrLf & _
                           "<td>" & empresa(0) & "-" & empresa(1) & "</td>" & vbCrLf & _
                           "<td>" & .Item("Cliente_Id") & "-" & .Item("EndCliente_Id") & "</td>" & vbCrLf & _
                           "<td>" & .Item("Nota_Id") & "</td>" & vbCrLf & _
                           "<td>" & .Item("Serie_Id") & "</td>" & vbCrLf & _
                           "<td>" & .Item("EntradaSaida_Id") & "</td>" & vbCrLf & _
                           "<td>" & CDate(.Item("DataDaNota")).ToStrDate() & "</td>" & vbCrLf & _
                           "<td>" & .Item("Operacao_Id") & "-" & .Item("SubOperacoes_Id") & "</td>" & vbCrLf & _
                           "<td>" & .Item("Cfop_Id") & "</td>" & vbCrLf & _
                           "<td>" & .Item("QuantidadeFiscal") & "</td>" & vbCrLf & _
                           "<td>" & .Item("Valor") & "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf

                End With
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1

            Next
        Next

        Sql &= "</table>"
        Try
            strmHtml = New StreamWriter(arquivo1, True)
            strmHtml.WriteLine(Sql)
            strmHtml.Close()
            'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo3 & "');", True)
        Catch ex As Exception
            strmHtml.Close()
            MsgBox(Me.Page, ex.Message.ToString)
        End Try

        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo433()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String

        Dim strSQL As String
        'Dim NomeArquivo As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)
        Dim Registros As Integer

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ433.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            'strSQL = "SELECT DISTINCT ModeloDocumento, SerieSubserieDocumento, NumeroDocumento, DataEmissaoDocumento, "
            'strSQL &= "CodigoParticipante, DataEntrada, ValorTotalMercadorias, COALESCE(ValorTotalDesconto, 0) AS ValorTotalDesconto, "
            'strSQL &= "ValorFrete, COALESCE(ValorSeguro, 0) AS ValorSeguro, COALESCE(ValorOutrasDespesas, 0) AS ValorOutrasDespesas, "
            'strSQL &= "ValorTotalIPI, ValorTotalICMSSubsTrib, ValorTotalNotaFiscal, InscricaoEstadualSubsTrib, TipoFatura, "
            'strSQL &= "COALESCE(Observacao, '" & StrDup(45, " ") & "') AS Observacao "
            'strSQL &= "FROM ArquivoMestreMercadoriasServicosEntradas433"

            strSQL = "  SELECT      consulta1.Movimento, consulta1.DataDaNota, consulta1.Empresa_Id, consulta1.EndEmpresa_Id, consulta1.Cliente_Id, consulta1.EndCliente_Id, " & vbCrLf & _
                     "                      consulta1.EntradaSaida_Id, consulta1.Serie_Id, consulta1.Nota_Id, consulta1.Pedido, consulta1.TotalMercadorias, consulta1.TotalDescontos, consulta1.TotalFrete, " & vbCrLf & _
                     "                      consulta1.TotalSeguro, consulta1.TotalOutras, consulta1.TotalIpi, consulta1.TotalIcmsST, consulta1.TotalDaNota, isnull(consulta1.Proprietario,'') as Proprietario, consulta1.EndProprietario, " & vbCrLf & _
                     "                      isnull(consulta1.Quantidade,0) as quantidade, isnull(consulta1.Especie,0) as Especie, isnull(consulta1.PesoBruto,0) as PesoBruto, isnull(consulta1.PesoLiquido,0) as PesoLiquido, consulta1.FreteCifFob, consulta1.Inscricao, isnull(NotasFiscais_2.Observacoes,'') as Observacoes ,isnull(NotasFiscais_2.Situacao,1) as Situacao " & vbCrLf & _
                     " FROM         (" & vbCrLf & _
                     "							SELECT      consulta.Movimento, consulta.DataDaNota, consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, " & vbCrLf & _
                     "                                              consulta.EntradaSaida_Id, consulta.Serie_Id, consulta.Nota_Id, consulta.Pedido, consulta.TotalMercadorias - SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'AFIXAR' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) as TotalMercadorias, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalDescontos, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'FRETE' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalFrete, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'SEGURO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalSeguro, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'OUTRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalOutras, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalIpi, " & vbCrLf & _
                     "                                              SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS TotalIcmsST, " & vbCrLf & _
                     "                                              consulta.TotalMercadorias + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END)" & vbCrLf & _
                     "                                               + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'FRETE' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'SEGURO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'OUTRAS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              + SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              - SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'AFIXAR' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) " & vbCrLf & _
                     "                                              AS TotalDaNota, " & vbCrLf & _
                     "                                              NotasFiscaisXTransportadores.Proprietario, NotasFiscaisXTransportadores.EndProprietario, NotasXEmbalagens.Quantidade, NotasXEmbalagens.Especie, " & vbCrLf & _
                     "                                              NotasXEmbalagens.PesoBruto, NotasXEmbalagens.PesoLiquido, Pedidos.FreteCifFob, Clientes.Inscricao" & vbCrLf & _
                     "						FROM          (" & vbCrLf & _
                     "											SELECT      NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                                                      NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, SUM(NotasFiscaisXItens.Valor) AS TotalMercadorias, " & vbCrLf & _
                     "                                                                      NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscais.Pedido, NotasFiscais.Situacao" & vbCrLf & _
                     "                                               FROM          NotasFiscais INNER JOIN " & vbCrLf & _
                     "                                                                      NotasFiscaisXItens ON  " & vbCrLf & _
                     "                                                                      NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf & _
                     "                                                                      NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf & _
                     "                                                                      NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf & _
                     "                                                                      NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                     "                                               WHERE      (year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ")" & vbCrLf & _
                     "                                                                      AND (NotasFiscais.EntradaSaida_Id = 'E')   and (NotasFiscais.NossaEmissao = 'N')  " & vbCrLf & _
                     "                                                                      And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "                                               GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                                                      NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscais.DataDaNota, " & vbCrLf & _
                     "                                                                      NotasFiscais.Pedido, NotasFiscais.Situacao) AS consulta INNER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXEncargos ON  consulta.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                     "                                             consulta.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND consulta.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                     "                                              consulta.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              consulta.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND consulta.Nota_Id = NotasFiscaisXEncargos.Nota_Id LEFT JOIN" & vbCrLf & _
                     "                                              Pedidos ON consulta.Empresa_Id = Pedidos.Empresa_Id AND " & vbCrLf & _
                     "                                              consulta.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND consulta.Pedido = Pedidos.Pedido_Id INNER JOIN" & vbCrLf & _
                     "                                              Clientes ON consulta.Cliente_Id = Clientes.Cliente_Id AND " & vbCrLf & _
                     "                                              consulta.EndCliente_Id = Clientes.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                     "                                              NotasXEmbalagens ON consulta.Empresa_Id = NotasXEmbalagens.Empresa_Id AND " & vbCrLf & _
                     "                                              consulta.EndEmpresa_Id = NotasXEmbalagens.EndEmpresa_Id AND consulta.Cliente_Id = NotasXEmbalagens.Cliente_Id AND " & vbCrLf & _
                     "                                              consulta.EndCliente_Id = NotasXEmbalagens.EndCliente_Id AND consulta.EntradaSaida_Id = NotasXEmbalagens.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              consulta.Serie_Id = NotasXEmbalagens.Serie_Id AND consulta.Nota_Id = NotasXEmbalagens.Nota_Id LEFT OUTER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXTransportadores ON " & vbCrLf & _
                     "                                              consulta.Empresa_Id = NotasFiscaisXTransportadores.Empresa_Id AND consulta.EndEmpresa_Id = NotasFiscaisXTransportadores.EndEmpresa_Id AND " & vbCrLf & _
                     "                                              consulta.Cliente_Id = NotasFiscaisXTransportadores.Cliente_Id AND consulta.EndCliente_Id = NotasFiscaisXTransportadores.EndCliente_Id AND " & vbCrLf & _
                     "                                              consulta.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id AND consulta.Serie_Id = NotasFiscaisXTransportadores.Serie_Id AND " & vbCrLf & _
                     "                                              consulta.Nota_Id = NotasFiscaisXTransportadores.Nota_Id " & vbCrLf & _
                     "                       GROUP BY consulta.Movimento, consulta.DataDaNota, consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, " & vbCrLf & _
                     "                                              consulta.EntradaSaida_Id, consulta.Serie_Id, consulta.Nota_Id, consulta.Pedido, consulta.TotalMercadorias, NotasFiscaisXTransportadores.Proprietario, " & vbCrLf & _
                     "                                              NotasFiscaisXTransportadores.EndProprietario, NotasXEmbalagens.Quantidade, NotasXEmbalagens.Especie, NotasXEmbalagens.PesoBruto, " & vbCrLf & _
                     "                                              NotasXEmbalagens.PesoLiquido, Pedidos.FreteCifFob, Clientes.Inscricao) AS consulta1 INNER JOIN" & vbCrLf & _
                     "                      NotasFiscais AS NotasFiscais_2 ON consulta1.Empresa_Id = NotasFiscais_2.Empresa_Id AND " & vbCrLf & _
                     "                      consulta1.EndEmpresa_Id = NotasFiscais_2.EndEmpresa_Id AND consulta1.Cliente_Id = NotasFiscais_2.Cliente_Id AND " & vbCrLf & _
                     "                      consulta1.EndCliente_Id = NotasFiscais_2.EndCliente_Id AND consulta1.EntradaSaida_Id = NotasFiscais_2.EntradaSaida_Id AND " & vbCrLf & _
                     "                      consulta1.Serie_Id = NotasFiscais_2.Serie_Id AND consulta1.Nota_Id = NotasFiscais_2.Nota_Id" & vbCrLf & _
                     "                     AND      (consulta1.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (consulta1.Serie_Id <> '101' ) AND (consulta1.TotalDaNota <> 0) " & vbCrLf & _
                     " ORDER BY consulta1.Empresa_Id, consulta1.EndEmpresa_Id, consulta1.Cliente_Id, consulta1.EndCliente_Id, consulta1.EntradaSaida_Id, consulta1.Serie_Id, consulta1.Nota_Id" & vbCrLf


            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr

                    strLinha = Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    strLinha &= Funcoes.AlinharEsquerda(CDate(.Item("DataDaNota")).ToStrDate(), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Cliente_Id")), 14, " ") 'codigo participante
                    strLinha &= Funcoes.AlinharEsquerda(CDate(.Item("Movimento")).ToStrDate(), 8, " ")  'data saida/entrada
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalMercadorias") * 100, 0).Replace(".", ""), 17, " ")     'total mercadorias
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalDescontos") * 100, 0).Replace(".", ""), 17, " ")   'total descontos
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalFrete") * 100, 0).Replace(".", ""), 17, " ")   'total frete
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalSeguro") * 100, 0).Replace(".", ""), 17, " ")    'total seguro
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalOutras") * 100, 0).Replace(".", ""), 17, " ")     'total outras
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalIPI") * 100, 0).Replace(".", ""), 17, " ")  ' total ipi
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalIcmsST") * 100, 0).Replace(".", ""), 17, " ")   'total icms st
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("TotalDaNota") * 100, 0).Replace(".", ""), 17, " ") ' total da nota
                    strLinha &= IIf(.Item("TotalIcmsST") > 0, Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Inscricao")), 14, " "), Funcoes.AlinharEsquerda(" ", 14, " ")) ' insricao substituido 
                    strLinha &= Funcoes.AlinharEsquerda("2", 1, " ")   'TipoFatura
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.RemoveAllEnterKey(Funcoes.EliminarCaracteresEspeciais(.Item("Observacoes"))), 45, " ")      'observação
                    strLinha &= Funcoes.AlinharEsquerda(" ", 50, " ")      'Ato declaratorio
                    strLinha &= Funcoes.AlinharEsquerda(" ", 2, " ")      'modelo documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda(" ", 5, " ")      'serie documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda("0", 9, " ")      'numero documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda(" ", 8, " ")      'data documento referenciado
                    strLinha &= Funcoes.AlinharEsquerda(" ", 14, " ")      'codigo do participante referenciado


                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next
        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo434()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String

        Dim strSQL As String
        'Dim NomeArquivo As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)
        Dim Registros As Integer

        Registros = 0


        Dim NomeArquivo2 As String = "Sva/434.html"
        Dim NomeArquivo3 As String = Server.MapPath(NomeArquivo2)
        Dim arquivo1 As String = NomeArquivo3
        Dim Sql As String = ""
        Dim strmHtml As StreamWriter = Nothing

        Sql = "<html>" & vbCrLf & _
              "<head>" & vbCrLf & _
              "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
              "<title>NGS - Anotação de Responsabilidade Técnica</title>" & vbCrLf & _
              "<style type='text/css'>" & vbCrLf & _
              "H6 {page-break-after:always}" & vbCrLf & _
              "A.A1 {font-family: Ms serif;font-size: 5pt;line height: 15px}" & vbCrLf & _
              "A.A2 {font-family: Ms serif;font-size: 6pt;line height: 15px}" & vbCrLf & _
              "A.A3 {font-family: Ms serif;font-size: 7pt;line height: 15px}" & vbCrLf & _
              "A.A4 {font-family: Ms serif;font-size: 8pt;line height: 15px}" & vbCrLf & _
              "A.A5 {font-family: Ms serif;font-size: 9pt;line height: 15px}" & vbCrLf & _
              "A.A6 {font-family: Ms serif;font-size: 10pt;line height: 15px}" & vbCrLf & _
              "A.A7 {font-family: Ms serif bold;font-size: 8pt;line height: 15px}" & vbCrLf & _
              "A.A8 {font-family: Ms serif bold;font-size: 9pt;line height: 15px}" & vbCrLf & _
              "A.A9 {font-family: Ms serif bold;font-size: 10pt;line height: 15px}" & vbCrLf & _
              "A.A10 {font-family: Ms serif bold;font-size: 11pt;line height: 15px}" & vbCrLf & _
              "A.A11 {font-family: Ms serif bold;font-size: 12pt;line height: 15px}" & vbCrLf & _
              "A.A12 {font-family: Ms serif bold;font-size: 13pt;line height: 15px}" & vbCrLf & _
              "A.A13 {font-family: Ms serif bold;font-size: 14pt;line height: 15px}" & vbCrLf & _
              "</style>" & vbCrLf & _
              "</head>" & vbCrLf & _
              "<body text=#000000 bgcolor=#FFFFFF>" & vbCrLf

        Sql = "<table border='0' width='100%' cellpadding='0' cellspacing='0'>" & vbCrLf & _
              "<tr>" & vbCrLf & _
              "<td>Empresa</td>" & vbCrLf & _
              "<td>Cliente</td>" & vbCrLf & _
              "<td>Nota</td>" & vbCrLf & _
              "<td>Serie</td>" & vbCrLf & _
              "<td>E/S</td>" & vbCrLf & _
              "<td>EmissaoNf</td>" & vbCrLf & _
              "<td>Operacao</td>" & vbCrLf & _
              "<td>CFOP</td>" & vbCrLf & _
              "<td>QuntFiscal</td>" & vbCrLf & _
              "<td>Valor</td>" & vbCrLf & _
              "</tr>" & vbCrLf

        'If Dir(arquivo1).Length > 0 Then Kill(arquivo1)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ434.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            'strSQL = "SELECT DISTINCT ModeloDocumento, SerieSubserieDocumento, NumeroDocumento, DataEmissaoDocumento, CodigoParticipante, "
            'strSQL &= "NumeroItem, CodigoMercadoriaServico, DescricaoComplementarProduto, CFOP, NaturezaOperacao, "
            'strSQL &= "COALESCE(ClassificacaoFiscalMercadoria, '" & StrDup(8, " ") & "') AS ClassificacaoFiscalMercadoria, Quantidade, "
            'strSQL &= "Unidade, ValorUnitario, ValorTotalItem, COALESCE(ValorDescontoItem, 0) AS ValorDescontoItem, "
            'strSQL &= "IndicadorTributacaoIPI, COALESCE(AliquotaIPI, 0) AS AliquotaIPI, COALESCE(BaseCalculoIPI, 0) AS BaseCalculoIPI, "
            'strSQL &= "COALESCE(ValorIPI, 0) AS ValorIPI, SituacaoTributariaEstadual, IndicadorTributacaoICMS, AliquotaICMS, "
            'strSQL &= "BaseCalculoICMSProprio, ValorICMSProprio, COALESCE(BaseCalculoICMSSubTrib, 0) AS BaseCalculoICMSSubTrib, "
            'strSQL &= "COALESCE(ValorICMSSubsTrib, 0) AS ValorICMSSubsTrib, IndicadorMovFisMercadoria "
            'strSQL &= "FROM ArquivoItensMercadoriasServicosEntradas434"

            strSQL = "   SELECT      consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, isnull(consulta.Operacao_Id,1) as Operacao_Id, isnull(consulta.SubOperacoes_Id,1) as SubOperacoes_Id, consulta.NCM,consulta.Unidade, " & vbCrLf & _
                     "                      isnull(consulta.QuantidadeFiscal,1) as QuantidadeFiscal, consulta.Unitario,  " & vbCrLf & _
                     "                      consulta.Valor - isnull(SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'AFIXAR' THEN NotasFiscaisXEncargos.Valor  END),0)  as Valor,  " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseIPI, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaIPI, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IPI' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorIPI," & vbCrLf & _
                     "					  SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaICMS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMS," & vbCrLf & _
                     "					  SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseICMSST, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaICMSST, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS-ST' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorICMSST" & vbCrLf & _
                     "   FROM         (SELECT      NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                              NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.DataDaNota, NotasFiscaisXItens.Produto_Id, Produtos.Nome, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Cfop_Id, SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, Produtos.NCM,Produtos.Unidade, NotasFiscaisXItens.QuantidadeFiscal, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor" & vbCrLf & _
                     "                       FROM          NotasFiscais INNER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                     "                                              Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id LEFT JOIN" & vbCrLf & _
                     "                                              SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                     "                       WHERE       ((year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") and (NotasFiscais.NossaEmissao = 'N') and (NotasFiscais.EntradaSaida_Id = 'E')   " & vbCrLf & _
                     "                                                                      And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "                                             )) AS consulta INNER JOIN" & vbCrLf & _
                     "                      NotasFiscaisXEncargos ON consulta.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                     "                      consulta.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND consulta.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                     "                      consulta.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                     "                      consulta.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND consulta.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf & _
                     "                      consulta.Produto_Id = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                     "                     AND      (consulta.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (consulta.Serie_Id <> '101' ) AND (consulta.Valor <> 0) " & vbCrLf & _
                     "   GROUP BY  consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, consulta.Operacao_Id, consulta.SubOperacoes_Id, consulta.NCM,consulta.Unidade, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor" & vbCrLf


            Dim nota As String = ""
            Dim item As Integer = 1
            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr
                    If nota = CStr(.Item("Nota_Id")) Then
                        item = item + 1
                    Else
                        item = 1
                    End If
                    strLinha = Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    nota = .Item("Nota_Id")
                    strLinha &= Funcoes.AlinharEsquerda(CDate(.Item("DataDaNota")).ToStrDate(), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Cliente_Id")), 14, " ")     'codigo participante
                    strLinha &= Funcoes.AlinharEsquerda(item, 3, " ") 'numero do item
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Produto_Id"), 20, " ") 'codigo produto
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.RemoveAllEnterKey(Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))), 45, " ")      'Nome produto
                    'strLinha &= Funcoes.AlinharEsquerda(.Item("Cfop_Id"), 4, " ")     'CFOP
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Cfop_Id"), 4, "0")     'CFOP
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Operacao_Id") & .Item("SubOperacoes_Id"), 6, " ")     'Operacao e subOperacao
                    strLinha &= Funcoes.AlinharEsquerda(.Item("NCM"), 8, " ")   'total descontos

                    If .Item("BaseDeCalculo") > 1 Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(1 * 1000, 0).Replace(".", ""), 17, " ")   'Quantidade
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(IIf(.Item("QuantidadeFiscal") = 0, 1, .Item("QuantidadeFiscal")) * 1000, 0).Replace(".", ""), 17, " ")   'Quantidade
                    End If

                    strLinha &= Funcoes.AlinharEsquerda(.Item("Unidade"), 3, " ")     ' numero nota

                    If .Item("BaseDeCalculo") > 1 Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(FormatNumber((.Item("Valor") / 1), 4).Replace(".", "") * 10000, 0).Replace(".", ""), 17, " ")   'Valor unitario
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(FormatNumber((.Item("Valor") / IIf(.Item("QuantidadeFiscal") = 0, 1, .Item("QuantidadeFiscal"))), 4).Replace(".", "") * 10000, 0).Replace(".", ""), 17, " ")   'Valor unitario
                    End If

                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("Valor") * 100, 0).Replace(".", ""), 17, " ")   'Valor unitario
                    strLinha &= Funcoes.AlinharEsquerda(0, 17, " ")  ' valor desconto item
                    If .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(1, 1, " ")  ' Indicador tributação ipi
                    ElseIf .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(2, 1, " ")  ' Indicador tributação ipi
                    ElseIf .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(1, 1, " ")  ' Indicador tributação ipi
                    ElseIf .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(2, 1, " ")  ' Indicador tributação ipi
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaIPI") * 100, 0).Replace(".", ""), 5, "0")   'Aliquota IPI
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BaseIPI") * 100, 0).Replace(".", ""), 17, " ")   'Base IPI
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorIPI") * 100, 0).Replace(".", ""), 17, " ")   'Valor IPI
                    strLinha &= Funcoes.AlinharEsquerda("0", 3, " ")   'Situacao tributaria estadual
                    If .Item("ValorICMS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda(1, 1, " ")  ' total ipi
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(2, 1, " ")  ' total ipi
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaICMS") * 100, 0).Replace(".", ""), 5, "0")   'Aliquota ICMS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BaseICMS") * 100, 0).Replace(".", ""), 17, " ")   'Base ICMS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorICMS") * 100, 0).Replace(".", ""), 17, " ")   'Valor ICMS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BaseICMSST") * 100, 0).Replace(".", ""), 17, " ")   'Base ICMS-ST
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorICMSST") * 100, 0).Replace(".", ""), 17, " ")   'Valor ICMS-ST
                    strLinha &= Funcoes.AlinharEsquerda("N", 1, " ") ' Indicador movimentacao

                    If .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("00", 2, " ")   'Situacao tributaria IPI
                    ElseIf .Item("EntradaSaida_Id") = "E" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("49", 2, " ")   'Situacao tributaria IPI
                    ElseIf .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("50", 2, " ")   'Situacao tributaria IPI
                    ElseIf .Item("EntradaSaida_Id") = "S" And .Item("ValorIPI") = 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("99", 2, " ")   'Situacao tributaria IPI
                    End If


                    Sql = "<tr>" & vbCrLf & _
                          "<td>" & empresa(0) & "-" & empresa(1) & "</td>" & vbCrLf & _
                          "<td>" & .Item("Cliente_Id") & "-" & .Item("EndCliente_Id") & "</td>" & vbCrLf & _
                          "<td>" & .Item("Nota_Id") & "</td>" & vbCrLf & _
                          "<td>" & .Item("Serie_Id") & "</td>" & vbCrLf & _
                          "<td>" & .Item("EntradaSaida_Id") & "</td>" & vbCrLf & _
                          "<td>" & CDate(.Item("DataDaNota")).ToStrDate() & "</td>" & vbCrLf & _
                          "<td>" & .Item("Operacao_Id") & "-" & .Item("SubOperacoes_Id") & "</td>" & vbCrLf & _
                          "<td>" & .Item("Cfop_Id") & "</td>" & vbCrLf & _
                          "<td>" & .Item("QuantidadeFiscal") & "</td>" & vbCrLf & _
                          "<td>" & .Item("Valor") & "</td>" & vbCrLf & _
                          "</tr>" & vbCrLf & vbCrLf

                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()

                Registros += 1
            Next
        Next

        Sql = "</table>"
        Try
            strmHtml = New StreamWriter(arquivo1, True)
            strmHtml.WriteLine(Sql)
            strmHtml.Close()
            'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo3 & "');", True)
        Catch ex As Exception
            strmHtml.Close()
            MsgBox(Me.Page, ex.Message.ToString)
        End Try

        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo491()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String

        Dim strSQL As String
        'Dim NomeArquivo As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)
        Dim Registros As Integer


        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ491.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)
            'strSQL = "SELECT DISTINCT DataAtualizacao, CodigoParticipante, CNPJCPF, InscricaoEstadual, NomeRazaoSocial, "
            'strSQL &= "Endereco, Bairro, Municipio, UF, Pais, Cep "
            'strSQL &= "FROM CadastroPessoasJuridicaFisica491"


            strSQL = "SELECT distinct    str((year(getdate()))-6) + '/01/01' as DataAtualizacao, Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Estado, Clientes.Nome, Clientes.Fantasia, " & vbCrLf & _
                     "              Clientes.Endereco, Clientes.Bairro, Clientes.Cep, Clientes.Cidade, Clientes.Inscricao, Clientes.Telefone, Clientes.Fax, Clientes.email, Clientes.Imagem, " & vbCrLf & _
                     "              Clientes.Reduzido, Clientes.Numero, Clientes.Complemento, Clientes.Situacao, " & vbCrLf & _
                     " Clientes.CodigoDoMunicipio, Clientes.Habilitacao, Clientes.EmailNFE, isnull(Clientes.Pais,' ') as Pais" & vbCrLf & _
                     " FROM         Clientes INNER JOIN" & vbCrLf & _
                     "             NotasFiscais ON Clientes.Cliente_Id = NotasFiscais.Cliente_Id AND Clientes.Endereco_Id = NotasFiscais.EndCliente_Id" & vbCrLf & _
                     "                     AND      (NotasFiscais.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (NotasFiscais.Serie_Id <> '101' ) " & vbCrLf & _
                     " where (year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") " & vbCrLf & _
                     " And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf





            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr
                    strLinha = Funcoes.AlinharEsquerda(Format(CDate(.Item("DataAtualizacao")), "ddMMyyyy"), 8, " ")
                    'strLinha = Funcoes.AlinharEsquerda(Format(CDate(TxtDataIni.Text), "ddMMyyyy"), 8, " ")
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Cliente_Id")), 14, " ")
                    strLinha &= IIf(.Item("Estado") = "EX", Funcoes.AlinharEsquerda("0", 14, "0"), Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Cliente_Id")), 14, " "))
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Inscricao")), 14, " ")
                    strLinha &= Funcoes.AlinharEsquerda(" ", 14, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nome"), 70, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Endereco"), 60, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Bairro"), 20, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Cidade"), 20, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Estado"), 2, " ")
                    strLinha &= Funcoes.AlinharEsquerda(IIf(.Item("Pais") = "1058", "", .Item("Pais")), 20, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Cep"), 8, " ")
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next
        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo494()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String
        Dim Registros As Integer
        Registros = 0

        Dim strSQL As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ494.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            'strSQL = "SELECT DISTINCT DataAtualizacao, CodigoNaturezaOperacao, Descricao "
            'strSQL &= "FROM TabelaNaturezaOperacao494"


            'strSQL = " SELECT distinct str((year(getdate()))-6) + '/01/01' as DataAtualizacao,SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao"
            'strSQL &= " FROM  SubOperacoes INNER JOIN"
            'strSQL &= " NotasFiscais ON  SubOperacoes.Operacao_Id = NotasFiscais.Operacao AND "
            'strSQL &= " SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao"
            'strSQL &= "  AND      (NotasFiscais.DataDaNota <> STR(YEAR(GETDATE())) + '/01/01') AND (NotasFiscais.Serie_Id <> '101' ) "
            'strSQL &= " where NotasFiscais.Movimento between '" & Format(CDate(TxtDataIni.Text), "yyyy/MM/dd") & "' AND '" & Format(CDate(TxtDataFin.Text), "yyyy/MM/dd") & "' "
            'strSQL &= " And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1)

            strSQL = "  SELECT " & vbCrLf & _
                     "   distinct str((year(getdate()))-6) + '/01/01' as DataAtualizacao, " & vbCrLf & _
                     "   isnull(SubOperacoes.Operacao_Id,1) as Operacao_Id, " & vbCrLf & _
                     "   isnull(SubOperacoes.SubOperacoes_Id,1) as SubOperacoes_Id, " & vbCrLf & _
                     "   SubOperacoes.Descricao " & vbCrLf & _
                     "   FROM NotasFiscais " & vbCrLf & _
                     "   Left JOIN SubOperacoes ON  case when isnull(NotasFiscais.Operacao,1) = 0 then 1 else isnull(NotasFiscais.Operacao,1) end = SubOperacoes.Operacao_Id " & vbCrLf & _
                     "   AND  case when isnull(NotasFiscais.SubOperacao,1) = 0 then 1 else isnull(NotasFiscais.SubOperacao,1) end = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                     " where (year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") " & vbCrLf & _
                     " And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "   AND  (NotasFiscais.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') " & vbCrLf & _
                     "   AND (NotasFiscais.Serie_Id <> '101' ) " & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr
                    strLinha = Funcoes.AlinharEsquerda(Format(CDate(.Item("DataAtualizacao")), "ddMMyyyy"), 8, " ")
                    'strLinha = Funcoes.AlinharEsquerda(Format(CDate(TxtDataIni.Text), "ddMMyyyy"), 8, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Operacao_Id") & .Item("SubOperacoes_Id"), 6, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Descricao"), 45, " ")
                End With
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next

        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo495()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String
        Dim Registros As Integer
        Registros = 0

        Dim strSQL As String

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ495.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            'strSQL = "SELECT DISTINCT DataAtualizacao, CodigoMercadoria, Descricao "
            'strSQL &= "FROM TabelaMercadoriaServicos495"

            strSQL = " SELECT  distinct   str((year(getdate()))-6) + '/01/01' as DataAtualizacao, Produtos.Produto_Id,Produtos.Nome" & vbCrLf & _
                     "   FROM Produtos INNER JOIN" & vbCrLf & _
                     "   NotasFiscaisXItens ON  Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id INNER JOIN" & vbCrLf & _
                     "   NotasFiscais ON  NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id AND " & vbCrLf & _
                     "   NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id AND " & vbCrLf & _
                     "   NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id AND " & vbCrLf & _
                     "   NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id" & vbCrLf & _
                     "                     AND      (NotasFiscais.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (NotasFiscais.Serie_Id <> '101' ) " & vbCrLf & _
                     " where (year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") " & vbCrLf & _
                     " And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr
                    strLinha = Funcoes.AlinharEsquerda(Format(CDate(.Item("DataAtualizacao")), "ddMMyyyy"), 8, " ")
                    'strLinha = Funcoes.AlinharEsquerda(Format(CDate(TxtDataIni.Text), "ddMMyyyy"), 8, " ")
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Produto_Id"), 20, " ")
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.RemoveAllEnterKey(.Item("Nome")), 45, " ")
                End With
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next

        Texto = Registros.ToString

        Return Texto

    End Function

    Public Function Arquivo4101()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String
        Dim Registros As Integer
        Dim strSQL As String

        Registros = 0
        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ4101.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            strSQL = "   SELECT              consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id, consulta.Movimento,consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, isnull(consulta.Operacao_Id,51) as Operacao_Id, isnull(consulta.SubOperacoes_Id,1) as SubOperacoes_Id, consulta.NCM, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.SituacaoTributaria ELSE 0 END) AS SituacaoTributariaPIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BasePIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaPIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorPIS," & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.SituacaoTributaria ELSE 0 END) AS SituacaoTributariaCOFINS, " & vbCrLf & _
                     "					     SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseCOFINS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaCOFINS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorCOFINS " & vbCrLf & _
                     "   FROM         (SELECT     NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                              NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id,NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscaisXItens.Produto_Id, Produtos.Nome, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Cfop_Id, SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, Produtos.NCM, NotasFiscaisXItens.QuantidadeFiscal, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor" & vbCrLf & _
                     "                       FROM          NotasFiscais INNER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                     "                                              Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id LEFT JOIN" & vbCrLf & _
                     "                                              SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                     "                       WHERE      ((year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") and (NotasFiscais.NossaEmissao = 'S') and NotasFiscais.EntradaSaida_Id = 'S' " & vbCrLf & _
                     "                                   And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "                                             )) AS consulta INNER JOIN" & vbCrLf & _
                     "                      NotasFiscaisXEncargos ON consulta.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                     "                      consulta.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND consulta.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                     "                      consulta.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                     "                      consulta.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND consulta.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf & _
                     "                      consulta.Produto_Id = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                     "                     AND      (consulta.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (consulta.Serie_Id <> '101' ) AND (consulta.Valor <> 0) " & vbCrLf & _
                     "   GROUP BY consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id,consulta.Movimento, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, consulta.Operacao_Id, consulta.SubOperacoes_Id, consulta.NCM, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor" & vbCrLf

            Dim nota As String = ""
            Dim item As Integer = 1
            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr

                    If nota = CStr(.Item("Nota_Id")) Then
                        item = item + 1
                    Else
                        item = 1
                    End If
                    strLinha = Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    nota = .Item("Nota_Id")
                    strLinha &= Funcoes.AlinharEsquerda(Format(.Item("DataDaNota"), "ddMMyyyy"), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(item, 3, " ") 'numero do item
                    If .Item("ValorPIS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("49", 2, " ") 'situacao tributaria PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda("07", 2, " ") 'situacao tributaria PIS
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaPIS") * 10000, 0).Replace(".", ""), 8, "0")   'Aliquota PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BasePIS") * 1000, 0).Replace(".", ""), 17, " ")   'Base PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    If .Item("ValorCOFINS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("49", 2, " ") 'situacao tributaria COFINS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda("07", 2, " ") 'situacao tributaria COFINS
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaCOFINS") * 10000, 0).Replace(".", ""), 8, "0")   'Aliquota PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BaseCOFINS") * 1000, 0).Replace(".", ""), 17, " ")   'Base PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(Format(.Item("Movimento"), "ddMMyyyy"), 8, " ") 'data emissão documento
                End With
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next
        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo4104()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String
        Dim Registros As Integer
        Dim strSQL As String

        Registros = 0

        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ4104.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            strSQL = "   SELECT      consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id, consulta.Movimento,consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, isnull(consulta.Operacao_Id,1) as Operacao_Id, isnull(consulta.SubOperacoes_Id,1) as SubOperacoes_Id, consulta.NCM, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.SituacaoTributaria ELSE 0 END) AS SituacaoTributariaPIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BasePIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaPIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorPIS," & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.SituacaoTributaria ELSE 0 END) AS SituacaoTributariaCOFINS, " & vbCrLf & _
                     "					     SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseCOFINS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaCOFINS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorCOFINS " & vbCrLf & _
                     "   FROM         (SELECT      NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                              NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id,NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscaisXItens.Produto_Id, Produtos.Nome, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Cfop_Id, SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, Produtos.NCM, NotasFiscaisXItens.QuantidadeFiscal, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor" & vbCrLf & _
                     "                       FROM          NotasFiscais INNER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                     "                                              Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id LEFT JOIN" & vbCrLf & _
                     "                                              SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                     "                       WHERE      ((year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") and (NotasFiscais.NossaEmissao = 'S') and NotasFiscais.EntradaSaida_Id = 'E' " & vbCrLf & _
                     "                                  And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "                                             )) AS consulta INNER JOIN" & vbCrLf & _
                     "                      NotasFiscaisXEncargos ON consulta.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                     "                      consulta.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND consulta.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                     "                      consulta.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                     "                      consulta.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND consulta.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf & _
                     "                      consulta.Produto_Id = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                     "                     AND      (consulta.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (consulta.Serie_Id <> '101' ) AND (consulta.Valor <> 0) " & vbCrLf & _
                     "   GROUP BY  consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id,consulta.Movimento, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, consulta.Operacao_Id, consulta.SubOperacoes_Id, consulta.NCM, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor " & vbCrLf

            Dim nota As String = ""
            Dim item As Integer = 1
            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr

                    If nota = CStr(.Item("Nota_Id")) Then
                        item = item + 1
                    Else
                        item = 1
                    End If
                    strLinha = Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    nota = .Item("Nota_Id")
                    strLinha &= Funcoes.AlinharEsquerda(Format(.Item("DataDaNota"), "ddMMyyyy"), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(item, 3, " ") 'numero do item
                    If .Item("ValorPIS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("50", 2, " ") 'situacao tributaria PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda("07", 2, " ") 'situacao tributaria PIS 
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaPIS") * 10000, 0).Replace(".", ""), 8, "0")   'Aliquota PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BasePIS") * 1000, 0).Replace(".", ""), 17, " ")   'Base PIS
                    If Mid(.Item("Cfop_Id"), 1, 1) = 3 Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    If (Mid(.Item("Cfop_Id"), 1, 1) = 1) Or (Mid(.Item("Cfop_Id"), 1, 1) = 2) Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS

                    If .Item("ValorCOFINS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("50", 2, " ") 'situacao tributaria COFINS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda("07", 2, " ") 'situacao tributaria COFINS 
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaCOFINS") * 10000, 0).Replace(".", ""), 8, "0")   'Aliquota PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BaseCOFINS") * 1000, 0).Replace(".", ""), 17, " ")   'Base PIS

                    If Mid(.Item("Cfop_Id"), 1, 1) = 3 Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    If (Mid(.Item("Cfop_Id"), 1, 1) = 1) Or (Mid(.Item("Cfop_Id"), 1, 1) = 2) Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(Format(.Item("Movimento"), "ddMMyyyy"), 8, " ") 'data emissão documento
                End With
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next
        Texto = Registros.ToString

        Return Texto
    End Function

    Public Function Arquivo4105()
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim Texto As String
        Dim Registros As Integer
        Registros = 0

        Dim strSQL As String
        CriarPastaEmpresa()
        CriarAno(ddlAno.SelectedValue)

        For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
            Dim arquivo As String = CriarMes(ddlAno.SelectedValue, Mes) & "/ARQ4105.TXT"
            Dim strm As StreamWriter
            Dim strLinha As String

            Registros = 0

            If Dir(arquivo).Length > 0 Then Kill(arquivo)


            strSQL = "   SELECT      consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id, consulta.Movimento,consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, isnull(consulta.Operacao_Id,1) as Operacao_Id, isnull(consulta.SubOperacoes_Id,1) as SubOperacoes_Id, consulta.NCM, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.SituacaoTributaria ELSE 0 END) AS SituacaoTributariaPIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BasePIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaPIS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorPIS," & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.SituacaoTributaria ELSE 0 END) AS SituacaoTributariaCOFINS, " & vbCrLf & _
                     "					     SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) AS BaseCOFINS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) AS AliquotaCOFINS, " & vbCrLf & _
                     "                      SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorCOFINS " & vbCrLf & _
                     "   FROM         (SELECT      NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
                     "                                              NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id,NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscaisXItens.Produto_Id, Produtos.Nome, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Cfop_Id, SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, Produtos.NCM, NotasFiscaisXItens.QuantidadeFiscal, " & vbCrLf & _
                     "                                              NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor" & vbCrLf & _
                     "                       FROM          NotasFiscais INNER JOIN" & vbCrLf & _
                     "                                              NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                     "                                              Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id LEFT JOIN" & vbCrLf & _
                     "                                              SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf & _
                     "                                              NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                     "                       WHERE      ((year(NotasFiscais.Movimento) = " & ddlAno.SelectedValue & " and Month(NotasFiscais.Movimento) = " & Mes & ") and (NotasFiscais.NossaEmissao = 'N') and NotasFiscais.EntradaSaida_Id = 'E' " & vbCrLf & _
                     "                                  And NotasFiscais.Empresa_Id = '" & empresa(0) & "' and NotasFiscais.EndEmpresa_id = " & empresa(1) & vbCrLf & _
                     "                                             )) AS consulta INNER JOIN" & vbCrLf & _
                     "                      NotasFiscaisXEncargos ON consulta.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                     "                      consulta.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND consulta.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                     "                      consulta.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND consulta.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                     "                      consulta.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND consulta.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf & _
                     "                      consulta.Produto_Id = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                     "                     AND      (consulta.DataDaNota <> '" & ddlAno.SelectedValue & "/01/01') AND (consulta.Serie_Id <> '101' ) AND (consulta.Valor <> 0) " & vbCrLf & _
                     "   GROUP BY  consulta.Empresa_Id, consulta.EndEmpresa_Id, consulta.Cliente_Id, consulta.EndCliente_Id, consulta.EntradaSaida_Id, consulta.Serie_Id, " & vbCrLf & _
                     "                      consulta.Nota_Id,consulta.Movimento, consulta.DataDaNota, consulta.Produto_Id, consulta.Nome, consulta.Cfop_Id, consulta.Operacao_Id, consulta.SubOperacoes_Id, consulta.NCM, " & vbCrLf & _
                     "                      consulta.QuantidadeFiscal, consulta.Unitario, consulta.Valor" & vbCrLf

            Dim nota As String = ""
            Dim item As Integer = 1
            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Clientes").Tables(0).Rows
                With Dr

                    If nota = CStr(.Item("Nota_Id")) Then
                        item = item + 1
                    Else
                        item = 1
                    End If
                    strLinha = Funcoes.AlinharEsquerda("01", 2, " ") 'modelo documento
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Serie_Id"), 5, " ") ' serie da nota
                    strLinha &= Funcoes.AlinharEsquerda(.Item("Nota_Id"), 9, " ")     ' numero nota
                    nota = .Item("Nota_Id")
                    strLinha &= Funcoes.AlinharEsquerda(Format(.Item("DataDaNota"), "ddMMyyyy"), 8, " ") 'data emissão documento
                    strLinha &= Funcoes.AlinharEsquerda(Funcoes.EliminarCaracteresEspeciais(.Item("Cliente_Id")), 14, " ") 'Codigo participante
                    strLinha &= Funcoes.AlinharEsquerda(item, 3, " ") 'numero do item
                    If .Item("ValorPIS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("50", 2, " ") 'situacao tributaria PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda("98", 2, " ") 'situacao tributaria PIS 
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaPIS") * 10000, 0).Replace(".", ""), 8, "0")   'Aliquota PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BasePIS") * 1000, 0).Replace(".", ""), 17, " ")   'Base PIS
                    If Mid(.Item("Cfop_Id"), 1, 1) = 3 Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    If (Mid(.Item("Cfop_Id"), 1, 1) = 1) Or (Mid(.Item("Cfop_Id"), 1, 1) = 2) Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorPIS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS

                    If .Item("ValorCOFINS") > 0 Then
                        strLinha &= Funcoes.AlinharEsquerda("50", 2, " ") 'situacao tributaria COFINS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda("98", 2, " ") 'situacao tributaria COFINS 
                    End If
                    strLinha &= Funcoes.AlinharDireita(FormatNumber(.Item("AliquotaCOFINS") * 10000, 0).Replace(".", ""), 8, "0")   'Aliquota PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("BaseCOFINS") * 1000, 0).Replace(".", ""), 17, " ")   'Base PIS 

                    If Mid(.Item("Cfop_Id"), 1, 1) = 3 Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    If (Mid(.Item("Cfop_Id"), 1, 1) = 1) Or (Mid(.Item("Cfop_Id"), 1, 1) = 2) Then
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    Else
                        strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    End If
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber("0", 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(FormatNumber(.Item("ValorCOFINS") * 100, 0).Replace(".", ""), 17, " ")   'Valor PIS
                    strLinha &= Funcoes.AlinharEsquerda(Format(.Item("Movimento"), "ddMMyyyy"), 8, " ") 'data emissão documento
                End With
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(strLinha)
                strm.Close()
                Registros += 1
            Next
        Next
        Texto = Registros.ToString

        Return Texto
    End Function

#Region "Botões"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo431() = True Then CheckBox1.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button2_Click2(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo432() Then CheckBox2.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button3_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo433() Then CheckBox3.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button4_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo434() Then CheckBox4.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button5_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                ' If  Arquivo461()Then CheckBox5.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button6_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo491() Then CheckBox6.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button7_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo494() Then CheckBox7.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button8_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo495() Then CheckBox8.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button9_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo4101() Then CheckBox9.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button10_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo4104() Then CheckBox10.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button11_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidaDados() Then
                If Arquivo4105() Then CheckBox11.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Function ValidaDados()
        'If Trim(TxtDataIni.Text) = "" Or Trim(TxtDataFin.Text) = "" Then
        '    MsgBox(Me.Page, "Informe o Periodo.")
        '    Return False
        'End If

        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione uma empresa, antes de processar.")
            Return False
        End If

        Return True
    End Function

#End Region

    Public Function CriarPastaEmpresa() As String
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")

        If Not Directory.Exists(Server.MapPath("Sva/" & empresa(0))) Then
            Directory.CreateDirectory(Server.MapPath("Sva/" & empresa(0)))
        End If

        Return Server.MapPath("Sva/" & empresa(0))
    End Function

    Public Function CriarAno(ByVal pAno As Integer) As String
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")

        If Not Directory.Exists(Server.MapPath("Sva/" & empresa(0) & "/" & pAno)) Then
            Directory.CreateDirectory(Server.MapPath("Sva/" & empresa(0) & "/" & pAno))
        End If

        Return Server.MapPath("Sva/" & empresa(0) & "/" & pAno)
    End Function

    Public Function CriarMes(ByVal pAno As Integer, ByVal pMes As Integer) As String
        Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")

        If Not Directory.Exists(Server.MapPath("Sva/" & empresa(0) & "/" & pAno & "/" & pMes)) Then
            Directory.CreateDirectory(Server.MapPath("Sva/" & empresa(0) & "/" & pAno & "/" & pMes))
        End If

        Return Server.MapPath("Sva/" & empresa(0) & "/" & pAno & "/" & pMes)
    End Function

End Class