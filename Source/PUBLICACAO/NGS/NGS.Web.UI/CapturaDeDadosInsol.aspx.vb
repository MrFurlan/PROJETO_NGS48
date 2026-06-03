Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CapturaDeDadosInsol
    Inherits BasePage

    Dim Array As New ArrayList
    Dim Sql As String
    Dim Sqla As String
    Dim Circular As String = "S"
    Dim Empresa() As String
    Dim i As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/ApuracaoDeCustos.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("CapturaDeDados", "ACESSAR") Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.EmpresasConsolidadas, "", False)
                ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/ApuracaoDeCustos.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Sub Limpar()
        DdlAno.SelectedValue = IIf(Now.Month = 1, Now.Year - 1, Now.Year)
        DdlMesInicial.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)
        DdlMesFinal.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)
        txtCiclos.Text = "5"
        ChkResultado.Items.Clear()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Function AjusteComplementacaoDePreco(ByVal pMes As Integer) As Boolean
        Dim sql As String
        Dim sqls As New ArrayList
        sql = " Select  Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id,  Quantidade, ValorDoProduto" & vbCrLf & _
              "   Into #Temp " & vbCrLf & _
              "   from ApuracaoDeCustos " & vbCrLf & _
              "  where CodigoDeCusto_id = 111" & vbCrLf & _
              "    and Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
              "    and Mes_Id           = " & pMes & vbCrLf


        sql &= " Update ApuracaoDeCustos set" & vbCrLf & _
               "    Quantidade     = AC.Quantidade  - T.Quantidade" & vbCrLf & _
               " from #Temp T" & vbCrLf & _
               " Inner Join ApuracaoDeCustos AC" & vbCrLf & _
               "    on T.Empresa_Id            = AC.Empresa_Id" & vbCrLf & _
               "   and T.EndEmpresa_Id         = AC.EndEmpresa_Id" & vbCrLf & _
               "   and T.Deposito_Id           = AC.Deposito_Id " & vbCrLf & _
               "   and T.EndDeposito_Id        = AC.EndDeposito_Id" & vbCrLf & _
               "   and T.Ano_Id                = AC.Ano_Id" & vbCrLf & _
               "   and T.Mes_Id                = AC.Mes_Id " & vbCrLf & _
               "   and T.Produto_Id            = AC.Produto_Id" & vbCrLf & _
               "   and T.EmpresaDestino_Id     = AC.EmpresaDestino_Id" & vbCrLf & _
               "   and T.EndEmpresaDestino_Id  = AC.EndEmpresaDestino_Id" & vbCrLf & _
               "   and T.DepositoDestino_Id    = AC.DepositoDestino_Id " & vbCrLf & _
               "   and T.EndDepositoDestino_Id = AC.EndDepositoDestino_Id" & vbCrLf & _
               "   and T.ProdutoDerivado_Id    = AC.ProdutoDerivado_Id" & vbCrLf & _
               "   and AC.CodigoDeCusto_Id     = 110" & vbCrLf & _
               " where AC.Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
               "   and AC.Mes_Id            = " & pMes
        sqls.Add(sql)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function AvaliarEstoqueAFixar(ByVal pMes As Integer) As Boolean
        Dim sql As String
        Dim sqls As New ArrayList
        Dim Ano As Integer = DdlAno.SelectedValue

        '"              then  (((ac.Quantidade*P.BaseDeCalculo) * PM.ValorOficial) / PM.basedecalculo) - AC.ValorDoProduto " & vbCrLf & _
        sql = " SELECT AC.Empresa_Id, AC.EndEmpresa_Id," & vbCrLf & _
              "        AC.Deposito_Id, AC.EndDeposito_Id," & vbCrLf & _
              "        AC.Ano_Id, AC.Mes_Id," & vbCrLf & _
              "        AC.Produto_Id, AC.CodigoDeCusto_Id, " & vbCrLf & _
              "        AC.EmpresaDestino_Id, AC.EndEmpresaDestino_Id, " & vbCrLf & _
              "        AC.DepositoDestino_Id, AC.EndDepositoDestino_Id," & vbCrLf & _
              "        AC.ProdutoDerivado_Id, AC.Etapa," & vbCrLf & _
              "        AC.Quantidade," & vbCrLf & _
              "        AC.ValorDoFrete, AC.ValorAuxiliar," & vbCrLf & _
              "        AC.ProdutoDestino, AC.CodigoDestino, AC.Reduzido, " & vbCrLf & _
              "        case when AC.Quantidade >= 0" & vbCrLf & _
              "              then  ((ac.Quantidade * PM.ValorOficial) / PM.basedecalculo) - AC.ValorDoProduto " & vbCrLf & _
              "              else     ac.Quantidade * (MesAnterior.ValorDoProduto / MesAnterior.Quantidade) -  AC.ValorDoProduto " & vbCrLf & _
              "        end as ValorDoProduto " & vbCrLf & _
              "  FROM ApuracaoDeCustos AC" & vbCrLf & _
              " inner join TabelaDePrecosDeMercado PM" & vbCrLf & _
              "    on AC.Empresa_Id    = PM.Empresa_Id" & vbCrLf & _
              "   And AC.EndEmpresa_Id = Pm.EndEmpresa_Id" & vbCrLf & _
              "   And AC.Deposito_Id   = PM.Deposito_Id" & vbCrLf & _
              "   And AC.EndEmpresa_Id = Pm.EndDeposito_id" & vbCrLf & _
              "   And AC.Produto_Id    = PM.Produto_Id" & vbCrLf & _
              "   And PM.Data_Id       = '" & DdlAno.SelectedValue & "-" & pMes & "-" & Date.DaysInMonth(DdlAno.SelectedValue, pMes) & "'" & vbCrLf & _
              "  left Join ( Select * from ApuracaoDeCustos) MesAnterior" & vbCrLf & _
              "    on AC.Empresa_Id            = MesAnterior.Empresa_Id" & vbCrLf & _
              "   and AC.EndEmpresa_Id         = MesAnterior.EndEmpresa_Id" & vbCrLf & _
              "   and AC.Deposito_Id           = MesAnterior.Deposito_Id " & vbCrLf & _
              "   and AC.EndDeposito_Id        = MesAnterior.EndDeposito_Id" & vbCrLf & _
              "   and case when AC.Mes_Id = 1 then AC.Ano_Id -1 else Ac.Ano_id end = MesAnterior.Ano_Id" & vbCrLf & _
              "   and case when AC.Mes_Id = 1 then 12           else Ac.Mes_Id end = MesAnterior.Mes_Id " & vbCrLf & _
              "   and AC.Produto_Id            = MesAnterior.Produto_Id" & vbCrLf & _
              "   and AC.EmpresaDestino_Id     = MesAnterior.EmpresaDestino_Id" & vbCrLf & _
              "   and AC.EndEmpresaDestino_Id  = MesAnterior.EndEmpresaDestino_Id" & vbCrLf & _
              "   and AC.DepositoDestino_Id    = MesAnterior.DepositoDestino_Id " & vbCrLf & _
              "   and AC.EndDepositoDestino_Id = MesAnterior.EndDepositoDestino_Id" & vbCrLf & _
              "   and AC.ProdutoDerivado_Id    = MesAnterior.ProdutoDerivado_Id" & vbCrLf & _
              "   and AC.CodigoDeCusto_Id      = MesAnterior.CodigoDeCusto_Id" & vbCrLf & _
              " inner join Produtos P" & vbCrLf & _
              "    on P.Produto_Id = AC.Produto_Id" & vbCrLf & _
              "  WHERE AC.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "    AND AC.Ano_Id = " & Ano & vbCrLf & _
              "    AND AC.Mes_Id = " & pMes & vbCrLf & _
              "    AND AC.CodigoDeCusto_Id = 110" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            sql = " Declare" & vbCrLf & _
                  " @Exist as varchar(1)" & vbCrLf & _
                  " set @Exist = (select case " & vbCrLf & _
                  "                        when exists (" & vbCrLf & _
                  "                                      select 1 " & vbCrLf & _
                  "                                        from ApuracaoDeCustos " & vbCrLf & _
                  "                                       Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "                                         And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "                                         And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "                                         And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "                                         And Mes_Id                = " & pMes & vbCrLf & _
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                         And CodigoDeCusto_Id      = 113" & vbCrLf & _
                  "                                         And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "                                         And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "                                         And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  "                                    )" & vbCrLf & _
                  "                           then 'S'" & vbCrLf & _
                  "                           else 'N'" & vbCrLf & _
                  "                       end) ;" & vbCrLf & _
                  " if @Exist = 'N' " & vbCrLf & _
                  " INSERT INTO ApuracaoDeCustos (" & vbCrLf & _
                  "                                Empresa_Id" & vbCrLf & _
                  "                               ,EndEmpresa_Id" & vbCrLf & _
                  "                               ,Deposito_Id" & vbCrLf & _
                  "                               ,EndDeposito_Id" & vbCrLf & _
                  "                               ,Ano_Id" & vbCrLf & _
                  "                               ,Mes_Id" & vbCrLf & _
                  "                               ,Produto_Id" & vbCrLf & _
                  "                               ,CodigoDeCusto_Id" & vbCrLf & _
                  "                               ,EmpresaDestino_Id" & vbCrLf & _
                  "                               ,EndEmpresaDestino_Id" & vbCrLf & _
                  "                               ,DepositoDestino_Id" & vbCrLf & _
                  "                               ,EndDepositoDestino_Id" & vbCrLf & _
                  "                               ,ProdutoDerivado_Id" & vbCrLf & _
                  "                               ,Quantidade" & vbCrLf & _
                  "                               ,ValorDoProduto" & vbCrLf & _
                  "                               ,ValorDoFrete" & vbCrLf & _
                  "                               ,ProdutoDestino" & vbCrLf & _
                  "                               ,CodigoDestino)" & vbCrLf & _
                  " VALUES( '" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "        ,'" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndDeposito_Id") & vbCrLf & _
                  "        , " & DdlAno.SelectedValue & vbCrLf & _
                  "        , " & pMes & vbCrLf & _
                  "        ,'" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "        , 113" & vbCrLf & _
                  "        ,'" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "        ,'" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "        ,'" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  "        , 0 " & vbCrLf & _
                  "        , " & Dr("ValorDoProduto").ToString.Replace(",", ".") & vbCrLf & _
                  "        , " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  "        ,'" & Dr("ProdutoDestino") & "'" & vbCrLf & _
                  "        ,0" & vbCrLf & _
                  "       )" & vbCrLf & _
                  " Else" & vbCrLf & _
                  "  Update ApuracaoDeCustos set " & vbCrLf & _
                  "     Quantidade     = 0" & vbCrLf & _
                  "    ,ValorDoProduto = " & Dr("ValorDoProduto").ToString.Replace(",", ".") & vbCrLf & _
                  "    ,ValorDoFrete   = " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  "  Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "    And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "    And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "    And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "    And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "    And Mes_Id                = " & pMes & vbCrLf & _
                  "    And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "    And CodigoDeCusto_Id      = 113" & vbCrLf & _
                  "    And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "    And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "    And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "    And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "    And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf

            Array.Add(sql)
            If Not Banco.GravaBanco(Array) Then Return False
        Next
        Return True
    End Function

    Public Function EstornaAfixarMesAnterior(ByVal pMes As Integer) As Boolean
        Dim erros As Integer = 0
        i = 0

        Dim Ano As Integer = DdlAno.SelectedValue
        Dim MesAnterior As Integer = pMes

        If pMes = 1 Then
            MesAnterior = 12
            Ano = Ano - 1
        Else
            MesAnterior = pMes - 1
        End If

        Sql = " SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, " & vbCrLf & _
              "        DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Etapa, Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, CodigoDestino, Reduzido" & vbCrLf & _
              "   FROM ApuracaoDeCustos" & vbCrLf & _
              "  WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "    AND Ano_Id = " & Ano & vbCrLf & _
              "    AND Mes_Id = " & MesAnterior & vbCrLf & _
              "    AND CodigoDeCusto_Id = 113" & vbCrLf & _
              "    AND ValorDoProduto <> 0"

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
                  "                                         And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "                                         And Mes_Id                = " & pMes & vbCrLf & _
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                         And CodigoDeCusto_Id      = 112" & vbCrLf & _
                  "                                         And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "                                         And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "                                         And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  "                                    )" & vbCrLf & _
                  "                           then 'S'" & vbCrLf & _
                  "                           else 'N'" & vbCrLf & _
                  "                       end) ;" & vbCrLf & _
                  " if @Exist = 'N' " & vbCrLf & _
                  " INSERT INTO ApuracaoDeCustos (" & vbCrLf & _
                  "                                Empresa_Id" & vbCrLf & _
                  "                               ,EndEmpresa_Id" & vbCrLf & _
                  "                               ,Deposito_Id" & vbCrLf & _
                  "                               ,EndDeposito_Id" & vbCrLf & _
                  "                               ,Ano_Id" & vbCrLf & _
                  "                               ,Mes_Id" & vbCrLf & _
                  "                               ,Produto_Id" & vbCrLf & _
                  "                               ,CodigoDeCusto_Id" & vbCrLf & _
                  "                               ,EmpresaDestino_Id" & vbCrLf & _
                  "                               ,EndEmpresaDestino_Id" & vbCrLf & _
                  "                               ,DepositoDestino_Id" & vbCrLf & _
                  "                               ,EndDepositoDestino_Id" & vbCrLf & _
                  "                               ,ProdutoDerivado_Id" & vbCrLf & _
                  "                               ,Quantidade" & vbCrLf & _
                  "                               ,ValorDoProduto" & vbCrLf & _
                  "                               ,ValorDoFrete" & vbCrLf & _
                  "                               ,ProdutoDestino" & vbCrLf & _
                  "                               ,CodigoDestino)" & vbCrLf & _
                  " VALUES( '" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "        ,'" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndDeposito_Id") & vbCrLf & _
                  "        , " & DdlAno.SelectedValue & vbCrLf & _
                  "        , " & pMes & vbCrLf & _
                  "        ,'" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "        , 112" & vbCrLf & _
                  "        ,'" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "        ,'" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "        ,'" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  "        , " & Str(Dr("Quantidade")) & vbCrLf & _
                  "        , " & Dr("ValorDoProduto").ToString.Replace(",", ".") & vbCrLf & _
                  "        , " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  "        ,'" & Dr("ProdutoDestino") & "'" & vbCrLf & _
                  "        ,0" & vbCrLf & _
                  "       )" & vbCrLf & _
                  " Else" & vbCrLf & _
                  "  Update ApuracaoDeCustos set " & vbCrLf & _
                  "     Quantidade     = " & Str(Dr("Quantidade")) & vbCrLf & _
                  "    ,ValorDoProduto = " & Dr("ValorDoProduto").ToString.Replace(",", ".") & vbCrLf & _
                  "    ,ValorDoFrete   = " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  "  Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "    And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "    And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "    And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "    And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "    And Mes_Id                = " & pMes & vbCrLf & _
                  "    And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "    And CodigoDeCusto_Id      = 112" & vbCrLf & _
                  "    And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "    And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "    And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "    And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "    And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf

            Array.Add(Sql)
            If Not Banco.GravaBanco(Array) Then Return False
        Next
        Return True
    End Function

    Public Function Transferencias(ByVal pMes As Integer) As Boolean
        Dim erros As Integer = 0
        i = 0


        Sql = "SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito AS Deposito_Id, NotasFiscaisXItens.EndDeposito AS EndDeposito_Id," & vbCrLf & _
              "       NotasFiscais.Cliente_Id AS EmpresaDestino_Id, NotasFiscais.EndCliente_Id AS EndEmpresaDestino_Id, NotasFiscais.Cliente_Id AS DepositoDestino_Id," & vbCrLf & _
              "       NotasFiscais.EndCliente_Id AS EndDepositoDestino_Id," & vbCrLf & _
              "       case" & vbCrLf & _
              "         when NotasFiscaisXItens.Produto_Id = '101010003' then '101010001'" & vbCrLf & _
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

        If CInt(txtDia.Text) > 0 Then
            Sql &= "   AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue & vbCrLf & _
                   "   AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes & vbCrLf & _
                   "   AND (isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  <= '" & DdlAno.SelectedValue & "-" & pMes.ToString("00") & "-" & CInt(txtDia.Text).ToString("00") & "'" & vbCrLf
        Else
            Sql &= "   AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue & vbCrLf & _
                   "   AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes & vbCrLf
        End If

        Sql &= "   AND (SubOperacoes.ApuracaoDeCustosContraPartida > 0)" & vbCrLf & _
               "   And  NotasFiscais.Situacao  = 1" & vbCrLf & _
               "   And  GruposDeEstoques.custo = 1" & vbCrLf & _
               "   And  PlanoDeCustos.Classe   = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf & _
               " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito," & vbCrLf & _
               "          NotasFiscaisXItens.EndDeposito, YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)), MONTH(ISNULL(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), " & vbCrLf & _
               "          case" & vbCrLf & _
               "            when NotasFiscaisXItens.Produto_Id = '101010003' then '101010001'" & vbCrLf & _
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

            Array.Add(Sql)
            If Not Banco.GravaBanco(Array) Then Return False
        Next
        Return True
    End Function

    Public Function Fretes(ByVal pMes As Integer) As Boolean
        Dim erros As Integer = 0
        i = 0

        Sql = " SELECT  Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id," & vbCrLf & _
              "        DepositoDestino_Id as EmpresaDestino_Id, EndDepositoDestino_Id as EndEmpresaDestino_Id," & vbCrLf & _
              "        DepositoDestino_Id, EndDepositoDestino_Id,  Produto_Id," & vbCrLf & _
              "        Mes_Id, Ano_Id, Placus_Id,  ApuracaodeCustosContraPartida," & vbCrLf & _
              "        Sum(ValorDoFrete - ICMS - PIs - COFINS) as ValorDoFrete" & vbCrLf & _
              " From (" & vbCrLf & _
              "       SELECT CTRC.Empresa_Id, CTRC.EndEmpresa_Id, NotasFiscais.Deposito AS Deposito_Id, NotasFiscais.EndDeposito AS EndDeposito_Id," & vbCrLf & _
              "              Case When SubOperacoes.Classe IN ('" & eClassesOperacoes.COMPRAS.ToString & "') Then '' else NotasFiscais.Cliente_Id    End AS DepositoDestino_Id," & vbCrLf & _
              "              Case When SubOperacoes.Classe IN ('" & eClassesOperacoes.COMPRAS.ToString & "') Then 0  else NotasFiscais.EndCliente_Id End AS EndDepositoDestino_Id," & vbCrLf & _
              "              NotasFiscaisXItens.Produto_Id, MONTH(NotasFiscais.Movimento) AS Mes_Id," & vbCrLf & _
              "              YEAR(NotasFiscais.Movimento) AS Ano_Id, SubOperacoes.ApuracaoDeCustos AS Placus_Id,  SubOperacoes.ApuracaodeCustosContraPartida," & vbCrLf & _
              "              Sum(Case When CTRCXEncargos.Encargo_Id = 'PRODUTO' THEN CTRCXItens.Valor else 0 End) AS ValorDoFrete," & vbCrLf & _
              "              Sum(Case When CTRCXEncargos.Encargo_Id = 'ICMS' THEN CTRCXEncargos.Valor else 0 End) AS ICMS," & vbCrLf & _
              "              Sum(Case When CTRCXEncargos.Encargo_Id = 'PIS' THEN CTRCXEncargos.Valor else 0 End) AS PIS," & vbCrLf & _
              "              Sum(Case When CTRCXEncargos.Encargo_Id = 'COFINS' THEN CTRCXEncargos.Valor else 0 End) AS COFINS" & vbCrLf & _
              "         FROM NotasFiscais" & vbCrLf & _
              "        INNER JOIN NotasFiscaisXItens" & vbCrLf & _
              "           ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
              "          AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
              "          AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
              "          AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
              "          AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
              "          AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
              "          AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
              "        INNER JOIN SubOperacoes" & vbCrLf & _
              "           ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
              "          AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
              "         FULL JOIN NotasXNotas" & vbCrLf & _
              "           ON NotasFiscais.Empresa_Id      = NotasXNotas.OrigemEmpresa_Id" & vbCrLf & _
              "          AND NotasFiscais.EndEmpresa_Id   = NotasXNotas.OrigemEndEmpresa_Id" & vbCrLf & _
              "          AND NotasFiscais.Cliente_Id      = NotasXNotas.OrigemCliente_Id" & vbCrLf & _
              "          AND NotasFiscais.EndCliente_Id   = NotasXNotas.OrigemEndCliente_Id" & vbCrLf & _
              "          AND NotasFiscais.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id" & vbCrLf & _
              "          AND NotasFiscais.Serie_Id        = NotasXNotas.OrigemSerie_Id" & vbCrLf & _
              "          AND NotasFiscais.Nota_Id         = NotasXNotas.OrigemNota_Id" & vbCrLf & _
              "         FULL JOIN NotasFiscais AS CTRC" & vbCrLf & _
              "        INNER JOIN NotasFiscaisXItens AS CTRCXItens" & vbCrLf & _
              "           ON CTRC.Empresa_Id      = CTRCXItens.Empresa_Id" & vbCrLf & _
              "          AND CTRC.EndEmpresa_Id   = CTRCXItens.EndEmpresa_Id" & vbCrLf & _
              "          AND CTRC.Cliente_Id      = CTRCXItens.Cliente_Id" & vbCrLf & _
              "          AND CTRC.EndCliente_Id   = CTRCXItens.EndCliente_Id" & vbCrLf & _
              "          AND CTRC.EntradaSaida_Id = CTRCXItens.EntradaSaida_Id" & vbCrLf & _
              "          AND CTRC.Serie_Id        = CTRCXItens.Serie_Id" & vbCrLf & _
              "          AND CTRC.Nota_Id         = CTRCXItens.Nota_Id" & vbCrLf & _
              "         LEFT JOIN NotasFiscaisXEncargos AS CTRCxEncargos" & vbCrLf & _
              "           ON CTRCXItens.Empresa_Id      = CTRCxEncargos.Empresa_Id" & vbCrLf & _
              "          AND CTRCXItens.EndEmpresa_Id   = CTRCxEncargos.EndEmpresa_Id" & vbCrLf & _
              "          AND CTRCXItens.Cliente_Id      = CTRCxEncargos.Cliente_Id" & vbCrLf & _
              "          AND CTRCXItens.EndCliente_Id   = CTRCxEncargos.EndCliente_Id" & vbCrLf & _
              "          AND CTRCXItens.EntradaSaida_Id = CTRCxEncargos.EntradaSaida_Id" & vbCrLf & _
              "          AND CTRCXItens.Serie_Id        = CTRCxEncargos.Serie_Id" & vbCrLf & _
              "          AND CTRCXItens.Nota_Id         = CTRCxEncargos.Nota_Id" & vbCrLf & _
              "          AND CTRCXItens.Produto_Id      = CTRCxEncargos.Produto_Id" & vbCrLf & _
              "          AND CTRCXItens.CFOP_Id         = CTRCxEncargos.CFOP_Id" & vbCrLf & _
              "           ON NotasXNotas.Empresa_Id      = Ctrc.Empresa_Id" & vbCrLf & _
              "          And NotasXNotas.EndEmpresa_Id   = Ctrc.EndEmpresa_Id" & vbCrLf & _
              "          And NotasXNotas.Cliente_Id      = Ctrc.Cliente_Id" & vbCrLf & _
              "          And NotasXNotas.EndCliente_Id   = Ctrc.EndCliente_Id" & vbCrLf & _
              "          And NotasXNotas.EntradaSaida_Id = Ctrc.EntradaSaida_Id" & vbCrLf & _
              "          And NotasXNotas.Serie_Id        = Ctrc.Serie_Id" & vbCrLf & _
              "          And NotasXNotas.Nota_Id = Ctrc.Nota_Id" & vbCrLf & _
              "        WHERE (CTRC.Empresa_Id like '" & Left(Empresa(0), 8) & "%')" & vbCrLf & _
              "       	 AND (CTRC.Situacao = 1) " & vbCrLf & _
              "       	 AND (CTRCXItens.QuantidadeFiscal > 1) " & vbCrLf & _
              "       	 AND (CTRCXItens.CFOP_Id BETWEEN 1350 AND 1360 OR" & vbCrLf & _
              "              CTRCXItens.CFOP_Id BETWEEN 2350 AND 2360 OR" & vbCrLf & _
              "              CTRCXItens.CFOP_Id BETWEEN 5350 AND 5360 OR" & vbCrLf & _
              "              CTRCXItens.CFOP_Id BETWEEN 6350 AND 6360) " & vbCrLf & _
              "          AND  year(NotasFiscais.Movimento) = " & DdlAno.SelectedValue & vbCrLf & _
              "          AND  month(NotasFiscais.Movimento) = " & pMes & vbCrLf & _
              "          AND (SubOperacoes.Classe IN ('" & eClassesOperacoes.COMPRAS.ToString & "', '" & eClassesOperacoes.DEPOSITOS.ToString & "', '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'))" & vbCrLf & _
              "          AND CTRCXEncargos.Encargo_Id in ('PRODUTO', 'ICMS', 'PIS', 'COFINS') " & vbCrLf & _
              "        GROUP BY   CTRC.Empresa_Id, CTRC.EndEmpresa_Id, NotasFiscais.Deposito, NotasFiscais.EndDeposito, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, " & vbCrLf & _
              "                  NotasFiscaisXItens.Produto_Id, Month(NotasFiscais.Movimento), Year(NotasFiscais.Movimento), SubOperacoes.ApuracaoDeCustos," & vbCrLf & _
              "                  SubOperacoes.ApuracaodeCustosContraPartida, SubOperacoes.Classe" & vbCrLf & _
              " ) AS Consulta" & vbCrLf & _
              " 	Group By	Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, " & vbCrLf & _
              "                DepositoDestino_Id, EndDepositoDestino_Id, Produto_Id," & vbCrLf & _
              "                Mes_Id, Ano_Id, Placus_Id, ApuracaodeCustosContraPartida" & vbCrLf



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
                  ", 0" & vbCrLf & _
                  ", 0" & vbCrLf & _
                  ", 0" & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  ", 0" & vbCrLf & _
                  ",''" & vbCrLf & _
                  ", 0" & vbCrLf & _
                  ")" & vbCrLf & _
                  " Else" & vbCrLf & _
                  "   Update ApuracaoDeCustos set " & vbCrLf & _
                  "   ValorDoFrete = " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
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

            Array.Add(Sql)
            If Not Banco.GravaBanco(Array) Then Return False
        Next
        Return True
    End Function



    Function Valida()
        If Not IsNumeric(txtCiclos.Text) Then txtCiclos.Text = "5"

        Empresa = ddlEmpresa.SelectedValue.Split("-")

        If ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório .", eTitulo.Info)
            Return False
        End If

        Return True
    End Function

    Function LimpaMovimentoDoMes(ByVal pMes As Integer) As Boolean

        'Sql = "  Update ApuracaoDeCustos"
        'Sql &= " SET    Quantidade = 0, ValorDoProduto = 0, ValorDoFrete = 0, ValorAuxiliar = 0"

        Sql = "  Delete ApuracaoDeCustos"
        Sql &= " WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        Sql &= "   AND Ano_Id = " & DdlAno.SelectedValue
        Sql &= "   AND Mes_Id = " & pMes
        Array.Add(Sql)

        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
            Return True
        Else
            Return False
        End If
    End Function

    Function ApuraSaldoInicial(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        i = 0

        Dim Ano As Integer = DdlAno.SelectedValue
        Dim MesAnterior As Integer = pMes

        If pMes = 1 Then
            MesAnterior = 12
            Ano = Ano - 1
        Else
            MesAnterior = pMes - 1
        End If

        Sql = " SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, CodigoDeCusto_Id, EmpresaDestino_Id, EndEmpresaDestino_Id, " & vbCrLf & _
              "        DepositoDestino_Id, EndDepositoDestino_Id, ProdutoDerivado_Id, Etapa, Quantidade, ValorDoProduto, 0 as ValorDoFrete, 0 as ValorAuxiliar, ProdutoDestino, CodigoDestino, Reduzido" & vbCrLf & _
              "   FROM ApuracaoDeCustos" & vbCrLf & _
              "  WHERE Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "    AND Ano_Id = " & Ano & vbCrLf & _
              "    AND Mes_Id = " & MesAnterior & vbCrLf & _
              "    AND CodigoDeCusto_Id = 920" & vbCrLf & _
              "    AND (Quantidade <> 0 OR ValorDoProduto <> 0 OR ValorDoFrete <> 0 OR ValorAuxiliar <> 0)" & vbCrLf

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
                  "                                         And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "                                         And Mes_Id                = " & pMes & vbCrLf & _
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                         And CodigoDeCusto_Id      = 101" & vbCrLf & _
                  "                                    )" & vbCrLf & _
                  "                           then 'S'" & vbCrLf & _
                  "                           else 'N'" & vbCrLf & _
                  "                       end) ;" & vbCrLf & _
                  " if @Exist = 'N' " & vbCrLf & _
                  " INSERT INTO ApuracaoDeCustos (" & vbCrLf & _
                  "                                Empresa_Id" & vbCrLf & _
                  "                               ,EndEmpresa_Id" & vbCrLf & _
                  "                               ,Deposito_Id" & vbCrLf & _
                  "                               ,EndDeposito_Id" & vbCrLf & _
                  "                               ,Ano_Id" & vbCrLf & _
                  "                               ,Mes_Id" & vbCrLf & _
                  "                               ,Produto_Id" & vbCrLf & _
                  "                               ,CodigoDeCusto_Id" & vbCrLf & _
                  "                               ,EmpresaDestino_Id" & vbCrLf & _
                  "                               ,EndEmpresaDestino_Id" & vbCrLf & _
                  "                               ,DepositoDestino_Id" & vbCrLf & _
                  "                               ,EndDepositoDestino_Id" & vbCrLf & _
                  "                               ,ProdutoDerivado_Id" & vbCrLf & _
                  "                               ,Quantidade" & vbCrLf & _
                  "                               ,ValorDoProduto" & vbCrLf & _
                  "                               ,ValorDoFrete" & vbCrLf & _
                  "                               ,ProdutoDestino" & vbCrLf & _
                  "                               ,CodigoDestino)" & vbCrLf & _
                  " VALUES( '" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "        ,'" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndDeposito_Id") & vbCrLf & _
                  "        , " & DdlAno.SelectedValue & vbCrLf & _
                  "        , " & pMes & vbCrLf & _
                  "        ,'" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "        , 101" & vbCrLf & _
                  "        ,'" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "        ,'" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "        , " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "        ,'" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  "        , " & Str(Dr("Quantidade")) & vbCrLf & _
                  "        , " & Dr("ValorDoProduto").ToString.Replace(",", ".") & vbCrLf & _
                  "        , " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  "        ,'" & Dr("ProdutoDestino") & "'" & vbCrLf & _
                  "        ,0" & vbCrLf & _
                  "       )" & vbCrLf & _
                  " Else" & vbCrLf & _
                  "  Update ApuracaoDeCustos set " & vbCrLf & _
                  "     Quantidade     = " & Str(Dr("Quantidade")) & vbCrLf & _
                  "    ,ValorDoProduto = " & Dr("ValorDoProduto").ToString.Replace(",", ".") & vbCrLf & _
                  "    ,ValorDoFrete   = " & Dr("ValorDoFrete").ToString.Replace(",", ".") & vbCrLf & _
                  "  Where Empresa_Id            ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "    And EndEmpresa_Id         = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "    And Deposito_Id           ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "    And EndDeposito_Id        = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "    And Ano_Id                = " & DdlAno.SelectedValue & vbCrLf & _
                  "    And Mes_Id                = " & pMes & vbCrLf & _
                  "    And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "    And CodigoDeCusto_Id      = 101" & vbCrLf & _
                  "    And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "    And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "    And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "    And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "    And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf

            Array.Add(Sql)
            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
                i += 1
            Else
                erros += 1
            End If
        Next
        Return erros
    End Function

    Function CapturaRazao(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        i = 0

        Sql = "SELECT * from ApuracaoDeCustosXFiltroRazao"
        Sqla = "; "
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            i = 0
            If Dr("Processo") = "D" Then
                Sqla &= "Delete #Razao where "
            End If

            If Dr("Processo") = "U" Then
                Sqla &= " Update #Razao "
                Sqla &= " Set Produto = '" & Dr("Produto") & "'"
                Sqla &= " where "
            End If

            If Dr("Empresa_Id") <> "" Then
                If Dr("EmpresaSinal_Id") = "LIKE" Then
                    Sqla &= " Empresa_Id Like ('" & Dr("Empresa_Id") & "%')"
                    i += 1
                Else
                    Sqla &= " Empresa_Id " & Dr("EmpresaSinal_Id") & " '" & Dr("Empresa_Id") & "'"
                    i += 1
                End If
            End If


            If Dr("Conta_Id") <> "" Then
                If Dr("ContaSinal_Id") = "LIKE" Then
                    Sqla &= IIf(i = 0, " Conta_Id LIKE " & "('" & Dr("Conta_Id") & "%')", " And Conta_Id  LIKE " & "('" & Dr("Conta_Id") & "%')")
                    i += 1
                Else
                    Sqla &= IIf(i = 0, " Conta_Id " & Dr("ContaSinal_Id") & " '" & Dr("Conta_Id") & "'", " And Conta_Id " & Dr("ContaSinal_Id") & " '" & Dr("Conta_Id") & "'")
                    i += 1
                End If
            End If


            If Dr("Produto_Id") <> "" Then
                If Dr("ProdutoSinal_Id") = "LIKE" Then
                    Sqla &= IIf(i = 0, " Produto LIKE " & "('" & Dr("Produto_Id") & "%')", " And Produto  LIKE " & "('" & Dr("Produto_Id") & "%')")
                    i += 1
                Else
                    Sqla &= IIf(i = 0, " Produto " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'", " And Produto " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'")
                    i += 1
                End If
            End If


            If Dr("Lote_Id") > 0 Then
                Sqla &= IIf(i = 0, " Lote_Id " & Dr("LoteSinal_Id") & " " & Dr("Lote_Id"), " And Lote_Id " & Dr("LoteSinal_Id") & " " & Dr("Lote_Id"))
                i += 1
            End If

            If Dr("Placus_Id") > 0 Then
                Sqla &= IIf(i = 0, " Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"), " And Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"))
            End If

            If Dr("Sql") <> "" Then
                Sqla &= IIf(i = 0, " " & Dr("sql"), " And " & Dr("sql"))
            End If

            Sqla &= "; "
        Next

        Sql = "  SELECT	Razao.Empresa_Id, Razao.EndEmpresa_Id,  Razao.Conta_Id, Razao.Lote_Id, PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto,"
        Sql &= "        isnull(Razao.Deposito, Razao.Empresa_Id) AS Deposito, "
        Sql &= "        isnull(Razao.EndDeposito, Razao.EndEmpresa_Id)AS EndDeposito, "
        Sql &= "        Sum(Razao.DebitoOficial) As DebitoOficial, Sum(Razao.CreditoOficial) AS CreditoOficial,"
        Sql &= "        Sum(Razao.DebitoQuantidade) As DebitoQuantidade, Sum(Razao.CreditoQuantidade) AS CreditoQuantidade"
        Sql &= "        Into #Razao"
        Sql &= " FROM   Razao INNER JOIN"
        Sql &= "        PlanoDeCustosXOrigem ON Razao.Conta_Id LIKE PlanoDeCustosXOrigem.Conta_Id + '%'"
        Sql &= " INNER JOIN Produtos "
        Sql &= "    ON Razao.Produto = Produtos.Produto_Id"
        Sql &= " INNER JOIN GruposDeEstoques"
        Sql &= "    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id"
        Sql &= " WHERE	Empresa_Id like '" & Left(Empresa(0), 8) & "%'"

        If CInt(txtDia.Text) > 0 Then
            Sql &= "        AND year(Razao.Movimento_Id) = " & DdlAno.SelectedValue
            Sql &= "        AND month(Razao.Movimento_Id) = " & pMes
            Sql &= "        AND (Razao.Movimento_Id)  <= '" & DdlAno.SelectedValue & "-" & pMes.ToString("00") & "-" & CInt(txtDia.Text).ToString("00") & "'"
        Else
            Sql &= "        AND year(Razao.Movimento_Id) = " & DdlAno.SelectedValue
            Sql &= "        AND month(Razao.Movimento_Id) = " & pMes
        End If


        Sql &= "        And GruposDeEstoques.custo = 1"
        Sql &= "        And Razao.Produto <> '' "
        Sql &= "        And  Razao.Lote_Id not in('7000') "
        Sql &= " GROUP  BY	Razao.Empresa_Id, Razao.EndEmpresa_Id, "
        Sql &= "        isnull(Razao.Deposito, Razao.Empresa_Id) , "
        Sql &= "        isnull(Razao.EndDeposito, Razao.EndEmpresa_Id), "
        Sql &= "        PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto, Razao.Conta_Id, Razao.Lote_Id "

        Sql &= Sqla & "; "

        Sql &= " SELECT Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Codigo_Id, Produto, " & vbCrLf & _
               "        case when codigo_id = 320 then sum(CreditoOficial) " & vbCrLf & _
               "             when codigo_id = 120 then sum(DebitoOficial)" & vbCrLf & _
               "             else SUM(DebitoOficial - CreditoOficial)" & vbCrLf & _
               "        end Valor," & vbCrLf & _
               "        case when codigo_id = 320 then sum(isnull(CreditoQuantidade,0)) " & vbCrLf & _
               "             when codigo_id = 120 then sum(isnull(DebitoQuantidade,0)) " & vbCrLf & _
               "             else isnull(SUM(DebitoQuantidade - CreditoQuantidade), 0) " & vbCrLf & _
               "        end Quantidade" & vbCrLf & _
               "   FROM #Razao " & vbCrLf & _
               "  INNER JOIN Produtos" & vbCrLf & _
               "     ON #Razao.Produto LIKE Produtos.Produto_Id" & vbCrLf & _
               "  WHERE (DebitoOficial - CreditoOficial) <> 0" & vbCrLf & _
               "  GROUP BY Empresa_Id, EndEmpresa_Id, Codigo_Id, Produto, Deposito, EndDeposito; " & vbCrLf

        Sql &= "  Drop Table #Razao;"


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare"
            Sql &= " @Exist as varchar(1)"
            Sql &= " set @Exist = (select case "
            Sql &= "       when exists ("
            Sql &= "                 select * from ApuracaoDeCustos "
            Sql &= " Where "
            Sql &= "     Empresa_Id = '" & Dr("Empresa_Id") & "'"
            Sql &= " And EndEmpresa_Id = " & Dr("EndEmpresa_Id")
            Sql &= " And Deposito_Id = '" & Dr("Deposito") & "'"
            Sql &= " And EndDeposito_Id = " & Dr("EndDeposito")
            Sql &= " And Ano_Id = " & DdlAno.SelectedValue
            Sql &= " And Mes_Id = " & pMes
            Sql &= " And Produto_Id = '" & Dr("Produto") & "'"
            Sql &= " And CodigoDeCusto_Id = " & Dr("Codigo_Id")

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
            Sql &= " VALUES('" & Dr("Empresa_Id") & "'"
            Sql &= ", " & Dr("EndEmpresa_Id")
            Sql &= ", '" & Dr("Deposito") & "'"
            Sql &= ", " & Dr("EndDeposito")
            Sql &= ", " & DdlAno.SelectedValue
            Sql &= ", " & pMes
            Sql &= ", '" & Dr("Produto") & "'"
            Sql &= ", " & Dr("Codigo_Id")
            Sql &= ", ''" 'Empresa Destino
            Sql &= ", 0"  'End Empresa Destino
            Sql &= ", ''" 'Deposito Destino
            Sql &= ", 0"  'End DepositoDestino
            Sql &= ", ''" 'Produto Derivado

            Sql &= ", " & Str(Dr("Quantidade"))  'Quantidade
            Sql &= ", " & Str(Dr("Valor"))
            Sql &= ", 0" '& Dr("ValorDoFrete").ToString.Replace(",", ".")
            Sql &= ", 0" '& Dr("ValorAuxiliar").ToString.Replace(",", ".")
            Sql &= ", ''" '& Dr("ProdutoDestino") & "'"
            Sql &= ", 0" '& Dr("CodigoDestino")
            Sql &= ")"

            Sql &= " Else"

            Sql &= "  Update ApuracaoDeCustos set "
            Sql &= "  ValorDoProduto = " & Str(Dr("Valor"))
            Sql &= ", Quantidade = " & Str(Dr("Quantidade"))

            Sql &= " Where Empresa_Id = '" & Dr("Empresa_Id") & "'"
            Sql &= "   And EndEmpresa_Id = " & Dr("EndEmpresa_Id")
            Sql &= "   And Deposito_Id = '" & Dr("Deposito") & "'"
            Sql &= "   And EndDeposito_Id = " & Dr("EndDeposito")
            Sql &= "   And Ano_Id = " & DdlAno.SelectedValue
            Sql &= "   And Mes_Id = " & pMes
            Sql &= "   And Produto_Id = '" & Dr("Produto") & "'"
            Sql &= "   And CodigoDeCusto_Id = " & Dr("Codigo_Id")

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
                i += 1
            Else
                erros += 1
            End If
        Next
        Return erros
    End Function

    Function CapturaEstoques(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        i = 0

        Sql = "SELECT * from ApuracaoDeCustosXFiltroEstoques"
        Sqla = "; "
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            i = 0
            Sqla &= "Delete #Estoques where "

            If Dr("Empresa_Id") <> "" Then
                Sqla &= " Empresa_Id like '" & Dr("Empresa_Id") & "%'"
                i += 1
            End If

            If Dr("Deposito_Id") <> "" Then
                Sqla &= IIf(i = 0, " Deposito_Id = '" & Dr("Deposito_Id"), " And Deposito_Id = " & Dr("Deposito_Id"))
                i += 1
            End If

            If Dr("Produto_Id") <> "" Then
                Sqla &= IIf(i = 0, " Produto = " & Dr("Produto_Id"), " And Produto = " & Dr("Produto_Id"))
                i += 1
            End If

            If Dr("ProdutoDerivado_Id") <> "" Then
                Sqla &= IIf(i = 0, " ProdutoDerivado_Id = " & Dr("ProdutoDerivado_Id"), " And ProdutoDerivado_Id = " & Dr("ProdutoDerivado_Id"))
                i += 1
            End If

            If Dr("Operacao_Id") > 0 Then
                Sqla &= IIf(i = 0, " Operacao_Id = " & Dr("Operacao_Id"), " And Operacao_Id = " & Dr("Operacao_Id"))
                i += 1
            End If

            If Dr("SubOperacao_Id") > 0 Then
                Sqla &= IIf(i = 0, " SubOperacao_Id = " & Dr("SubOperacao_Id"), " And SubOperacao_Id = " & Dr("SubOperacao_Id"))
            End If
            Sqla &= "; "
        Next

        Sql = " SELECT Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Operacao_Id, " & vbCrLf & _
              "        Producao.SubOperacao_Id, Producao.ProdutoDerivado_Id, SubOperacoes.ApuracaoDeCustos AS Placus_Id,  isnull(SubOperacoes.ApuracaoDeCustosContraPartida,0) as PlacusContraPartida," & vbCrLf & _
              "        Convert(Decimal(18,4),ABS(SUM(Producao.Entradas - Producao.Saidas)))  AS Quantidade" & vbCrLf & _
              "   Into #Estoques" & vbCrLf & _
              "   FROM Producao" & vbCrLf & _
              "  INNER JOIN SubOperacoes" & vbCrLf & _
              "     ON Producao.Operacao_Id    = SubOperacoes.Operacao_Id " & vbCrLf & _
              "    AND Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
              "  INNER JOIN Produtos " & vbCrLf & _
              "     ON Producao.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
              "  INNER JOIN GruposDeEstoques" & vbCrLf & _
              "    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf & _
              " WHERE Producao.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'" & vbCrLf

        If CInt(txtDia.Text) > 0 Then
            Sql &= "        AND year(Producao.Movimento_Id) = " & DdlAno.SelectedValue
            Sql &= "        AND month(Producao.Movimento_Id) = " & pMes
            Sql &= "        AND (Producao.Movimento_Id)  <= '" & DdlAno.SelectedValue & "-" & pMes.ToString("00") & "-" & CInt(txtDia.Text).ToString("00") & "'"
        Else
            Sql &= "        AND year(Producao.Movimento_Id) = " & DdlAno.SelectedValue
            Sql &= "        AND month(Producao.Movimento_Id) = " & pMes
        End If

        Sql &= "   AND Producao.FisicoFiscal_Id     = 2" & vbCrLf & _
               "   AND GruposDeEstoques.custo       = 1" & vbCrLf & _
               " GROUP BY Producao.Empresa_Id,  Producao.EndEmpresa_Id,  Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.ProdutoDerivado_Id, " & vbCrLf & _
               "          Producao.Operacao_Id, Producao.SubOperacao_Id, SubOperacoes.ApuracaoDeCustos,  isnull(SubOperacoes.ApuracaoDeCustosContraPartida,0)" & vbCrLf

        Sql &= Sqla

        Sql &= " SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, #Estoques.Produto_Id,  " & vbCrLf & _
               "        ProdutoDerivado_Id, Placus_Id, PlacusContraPartida," & vbCrLf & _
               "        SUM(quantidade) AS Quantidade " & vbCrLf & _
               "   FROM #Estoques" & vbCrLf & _
               "  INNER JOIN Produtos " & vbCrLf & _
               "     ON #Estoques.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
               "  GROUP BY Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, #Estoques.Produto_Id,  " & vbCrLf & _
               "           ProdutoDerivado_Id, Placus_Id, PlacusContraPartida; " & vbCrLf

        Sql &= "  Drop Table #Estoques;"


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows

            Sql = " Declare"
            Sql &= " @Exist as varchar(1)"
            Sql &= " set @Exist = (select case "
            Sql &= "                        when exists ("
            Sql &= "                                     select 1"
            Sql &= "                                       from ApuracaoDeCustos "
            Sql &= "                                      Where Empresa_Id         ='" & Dr("Empresa_Id") & "'"
            Sql &= "                                        And EndEmpresa_Id      = " & Dr("EndEmpresa_Id")
            Sql &= "                                        And Deposito_Id        ='" & Dr("Deposito_Id") & "'"
            Sql &= "                                        And EndDeposito_Id     = " & Dr("EndDeposito_Id")
            Sql &= "                                        And Ano_Id             = " & DdlAno.SelectedValue
            Sql &= "                                        And Mes_Id             = " & pMes
            Sql &= "                                        And Produto_Id         ='" & Dr("Produto_Id") & "'"
            Sql &= "                                        And ProdutoDerivado_Id ='" & Dr("ProdutoDerivado_Id") & "'"
            Sql &= "                                        And CodigoDeCusto_Id   = " & Dr("Placus_Id")
            Sql &= "                                        And CodigoDestino      = " & Dr("PlacusContraPartida")
            Sql &= "                                     )"
            Sql &= "                           then 'S'"
            Sql &= "                           else 'N'"
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
            Sql &= ", CodigoDestino"
            Sql &= ", Reduzido)"
            Sql &= " VALUES('" & Dr("Empresa_Id") & "'"
            Sql &= ", " & Dr("EndEmpresa_Id")
            Sql &= ", '" & Dr("Deposito_Id") & "'"
            Sql &= ", " & Dr("EndDeposito_Id")
            Sql &= ", " & DdlAno.SelectedValue
            Sql &= ", " & pMes
            Sql &= ", '" & Dr("Produto_Id") & "'"
            Sql &= ", " & Dr("Placus_Id")
            Sql &= ", ''" 'Empresa Destino
            Sql &= ", 0"  'End Empresa Destino
            Sql &= ", ''" 'Deposito Destino
            Sql &= ", 0"  'End DepositoDestino
            Sql &= ", '" & Dr("ProdutoDerivado_Id") & "'"
            Sql &= ", " & Str(Dr("Quantidade"))
            Sql &= ", 0" '& Dr("Valor").ToString.Replace(",", ".")
            Sql &= ", 0" '& Dr("ValorDoFrete").ToString.Replace(",", ".")
            Sql &= ", 0" '& Dr("ValorAuxiliar").ToString.Replace(",", ".")
            Sql &= ", ''" '& Dr("ProdutoDestino") & "'"
            Sql &= ", " & Dr("PlacusContraPartida")
            Sql &= ", ''" '& Dr("Reduzido") & "'"
            Sql &= ")"
            Sql &= " Else"
            Sql &= "  Update ApuracaoDeCustos set "
            Sql &= "  Quantidade = " & Str(Dr("Quantidade"))
            Sql &= " Where Empresa_Id         ='" & Dr("Empresa_Id") & "'"
            Sql &= "   And EndEmpresa_Id      = " & Dr("EndEmpresa_Id")
            Sql &= "   And Deposito_Id        ='" & Dr("Deposito_Id") & "'"
            Sql &= "   And EndDeposito_Id     = " & Dr("EndDeposito_Id")
            Sql &= "   And Ano_Id             = " & DdlAno.SelectedValue
            Sql &= "   And Mes_Id             = " & pMes
            Sql &= "   And Produto_Id         ='" & Dr("Produto_Id") & "'"
            Sql &= "   And ProdutoDerivado_Id ='" & Dr("ProdutoDerivado_Id") & "'"
            Sql &= "   And CodigoDeCusto_Id   = " & Dr("Placus_Id")
            Sql &= "   And CodigoDestino      = " & Dr("PlacusContraPartida")

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
                i += 1
            Else
                erros += 1
            End If
        Next
        Return erros
    End Function

    Function CapturaNotasFiscais(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        Dim sqlWhere As String = ""
        Dim sqlFiltros As String = ""
        Dim ds As DataSet
        i = 0

        Sql = "SELECT * from ApuracaoDeCustosXFiltroNotas order by ordem"
        ds = Banco.ConsultaDataSet(Sql, "Consulta")

        For Each Dr As DataRow In ds.Tables(0).Rows
            sqlWhere = ""
            i = 0

            If Dr("Processo") = "D" Then
                sqlFiltros &= "Delete #Notas" & vbCrLf & _
                       " where "
            End If

            If Dr("Processo") = "U" Then
                sqlFiltros &= " Update #Notas Set " & vbCrLf

                If Dr("Produto") <> "" Then
                    sqlFiltros &= " Produto_Id = '" & Dr("Produto") & "'" & vbCrLf
                End If

                If Dr("Empresa_Id") <> Dr("Empresa") And Dr("Empresa").ToString().Length = 14 Then
                    If Dr("Produto") <> "" Then
                        sqlFiltros &= ", "
                    End If

                    sqlFiltros &= " Empresa_Id     ='" & Dr("Empresa") & "'," & vbCrLf & _
                           " EndEmpresa_Id  = " & Dr("EndEmpresa") & "," & vbCrLf

                    If Dr("Deposito_Id") <> Dr("Deposito") And Dr("Empresa").ToString().Length = 14 Then
                        sqlFiltros &= " Deposito_Id    ='" & Dr("Deposito") & "'," & vbCrLf
                        sqlFiltros &= " EndDeposito_Id = " & Dr("EndDeposito") & vbCrLf
                    End If

                    If Dr("Deposito_Id") <> "" Then
                        sqlWhere &= " and Deposito_id = '" & Dr("Deposito_Id") & "'" & vbCrLf & _
                                    " and EndDeposito_id = '" & Dr("EndDeposito_id") & "'" & vbCrLf
                    End If
                ElseIf Dr("Deposito_Id") <> Dr("Deposito") And Dr("Empresa").ToString().Length = 14 Then
                    If Dr("Produto") <> "" Then
                        sqlFiltros &= ", "
                    End If

                    sqlFiltros &= " Deposito_Id    = '" & Dr("Deposito") & "'," & vbCrLf & _
                           " EndDeposito_Id = " & Dr("EndDeposito")

                    If Dr("Deposito_Id") <> "" Then
                        sqlWhere &= " and Deposito_id = '" & Dr("Deposito_Id") & "'"
                        sqlWhere &= " and EndDeposito_id = " & Dr("EndDeposito_id")
                    End If
                End If

                sqlFiltros &= " where "
            End If

            If Dr("Empresa_Id") <> "" Then
                If Dr("EmpresaSinal_Id") = "LIKE" Then
                    sqlFiltros &= " Empresa_Id Like ('" & Dr("Empresa_Id") & "%')" & vbCrLf
                    i += 1
                Else
                    sqlFiltros &= " Empresa_Id " & Dr("EmpresaSinal_Id") & " '" & Dr("Empresa_Id") & "'" & vbCrLf
                    i += 1
                End If
            End If


            If Dr("Produto_Id") <> "" Then
                If Dr("ProdutoSinal_Id") = "LIKE" Then
                    sqlFiltros &= IIf(i = 0, " Produto_Id LIKE " & "('" & Dr("Produto_Id") & "%')", " And Produto_Id  LIKE " & "('" & Dr("Produto_Id") & "%')") & vbCrLf
                    i += 1
                Else
                    sqlFiltros &= IIf(i = 0, " Produto_Id " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'", " And Produto_Id " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'") & vbCrLf
                    i += 1
                End If
            End If


            If Dr("Placus_Id") > 0 Then
                sqlFiltros &= IIf(i = 0, " Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"), " And Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id")) & vbCrLf
            End If

            sqlFiltros &= sqlWhere & "; " & vbCrLf
        Next


        Sql = "SELECT NotasFiscais.Empresa_Id," & vbCrLf & _
              "       NotasFiscais.EndEmpresa_Id," & vbCrLf & _
              "       NotasFiscais.Empresa_Id AS Deposito_Id, " & vbCrLf & _
              "       NotasFiscais.EndEmpresa_Id AS EndDeposito_Id, " & vbCrLf & _
              "       YEAR(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)) AS Ano_Id, " & vbCrLf & _
              "       MONTH(ISNULL(NotasFiscais.DataParaCusto, NotasFiscais.Movimento)) AS Mes_Id," & vbCrLf & _
              "       NotasFiscaisXItens.Produto_Id," & vbCrLf & _
              "       SubOperacoes.ApuracaoDeCustos AS Placus_Id, " & vbCrLf & _
              "       CASE " & vbCrLf & _
              "          WHEN SubOperacoes.ApuracaodeCustosContraPartida <> 0 " & vbCrLf & _
              "            THEN NotasFiscais.Cliente_Id " & vbCrLf & _
              "            ELSE ''" & vbCrLf & _
              "       END AS DepositoDestino_Id," & vbCrLf & _
              "       CASE" & vbCrLf & _
              "          WHEN SubOperacoes.ApuracaodeCustosContraPartida <> 0 " & vbCrLf & _
              "            THEN NotasFiscais.EndCliente_Id" & vbCrLf & _
              "            ELSE 0" & vbCrLf & _
              "       END AS EndDepositoDestino_Id," & vbCrLf & _
              "       SubOperacoes.ApuracaodeCustosContraPartida," & vbCrLf & _
              "       SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade," & vbCrLf & _
              "       SUM(NotasFiscaisXItens.Valor) AS ValorDoProduto, " & vbCrLf & _
              "       NotasFiscais.EntradaSaida_Id," & vbCrLf & _
              "       SUM(Enc.ValorIPI) AS ValorIPI," & vbCrLf & _
              "       SUM(Enc.ValorICMS)AS ValorICMS," & vbCrLf & _
              "       SUM(Enc.ValorCOFINS) AS ValorCOFINS," & vbCrLf & _
              "       SUM(Enc.ValorPIS) AS ValorPIS," & vbCrLf & _
              "       SUM(Enc.ValorFunrural+Enc.ValorFunruralJudicial+Enc.ValorSenar) AS ValorFunrural," & vbCrLf & _
              "       SubOperacoes.Devolucao" & vbCrLf & _
              "  Into #Notas" & vbCrLf & _
              "  FROM GruposDeEstoques" & vbCrLf & _
              "  LEFT JOIN Produtos" & vbCrLf & _
              "    ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & vbCrLf & _
              " RIGHT JOIN NotasFiscais" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
              "    ON Produtos.Produto_Id          = NotasFiscaisXItens.Produto_Id" & vbCrLf & _
              "  LEFT JOIN SubOperacoes" & vbCrLf & _
              "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
              "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
              "  LEFT JOIN (" & vbCrLf & _
              "             SELECT  NFxE.Empresa_Id, NFxE.EndEmpresa_Id," & vbCrLf & _
              "                     NFxE.Cliente_Id, NFxE.EndCliente_Id," & vbCrLf & _
              "                     NFxE.EntradaSaida_Id, NFxE.Serie_Id, NFxE.Nota_Id," & vbCrLf & _
              "                     NFxE.Produto_Id, NFxE.CFOP_Id," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'IPI'               THEN NFxE.Valor ELSE 0 END), 0) AS ValorIPI," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'ICMS'              THEN NFxE.Valor ELSE 0 END), 0) AS ValorICMS," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'FUNRURAL'          THEN NFxE.Valor ELSE 0 END), 0) AS ValorFunrural," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'FUNRURAL JUDICIAL' THEN NFxE.Valor ELSE 0 END), 0) AS ValorFunruralJudicial," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'SENAR'             THEN NFxE.Valor ELSE 0 END), 0) AS ValorSenar," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'PIS'               THEN NFxE.Valor ELSE 0 END), 0) AS ValorPIS," & vbCrLf & _
              "                     ISNULL(SUM(CASE WHEN UPPER(NFxE.Encargo_Id) = 'COFINS'            THEN NFxE.Valor ELSE 0 END), 0) AS ValorCOFINS" & vbCrLf & _
              "               FROM NotasFiscaisXEncargos NFxE" & vbCrLf & _    
              "               Left join ClientesxEmpresas CxE" & vbCrLf & _
              "                 on NFxE.Empresa_Id    = CxE.Empresa_Id" & vbCrLf & _
              "                and NFxE.EndEmpresa_Id = CxE.EndEmpresa_Id" & vbCrLf & _
              "              GROUP BY NFxE.Empresa_Id, NFxE.EndEmpresa_Id," & vbCrLf & _
              "                       NFxE.Cliente_Id, NFxE.EndCliente_Id," & vbCrLf & _
              "                       NFxE.EntradaSaida_Id, NFxE.Serie_Id, NFxE.Nota_Id," & vbCrLf & _
              "                       NFxE.Produto_Id, NFxE.CFOP_Id" & vbCrLf & _
              "              ) AS Enc" & vbCrLf & _
              "    ON NotasFiscaisXItens.Empresa_Id      = Enc.Empresa_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.Cliente_Id      = Enc.Cliente_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.Serie_Id        = Enc.Serie_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.Nota_Id         = Enc.Nota_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.Produto_Id      = Enc.Produto_Id" & vbCrLf & _
              "   AND NotasFiscaisXItens.CFOP_Id         = Enc.CFOP_Id" & vbCrLf & _
              " WHERE NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%'" & vbCrLf

        If CInt(txtDia.Text) > 0 Then
            Sql &= "        AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue
            Sql &= "        AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes
            Sql &= "        AND isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento) <= '" & DdlAno.SelectedValue & "-" & pMes.ToString("00") & "-" & CInt(txtDia.Text).ToString("00") & "'"
        Else
            Sql &= "        AND year(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlAno.SelectedValue
            Sql &= "        AND month(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & pMes
        End If

        Sql &= "   AND SubOperacoes.ApuracaoDeCustos > 0" & vbCrLf & _
               "   And NotasFiscais.Situacao         = 1" & vbCrLf & _
               "   And GruposDeEstoques.custo        = 1" & vbCrLf & _
               "   AND Not SubOperacoes.ApuracaoDeCustos in (200, 600)" & vbCrLf & _
               " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito, NotasFiscaisXItens.EndDeposito, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))," & vbCrLf & _
               "          MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro," & vbCrLf & _
               "          NotasFiscaisXItens.EndDepositoTerceiro, SubOperacoes.ApuracaodeCustosContrapartida, NotasFiscais.entradaSaida_Id , SubOperacoes.Devolucao" & vbCrLf


        Sql &= sqlFiltros & "; "

        Sql &= " SELECT Empresa_Id, EndEmpresa_Id," & vbCrLf & _
               "        Deposito_Id, EndDeposito_Id," & vbCrLf & _
               "        DepositoDestino_Id, EndDepositoDestino_Id," & vbCrLf & _
               "        Placus_Id," & vbCrLf & _
               "        isnull(ApuracaodeCustosContrapartida,0) as ApuracaodeCustosContrapartida," & vbCrLf & _
               "        Produto_Id," & vbCrLf & _
               "        sum(Quantidade) as Quantidade, " & vbCrLf & _
               "        Sum((ValorDoProduto + ValorIPI) - (ValorCOFINS + ValorPIS + ValorICMS + ValorFunrural)) as ValorDoProduto" & vbCrLf & _
               "  FROM #Notas" & vbCrLf & _
               "  Where Placus_Id <> 1" & vbCrLf & _
               "  GROUP BY Empresa_Id, EndEmpresa_Id, Placus_Id,isnull(ApuracaodeCustosContrapartida,0), Produto_Id, Deposito_Id, EndDeposito_Id, DepositoDestino_Id, EndDepositoDestino_Id;" & vbCrLf

        Sql &= "  Drop Table #Notas;"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare" & vbCrLf & _
                  " @Exist as varchar(1)" & vbCrLf & _
                  " set @Exist = (select case " & vbCrLf & _
                  "                        when exists (" & vbCrLf & _
                  "                                     Select 1 " & vbCrLf & _
                  "                                       From ApuracaoDeCustos " & vbCrLf & _
                  "                                      Where Empresa_Id       ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "                                        And EndEmpresa_Id    = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "                                        And Deposito_Id      ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "                                        And EndDeposito_Id   = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "                                        And Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
                  "                                        And Mes_Id           = " & pMes & vbCrLf & _
                  "                                        And Produto_Id       ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                        And CodigoDeCusto_Id = " & Dr("Placus_Id") & vbCrLf & _
                  "                                     )" & vbCrLf & _
                  "                           then 'S'" & vbCrLf & _
                  "                           else 'N'" & vbCrLf & _
                  "                      end) ;" & vbCrLf

            Sql &= " if @Exist = 'N' "
            Sql &= "  INSERT INTO ApuracaoDeCustos ( "
            Sql &= "  Empresa_Id, EndEmpresa_Id"
            Sql &= ", Deposito_Id, EndDeposito_Id"
            Sql &= ", Ano_Id, Mes_Id"
            Sql &= ", Produto_Id"
            Sql &= ", CodigoDeCusto_Id"
            Sql &= ", EmpresaDestino_Id, EndEmpresaDestino_Id"
            Sql &= ", DepositoDestino_Id, EndDepositoDestino_Id"
            Sql &= ", ProdutoDerivado_Id"
            Sql &= ", Quantidade"
            Sql &= ", ValorDoProduto"
            Sql &= ", ValorDoFrete"
            Sql &= ", ValorAuxiliar"
            Sql &= ", ProdutoDestino"
            Sql &= ", CodigoDestino)"
            Sql &= " VALUES('" & Dr("Empresa_Id") & "'"
            Sql &= ", " & Dr("EndEmpresa_Id")
            Sql &= ", '" & Dr("Deposito_Id") & "'"
            Sql &= ", " & Dr("EndDeposito_Id")
            Sql &= ", " & DdlAno.SelectedValue
            Sql &= ", " & pMes
            Sql &= ", '" & Dr("Produto_Id") & "'"
            Sql &= ", " & Dr("Placus_Id")
            Sql &= ", ''"
            Sql &= ", 0"

            If Dr("ApuracaodeCustosContrapartida") <> 0 Then
                Sql &= ",  '" & Dr("DepositoDestino_Id") & "'"
                Sql &= ", " & Dr("EndDeposito_Id")
                Sql &= ", '" & Dr("Produto_Id") & "'"
            Else
                Sql &= ", ''"
                Sql &= ", 0"
                Sql &= ", ''"

            End If

            Sql &= ", " & Str(Dr("Quantidade"))
            Sql &= ", " & Str(Dr("ValorDoProduto"))
            Sql &= ", 0" '& Dr("ValorDoFrete").ToString.Replace(",", ".")
            Sql &= ", " & Str(Dr("ValorDoProduto"))

            If Dr("ApuracaodeCustosContrapartida") <> 0 Then
                Sql &= ", '" & Dr("Produto_Id") & "'"
                Sql &= ", '" & Dr("ApuracaodeCustosContrapartida") & "'"
            Else
                Sql &= ", ''" '& Dr("ProdutoDestino") & "'"
                Sql &= ", 0" '& Dr("CodigoDestino")
            End If

            Sql &= ")"

            Sql &= " Else"

            Sql &= "   Update ApuracaoDeCustos set "
            Sql &= "   Quantidade = " & Str(Dr("Quantidade"))
            Sql &= ",  ValorDoProduto = " & Str(Dr("ValorDoProduto"))
            Sql &= ",  ValorAuxiliar = " & Str(Dr("ValorDoProduto"))

            Sql &= " Where Empresa_Id       = '" & Dr("Empresa_Id") & "'"
            Sql &= "   And EndEmpresa_Id    = " & Dr("EndEmpresa_Id")
            Sql &= "   And Deposito_Id      = '" & Dr("Deposito_Id") & "'"
            Sql &= "   And EndDeposito_Id   = " & Dr("EndDeposito_Id")
            Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
            Sql &= "   And Mes_Id           = " & pMes
            Sql &= "   And Produto_Id       = '" & Dr("Produto_Id") & "'"
            Sql &= "   And CodigoDeCusto_Id = " & Dr("Placus_Id")

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
                i += 1
            Else
                erros += 1
            End If
        Next
        Return erros
    End Function

    Sub CapturaNotasFiscaisContraPartida(ByVal pMes As Integer)
        i = 0
        ChkResultado.Items.Add(New ListItem(" 10 - Capturando dados das Notas Fiscais Contra Partida- Criando Filtro"))

        Sql = "  SELECT     NotasFiscais.Empresa_Id, "
        Sql &= " 	        NotasFiscais.EndEmpresa_Id, "
        Sql &= "            NotasFiscais.Cliente_Id AS Deposito_Id,  	   "
        Sql &= " 	        NotasFiscais.EndCliente_Id AS EndDeposito_Id,"
        Sql &= "            YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) AS Ano_Id, "
        Sql &= " 		    MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) AS Mes_Id, NotasFiscaisXItens.Produto_Id, "
        Sql &= "            SubOperacoes.ApuracaoDeCustosContraPartida AS Placus_Id, "
        Sql &= " 	        NotasFiscaisXItens.Deposito AS DepositoDestino_Id, "
        Sql &= "            NotasFiscaisXItens.EndDeposito AS EndDepositoDestino_Id, "
        Sql &= " 		    SubOperacoes.ApuracaodeCustos, "
        Sql &= "            SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade, "
        Sql &= " 	        SUM(NotasFiscaisXItens.Valor) AS ValorDoProduto, "
        Sql &= " 	        NotasFiscais.EntradaSaida_Id, "
        Sql &= " 	        Sum(ValorIPI) as ValorIPI, "
        Sql &= " 	        Sum(ValorICMS) as ValorICMS, "
        Sql &= " 	        Sum(ValorCOFINS) as ValorCOFINS, "
        Sql &= " 	        Sum(ValorPIS) as ValorPIS  "
        Sql &= " Into #Notas"
        Sql &= " FROM       NotasFiscais INNER JOIN"
        Sql &= "            NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND "
        Sql &= "            NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND "
        Sql &= "            NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND "
        Sql &= "            NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN"
        Sql &= "            SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN"
        Sql &= "            Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id "
        Sql &= "    INNER JOIN GruposDeEstoques "
        Sql &= "     ON Produtos.Grupo = GruposDeEstoques.Grupo_Id "
        Sql &= "    LEFT JOIN (Select Empresa_Id,EndEmpresa_Id,Cliente_Id,EndCliente_Id,EntradaSaida_Id,Serie_Id,Nota_Id,Produto_Id,Cfop_Id, "
        Sql &= "    isnull(sum(case  	        when UPPER(Encargo_Id) = 'IPI'  	        then NotasFiscaisXEncargos.Valor 	        else 0 	        end),0) as ValorIPI,  "
        Sql &= "    isnull(sum(case  	        when UPPER(Encargo_Id) = 'ICMS'  	        then NotasFiscaisXEncargos.Valor 	        else 0 	        end),0) as ValorICMS,  "
        Sql &= "    isnull(sum(case  	        when UPPER(Encargo_Id) = 'COFINS'  	        then NotasFiscaisXEncargos.Valor 	        else 0 	        end),0) as ValorCOFINS,  "
        Sql &= "    isnull(sum(case 	        when UPPER(Encargo_Id) = 'PIS'  	        then NotasFiscaisXEncargos.Valor  	        else 0 	        end),0) as ValorPIS  "
        Sql &= "    from NotasFiscaisXEncargos "
        Sql &= "    Group by  Empresa_Id,EndEmpresa_Id,Cliente_Id,EndCliente_Id,EntradaSaida_Id,Serie_Id,Nota_Id,Produto_Id,Cfop_Id)   Enc    "
        Sql &= "    ON NotasFiscaisXItens.Empresa_Id  = Enc.Empresa_Id      AND NotasFiscaisXItens.EndEmpresa_Id   = Enc.EndEmpresa_Id      "
        Sql &= "    AND NotasFiscaisXItens.Cliente_Id      = Enc.Cliente_Id      AND NotasFiscaisXItens.EndCliente_Id   = Enc.EndCliente_Id      "
        Sql &= "    AND NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id      AND NotasFiscaisXItens.Serie_Id        = Enc.Serie_Id      "
        Sql &= "    AND NotasFiscaisXItens.Nota_Id         = Enc.Nota_Id  AND NotasFiscaisXItens.Produto_Id         = Enc.Produto_Id "
        Sql &= "    AND NotasFiscaisXItens.Cfop_Id         = Enc.Cfop_Id "
        Sql &= " WHERE      (NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"

        If CInt(txtDia.Text) > 0 Then
            Sql &= " 		AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & pMes & ")"
            Sql &= "        AND (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))   = " & DdlAno.SelectedValue & ")"
            Sql &= "        AND isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento) <= '" & DdlAno.SelectedValue & "-" & pMes.ToString("00") & "-" & CInt(txtDia.Text).ToString("00") & "'"
        Else
            Sql &= " 		AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & pMes & ")"
            Sql &= "        AND (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))   = " & DdlAno.SelectedValue & ")"
        End If

        Sql &= "        AND (SubOperacoes.ApuracaoDeCustos > 0)"
        Sql &= "        And NotasFiscais.Situacao = 1"
        Sql &= "        And  GruposDeEstoques.custo = 1"
        Sql &= " GROUP BY	NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscaisXItens.Deposito, NotasFiscaisXItens.EndDeposito, YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), "
        Sql &= "            MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)), NotasFiscaisXItens.Produto_Id, SubOperacoes.ApuracaoDeCustos, NotasFiscaisXItens.DepositoTerceiro, "
        Sql &= "            NotasFiscaisXItens.EndDepositoTerceiro, SubOperacoes.ApuracaodeCustosContrapartida, NotasFiscais.entradaSaida_Id"


        Sql &= Sqla & "; "

        Sql &= " SELECT  Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Placus_Id, Produto_Id, sum(Quantidade) as Quantidade, "
        Sql &= " isnull(sum(case "
        Sql &= " when EntradaSaida_Id = 'E' "
        Sql &= " then ValorDoProduto - (ValorCOFINS + ValorPIS + ValorICMS) "
        Sql &= " else ValorDoProduto "
        Sql &= " end),0) as ValorDoProduto "
        Sql &= " FROM    #Notas "
        Sql &= " Where Placus_Id > 0  "
        Sql &= " GROUP   BY Empresa_Id, EndEmpresa_Id, Placus_Id, Produto_Id, Deposito_Id, EndDeposito_Id, Placus_Id, Produto_Id; "

        Sql &= "  Drop Table #Notas;"


        ChkResultado.Items.Add(New ListItem(" 11 - Capturando dados das Notas Fiscais ContraPartida- Atualizando Tabela de Apuração"))


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare"
            Sql &= " @Exist as varchar(1)"
            Sql &= " set @Exist = (select case "
            Sql &= "       when exists ("
            Sql &= "                 select * from ApuracaoDeCustos "
            Sql &= " Where "
            Sql &= "     Empresa_Id = '" & Dr("Empresa_Id") & "'"
            Sql &= " And EndEmpresa_Id = " & Dr("EndEmpresa_Id")
            Sql &= " And Deposito_Id = '" & Dr("Deposito_Id") & "'"
            Sql &= " And EndDeposito_Id = " & Dr("EndDeposito_Id")
            Sql &= " And Ano_Id = " & DdlAno.SelectedValue
            Sql &= " And Mes_Id = " & pMes
            Sql &= " And Produto_Id = '" & Dr("Produto_Id") & "'"
            Sql &= " And CodigoDeCusto_Id = " & Dr("Placus_Id")

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

            Sql &= " VALUES('" & Dr("Empresa_Id") & "'"
            Sql &= ", " & Dr("EndEmpresa_Id")
            Sql &= ", '" & Dr("Deposito_Id") & "'"
            Sql &= ", " & Dr("EndDeposito_Id")
            Sql &= ", " & DdlAno.SelectedValue
            Sql &= ", " & pMes
            Sql &= ", '" & Dr("Produto_Id") & "'"
            Sql &= ", " & Dr("Placus_Id")
            Sql &= ", ''" 'Empresa Destino
            Sql &= ", 0"  'End Empresa Destino
            Sql &= ", ''" 'Deposito Destino
            Sql &= ", 0"  'End DepositoDestino
            Sql &= ", '" & Dr("Produto_Id") & "'"

            Sql &= ", " & Str(Dr("Quantidade"))
            Sql &= ", 0" '& Str(Dr("ValorDoProduto"))
            Sql &= ", 0" '& Dr("ValorDoFrete").ToString.Replace(",", ".")
            Sql &= ", 0" '& Dr("ValorAuxiliar").ToString.Replace(",", ".")
            Sql &= ", ''" '& Dr("Produto_Id") & "'"
            Sql &= ", 0" '& Dr("CodigoDestino")
            Sql &= ")"

            Sql &= " Else"

            Sql &= "   Update ApuracaoDeCustos set "
            Sql &= "   Quantidade = " & Str(Dr("Quantidade"))
            'Sql &= ",  ValorDoProduto = " & Str(Dr("ValorDoProduto"))

            Sql &= " Where Empresa_Id       = '" & Dr("Empresa_Id") & "'"
            Sql &= "   And EndEmpresa_Id    = " & Dr("EndEmpresa_Id")
            Sql &= "   And Deposito_Id      = '" & Dr("Deposito_Id") & "'"
            Sql &= "   And EndDeposito_Id   = " & Dr("EndDeposito_Id")
            Sql &= "   And Ano_Id           = " & DdlAno.SelectedValue
            Sql &= "   And Mes_Id           = " & pMes
            Sql &= "   And Produto_Id       = '" & Dr("Produto_Id") & "'"
            Sql &= "   And CodigoDeCusto_Id = " & Dr("Placus_Id")

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
                i += 1
                If i = 1 Then
                    ChkResultado.Items.Add(New ListItem(" 12 - Captura dos Dados das Notas Fiscais"))
                    'ChkResultado.Items(11).Selected = True
                End If
            Else
                ChkResultado.Items.Add(New ListItem(" 12 - Erro no Processo de captura dos dados das Notas Fiscais"))
            End If
        Next
    End Sub

    Sub AjustaTotalizadores(ByVal pMes As Integer)
        Dim Fase As Integer

        Sql = " Update ApuracaoDeCustos SET" & vbCrLf & _
              "     Quantidade     = 0" & vbCrLf & _
              "    ,ValorDoProduto = 0" & vbCrLf & _
              "    ,ValorDoFrete   = 0" & vbCrLf & _
              "    ,ValorAuxiliar  = 0" & vbCrLf & _
              "  from PlanoDeCustos PC" & vbCrLf & _
              " inner join ApuracaoDeCustos" & vbCrLf & _
              "    ON ApuracaoDeCustos.CodigoDeCusto_Id = PC.Codigo_Id" & vbCrLf & _
              " WHERE ApuracaoDeCustos.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "   AND ApuracaoDeCustos.Ano_Id = " & DdlAno.SelectedValue & vbCrLf & _
              "   AND ApuracaoDeCustos.Mes_Id = " & pMes & vbCrLf & _
              "   AND PC.FaseDoTotalizador > 0" & vbCrLf
        Array.Add(Sql)

        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
        Else
            ChkResultado.Items.Add(New ListItem("Erro no Zeramento dos Totalizadores"))
        End If

        For Fase = 0 To 5
            Sql = "SELECT AP.Empresa_Id, AP.EndEmpresa_Id, AP.Deposito_Id, AP.EndDeposito_Id, AP.Ano_Id," & vbCrLf & _
                  "       AP.Mes_Id, AP.Produto_Id, AP.CodigoDeCusto_Id, AP.EmpresaDestino_Id," & vbCrLf & _
                  "       AP.EndEmpresaDestino_Id, AP.DepositoDestino_Id, AP.EndDepositoDestino_Id," & vbCrLf & _
                  "       AP.ProdutoDerivado_Id, AP.Etapa, AP.Quantidade, AP.ValorDoProduto," & vbCrLf & _
                  "       AP.ValorDoFrete, AP.ValorAuxiliar, AP.ProdutoDestino, AP.CodigoDestino," & vbCrLf & _
                  "       AP.Reduzido, PC.Totalizador," & vbCrLf & _
                  "       ISNULL(PC.SinalPeso, '') AS SinalPeso," & vbCrLf & _
                  "       ISNULL(PC.SinalValor, '') AS SinalValor" & vbCrLf & _
                  "  FROM ApuracaoDeCustos AP" & vbCrLf & _
                  " INNER JOIN PlanoDeCustos PC" & vbCrLf & _
                  "    ON AP.CodigoDeCusto_Id = PC.Codigo_Id" & vbCrLf & _
                  " WHERE (AP.Empresa_Id      LIKE '" & Left(Empresa(0), 8) & "%')" & vbCrLf & _
                  "   AND (AP.Ano_Id          = " & DdlAno.SelectedValue & ")" & vbCrLf & _
                  "   AND (AP.Mes_Id          = " & pMes & ")" & vbCrLf & _
                  "   AND (PC.Totalizador        > 0)" & vbCrLf & _
                  "   AND (AP.Quantidade + AP.ValorDoProduto + AP.ValorDoFrete <> 0)" & vbCrLf & _
                  "   AND (PC.FaseDoTotalizador  = " & Fase & ")" & vbCrLf & _
                  " ORDER BY AP.Empresa_Id, AP.EndEmpresa_Id, AP.Deposito_Id, AP.EndDeposito_Id," & vbCrLf & _
                  "          AP.Ano_Id, AP.Mes_Id, AP.Produto_Id, AP.CodigoDeCusto_Id, AP.DepositoDestino_Id," & vbCrLf & _
                  "          AP.EndDepositoDestino_Id" & vbCrLf


            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
                Sql = " Declare" & vbCrLf & _
                      " @Exist as varchar(1)" & vbCrLf & _
                      " set @Exist = (select case " & vbCrLf & _
                      "                        when exists (" & vbCrLf & _
                      "                                      select *" & vbCrLf & _
                      "                                        from ApuracaoDeCustos " & vbCrLf & _
                      "                                       Where Empresa_Id       = '" & Dr("Empresa_Id") & "'" & vbCrLf & _
                      "                                         And EndEmpresa_Id    = " & Dr("EndEmpresa_Id") & vbCrLf & _
                      "                                         And Deposito_Id      = '" & Dr("Deposito_Id") & "'" & vbCrLf & _
                      "                                         And EndDeposito_Id   = " & Dr("EndDeposito_Id") & vbCrLf & _
                      "                                         And Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
                      "                                         And Mes_Id           = " & pMes & vbCrLf & _
                      "                                         And Produto_Id       = '" & Dr("Produto_Id") & "'" & vbCrLf & _
                      "                                         And CodigoDeCusto_Id = " & Dr("Totalizador") & vbCrLf & _
                      "                                    )" & vbCrLf & _
                      "                            then 'S'" & vbCrLf & _
                      "                            else 'N'" & vbCrLf & _
                      "                      end) ;" & vbCrLf & _
                      " if @Exist = 'N' " & vbCrLf & _
                      "  INSERT INTO ApuracaoDeCustos ( " & vbCrLf & _
                      "                                Empresa_Id" & vbCrLf & _
                      "                               ,EndEmpresa_Id" & vbCrLf & _
                      "                               ,Deposito_Id" & vbCrLf & _
                      "                               ,EndDeposito_Id" & vbCrLf & _
                      "                               ,Ano_Id" & vbCrLf & _
                      "                               ,Mes_Id" & vbCrLf & _
                      "                               ,Produto_Id" & vbCrLf & _
                      "                               ,CodigoDeCusto_Id" & vbCrLf & _
                      "                               ,EmpresaDestino_Id" & vbCrLf & _
                      "                               ,EndEmpresaDestino_Id" & vbCrLf & _
                      "                               ,DepositoDestino_Id" & vbCrLf & _
                      "                               ,EndDepositoDestino_Id" & vbCrLf & _
                      "                               ,ProdutoDerivado_Id" & vbCrLf & _
                      "                               ,Quantidade" & vbCrLf & _
                      "                               ,ValorDoProduto" & vbCrLf & _
                      "                               ,ValorDoFrete" & vbCrLf & _
                      "                               ,ValorAuxiliar" & vbCrLf & _
                      "                               ,ProdutoDestino" & vbCrLf & _
                      "                               ,CodigoDestino)" & vbCrLf & _
                      " VALUES( '" & Dr("Empresa_Id") & "'" & vbCrLf & _
                      "        , " & Dr("EndEmpresa_Id") & vbCrLf & _
                      "        ,'" & Dr("Deposito_Id") & "'" & vbCrLf & _
                      "        , " & Dr("EndDeposito_Id") & vbCrLf & _
                      "        , " & DdlAno.SelectedValue & vbCrLf & _
                      "        , " & pMes & vbCrLf & _
                      "        ,'" & Dr("Produto_Id") & "'" & vbCrLf & _
                      "        , " & Dr("Totalizador") & vbCrLf & _
                      "        , ''" & vbCrLf & _
                      "        , 0" & vbCrLf & _
                      "        , ''" & vbCrLf & _
                      "        , 0" & vbCrLf & _
                      "        , ''" & vbCrLf & _
                      "        , 0" & vbCrLf & _
                      "        , 0" & vbCrLf & _
                      "        , 0" & vbCrLf & _
                      "        , 0" & vbCrLf & _
                      "        , '" & Dr("ProdutoDestino") & "'" & vbCrLf & _
                      "        , " & Dr("CodigoDestino") & vbCrLf & _
                      "       )" & vbCrLf

                If Dr("SinalPeso") <> "" Or Dr("SinalValor") <> "" Then
                    Sql &= "  Update ApuracaoDeCustos set "

                    If Dr("SinalPeso") <> "" Then
                        Sql &= "  Quantidade = Quantidade " & Dr("SinalPeso") & "(" & Str(Dr("Quantidade")) & ")"
                    End If

                    If Dr("SinalValor") <> "" Then
                        If Dr("SinalPeso") <> "" Then Sql &= ", "

                        Sql &= " ValorDoProduto = ValorDoProduto " & Dr("SinalValor") & " (" & Str(Dr("ValorDoProduto")) & ")"
                        Sql &= ", ValorDoFrete  = ValorDoFrete   " & Dr("SinalValor") & " (" & Str(Dr("ValorDoFrete")) & ")"
                        Sql &= ", ValorAuxiliar = ValorAuxiliar  " & Dr("SinalValor") & " (" & Str(Dr("ValorAuxiliar")) & ")"
                    End If

                    Sql &= " Where Empresa_Id       ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                           "   And EndEmpresa_Id    = " & Dr("EndEmpresa_Id") & vbCrLf & _
                           "   And Deposito_Id      ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                           "   And EndDeposito_Id   = " & Dr("EndDeposito_Id") & vbCrLf & _
                           "   And Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
                           "   And Mes_Id           = " & pMes & vbCrLf & _
                           "   And Produto_Id       ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                           "   And CodigoDeCusto_Id = " & Dr("Totalizador")
                End If

                Array.Add(Sql)

                If Banco.GravaBanco(Array) = True Then
                    Array.Clear()
                Else
                    ChkResultado.Items.Add(New ListItem("Erro Durante os Ajuste dos Totalizadores"))
                End If
            Next
        Next
    End Sub

    Function AjustaProdutoDerivado(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        Sql = "   select case when len(EmpresaDestino_Id) > 0 then EmpresaDestino_Id else Empresa_Id end as Empresa_Id," & vbCrLf & _
              "          case when len(EmpresaDestino_Id) > 0 then EndEmpresaDestino_Id else EndEmpresa_Id end as EndEmpresa_Id," & vbCrLf & _
              "          case when len(DepositoDestino_Id) > 0 then DepositoDestino_Id else Deposito_Id end as Deposito_Id," & vbCrLf & _
              "          case when len(DepositoDestino_Id) > 0 then EndDepositoDestino_Id else EndDeposito_Id end as EndDeposito_Id," & vbCrLf & _
              "          Ano_Id, Mes_Id,Produto_Id as ProdutoOrigem, ProdutoDerivado_Id as Produto_Id, CodigoDestino as CodigoDeCusto_Id, '' as EmpresaDestino_Id, 0 as EndEmpresaDestino_Id, " & vbCrLf & _
              "         '' as DepositoDestino_Id, 0 as EndDepositoDestino_Id, '' as ProdutoDerivado_Id, isnull(Etapa,0) as Etapa, 0 as Quantidade, ValorDoProduto, ValorDoFrete, ValorAuxiliar, ProdutoDestino, " & vbCrLf & _
              "         0 as CodigoDestino, Reduzido" & vbCrLf & _
              "   from ApuracaoDeCustos" & vbCrLf & _
              "  Where Empresa_Id    like '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "    and Ano_id        = " & DdlAno.SelectedValue & vbCrLf & _
              "    And Mes_id        = " & pMes & vbCrLf & _
              "    And ProdutoDerivado_Id > 0" & vbCrLf & _
              "    and isnull(codigodestino,0) > 0" & vbCrLf & _
              "    And CodigodeCusto_Id not in (200, 600)" & vbCrLf

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
                  "                                         And Ano_Id                = " & Dr("Ano_Id") & vbCrLf & _
                  "                                         And Mes_Id                = " & Dr("Mes_Id") & vbCrLf & _
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                         And CodigoDeCusto_Id      = " & Dr("CodigoDeCusto_Id") & vbCrLf & _
                  "                                         And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "                                         And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "                                         And ProdutoDerivado_Id    ='" & Dr("ProdutoOrigem") & "'" & vbCrLf & _
                  "                                    )" & vbCrLf & _
                  "                          then 'S'" & vbCrLf & _
                  "                          else 'N'" & vbCrLf & _
                  "                      end) ;" & vbCrLf

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
            Sql &= " VALUES('" & Dr("Empresa_Id") & "'"
            Sql &= ", " & Dr("EndEmpresa_Id")
            Sql &= ",'" & Dr("Deposito_Id") & "'"
            Sql &= ", " & Dr("EndDeposito_Id")
            Sql &= ", " & Dr("Ano_Id")
            Sql &= ", " & Dr("Mes_Id")
            Sql &= ",'" & Dr("Produto_Id") & "'"
            Sql &= ", " & Dr("CodigoDeCusto_Id")
            Sql &= ",'" & Dr("EmpresaDestino_Id") & "'"
            Sql &= ", " & Dr("EndEmpresaDestino_Id")
            Sql &= ",'" & Dr("DepositoDestino_Id") & "'"
            Sql &= ", " & Dr("EndDepositoDestino_Id")
            Sql &= ",'" & Dr("ProdutoOrigem") & "'"
            Sql &= ", " & Dr("Etapa")
            Sql &= ", " & Dr("Quantidade")
            Sql &= ", " & Dr("ValorDoProduto").ToString.Replace(",", ".")
            Sql &= ", " & Dr("ValorDoFrete").ToString.Replace(",", ".")
            Sql &= ", " & Dr("ValorAuxiliar").ToString.Replace(",", ".")
            Sql &= ",'" & Dr("ProdutoDestino") & "'"
            Sql &= ", " & Dr("CodigoDestino")
            Sql &= ",'" & Dr("Reduzido") & "'"
            Sql &= ")"

            Sql &= "  Update ApuracaoDeCustos set "
            Sql &= "         ValorDoProduto     =  " & Dr("ValorDoProduto").ToString.Replace(",", ".")
            Sql &= "        ,ValorDoFrete       =  " & Dr("ValorDoFrete").ToString.Replace(",", ".")
            Sql &= " Where Empresa_Id            ='" & Dr("Empresa_Id") & "'"
            Sql &= "   And EndEmpresa_Id         = " & Dr("EndEmpresa_Id")
            Sql &= "   And Deposito_Id           ='" & Dr("Deposito_Id") & "'"
            Sql &= "   And EndDeposito_Id        = " & Dr("EndDeposito_Id")
            Sql &= "   And Ano_Id                = " & Dr("Ano_Id")
            Sql &= "   And Mes_Id                = " & Dr("Mes_Id")
            Sql &= "   And Produto_Id            ='" & Dr("Produto_Id") & "'"
            Sql &= "   And CodigoDeCusto_Id      = " & Dr("CodigoDeCusto_Id")
            Sql &= "   And EmpresaDestino_Id     ='" & Dr("EmpresaDestino_Id") & "'"
            Sql &= "   And EndEmpresaDestino_Id  = " & Dr("EndEmpresaDestino_Id")
            Sql &= "   And DepositoDestino_Id    ='" & Dr("DepositoDestino_Id") & "'"
            Sql &= "   And EndDepositoDestino_Id = " & Dr("EndDepositoDestino_Id")
            Sql &= "   And ProdutoDerivado_Id    ='" & Dr("ProdutoOrigem") & "'"

            Array.Add(Sql)

        Next


        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
        Else
            erros += 1
        End If
        Return erros
    End Function

    Function AjustaConsumoXProducao(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        Sql = "select distinct ProdutoOrigem_Id, CodigoCustoOrigem_Id " & vbCrLf & _
              "  from ConsumoxProducao "

        For Each Row As DataRow In Banco.ConsultaDataSet(Sql, "ProdutosOrigem").Tables(0).Rows
            Sql = "	declare" & vbCrLf & _
                  "	@Soma as numeric(18,9)" & vbCrLf & _
                  "	select ACO.Empresa_Id, 		   ACO.EndEmpresa_id," & vbCrLf & _
                  "		   ACO.Deposito_Id,		   ACO.EndDeposito_Id," & vbCrLf & _
                  "        ACO.Ano_Id," & vbCrLf & _
                  "        ACO.Mes_Id," & vbCrLf & _
                  "		   ACO.produto_Id as ProdutoOrigem," & vbCrLf & _
                  "		   ACO.Quantidade as QuantidadeOrigem," & vbCrLf & _
                  "		   ACO.ValorDoProduto as VlrProdutoOrigem," & vbCrLf & _
                  "		   ACD.produto_Id as ProdutoDestino," & vbCrLf & _
                  "          ACD.CodigoDeCusto_Id as CodigoDeCustoDestino," & vbCrLf & _
                  "		   ACD.Quantidade as QuantidadeDestino," & vbCrLf & _
                  "		   TPM.ValorOficial as ValorDeMercado," & vbCrLf & _
                  "		  (ACD.Quantidade/TPM.BaseDeCalculo)*TPM.ValorOficial as VlrTotal," & vbCrLf & _
                  "		  convert(numeric(18,9),0) as Soma" & vbCrLf & _
                  "	  into #Temp" & vbCrLf & _
                  "	  from ApuracaoDeCustos ACO" & vbCrLf & _
                  "	 Inner Join(select AC.Empresa_Id," & vbCrLf & _
                  "					   AC.EndEmpresa_Id," & vbCrLf & _
                  "					   AC.Ano_Id," & vbCrLf & _
                  "					   AC.Mes_Id," & vbCrLf & _
                  "					   AC.Produto_Id," & vbCrLf & _
                  "                      AC.CodigoDeCusto_Id," & vbCrLf & _
                  "					   AC.Quantidade" & vbCrLf & _
                  "				  from ApuracaoDeCustos AC " & vbCrLf & _
                  "                Inner Join ConsumoxProducao CP " & vbCrLf & _
                  "                   on CP.ProdutoOrigem_Id    ='" & Row("ProdutoOrigem_Id") & "'" & vbCrLf & _
                  "				   and CP.CodigoCustoOrigem_Id  = " & Row("CodigoCustoOrigem_Id") & vbCrLf & _
                  "				   and CP.ProdutoDestino_Id     = AC.Produto_Id" & vbCrLf & _
                  "				   and CP.CodigoCustoDestino_Id = AC.CodigoDeCusto_Id" & vbCrLf & _
                  "				) ACD " & vbCrLf & _
                  "		on ACD.Empresa_Id    = ACO.Empresa_Id" & vbCrLf & _
                  "	   and ACD.EndEmpresa_Id = ACO.EndEmpresa_Id" & vbCrLf & _
                  "	   and ACD.Ano_Id        = ACO.Ano_Id" & vbCrLf & _
                  "	   and ACD.Mes_Id        = ACO.Mes_Id" & vbCrLf & _
                  "	 Inner Join (SELECT TPM1.Empresa_Id,    TPM1.EndEmpresa_Id," & vbCrLf & _
                  "                       TPM1.Deposito_Id,   TPM1.EndDeposito_Id," & vbCrLf & _
                  "                       TPM1.Produto_Id,    TPM1.Data_Id," & vbCrLf & _
                  "                       TPM1.ValorOficial,  TPM1.ValorMoeda," & vbCrLf & _
                  "                       TPM1.BaseDeCalculo" & vbCrLf & _
                  "                  FROM TabelaDePrecosDeMercado TPM1" & vbCrLf & _
                  "                 Inner Join (SELECT Empresa_Id," & vbCrLf & _
                  "									 EndEmpresa_Id," & vbCrLf & _
                  "									 Deposito_Id," & vbCrLf & _
                  "									 EndDeposito_Id," & vbCrLf & _
                  "									 Produto_Id," & vbCrLf & _
                  "									 max(Data_Id) as Data_Id" & vbCrLf & _
                  "                                from TabelaDePrecosDeMercado" & vbCrLf & _
                  "                               Group by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id" & vbCrLf & _
                  "                             ) TPM2" & vbCrLf & _
                  "                     on TPM1.Empresa_Id     = TPM2.Empresa_Id" & vbCrLf & _
                  "                    and TPM1.EndEmpresa_Id  = TPM2.EndEmpresa_Id" & vbCrLf & _
                  "                    and TPM1.Deposito_Id    = TPM2.Deposito_Id" & vbCrLf & _
                  "                    and TPM1.EndDeposito_Id = TPM2.EndDeposito_Id" & vbCrLf & _
                  "                    and TPM1.Produto_Id     = TPM2.Produto_Id" & vbCrLf & _
                  "                    and TPM1.Data_Id        = TPM2.Data_Id" & vbCrLf & _
                  "                ) TPM" & vbCrLf & _
                  "		on ACO.Empresa_Id     = TPM.Empresa_Id" & vbCrLf & _
                  "	   and ACO.EndEmpresa_id  = TPM.EndEmpresa_id" & vbCrLf & _
                  "	   and ACO.Deposito_Id    = TPM.Deposito_Id" & vbCrLf & _
                  "	   And ACO.EndDeposito_Id = TPM.EndDeposito_Id" & vbCrLf & _
                  "	   And ACO.Ano_Id         = year(TPM.Data_Id)" & vbCrLf & _
                  "	   And ACO.Mes_Id         = Month(TPM.Data_Id)" & vbCrLf & _
                  "	   And ACD.Produto_Id     = TPM.Produto_Id " & vbCrLf & _
                  "	 where ACO.Empresa_Id    like '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
                  "    AND ACO.Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
                  "	   and ACO.Mes_Id           = " & pMes & vbCrLf & _
                  "	   And ACO.Produto_Id       ='" & Row("ProdutoOrigem_Id") & "'" & vbCrLf & _
                  "	   and ACO.CodigoDeCusto_Id = " & Row("CodigoCustoOrigem_Id") & vbCrLf & _
                  "	set @Soma =(Select Sum(VlrTotal) from #Temp)" & vbCrLf & _
                  "	Update #Temp set" & vbCrLf & _
                  "	  Soma = @Soma" & vbCrLf & _
                  "	Select Empresa_Id," & vbCrLf & _
                  "		   EndEmpresa_id," & vbCrLf & _
                  "		   Deposito_Id," & vbCrLf & _
                  "		   EndDeposito_Id," & vbCrLf & _
                  "        Ano_Id," & vbCrLf & _
                  "        Mes_Id," & vbCrLf & _
                  "		   ProdutoOrigem," & vbCrLf & _
                  "		   QuantidadeOrigem," & vbCrLf & _
                  "		   VlrProdutoOrigem," & vbCrLf & _
                  "		   ProdutoDestino," & vbCrLf & _
                  "        CodigoDeCustoDestino," & vbCrLf & _
                  "		   QuantidadeDestino," & vbCrLf & _
                  "		   ValorDeMercado," & vbCrLf & _
                  "		   VlrTotal," & vbCrLf & _
                  "		   (VlrTotal * 100) / Soma as Percentual," & vbCrLf & _
                  "		   (VlrProdutoOrigem *  ((VlrTotal * 100) / Soma)) / 100 as CustoDoProdutoDestino " & vbCrLf & _
                  "	   from #Temp" & vbCrLf


            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows

                Sql = "  Update ApuracaoDeCustos set "
                Sql &= "         ValorDoProduto     =  " & Dr("CustoDoProdutoDestino").ToString.Replace(",", ".")
                Sql &= " Where Empresa_Id            ='" & Dr("Empresa_Id") & "'"
                Sql &= "   And EndEmpresa_Id         = " & Dr("EndEmpresa_Id")
                Sql &= "   And Deposito_Id           ='" & Dr("Deposito_Id") & "'"
                Sql &= "   And EndDeposito_Id        = " & Dr("EndDeposito_Id")
                Sql &= "   And Ano_Id                = " & Dr("Ano_Id")
                Sql &= "   And Mes_Id                = " & Dr("Mes_Id")
                Sql &= "   And Produto_Id            ='" & Dr("ProdutoDestino") & "'"
                Sql &= "   And CodigoDeCusto_Id      = " & Dr("CodigoDeCustoDestino")
                Sql &= "   And EmpresaDestino_Id     =''"
                Sql &= "   And EndEmpresaDestino_Id  = 0"
                Sql &= "   And DepositoDestino_Id    =''"
                Sql &= "   And EndDepositoDestino_Id = 0"
                Sql &= "   And ProdutoDerivado_Id    =''"

                Array.Add(Sql)
            Next
        Next

        If Banco.GravaBanco(Array) = True Then
            Array.Clear()
        Else
            erros += 1
        End If
        Return erros
    End Function

    Function AjustaCustosDeSaidas(ByVal pMes As Integer) As Integer

        Dim erros As Integer

        Sql = "SELECT Empresa_Id, EndEmpresa_ID, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, " & vbCrLf & _
              "       CodigoDeCusto_Id, DepositoDestino_ID, EndDepositoDestino_Id, ProdutoDerivado_Id, " & vbCrLf & _
              "       Convert(Decimal(18,9),((ValorDoProduto + ValorDoFrete) / Quantidade)) as Medio  " & vbCrLf & _
              "  FROM ApuracaoDeCustos" & vbCrLf & _
              " Where Empresa_ID like '" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "   And Ano_Id           = " & DdlAno.SelectedValue & vbCrLf & _
              "	  And Mes_ID           = " & pMes & vbCrLf & _
              "   And CodigoDeCusto_Id = 495" & vbCrLf & _
              "   And Quantidade       > 0" & vbCrLf & _
              " Order by Empresa_Id, EndEmpresa_ID, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, Produto_Id, " & vbCrLf & _
              "          CodigoDeCusto_Id, DepositoDestino_ID, EndDepositoDestino_Id, ProdutoDerivado_Id" & vbCrLf



        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = "  Update ApuracaoDeCustos set " & vbCrLf & _
                  "    ValorDoProduto = Quantidade * " & Dr("Medio").ToString.Replace(",", ".") & vbCrLf & _
                  " Where Empresa_Id       ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "   And EndEmpresa_Id    = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "   And Deposito_Id      ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "   And EndDeposito_Id   = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "   And Ano_Id           = " & Dr("Ano_Id") & vbCrLf & _
                  "   And Mes_Id           = " & Dr("Mes_Id") & vbCrLf & _
                  "   And Produto_Id       ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "   And CodigoDeCusto_Id > 495" & vbCrLf

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
            Else
                erros += 1
            End If
        Next
        Return erros
    End Function

    Function AjustaTransferencias(ByVal pMes As Integer) As Integer

        Dim erros As Integer

        Sql = "SELECT ApuracaoDeCustos.Empresa_Id," & vbCrLf & _
              "       ApuracaoDeCustos.EndEmpresa_Id," & vbCrLf & _
              "       ApuracaoDeCustos.Deposito_Id," & vbCrLf & _
              "       ApuracaoDeCustos.EndDeposito_Id," & vbCrLf & _
              "       ApuracaoDeCustos.Ano_Id," & vbCrLf & _
              "       ApuracaoDeCustos.Mes_Id," & vbCrLf & _
              "       ApuracaoDeCustos.Produto_Id," & vbCrLf & _
              "       ApuracaoDeCustos.CodigoDeCusto_Id," & vbCrLf & _
              "       ApuracaoDeCustos.DepositoDestino_Id as EmpresaDestino_Id," & vbCrLf & _
              "       ApuracaoDeCustos.EndDepositoDestino_Id as EndEmpresaDestino_Id," & vbCrLf & _
              "       ApuracaoDeCustos.DepositoDestino_Id," & vbCrLf & _
              "       ApuracaoDeCustos.EndDepositoDestino_Id," & vbCrLf & _
              "       ApuracaoDeCustos.ProdutoDerivado_Id," & vbCrLf & _
              "       ApuracaoDeCustos.Quantidade," & vbCrLf & _
              "       ApuracaoDeCustos.ValorDoProduto," & vbCrLf & _
              "       ApuracaoDeCustos.ValorDoFrete," & vbCrLf & _
              "       ApuracaoDeCustos.ValorAuxiliar," & vbCrLf & _
              "       ApuracaoDeCustos.ProdutoDestino," & vbCrLf & _
              "       ApuracaoDeCustos.CodigoDestino " & vbCrLf & _
              "  FROM ApuracaoDeCustos " & vbCrLf & _
              " INNER JOIN PlanoDeCustos" & vbCrLf & _
              "    ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id" & vbCrLf & _
              " WHERE ApuracaoDeCustos.Empresa_ID           like'" & Left(Empresa(0), 8) & "%'" & vbCrLf & _
              "   and ApuracaoDeCustos.Ano_id                  = " & DdlAno.SelectedValue & vbCrLf & _
              "   And ApuracaoDeCustos.Mes_id                  = " & pMes & vbCrLf & _
              "   And ApuracaoDeCustos.CodigoDeCusto_Id        > 500" & vbCrLf & _
              "   And PlanoDeCustos.Classe                     = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf & _
              "   and isnull(ApuracaoDeCustos.codigodestino,0) > 0" & vbCrLf


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Sql = " Declare" & vbCrLf & _
                  " @Exist as varchar(1)" & vbCrLf & _
                  " set @Exist = (select case " & vbCrLf & _
                  "                        when exists (" & vbCrLf & _
                  "                                      select 1 " & vbCrLf & _
                  "                                        from ApuracaoDeCustos " & vbCrLf & _
                  "                                       Where Empresa_Id            ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresa_Id         = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "                                         And Deposito_Id           ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "                                         And EndDeposito_Id        = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "                                         And Ano_Id                = " & Dr("Ano_Id") & vbCrLf & _
                  "                                         And Mes_Id                = " & Dr("Mes_Id") & vbCrLf & _
                  "                                         And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "                                         And CodigoDeCusto_Id      = " & Dr("CodigoDestino") & vbCrLf & _
                  "                                         And EmpresaDestino_Id     ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "                                         And EndEmpresaDestino_Id  = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "                                         And DepositoDestino_Id    ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "                                         And EndDepositoDestino_Id = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "                                         And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  "                                    )" & vbCrLf & _
                  "                          then 'S'" & vbCrLf & _
                  "                          else 'N'" & vbCrLf & _
                  "                     end) ;" & vbCrLf & _
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
                  " VALUES('" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  ",'" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  ", " & Dr("Ano_Id") & vbCrLf & _
                  ", " & Dr("Mes_Id") & vbCrLf & _
                  ",'" & Dr("Produto_Id") & "'" & vbCrLf & _
                  ", " & Dr("CodigoDestino") & vbCrLf & _
                  ",'" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndEmpresa_Id") & vbCrLf & _
                  ",'" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  ", " & Dr("EndDeposito_Id") & vbCrLf & _
                  ",'" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf & _
                  ", " & Replace(Dr("Quantidade"), ",", ".") & vbCrLf & _
                  ", " & Replace(Dr("ValorDoProduto"), ",", ".") & vbCrLf & _
                  ", " & Replace(Dr("ValorDoFrete"), ",", ".") & vbCrLf & _
                  ", " & Replace(Dr("ValorAuxiliar"), ",", ".") & vbCrLf & _
                  ",'" & Dr("ProdutoDestino") & "'" & vbCrLf & _
                  ", " & Dr("CodigoDeCusto_Id") & vbCrLf & _
                  ")" & vbCrLf & _
                  "  Update ApuracaoDeCustos set " & vbCrLf & _
                  "         ValorDoProduto     =  " & Replace(Dr("ValorDoProduto"), ",", ".") & vbCrLf & _
                  "        ,ValorDoFrete       =  " & Replace(Dr("ValorDoFrete"), ",", ".") & vbCrLf & _
                  " Where Empresa_Id            ='" & Dr("EmpresaDestino_Id") & "'" & vbCrLf & _
                  "   And EndEmpresa_Id         = " & Dr("EndEmpresaDestino_Id") & vbCrLf & _
                  "   And Deposito_Id           ='" & Dr("DepositoDestino_Id") & "'" & vbCrLf & _
                  "   And EndDeposito_Id        = " & Dr("EndDepositoDestino_Id") & vbCrLf & _
                  "   And Ano_Id                = " & Dr("Ano_Id") & vbCrLf & _
                  "   And Mes_Id                = " & Dr("Mes_Id") & vbCrLf & _
                  "   And Produto_Id            ='" & Dr("Produto_Id") & "'" & vbCrLf & _
                  "   And CodigoDeCusto_Id      = " & Dr("CodigoDestino") & vbCrLf & _
                  "   And EmpresaDestino_Id     ='" & Dr("Empresa_Id") & "'" & vbCrLf & _
                  "   And EndEmpresaDestino_Id  = " & Dr("EndEmpresa_Id") & vbCrLf & _
                  "   And DepositoDestino_Id    ='" & Dr("Deposito_Id") & "'" & vbCrLf & _
                  "   And EndDepositoDestino_Id = " & Dr("EndDeposito_Id") & vbCrLf & _
                  "   And ProdutoDerivado_Id    ='" & Dr("ProdutoDerivado_Id") & "'" & vbCrLf

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
            Else
                erros += 1
            End If
        Next
        Return erros
    End Function

    Function CalculoCircular(ByVal pMes As Integer) As Integer
        Dim erros As Integer = 0
        Sql = "  SELECT Origem.Empresa_Id, "
        Sql &= " 	    Origem.EndEmpresa_Id, "
        Sql &= " 	    Origem.Deposito_Id, "
        Sql &= "        Origem.EndDeposito_Id, "
        Sql &= "        Origem.Produto_Id,"
        Sql &= "        Origem.CodigoDeCusto_Id, "
        Sql &= "        Origem.Ano_Id, "
        Sql &= "        Origem.Mes_Id, "
        Sql &= "        Origem.EmpresaDestino_Id, "
        Sql &= "        Origem.EndEmpresaDestino_Id, "
        Sql &= "        Origem.DepositoDestino_Id, "
        Sql &= "        Origem.EndDepositoDestino_Id, "
        Sql &= "        CASE WHEN Origem.ProdutoDerivado_Id = '000000' THEN '' ELSE Origem.ProdutoDerivado_Id END AS ProdutoDerivado_Id,"
        Sql &= " 	    Origem.CodigoDestino,"
        Sql &= " 	    Origem.ValorDoProduto as ValorOrigem, "
        Sql &= " 	    Origem.ValorDoFrete as FreteOrigem, "
        Sql &= " 	    isnull(Destino.ValorDoProduto, 0) AS ValorDestino"
        Sql &= " FROM   ApuracaoDeCustos AS Origem LEFT OUTER JOIN ApuracaoDeCustos AS Destino ON"
        Sql &= "            Origem.CodigoDeCusto_Id = Destino.CodigoDestino"
        Sql &= "        AND Origem.Produto_Id = Destino.ProdutoDerivado_Id  "
        Sql &= "        AND Origem.EndDeposito_Id = Destino.EndDepositoDestino_Id "
        Sql &= "        AND Origem.Deposito_Id = Destino.DepositoDestino_Id "
        Sql &= "        AND Origem.EndEmpresa_Id = Destino.EndEmpresaDestino_Id "
        Sql &= "        AND Origem.Empresa_Id = Destino.EmpresaDestino_Id "
        Sql &= "        AND Origem.EmpresaDestino_Id = Destino.Empresa_Id "
        Sql &= "        AND Origem.EndEmpresaDestino_Id = Destino.EndEmpresa_Id "
        Sql &= "        AND Origem.DepositoDestino_Id = Destino.Deposito_Id "
        Sql &= "        AND Origem.EndDepositoDestino_Id = Destino.EndDeposito_Id "
        Sql &= "        AND Origem.ProdutoDerivado_Id = Destino.Produto_Id "
        Sql &= "        AND Origem.Ano_Id = Destino.Ano_Id "
        Sql &= "        AND Origem.Mes_Id = Destino.Mes_Id "
        Sql &= "        AND Origem.CodigoDestino = Destino.CodigoDeCusto_Id"
        Sql &= " WHERE	(Origem.Empresa_Id like '" & Left(Empresa(0), 8) & "%')"
        Sql &= " 	    And   (Origem.Ano_Id = " & DdlAno.SelectedValue & ")"
        Sql &= " 	    And   (Origem.Mes_Id = " & pMes & ")"
        Sql &= " 	    And   (Origem.CodigoDeCusto_Id > 495)"
        Sql &= "        And   (Origem.CodigoDestino <> 0)"
        Sql &= "        And   (Origem.ValorDoProduto > 0)"
        Sql &= " 	    And   (Origem.EmpresaDestino_Id <> '')"
        Sql &= " 	    And   (Origem.DepositoDestino_ID <> '')"
        Sql &= "        And   (Origem.ValorDoProduto <> Destino.ValorDoProduto)"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            Circular = "S"

            Sql = " Declare"
            Sql &= " @Exist as varchar(1)"
            Sql &= " set @Exist = (select case "
            Sql &= "       when exists ("
            Sql &= "                 select * from ApuracaoDeCustos "
            Sql &= " Where "
            Sql &= "     Empresa_Id = '" & Dr("EmpresaDestino_Id") & "'"
            Sql &= " And EndEmpresa_Id = " & Dr("EndEmpresaDestino_Id")
            Sql &= " And Deposito_Id = '" & Dr("DepositoDestino_Id") & "'"
            Sql &= " And EndDeposito_Id = " & Dr("EndDepositoDestino_Id")
            Sql &= " And Ano_Id = " & Dr("Ano_Id")
            Sql &= " And Mes_Id = " & Dr("Mes_Id")
            Sql &= " And Produto_Id = '" & Dr("ProdutoDerivado_Id") & "'"
            Sql &= " And CodigoDeCusto_Id = " & Dr("CodigoDestino")

            Sql &= " And EmpresaDestino_Id = '" & Dr("Empresa_Id") & "'"
            Sql &= " And EndEmpresaDestino_Id = " & Dr("EndEmpresa_Id")
            Sql &= " And DepositoDestino_Id = '" & Dr("Deposito_Id") & "'"
            Sql &= " And EndDepositoDestino_Id = " & Dr("EndDeposito_Id")
            Sql &= " And ProdutoDerivado_Id = '" & Dr("Produto_Id") & "'"

            Sql &= ")"
            Sql &= "            then  'S'"
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

            Sql &= ", Etapa"
            Sql &= ", Quantidade"
            Sql &= ", ValorDoProduto"
            Sql &= ", ValorDoFrete"
            Sql &= ", ValorAuxiliar"
            Sql &= ", ProdutoDestino"
            Sql &= ", CodigoDestino"
            Sql &= ", Reduzido)"

            Sql &= " VALUES('" & Dr("EmpresaDestino_Id") & "'"
            Sql &= ", " & Dr("EndEmpresaDestino_Id")
            Sql &= ", '" & Dr("DepositoDestino_Id") & "'"
            Sql &= ", " & Dr("EndDepositoDestino_Id")
            Sql &= ", " & Dr("Ano_Id")
            Sql &= ", " & Dr("Mes_Id")
            Sql &= ", '" & Dr("ProdutoDerivado_Id") & "'"
            Sql &= ", " & Dr("CodigoDestino")

            Sql &= ", '" & Dr("Empresa_Id") & "'"
            Sql &= ", " & Dr("EndEmpresa_Id")
            Sql &= ", '" & Dr("Deposito_Id") & "'"
            Sql &= ", " & Dr("EndDeposito_Id")
            Sql &= ", '" & Dr("Produto_Id") & "'"

            Sql &= ", 1"    'Etapa
            Sql &= ", 0"    'Quantidade
            Sql &= ", " & Dr("ValorOrigem").ToString.Replace(",", ".")
            Sql &= ", " & Dr("FreteOrigem").ToString.Replace(",", ".")
            Sql &= ", 0"    'Valor Auxiliar

            Sql &= ", ''" '& Dr("ProdutoDestino") & "'"
            Sql &= ", " & Dr("CodigoDeCusto_Id")
            Sql &= ", ''"
            Sql &= ")"


            Sql &= "  Update ApuracaoDeCustos set "

            Sql &= "         ValorDoProduto = " & Dr("ValorOrigem").ToString.Replace(",", ".")
            Sql &= ",        ValorDoFrete = " & Dr("FreteOrigem").ToString.Replace(",", ".")

            Sql &= " Where "
            Sql &= "     Empresa_Id = '" & Dr("EmpresaDestino_Id") & "'"
            Sql &= " And EndEmpresa_Id = " & Dr("EndEmpresaDestino_Id")
            Sql &= " And Deposito_Id = '" & Dr("DepositoDestino_Id") & "'"
            Sql &= " And EndDeposito_Id = " & Dr("EndDepositoDestino_Id")
            Sql &= " And Ano_Id = " & Dr("Ano_Id")
            Sql &= " And Mes_Id = " & Dr("Mes_Id")
            Sql &= " And Produto_Id = '" & Dr("ProdutoDerivado_Id") & "'"
            Sql &= " And CodigoDeCusto_Id = " & Dr("CodigoDestino")

            Sql &= " And EmpresaDestino_Id = '" & Dr("Empresa_Id") & "'"
            Sql &= " And EndEmpresaDestino_Id = " & Dr("EndEmpresa_Id")
            Sql &= " And DepositoDestino_Id = '" & Dr("Deposito_Id") & "'"
            Sql &= " And EndDepositoDestino_Id = " & Dr("EndDeposito_Id")
            Sql &= " And ProdutoDerivado_Id = '" & Dr("Produto_Id") & "'"

            Array.Add(Sql)

            If Banco.GravaBanco(Array) = True Then
                Array.Clear()
            Else
                erros += 1
            End If

        Next
        Return erros
    End Function


    Private Sub RateioDosCustosDeProducao_Soja(ByVal pMes As Integer)
        Dim sql As String = "  Update ApuracaoDeCustos "
        sql &= "    Set     ValorAuxiliar = ((ApuracaoDeCustos.Quantidade * TabelaDePrecosDeMercado.ValorOficial) / TabelaDePrecosDeMercado.BaseDeCalculo)"
        sql &= "    FROM    ApuracaoDeCustos INNER JOIN"
        sql &= "            TabelaDePrecosDeMercado ON ApuracaoDeCustos.Empresa_Id = TabelaDePrecosDeMercado.Empresa_Id AND "
        sql &= "            ApuracaoDeCustos.EndEmpresa_Id = TabelaDePrecosDeMercado.EndEmpresa_Id And"
        sql &= "            ApuracaoDeCustos.Deposito_Id = TabelaDePrecosDeMercado.Deposito_Id And"
        sql &= "            ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id And"
        sql &= "            ApuracaoDeCustos.Produto_Id = TabelaDePrecosDeMercado.Produto_Id And ApuracaoDeCustos.Ano_Id = Year(TabelaDePrecosDeMercado.Data_Id) And"
        sql &= "            ApuracaoDeCustos.Mes_Id = Month(TabelaDePrecosDeMercado.Data_Id)"
        sql &= "    WHERE   ApuracaoDeCustos.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "            And Ano_Id = " & DdlAno.SelectedValue
        sql &= "            And Mes_Id = " & pMes
        sql &= "            And (ApuracaoDeCustos.CodigoDeCusto_Id IN (301, 502)) AND "
        sql &= "                (ApuracaoDeCustos.Produto_Id IN (101010003, 102010001, 102020001, 102030001))"

        Banco.GravaBanco(sql)

        sql = "  Update ApuracaoDeCustos"
        sql &= " Set    ValorAuxiliar = Consulta.Valor"
        sql &= " From ("

        sql &= " SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id, '101010003' as Produto_Id, 502 as CodigoDeCusto_Id,"
        sql &= "        Sum (ValorAuxiliar) as Valor"
        sql &= " FROM   ApuracaoDeCustos as AC"
        sql &= " WHERE  AC.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "            And Ano_Id = " & DdlAno.SelectedValue
        sql &= "            And Mes_Id = " & pMes
        sql &= "            And (CodigoDeCusto_Id IN (301 )) AND (Produto_Id IN (102010001, 102020001, 102030001))"
        sql &= " Group By Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id,  CodigoDeCusto_Id"

        sql &= " ) as  Consulta Inner join ApuracaoDeCustos ON"
        sql &= "       Consulta.Empresa_Id = ApuracaoDeCustos.Empresa_Id And"
        sql &= "       Consulta.EndEmpresa_Id = ApuracaoDeCustos.EndEmpresa_Id And"
        sql &= "       Consulta.Deposito_Id = ApuracaoDeCustos.Deposito_Id And"
        sql &= "       Consulta.EndDeposito_Id = ApuracaoDeCustos.EndDeposito_Id And"
        sql &= "       Consulta.Ano_Id = ApuracaoDeCustos.Ano_Id And"
        sql &= "       Consulta.Mes_Id = ApuracaoDeCustos.Mes_Id And"
        sql &= "       Consulta.Produto_Id = ApuracaoDeCustos.Produto_Id And"
        sql &= "       Consulta.CodigoDeCusto_Id = ApuracaoDeCustos.CodigoDeCusto_Id"

        Banco.GravaBanco(sql)

        sql = "  Update ApuracaoDeCustos"
        sql &= " Set    ValorDoProduto =  (Consulta.Valor * (ValorAuxiliar * 100 / Consulta.Auxiliar)) / 100"

        sql &= " From ("

        sql &= " SELECT Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Ano_Id, Mes_Id,  Produto_Id,  CodigoDeCusto_Id,"
        sql &= "        ValorDoProduto as Valor, ValorAuxiliar AS Auxiliar"

        sql &= " FROM   ApuracaoDeCustos as AC"
        sql &= " WHERE  AC.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%' "
        sql &= "        And Ano_Id = " & DdlAno.SelectedValue
        sql &= "        And Mes_Id = " & pMes
        sql &= "		And (CodigoDeCusto_Id IN (502 ))"
        sql &= "		And (Produto_Id       IN (101010003))"

        sql &= " ) as Consulta Inner join ApuracaoDeCustos On"
        sql &= "        Consulta.Empresa_Id = ApuracaoDeCustos.Empresa_Id And"
        sql &= "        Consulta.EndEmpresa_Id = ApuracaoDeCustos.EndEmpresa_Id And"
        sql &= "        Consulta.Deposito_Id = ApuracaoDeCustos.Deposito_Id And"
        sql &= "        Consulta.EndDeposito_Id = ApuracaoDeCustos.EndDeposito_Id And"
        sql &= "        Consulta.Ano_Id = ApuracaoDeCustos.Ano_Id And"
        sql &= "        Consulta.Mes_Id = ApuracaoDeCustos.Mes_Id And"
        sql &= "        ApuracaoDeCustos.CodigoDeCusto_Id = 301"
        sql &= " Where (ApuracaoDeCustos.Produto_Id IN ( 102010001, 102020001, 102030001))   "

        Banco.GravaBanco(sql)


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

    Sub Contabiliza(ByVal pMes As Integer)
        Dim AnoC As Integer
        AnoC = DdlAno.SelectedValue

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
              " Update #Temp  " & vbCrLf & _
              "    Set Produto_Id = ''" & vbCrLf & _
              "   From #temp " & vbCrLf & _
              "   Left join PlanoDeContas PC  " & vbCrLf & _
              "     on #Temp.Empresa_Id    = PC.Empresa_Id  " & vbCrLf & _
              "    And #Temp.EndEmpresa_Id = PC.EndEmpresa_Id  " & vbCrLf & _
              "    AND #Temp.Conta_Id      = PC.Conta_Id " & vbCrLf & _
              "  Where PC.Produto = 'N'; " & vbCrLf & _
              "                                                        " & vbCrLf & _
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        If Valida() Then

            If Funcoes.VerificaPermissao("CapturaDeDados", "RELATORIO") Then

                Dim crpt As New ReportDocument()

                Try
                    Dim erro As Integer
                    ChkResultado.Items.Clear()
                    For Mes As Integer = DdlMesInicial.SelectedValue To DdlMesFinal.SelectedValue
                        ChkResultado.Items.Add(New ListItem("*********************************************************************"))
                        ChkResultado.Items.Add(New ListItem("************  Mes: " & Mes & "  Ano: " & DdlAno.SelectedValue))
                        ChkResultado.Items.Add(New ListItem("*********************************************************************"))
                        ChkResultado.Items.Add(New ListItem(""))

                        '****************************************************************************************
                        '***************************  LIMPAR MOVIMENTO DO MES   *********************************
                        '****************************************************************************************
                        If LimpaMovimentoDoMes(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Movimento do Mes Limpo"))
                        Else
                            ChkResultado.Items.Add(New ListItem("Erro ao Limpar o Movimento do Mes"))
                        End If

                        '****************************************************************************************
                        '*****************************   APURAR SALDO INICIAL   *********************************
                        '****************************************************************************************
                        erro = ApuraSaldoInicial(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Saldos Iniciais Apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) na Apuracao de Saldos Iniciais"))
                        End If

                        '****************************************************************************************
                        '********************************   CAPTURAR RAZAO   ************************************
                        '****************************************************************************************
                        erro = CapturaRazao(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Dados do Razão Apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) na Apuracao dos dados do Razao"))
                        End If

                        '****************************************************************************************
                        '********************************   CAPTURAR ESTOQUES  **********************************
                        '****************************************************************************************
                        erro = CapturaEstoques(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Dados do Estoque Apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) na Apuracao dos dados do Estoque"))
                        End If

                        '****************************************************************************************
                        '*****************************   CAPTURAR NOTAS FISCAIS  ********************************
                        '****************************************************************************************
                        erro = CapturaNotasFiscais(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Dados das Notas Fiscais Apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) na Apuracao dos dados das Notas Fiscais"))
                        End If


                        '****************************************************************************************
                        '*****************************   CAPTURAR NOTAS FISCAIS  DE TRANSFERENCIAS **************
                        '****************************************************************************************

                        If Transferencias(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Dados das Transferencias Apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & "Erro(s) na Apuracao das Notas de Transferencias"))
                        End If

                        '****************************************************************************************
                        '*****************************   CAPTURAR Fretes (Compras, Depositos, Transferencias) ***
                        '****************************************************************************************

                        If Fretes(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Dados dos Fretes Apurados"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & "Erro(s) na Apuracao dos Fretes"))
                        End If

                        'CapturaNotasFiscaisContraPartida(Mes)

                        '****************************************************************************************
                        '*************************   AJUSTAR COMPLEMENTACAO DE PRECO   **************************
                        '****************************************************************************************
                        If AjusteComplementacaoDePreco(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Ajustando Complementacao de Preco"))
                        Else
                            ChkResultado.Items.Add(New ListItem("Erro Durante o ajuste de complementacao de Preco"))
                        End If

                        '****************************************************************************************
                        '************   ESTORNA VALOR DE AJUSTE DA AVALIACAO DO AFIXAR DO MES ANTERIOR  *********
                        '****************************************************************************************
                        If EstornaAfixarMesAnterior(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Valor estornado"))
                        Else
                            ChkResultado.Items.Add(New ListItem("Erro Durante o estorno do valor reajustado do A fixar"))
                        End If

                        '****************************************************************************************
                        '*****************************   AVALIAR ESTOQUE A FIXAR   ******************************
                        '****************************************************************************************
                        If AvaliarEstoqueAFixar(Mes) Then
                            ChkResultado.Items.Add(New ListItem("Estoque A fixar Avaliado"))
                        Else
                            ChkResultado.Items.Add(New ListItem("Erro Durante a avaliacao do estoque A fixar"))
                        End If


                        AjustaTotalizadores(Mes)

                        '****************************************************************************************
                        '***************************   AJUSTAR CUSTOS DE SAIDA   ********************************
                        '****************************************************************************************
                        erro = AjustaCustosDeSaidas(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Custos de Saida Ajustador"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes nos custos de saida"))
                        End If

                        '****************************************************************************************
                        '***************************   Ajustar Saidas Por Transferencias*************************
                        '****************************************************************************************
                        erro = AjustaTransferencias(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Saidas Por Transferencias Ajustador"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes das Saidas por Transferencias"))
                        End If

                        AjustaTotalizadores(Mes)

                        '****************************************************************************************
                        '***************************   AJUSTAR PRODUTO DERIVADO   *******************************
                        '****************************************************************************************
                        erro = AjustaProdutoDerivado(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Produto Derivado Ajustado"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes do Produto Derivado"))
                        End If

                        '****************************************************************************************
                        '***************************   AJUSTAR CONSUMO X PRODUCAO   *****************************
                        '****************************************************************************************
                        erro = AjustaConsumoXProducao(Mes)
                        If erro = 0 Then
                            ChkResultado.Items.Add(New ListItem("Consumo x Producao Ajustado"))
                        Else
                            ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes do Consumo x Producao"))
                        End If

                        AjustaTotalizadores(Mes)

                        '****************************************************************************************
                        '********************************  CALCULO CIRCULAR   ***********************************
                        '****************************************************************************************
                        ChkResultado.Items.Add(New ListItem("Iniciando Calculo Circular"))
                        'Circular = "S"

                        For ii = 0 To CInt(txtCiclos.Text)

                            'Circular = "N"
                            'erro = CalculoCircular(Mes)
                            'If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante o Calculo Circular"))

                            RateioDosCustosDeProducao_Soja(Mes)

                            erro = AjustaCustosDeSaidas(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes nos custos de saida"))

                            erro = AjustaTransferencias(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes nas Saidas Por Transferencias"))

                            erro = AjustaConsumoXProducao(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes do Consumo x Producao"))

                            AjustaTotalizadores(Mes)

                            erro = AjustaProdutoDerivado(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes do Produto Derivado"))

                            AjustaTotalizadores(Mes)

                            erro = AjustaCustosDeSaidas(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes nos custos de saida"))

                            erro = AjustaTransferencias(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes nas Saidas Por Transferencias"))

                            AjustaTotalizadores(Mes)

                            erro = AjustaProdutoDerivado(Mes)
                            If erro > 0 Then ChkResultado.Items.Add(New ListItem(erro & " - Erro(s) durante os Ajustes do Produto Derivado"))
                            AjustaTotalizadores(Mes)
                            ii += 1
                        Next

                        ContabilizaCustos(Mes)
                    Next
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                Finally
                    crpt.Close()
                    crpt.Dispose()
                End Try
            Else
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
            End If
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CapturaDeDados")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class