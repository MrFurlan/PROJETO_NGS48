Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucSelecaoProduto
    Inherits BaseUserControl

    Dim ListGrupo As [Lib].Negocio.ListGrupoProduto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            HIDSelecaoProduto.Value = Guid.NewGuid.ToString
            'HWhereProduto.Value = IIf(HWhereProduto.Value.Length > 0, HWhereProduto.Value + " and ", "") + "Situacao = 1"
        End If
    End Sub

#Region "SESSAO"
    Private Sub SessaoSalvarGrupo()
        Session("ListGrupo" + HIDSelecaoProduto.Value) = ListGrupo
    End Sub

    Private Sub SessaoRecuperaGrupo()
        If Session("ListGrupo" + HIDSelecaoProduto.Value) Is Nothing Then
            ListGrupo = New [Lib].Negocio.ListGrupoProduto()
        Else
            ListGrupo = Session("ListGrupo" + HIDSelecaoProduto.Value)
        End If
    End Sub
#End Region

#Region "Metodos Publicos"
    '************  Setar o nome visivel do User Control por padrao esta "Seleção de Produtos" *******************
    Public Property NomeUC
        Get
            Return lblNome.Text
        End Get
        Set(ByVal value)
            lblNome.Text = value
        End Set
    End Property

    Public Property WhereProduto As String
        Get
            Return HWhereProduto.Value
        End Get
        Set(ByVal value As String)
            HWhereProduto.Value = value
        End Set
    End Property

    '*************** Seta como visível ou invisível a coluna de produtos bem como a carga dos mesmos. *************
    Public Property ProdutosVisiveis As Boolean
        Get
            ProdutosVisiveis = pnlSelecaoProdutos.Visible
        End Get
        Set(ByVal value As Boolean)
            pnlSelecaoProdutos.Visible = value
        End Set
    End Property

    '*************** Parâmetro pSeleciona (True ou False) Seleciona ou retira a seleção de todos os grupos e o pExpandAccordion (True ou false) para expandir ou não o corpo do user control . *************
    Public Sub SelecionaTodosOsGrupos(ByVal pSeleciona As Boolean, Optional ByVal pExpandAccordion As Boolean = False)
        SessaoRecuperaGrupo()
        If Not ListGrupo Is Nothing Then
            For Each grp As [Lib].Negocio.GrupoProduto In ListGrupo
                grp.Selecionado = pSeleciona
            Next
            SessaoSalvarGrupo()
            gridGrupo.DataSource = ListGrupo
            gridGrupo.DataBind()
            If pExpandAccordion Then
                AbrirAccordion(Me.ID, True, True, 100)
            End If
        End If
    End Sub



    '**********************************************************************************************************************************************************************************
    '**************** Devolve os grupo e produtos selecionados, Formatacao da string:  Tamanho do Grupo | Grupos ; Produtos  os Grupos e produtos estao separados por "," *************
    '* Faça um split com ; para dividir Grupo e produto, a string no index 0 ficara o grupo e no index 1 os Produtos ******************************************************************
    '* Quando o index 0 que se refere ao grupo conter valores significa q todos os produtos daquele grupo foram selecionados
    '* para montar o sql do grupo faca um split no index 0 usando o caracter | na posicao 0 lhe mostrara o nivel do grupo e na posicao 1 estarao os grupos selecionados
    '* EX: 2|21,30,40;                           "2" nivel do grupo  "|" Separador  "21,30,40" Grupos ";" Separado " " Produtos nesse caso nao foi selecionado um produto + sim todos os produtos q estao no grupo 21,30 e  40
    '*     3|301,302,303,304,305,306;            "3" nivel do grupo  "|" Separador  "301,302,303,304,305,306" Grupos ";" Separado " " Produtos nesse caso nao foi selecionado um produto + sim todos os produtos q estao no grupo 301,302,303,304,305 e 306
    '*     5|10102,10104;101010001,101010003     "5" nivel do grupo  "|" Separador  "10102,10104" Grupos ";" Separado "101010001,101010003" Produtos nesse caso nao foi selecionado dois grupos com todos os seus produtos + os produtos 101010001 e 101010003
    '*     ;101010001,101010003                  " " nao foi selecionado nenhum grupo  ";" Separado "101010001,101010003" Produtos nesse caso nao foi selecionado apenas os produtos 101010001 e 101010003 
    '* olhe o exemplo para montar o sql no proximo metodo neste mesmo User control GetSqlEParametrosRelatorio 
    Public Function GetStringGrupoProdutoSelecionado() As String
        Dim strGrupo As String = ""
        Dim strProduto As String = ""
        Dim straux1 As String = ""
        Dim TemFalso As Boolean
        SessaoRecuperaGrupo()

        For Each gp As [Lib].Negocio.GrupoProduto In ListGrupo
            If gp.Selecionado Then
                TemFalso = False
                straux1 = ""
                For Each prd As [Lib].Negocio.Produto In gp.Produtos
                    If prd.Selecionado = True Then
                        straux1 &= IIf(straux1.Length > 0, ",", "") & prd.Codigo
                    Else
                        TemFalso = True
                    End If
                Next

                If TemFalso Then
                    strProduto &= IIf(strProduto.Length > 0, ",", "") & straux1
                Else
                    strGrupo &= IIf(strGrupo.Length > 0, ",", "") & gp.Codigo
                End If
            End If
        Next

        Dim GrupoProduto As String
        If strGrupo.Length > 0 Then
            GrupoProduto = ListGrupo(0).Codigo.Length & "|" & strGrupo & ";" & strProduto
        Else
            GrupoProduto = ";" & strProduto
        End If

        'Return IIf(strGrupo.Length > 0, ListGrupo(0).Codigo.Length & "|" & strGrupo, "") & ";" & strProduto
        Return GrupoProduto

    End Function

    '************  Simplifica o retorno ja montando sql e devolvendo os paramentros para usar na descricao do relatorio *******************
    '************  Voce tem q passar somente os nomes das colunas do grupo e do produto usado no sql q vc esta montando de preferencia com o alias 
    '************  EX: GetSqlEParametrosRelatorio("prd.Grupo", "nfi.Produto_Id","") As String()  o terceiro parametro é opcional se vc ja tiver a string e quizer somente montar o sql
    '************  - ai vc escolheu 2 grupos e 2 produtos a funcao anterior vai retornar 5|10102,10104;101010001,101010003
    '************  e vai montar o sql  "(left(prd.grupo,5) in('10102','10104') or nfi.produto_id in ('101010001','101010003')) "
    '************  e vai montar o descritivo para o paramentro assim :
    '************  Grupo(s) de Produto: 10102,10104
    '************  Produto(s): 101010001,101010003
    '************  Lembrando q o retorno é um arraylist entao o index 0 é o sql e o index 1 é o Descritivo do parametro
    Public Function GetSqlEParametrosRelatorio(ByVal ColunaGrupoNaConsulta As String, ByVal ColunaProdutoNaConsulta As String, Optional ByVal stringGrupoProduto As String = "", Optional ByVal DescricaoGrupoeProduto As Boolean = False) As ArrayList
        Dim Sql As String = ""
        Dim Parametros As String = ""
        '****************************************************************************
        '************************ GRUPOS E PRODUTOS  ********************************
        '****************************************************************************
        If stringGrupoProduto.Length = 0 Then
            stringGrupoProduto = GetStringGrupoProdutoSelecionado()
        End If
        Dim GruposEProdutos As String() = stringGrupoProduto.Split(";")
        Dim sqlAuxPv As String = ""

        '*************************  GRUPOS ******************************************
        If GruposEProdutos(0).Length > 0 Then
            Dim Grupo As String() = GruposEProdutos(0).Split("|")
            Dim ParametroGrupo As String = ""

            If Not DescricaoGrupoeProduto Then
                ParametroGrupo = Grupo(1)
            End If

            Dim Grupos As String() = Grupo(1).Split(",")
            Dim grp As String = ""

            For x As Integer = 0 To Grupos.Length - 1
                grp += IIf(grp.Length > 0, ",", "") & "'" & Grupos(x) & "'"
                If DescricaoGrupoeProduto Then
                    Dim ConsGrupo As New GrupoProduto(Grupos(x))
                    ParametroGrupo &= IIf(ParametroGrupo = "", "", ", ") & ConsGrupo.Codigo & "-" & ConsGrupo.Descricao
                End If
            Next
            sqlAuxPv &= ColunaProdutoNaConsulta & " in (Select Produto_id from Produtos where LEFT(grupo," & Grupo(0) & ") in (" & grp & "))"

            Parametros &= "Grupo(s) de Produto:" & ParametroGrupo & vbCrLf
        End If

        '*************************  PRODUTO ******************************************
        If GruposEProdutos(1).Length > 0 Then
            Dim ParametrosProduto As String = ""
            Dim sqlprd As String = ""

            If Not DescricaoGrupoeProduto Then
                'ParametrosProduto &= "Produto(s): " & GruposEProdutos(1)
                ParametrosProduto &= GruposEProdutos(1)
            End If

            Dim xvPrd As String() = GruposEProdutos(1).Split(",")
            Dim prd As String = ""
            For x As Integer = 0 To xvPrd.Length - 1
                prd += IIf(prd.Length > 0, ",", "") & "'" & xvPrd(x) & "'"
                If DescricaoGrupoeProduto Then
                    Dim ConsPrd As New Produto(xvPrd(x))
                    ParametrosProduto &= IIf(ParametrosProduto = "", "", ", ") & ConsPrd.Codigo & "-" & ConsPrd.Descricao
                End If
            Next

            sqlprd = ColunaProdutoNaConsulta & " in (Select Produto_id from Produtos where Produto_Id in (" & prd & "))"

            Sql &= IIf(sqlAuxPv.Length > 0, "(" & sqlAuxPv & " OR " & sqlprd & ")", sqlprd)

            Parametros &= "Produto(s): " & ParametrosProduto & vbCrLf
        Else
            Sql &= sqlAuxPv
        End If

        Dim retorno As New ArrayList()
        retorno.Add(Sql)
        retorno.Add(Parametros)
        Return retorno
    End Function

    '******************* Retorna um Boolean dizendo se alguma selecao foi realizada *****************************
    Public Function TemSelecionado() As Boolean
        SessaoRecuperaGrupo()
        For Each grp In ListGrupo
            If grp.Selecionado Then
                For Each prd In grp.Produtos
                    If prd.Selecionado Then
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function

    '******************* Carrega o grid com a string recuperada, informada ou montada  **************************
    Public Sub PreenheGrid(ByVal pSelecao As String)
        If pSelecao.Length = 0 Then Exit Sub

        Dim x As String() = pSelecao.Split(";")
        x = x(0).Split("|")
        Dim TG As Integer = IIf(x(0).Length = 0, 5, x(0))

        Select Case TG
            Case 1 : rdnivel1.Checked = True
            Case 2 : rdnivel10.Checked = True
            Case 3 : rdnivel100.Checked = True
            Case 5 : rdnivel10000.Checked = True
        End Select

        CarregarNivel(TG)

        Dim sql As String
        sql = "Select left(grupo," & TG & ") as Grupo, Produto_Id" & vbCrLf & _
              "  from Produtos " & vbCrLf & _
              " Where " & GetSqlEParametrosRelatorio("Grupo", "Produto_Id", pSelecao)(0) & vbCrLf & _
              " Order By left(Grupo," & TG & "), Produto_id"

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Consulta")

        Dim Grupo As String = ""
        Dim produto As String = ""
        Dim GrupoId As Integer = 0
        Dim ProdutoId As Integer = 0
        Dim GrupoDeProduto As GrupoProduto = Nothing

        SessaoRecuperaGrupo()
        For Each row As DataRow In ds.Tables(0).Rows
            If Grupo <> row("Grupo") Then
                Grupo = row("Grupo")
                GrupoId = CInt(row("Grupo"))
                GrupoDeProduto = ListGrupo.Find(Function(s) s.Codigo = GrupoId)
                GrupoDeProduto.Selecionado = True
                GrupoDeProduto.Produtos.ForEach(Function(s)
                                                    s.Selecionado = False
                                                    Return True
                                                End Function)
            End If
            ProdutoId = CInt(row("Produto_id"))
            GrupoDeProduto.Produtos.Find(Function(s) s.Codigo = ProdutoId).Selecionado = True
        Next

        gridGrupo.DataSource = ListGrupo.ToArray()
        gridGrupo.DataBind()
    End Sub

    '******************* Carrega o Nivel do produto  **************************
    Public Sub CarregarNivel(ByVal nivel As Integer)
        Dim sql As String = ""
        sql = " SELECT Grupo_id, Descricao " & _
              "   FROM GruposDeEstoques " & _
              "  Where Len(Grupo_Id) = " & nivel & _
              "    and exists (select 1 from produtos where left(grupo," & nivel & ") = grupo_id " & IIf(HWhereProduto.Value.Length > 0, " and " & HWhereProduto.Value, "") & ")" & vbCrLf & _
              "  ORDER BY Grupo_id"
        ListGrupo = New [Lib].Negocio.ListGrupoProduto
        For Each row As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            Dim gp As New [Lib].Negocio.GrupoProduto
            gp.Codigo = row("Grupo_id")
            gp.Descricao = row("Descricao")
            gp.Selecionado = False
            ListGrupo.Add(gp)
        Next

        Select Case nivel
            Case 1 : rdnivel1.Checked = True
            Case 2 : rdnivel10.Checked = True
            Case 3 : rdnivel100.Checked = True
            Case 5 : rdnivel10000.Checked = True
        End Select

        gridGrupo.DataSource = ListGrupo.ToArray()
        gridGrupo.DataBind()
        gridProduto.DataSource = Nothing
        gridProduto.DataBind()

        SessaoSalvarGrupo()

        pnlSelecaoProdutos.Visible = Me.ProdutosVisiveis

        btnMarcarTodos.Visible = True
        btnDesmarcarTodos.Visible = True
        btnInverte.Visible = True
        AbrirAccordion(Me.ID, True, True, 100)

    End Sub

    Public Sub AbrirAccordion(ByVal AccordionId As String, ByVal AbrirAccordion As Boolean, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & AccordionId & "').accordion('activate',  " & IIf(AbrirAccordion, 0, 1) & ");"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    '****************** Limpar ***********************************
    Public Overrides Sub Limpar()
        rdnivel1.Checked = False
        rdnivel10.Checked = False
        rdnivel100.Checked = False
        rdnivel10000.Checked = False

        gridGrupo.DataSource = Nothing
        gridGrupo.DataBind()
        gridProduto.DataSource = Nothing
        gridProduto.DataBind()
    End Sub

#End Region

    Protected Sub rdnivel1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CarregarNivel(1)
    End Sub

    Protected Sub rdnivel10_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdnivel10.CheckedChanged
        CarregarNivel(2)
    End Sub

    Protected Sub rdnivel100_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdnivel100.CheckedChanged
        CarregarNivel(3)
    End Sub

    Protected Sub rdnivel10000_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CarregarNivel(5)
    End Sub

    Protected Sub gridGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridGrupo.SelectedIndexChanged
        If ProdutosVisiveis Then
            SessaoRecuperaGrupo()
            HGrupo.Value = gridGrupo.SelectedIndex
            If ListGrupo(gridGrupo.SelectedIndex).Selecionado Then
                gridProduto.DataSource = ListGrupo(gridGrupo.SelectedIndex).Produtos.ToArray()
                gridProduto.DataBind()
            Else
                gridProduto.DataSource = Nothing
                gridProduto.DataBind()
            End If
            AbrirAccordion(Me.ID, True, True, 100)
        End If
    End Sub

    Protected Sub chkGrupo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkGrupo As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkGrupo.NamingContainer, GridViewRow)
        HGrupo.Value = row.RowIndex
        SessaoRecuperaGrupo()
        If chkGrupo.Checked Then
            ListGrupo(row.RowIndex).Selecionado = True
            ListGrupo(row.RowIndex).Produtos = New ListProduto(ListGrupo(row.RowIndex).Codigo, "", "", "", "", HWhereProduto.Value.ToString, True)
            gridProduto.DataSource = ListGrupo(row.RowIndex).Produtos.ToArray()
            gridProduto.DataBind()
        Else
            ListGrupo(row.RowIndex).Selecionado = False
            ListGrupo(row.RowIndex).Produtos = Nothing
            gridProduto.DataSource = Nothing
            gridProduto.DataBind()
        End If
        SessaoSalvarGrupo()
        AbrirAccordion(Me.ID, True, True, 100)
    End Sub

    Protected Sub chkProduto_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkProduto As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkProduto.NamingContainer, GridViewRow)
        SessaoRecuperaGrupo()
        ListGrupo(HGrupo.Value).Produtos(row.RowIndex).Selecionado = chkProduto.Checked
        SessaoSalvarGrupo()
        AbrirAccordion(Me.ID, True, True, 100)
    End Sub

    Protected Sub btnMarcarTodos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaGrupo()
        If ListGrupo Is Nothing Then Exit Sub
        For Each prd As [Lib].Negocio.Produto In ListGrupo(HGrupo.Value).Produtos
            prd.Selecionado = True
        Next
        SessaoSalvarGrupo()
        gridProduto.DataSource = ListGrupo(HGrupo.Value).Produtos
        gridProduto.DataBind()
        AbrirAccordion(Me.ID, True, True, 100)
    End Sub

    Protected Sub btnDesmarcarTodos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDesmarcarTodos.Click
        SessaoRecuperaGrupo()

        If ListGrupo Is Nothing Then Exit Sub

        If Not String.IsNullOrWhiteSpace(HGrupo.Value) Then
            For Each prd As [Lib].Negocio.Produto In ListGrupo(HGrupo.Value).Produtos
                prd.Selecionado = False
            Next
            SessaoSalvarGrupo()
            gridProduto.DataSource = ListGrupo(HGrupo.Value).Produtos
            gridProduto.DataBind()
            AbrirAccordion(Me.ID, True, True, 100)
        End If
    End Sub

    Protected Sub btnInverte_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaGrupo()
        If ListGrupo Is Nothing Then Exit Sub
        For Each prd As [Lib].Negocio.Produto In ListGrupo(HGrupo.Value).Produtos
            prd.Selecionado = Not prd.Selecionado
        Next
        SessaoSalvarGrupo()
        gridProduto.DataSource = ListGrupo(HGrupo.Value).Produtos
        gridProduto.DataBind()
        AbrirAccordion(Me.ID, True, True, 100)
    End Sub

    Function FindControlRecursive(p1 As String) As Object
        Throw New NotImplementedException
    End Function

End Class