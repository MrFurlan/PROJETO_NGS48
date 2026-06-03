Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucSelecaoOperacoes
    Inherits BaseUserControl

    Dim ListOperacao As [Lib].Negocio.ListOperacao
    Dim ListSubOperacao As [Lib].Negocio.ListSubOperacao

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            HIDSelecaoOperacoes.Value = Guid.NewGuid.ToString
            Limpar()
        End If
    End Sub

    Public Overrides Sub Limpar()
        SessaoRecuperaOperacao()
        If ListOperacao.Selecionar() Then
            gridOperacoes.DataSource = ListOperacao.ToArray()
            gridOperacoes.DataBind()
        Else
            gridOperacoes.DataSource = Nothing
            gridOperacoes.DataBind()
        End If
        SessaoSalvarOperacao()
    End Sub

#Region "SESSAO"
    Private Sub SessaoSalvarOperacao()
        Session("ListOperacao" + HIDSelecaoOperacoes.Value) = ListOperacao
    End Sub

    Private Sub SessaoRecuperaOperacao()
        If Session("ListOperacao" + HIDSelecaoOperacoes.Value) Is Nothing Then
            ListOperacao = New [Lib].Negocio.ListOperacao()
        Else
            ListOperacao = Session("ListOperacao" + HIDSelecaoOperacoes.Value)
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

    Public Property WhereOperacoes As String
        Get
            Return HWhereOperacoes.Value
        End Get
        Set(ByVal value As String)
            HWhereOperacoes.Value = value
        End Set
    End Property

    '*************** Seta como visível ou invisível a coluna de produtos bem como a carga dos mesmos. *************
    Public Property OperacoesVisiveis As Boolean
        Get
            OperacoesVisiveis = pnlSelecaoOperacoes.Visible
        End Get
        Set(ByVal value As Boolean)
            pnlSelecaoOperacoes.Visible = value
        End Set
    End Property

    Protected Sub gridOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridOperacoes.SelectedIndexChanged
        If OperacoesVisiveis Then
            HIDOperacoes.Value = gridOperacoes.SelectedIndex
            SessaoRecuperaOperacao()
            If ListOperacao(gridOperacoes.SelectedIndex).Selecionado Then
                gridSubOperacoes.DataSource = ListOperacao(gridOperacoes.SelectedIndex).SubOperacoes.ToArray()
                gridSubOperacoes.DataBind()
            Else
                gridSubOperacoes.DataSource = Nothing
                gridSubOperacoes.DataBind()
            End If
            SessaoSalvarOperacao()
            AbrirAccordion(Me.ID, True, True, 100)
        End If
    End Sub

    Protected Sub chkOperacoes_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkOperacoes As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkOperacoes.NamingContainer, GridViewRow)
        HIDOperacoes.Value = row.RowIndex
        SessaoRecuperaOperacao()
        If chkOperacoes.Checked Then
            ListOperacao(row.RowIndex).Selecionado = True
            ListOperacao(row.RowIndex).Selecionado = chkOperacoes.Checked
            ListOperacao(row.RowIndex).SubOperacoes = New ListSubOperacao(ListOperacao(row.RowIndex).Codigo)
            gridSubOperacoes.DataSource = ListOperacao(row.RowIndex).SubOperacoes.ToArray()
            gridSubOperacoes.DataBind()
        Else
            ListOperacao(row.RowIndex).Selecionado = False
            ListOperacao(row.RowIndex).SubOperacoes = Nothing
            gridSubOperacoes.DataSource = Nothing
            gridSubOperacoes.DataBind()
        End If
        SessaoSalvarOperacao()
        AbrirAccordion(Me.ID, True, True, 100)
    End Sub

    Protected Sub chkSubOperacoes_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkSubOperacoes As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkSubOperacoes.NamingContainer, GridViewRow)
        SessaoRecuperaOperacao()
        ListOperacao(HIDOperacoes.Value).SubOperacoes(row.RowIndex).Selecionado = chkSubOperacoes.Checked
        SessaoSalvarOperacao()
        AbrirAccordion(Me.ID, True, True, 100)
    End Sub

    '*************** Parâmetro pSeleciona (True ou False) Seleciona ou retira a seleção de todos os grupos e o pExpandAccordion (True ou false) para expandir ou não o corpo do user control . *************
    'Public Sub SelecionaTodosOsGrupos(ByVal pSeleciona As Boolean, Optional ByVal pExpandAccordion As Boolean = False)
    '    SessaoRecuperaGrupo()
    '    If Not ListGrupo Is Nothing Then
    '        For Each grp As [Lib].Negocio.GrupoProduto In ListGrupo
    '            grp.Selecionado = pSeleciona
    '        Next
    '        SessaoSalvarGrupo()
    '        gridGrupo.DataSource = ListGrupo
    '        gridGrupo.DataBind()
    '        If pExpandAccordion Then
    '            AbrirAccordion(Me.ID, True, True, 100)
    '        End If
    '    End If
    'End Sub



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
    Public Function GetStringOperacoesSelecionado() As String
        Dim strOperacao As String = ""
        Dim strSubOperacao As String = ""
        Dim parametrosOperacao As String = ""
        Dim straux1 As String = ""
        Dim TemFalso As Boolean
        Dim primeraVez As Boolean = True

        SessaoRecuperaOperacao()
        For Each op As [Lib].Negocio.Operacao In ListOperacao
            If op.Selecionado = True Then
                TemFalso = False
                straux1 = ""
                For Each opSub As [Lib].Negocio.SubOperacao In op.SubOperacoes
                    If opSub.Selecionado = True Then
                        strSubOperacao &= IIf(strSubOperacao.Length > 0, ",", "") & opSub.Codigo
                    Else
                        TemFalso = True
                    End If
                Next

                strOperacao &= op.Codigo & "-" & strSubOperacao & "|"

                strSubOperacao = ""
            End If
        Next

        Return strOperacao
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
    Public Function GetSqlEParametrosRelatorioOperacoes(ByVal ColunaOperacoesNaConsulta As String, ByVal ColunaSubOperacoesNaConsulta As String, Optional ByVal stringOPSubOperacoes As String = "", Optional ByVal DescricaoOPSubOperacoes As Boolean = False) As ArrayList
        Dim Sql As String = ""
        Dim Parametros As String = ""
        Dim Descricao As String = ""
        '****************************************************************************
        '************************ OPERAÇÕES E SUBOPERAÇÕES  ********************************
        '****************************************************************************
        If stringOPSubOperacoes.Length = 0 Then
            stringOPSubOperacoes = GetStringOperacoesSelecionado()
        End If

        Dim OPe As String() = stringOPSubOperacoes.Split("|")

        Dim PrimeiraVez As Boolean = True
        Dim PrimeiraDescricao As Boolean = True
        Dim Separador As String = ""

        If stringOPSubOperacoes.Length > 1 Then
            Sql = "("
            For x As Integer = 0 To OPe.Length - 1

                If OPe(x).Length > 0 Then

                    Dim OPSub As String() = OPe(x).Split("-")

                    If PrimeiraVez Then
                        If OPSub(1).Length > 0 Then
                            Sql &= "(Operacao_Id = " & OPSub(0) & " AND SubOperacoes_Id in(" & OPSub(1) & "))"
                        Else
                            Sql &= "(Operacao_Id = " & OPSub(0) & ")"
                        End If
                        PrimeiraVez = False
                    Else
                        If OPSub(1).Length > 0 Then
                            Sql &= " OR (Operacao_Id = " & OPSub(0) & " AND SubOperacoes_Id in(" & OPSub(1) & "))"
                        Else
                            Sql &= " OR (Operacao_Id = " & OPSub(0) & ")"
                        End If
                    End If

                    If OPSub(1).Length > 0 Then
                        Dim OPSubs As String() = OPSub(1).Split(",")

                        For y As Integer = 0 To OPSubs.Length - 1

                            If PrimeiraDescricao Then
                                Descricao += CInt(OPSub(0)).ToString("00") & "-" & CInt(OPSubs(y)).ToString("00")
                                PrimeiraDescricao = False
                            Else
                                Descricao += ", " & CInt(OPSub(0)).ToString("00") & "-" & CInt(OPSubs(y)).ToString("00")
                            End If

                        Next
                    Else
                        If PrimeiraDescricao Then
                            Descricao += CInt(OPSub(0)).ToString("00") & "-TODAS"
                            PrimeiraDescricao = False
                        Else
                            Descricao += ", " & CInt(OPSub(0)).ToString("00") & "-TODAS"
                        End If

                    End If
                End If
            Next

            Parametros &= "Operação:" & Descricao & vbCrLf

            Sql &= ")"
        End If

        Dim retorno As New ArrayList()
        retorno.Add(Sql)
        retorno.Add(Parametros)
        Return retorno
    End Function

    '******************* Retorna um Boolean dizendo se alguma selecao foi realizada *****************************
    Public Function TemSelecionado() As Boolean
        SessaoRecuperaOperacao()
        For Each op In ListOperacao
            If op.Selecionado Then
                For Each subOP In op.SubOperacoes
                    If subOP.Selecionado Then
                        Return True
                    End If
                Next
            End If
        Next
        Return False
    End Function

    '******************* Carrega o grid com a string recuperada, informada ou montada  **************************
    'Public Sub PreenheGrid(ByVal pSelecao As String)

    '    Dim sql As String
    '    sql = "Select left(grupo," & TG & ") as Grupo, Produto_Id" & vbCrLf &
    '          "  from Produtos " & vbCrLf &
    '          " Where " & GetSqlEParametrosRelatorio("Grupo", "Produto_Id", pSelecao)(0) & vbCrLf &
    '          " Order By left(Grupo," & TG & "), Produto_id"

    '    Dim ds As DataSet
    '    ds = Banco.ConsultaDataSet(sql, "Consulta")

    '    Dim Grupo As String = ""
    '    Dim produto As String = ""
    '    Dim GrupoId As Integer = 0
    '    Dim ProdutoId As Integer = 0
    '    Dim GrupoDeProduto As GrupoProduto = Nothing

    '    SessaoRecuperaOperacao()
    '    For Each row As DataRow In ds.Tables(0).Rows
    '        If Grupo <> row("Grupo") Then
    '            Grupo = row("Grupo")
    '            GrupoId = CInt(row("Grupo"))
    '            GrupoDeProduto = ListGrupo.Find(Function(s) s.Codigo = GrupoId)
    '            GrupoDeProduto.Selecionado = True
    '            GrupoDeProduto.Produtos.ForEach(Function(s)
    '                                                s.Selecionado = False
    '                                                Return True
    '                                            End Function)
    '        End If
    '        ProdutoId = CInt(row("Produto_id"))
    '        GrupoDeProduto.Produtos.Find(Function(s) s.Codigo = ProdutoId).Selecionado = True
    '    Next

    '    gridOperacoes.DataSource = ListOperacao.ToArray()
    '    gridOperacoes.DataBind()
    'End Sub

    Public Sub AbrirAccordion(ByVal AccordionId As String, ByVal AbrirAccordion As Boolean, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & AccordionId & "').accordion('activate',  " & IIf(AbrirAccordion, 0, 1) & ");"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub
#End Region

    Function FindControlRecursive(p1 As String) As Object
        Throw New NotImplementedException
    End Function

End Class