Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class DuplicataAvulsa
    Inherits BasePage

    Private dateEmissao As New DateTime
    Private intPagina As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("DuplicataAvulsa", "ACESSAR") Then
                BuscarUnidadeDeNegocio()
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Sub BuscarUnidadeDeNegocio()
        Try
            ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If txtDataVenc.Text.Length = 0 Or IsDate(txtDataVenc.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida")
            txtDataVenc.Focus()
            Return False
        ElseIf ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Informe a Empresa.")
            ddlEmpresa.Focus()
            Return False
        ElseIf txtCodigoCliente.Value = "" Then
            MsgBox(Me.Page, "Informe o Cliente.")
            txtClientes.Focus()
            Return False
        ElseIf txtNumNota.Text = "" Then
            MsgBox(Me.Page, "Informe Numero Da Nota.")
            txtNumNota.Focus()
            Return False
        ElseIf txtDuplicata.Text = "" Then
            MsgBox(Me.Page, "Informe Numero Da Duplicata.")
            txtDuplicata.Focus()
            Return False
        ElseIf txtParcelas.Text = "" Then
            MsgBox(Me.Page, "Informe Numero de Parcelas.")
            txtParcelas.Focus()
            Return False
        ElseIf txtTotalNota.Text = "" Then
            MsgBox(Me.Page, "Informe Total Da Nota.")
            txtTotalNota.Focus()
            Return False
        End If

        Dim totalParc As Decimal
        For i As Integer = 0 To GridParcelas.Rows.Count - 1
            totalParc += GridParcelas.Rows(i).Cells(2).Text
        Next
        If totalParc <> CDec(Me.txtTotalNota.Text) Then
            MsgBox(Me.Page, "Total das parcelas difere do valor total, Verifique.")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        txtClientes.Text = ""
        txtPraca.Text = ""
        txtNumNota.Text = ""
        txtDuplicata.Text = ""
        txtParcelas.Text = ""
        txtDataVenc.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataEmissao.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtTotalNota.Text = ""
        txtValor.Text = ""
        txtCodigoCliente.Value = ""
        txtCodigoPraca.Value = ""
        GridParcelas.DataSource = ""
        GridParcelas.DataBind()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Session("objClienteDupAvPraca" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteDupAvPraca" & HID.Value), [Lib].Negocio.Cliente))
                txtPraca.Text = itemCliente.Text
                txtCodigoPraca.Value = itemCliente.Value
                Session.Remove("objClienteDupAvPraca" & HID.Value)
            ElseIf Session("objClienteDupAv" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteDupAv" & HID.Value), [Lib].Negocio.Cliente))
                txtClientes.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteDupAv" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlBuscaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDupAv" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaPraca_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDupAvPraca" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarRegistros()
        Try
            If Funcoes.VerificaPermissao("DuplicataAvulsa", "ACESSAR") Then
                Dim sqlProd As String = ""
                Dim DescricaoProduto As String = ""
                Dim dsProduto As New DataSet
                Dim ds As New DataSet
                Dim SQL As String
                Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                Dim strPraca() As String = txtCodigoPraca.Value.Split("-")

                SQL = " Select distinct sb.Parcela," & vbCrLf & _
                    " Empresas.nome as Empresa, " & vbCrLf & _
                    " Empresas.Cidade as Cidade, " & vbCrLf & _
                    " Empresas.Endereco + ' ' + isnull(convert(varchar(10),Empresas.Numero),'')  as Endereco,  " & vbCrLf & _
                    " Empresas.Bairro as Bairro, " & vbCrLf & _
                    " Empresas.Telefone as Fone, " & vbCrLf & _
                    " Empresas.Cep as Cep, " & vbCrLf & _
                    " Empresas.Estado as UF, " & vbCrLf & _
                    " Empresas.Cliente_Id as Cnpj, " & vbCrLf & _
                    " Empresas.Inscricao as Insc, " & vbCrLf & _
                    " 0.00 as Valnot, " & vbCrLf & _
                    " ''  as NumNot," & vbCrLf & _
                    " '' as NumOrd," & vbCrLf & _
                    " '' as SeqOrd, " & vbCrLf & _
                    " '' as Vencto, " & vbCrLf & _
                    " Clientes.Nome as Cliente," & vbCrLf & _
                    " Clientes.Endereco + ' ' + isnull(convert(varchar(10),Clientes.Numero),'') + '         ' + Clientes.Telefone  as EndCli, " & vbCrLf & _
                    " Clientes.Cep as CepCli, " & vbCrLf & _
                    " Clientes.Cidade as CidCli, " & vbCrLf & _
                    " Clientes.Estado as UFCli, " & vbCrLf & _
                    " Clientes.Cliente_Id as CnpjCli, " & vbCrLf

                If strPraca(0).Length > 0 Then
                    SQL &= " ClientePraca.Endereco + ' ' + isnull(convert(varchar(10),ClientePraca.Numero),'')  + '  ' + ClientePraca.Cidade  + '-' + ClientePraca.Estado + ' - ' + ClientePraca.Cep as Praca,  " & vbCrLf
                Else
                    SQL &= " '' as Praca, " & vbCrLf
                End If

                SQL &= " '' as Ext001, " & vbCrLf & _
                    " 'REG.:' as Ext002, " & vbCrLf & _
                    " Clientes.Reduzido as Ext003, " & vbCrLf & _
                    " Clientes.Inscricao as Inscli, " & vbCrLf & _
                    " 1 as Via, " & vbCrLf & _
                    " '' as Tipo, " & vbCrLf & _
                    " '' as ValTot1, " & vbCrLf & _
                    " '' AS Emissa , " & vbCrLf & _
                    " Clientes.Bairro as BaiCli " & vbCrLf & _
                    " FROM  Clientes AS Empresas " & vbCrLf & _
                    " CROSS JOIN Clientes AS Clientes " & vbCrLf & _
                    " CROSS JOIN Clientes AS ClientePraca " & vbCrLf & _
                    " Inner join (select 1 Parcela " & vbCrLf

                For x As Integer = 2 To CInt(txtParcelas.Text)
                    SQL &= " union select " & x
                Next
                SQL &= ") sb " & vbCrLf & _
                       "  on 1 = 1" & vbCrLf

                SQL &= " WHERE "
                ''Filtros
                If strEmpresa(0).Length > 0 Then
                    SQL &= " Empresas.Cliente_Id = '" & strEmpresa(0) & "'" & vbCrLf & _
                           " AND Empresas.Endereco_Id= " & strEmpresa(1) & " " & vbCrLf
                End If
                If strCliente(0).Length > 0 Then
                    SQL &= " AND Clientes.Cliente_Id = '" & strCliente(0) & "'" & vbCrLf & _
                           " AND Clientes.Endereco_Id= " & strCliente(1) & " " & vbCrLf
                End If
                If strPraca(0).Length > 0 Then
                    SQL &= " AND ClientePraca.Cliente_Id = '" & strPraca(0) & "'" & vbCrLf & _
                           " AND ClientePraca.Endereco_Id= " & strPraca(1) & " " & vbCrLf
                End If

                ds = Banco.ConsultaDataSet(SQL, "Cabecalho")

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Erro")
                    Exit Sub
                End If

                For i As Integer = 0 To GridParcelas.Rows.Count - 1
                    Dim row As DataRow = ds.Tables(0).Rows(i)
                    row("Valnot") = CDec(txtTotalNota.Text)
                    row("Numnot") = txtNumNota.Text
                    row("NumOrd") = i + 1
                    row("SeqOrd") = i + 1
                    row("Vencto") = GridParcelas.Rows(i).Cells(1).Text 'Format(CDate(Me.txtDataVenc.Text).AddMonths(i), "dd/MM/yyyy") 'GridParcelas.Rows(i)
                    'row("ValTot1") = CDec(GridParcelas.Rows(i).Cells(2).Text) 
                    row("ValTot1") = FormatNumber(GridParcelas.Rows(i).Cells(2).Text, 2)
                    row("Emissa") = IIf(String.IsNullOrEmpty(txtDataEmissao.Text), DateTime.Now.ToString("dd/MM/yyyy"), txtDataEmissao.Text)
                    row("Ext001") = Funcoes.Extenso(row("ValTot1"), "Real", "Reais").ToUpper
                    row("Ext002") = row("Ext002") & txtDuplicata.Text
                Next

                Montaduplicatas(ds)

                'AlimentaCrptRelatorios(ds, "~/Reports/Cr_Duplicata")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para fazer uma emissão.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Montaduplicatas(ByVal ds As DataSet)
        Try
            Dim drAux As DataRow
            Dim dsDuplicatasCopy As New DataSet

            'For Each drAux In dsDuplicatas.Tables(0).Rows 'Cria valor por extenso para duplicatas
            '    drAux("Ext001") = Utilitarios.Funcoes.Extenso(drAux("ValTot1"), "Reais", "Real").ToUpper
            'Next

            dsDuplicatasCopy = ds.Copy() 'copia ds duplicatas para criar segunda via
            For Each drAux In ds.Tables(0).Rows
                Dim dr As DataRow
                dr = dsDuplicatasCopy.Tables(0).NewRow
                dr("Empresa") = drAux("Empresa")
                dr("Cidade") = drAux("Cidade")
                dr("Endereco") = drAux("Endereco")
                dr("Bairro") = drAux("Bairro")
                dr("Fone") = drAux("Fone")
                dr("Cep") = drAux("Cep")
                dr("UF") = drAux("UF")
                dr("Cnpj") = drAux("Cnpj")
                dr("Insc") = drAux("Insc")
                dr("ValTot1") = drAux("ValTot1")
                dr("NumNot") = drAux("NumNot")
                dr("NumOrd") = drAux("NumOrd")
                dr("SeqOrd") = drAux("SeqOrd")
                dr("Vencto") = drAux("Vencto")
                dr("Cliente") = drAux("Cliente")
                dr("EndCli") = drAux("EndCli")
                dr("CepCli") = drAux("CepCli")
                dr("CidCli") = drAux("CidCli")
                dr("UFCli") = drAux("UFCli")
                dr("CnpjCli") = drAux("CnpjCli")
                dr("Praca") = drAux("Praca")
                dr("Ext001") = drAux("Ext001")
                dr("Ext002") = drAux("Ext002")
                dr("Ext003") = drAux("Ext003")
                dr("Inscli") = drAux("Inscli")
                dr("Via") = 2
                dr("Tipo") = drAux("Tipo")
                dr("Valnot") = drAux("Valnot")
                dr("Emissa") = drAux("Emissa")
                dr("BaiCli") = drAux("BaiCli")

                dsDuplicatasCopy.Tables(0).Rows.Add(dr)
            Next

            dsDuplicatasCopy.AcceptChanges()
            AlimentaCrptRelatorios(dsDuplicatasCopy, "~/Reports/Cr_Duplicata")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal Caminho As String)
        Try
            'LOGOTIPO
            Dim dt As DataTable = Ds.Tables.Add("Logotipo")

            dt.Columns.Add("Imagem", GetType(System.Byte()))

            Dim drImagem As DataRow = dt.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
            drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            dt.Rows.Add(drImagem)

            Dim crptRelatorio As New ReportDocument()
            crptRelatorio.FileName = Server.MapPath(Caminho & ".rpt")
            crptRelatorio.Load(crptRelatorio.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
            Dim arquivo As String = NomeArquivo

            crptRelatorio.SetDataSource(Ds)

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions

            crParameterFieldDefinitions = crptRelatorio.DataDefinition.ParameterFields()

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            Try
                crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message.ToString)
            Finally
                crptRelatorio.Close()
                crptRelatorio.Dispose()
                If IO.File.Exists(arquivo) Then
                    Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
                End If
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function ToDataTableToDsParcelas() As DataTable
        Dim dtParcelas As New DataTable("Parcelas")
        Try
            dtParcelas.Columns.Add("Vencimento", GetType(String))
            dtParcelas.Columns.Add("Valor", GetType(String))
            dtParcelas.Columns.Add("Saldo", GetType(String))

            Dim drItem As DataRow = dtParcelas.NewRow()

            drItem("Vencimento") = Me.txtData.Text
            drItem("Valor") = Me.txtValor.Text
            drItem("Saldo") = CDec(Me.txtTotalNota.Text) + CDec(Me.txtValor.Text)

            dtParcelas.Rows.Add(drItem)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return dtParcelas
    End Function

    Public Sub ParcelasAuto()
        Try
            If Not String.IsNullOrWhiteSpace(txtParcelas.Text) AndAlso CDec(txtParcelas.Text) > 0 And Not String.IsNullOrWhiteSpace(txtTotalNota.Text) AndAlso CDec(txtTotalNota.Text) > 0 And txtDataVenc.Text <> "" Then
                Dim intPos As Integer
                Dim dtParcelas As New DataTable("Parcelas")
                dtParcelas.Columns.Add("Vencimento", GetType(String))
                dtParcelas.Columns.Add("Valor", GetType(String))
                dtParcelas.Columns.Add("Saldo", GetType(String))

                Dim difUltima As Decimal = (CDec(txtParcelas.Text) * (Math.Round(CDec(Me.txtTotalNota.Text) / CDec(txtParcelas.Text), 2))) - CDec(Me.txtTotalNota.Text)
                Dim saldo As Decimal = Me.txtTotalNota.Text
                For intPos = 0 To (CDec(txtParcelas.Text) - 1)
                    Dim drItem As DataRow = dtParcelas.NewRow()
                    drItem("Vencimento") = IIf(intPos = 0, Me.txtDataVenc.Text, Format(CDate(Me.txtDataVenc.Text).AddMonths(intPos), "dd/MM/yyyy"))
                    drItem("Valor") = Math.Round(IIf(intPos + 1 <> CDec(txtParcelas.Text), CDec(Me.txtTotalNota.Text) / CDec(txtParcelas.Text), CDec(Me.txtTotalNota.Text) / CDec(txtParcelas.Text) - CDec(difUltima)), 2)
                    saldo = saldo - drItem("Valor")
                    drItem("Saldo") = CDec(saldo)
                    dtParcelas.Rows.Add(drItem)
                Next

                GridParcelas.DataSource = dtParcelas
                GridParcelas.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtTotalNota_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ParcelasAuto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridParcelas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtData.Text = GridParcelas.Rows(GridParcelas.SelectedIndex).Cells(1).Text
            txtValor.Text = GridParcelas.Rows(GridParcelas.SelectedIndex).Cells(2).Text
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Alterar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim saldo As Decimal = CDec(Me.txtTotalNota.Text)
            GridParcelas.Rows(GridParcelas.SelectedIndex).Cells(1).Text = txtData.Text
            GridParcelas.Rows(GridParcelas.SelectedIndex).Cells(2).Text = txtValor.Text
            For i As Integer = 0 To GridParcelas.Rows.Count - 1
                GridParcelas.Rows(i).Cells(3).Text = saldo - GridParcelas.Rows(i).Cells(2).Text
                saldo = saldo - GridParcelas.Rows(i).Cells(2).Text
            Next
            txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtValor.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEmitir_Click(sender As Object, e As EventArgs) Handles lnkEmitir.Click
        Try
            If ValidarCampos() Then BuscarRegistros()
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
            Funcoes.Ajuda(Me.Page, "DuplicataAvulsa")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class