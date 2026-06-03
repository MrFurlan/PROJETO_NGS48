Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ImportaFolha
    Inherits BasePage

    Private Mensagem As String
    Private Mes As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ImportaFolha", "ACESSAR") Then
                ddl.Carregar(lstUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Limpar_Campos()
        lstUnidadeNegocio.SelectedIndex = 0
        lstEmpresa.SelectedIndex = 0
        filUpload.Value = ""

        Session.Remove("ssFolha")
    End Sub

    Protected Sub lstUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(lstEmpresa, CarregarDDL.Tabela.Empresas, lstUnidadeNegocio.SelectedValue.ToString, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function ValidaDados(ByVal infoarquivo As Object) As Boolean
        Try
            If lstEmpresa.SelectedIndex = 0 Then
                Mensagem = "Empresa não foi selecionanda."
                Return False
            ElseIf infoarquivo.Name.ToString.Length = 0 Then
                Mensagem = "Arquivo não foi informado."
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Function ValidarArquivo(ByVal infoarquivo As Object) As Boolean
        If infoarquivo Is Nothing OrElse String.IsNullOrWhiteSpace(infoarquivo.Name) Then
            Return False
        End If

        Dim arquivo As String = Server.MapPath("Files/") & infoarquivo.Name
        Using objArquivo As New StreamReader(arquivo)
            Dim strLinha As String
            Dim campos() As String
            Dim intLinha As Integer
            Dim dsArquivo As New DataSet
            Dim arrEmpresa() As String
            Dim primeiraVez As Boolean = True
            arrEmpresa = lstEmpresa.SelectedValue.Split("-")
            Mensagem = ""

            dsArquivo = New DataSet
            Dim rowNew As DataRow

            With dsArquivo
                .Tables.Add()
                With .Tables(0).Columns
                    .Add("CentroCusto", Type.GetType("System.Int32"))
                    .Add("Movimento", Type.GetType("System.DateTime"))
                    .Add("ContaDebito", Type.GetType("System.String"))
                    .Add("ContaCredito", Type.GetType("System.String"))
                    .Add("Valor", Type.GetType("System.Decimal"))
                    .Add("Historico", Type.GetType("System.Int32"))
                    .Add("Descricao", Type.GetType("System.String"))
                    .Add("CNPJEmpresa", Type.GetType("System.String"))
                    .Add("EndEmpresa", Type.GetType("System.Int16"))
                End With
            End With

            Session("ssFolha") = dsArquivo

            Try
                If ddlSistema.SelectedValue = "PADRÃO" Then
                    Do While objArquivo.Peek >= 0
                        intLinha += 1

                        strLinha = objArquivo.ReadLine()

                        If String.IsNullOrWhiteSpace(strLinha) Then
                            Continue Do
                        End If

                        campos = strLinha.Split(",")

                        rowNew = dsArquivo.Tables(0).NewRow

                        Dim CentroDeCusto As New [Lib].Negocio.CentroDeCusto(campos(1))

                        If CentroDeCusto.Codigo = 0 Then
                            Mensagem = "Centro de Custo " & campos(1) & " inexistente na linha " & intLinha
                            Exit Do
                        End If
                        rowNew("CentroCusto") = CentroDeCusto.Codigo

                        rowNew("Movimento") = Left(campos(2), 4) & "/" & Mid(campos(2), 5, 2) & "/" & Mid(campos(2), 7, 2)

                        If primeiraVez Then
                            If Not Funcoes.VerificaAcesso(arrEmpresa(0), arrEmpresa(1), rowNew("Movimento"), "CONTABIL") Then
                                MsgBox(Me.Page, "Movimento Contábil Fechado para esta data: " & rowNew("Movimento"))
                                Return False
                            End If
                            primeiraVez = False
                        End If

                        Dim ContaDebito As New [Lib].Negocio.PlanoDeConta("", 0, campos(3))
                        If ContaDebito.Conta = "" Then
                            Mensagem = "Conta contábil de debito " & campos(4) & " inexistente na linha " & intLinha
                            Exit Do
                        End If
                        rowNew("ContaDebito") = ContaDebito.Conta

                        Dim ContaCredito As New [Lib].Negocio.PlanoDeConta("", 0, campos(4))
                        If ContaCredito.Conta = "" Then
                            Mensagem = "Conta contábil de Credito " & campos(3) & " inexistente na linha " & intLinha
                            Exit Do
                        End If
                        rowNew("ContaCredito") = ContaCredito.Conta

                        rowNew("Valor") = campos(5) / 100

                        Dim Historico As New [Lib].Negocio.Historico(Trim(campos(6)))
                        If Historico.Codigo = 0 Then
                            Mensagem = "Código Histórico " & campos(6) & " inexistente na linha " & intLinha
                            Exit Do
                        End If

                        rowNew("Historico") = Historico.Codigo
                        rowNew("Descricao") = Historico.Descricao

                        rowNew("CNPJEmpresa") = campos(7).Replace(".", "").Replace("/", "").Replace("-", "")
                        rowNew("EndEmpresa") = 0

                        If arrEmpresa(0) <> rowNew("CNPJEmpresa") Then
                            Mensagem = "Arquivo selecionado para importação não corresponde com empresa selecionada." & intLinha
                            Exit Do
                        End If

                        dsArquivo.Tables(0).Rows.Add(rowNew)
                        dsArquivo.AcceptChanges()
                    Loop

                ElseIf ddlSistema.SelectedValue = "GCI" Then

                    Do While objArquivo.Peek >= 0
                        intLinha += 1

                        strLinha = objArquivo.ReadLine()

                        If String.IsNullOrWhiteSpace(strLinha) Then
                            Continue Do
                        End If

                        campos = strLinha.Split(",")

                        rowNew = dsArquivo.Tables(0).NewRow

                        Dim CentroDeCusto As New [Lib].Negocio.CentroDeCusto("10701") 'campos(1))
                        rowNew("CentroCusto") = CentroDeCusto.Codigo

                        rowNew("Movimento") = "20" & Right(campos(2), 2) & "/" & Mid(campos(2), 3, 2) & "/" & Left(campos(2), 2)

                        If primeiraVez Then
                            If Not Funcoes.VerificaAcesso(arrEmpresa(0), arrEmpresa(1), rowNew("Movimento"), "CONTABIL") Then
                                MsgBox(Me.Page, "Movimento Contábil Fechado para esta data: " & rowNew("Movimento"))
                                Return False
                            End If
                            primeiraVez = False
                        End If

                        Dim ContaDebito As New [Lib].Negocio.PlanoDeConta("", 0, campos(4))
                        If ContaDebito.Conta = "" Then
                            'Mensagem = "Conta contábil de debito " & campos(4) & " inexistente na linha " & intLinha
                            'Exit Do
                        End If
                        rowNew("ContaDebito") = ContaDebito.Conta

                        Dim ContaCredito As New [Lib].Negocio.PlanoDeConta("", 0, campos(3))
                        If ContaCredito.Conta = "" Then
                            'Mensagem = "Conta contábil de Credito " & campos(3) & " inexistente na linha " & intLinha
                            'Exit Do
                        End If
                        rowNew("ContaCredito") = ContaCredito.Conta

                        rowNew("Valor") = campos(5) / 100

                        Dim Historico As New [Lib].Negocio.Historico(campos(6))
                        If Historico.Codigo = 0 Then
                            Mensagem = "Código Histórico " & campos(6) & " inexistente na linha " & intLinha
                            Exit Do
                        End If
                        rowNew("Historico") = Historico.Codigo
                        rowNew("Descricao") = Historico.Descricao

                        rowNew("CNPJEmpresa") = arrEmpresa(0)
                        rowNew("EndEmpresa") = arrEmpresa(1)

                        If arrEmpresa(0) <> rowNew("CNPJEmpresa") Then
                            Mensagem = "Arquivo selecionado para importação não corresponde com empresa selecionada." & intLinha
                            Exit Do
                        End If

                        dsArquivo.Tables(0).Rows.Add(rowNew)
                        dsArquivo.AcceptChanges()
                    Loop


                ElseIf ddlSistema.SelectedValue = "DOMINIO" Then

                    Do While objArquivo.Peek >= 0
                        intLinha += 1

                        strLinha = objArquivo.ReadLine()

                        If String.IsNullOrWhiteSpace(strLinha) Then
                            Continue Do
                        End If

                        campos = strLinha.Split(",")

                        rowNew = dsArquivo.Tables(0).NewRow

                        Dim CentroDeCusto As New [Lib].Negocio.CentroDeCusto("10701") 'campos(1))
                        rowNew("CentroCusto") = CentroDeCusto.Codigo

                        rowNew("Movimento") = "20" & Right(campos(2), 2) & "/" & Mid(campos(2), 3, 2) & "/" & Left(campos(2), 2)

                        If primeiraVez Then
                            If Not Funcoes.VerificaAcesso(arrEmpresa(0), arrEmpresa(1), rowNew("Movimento"), "CONTABIL") Then
                                MsgBox(Me.Page, "Movimento Contábil Fechado para esta data: " & rowNew("Movimento"))
                                Return False
                            End If
                            primeiraVez = False
                        End If

                        Dim ContaDebito As New [Lib].Negocio.PlanoDeConta("", 0, campos(3))
                        If ContaDebito.Conta = "" Then
                            'Mensagem = "Conta contábil de debito " & campos(4) & " inexistente na linha " & intLinha
                            'Exit Do
                        End If
                        rowNew("ContaDebito") = ContaDebito.Conta

                        Dim ContaCredito As New [Lib].Negocio.PlanoDeConta("", 0, campos(4))
                        If ContaCredito.Conta = "" Then
                            'Mensagem = "Conta contábil de Credito " & campos(3) & " inexistente na linha " & intLinha
                            'Exit Do
                        End If
                        rowNew("ContaCredito") = ContaCredito.Conta

                        rowNew("Valor") = campos(5) / 100

                        Dim Historico As New [Lib].Negocio.Historico(campos(6))
                        If Historico.Codigo = 0 Then
                            'AJUSTADO LAYOUT PARA NUTRI/BAXI - FURLAN 11/12/2024
                            If Left(Session("ssEmpresa").ToString, 8) = "05366261" OrElse Left(Session("ssEmpresa").ToString, 8) = "40938762" Then
                                'LIBERA PORQUE VAMOS PEGAR O HISTÓRIOCO DA COLUNA DO ARQUIVO
                                Historico.Codigo = campos(6)
                            Else
                                Mensagem = "Código Histórico " & campos(6) & " inexistente na linha " & intLinha
                                Exit Do
                            End If
                        End If

                        rowNew("Historico") = Historico.Codigo

                        'AJUSTADO LAYOUT PARA NUTRI/BAXI - FURLAN 11/12/2024
                        If Left(Session("ssEmpresa").ToString, 8) = "05366261" OrElse Left(Session("ssEmpresa").ToString, 8) = "40938762" Then
                            rowNew("Descricao") = Funcoes.SubstituirCaracteresEspeciaisMunicipio(campos(7))
                        Else
                            rowNew("Descricao") = Historico.Descricao
                        End If

                        rowNew("CNPJEmpresa") = arrEmpresa(0)
                        rowNew("EndEmpresa") = arrEmpresa(1)

                        If arrEmpresa(0) <> rowNew("CNPJEmpresa") Then
                            Mensagem = "Arquivo selecionado para importação não corresponde com empresa selecionada." & intLinha
                            Exit Do
                        End If

                        dsArquivo.Tables(0).Rows.Add(rowNew)
                        dsArquivo.AcceptChanges()
                    Loop

                End If


                Session("ssFolha") = dsArquivo

                If Mensagem.Length = 0 And dsArquivo.Tables(0).Rows.Count > 0 Then
                    Return True
                ElseIf Mensagem.Length = 0 And dsArquivo.Tables(0).Rows.Count = 0 Then
                    Mensagem = "Arquivo selecionado para importação vázio ou com problema."
                    Return False
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw New Exception(Funcoes.EliminarCaracteresEspeciais(ex.Message))
            End Try

            If Mensagem.Length > 0 Then
                Return False
            Else
                Return True
            End If
        End Using
    End Function

    Private Sub ImportarFolha()
        Dim infoarquivo As New IO.FileInfo(filUpload.PostedFile.FileName)
        If Upload(infoarquivo) Then
            If ValidaDados(infoarquivo) Then
                Dim strSQL As String
                Dim alSQL As New ArrayList
                Dim arrEmpresa() As String
                arrEmpresa = lstEmpresa.SelectedValue.Split("-")
                If ValidarArquivo(infoarquivo) Then

                    Dim Datas As String = ""
                    Dim Delimitador As String = ""
                    For Each row As DataRow In CType(Session("ssFolha"), DataSet).Tables(0).Rows
                        If Not Datas.Contains(row.Item("Movimento")) Then
                            Datas += Delimitador & row.Item("Movimento")
                            Delimitador = ";"
                        End If
                    Next

                    Dim Movimento() As String = Datas.Split(";")
                    For i As Integer = 0 To Movimento.GetUpperBound(0)
                        If Movimento(i).ToString.Length > 0 Then
                            strSQL = "DELETE Razao " & vbCrLf & _
                                     "WHERE Empresa_Id = '" & arrEmpresa(0) & "' " & vbCrLf & _
                                     "AND EndEmpresa_Id = " & arrEmpresa(1) & " " & vbCrLf & _
                                     "AND Movimento_Id = '" & Movimento(i).ToSqlDate() & "' " & vbCrLf & _
                                     "AND Lote_Id = 28 " & vbCrLf
                            alSQL.Add(strSQL)
                        End If
                    Next

                    Dim intSeq As Integer
                    Dim ds_Cotacoes As New DataSet
                    Dim drCotacao As DataRow
                    Dim valorcotacao As Double
                    Datas = ""

                    For Each row As DataRow In CType(Session("ssFolha"), DataSet).Tables(0).Rows
                        If Datas.Length = 0 OrElse Datas <> row.Item("Movimento").ToString Then
                            Datas = row.Item("Movimento").ToString
                            'Busca Valor da Contação da Moeda
                            strSQL = "Select Data_Id,Realizado,Indice From Cotacoes " & vbCrLf & _
                                     " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & CDate(row.Item("Movimento")).ToSqlDate() & "')" & vbCrLf
                            ds_Cotacoes = Banco.ConsultaDataSet(strSQL, "Cotacoes")
                            For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                                valorcotacao = drCotacao("Indice")
                            Next
                        End If

                        If row.Item("ContaDebito") <> "" Then
                            intSeq += 1
                            strSQL = "INSERT INTO Razao " & vbCrLf & _
                                     "(Empresa_Id, EndEmpresa_Id,UnidadeDeNegocio, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, " & vbCrLf & _
                                     "Lote_Id, Sequencia_Id, Titulo, Pedido, Produto, Custo, Indexador, DataMoeda, " & vbCrLf & _
                                     "DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado) " & vbCrLf & _
                                     "VALUES (" & vbCrLf & _
                                     "'" & row.Item("CNPJEmpresa") & "', " & row.Item("EndEmpresa") & ",'" & lstUnidadeNegocio.SelectedValue & "',  " & vbCrLf & _
                                     "'" & row.Item("ContaDebito") & "', '', 0, " & vbCrLf & _
                                     "'" & CDate(row.Item("Movimento")).ToSqlDate() & "', " & vbCrLf & _
                                     "28, " & intSeq & ", 0, NULL, '0',  " & row.Item("CentroCusto") & ", 3, " & vbCrLf & _
                                     "'" & CDate(row.Item("Movimento")).ToSqlDate() & "', " & vbCrLf & _
                                     "" & row.Item("Valor").ToString.Replace(",", ".") & ", 0, " & vbCrLf & _
                                     "" & Math.Round(row.Item("Valor") / valorcotacao, 2).ToString.Replace(",", ".") & ", 0, " & vbCrLf

                            If Left(Session("ssEmpresa").ToString, 8) = "05366261" OrElse Left(Session("ssEmpresa").ToString, 8) = "40938762" Then
                                strSQL &= "'" & RTrim(row.Item("Descricao")) & "', 'P')" & vbCrLf
                            Else
                                strSQL &= "'" & RTrim(row.Item("Descricao")) & " REF. MÊS " & CDate(row.Item("Movimento")).ToSqlDate() & "', 'P')" & vbCrLf
                            End If

                            alSQL.Add(strSQL)
                        End If


                        If row.Item("ContaCredito") <> "" Then
                            intSeq += 1
                            strSQL = "INSERT INTO Razao " & vbCrLf & _
                                     "(Empresa_Id, EndEmpresa_Id,UnidadeDeNegocio, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, " & vbCrLf & _
                                     "Lote_Id, Sequencia_Id, Titulo, Pedido, Produto, Custo, Indexador, DataMoeda, " & vbCrLf & _
                                     "DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado) " & vbCrLf & _
                                     "VALUES (" & vbCrLf & _
                                     "'" & row.Item("CNPJEmpresa") & "', " & row.Item("EndEmpresa") & ",'" & lstUnidadeNegocio.SelectedValue & "',  " & vbCrLf & _
                                     "'" & row.Item("ContaCredito") & "', '', 0, " & vbCrLf & _
                                     "'" & CDate(row.Item("Movimento")).ToSqlDate() & "', " & vbCrLf & _
                                     "28, " & intSeq & ", 0, NULL, '0',  " & row.Item("CentroCusto") & ", 3, " & vbCrLf & _
                                     "'" & CDate(row.Item("Movimento")).ToSqlDate() & "', " & vbCrLf & _
                                     "0, " & row.Item("Valor").ToString.Replace(",", ".") & ", " & vbCrLf & _
                                     "0, " & Math.Round(row.Item("Valor") / valorcotacao, 2).ToString.Replace(",", ".") & ", " & vbCrLf

                            If Left(Session("ssEmpresa").ToString, 8) = "05366261" OrElse Left(Session("ssEmpresa").ToString, 8) = "40938762" Then
                                strSQL &= "'" & RTrim(row.Item("Descricao")) & "', 'P')" & vbCrLf
                            Else
                                strSQL &= "'" & RTrim(row.Item("Descricao")) & " REF. MÊS " & CDate(row.Item("Movimento")).ToSqlDate() & "', 'P')" & vbCrLf
                            End If


                            alSQL.Add(strSQL)
                        End If
                    Next

                    If Banco.GravaBanco(alSQL) = True Then
                        MsgBox(Me.Page, "Importação de " & MonthName(CDate(Datas).Month) & "/" & CDate(Datas).Year & " realizada com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, Mensagem.ToString)
                End If
            Else
                MsgBox(Me.Page, Mensagem.ToString)
            End If
        Else
            MsgBox(Me.Page, Mensagem.ToString)
        End If
    End Sub

    Private Function Upload(ByVal infoarquivo As Object) As Boolean
        Try
            'Verificamos se tem alguma coisa postada 
            If Not IsNothing(filUpload.PostedFile) Then
                'Pegamos as informacoes do arquivo postado 
                'Definimos onde ele será salvo 
                Dim strCaminho As String = Server.MapPath("Files/") & infoarquivo.Name

                If File.Exists(strCaminho) Then File.Delete(strCaminho)

                'Salvamos o mesmo 
                filUpload.PostedFile.SaveAs(strCaminho)
            Else
                MsgBox(Me.Page, "Selecione um arquivo.")
                Return False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return True
    End Function

    Protected Sub lnkImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportar.Click
        Try
            ImportarFolha()
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ImportaFolha")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class