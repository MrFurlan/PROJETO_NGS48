Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ListTituloXContaContabil
        Inherits List(Of Novo.TituloXContaContabil)

#Region "Fields"
        Private _Titulo As Novo.TituloNovo
        Private _TituloOld As Titulo
        Private _EncargoValorDocumento As Novo.TituloXContaContabil
        Private _EncargoValorLiquido As Novo.TituloXContaContabil
#End Region

#Region "Property"

        Public Property Titulo() As Novo.TituloNovo
            Get
                Return _Titulo
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _Titulo = value
            End Set
        End Property

        Public Property TituloOld() As Titulo
            Get
                Return _TituloOld
            End Get
            Set(ByVal value As Titulo)
                _TituloOld = value
            End Set
        End Property

        Public Property EncargoValorDocumento() As Novo.TituloXContaContabil
            Get
                Return _EncargoValorDocumento
            End Get
            Set(ByVal value As Novo.TituloXContaContabil)
                _EncargoValorDocumento = value
            End Set
        End Property

        Public Property EncargoValorLiquido() As Novo.TituloXContaContabil
            Get
                Return _EncargoValorLiquido
            End Get
            Set(ByVal value As Novo.TituloXContaContabil)
                _EncargoValorLiquido = value
            End Set
        End Property

        Public ReadOnly Property TotalDebitos() As Decimal
            Get
                Dim vlr As Decimal = 0
                For Each row As Novo.TituloXContaContabil In Me
                    If row.DC = "D" Then vlr += IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, row.ValorOficial, row.ValorMoeda)
                Next
                Return vlr
            End Get
        End Property

        Public ReadOnly Property TotalCreditos() As Decimal
            Get
                Dim vlr As Decimal = 0
                For Each row As Novo.TituloXContaContabil In Me
                    If row.DC = "C" Then vlr += IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, row.ValorOficial, row.ValorMoeda)
                Next
                Return vlr
            End Get
        End Property
#End Region

#Region "Contrutor"
        Public Sub New()

        End Sub

        Public Sub New(ByRef pTitulo As Novo.TituloNovo)

            _Titulo = pTitulo
            _Titulo.DesligarControles()
            Dim sql As String = ""
            sql = " Select sum(SB.TemTitulo) TemTitulo," & vbCrLf &
                  "        SB.Conta_Id," & vbCrLf &
                  "        SB.Descricao," & vbCrLf &
                  "        sum(ValorOficial) as ValorOficial," & vbCrLf &
                  "        SUM(ValorMoeda) as ValorMoeda," & vbCrLf &
                  "        max(sb.DC_Cadastro) as DC_Cadastro," & vbCrLf &
                  "        max(sb.dc_titulo) as DC_Titulo " & vbCrLf &
                  "   into #Temp " & vbCrLf &
                  "   from" & vbCrLf &
                  "      (" & vbCrLf &
                  "       SELECT 0 as TemTitulo," & vbCrLf &
                  "              EPC.ContaEncargo_Id as Conta_Id," & vbCrLf &
                  "              PC.Titulo as Descricao," & vbCrLf &
                  "              convert(numeric(18,2),0) as ValorOficial," & vbCrLf &
                  "              convert(numeric(18,2),0) as ValorMoeda," & vbCrLf &
                  "              case when 'R' = '" & pTitulo.ReceberPagar & "' then PC.Receber" & vbCrLf &
                  "                   when 'P' = '" & pTitulo.ReceberPagar & "' then PC.Pagar" & vbCrLf &
                  "                   else ''" & vbCrLf &
                  "              end as DC_Cadastro," & vbCrLf &
                  "              '' as DC_Titulo" & vbCrLf &
                  "         FROM EncargosPlanoDeContas EPC" & vbCrLf &
                  "        inner Join PlanoDeContas PC" & vbCrLf &
                  "           on EPC.Empresa_id      = PC.Empresa_Id" & vbCrLf &
                  "          and EPC.EndEmpresa_id   = PC.EndEmpresa_Id" & vbCrLf &
                  "          and EPC.ContaEncargo_Id = PC.Conta_Id" & vbCrLf &
                  "        where EPC.Conta_id        ='" & pTitulo.CodigoContaContabilCliFor & "'" & vbCrLf &
                  "        union " & vbCrLf &
                  "       SELECT 1 as TemTitulo," & vbCrLf &
                  "              TCC.Conta_Id," & vbCrLf &
                  "              PC.Titulo as Descricao," & vbCrLf &
                  "              isnull(TCC.ValorOficial,0) as ValorOficial," & vbCrLf &
                  "              isnull(TCC.ValorMoeda,0) as ValorMoeda," & vbCrLf &
                  "              '' as DC_Cadastro," & vbCrLf &
                  "              TCC.DC_Id " & vbCrLf &
                  "         FROM TitulosxContaContabil TCC" & vbCrLf &
                  "        inner Join PlanoDeContas PC" & vbCrLf &
                  "           on PC.Empresa_id     = '99999999999999'" & vbCrLf &
                  "          and PC.EndEmpresa_id  = 0" & vbCrLf &
                  "          and PC.Conta_Id       = TCC.Conta_Id" & vbCrLf &
                  "        where TCC.Titulo_Id    = " & pTitulo.Codigo & vbCrLf &
                  "       )as sb" & vbCrLf &
                  "       group by SB.Conta_id," & vbCrLf &
                  "                SB.Descricao" & vbCrLf &
                  "       Order by case when SB.Conta_Id ='" & Titulo.CodigoContaContabilCliFor & "' then 1" & vbCrLf &
                  "                     when SB.Conta_Id in ('" & Titulo.CodigoContaContabilRecPag & "','" & Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco & "') then 3" & vbCrLf &
                  "                     else 2 " & vbCrLf &
                  "                 end," & vbCrLf &
                  "                 SB.Conta_Id;" & vbCrLf

            sql &= "Select Conta_Id," & vbCrLf &
                   "       Descricao," & vbCrLf &
                   "       ValorOficial," & vbCrLf &
                   "       ValorMoeda," & vbCrLf &
                   "       Case " & vbCrLf &
                   "         When TemTitulo = 1" & vbCrLf &
                   "           then DC_Titulo" & vbCrLf &
                   "           else case when len(isnull(DC_Cadastro,'')) = 0 then 'I' else DC_Cadastro end" & vbCrLf &
                   "       end as DC" & vbCrLf &
                   "  from #Temp" & vbCrLf &
                   " Where ValorOficial +  ValorMoeda > 0 or Conta_Id in ('" & Titulo.CodigoContaContabilCliFor & "','" & Titulo.CodigoContaContabilRecPag & "')"

            Dim ds As DataSet
            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(sql, "EncargosPlanoDeConta")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim TxE As New Novo.TituloXContaContabil(pTitulo)
                TxE.CodigoContaEncargo = row("Conta_Id")
                TxE.Descricao = row("Descricao")
                TxE.ValorOficial = row("ValorOficial")
                TxE.ValorMoeda = row("ValorMoeda")
                TxE.DC = row("DC")
                Me.Add(TxE)
            Next
            AtualizaValores()
            _Titulo.LigarControles()
        End Sub

        Public Sub New(ByRef pTitulo As Titulo)

            _TituloOld = pTitulo
            _TituloOld.DesligarControles()
            Dim sql As String = ""
            sql = " Select sum(SB.TemTitulo) TemTitulo," & vbCrLf &
                  "        SB.Conta_Id," & vbCrLf &
                  "        SB.Descricao," & vbCrLf &
                  "        sum(ValorOficial) as ValorOficial," & vbCrLf &
                  "        SUM(ValorMoeda) as ValorMoeda," & vbCrLf &
                  "        max(sb.DC_Cadastro) as DC_Cadastro," & vbCrLf &
                  "        max(sb.dc_titulo) as DC_Titulo " & vbCrLf &
                  "   into #Temp " & vbCrLf &
                  "   from" & vbCrLf &
                  "      (" & vbCrLf &
                  "       SELECT 0 as TemTitulo," & vbCrLf &
                  "              EPC.ContaEncargo_Id as Conta_Id," & vbCrLf &
                  "              PC.Titulo as Descricao," & vbCrLf &
                  "              convert(numeric(18,2),0) as ValorOficial," & vbCrLf &
                  "              convert(numeric(18,2),0) as ValorMoeda," & vbCrLf &
                  "              case when 'R' = '" & pTitulo.ReceberPagar & "' then PC.Receber" & vbCrLf &
                  "                   when 'P' = '" & pTitulo.ReceberPagar & "' then PC.Pagar" & vbCrLf &
                  "                   else ''" & vbCrLf &
                  "              end as DC_Cadastro," & vbCrLf &
                  "              '' as DC_Titulo" & vbCrLf &
                  "         FROM EncargosPlanoDeContas EPC" & vbCrLf &
                  "        inner Join PlanoDeContas PC" & vbCrLf &
                  "           on EPC.Empresa_id      = PC.Empresa_Id" & vbCrLf &
                  "          and EPC.EndEmpresa_id   = PC.EndEmpresa_Id" & vbCrLf &
                  "          and EPC.ContaEncargo_Id = PC.Conta_Id" & vbCrLf &
                  "        where EPC.Conta_id        ='" & pTitulo.CodigoContaContabilCliFor & "'" & vbCrLf &
                  "        union " & vbCrLf &
                  "       SELECT 1 as TemTitulo," & vbCrLf &
                  "              TCC.Conta_Id," & vbCrLf &
                  "              PC.Titulo as Descricao," & vbCrLf &
                  "              isnull(TCC.ValorOficial,0) as ValorOficial," & vbCrLf &
                  "              isnull(TCC.ValorMoeda,0) as ValorMoeda," & vbCrLf &
                  "              '' as DC_Cadastro," & vbCrLf &
                  "              TCC.DC_Id " & vbCrLf &
                  "         FROM TitulosxContaContabil TCC" & vbCrLf &
                  "        inner Join PlanoDeContas PC" & vbCrLf &
                  "           on PC.Empresa_id     = '99999999999999'" & vbCrLf &
                  "          and PC.EndEmpresa_id  = 0" & vbCrLf &
                  "          and PC.Conta_Id       = TCC.Conta_Id" & vbCrLf &
                  "        where TCC.Titulo_Id    = " & pTitulo.Codigo & vbCrLf &
                  "       )as sb" & vbCrLf &
                  "       group by SB.Conta_id," & vbCrLf &
                  "                SB.Descricao" & vbCrLf &
                  "       Order by case when SB.Conta_Id ='" & Titulo.CodigoContaContabilCliFor & "' then 1" & vbCrLf &
                  "                     when SB.Conta_Id in ('" & Titulo.CodigoContaContabilRecPag & "','" & Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco & "') then 3" & vbCrLf &
                  "                     else 2 " & vbCrLf &
                  "                 end," & vbCrLf &
                  "                 SB.Conta_Id;" & vbCrLf

            sql &= "Select Conta_Id," & vbCrLf &
                   "       Descricao," & vbCrLf &
                   "       ValorOficial," & vbCrLf &
                   "       ValorMoeda," & vbCrLf &
                   "       Case " & vbCrLf &
                   "         When TemTitulo = 1" & vbCrLf &
                   "           then DC_Titulo" & vbCrLf &
                   "           else case when len(isnull(DC_Cadastro,'')) = 0 then 'I' else DC_Cadastro end" & vbCrLf &
                   "       end as DC" & vbCrLf &
                   "  from #Temp" & vbCrLf &
                   " Where ValorOficial +  ValorMoeda > 0 or Conta_Id in ('" & Titulo.CodigoContaContabilCliFor & "','" & Titulo.CodigoContaContabilRecPag & "')"

            Dim ds As DataSet
            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(sql, "EncargosPlanoDeConta")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim TxE As New Novo.TituloXContaContabil(pTitulo)
                TxE.CodigoContaEncargo = row("Conta_Id")
                TxE.Descricao = row("Descricao")
                TxE.ValorOficial = row("ValorOficial")
                TxE.ValorMoeda = row("ValorMoeda")
                TxE.DC = row("DC")
                Me.Add(TxE)
            Next
            AtualizaValores()
            _TituloOld.LigarControles()
        End Sub

#End Region

#Region "Methods"
        Public Function Salvar(Optional ByVal UsaNumerador As Boolean = True) As Boolean
            Dim Banco As New AcessaBanco
            Dim sqls As New ArrayList

            sqls.Clear()
            Me.SalvarSQL(sqls)

            If sqls.Count = 0 OrElse Banco.GravaBanco(sqls) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub SalvarSQL(ByRef Sqls As ArrayList)
            'Apaga Todos
            Dim sql As String
            sql = "Delete TitulosxContaContabil WHERE Titulo_Id = " & Titulo.Codigo
            Sqls.Add(sql)

            If Titulo.IUD = "D" Then Exit Sub

            'Salva lista Atual
            For Each TxE As Novo.TituloXContaContabil In Me
                TxE.IUD = "I"
                TxE.SalvarSql(Sqls)
            Next
        End Sub

        Public Sub AtualizaValores()
            If Titulo.IndiceTitulo = 0 AndAlso Titulo.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then Exit Sub

            Dim EstavaControlando As Boolean = Titulo.Controlando
            Titulo.DesligarControles()

            'Nos Lancamentos contabeis qdo a conta nao tiver relacionamento entre contas assume a principal como Produto
            If Me.Count = 0 Then
                Dim Produto As New Novo.TituloXContaContabil(Titulo)
                Produto.CodigoContaEncargo = Titulo.CodigoContaContabilCliFor
                Produto.Descricao = Titulo.ContaContabilCliFor.Titulo
                Produto.ValorMoeda = 0
                Produto.ValorOficial = 0
                Me.Add(Produto)
                Me.EncargoValorDocumento = Produto
            End If

            Dim VlrliquidoOficial As Decimal
            Dim VlrliquidoMoeda As Decimal
            Dim Liquido As Novo.TituloXContaContabil = Nothing

            If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                For Each enc As Novo.TituloXContaContabil In Me
                    If enc.CodigoContaEncargo = Titulo.CodigoContaContabilRecPag Or enc.CodigoContaEncargo = Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco Then
                        Liquido = enc
                    Else
                        If enc.CodigoContaEncargo = Titulo.CodigoContaContabilCliFor Then
                            Me.EncargoValorDocumento = enc
                            If Titulo.IUD = "I" Or Me.EncargoValorDocumento.DC = "" Then
                                If Titulo.ReceberPagar = "P" Or Titulo.ReceberPagar = "C" Then
                                    Me.EncargoValorDocumento.DC = "D"
                                Else
                                    Me.EncargoValorDocumento.DC = "C"
                                End If
                            End If
                        End If
                        enc.ValorMoeda = Funcoes.ConverteMoeda(enc.ValorOficial, Titulo.IndiceTitulo, eTiposMoeda.MoedaEstrangeira)
                        VlrliquidoOficial += IIf(enc.DC = "D", enc.ValorOficial * -1, enc.ValorOficial)
                        VlrliquidoMoeda += IIf(enc.DC = "D", enc.ValorMoeda * -1, enc.ValorMoeda)
                    End If
                Next
            Else
                For Each enc As Novo.TituloXContaContabil In Me
                    If enc.CodigoContaEncargo = Titulo.CodigoContaContabilRecPag Or enc.CodigoContaEncargo = Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco Then
                        Liquido = enc
                    Else
                        If enc.CodigoContaEncargo = Titulo.CodigoContaContabilCliFor Then
                            Me.EncargoValorDocumento = enc
                            If Titulo.IUD = "I" Or Me.EncargoValorDocumento.DC = "" Then
                                If Titulo.ReceberPagar = "P" Or Titulo.ReceberPagar = "C" Then
                                    Me.EncargoValorDocumento.DC = "D"
                                Else
                                    Me.EncargoValorDocumento.DC = "C"
                                End If
                            End If
                        End If
                        If Titulo.Controlando Then enc.ValorOficial = Funcoes.ConverteMoeda(enc.ValorMoeda, Titulo.IndiceTitulo, eTiposMoeda.Oficial)
                        VlrliquidoOficial += IIf(enc.DC = "D", enc.ValorOficial * -1, enc.ValorOficial)
                        VlrliquidoMoeda += IIf(enc.DC = "D", enc.ValorMoeda * -1, enc.ValorMoeda)
                    End If
                Next
            End If

            ConfigurarLiquido(Liquido, VlrliquidoMoeda, VlrliquidoOficial)

            If EstavaControlando Then Titulo.LigarControles()
        End Sub

        Public Sub AtualizaLiquido()
            Dim VlrliquidoOficial As Decimal
            Dim VlrliquidoMoeda As Decimal
            Dim Liquido As Novo.TituloXContaContabil = Nothing

            For Each enc As Novo.TituloXContaContabil In Me
                If enc.CodigoContaEncargo = Titulo.CodigoContaContabilRecPag Or enc.CodigoContaEncargo = Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco Then
                    Liquido = enc
                Else
                    VlrliquidoOficial += IIf(enc.DC = "D", enc.ValorOficial * -1, enc.ValorOficial)
                    VlrliquidoMoeda += IIf(enc.DC = "D", enc.ValorMoeda * -1, enc.ValorMoeda)
                End If
            Next

            Dim EstavaControlando As Boolean = Titulo.Controlando
            Titulo.DesligarControles()

            ConfigurarLiquido(Liquido, VlrliquidoMoeda, VlrliquidoOficial)

            If EstavaControlando Then Titulo.LigarControles()
        End Sub

        Private Sub ConfigurarLiquido(ByRef Liquido As Novo.TituloXContaContabil, ByVal VlrLiquidoMoeda As Decimal, ByVal VlrLiquidoOficial As Decimal)
            If Not Liquido Is Nothing Then
                Liquido.ValorMoeda = Math.Abs(VlrLiquidoMoeda)
                Liquido.ValorOficial = Math.Abs(VlrLiquidoOficial)
            Else
                Liquido = New Novo.TituloXContaContabil(Titulo)

                If Titulo.CodigoContaContabilRecPag.Length > 0 Then
                    Liquido.CodigoContaEncargo = Titulo.CodigoContaContabilRecPag
                    Liquido.Descricao = Titulo.ContaContabilRecPag.Titulo
                Else
                    Liquido.CodigoContaEncargo = Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco
                    Liquido.Descricao = Titulo.EmpresaRecPag.Empresa.ContaGrupoBanco.Titulo
                    Titulo.CodigoContaContabilRecPag = Titulo.EmpresaRecPag.Empresa.CodigoContaGrupoBanco
                End If

                Liquido.ValorMoeda = Math.Abs(VlrLiquidoMoeda)
                Liquido.ValorOficial = Math.Abs(VlrLiquidoOficial)
                Me.Add(Liquido)
            End If
            Me.EncargoValorLiquido = Liquido

            'Deb/Cred do liquido é sempre o contrario do documento
            If Titulo.IUD = "I" Or Me.EncargoValorLiquido.DC = "" Then
                If Titulo.ReceberPagar = "P" Or Titulo.ReceberPagar = "C" Then
                    Me.EncargoValorLiquido.DC = "C"
                Else
                    Me.EncargoValorLiquido.DC = "D"
                End If
            End If
        End Sub

        Public Sub LimpaContas()
            Me.RemoveAll(Function(s) s.ValorOficial + s.ValorMoeda = 0 And Not s.Equals(Me.EncargoValorDocumento) And Not s.Equals(EncargoValorLiquido))
        End Sub
#End Region
    End Class

    <Serializable()> _
    Public Class TituloXContaContabil
#Region "Construtor"

        Public Sub New(ByVal pTitulo As Novo.TituloNovo)
            _Titulo = pTitulo
        End Sub

        Public Sub New(ByVal pTitulo As Titulo)
            _TituloOld = pTitulo
        End Sub

        Public Sub New()

        End Sub
#End Region

#Region "Fields"
        Private _IUD As String = ""
        Private _Titulo As Novo.TituloNovo
        Private _TituloOld As Titulo
        Private _CodigoContaEncargo As String = ""
        Private _ContaEncargo As PlanoDeConta
        Private _Descricao As String
        Private _ValorOficial As Decimal
        Private _ValorMoeda As Decimal
        Private _DC As String = ""
        Private _CodigoCentroDeCusto As String = ""
        Private _CentroDeCusto As CentroDeCusto
#End Region

#Region "Property"
        Public Property IUD() As String
            Get
                Return _IUD
            End Get
            Set(ByVal value As String)
                _IUD = value
            End Set
        End Property

        Public Property Titulo() As Novo.TituloNovo
            Get
                Return _Titulo
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _Titulo = value
            End Set
        End Property

        Public Property TituloOld() As Titulo
            Get
                Return _TituloOld
            End Get
            Set(ByVal value As Titulo)
                _TituloOld = value
            End Set
        End Property

        Public Property CodigoContaEncargo() As String
            Get
                Return _CodigoContaEncargo
            End Get
            Set(ByVal value As String)
                _CodigoContaEncargo = value
                _ContaEncargo = Nothing
                _Descricao = ""
            End Set
        End Property

        Public Property Descricao() As String
            Get
                If _Descricao.Length = 0 Then _Descricao = ContaEncargo.Titulo
                Return _Descricao
            End Get
            Set(ByVal value As String)
                _Descricao = value
            End Set
        End Property

        Public Property ContaEncargo() As PlanoDeConta
            Get
                If _ContaEncargo Is Nothing And CodigoContaEncargo.Length > 0 Then _ContaEncargo = New PlanoDeConta("", 0, CodigoContaEncargo)
                Return _ContaEncargo
            End Get
            Set(ByVal value As PlanoDeConta)
                _ContaEncargo = value
            End Set
        End Property

        Public Property ValorOficial() As Decimal
            Get
                Return _ValorOficial
            End Get
            Set(ByVal value As Decimal)
                If Titulo.Controlando Then
                    If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                        _ValorOficial = value
                        '_ValorMoeda = Funcoes.ConverteMoeda(_ValorOficial, Titulo.IndiceTitulo, eTiposMoeda.MoedaEstrangeira, True, False, 2)
                        Titulo.Valores.AtualizaLiquido()
                    Else
                        _ValorOficial = value
                    End If
                Else
                    _ValorOficial = value
                End If
            End Set
        End Property

        Public Property ValorMoeda() As Decimal
            Get
                Return _ValorMoeda
            End Get
            Set(ByVal value As Decimal)
                If Titulo.Controlando Then
                    If Titulo.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                        _ValorMoeda = value
                        _ValorOficial = Funcoes.ConverteMoeda(_ValorMoeda, Titulo.IndiceTitulo, eTiposMoeda.Oficial, True, False, 2)
                        Titulo.Valores.AtualizaLiquido()
                    End If
                Else
                    _ValorMoeda = value
                End If
            End Set
        End Property

        'Segue a moeda do titulo
        Public Property Valor As Decimal
            Get
                If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return ValorOficial
                Else
                    Return ValorMoeda
                End If
            End Get
            Set(value As Decimal)
                If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    ValorOficial = value
                Else
                    ValorMoeda = value
                End If
            End Set
        End Property

        Public Property DC() As String
            Get
                Return _DC
            End Get
            Set(ByVal value As String)
                _DC = value
            End Set
        End Property

        Public Property CodigoCentroDeCusto() As String
            Get
                Return _CodigoCentroDeCusto
            End Get
            Set(ByVal value As String)
                _CodigoCentroDeCusto = value
                _CentroDeCusto = Nothing
            End Set
        End Property

        Public ReadOnly Property CentroDeCusto() As CentroDeCusto
            Get
                If _CentroDeCusto Is Nothing And _CodigoCentroDeCusto.Length > 0 Then _CentroDeCusto = New CentroDeCusto(_CodigoCentroDeCusto)
                Return _CentroDeCusto
            End Get
        End Property
#End Region

#Region "Methods"
        Public Function Salvar() As Boolean
            Dim Banco As New AcessaBanco
            Dim sqls As New ArrayList
            SalvarSql(sqls)
            Return Banco.GravaBanco(sqls)
        End Function

        Public Function SalvarSql(ByRef sqls As ArrayList) As ArrayList
            Dim strSQL As String
            Dim ObjBanco As New AcessaBanco

            If Me.IUD = "U" And IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, ValorOficial, ValorMoeda) = 0 Then Me.IUD = "D"
            If Me.IUD = "I" And IIf(Titulo.Moeda.Classificacao = eTiposMoeda.Oficial, ValorOficial, ValorMoeda) = 0 AndAlso Not Titulo.CodigoProvisao = eProvisao.Previsao Then Me.IUD = ""

            Select Case Me.IUD
                Case "I"
                    strSQL = "INSERT INTO TitulosxContaContabil(Titulo_Id, Conta_Id, ValorOficial, ValorMoeda, DC_Id, CentroDeCusto) " & vbCrLf & _
                             " VALUES (" & Titulo.Codigo & ",'" & CodigoContaEncargo & "'," & Str(Math.Round(ValorOficial, 2)) & "," & Str(Math.Round(ValorMoeda, 2)) & ",'" & DC & "'," & IIf(CodigoCentroDeCusto.Length > 0, "'" & CodigoCentroDeCusto & "'", "NULL") & ")"
                    sqls.Add(strSQL)
                Case "U"
                    strSQL = " UPDATE TitulosxContaContabil Set " & vbCrLf & _
                             "   ValorOficial  = " & Str(Math.Round(ValorOficial, 2)) & vbCrLf & _
                             "  ,ValorMoeda    = " & Str(Math.Round(ValorMoeda, 2)) & vbCrLf & _
                             "  ,DC_Id         ='" & DC & "'" & vbCrLf & _
                             "  ,CentroDeCusto = " & IIf(CodigoCentroDeCusto.Length > 0, "'" & CodigoCentroDeCusto & "'", "NULL") & vbCrLf & _
                             " WHERE Titulo_Id       = " & Titulo.Codigo & vbCrLf & _
                             "   AND Conta_Id ='" & CodigoContaEncargo & "'"
                    sqls.Add(strSQL)
                Case "D"
                    strSQL = "Delete TitulosxContaContabil " & vbCrLf & _
                             " WHERE Titulo_Id       = " & Titulo.Codigo & vbCrLf & _
                             "   AND Conta_Id ='" & CodigoContaEncargo & "'"
                    sqls.Add(strSQL)
            End Select

            Return sqls
        End Function
#End Region
    End Class

End Namespace
