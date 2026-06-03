Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXProduto
    Inherits List(Of ClienteXProduto)
    Implements IBaseEntity

    Dim ds As DataSet
    Dim sql As String
    Dim objBanco As New AcessaBanco

#Region "Constructor"
    Public Sub New()
        sql = " SELECT Tp.Produto_Id AS codProduto, Tp.Situacao ProdutoSituacao,				" & vbCrLf & _
              "		 Tp.Cliente_Id AS codCliente, Tp.EndCliente_Id AS codEndCliente,		    " & vbCrLf & _
              "		 Tp.Empresa_Id AS codEmpresa, Tp.EndEmpresa_Id AS codEndEmpresa,		    " & vbCrLf & _
              "		 TpXE.Encargo_Id AS codEncargo, TpXE.Situacao AS EncargoSituacao,		    " & vbCrLf & _
              "		 Tp.Dolar AS Dolar, Tp.DataPTax AS DataPTax, Tp.ValorPTax AS ValorPTax,	    " & vbCrLf & _
              "		 Tp.ValorMoeda AS ValorMoeda, Tp.ValorOficial AS ValorOficial			    " & vbCrLf & _
              " FROM TabelaDePrecos Tp														    " & vbCrLf & _
              " INNER JOIN TabelaDePrecosXEncargos TpXE										    " & vbCrLf & _
              " ON TpXE.Produto_Id = Tp.Produto_Id											    " & vbCrLf & _
              " AND TpXE.Cliente_Id = Tp.Cliente_Id											    " & vbCrLf & _
              " AND TpXE.EndCliente_Id = Tp.EndCliente_Id									    " & vbCrLf & _
              " AND TpXE.Empresa_Id = Tp.Empresa_Id											    " & vbCrLf & _
              " AND TpXE.EndEmpresa_Id = Tp.EndEmpresa_Id										"
        setObject(sql)
    End Sub

    Public Sub New(Cliente As String, EndCliente As String)
        sql = " SELECT Tp.Produto_Id AS codProduto, Tp.Situacao ProdutoSituacao,				" & vbCrLf & _
              "		 Tp.Cliente_Id AS codCliente, Tp.EndCliente_Id AS codEndCliente,		    " & vbCrLf & _
              "		 Tp.Empresa_Id AS codEmpresa, Tp.EndEmpresa_Id AS codEndEmpresa,		    " & vbCrLf & _
              "		 TpXE.Encargo_Id AS codEncargo, TpXE.Situacao AS EncargoSituacao,		    " & vbCrLf & _
              "		 Tp.Dolar AS Dolar, Tp.DataPTax AS DataPTax, Tp.ValorPTax AS ValorPTax,	    " & vbCrLf & _
              "		 Tp.ValorMoeda AS ValorMoeda, Tp.ValorOficial AS ValorOficial			    " & vbCrLf & _
              " FROM TabelaDePrecos Tp														    " & vbCrLf & _
              " INNER JOIN TabelaDePrecosXEncargos TpXE										    " & vbCrLf & _
              " ON TpXE.Produto_Id = Tp.Produto_Id											    " & vbCrLf & _
              " AND TpXE.Cliente_Id = Tp.Cliente_Id											    " & vbCrLf & _
              " AND TpXE.EndCliente_Id = Tp.EndCliente_Id									    " & vbCrLf & _
              " AND TpXE.Empresa_Id = Tp.Empresa_Id											    " & vbCrLf & _
              " AND TpXE.EndEmpresa_Id = Tp.EndEmpresa_Id									    " & vbCrLf & _
              " WHERE Tp.Cliente_Id = '" & Cliente & "'                                         " & vbCrLf & _
              " AND Tp.EndCliente_Id = " & EndCliente & "                                       "
        setObject(sql)
    End Sub

    Public Sub New(Produto As String)
        sql = " SELECT Tp.Produto_Id AS codProduto, Tp.Situacao ProdutoSituacao,				" & vbCrLf & _
              "		 Tp.Cliente_Id AS codCliente, Tp.EndCliente_Id AS codEndCliente,		    " & vbCrLf & _
              "		 Tp.Empresa_Id AS codEmpresa, Tp.EndEmpresa_Id AS codEndEmpresa,		    " & vbCrLf & _
              "		 TpXE.Encargo_Id AS codEncargo, TpXE.Situacao AS EncargoSituacao,		    " & vbCrLf & _
              "		 Tp.Dolar AS Dolar, Tp.DataPTax AS DataPTax, Tp.ValorPTax AS ValorPTax,	    " & vbCrLf & _
              "		 Tp.ValorMoeda AS ValorMoeda, Tp.ValorOficial AS ValorOficial			    " & vbCrLf & _
              " FROM TabelaDePrecos Tp														    " & vbCrLf & _
              " INNER JOIN TabelaDePrecosXEncargos TpXE										    " & vbCrLf & _
              " ON TpXE.Produto_Id = Tp.Produto_Id											    " & vbCrLf & _
              " AND TpXE.Cliente_Id = Tp.Cliente_Id											    " & vbCrLf & _
              " AND TpXE.EndCliente_Id = Tp.EndCliente_Id									    " & vbCrLf & _
              " AND TpXE.Empresa_Id = Tp.Empresa_Id											    " & vbCrLf & _
              " AND TpXE.EndEmpresa_Id = Tp.EndEmpresa_Id									    " & vbCrLf & _
              " WHERE Tp.Produto_Id = '" & Produto & "'                                         "
        setObject(sql)
    End Sub

    Private Sub setObject(sql)
        Try
            ds = objBanco.ConsultaDataSet(sql, "TabelaDePrecos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objClienteXProduto As New ClienteXProduto
                objClienteXProduto.codProduto = row("codProduto")
                objClienteXProduto.codCliente = row("codCliente")
                objClienteXProduto.codEndCliente = row("codEndCliente")
                objClienteXProduto.codEmpresa = row("codEmpresa")
                objClienteXProduto.codEndEmpresa = row("codEndEmpresa")
                objClienteXProduto.Dolar = CBool(row("Dolar"))
                objClienteXProduto.DataPTax = row("DolarPTax")
                objClienteXProduto.ValorPTax = row("ValorPTax")
                objClienteXProduto.ValorMoeda = row("ValorMoeda")
                objClienteXProduto.ValorOficial = row("ValorOficial")
                objClienteXProduto.Encargo = row("codEncargo")
                objClienteXProduto.SituacaoEncargo = CBool(row("EncargoSituacao"))
                Me.Add(objClienteXProduto)
            Next
        Catch ex As Exception
            MsgBox("Houve um erro ao consultar as alterações. Erro: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Methods"
    Public Sub salvar()
        For Each ClienteXProduto As ClienteXProduto In Me
            ClienteXProduto.salvar()
        Next
    End Sub
#End Region
End Class

<Serializable()> _
Public Class ClienteXProduto
    Implements IBaseEntity

#Region "Construtor"
    Dim ds As DataSet
    Dim sql As String
    Dim objBanco As New AcessaBanco

    Public Sub New()

    End Sub

    Public Sub New(Empresa As String, EndEmpresa As String, Cliente As String, EndCliente As String, Produto As String)
        Try
            sql = " SELECT Tp.Produto_Id AS codProduto, Tp.Situacao ProdutoSituacao,				" & vbCrLf & _
                  "		 Tp.Cliente_Id AS codCliente, Tp.EndCliente_Id AS codEndCliente,		    " & vbCrLf & _
                  "		 Tp.Empresa_Id AS codEmpresa, Tp.EndEmpresa_Id AS codEndEmpresa,		    " & vbCrLf & _
                  "		 TpXE.Encargo_Id AS codEncargo, TpXE.Situacao AS EncargoSituacao,		    " & vbCrLf & _
                  "		 Tp.Dolar AS Dolar, Tp.DataPTax AS DataPTax, Tp.ValorPTax AS ValorPTax,	    " & vbCrLf & _
                  "		 Tp.ValorMoeda AS ValorMoeda, Tp.ValorOficial AS ValorOficial			    " & vbCrLf & _
                  " FROM TabelaDePrecos Tp														    " & vbCrLf & _
                  " INNER JOIN TabelaDePrecosXEncargos TpXE										    " & vbCrLf & _
                  " ON TpXE.Produto_Id = Tp.Produto_Id											    " & vbCrLf & _
                  " AND TpXE.Cliente_Id = Tp.Cliente_Id											    " & vbCrLf & _
                  " AND TpXE.EndCliente_Id = Tp.EndCliente_Id									    " & vbCrLf & _
                  " AND TpXE.Empresa_Id = Tp.Empresa_Id											    " & vbCrLf & _
                  " AND TpXE.EndEmpresa_Id = Tp.EndEmpresa_Id									    " & vbCrLf & _
                  " WHERE Tp.Empresa_Id = '" & Empresa & "'                                         " & vbCrLf & _
                  " AND Tp.EndEmpresa_Id = " & EndEmpresa & "                                       " & vbCrLf & _
                  " AND Tp.Cliente_Id = '" & Cliente & "'                                           " & vbCrLf & _
                  " AND Tp.EndCliente_Id = " & EndCliente & "                                       " & vbCrLf & _
                  " AND Tp.CodProduto = '" & Produto & "'                                           "

            ds = objBanco.ConsultaDataSet(sql, "TabelaDePrecos")

            For Each row As DataRow In ds.Tables(0).Rows
                _codProduto = row("codProduto")
                _codCliente = row("codCliente")
                _codEndCliente = row("codEndCliente")
                _codEmpresa = row("codEmpresa")
                _codEndEmpresa = row("codEndEmpresa")
                _Dolar = CBool(row("Dolar"))
                _DataPTax = row("DolarPTax")
                _ValorPTax = row("ValorPTax")
                _ValorMoeda = row("ValorMoeda")
                _ValorOficial = row("ValorOficial")
                _Encargo = row("codEncargo")
                _SituacaoEncargo = CBool(row("EncargoSituacao"))
            Next
        Catch ex As Exception
            MsgBox("Houve um erro ao consultar as alterações. Erro: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Bacen As BACEN
    Private _codProduto As String
    Private _Produto As Produto
    Private _codCliente As String
    Private _codEndCliente As Integer
    Private _Cliente As Cliente
    Private _codEmpresa As String
    Private _codEndEmpresa As Integer
    Private _Empresa As Cliente
    Private _Encargo As String
    Private _SituacaoEncargo As Boolean
    Private _EncargosAtivos As ListEncargo
    Private _EncargosInativos As ListEncargo
    Private _Dolar As Boolean
    Private _ValorPTax As Decimal
    Private _DataPTax As DateTime
    Private _ValorMoeda As Decimal
    Private _ValorOficial As Decimal
    Private _Usuario As Usuario
#End Region

#Region "Property"
    Public WriteOnly Property IUD As String
        Set(value As String)
            _IUD = value
        End Set
    End Property
    Public ReadOnly Property Bacen As BACEN
        Get
            Return _Bacen
        End Get
    End Property
    Public WriteOnly Property codProduto As String
        Set(value As String)
            _codProduto = value
        End Set
    End Property
    Public ReadOnly Property Produto As Produto
        Get
            If _Produto Is Nothing Then
                _Produto = New Produto(_codProduto)
            End If
            Return _Produto
        End Get
    End Property
    Public WriteOnly Property codCliente As String
        Set(value As String)
            _codCliente = value
        End Set
    End Property
    Public WriteOnly Property codEndCliente As Integer
        Set(value As Integer)
            _codEndCliente = value
        End Set
    End Property
    Public ReadOnly Property Cliente As Cliente
        Get
            If _Cliente Is Nothing Then
                _Cliente = New Cliente(_codCliente, _codEndCliente)
            End If
            Return _Cliente
        End Get
    End Property
    Public WriteOnly Property codEmpresa As String
        Set(value As String)
            _codEmpresa = value
        End Set
    End Property
    Public WriteOnly Property codEndEmpresa As String
        Set(value As String)
            _codEndEmpresa = value
        End Set
    End Property
    Public ReadOnly Property Empresa As Cliente
        Get
            If _Empresa Is Nothing Then
                _Empresa = New Cliente(_codEmpresa, _codEmpresa)
            End If
            Return _Empresa
        End Get
    End Property
    Public WriteOnly Property Encargo As String
        Set(value As String)
            _Encargo = value
        End Set
    End Property
    Public WriteOnly Property SituacaoEncargo As Boolean
        Set(value As Boolean)
            If value Then
                _EncargosAtivos.Add(New Encargo(_Encargo))
            Else
                _EncargosInativos.Add(New Encargo(_Encargo))
            End If
        End Set
    End Property
    Public ReadOnly Property EncargosAtivos As ListEncargo
        Get
            Return _EncargosAtivos
        End Get
    End Property
    Public ReadOnly Property EncargosInativos As ListEncargo
        Get
            Return _EncargosInativos
        End Get
    End Property
    Public Property Dolar As Boolean
        Get
            Return _Dolar
        End Get
        Set(value As Boolean)
            _Dolar = value
        End Set
    End Property
    Public Property ValorPTax As Decimal
        Get
            Return _ValorPTax
        End Get
        Set(value As Decimal)
            _ValorPTax = value
        End Set
    End Property
    Public Property DataPTax As DateTime
        Get
            Return _DataPTax
        End Get
        Set(value As DateTime)
            _DataPTax = value
        End Set
    End Property
    Public Property ValorMoeda As Decimal
        Get
            Return _ValorMoeda
        End Get
        Set(value As Decimal)
            _ValorMoeda = value
        End Set
    End Property
    Public Property ValorOficial As Decimal
        Get
            Return _ValorOficial
        End Get
        Set(value As Decimal)
            _ValorOficial = value
        End Set
    End Property
    Public WriteOnly Property codUsuario As String
        Set(value As String)
            If _Usuario Is Nothing Then
                _Usuario = New Usuario(value)
            End If
        End Set
    End Property
    Private ReadOnly Property Data As DateTime
        Get
            Return Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub salvar()
        Try
            If Me._IUD = "I" Then
                sql = " INSERT INTO TabelaDePrecos " & vbCrLf & _
                      " VALUES ('" & Me.Empresa.Codigo & "'," & Me.Empresa.CodigoEndereco & ",'" & Me.Cliente.Codigo & "','" & Me.Produto.Codigo & "'," & Me.Dolar & vbCrLf & _
                      " ,'" & Me.DataPTax & "'," & Me.ValorPTax & "," & Me.ValorMoeda & "," & Me.ValorOficial & ",1,'" & Me._Usuario.Usuario_Id & "','" & Me.Data & "')"
                If Me.EncargosAtivos.Count > 0 Then
                    For Each encargo As Encargo In Me.EncargosAtivos
                        sql &= " INSERT INTO TabelaDePrecosXEncargos " & vbCrLf & _
                               " VALUES ('" & Me.Empresa.Codigo & "'," & Me.Empresa.CodigoEndereco & ",'" & Me.Cliente.Codigo & "'" & vbCrLf & _
                               " ,'" & Me.Produto.Codigo & "','" & encargo.Codigo & "',1,'" & Me._Usuario.Usuario_Id & "','" & Me.Data & "')"
                    Next
                End If
            ElseIf Me._IUD = "U" Then
                sql = " UPDATE TabelaDePrecos SET Dolar=" & Me.Dolar & ",DataPTax=" & Me.DataPTax & ",ValorPTax=" & Me.ValorPTax & "," & vbCrLf & _
                      " ValorMoeda=" & Me.ValorMoeda & ",ValorOficial=" & Me.ValorOficial & ",UsuarioAlteracao=" & Me._Usuario.Usuario_Id & ",UsuarioAlteracaoData=" & Me.Data & vbCrLf & _
                      " WHERE Empresa='" & Me.Empresa.Codigo & "'," & Me.Empresa.CodigoEndereco & ",'" & Me.Cliente.Codigo & "'," & Me.Cliente.CodigoEndereco & ",'" & Me.Produto.Codigo & "'"
            ElseIf Me._IUD = "D" Then
                sql &= " UPDATE TabelaDePrecos SET Situacao=0, UsuarioCancelamento='" & Me._Usuario.Usuario_Id & "',UsuarioCancelamentoData='" & Me.Data & "'" & vbCrLf & _
                       " WHERE Empresa_Id='" & Me.Empresa.Codigo & "',EndEmpresa_Id=" & Me.Empresa.CodigoEndereco & ",Cliente_Id='" & Me.Cliente.Codigo & "',EndCliente_Id=" & Me.Cliente.CodigoEndereco & ",Produto_Id='" & Me.Produto.Codigo & "'"
                If Me.EncargosInativos.Count > 0 Then
                    For Each encargo As Encargo In Me.EncargosInativos
                        sql &= " UPDATE TabelaDePrecosXEncargos SET Situacao=0,UsuarioCancelamento='" & Me._Usuario.Usuario_Id & "',UsuarioCancelamentoData='" & Me.Data & "'" & vbCrLf & _
                               " WHERE Empresa_Id='" & Me.Empresa.Codigo & "',EndEmpresa_Id=" & Me.Empresa.CodigoEndereco & ",Cliente_Id='" & Me.Cliente.Codigo & "',EndCliente=" & Me.Cliente.CodigoEndereco & vbCrLf & _
                               " ,Produto_Id = '" & Me.Produto.Codigo & "',Encargo_Id='" & encargo.Codigo & "'"
                    Next
                End If
            Else
                MsgBox("VALOR IUD INVALIDO.")
                Me._IUD = String.Empty
                Exit Sub
            End If

            objBanco.GravaBanco(sql)
        Catch ex As Exception
            MsgBox("Houve um erro ao consultar as alterações. Erro: " & ex.Message)
        End Try
    End Sub
#End Region
End Class