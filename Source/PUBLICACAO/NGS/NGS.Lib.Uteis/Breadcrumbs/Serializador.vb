Imports System.IO
Imports System.Xml.Serialization
Imports System.Web

Public Module Serializador

    Public Sub Serializar()
        Dim obj As New Breadcrumbs()
        Dim lstNodes As New List(Of BreadcrumbsNode)

        Dim node1 As New BreadcrumbsNode()
        node1.Title = "Contas À Pagar"
        node1.Description = "Contas À Pagar"
        node1.Url = "~/ContasAPagar.aspx"
        node1.CssClass = "link"
        lstNodes.Add(node1)

        Dim node2 As New BreadcrumbsNode()
        node2.Title = "Contas À Receber"
        node2.Description = "Contas À Receber"
        node2.Url = "~/ContasAReceber.aspx"
        node2.CssClass = "link"
        lstNodes.Add(node2)

        Dim start As New BreadcrumbsNode()
        start.Title = "Início"
        start.Description = "Início"
        start.Url = "~/Index.aspx"
        start.CssClass = "linkHome"
        start.Nodes = New List(Of BreadcrumbsNode)() From {
            New BreadcrumbsNode() With { _
              .Title = "Financeiro", _
              .Description = "Financeiro", _
              .Url = "~/Financeiro.aspx", _
              .CssClass = "link", _
              .Nodes = lstNodes _
             } _
        }

        obj.Nodes = New List(Of BreadcrumbsNode)
        obj.Nodes.Add(start)

        'criar o arquivo xml - breadcrumbs.xml - na raiz da aplicação
        Using writer As TextWriter = New StreamWriter(HttpContext.Current.Server.MapPath("~/breadcrumbs.xml"))
            'definir o type e o elemento raiz do xml (breadcrumbs)
            Dim xml As XmlSerializer = New XmlSerializer(obj.GetType())
            'serializar o list para o textwriter e salvar os dados no xml
            xml.Serialize(writer, obj)
        End Using

    End Sub

    Public Function Deserializar() As Breadcrumbs
        Dim obj As New Breadcrumbs()
        Dim xml As XmlSerializer = New XmlSerializer(obj.GetType())
        Using fs As New FileStream(HttpContext.Current.Server.MapPath("~/breadcrumbs.xml"), FileMode.Open)
            obj = CType(xml.Deserialize(fs), Breadcrumbs)
            fs.Close()
        End Using
        Return obj
    End Function

End Module
