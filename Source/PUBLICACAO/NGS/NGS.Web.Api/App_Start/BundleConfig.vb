Imports System.Web
Imports System.Web.Optimization

Public Module BundleConfig
    ' For more information on bundling, visit http://go.microsoft.com/fwlink/?LinkId=301862
    Public Sub RegisterBundles(ByVal bundles As BundleCollection)
        bundles.Add(New ScriptBundle("~/bundles/jquery").Include(
                   "~/Scripts/jquery-{version}.js"))

        ' Use the development version of Modernizr to develop with and learn from. Then, when you're
        ' ready for production, use the build tool at http://modernizr.com to pick only the tests you need.
        bundles.Add(New ScriptBundle("~/bundles/modernizr").Include(
                    "~/Scripts/modernizr-*"))

        bundles.Add(New ScriptBundle("~/bundles/bootstrap").Include(
                    "~/Scripts/bootstrap.js",
                    "~/Scripts/respond.js"))

        bundles.Add(New StyleBundle("~/Content/css").Include(
                    "~/Content/bootstrap.css",
                    "~/Content/site.css"))

        bundles.Add(New StyleBundle("~/admin-lte/css").Include(
                    "~/admin-lte/css/adminlte.css",
                    "~/admin-lte/css/adminlte.min.css",
                    "~/admin-lte/css/adminlte.rtl.css",
                    "~/admin-lte/css/adminlte.rtl.min.css"))

        bundles.Add(New ScriptBundle("~/admin-lte/js").Include(
                    "~/admin-lte/js/chart.min.js",
                    "~/admin-lte/js/jquery.min.js",
                    "~/admin-lte/js/adminlte.min.js"))

    End Sub
End Module