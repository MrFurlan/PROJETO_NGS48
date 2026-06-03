@ModelType Models.LoginViewModel
@Code
    Layout = Nothing
End Code

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>Dashboard</title>
    <meta content="width=device-width, initial-scale=1, maximum-scale=10, user-scalable=yes" name="viewport" />

    <link href="//netdna.bootstrapcdn.com/bootstrap/3.0.0/css/bootstrap.min.css" rel="stylesheet" id="bootstrap-css">
    <script src="//code.jquery.com/jquery-1.11.1.min.js"></script>
    <script src="//netdna.bootstrapcdn.com/bootstrap/3.0.0/js/bootstrap.min.js"></script>

    @Styles.Render("~/Content/css")
    @Scripts.Render("~/bundles/modernizr")
    <link id="favicon" rel="shortcut icon" href="@Url.Content("~/admin-lte/assets/img/favicon.png")" />
    <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,600,700,300italic,400italic,600italic" />

    <style>
        html, body {
            height: 100%;
            margin: 0;
            padding: 0;
            overflow: hidden;
        }

        *, *::before, *::after {
            box-sizing: border-box;
        }

        #background-carousel {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: -1;
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            transition: background-image 1s ease-in-out;
        }

        body {
            font-family: 'Source Sans Pro', sans-serif;
        }

        .form-signin {
            max-width: 350px;
            padding: 15px;
            margin: 0 auto;
        }

            .form-signin .form-signin-heading, .form-signin .checkbox {
                margin-bottom: 10px;
            }

            .form-signin .checkbox {
                font-weight: normal;
            }

            .form-signin .form-control {
                position: relative;
                font-size: 16px;
                height: auto;
                padding: 10px;
                box-sizing: border-box;
            }

            .form-signin input[type="text"] {
                margin-bottom: -1px;
                border-bottom-left-radius: 0;
                border-bottom-right-radius: 0;
            }

            .form-signin input[type="password"] {
                margin-bottom: 10px;
                border-top-left-radius: 0;
                border-top-right-radius: 0;
            }

        .account-wall {
            margin-top: 20px;
            padding: 40px 20px 20px 20px;
            background-color: rgba(255,255,255,0.95);
            box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.4);
            border-radius: 10px;
            max-width: 350px;
            width: 100%;
        }

            .account-wall input.form-control {
                height: 45px;
                padding: 10px 12px;
                font-size: 16px;
            }

        select.form-control {
            padding-top: 10px;
            padding-bottom: 10px;
            height: auto;
            min-height: 42px;
        }

        .login-title {
            color: #555;
            font-size: 18px;
            font-weight: 400;
            display: block;
        }

        footer.app-footer {
            color: white;
            padding: 15px;
            margin-top: auto;
        }

        .text-muted {
            color: #ccc !important;
        }

        .login-span {
            display: block;
            margin-top: 5px;
        }

        .wrapper {
            min-height: 100vh;
            display: flex;
            flex-direction: column;
        }

        .content {
            flex: 1;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }

        img, input, button, select {
            max-width: 100%;
        }
    </style>
</head>
<body class="hold-transition login-page">

    <div id="background-carousel"></div>

    <div class="wrapper">
        <div class="content">
            <div class="account-wall">
                @Code
                    Dim returnUrlValue As String = If(String.IsNullOrEmpty(ViewBag.ReturnUrl), "", ViewBag.ReturnUrl)
                    Using Html.BeginForm("Login", "Conta", New With {.ReturnUrl = returnUrlValue}, FormMethod.Post)
                        @Html.AntiForgeryToken()
                End Code

                <p class="login-box-msg"><strong>LOGIN</strong></p>

                @Html.ValidationSummary("", New With {.class = "validationerrors"})
                @Html.HiddenFor(Function(model) model.ReturnUrl)

                <div class="form-group has-feedback">
                    @Html.TextBoxFor(Function(m) m.UsuarioId, New With {.class = "form-control", .placeholder = "Login", .autofocus = "autofocus", .required = "required"})
                    <span class="glyphicon glyphicon-envelope form-control-feedback login-span"></span>
                </div>
                <div class="form-group has-feedback">
                    <input type="password" name="password" id="password" class="form-control" placeholder="Senha" required>
                    <span class="glyphicon glyphicon-lock form-control-feedback login-span"></span>
                </div>

                @If Not Model Is Nothing AndAlso Model.BancosDeDados IsNot Nothing AndAlso Model.BancosDeDados.Count > 1 Then
                    @<div class="form-group has-feedback">
                        @Html.DropDownListFor(Function(m) m.BancoId, CType(Model.BancosDeDados, List(Of SelectListItem)), "-- Selecione Banco de Dados--", New With {.class = "form-control"})
                    </div>
                End If

                <div class="row">
                    <div class="col-md-12">
                        <button class="btn btn-lg btn-primary btn-block" type="submit">Entrar</button>
                    </div>
                </div>

                @If ViewBag.Erro IsNot Nothing Then
                    @<div class="text-danger">@ViewBag.Erro</div>
                End If

                @Code
                    End Using
                End Code
            </div>
        </div>

        <footer class="app-footer d-flex justify-content-end align-items-center px-3 w-100">
            <div class="text-end">
                <strong>
                    Copyright &copy; @DateTime.Now.Year.ToString()
                    <a href="https://ngssolucoes.com.br/" class="text-decoration-none text-light">NGS</a>.
                </strong>
            </div>
        </footer>
    </div>

    <div id="progress" style="display: none;">
        <div class="divProgressBackground"></div>
        <div class="divProgressLoading">
            <i class="fa fa-refresh fa-2x fa-spin"></i><span class="fa-2x">&nbsp;Carregando</span>
        </div>
    </div>

    @Scripts.Render("~/bundles/jquery")
    @Scripts.Render("~/bundles/bootstrap")

    <script type="text/javascript">
        $(document).ready(function () {
            if ($("#UsuarioId").val() == "") {
                $("#UsuarioId").focus();
            } else if ($("#password").val() == "") {
                $("#password").focus();
            }

            $("form").on('submit', function () {
                $("#progress").show();
            });

            if (typeof (Storage) !== "undefined") {
                sessionStorage.removeItem("divSelecionada");
            }
        });

        // Carrossel de imagens de fundo
        const images = [
            '@Url.Content("~/Images/dashbord_1.png")',
            '@Url.Content("~/Images/dashbord_2.png")',
            '@Url.Content("~/Images/dashbord_3.png")',
            '@Url.Content("~/Images/dashbord_4.png")'
        ];

        let currentImage = 0;
        const carousel = document.getElementById("background-carousel");

        function changeBackground() {
            carousel.style.backgroundImage = `url('${images[currentImage]}')`;
            currentImage = (currentImage + 1) % images.length;
        }

        changeBackground();
        setInterval(changeBackground, 5000);
    </script>
</body>
</html>
