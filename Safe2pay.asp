json =  "{                                                                      "&_
        "            ""IsSandbox"": false,                                      "&_
        "            ""IpAddress"": """&ip&""",                                 "&_
        "            ""Application"": """&nome_site&""",                        "&_
        "            ""Vendor"": """&nome_site&""",                             "&_
        "            ""ShouldUseAntiFraud"": true,                              "&_
        "            ""PaymentMethod"": ""2"",                                  "&_
        "            ""Reference"": """&pedido&""",                             "&_
        "            ""Customer"":                                              "&_
        "                {                                                      "&_
        "                    ""Name"": """&nome&""",                            "&_
        "                    ""Identity"": """&cpf&""",                         "&_
        "                    ""Email"": """&Email&""",                          "&_
        "                    ""Phone"": """&phone&""",                          "&_
        "                    ""Address"": {""ZipCode"": """&cep&""",            "&_
        "                    ""Street"": """&logradouro&""",                    "&_
        "                    ""Number"": """&numero&""",                        "&_
        "                    ""Complement"": """&complemento&""",               "&_
        "                    ""District"": """&bairro&""",                      "&_
        "                    ""CityName"": """&cidade&""",                      "&_
        "                    ""StateInitials"": """&estado&""",                 "&_
        "                    ""CountryName"": ""Brasil""                        "&_
        "                }                                                      "&_
        "        },                                                             "&_
        "        ""Products"":                                                  "&_
        "            [{                                                         "&_
        "                ""Code"": ""1"",                                       "&_
        "                ""Description"": """&Nome_produto&""",                 "&_
        "                ""UnitPrice"": "&valor&",                              "&_
        "                ""Quantity"": "&quantidade&"                           "&_
        "            }],                                                        "&_
        "        ""PaymentObject"":                                             "&_
        "            {                                                          "&_
        "            ""Holder"": """&Holder&""",                                "&_
        "            ""CardNumber"": """&CardNumber&""",                        "&_
        "            ""ExpirationDate"": """&ExpirationDate&""",                "&_
        "            ""SecurityCode"": """&SecurityCode&""",                    "&_
        "            ""InstallmentQuantity"": """&parcelas&"""                  "&_
        "            }                                                          "&_
        "        }                                                              "
Set objhttpPagBol = Server.CreateObject ("Msxml2.ServerXMLHTTP.6.0")
objhttpPagBol.Open "POST", "https://payment.safe2pay.com.br/v2/Payment", False
objhttpPagBol.SetRequestHeader "Content-Type", "application/json"
objhttpPagBol.SetRequestHeader "x-api-key", afiliacao
objhttpPagBol.Send json
strResponseStatusw3 = objhttpPagBol.Status & " " & objhttpPagBol.StatusText
strResponseTextw3 = objhttpPagBol.ResponseText