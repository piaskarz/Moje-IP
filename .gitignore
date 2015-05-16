Option Explicit
Dim http : Set http = CreateObject( "MSXML2.ServerXmlHttp" )
http.Open "GET", "http://ip-api.com/line/?fields=country,city,isp,query", False
http.Send
Wscript.Echo http.responseText
Set http = Nothing
