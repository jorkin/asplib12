<%

/* http.asp - by Thomas Kjoernes <thomas@ipv.no> */

	// WinHttp Options
	var WINHTTP_OPTION_UserAgentString = 0
	var WINHTTP_OPTION_URL = 1
	var WINHTTP_OPTION_URLCodePage = 2
	var WINHTTP_OPTION_EscapePercentInURL = 3
	var WINHTTP_OPTION_SslErrorIgnoreFlags = 4
	var WINHTTP_OPTION_SelectCertificate = 5
	var WINHTTP_OPTION_EnableRedirects = 6
	var WINHTTP_OPTION_UrlEscapeDisable = 7
	var WINHTTP_OPTION_UrlEscapeDisableQuery = 8
	var WINHTTP_OPTION_SecureProtocols = 9
	var WINHTTP_OPTION_EnableTracing = 10

// XMLHttpRequest()
function XMLHttpRequest() {
	var msxmls = ["WinHttp.WinHttpRequest.5.1", "WinHttpRequest.5", "MSXML2.ServerXMLHTTP"];
	for (var i=0; i<msxmls.length; i++) {
		try {
			return new ActiveXObject(msxmls[i]);
		} catch (e) { }
	}
	return null;
}

%>