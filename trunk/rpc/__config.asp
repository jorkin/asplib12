<%

	asplib.rpc.config["debug"] = true;
	asplib.rpc.config["codepage"] = XMLRPC_DEFAULT_CODEPAGE;
	asplib.rpc.config["charset"] = XMLRPC_DEFAULT_CHARSET;
	asplib.rpc.config["apiurl"] = new URL("/api/svc.asp").toString();
	asplib.rpc.config["apikey"] = "";
	asplib.rpc.config["request-param-value"] = "textnode,attrnode";
	asplib.rpc.config["response-data-value"] = "textnode";

%>