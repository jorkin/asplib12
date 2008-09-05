<%

	var global = this;

	if (!this.asplib) asplib = {};
	if (!this.asplib.rpc) asplib.rpc = {
		config: [],
		cache: []
	};

	var XMLRPC_DEFAULT_CHARSET = "UTF-8";
	var XMLRPC_DEFAULT_CODEPAGE = 65001;

	var XMLRPC_ERROR = 0x80000000;
	var XMLRPC_ERROR_INVALID_HASH = 0x80000001;
	var XMLRPC_ERROR_INVALID_METHOD = 0x80000002;
	var XMLRPC_ERROR_INVALID_XML = 0x80000003;

	var	XMLRPC_REQUEST_GET = 0;
	var	XMLRPC_REQUEST_XML = 1;
	var XMLRPC_REQUEST_RAW = 2;
	var XMLRPC_REQUEST_MULTIPART = 3;
	var XMLRPC_REQUEST_POST = 4;

	var XMLRPC_RESPONSE_XML = 1;
	var XMLRPC_RESPONSE_RAW = 2;

	var XMLRPC_CONTENT_TYPES = {
		1: "application/rpc+xml",
		2: "application/octet-stream",
		3: "multipart/form-data",
		4: "application/x-www-form-urlencoded"
	};

	var XMLRPC_METHOD_PREFIXES = {
		0: "Get_",
		1: "Xml_",
		2: "Raw_",
		3: "Multipart_",
		4: "Post_"
	}

%>