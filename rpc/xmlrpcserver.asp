<%

//--------------------------------------------------------------------
// XmlRpcServer
//--------------------------------------------------------------------

asplib.rpc.XmlRpcServer = function(apiKey) {
	this.__name = "asplib.rpc.XmlRpcServer";
	this.__init(apiKey);
};

asplib.rpc.XmlRpcServer.prototype.__init = function(apiKey) {

	this.apiKey = apiKey || asplib.rpc.config["apikey"] || "";
	this.defaultScope = global;

	// read and split querystring
	var temp = Request.QueryString.Item().split(",");
	var fn = temp[0];
	var l = (this.apiKey) ? temp.length-1 : temp.length;

	// client Md5 hash of request
	this.clientHash = temp[l] || null;
	this.serverHash = null;

	// decode args
	this.args = [];
	for (var i=1; i<l; i++) {
		try {
			this.args.push(decodeURIComponent(temp[i]));
		} catch(e) {
			// will crash with invalid URI-encoding
			this.args.push(null);
		}
	}

	this.methodName = fn;
	this.requestMethod = Request.ServerVariables("REQUEST_METHOD").Item().toUpperCase();
	var contentType = Request.ServerVariables("CONTENT_TYPE").Item().split(";");
	this.contentType = contentType[0].toLowerCase();
	this.contentArgs = [];

	// contentType args
	for (var i=1; i<contentType.length; i++) {
		var temp = contentType[i];
		var l = temp.indexOf("=");
		if (temp && l) {
			var name = temp.trim().substring(0,l-1).toLowerCase();
			var value = temp.trim().substring(l);
			this.contentArgs[name] = value;
		}
	}

	this.status = 0;
	this.statusText = "";
	this.contentCharset = this.contentArgs["charset"] || asplib.rpc.config["charset"];
	this.contentLength = Request.TotalBytes;

	switch (this.requestMethod) {
		case "GET" : {
			this.requestType = XMLRPC_REQUEST_GET;
			this.__validateHash();
			break;
		}
		case "POST" : {
			switch (this.contentType) {
				case "application/rpc+xml" : {
					this.requestType = XMLRPC_REQUEST_XML;
					this.__loadStream();
					this.__loadXmlRequest();
					break;
				}
				case "multipart/form-data" : {
					this.requestType = XMLRPC_REQUEST_MULTIPART;
					this.__loadStream();
					this.__loadMultipartRequest();
					break;
				}
				case "application/x-www-form-urlencoded" : {
					this.requestType = XMLRPC_REQUEST_POST;
					this.__loadFormRequest();
					break;
				}
				default : {
					this.requestType = XMLRPC_REQUEST_RAW;
					this.__loadStream();
					this.__loadRawRequest();
				}
			}
			break;
		}
	}
};

asplib.rpc.XmlRpcServer.prototype.__loadStream = function() {
	this.contentStream = new ActiveXObject("ADODB.Stream");
	this.contentStream.Open();
	this.contentStream.Type = 1; // 1=adTypeBinary
	try {
		this.contentStream.Write(Request.BinaryRead(Request.TotalBytes));
	//	this.contentStream.SaveToFile(Server.MapPath("/request.txt"), 2);
	} catch(e) {
	}
};

asplib.rpc.XmlRpcServer.prototype.__cleanupStream = function() {
	this.contentStream = null;
};

asplib.rpc.XmlRpcServer.prototype.__saveStream = function(path, position, length) {
	try {
		if (typeof position !== "undefined" && typeof length !== "undefined") {
			var stream = new ActiveXObject("ADODB.Stream");
			stream.Open();
			stream.Type = 1; // adTypeBinary
			this.contentStream.Position = position;
			this.contentStream.CopyTo(stream, length);
			stream.SaveToFile(path, 2); // adSaveCreateOverwrite
			stream = null;
		} else {
			this.contentStream.SaveToFile(path, 2); // adSaveCreateOverwrite
		}
	} catch(e) {
	}
};

asplib.rpc.XmlRpcServer.prototype.__readStreamText = function(charset) {
	try {
		this.contentStream.Position = 0;
		this.contentStream.Type = 2 // adTypeText
		this.contentStream.Charset = charset || this.contentCharset;
		return this.contentStream.ReadText();
	} catch(e) {
		return null;
	}
};

asplib.rpc.XmlRpcServer.prototype.__readStreamBinary = function() {
	try {
		this.contentStream.Position = 0;
		this.contentStream.Type = 1; // adTypeBinary
		return this.contentStream.Read();
	} catch(e) {
		return null;
	}
};

asplib.rpc.XmlRpcServer.prototype.__validateHash = function(data) {
	this.serverHash = Hash.hmacMd5(this.apiKey, this.methodName + this.args.join("") + (data||""));
	if (this.apiKey && this.serverHash !== this.clientHash) {
		this.status = XMLRPC_ERROR_INVALID_HASH;
		this.statusText = "Invalid Hash";
		if (asplib.rpc.config["debug"] === true) {
			Response.AddHeader("X-ApiKey", this.serverHash);
		}
		return false;
	} else {
		return true;
	}
};

asplib.rpc.XmlRpcServer.prototype.__loadXmlRequest = function() {
	this.request = null;
	this.requestXML = this.__readStreamText();
	if (this.__validateHash(this.requestXML)) {
		var xmlDOM = new ActiveXObject("Microsoft.XMLDOM");
		xmlDOM.async = false;
		xmlDOM.validateOnParse = false;
		xmlDOM.loadXML(this.requestXML)
		this.parseError = xmlDOM.parseError.errorCode;
		this.parseErrorText = xmlDOM.parseError.reason;
		var root = xmlDOM.documentElement;
		var request = new asplib.rpc.XmlRequest();
		if (root) {
			var obj = parseXMLNode(root);
			if (obj.method && !this.methodName) this.methodName = obj.method.toString();
			forEach(obj.param, function(param) {
				var name = param.name ? param.name.toString() : "";
				var type = param.type ? param.type.toString() : "";
				var value = param.value ? param.value.toString() : param.nodeValue;
				if (value == "") {
					value = param;
					delete value.name;
					delete value.type;
					delete value.value;
					delete value.size;
					delete value.index;
				}
				request.AddParam(name, value, type);
			});
			this.request = request;
		}
		if (this.parseError !== 0) {
			this.status = XMLRPC_ERROR_INVALID_XML;
			this.statusText = "XML Parse Error " + this.parseError;
		}
	}
};

asplib.rpc.XmlRpcServer.prototype.__loadFormRequest = function() {
	this.request = {};
	if (this.__validateHash()) {
		for (var e = new Enumerator(Request.Form); !e.atEnd(); e.moveNext()) {
			var name = e.item();
			var value = Request(name).Item();
			this.request[name] = value;
		}
	}
};

asplib.rpc.XmlRpcServer.prototype.__loadMultipartRequest = function() {
	this.request = { parts: [], files: [], items: [], names: [] };
	if (this.__validateHash()) {
		var vbCrLf = String.fromCharCode(13,10);
		var convertTypes = {
			"image/pjpeg" : "image/jpeg",
			"application/msword" : "application/vnd.ms-word",
			"application/msaccess" : "application/vnd.ms-access"
		};
		this.contentStream.Position = 0;
		this.contentStream.Type = 2; // adTypeText
		this.contentStream.Charset = "ISO-8859-1";
		var formData = this.contentStream.ReadText();
		var boundaryStart = formData.indexOf(vbCrLf)+2;
		var boundaryTag = formData.substring(0, boundaryStart-2);
		var boundaryMax = formData.indexOf(boundaryTag + "--");
		var boundaryEnd = 0;
		// check each message part
		while (boundaryEnd < boundaryMax) {
			boundaryEnd = formData.indexOf(boundaryTag, boundaryStart);
			var contentStart = formData.indexOf(vbCrLf + vbCrLf, boundaryStart)+4;
			var contentEnd = boundaryEnd-2;
			var contentLength = contentEnd - contentStart;
			var contentHeader = formData.substring(boundaryStart, contentStart-4).split(vbCrLf);
			var contentData = formData.substring(contentStart, contentEnd);
			// any content present?
			if (contentLength>0) {
				var contentType = null;
				var contentName = null;
				var contentFileName = null;
				// check headers
				for (var i=0; i<contentHeader.length; i++) {
					var headerName = contentHeader[i].substring(0, contentHeader[i].indexOf(":")).toLowerCase();
					var headerContent = contentHeader[i].substring(contentHeader[i].indexOf(":")+2).split(";");
					switch (headerName) {
						case "content-type" : {
							// content-type header (MIME type)
							contentType = headerContent[0];
							break;
						}
						case "content-disposition" : {
							// content-disposition header
							for (var j=0; j<headerContent.length; j++) {
								var headerParam = headerContent[j].split("=");
								switch (headerParam[0].trim()) {
									// name param?
									case "name" : {
										if (headerParam[1]) contentName = headerParam[1].replace(/\"/g, "").toLowerCase(); // "
										break;
									}
									// fileName?
									case "filename" : {
										if (headerParam[1]) contentFileName = headerParam[1].replace(/\"/g, ""); // "
										if (global.Utf8 && Utf8.validate(contentFileName)) contentFileName = Utf8.decode(contentFileName);
										break;
									}
								}
							}
							break;
						}
					}
				}
				var i = (contentFileName) ? contentFileName.lastIndexOf("\\") : -1;
				var path = (i !== -1) ? contentFileName.substring(0,i+1) : "";
				var fileName = (i !== -1) ? contentFileName.substring(i+1) : contentFileName;
				var part = {
					name: contentName,
					position: contentStart,
					length: contentLength,
					path: path,
					fileName: fileName,
					mimeType: convertTypes[contentType] || contentType,
					value: contentData,
					toString: function() { return this.value.toString(); }
				}
				if (fileName) {
					this.request.files.push(part);
				} else {
					this.request.items.push(part);
				}
			}
			// find next message part
			boundaryStart = formData.indexOf(vbCrLf, boundaryEnd)+2;
		}
	}
};

asplib.rpc.XmlRpcServer.prototype.__loadRawRequest = function() {
	this.request = null;
	if (this.__validateHash()) {
		this.request = new asplib.rpc.RawRequest(this.__readStreamBinary());
	}
};

asplib.rpc.XmlRpcServer.prototype.defaultHandler = function(args, request, rpc) {
	this.status = XMLRPC_ERROR_INVALID_METHOD;
	this.statusText = "Invalid Method";
	this.SendResponse();
};

asplib.rpc.XmlRpcServer.prototype.__sendXmlResponse = function(response) {
	this.responseXML = response.toString();
	Response.ContentType = "text/xml";
	Response.Charset = asplib.rpc.config["charset"];
	Response.CodePage = asplib.rpc.config["codepage"];
	Response.Write(response);
};

asplib.rpc.XmlRpcServer.prototype.__sendTextResponse = function(response) {
	this.responseText = response.toString();
	Response.ContentType = "text/plain";
	Response.Charset = asplib.rpc.config["charset"];
	Response.CodePage = asplib.rpc.config["codepage"];
	Response.Write(response);
};

asplib.rpc.XmlRpcServer.prototype.__sendRawResponse = function(response) {
	Response.ContentType = "application/octet-stream";
	Response.BinaryWrite(response.toBinary());
};

asplib.rpc.XmlRpcServer.prototype.SendResponse = function(response, type) {
	if (this.status !== 0) {
		Response.Status = "400 Bad Request (XmlRpcServer.status = 0x" + this.status.uint().toPaddedString(8, 16) + ")";
	} else {
		if (response != null) {
			switch (response.type) {
				case XMLRPC_RESPONSE_XML : {
					this.__sendXmlResponse(response);
					break;
				}
				case XMLRPC_RESPONSE_RAW : {
					this.__sendRawResponse(response);
					break;
				}
				default : {
					this.__sendTextResponse(response);
					break;
				}
			}
		}
	}
};

asplib.rpc.XmlRpcServer.prototype.HandleRequest = function() {
	var response = null;
	if (this.status === 0) {
		var fn = XMLRPC_METHOD_PREFIXES[this.requestType] + this.methodName;
		if (fn && this.defaultScope && typeof this.defaultScope[fn] === "function") {
			response = this.defaultScope[fn](this.args, this.request, this);
		} else {
			if (typeof this.defaultHandler === "function") {
				response = this.defaultHandler(this.args, this.request, this);
			}
		}
	}
	if (typeof response != null) {
		this.SendResponse(response);
	}
	this.__cleanupStream();
};

%>