<%

//--------------------------------------------------------------------
// XmlRpcClient
//--------------------------------------------------------------------

asplib.rpc.XmlRpcClient = function(apiURL, apiKey) {
	this.__name = "XmlRpcClient";
	this.__init(apiURL, apiKey);
};

asplib.rpc.XmlRpcClient.prototype.__init = function(apiURL, apiKey) {
	this.apiURL = apiURL || asplib.rpc.config["apiurl"] || "";
	this.apiKey = apiKey || asplib.rpc.config["apikey"] || "";
};

asplib.rpc.XmlRpcClient.prototype.SendRequest = function(fn, args, request) {
	// encode args
	var args = toArray(args)||[];
	var temp = [];
	for (var i=0; i<args.length; i++) {
		temp.push(encodeURIComponent(args[i]));
	}
	// create hash and url
	var hash = (this.apiKey) ? Hash.hmacMd5(this.apiKey, fn + temp.join("") + request.toString()) : "";
	var url = this.apiURL + (fn ? "?" + fn : "") + (fn && temp.length>0 ? "," + temp.join(",") : "") + (fn && hash ? "," + hash : "");
	var xmlHttp = new XMLHttpRequest();
	var requestType = (request) ? request.type : 0;
	this.requestURL = url;
	// check request type
	switch (requestType) {
		case XMLRPC_REQUEST_XML : {
			xmlHttp.open("POST", url, false);
			xmlHttp.setRequestHeader("Content-Type", "application/rpc+xml; charset=UTF-8");
			xmlHttp.send(request);
			break;
		}
		case XMLRPC_REQUEST_RAW : {
			xmlHttp.open("POST", url, false);
			xmlHttp.setRequestHeader("Content-Type", "application/octet-stream");
			xmlHttp.send(request.toString());
			break;
		}
		default : {
			xmlHttp.open("GET", url, false);
			xmlHttp.send(null);
		}
	}
	var response = null;
	if (xmlHttp.status === 200) {
		this.status = 0;
		switch (requestType) {
			// return XML response
			case XMLRPC_REQUEST_XML : {
				response = new XmlResponse();
				var xml = parseXMLDocument(xmlHttp.responseXML);
				if (xml && xml.response) {
					if (xml.response.status) response.status = parseInt(xml.response.status.nodeValue);
					if (xml.response.data) {
						var data = toArray(xml.response.data);
						for (var i=0; i<data.length; i++) {
							var name = data[i].name ? data[i].name.toString() : "";
							var type = data[i].type ? data[i].type.toString() : "";
							var value = data[i].value ? data[i].value.toString() : data[i].nodeValue;
							if (value == "") {
								var value = data[i];
								delete value.name;
								delete value.type;
								delete value.value;
								delete value.size;
								delete value.index;
							}
							response.AddData(name, value, type);
						}
					}
				}
				response.xml = xmlHttp.responseText;
				break;
			}
			// return RAW response
			case XMLRPC_REQUEST_RAW : {
				response = new RawResponse(xmlHttp.responseText);
				break;
			}
			default : {
				var xml = parseXMLDocument(xmlHttp.responseXML);
				response = {
					__httpStatus: xmlHttp.status,
					__httpStatusText: xmlHttp.statusText,
					xml: xmlHttp.responseText, 
					type: requestType
				};
				if (xml && xml.__root) {
					response[xml.__root] = xml;
				}
			}
		}
	} else {
		this.status = XMLRPC_ERROR;
		this.statusText = xmlHttp.responseText;
	}
	xmlHttp = null;
	return response;
};

%>
