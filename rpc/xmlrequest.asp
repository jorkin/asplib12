<%

//--------------------------------------------------------------------
// XmlRequest
//--------------------------------------------------------------------

asplib.rpc.XmlRequest = function(xml, params) {
	trace("asplib.rpc.XmlRequest");
	this.__name = "asplib.rpc.XmlRequest";
	this.type = XMLRPC_REQUEST_XML;
	this.params = params || [];
	this.xml = xml || "";
};

asplib.rpc.XmlRequest.prototype.toString = function() {
	return this.toXMLString();
};

asplib.rpc.XmlRequest.prototype.toXMLString = function() {
	if (this.xml) {
		return this.xml;
	} else {
		var xml = [];
		xml.push(Xml.prolog("1.0", "UTF-8"));
		xml.push("<request>");
		if (this.session) {
			xml.push("<session>" + this.session + "</session>");
		}
		if (this.params) {
			for (var i=0; i<this.params.length; i++) {
				xml.push(this.params[i].toXMLString());
			}
		}
		xml.push("</request>");
		return xml.join("\n");
	}
};

asplib.rpc.XmlRequest.prototype.AddParam = function(name, value, type, size, index) {
	this.params.push(new asplib.rpc.XmlParam(name, value, type, size, index));
};

asplib.rpc.XmlRequest.prototype.GetParam = function(name) {
	if (this.params) {
		for (var i=0; i<this.params.length; i++) {
			var param = this.params[i];
			if (param.name && param.name.toString() === name) return param.valueOf();
		}
	} else {
		return null;
	}
};

%>