<%

//--------------------------------------------------------------------
// XmlResponse
//--------------------------------------------------------------------

asplib.rpc.XmlResponse = function(xml, data, status, statusText) {
	trace("asplib.rpc.XmlResponse");
	this.__name = "asplib.rpc.XmlResponse";
	this.type = XMLRPC_RESPONSE_XML;
	this.status = status || 0;
	this.statusText = statusText || "";
	this.data = data || [];
	this.xml = xml || "";
};

asplib.rpc.XmlResponse.prototype.toString = function() {
	return this.toXMLString();
};

asplib.rpc.XmlResponse.prototype.toXMLString = function() {
	if (this.xml) {
		return this.xml;
	} else {
		var xml = [];
		xml.push(Xml.prolog("1.0", "UTF-8"));
		xml.push("<response>");
		xml.push("<status" + Xml.attr("text", this.statusText) + ">" + this.status + "</status>");
		if (this.session) {
			xml.push("<session>" + this.session + "</session>");
		}
		if (this.data) {
			for (var i=0; i<this.data.length; i++) {
				xml.push(this.data[i].toXMLString());
			}
		}
		xml.push("</response>");
		return xml.join("\n");
	}
};

asplib.rpc.XmlResponse.prototype.AddData = function(name, value, type, size, index) {
	this.data.push(new asplib.rpc.XmlData(name, value, type, size, index));
};

asplib.rpc.XmlResponse.prototype.GetData = function(name) {
	if (this.data) {
		for (var i=0; i<this.data.length; i++) {
			var data = this.data[i];
			if (data.name && data.name.toString() === name) return data.valueOf();
		}
	} else {
		return null;
	}
};

%>