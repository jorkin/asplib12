<%

//--------------------------------------------------------------------
// XmlParam
//--------------------------------------------------------------------

asplib.rpc.XmlParam = function(name, value, type, size, index) {
	trace("asplib.rpc.XmlParam");
	this.__name = "asplib.rpc.XmlParam";
	this.name = name || "";
	this.type = type || (typeof value !== "undefined" ? "string" : "");
	this.value = value;
	this.xml = "";
	if (size) this.size = size
	if (index) this.index = index;
	if (typeof value === "string") {
		switch (type) {
			case "int" :
			case "decimal" :
			case "float" : this.value = parseNumber(value, null); break
			case "bool" : this.value = parseBoolean(value, null); break
			case "date" : this.value = parseDate(value, null); break
			case "string" : this.value = parseString(value, null); break;
			case "array" :
			case "object" : {
				this.xml = value;
				break;
			}
			case "string" :
		}
	} else {
		if (value != null) {
			switch (type) {
				case "int" :
				case "decimal" :
				case "float" : this.value = parseNumber(value, null); break
				case "bool" : this.value = parseBoolean(value, null); break
				case "date" : this.value = parseDate(value, null); break;
				case "string" : this.value = parseString(value, null); break;
				case "array" :
				case "object" : {
					if (typeof value.toXMLString === "function") {
						for (var i in value) {
							if (typeof value[i] === "object") {
								var temp =  toArray(value[i]);
								for (var j=0; j<temp.length; j++) {
									this.xml += temp[j].toXMLString();
								}
							}
						}
					} else if (typeof value.join === "function") {
						this.xml = Xml.arrayToXML(value);
					} else {
						this.xml = Xml.objectToXML(value);
					}
					break;
				}
			}
		}
	}
};

asplib.rpc.XmlParam.prototype.valueOf = function() {
	if (this.xml) {
		return this.xml;
	} else {
		return this.value;
	}
};

asplib.rpc.XmlParam.prototype.toString = function() {
	if (this.xml) {
		return this.xml;
	} else {
		if (this.value == null) {
			return "";
		} else if (typeof this.value.toUTCString === "function") {
			return this.value.toUTCString();
		} else {
			return Xml.text(this.value.toString());
		}
	}
};

asplib.rpc.XmlParam.prototype.toXMLString = function() {
	return "<param" + Xml.attr("name", this.name) + Xml.attr("type", this.type) + Xml.attr("size", this.size) + Xml.attr("index", this.index) + (this.value == null ? " />" : ">" + this.toString() + "</param>");
};

%>