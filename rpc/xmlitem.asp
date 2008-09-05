<%

//--------------------------------------------------------------------
// XmlItem
//--------------------------------------------------------------------

asplib.rpc.XmlItem = function(value, type, index, size) {
	trace("asplib.rpc.XmlItem");
	this.__name = "asplib.rpc.XmlItem";
	this.value = value;
	this.type = type;
	if (index) this.index = index;
	if (size) this.size = size
	if (typeof value === "string") {
		switch (type) {
			case "int" :
			case "decimal" :
			case "float" : this.value = parseNumber(value, null); break
			case "bool" : this.value = parseBoolean(value, null); break
			case "date" : this.value = parseDate(value, null); break
			case "string" : this.value = value; break;
			case "array" :
			case "object" : this.xml = value; break;
			default : this.value = null;
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
				case "array" : this.xml = Xml.arrayToXML(value); break;
				case "object" : this.xml = Xml.objectToXML(value); break;
				default : this.value = null;
			}
		}
	}
};

asplib.rpc.XmlItem.prototype.valueOf = function() {
	return this.value;
};

asplib.rpc.XmlItem.prototype.toString = function() {
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

asplib.rpc.XmlItem.prototype.toXMLString = function() {
	return "<item" + Xml.attr("index", this.index) + Xml.attr("type", this.type) + Xml.attr("size", this.size) + (this.value == null ? " />" : ">" + this.toString() + "</item>");
};

%>