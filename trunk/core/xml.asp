<%

/* xml.asp - by Thomas Kjoernes <thomas@ipv.no> */

// parseXMLNode()
parseXMLNode = function(node) {
	var obj = {};
	var element = node.firstChild;
	while (element) {
		if (element.nodeType === 1) {
			var name = element.nodeName;
			var sub = parseXMLNode(element)
			sub.nodeValue = "";
			sub.xml = element.xml;
			sub.toString = function() { return this.nodeValue; };
			sub.toXMLString = function() { return this.xml; }
			// get attributes
			if (element.attributes) {
				for (var i=0; i<element.attributes.length; i++) {
					var attribute = element.attributes[i];
					sub[attribute.nodeName] = attribute.nodeValue;
				}
			}
			// get nodeValue
			if (element.firstChild) {
				var nodeType = element.firstChild.nodeType;
				if (nodeType === 3 || nodeType === 4) {
					sub.nodeValue = element.firstChild.nodeValue;
				}
			}
			// node already exists?
			if (obj[name]) {
				// need to create array?
				if (!obj[name].length) {
					var temp = obj[name];
					obj[name] = [];
					obj[name].push(temp);
				}
				// append object to array
				obj[name].push(sub);
			} else {
				// create object
				obj[name] = sub;
			}
		}
		element = element.nextSibling;
	}
	return obj;
}

// parseXMLDocument()
function parseXMLDocument(xmldoc) {
	var obj = {};
	if (xmldoc && xmldoc.documentElement) {
		obj.__root = xmldoc.documentElement.nodeName;
		obj[obj.__root] = parseXMLNode(xmldoc.documentElement);
	}
	return obj;
}

// parseXMLString()
function parseXMLString(xml) {
	var obj = {};
	var xmlDOM = new ActiveXObject("Microsoft.XMLDOM");
	xmlDOM.async = false;
	xmlDOM.validateOnParse = false;
	var success = xmlDOM.loadXML(xml);
	if (success) {
		obj = parseXMLDocument(xmlDOM);
	}
	xmlDOM = null;
	return obj;
}

// parseXML()
function parseXML(url) {
	var obj = {};
	var xmlDOM = new ActiveXObject("Microsoft.XMLDOM");
	xmlDOM.async = false;
	xmlDOM.validateOnParse = false;
	var success = xmlDOM.load(url);
	if (!success) {
		var xmlHttp = new ActiveXObject("MSXML2.ServerXMLHTTP");
		xmlHttp.open("GET", new URL(url), false);
		xmlHttp.send();
		if (xmlHttp.status === 200) {
			success = xmlDOM.loadXML(xmlHttp.responseText);
			if (!success) {
				obj.__error = "Parse Error 0x" + xmlDOM.parseError.errorCode.uint().toString(16) + " on line " + xmlDOM.parseError.line + " (column " + xmlDOM.parseError.linepos  + ")";
			}
		} else {
			obj.__error = "HTTP Error: " + xmlHttp.status + " " + xmlHttp.statusText;
		}
		xmlHttp = null;
	}
	if (success) {
		obj = parseXMLDocument(xmlDOM);
	}
	xmlDOM = null;
	return obj;
}

// XML()
function XML() {
	for (var i=0; i<arguments.length; i++) {
		var arg = arguments[i];
		if (typeof arg === "function") arg = arg();
		if (typeof arg !== "string") arg = arg.toString();
		var obj = (arg.indexOf("<?xml") === 0) ? parseXMLString(arg) : parseXML(arg);
		for (var j in obj) this[j] = obj[j];
	}
	this.__name = "XML";
}

// Xml()
var Xml = {

	prolog: function(version, encoding, standalone) {
		var version = version || "1.0";
		var encoding = encoding || Response.Charset || "ISO-8859-1";
		return "<?xml version=\"" + version + "\" encoding=\"" + encoding + "\"" + (standalone ? " standalone=\"yes\"" : "") + " ?>";
	},

	text: function(text) {
		return (text != null) ? text.toString().replace(/\&/g, "&amp;").replace(/\</g, "&lt;").replace(/\>/g, "&gt;") : "";
	},

	attr: function(name, value) {
		return (name && value != null) ? (" " + name + "=\"" + this.text(value).replace(/\x22/g, "&quot;") + "\"") : "";
	},

	cdata: function(text) {
		return (text) ? "<![CDATA[" + text.toString().replace("]>>", "]&gt;&gt;") + "]>>" : "";
	},

	comment: function(text) {
		return (text) ? "<!--" + text.toString().replace("--", "\x2D\x2D") + "-->" : "";
	},

	parse: function(url) {
		return parseXML(url);
	},

	toXMLString: function(obj) {
		var xml = [];
		function serialize(obj) {
			var xml = [];
			var type = typeof obj;
			switch (type) {
				case "object" : {
					if (obj === null) {
						xml.push("<"+i+" />\n");
					} else if (typeof obj.getTime === "function") {
						xml.push("<"+i+Xml.attr("type", "date")+(obj !== "" ? ">"+obj.toUTCString()+"</"+i : " /")+">\n");
					} else if (typeof obj.join === "function") {
						for (var j=0; j<obj.length; j++) {
							xml.push(serialize(obj[j]));
						}
					} else {
						var sub = Xml.toXMLString(obj);
						xml.push("<"+i+Xml.attr("type", type)+(sub !== "" ? ">\n"+sub+"</"+i : " /")+">\n");
					}
					break;
				}
				case "string" :
					xml.push("<"+i+Xml.attr("type", type)+(obj !== "" ? ">"+Xml.text(obj)+"</"+i : " /")+">\n");
					break;
				case "date" :
				case "boolean" :
				case "number" : {
					xml.push("<"+i+Xml.attr("type", type)+">"+obj+"</"+i+">\n");
					break;
				}
			}
			return xml.join("");
		}
		for (var i in obj) {
			xml.push(serialize(obj[i]));
		}
		return xml.join("");
	},

	arrayToXML: function(obj) {
		return this.toXMLString(obj);
	},

	objectToXML: function(obj) {
		return this.toXMLString(obj);
	}

};

%>