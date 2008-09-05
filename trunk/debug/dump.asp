<%

/* dump.asp - by Thomas Kjoernes <thomas@ipv.no> */

function dump(obj, name, open, sub, depth, buttons) {
	var html = [];
	var obj = obj;
	var xml = (asplib && asplib.hyperweb && asplib.hyperweb.document) ? asplib.hyperweb.document.__xml() : "";
	if (!depth) depth = 1;
	if (depth>20) return html;
	if (typeof obj !== "undefined") {
		var item = [];
		var canenum = true;
		try {
			for (var i in obj) {}
		} catch(e) {
			canenum = false;
		}
		if (canenum) {
			for (var i in obj) {
				var prop = obj[i];
				var details = "";
				var propstr = "";
				try {
					propstr = prop.toString();
					details = propstr.replace(/\n/g, "\\n");
					details = details.replace(/\t/g, "\\t");
					details = details.replace(/</g, "&lt;");
					details = details.replace(/>/g, "&gt;");
					details = (details.length>50) ? (details.substring(0,47) + "...") : details;
				} catch(e) { }
				var prefix = (i.indexOf("__")===0 ? "private " : "");
				var image = "/asplib1.2/debug/images/" + (i.indexOf("__")===0 ? "private_" : "") + "property.gif";
				var title = (i.indexOf("__")===0 ? "private " : "") + typeof prop;
				switch (typeof prop) {
					case "function" : {
						var isClass = (propstr.indexOf(".__name") !== -1);
						var p1 = details.indexOf("(");
						var p2 = details.indexOf(")");
						var params = (p1<p2) ? details.substring(p1+1, p2) : "";
						details = "(<span style=\"color:gray\">" + params + "</span>)";
						if (isClass) {
							image = "/asplib1.2/debug/images/baseclass.gif";
							title = prefix + "class";
						} else {
							image = "/asplib1.2/debug/images/" + (i.indexOf("__")===0 ? "private_" : "") + "method.gif";
							title = prefix + "method";
						}
						if (i.substring(0,2) === "on") {
							image = "/asplib1.2/debug/images/" + "eventhandler.gif";
							title = prefix + "eventhandler";
						}
						temp = null;
						break;
					}
					case "object" : {
						var isDate = (prop && typeof prop.getDate === "function");
						image = "/asplib1.2/debug/images/object.gif";
						if (prop != null) {
						//	details = "<span style=\"color:green\">\"" + details + "\"</span>";
							try {
								if (prop.__name) {
									title = prefix + "object instance";
									image = "/asplib1.2/debug/images/instance.gif";
									details = "";
								}
								if (typeof prop.join === "function") {
									title = prefix + "array";
									image = "/asplib1.2/debug/images/object.gif";
									details = "";
								}
								if (typeof prop.join === "function" && !prop.length) {
									title = prefix + "collection";
									image = "/asplib1.2/debug/images/object.gif";
									details = "";
								}
								if (typeof prop.getDate === "function") {
									title = prefix + "date";
									image = "/asplib1.2/debug/images/property.gif";
									details = " = <span style=\"color:green\">\"" + details + "\"</span>";
								} else {
									if (details) details = " = <span style=\"color:green\">\"" + details + "\"</span>";
								}
							} catch(e) {
								details = "<span style=\"color:blue\">" + "ActiveXObject" + "</span>";
								image = "/asplib1.2/debug/images/activex.gif";
							}
							if (details.indexOf("object Object") !== -1) details = "";
						} else {
						//	title = prefix + "date";
							image = "/asplib1.2/debug/images/property.gif";
							details = " = <span style=\"color:blue\">" + "null" + "</span>";
						}
						break;
					}
					case "string" : {
						details = Server.HTMLEncode(details);
						details = " = <span style=\"color:green\">\"" + details + "\"</span>";
						break;
					}
					case "number" : {
						details = " = <span style=\"color:maroon\">" + details + "</span>";
						break;
					}
					case "date" : {
						details = " = <span style=\"color:blue\">" + details + "</span>";
						break;
					}
					case "boolean" : {
						details = " = <span style=\"color:blue\">" + details + "</span>";
						break;
					}
					case "undefined" : {
						details = " = <span style=\"color:blue\">" + details + "</span>";
						break;
					}
				}
				if (i !== "__name") {
					item.push("<li" + (open ? "" :" class=\"collapsed\"") + ">");
					item.push("<img src=\"" + image + "\" alt=\"\"" + xml + ">");
					item.push("<a onclick=\"return false\" title=\"" + title + "\">");
					item.push("" + i + "");
					item.push("</a>");
					if (details) item.push(details);
					if (typeof prop === "object" && prop !== null) {
						try {
							if (prop && prop.__name) item.push("<b>" + prop.__name + "</b>");
						} catch(e) { }
						if (i.charAt(0) !== "_") {
							item.push(dump(prop, null, open, true, depth+1));
							if (typeof prop.prototype === "object") {
								item.push(dump(prop.prototype, null, open, true, depth+1));
							}
						}
					}
					item.push("</li>\n");
				}
			}
		}
	/*	if (!propLength && obj.length) {
			item.push("<li>");
			item.push("<img src=\"/asplib1.2/debug/images/property.gif\" alt=\"\"" + xml + ">");
			item.push("<a onclick=\"return false\" title=\"" + "number" + "\">");
			item.push("" + "length" + "");
			item.push("</a>");
			item.push(" = <span style=\"color:maroon\">" + obj.length + "</span>");
		}*/
		if (item.length) {
			if (!sub) {
				html.push("<link rel=\"stylesheet\" type=\"text/css\" href=\"/asplib1.2/debug/styles/dump.css\"" + xml + ">\n");
				html.push("<script type=\"text/javascript\" src=\"/asplib1.2/debug/scripts/dom.js\"></script>\n");
				html.push("<script type=\"text/javascript\" src=\"/asplib1.2/debug/scripts/tree.js\"></script>\n");
				html.push("<ul class=\"tree" + (buttons === false ? "" : " buttons") + "\">\n");
			} else {
				html.push("<ul>\n");
			}
			if (name) {
				html.push("<li>");
				html.push("<img src=\"/asplib1.2/debug/images/namespace.gif\" alt=\"\"" + xml + ">");
				html.push("<a onclick=\"return false\">" + name + "</a>\n");
				html.push("<ul>\n");
			}
			html = html.concat(item);
			if (name) {
				html.push("</ul>\n");
				html.push("</li>\n");
			}
			html.push("</ul>\n");
		}
	}
	return html.join("");
}

function dumpForm() {
	var html = [];
	for (var e = new Enumerator(Request.Form); !e.atEnd(); e.moveNext()) {
		var name = e.item();
		var value = Request.Form(name).Item();
		var color = "green";
		html.push(name + "=<span style=\"color:" + color + "\">" + value + "</span><br>\n");
	}
	return html.join("");
}

%>