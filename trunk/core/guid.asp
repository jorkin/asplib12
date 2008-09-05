<%

/* guid.asp - by Thomas Kjoernes <thomas@ipv.no> */

// Guid()
function Guid() {
	this.__name = "Guid";
	var scriptletTypelib = new ActiveXObject("Scriptlet.Typelib");
	this.value = scriptletTypelib.Guid.substring(0,38);
	scriptletTypelib = null;
}

// toString()
Guid.prototype.toString = function() {
	return this.value.toString();
};

%>