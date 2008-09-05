<%

/* flash.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.html) asplib.html = {};

//--------------------------------------------------------------------
// asplib.html.Flash()
//--------------------------------------------------------------------

asplib.html.Flash = function(url, width, height, bgcolor, params, name, version) {
	trace("asplib.html.Flash");
	this.__name = "asplib.html.Flash";
	this.Init(url, width, height);
};

asplib.html.Flash.prototype.Init = function(url, width, height, bgcolor, params, name, version) {
	trace("asplib.html.Flash.Init");
	this.id = "flash_" + new Date().valueOf();
	this.width = width || "";
	this.height = height || "";
	this.name = name || "";
	this.bgcolor = bgcolor || "#FFF";
	this.version = version || "7";
	this.url = (url) ? url.toString() : "";
	this.classNames = [];
	this.document = {
		__xml: function() { return (this.xml) ? " /" : "";	},
		__attr: function(name, value) { return (name && value != null && value !== "") ? (" " + name + "=\"" + Server.HTMLEncode(value) + "\"") : ""; }
	};
	this.xml = false;
	this.params = [];
	this.AddParam("movie", url);
	if (params) {
		for (var i in params) this.AddParam(i, params[i]);
	}
};

asplib.html.Flash.prototype.__render = function() {
	var html = [];
	html.push(
		"<object" +
		this.document.__attr("classid", "clsid:D27CDB6E-AE6D-11cf-96B8-444553540000") +
		this.document.__attr("codebase", "http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,40,0") +
		this.document.__attr("id", this.id) +
		this.document.__attr("name", this.name) +
		this.document.__attr("width", this.width) +
		this.document.__attr("height", this.height) +
		">\n"
	);
	for (var i in this.params) {
		html.push(
		"<param" +
		this.document.__attr("name", this.params[i].name) +
		this.document.__attr("value", this.params[i].value) +
		this.document.__xml() +
		">\n"
		);
	}
	html.push(
		"<embed" +
		this.document.__attr("src", this.url) +
		this.document.__attr("name", this.name) +
		this.document.__attr("width", this.width) +
		this.document.__attr("height", this.height) +
		this.document.__attr("bgcolor", this.bgcolor) +
		this.document.__attr("quality", this.quality) +
		">\n"
	);
	html.push(
		"</object>\n"
	);
	return html.join("");
};

asplib.html.Flash.prototype.AddParam = function(name, value) {
	trace("asplib.html.Flash.AddParam");
	if (typeof name !== "undefined") {
		this.params.push({ name: name, value: value });
	}
};

asplib.html.Flash.prototype.Render = function() {
	trace("asplib.html.Flash.Render");
	if (this.document && this.document.Put) {
		this.document.Put(this.__render());
	}
};

asplib.html.Flash.prototype.toString = function() {
	return this.__render();
};

%>