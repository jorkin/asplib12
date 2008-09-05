<%

/* template.asp - by Thomas Kjoernes <thomas@ipv.no> */

	global = this;

	if (!this.asplib) asplib = {};
	if (!this.asplib.html) asplib.html = {};

//--------------------------------------------------------------------
// asplib.html.Template()
//--------------------------------------------------------------------

asplib.html.Template = function(path, fields, defaultHandler) {
	this.__name = "asplib.html.Template";
	trace("asplib.html.Template");
	this.charset = Response.CharSet || "UTF-8";
	this.html = [];
	this.fields = [];
	this.defaultHandler = function(inst, fn) { return "<var>" + fn + "</var>"; };
	this.removeTabs = false;
	this.removeDoubleLinebreaks = false;
	if (typeof fields === "object" && fields != null) this.fields = fields;
	if (typeof defaultHandler !== "undefined") this.defaultHandler = defaultHandler;
	this.LoadFromFile(path);
};

asplib.html.Template.prototype.toString = function() {
	var that = this;
	var text = this.html.join("").replace(/{\s*(.+?)\s*}/g, function(str, fn) {
		if (typeof that.fields[fn] === "string") return that.fields[fn];
		if (typeof that.fields[fn] === "function") return that.fields[fn](that, str);
		if (typeof global[fn] === "string") return global[fn];
		if (typeof global[fn] === "function") return global[fn](that, str);
		if (typeof that.defaultHandler === "string") return that.defaultHandler;
		if (typeof that.defaultHandler === "function") return that.defaultHandler(that, str);
	});
	if (this.removeTabs) {
		text = text.replace(/\t/g, "");
	}
	if (this.removeDoubleLinebreaks) {
		text = text.replace(/\r\n/g, "\n");
		text = text.replace(/\n\n/g, "\n");
	}
	return text;
};

asplib.html.Template.prototype.Get = function() {
	trace("asplib.html.Template.Get");
	return this.toString();
};

asplib.html.Template.prototype.Put = function(text) {
	trace("asplib.html.Template.Put");
	this.html.push(text);
};

asplib.html.Template.prototype.Set = function(text) {
	trace("asplib.html.Template.Put");
	this.html = [];
	if (text && typeof text.toString === "function") {
		this.html.push(text.toString());
	}
	return this.toString();
};

asplib.html.Template.prototype.LoadFromFile = function(path, charset) {
	trace("asplib.html.Template.LoadFromFile");
	if (path) {
		var stream = new ActiveXObject("ADODB.Stream");
		stream.Open();
		stream.Type = 2; // adTypeBinary=1
		stream.CharSet = charset || this.charset;
		try {
			stream.LoadFromFile(path);
			this.html.push(stream.ReadText());
		} catch(e) { }
		stream.Close();
		stream = null;
	}
};

%>