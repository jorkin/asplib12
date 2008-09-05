<%

/* document.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.html) asplib.html = {};

	var W3_XMLNS_URL = "http://www.w3.org/1999/xhtml";
	var W3_VALIDATOR_URL = "http://validator.w3.org/check?uri=referer";
	var W3_VALIDATOR_HTML_IMG = "http://www.w3.org/Icons/valid-html401";
	var W3_VALIDATOR_XHTML_IMG = "http://www.w3.org/Icons/valid-xhtml10";
	var W3_XHTML_CONTENT_TYPE = "text/html";

	var HTML401Strict = {
		xml: false,
		text: "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01//EN\"\n \"http://www.w3.org/TR/html4/strict.dtd\">\n",
		validator: {
			alt: "Valid HTML 4.01 Strict",
			src: W3_VALIDATOR_HTML_IMG,
			width: 88,
			height: 31
		}
	};

	var HTML401Transitional = {
		xml: false,
		text: "<!DOCTYPE html PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\n \"http://www.w3.org/TR/html4/loose.dtd\">\n",
		validator: {
			alt: "Valid HTML 4.01 Transitional",
			src: W3_VALIDATOR_HTML_IMG,
			width: 88,
			height: 31
		}
	};

	var HTML401Frameset = {
		xml: false,
		text: "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Frameset//EN\"\n \"http://www.w3.org/TR/html4/frameset.dtd\">\n",
		validator: {
			alt: "Valid HTML 4.01 Frameset",
			src: W3_VALIDATOR_XHTML_IMG,
			width: 88,
			height: 31
		}
	};

	var XHTML10Strict = {
		xml: true,
		text: "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\"\n \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">\n",
		validator: {
			alt: "Valid XHTML 1.0 Strict",
			src: W3_VALIDATOR_XHTML_IMG,
			width: 88,
			height: 31
		},
		contentType: W3_XHTML_CONTENT_TYPE
	};

	var XHTML10Transitional = {
		xml: true,
		text: "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\"\n \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\n",
		validator: {
			alt: "Valid XHTML 1.0 Transitional",
			src: W3_VALIDATOR_XHTML_IMG,
			width: 88,
			height: 31
		},
		contentType: W3_XHTML_CONTENT_TYPE

	};

	var XHTML10Frameset = {
		xml: true,
		text: "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Frameset//EN\"\n \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd\">\n",
		validator: {
			alt: "Valid XHTML 1.0 Frameset",
			src: "http://www.w3.org/Icons/valid-xhtml10",
			width: 88,
			height: 31
		},
		contentType: W3_XHTML_CONTENT_TYPE

	};

	var PART_HTML_BEGIN = 1;
	var PART_HTML_END = 2;
	var PART_BODY_BEGIN = 3;
	var PART_BODY_END = 4;
	var PART_HEAD = 5;
	var PART_BODY = 6;

//--------------------------------------------------------------------
// asplib.html.Document()
//--------------------------------------------------------------------

asplib.html.Document = function(doctype, lang, dir, title, base) {
	this.__name = "asplib.html.Document";
	trace("asplib.html.Document");
	this.buffer = true;
	this.opened = false;
	this.dirty = false;
	this.doctype = doctype || HTML401Transitional;
	this.html = {
		lang: lang || "",
		dir: dir || "",
		head: {
			id: "",
			meta: [],
			link: [],
			script: [],
			style: [],
			title: title || "",
			base: base || "",
			innerHTML: []
		},
		body: {
			id: "",
			classNames: [],
			innerHTML: []
		}
	};
	this.__xml = function() {
		return (this.doctype.xml) ? " /" : "";
	};
	this.__attr = function(name, value) {
		return (name && value != null && value !== "") ? (" " + name + "=\"" + Server.HTMLEncode(value) + "\"") : "";
	};

};

asplib.html.Document.prototype.__render = function(part) {
	var html = [];
	if (!part || part===PART_HTML_BEGIN) {
		html.push(this.doctype.text);
		html.push("<html" + (this.doctype.xml ? (this.__attr("xmlns",  W3_XMLNS_URL) + this.__attr("xml:lang", this.html.lang)) : "") + this.__attr("lang", this.html.lang) + this.__attr("dir", this.html.dir) + ">\n");
	}
	if (!part || part===PART_HEAD) {
		html.push("<head" + this.__attr("id", this.html.head.id) + ">\n");
		for (var i in this.html.head.meta) {
			var meta = this.html.head.meta[i];
			html.push(
				"<meta" +
				this.__attr("name", meta.name) +
				this.__attr("http-equiv", meta.httpEquiv) +
				this.__attr("content", meta.content) +
				this.__attr("scheme", meta.scheme) +
				this.__xml() +
				">\n"
			);
		}
		html.push("<title>" + this.html.head.title + "</title>\n");
		if (this.html.head.base) {
			html.push("<base" + this.__attr("href", this.html.head.base) + this.__xml());
		}
		for (var i in this.html.head.link) {
			var link = this.html.head.link[i];
			if (link.cond) html.push("<!--[if " + link.cond + "]>\n");
			html.push(
				"<link" +
				this.__attr("href", link.href) +
				this.__attr("rel", link.rel) +
				this.__attr("type", link.type) +
				this.__xml() +
				">\n"
			);
			if (link.cond) html.push("<![endif]-->\n");
		}
		for (var i in this.html.head.script) {
			var script = this.html.head.script[i];
			if (script.cond) html.push("<!--[if " + script.cond + "]>\n");
			html.push(
				"<script" +
				this.__attr("src", script.src) +
				this.__attr("type", script.type) +
				">" +
				((script.text) ? "\n" + script.text : "") +
				"</script>\n"
			);
			if (script.cond) html.push("<![endif]-->\n");
		}
		if (this.html.head.style.length) {
			html.push("<style type=\"text/css\">\n" + this.html.head.style.join("\n") + "\n</style>\n");
		}
		html.push(this.html.head.innerHTML.join(""));
		html.push("</head>\n");
	}
	if (!part || part===PART_BODY_BEGIN) {
		html.push("<body" + this.__attr("id", this.html.body.id) + this.__attr("class", this.html.body.classNames.join(" ")) + ">\n");
	}
	if (!part || part===PART_BODY) {
		html.push(this.html.body.innerHTML.join(""));
	}
	if (!part || part===PART_BODY_END) {
		html.push("</body>\n");
	}
	if (!part || part===PART_HTML_END) {
		html.push("</html>\n");
	}
	return html.join("");
};

asplib.html.Document.prototype.Open = function() {
	trace("asplib.html.Document.Open");
	this.buffer = false;
	Response.Write(this.__render(PART_HTML_BEGIN));
};

asplib.html.Document.prototype.Close = function() {
	trace("asplib.html.Document.Close");
	Response.Write(this.__render(PART_BODY_END));
	Response.Write(this.__render(PART_HTML_END));
};

asplib.html.Document.prototype.SetTitle = function(title) {
	trace("asplib.html.Document.SetTitle");
	this.html.head.title = title || "";
};

asplib.html.Document.prototype.AddMeta = function(name, content, httpEquiv, scheme) {
	trace("asplib.html.Document.AddMeta");
	var meta = { name: name, content: content, httpEquiv: httpEquiv, scheme: scheme };
	if ((name || httpEquiv) && content) {
		this.dirty = true;
		this.html.head.meta.push(meta);
	}
};

asplib.html.Document.prototype.AddLink = function(href, rel, type, cond) {
	trace("asplib.html.Document.AddLink");
	var link = { href: href, rel: rel, type: type, cond: cond };
	if (href && rel) {
		this.dirty = true;
		this.html.head.link.push(link);
	}
};

asplib.html.Document.prototype.AddScript = function(src, text, type, cond) {
	trace("asplib.html.Document.AddScript");
	var script = { src: src, text: text, type: type ? type : "text/javascript", cond: cond };
	if (src || text) this.dirty = true;
	if (src) {
		this.html.head.script[src] = script;
	} else {
		if (text) {
			this.html.head.script.push(script);
		}
	}
};

asplib.html.Document.prototype.GetScript = function(src, text, type) {
	trace("asplib.html.Document.GetScript");
	return "<script" + this.__attr("src", src) + this.__attr("type", ((type) ? type : "text/javascript")) + ">" + text + "</script>\n";
};

asplib.html.Document.prototype.PutScript = function(src, text, type) {
	trace("asplib.html.Document.PutScript");
	this.Put(this.GetScript(src, text, type));
};

asplib.html.Document.prototype.AddStyleSheet = function(href, cond) {
	trace("asplib.html.Document.AddStyleSheet");
	this.AddLink(href, "stylesheet", "text/css", cond);
};

asplib.html.Document.prototype.AddFavIcon = function(href, cond) {
	trace("asplib.html.Document.AddFavIcon");
	this.AddLink(href, "shortcut icon", "image/x-icon", cond);
};

asplib.html.Document.prototype.AddStyle = function(text) {
	trace("asplib.html.Document.AddStyle");
	this.dirty = true;
	this.html.head.style.innerHTML.push(text);
};

asplib.html.Document.prototype.Put = function(text) {
	trace("asplib.html.Document.Put");
	if (this.buffer) {
		this.html.body.innerHTML.push(text);
		this.dirty = true;
	} else {
		if (!this.opened) {
			this.opened = true;
			Response.Write(this.__render(PART_HEAD));
			Response.Write(this.__render(PART_BODY_BEGIN));
		}
		Response.Write(text)
	}
};

asplib.html.Document.prototype.Write = function(text) {
	this.Put(text);
};

asplib.html.Document.prototype.Get = function() {
	trace("asplib.html.Document.Get");
	return this.__render();
};

asplib.html.Document.prototype.toString = function() {
	return this.__render();
};

asplib.html.Document.prototype.Render = function() {
	trace("asplib.html.Document.Render");
	if (this.buffer) {
		Response.Write(this.__render());
	}
};

asplib.html.Document.prototype.ValidatorSnippet = function(fragment) {
    var html = [];
    var validator = this.doctype.validator;
    if (!fragment) html.push("<p>");
    html.push(
    	"<a href=\"" + W3_VALIDATOR_URL + "\"><img" +
    	(!this.doctype.xml ? this.__attr("border", "0") : this.__attr("style", "border:none")) +
    	this.__attr("src", validator.src) +
    	this.__attr("alt", validator.alt) +
    	this.__attr("width", validator.width) +
    	this.__attr("height", validator.height) +
    	this.__xml() +
	    "></a>"
    );
    if (!fragment) html.push("</p>\n");
    return html.join("");
};

asplib.html.Document.prototype.ValidatorURL = function() {
    return W3_VALIDATOR_URL;
};

%>