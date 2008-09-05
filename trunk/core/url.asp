<%

/* url.asp - by Thomas Kjoernes <thomas@ipv.no> */

// splitURI()
function splitURI(uri) {
	if (!uri) return [];
	var i = (uri.charAt(0) === "/") ? 1 : 0;
	var j = (uri.charAt(uri.length-1) === "/") ? uri.length-1 : uri.length;
	return uri.substring(i,j).split("/");
}

// completeURI()
function completeURI(uri) {
	if (!uri) return "";
	return (uri.match(/\.|\/$/g, "")) ? uri : uri + "/";
}

// filterURI()
function filterURI(uri) {
	if (!uri) return "";
	return uri.toString().toAscii().replace(/\s/g, "-").replace("_", "-").replace(/(-){2,}/g, "-").replace(/[^a-zA-Z0-9\-/]/g, "");
}

// filterURIComponent()
function filterURIComponent(uri) {
	if (!uri) return "";
	return uri.toString().toAscii().replace(/\s/g, "-").replace("_", "-").replace(/(-){2,}/g, "-").replace(/[^a-zA-Z0-9\-/]/g, "");
}

// URL()
function URL(url) {
	this.__name = "URL";
	this.protocol = Request.ServerVariables("SERVER_PORT_SECURE").Item() === "1" ? "https:" : "http:";
	this.hostname = Request.ServerVariables("SERVER_NAME").Item();
	this.port = Request.ServerVariables("SERVER_PORT").Item();
	this.pathname = Request.ServerVariables("PATH_INFO").Item();
	this.search = Request.QueryString.Item();
	this.hash = "";
	this.Parse(url);
	this.__rebuild();
}
// __rebuild()
URL.prototype.__rebuild = function(relative) {
	this.host = (this.port=="80" || this.port=="443") ? this.hostname : this.hostname + ":" + this.port;
	this.path = this.pathname + (this.search ? "?" + this.search : "") + (this.hash ? "#" + this.hash : "");
	this.href = this.protocol + "//" + this.host + this.path;
	return relative ? this.path : this.href;
};
// toString()
URL.prototype.toString = function(relative) {
	return this.__rebuild(relative);
};
// Split()
URL.prototype.Split = function(relative) {
	return splitURI(this.__rebuild(relative));
};
// Filter()
URL.prototype.Filter = function(relative) {
	return filterURI(this.__rebuild(relative));
};
// Complete()
URL.prototype.Complete = function(relative) {
	return completeURI(this.__rebuild(relative));
};
// Clear()
URL.prototype.Clear = function(relative) {
	this.search = "";
	return this.__rebuild(relative);
};
// Parse()
URL.prototype.Parse = function(url, relative) {
	var url = (url) ? url + "" : "";
	var i = url.indexOf("://");
	if (i<0) i=-3;
	var j = url.indexOf(":", i+3);
	var k = url.indexOf("/", i+3);
	if (i<0) k = 0;
	var l = url.indexOf("?");
	var m = url.indexOf("#");
	if (m<0) m = url.length;
	if (l<0) l = m;
	if (k<0) k = l;
	if (j<0) j = k;
	var protocol = (i>0) ? url.substring(0, i+1) : "";
	var hostname = (i>0) ? url.substring(i+3, j) : "";
	var port = (j<k) ? url.substring(j+1, k) : "";
	var pathname = url.substring(k, l);
	var search = (l<m) ? url.substring(l+1, m) : "";
	var hash = url.substring(m+1);
	var temp = this.pathname.substring(0, this.pathname.lastIndexOf("/"));
	this.filename = this.pathname.substring(this.pathname.lastIndexOf("/")+1);
	this.extension = this.filename.substring(this.filename.lastIndexOf("."));
	if (pathname && pathname.indexOf("/")!==0) pathname = temp + "/" + pathname;
	if (pathname && pathname.charAt(0) !== "/") pathname = "/" + pathname;
	if (pathname) this.pathname = pathname;
	if (protocol) this.protocol = protocol;
	if (hostname) this.hostname = hostname;
	if (port) this.port = port;
	if (search || pathname) this.search = search;
	if (hash || search) this.hash = hash;
	return this.__rebuild(relative);
};
// Get()
URL.prototype.Get = function(name) {
	var temp = [];
	var name = name + "=";
	if (name) {
		var search = (this.search) ? this.search.split("&") : [];
		for (var i=0; i<search.length; i++) {
			if (search[i].indexOf(name)===0) {
				temp.push(search[i].substring(name.length));
			}
		}
	}
	return temp.join(",");
};
// Set()
URL.prototype.Set = function(name, value, relative) {
	var name = (name) ? name+"=" : "";
	var value = encodeURIComponent(value != null ? value+"" : "");
	var search = (this.search) ? this.search.split("&") : [];
	for (var i=0; i<search.length; i++) {
		if (search[i].indexOf(name)===0) break;
	}
	if (search[i]) {
		search[i] = name+value;
	} else {
		search.push(name+value);
	}
	this.search = search.join("&");
	return this.__rebuild(relative);
};
// Add()
URL.prototype.Add = function(name, value, relative) {
	this.search += (this.search ? "&" : "") + name + "=" + encodeURIComponent(value != null ? value+"" : "");
	return this.__rebuild(relative);
};
// Remove()
URL.prototype.Remove = function(name, relative) {
	var temp = [];
	var name = name + "=";
	var search = (this.search) ? this.search.split("&") : [];
	for (var i=0; i<search.length; i++) {
		if (search[i].indexOf(name)!==0) temp.push(search[i]);
	}
	this.search = temp.join("&");
	return this.__rebuild(relative);
};
// NoCache()
URL.prototype.NoCache = function(relative) {
	return this.Set("nocache", new Date().valueOf(), relative);
};

// Extended Methods - returns a new URL instance

// SetEx()
URL.prototype.SetEx = function(name, value, relative) {
	return new URL(this.Set(name, value, relative));
};
// AddEx()
URL.prototype.AddEx = function(name, value, relative) {
	return new URL(this.Add(name, value, relative));
};
// RemoveEx()
URL.prototype.RemoveEx = function(name, relative) {
	return new URL(this.Remove(name, relative));
};
// NoCacheEx()
URL.prototype.NoCacheEx = function(relative) {
	return new URL(this.NoCache(relative));
};
// SplitEx()
URL.prototype.SplitEx = function(relative) {
	return new URL(this.Split(relative));
};
// FilterEx()
URL.prototype.FilterEx = function(relative) {
	return new URL(this.Filter(relative));
};
// CompleteEx()
URL.prototype.CompleteEx = function(relative) {
	return new URL(this.Complete(relative));
};
// ClearEx()
URL.prototype.ClearEx = function(relative) {
	return new URL(this.Clear(relative));
};

%>