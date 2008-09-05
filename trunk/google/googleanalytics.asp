<%

/* googleanalytics.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.google) asplib.google = {};

// GoogleAnalytics()
asplib.google.GoogleAnalytics = function(accountId) {
	this.__name = "asplib.google.GoogleAnalytics";
	this.loaded = false;
	this.accountId = accountId;
};

asplib.google.GoogleAnalytics.prototype.toString = function() {
	return this.Render();
};

asplib.google.GoogleAnalytics.prototype.URL = function() {
	var	secure = Request.ServerVariables("SERVER_PORT_SECURE").Item();
	var url = ((secure === "1") ? "https://ssl" : "http://www") + ".google-analytics.com/ga.js";
	return url;
};

asplib.google.GoogleAnalytics.prototype.Script = function() {
	return "<script type=\"text/javascript\" src=\"" + this.URL() + "\"></script>\n";
};

asplib.google.GoogleAnalytics.prototype.Render = function() {
	var html = [];
	if (this.accountId) {
		if (!this.loaded) {
			html.push(this.Script());
			this.loaded = true;
		}
		html.push("<script type=\"text/javascript\">\n");
		html.push("var pageTracker = _gat._getTracker(\"" + this.accountId + "\");\n");
		html.push("pageTracker._initData();\n");
		html.push("pageTracker._trackPageview();\n");
		html.push("</script>\n");
	}
	return html.join("");
};

%>
