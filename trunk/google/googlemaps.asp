<%

/* googlemaps.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.google) asplib.google = {};

// GoogleMaps()
asplib.google.GoogleMaps = function(apiKey, width, height, lat, lng, version) {
	this.__name = "asplib.google.GoogleMaps";
	this.loaded = false;
	this.apiKey = apiKey;
	this.width = width || 320;
	this.height = height || 240;
	this.lat = lat || 0;
	this.lng = lng || 0;
	this.version = version || 2;
};

asplib.google.GoogleMaps.prototype.toString = function() {
	return this.Render();
};

asplib.google.GoogleMaps.prototype.URL = function() {
	var	secure = Request.ServerVariables("SERVER_PORT_SECURE").Item();
	var url = "http://maps.google.com/maps?file=api&v=" + this.version + "&key=" + this.apiKey;
	return url;
};

asplib.google.GoogleMaps.prototype.Script = function() {
	return "<script type=\"text/javascript\" src=\"" + this.URL() + "\"></script>\n";
};

asplib.google.GoogleMaps.prototype.Render = function() {
	var html = [];
	if (this.accountId) {
		if (!this.loaded) {
			html.push(this.Script());
			this.loaded = true;
		}
		html.push("<script type=\"text/javascript\">\n");
		html.push("</script>\n");
	}
	return html.join("");
};

%>
