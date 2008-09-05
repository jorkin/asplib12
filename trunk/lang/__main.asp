<%

	if (!this.asplib) asplib = {};
	if (!this.asplib.lang) asplib.lang = {
		items: [],
		browser: Request.ServerVariables("HTTP_ACCEPT_LANGUAGE").Item()
	};

// getProperty()
asplib.lang.getProperty = function(prop, lang) {
	if (asplib && asplib.lang) {
		var lang = lang || asplib.lang.current;
		if (lang && !asplib.lang[lang]) lang = asplib.lang.system;
		if (lang && asplib.lang[lang]) {
			var temp = asplib.lang[lang][prop];
			var inherits = asplib.lang[lang].inherits;
			if (!temp && inherits && asplib.lang[inherits]) {
				temp = asplib.lang[inherits][prop];
			}
			return temp;
		}
	}
};

%>