<%

// toShortTime()
Date.prototype.toShortTime = function(lang) {
	var format = asplib.lang.getProperty("shorttime", lang);
	return this.format(format, lang);
};

// toShortDate()
Date.prototype.toShortDate = function(lang) {
	var format = asplib.lang.getProperty("shortdate", lang);
	return this.format(format, lang);
};

// toShortDateTime()
Date.prototype.toShortDateTime = function(lang) {
	var format = asplib.lang.getProperty("shortdatetime", lang);
	return this.format(format, lang);
};

// toLongTime()
Date.prototype.toLongTime = function(lang) {
	var format = asplib.lang.getProperty("longtime", lang);
	return this.format(format, lang);
};

// toLongDate()
Date.prototype.toLongDate = function(lang) {
	var format = asplib.lang.getProperty("longdate", lang);
	return this.format(format, lang);
};

// toLongDateTime()
Date.prototype.toLongDateTime = function(lang) {
	var format = asplib.lang.getProperty("longdatetime", lang);
	return this.format(format, lang);
};

// getDayName()
Date.prototype.getDayName = function(day, lang) {
	var day = (typeof day !== "undefined") ? day : this.getDay();
	var weekdays = asplib.lang.getProperty("weekdays", lang);
	return (weekdays && weekdays[day]) ? weekdays[day] : "";
};

// getMonthName()
Date.prototype.getMonthName = function(month, lang) {
	var month = (typeof month !== "undefined") ? month : this.getMonth();
	var months = asplib.lang.getProperty("months", lang);
	return (months && months[month]) ? months[month] : "";
};

%>