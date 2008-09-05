<%

/* parse.asp - by Thomas Kjoernes <thomas@ipv.no> */

	var global = this;

// parseNumber()
function parseNumber(value, nullValue) {
	if (value === null) return nullValue;
	if (typeof value === "undefined") return nullValue;
	if (typeof value === "number" && !isNaN(value)) return value;
	var temp = value.toString().replace(/\,/g, ".").toLowerCase();;
	if (temp.indexOf(".")!==-1) {
		temp = parseFloat(temp);
	} else {
		temp = parseInt(temp, 10);
	}
	return !isNaN(temp) ? temp : nullValue;
}

// parseDate()
function parseDate(value, nullValue) {
	if (value == null) return nullValue;
	if (typeof value === "undefined") return nullValue;
	if (typeof value === "date" && !isNaN(value)) return new Date(value);
	if (typeof value === "number" && !isNaN(value)) return new Date(value);
	if (typeof value === "object" && typeof value.getTime === "function") return value;
	if (typeof value !== "string") value = value.toString();
	if (value.indexOf("UTC") !== -1 || value.indexOf("GMT") !== -1) {
		var temp = new Date(Date.parse(value));
	} else {
		var temp = value.toLowerCase().replace(/[t]/g, " ").replace(/\s/g, " ").split(" ");
		var tempDate = (temp[0]||"").replace(/\//g, "-").replace(/\\/g, "-").replace(/\./g, "-").replace(/\:/g, "-").split("-");
		var tempTime = (temp[1]||"").replace(/\./g, ":").replace(/\,/g, ":").split(":");
		var year = parseInt(tempDate[0], 10);
		var month = parseInt(tempDate[1], 10)-1;
		var date = parseInt(tempDate[2], 10);
		var hours = parseInt(tempTime[0], 10);
		var minutes = parseInt(tempTime[1], 10);
		var seconds = parseInt(tempTime[2], 10);
		var fraction = parseInt(tempTime[3], 10);
		if (!date) {
			if (year<32) {
				date = new Date().getFullYear();
			} else {
				date = new Date().getDate();
			}
		}
		if (year<32 && date<32) year = 1900 + year;
		if (year<32 && date>31 || !date) {
			var tempYear = year;
			year = date;
			date = tempYear;
		}
		if (value.match(/pm/ig)) hours += 12;
		if (!hours) hours = 0;
		if (!minutes) minutes = 0;
		if (!seconds) seconds = 0;
		if (!fraction) fraction = 0;
		var temp = new Date(year, month, date, hours, minutes, seconds, fraction);
	}
	return (!isNaN(temp))  ? temp : nullValue;
}

// parseBoolean()
function parseBoolean(value, nullValue) {
	if (value === null) return nullValue;
	if (typeof value === "undefined") return nullValue;
	var temp = value.toString().toLowerCase();
	if (temp.match(/[1-9]/g)) return true;
	if (temp.indexOf("on") !== -1) return true;
	if (temp.indexOf("yes") !== -1) return true;
	if (temp.indexOf("true") !== -1) return true;
	return false;
}

// parseString()
function parseString(value, nullValue) {
	if (value === null) return nullValue;
	if (typeof value === "undefined") return nullValue;
	return value.toString();
}

// parseBits()
function parseBits(value, nullValue) {
	if (value === null) return nullValue;
	if (typeof value === "undefined") return nullValue;
	if (typeof value === "number" && !isNaN(value)) return value;
	var temp = 0;
	var list = (typeof value === "string") ? value.split(",") : value;
	try {
		for (var i in list) {
			temp = temp | list[i];
		}
	} catch(e) { }
	return temp;
}

// parseEnum
function parseEnum(value, nullValue) {
	if (value === null) return nullValue;
	if (typeof value === "undefined") return nullValue;
	if (typeof value === "string") {
		var temp = value.split("|");
		if (temp.length>1) {
			var value = 0;
			for (var i=0; i<temp.length; i++) {
				var name = temp[i].trim();
				value = value | parseNumber(global[name], 0);
			}
			return value;
		} else {
			return (typeof global[value] !== "undefined") ? global[value] : nullValue;
		}
	} else {
		return value;
	}
}

// parseCookies
function parseCookies() {
	var obj = {};
	for (var e = new Enumerator(Request.Cookies); !e.atEnd(); e.moveNext()) {
		var name = e.item();
		var value = Request.Cookies(name).Item();
		obj[name] = value;
	}
	return obj;
}

// parseForm
function parseForm() {
	var obj = {};
	for (var e = new Enumerator(Request.Form); !e.atEnd(); e.moveNext()) {
		var name = e.item();
		var value = Request.Form(name).Item();
		obj[name] = value;
	}
	return obj;
}

// parseQueryString
function parseQueryString() {
	var obj = {};
	for (var e = new Enumerator(Request.QueryString); !e.atEnd(); e.moveNext()) {
		var name = e.item();
		var value = Request.QueryString(name).Item();
		obj[name] = value;
	}
	return obj;
}

// parseServerVariables
function parseServerVariables() {
	var obj = {};
	for (var e = new Enumerator(Request.ServerVariables); !e.atEnd(); e.moveNext()) {
		var name = e.item();
		var value = Request.ServerVariables(name).Item();
		obj[name] = value;
	}
	return obj;
}

%>