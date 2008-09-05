<%

/* array.asp - by Thomas Kjoernes <thomas@ipv.no> */

// toArray()
function toArray(obj) {
	return (obj && typeof obj.join !== "function") ? [obj] : obj || [];
}

// forEach()
function forEach(obj, fn) {
	if (typeof obj === "object" && typeof fn === "function") {
		if (!obj.length) {
			fn(obj);
		} else {
			for (var i=0; i<obj.length; i++) {
				fn(obj[i]);
			}
		}
	}
}

%>