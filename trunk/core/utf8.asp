<%

/* utf8.asp - by Thomas Kjoernes <thomas@ipv.no> */

var Utf8 = {

	encode: function(text) {
		var text = (text || "");
		var temp = [];
		var i = 0;
		var l = text.length;
		while (i < l) {
			var a = text.charCodeAt(i++);
			if (a < 128) {
				temp.push(String.fromCharCode(a));
			} else if ((a > 127) && (a < 2048)) {
				temp.push(String.fromCharCode((a >> 6) | 192));
				temp.push(String.fromCharCode((a & 63) | 128));
			} else {
				temp.push(String.fromCharCode((a >> 12) | 224));
				temp.push(String.fromCharCode(((a >> 6) & 63) | 128));
				temp.push(String.fromCharCode((a & 63) | 128));
			}
		}
		return temp.join("");
	},

	decode: function(text) {
		var text = (text || "");
		var temp = [];
		var i = 0;
		var l = text.length;
		while (i < l) {
			var a = text.charCodeAt(i++);
			if (a < 128) {
				temp.push(String.fromCharCode(a));
			} else if ((a > 191) && (a < 224)) {
				var b = text.charCodeAt(i++);
				temp.push(String.fromCharCode(((a & 31) << 6) | (b & 63)));
			} else {
				var b = text.charCodeAt(i++);
				var c = text.charCodeAt(i++);
				temp.push(String.fromCharCode(((a & 15) << 12) | ((b & 63) << 6) | (c & 63)));
			}
		}
		return temp.join("");
	},

	validate: function(text) {
		return false;
	}

};

String.prototype.toUTF8 = function() {
	return Utf8.encode(this.valueOf());
};

String.prototype.fromUTF8 = function() {
	return Utf8.decode(this.valueOf());
}

%>