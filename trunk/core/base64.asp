<%

/* base64.asp - by Thomas Kjoernes <thomas@ipv.no> */

var Base64 = {

	chars:
	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" +
	"abcdefghijklmnopqrstuvwxyz" +
	"0123456789+/=",

	encode: function(text) {
		var text = (text || "");
		var temp = [];
		var i = 0;
		var l = text.length;
		while (i < l) {
			var c1 = text.charCodeAt(i++);
			var c2 = text.charCodeAt(i++);
			var c3 = text.charCodeAt(i++);
			var e1 = c1 >> 2;
			var e2 = ((c1 & 3) << 4) | (c2 >> 4);
			var e3 = ((c2 & 15) << 2) | (c3 >> 6);
			var e4 = c3 & 63;
			if (isNaN(c2)) {
				e3 = e4 = 64;
			} else if (isNaN(c3)) {
				e4 = 64;
			}
			temp.push(this.chars.charAt(e1));
			temp.push(this.chars.charAt(e2));
			temp.push(this.chars.charAt(e3));
			temp.push(this.chars.charAt(e4));
		}
		return temp.join("");
	},

	decode: function(text) {
		var text = (text || "").replace(/[^A-Za-z0-9\+\/\=]/g, "");
		var temp = [];
		var i = 0;
		var l = text.length;
		while (i < l) {
			var e1 = this.chars.indexOf(text.charAt(i++));
			var e2 = this.chars.indexOf(text.charAt(i++));
			var e3 = this.chars.indexOf(text.charAt(i++));
			var e4 = this.chars.indexOf(text.charAt(i++));
			var c1 = (e1 << 2) | (e2 >> 4);
			var c2 = ((e2 & 15) << 4) | (e3 >> 2);
			var c3 = ((e3 & 3) << 6) | e4;
			temp.push(String.fromCharCode(c1));
			if (e3 != 64) temp.push(String.fromCharCode(c2));
			if (e4 != 64) temp.push(String.fromCharCode(c3));
		}
		return temp.join("");
	}

};

%>
