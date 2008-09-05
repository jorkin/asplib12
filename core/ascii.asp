<%

/* ascii.asp - by Thomas Kjoernes <thomas@ipv.no> */

var Ascii = {

	charmap: {
		0xc0: "A",	0xc1: "A",	0xc2: "A",	0xc3: "A",
		0xc4: "A",	0xc5: "Aa",	0xc6: "Ae",	0xc7: "C",
		0xc8: "E",	0xc9: "E",	0xca: "E",	0xcb: "E",
		0xcc: "I",	0xcd: "I",	0xce: "I",	0xcf: "I",
		0xd0: "Th",	0xd1: "N",	0xd2: "O",	0xd3: "O",
		0xd4: "O",	0xd5: "O",	0xd6: "O",
		0xd8: "Oe",	0xd9: "U",	0xda: "U",	0xdb: "U",
		0xdc: "U",	0xdd: "Y",	0xde: "th",	0xdf: "ss",
		0xe0: "a",	0xe1: "a",	0xe2: "a",	0xe3: "a",
		0xe4: "a",	0xe5: "aa",	0xe6: "ae",	0xe7: "c",
		0xe8: "e",	0xe9: "e",	0xea: "e",	0xeb: "e",
		0xec: "i",	0xed: "i",	0xee: "i",	0xef: "i",
		0xf0: "th",	0xf1: "n",	0xf2: "o",	0xf3: "o",
		0xf4: "o",	0xf5: "o",	0xf6: "o",
		0xf8: "oe",	0xf9: "u",	0xfa: "u",	0xfb: "u",
		0xfc: "u",	0xfd: "y",	0xfe: "th",	0xff: "y"
	},

	filter: function(text) {
		var text = (text || "");
		var temp = [];
		var i = 0;
		var l = text.length;
		while (i < l) {
			var chr = text.charAt(i);
			var asc = text.charCodeAt(i++);
			if (asc > 127) {
				chr = this.charmap[asc] || "";
			}
			temp.push(chr);
		}
		return temp.join("");
	}

};

String.prototype.toAscii = function() {
	return Ascii.filter(this.valueOf());
};

%>