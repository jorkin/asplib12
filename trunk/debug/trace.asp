<%

/* trace.asp - by Thomas Kjoernes <thomas@ipv.no> */

// trace()
function trace(name, text, raw, url) {
	if (asplib.debug.trace.enabled) {
		if (!asplib.debug.trace.time) {
			asplib.debug.trace.time = new Date();
		}
		if (!asplib.debug.cache["fso"]) {
			asplib.debug.cache["fso"] = new ActiveXObject("Scripting.FileSystemObject");
		}
		var fso = asplib.debug.cache["fso"];
		var filename = url || asplib.debug.trace.filename;
		var stream = asplib.debug.cache[filename];
		var script = Request.ServerVariables("SCRIPT_NAME");
		try {
			if (!stream) {
				asplib.debug.cache[filename] = fso.OpenTextFile(Server.MapPath(filename), 8, true); // ForAppending=8
				stream = asplib.debug.cache[filename];
				if (!raw) stream.WriteLine("--------------------------------------------------------------------------------");
			}
			var stream = asplib.debug.cache[filename];
			if (!raw) {
				var timestamp = new Date().valueOf() - asplib.debug.trace.time.valueOf();
				var caller = arguments.callee.caller;
				var fn = caller+"";
				var i = fn.indexOf("function")+9;
				var j = fn.indexOf("(");
				var k = fn.indexOf(")");
				var __name = (i<j) ? fn.substring(i,j) : name;
				var __args = (k>j+1) ? fn.substring(j+1, k).split(",") : [];
				var temp = [];
				for (var i=0; i<__args.length; i++) {
					var v = caller.arguments[i];
					var s = (v + "");
					if (s.length>160) s = s.substring(0,160) + "...";
					s = s.replace(/\n/g, "\\n").replace(/\r/g, "\\r").replace(/\t/g, "\\t");
					switch (typeof v) {
						case "undefined" : s = "<>"; break;
						case "string" : s = "\"" + s + "\""; break;
						case "number" :
						case "boolean" :
						case "date" :
						case "object" :
						default :
					}
					temp.push(__args[i] + "=" + s);
				}
				stream.WriteLine("[" + asplib.debug.trace.time.toUTCString() + "] " + script + ": " + (__name+"()" + (temp.length ? " " + temp : "") + "") + (text ? " {" + text + "}" : "") + " (" + timestamp + "ms)");
			} else {
				stream.WriteLine(text);
			}
		} catch(e) {
		}
	}
}

%>