<%

//--------------------------------------------------------------------
// RawRequest
//--------------------------------------------------------------------

asplib.rpc.RawRequest = function(data) {
	this.__name = "asplib.rpc.RawRequest";
	this.type = XMLRPC_REQUEST_RAW;
	this.data = (typeof data !== "undefined") ? data : "";
};

asplib.rpc.RawRequest.prototype.toBinary = function() {
	if (typeof this.data === "unknown") {
		return this.data;
	} else {
		var stream = new ActiveXObject("ADODB.Stream");
		stream.Open();
		stream.Type = 2; // adTypeText
		stream.Charset = asplib.rpc.config["charset"];
		if (this.data != null) {
			stream.WriteText((this.data||"").toString());
		}
		stream.Position = 0;
		stream.Type = 1; // adTypeBinary
		var data = stream.Read();
		stream = null;
		return data;
	}
};

asplib.rpc.RawRequest.prototype.toString = function() {
	if (typeof this.data === "unknown") {
		var stream = new ActiveXObject("ADODB.Stream");
		stream.Open();
		stream.Type = 1; // adTypeBinary
		stream.Write(this.data);
		stream.Position = 0;
		stream.Type = 2; // adTypeText
		stream.Charset = asplib.rpc.config["charset"];
		var data = stream.ReadText();
		stream = null;
		return data;
	} else {
		return (this.data||"").toString();
	}
};

asplib.rpc.RawRequest.prototype.LoadFromFile = function(path) {
	var stream = new ActiveXObject("ADODB.Stream");
	stream.Open();
	stream.Type = 1; // adTypeBinary
	try {
		stream.LoadFromFile(path);
		this.data = stream.Read();
	} catch(e) { }
	stream = null;
};

asplib.rpc.RawRequest.prototype.SaveToFile = function(path) {
	var stream = new ActiveXObject("ADODB.Stream");
	stream.Open();
	if (typeof this.data === "unknown") {
		stream.Type = 1; // adTypeBinary
		stream.Write(this.data);
	} else {
		stream.Type = 2; // adTypeText
		stream.Charset = asplib.rpc.config["charset"];
		stream.WriteText((this.data||"").toString());
	}
	try {
		stream.SaveToFile(path, 2); // adSaveCreateOverwrite
	} catch(e) { }
	stream = null;
};

%>