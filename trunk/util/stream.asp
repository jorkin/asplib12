<%

/* stream.asp - by Thomas Kjoernes <thomas@ipv.no> */

//
function LoadTextFile(path, charset)  {
	trace("LoadTextFile", "path: " + path);
	try {
		var stream = new ActiveXObject("ADODB.Stream");
		if (typeof charset == "undefined") charset = "ISO-8859-1";
		stream.Type = 2; // adTypeText=2
		stream.Charset = charset;
		stream.Open();
		stream.LoadFromFile(path);
		var content = stream.ReadText();
		stream.Close();
		stream = null;
		return content;
	} catch(e) { }
}

//
function SaveTextFile(path, content, charset)  {
	trace("SaveTextFile", "path: " + path);
	try {
		var stream = new ActiveXObject("ADODB.Stream");
		if (typeof charset == "undefined") charset = "ISO-8859-1";
		stream.Type = 2; // adTypeText=2
		stream.Charset = charset;
		stream.Open();
		stream.WriteText(content);
		stream.SaveToFile(path);
		stream.Close();
		stream = null;
	} catch(e) { }
}

//
function LoadBinaryFile(path)  {
	trace("LoadBinaryFile", "path: " + path);
	try {
		var stream = new ActiveXObject("ADODB.Stream");
		stream.Type = 1; // adTypeBinary=1
		stream.Open();
		stream.LoadFromFile(path);
		var content = stream.Read();
		stream.Close();
		stream = null;
		return content;
	} catch(e) { }
}

//
function SaveBinaryFile(path, content)  {
	trace("SaveBinaryFile", "path: " + path);
	try {
		var stream = new ActiveXObject("ADODB.Stream");
		stream.Type = 1; // adTypeBinary=1
		stream.Open();
		stream.Write(content);
		stream.SaveToFile(path);
		stream.Close();
		stream = null;
	} catch(e) { }
}

%>