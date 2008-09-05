<%

/* connection.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.database) asplib.database = {};

	// Base Types
	var BASE_TYPE_UNKNOWN		= 0;
	var BASE_TYPE_NUMBER		= 1;
	var BASE_TYPE_DATE			= 2;
	var BASE_TYPE_BOOLEAN		= 3;
	var BASE_TYPE_TEXT			= 4;

	// ADO Types
	var ADO_TYPE_INT			= 3;
	var ADO_TYPE_FLOAT			= 5;
	var ADO_TYPE_MONEY			= 6;
	var ADO_TYPE_DATETIME		= 7;
	var ADO_TYPE_BIT			= 11;
	var ADO_TYPE_VARCHAR		= 200;
	var ADO_TYPE_TEXT			= 201;
	var ADO_TYPE_NVARCHAR		= 202;
	var ADO_TYPE_NTEXT			= 203;

	var __adoFieldTypes = {
		20:		{ name: "adBigInt", type: BASE_TYPE_NUMBER },
		128:	{ name: "adBinary", type: BASE_TYPE_UNKNOWN },
		11:		{ name: "adBoolean", type: BASE_TYPE_BOOLEAN },
		8:		{ name: "adBSTR", type: BASE_TYPE_UNKNOWN },
		136:	{ name: "adChapter", type: BASE_TYPE_UNKNOWN },
		129:	{ name: "adChar", type: BASE_TYPE_TEXT },
		6:		{ name: "adCurrency", type: BASE_TYPE_NUMBER },
		7:		{ name: "adDate", type: BASE_TYPE_DATE },
		133:	{ name: "adDBDate", type: BASE_TYPE_DATE },
		137:	{ name: "adDBFileTime", type: BASE_TYPE_DATE },
		134:	{ name: "adDBTime", type: BASE_TYPE_DATE },
		135:	{ name: "adDBTimeStamp", type: BASE_TYPE_DATE },
		14:		{ name: "adDecimal", type: BASE_TYPE_NUMBER },
		5:		{ name: "adDouble", type: BASE_TYPE_NUMBER },
		0:		{ name: "adEmpty", type: BASE_TYPE_UNKNOWN },
		10:		{ name: "adError", type: BASE_TYPE_UNKNOWN },
		64:		{ name: "adFileTime", type: BASE_TYPE_DATE },
		72:		{ name: "adGUID", type: BASE_TYPE_UNKNOWN },
		9:		{ name: "adIDispatch", type: BASE_TYPE_UNKNOWN },
		3:		{ name: "adInteger", type: BASE_TYPE_NUMBER },
		13:		{ name: "adIUnknown", type: BASE_TYPE_UNKNOWN },
		205:	{ name: "adLongVarBinary", type: BASE_TYPE_UNKNOWN },
		201:	{ name: "adLongVarChar", type: BASE_TYPE_TEXT },
		203:	{ name: "adLongVarWChar", type: BASE_TYPE_TEXT },
		131:	{ name: "adNumeric", type: BASE_TYPE_NUMBER },
		138:	{ name: "adPropVariant", type: BASE_TYPE_UNKNOWN },
		4:		{ name: "adSingle", type: BASE_TYPE_NUMBER },
		2:		{ name: "adSmallInt", type: BASE_TYPE_NUMBER },
		16:		{ name: "adTinyInt", type: BASE_TYPE_NUMBER },
		21:		{ name: "adUnsignedBigInt", type: BASE_TYPE_NUMBER },
		19:		{ name: "adUnsignedInt", type: BASE_TYPE_NUMBER },
		18:		{ name: "adUnsignedSmallInt", type: BASE_TYPE_NUMBER },
		17:		{ name: "adUnsignedTinyInt", type: BASE_TYPE_NUMBER },
		132:	{ name: "adUserDefined", type: BASE_TYPE_UNKNOWN },
		204:	{ name: "adVarBinary", type: BASE_TYPE_UNKNOWN },
		200:	{ name: "adVarChar", type: BASE_TYPE_TEXT },
		12:		{ name: "adVariant", type: BASE_TYPE_UNKNOWN },
		139:	{ name: "adVarNumeric", type: BASE_TYPE_NUMBER },
		202:	{ name: "adVarWChar", type: BASE_TYPE_TEXT },
		130:	{ name: "adWChar", type: BASE_TYPE_TEXT },
		8192:	{ name: "adArray", type: BASE_TYPE_UNKNOWN }
	};

// __getAdoTypeName()
function __getAdoTypeName(type) {
	return (__adoFieldTypes[type]) ? __adoFieldTypes[type].name : "";
}

// __getDateFormat()
function __getDateFormat(date, dateFormat) {
	var temp = dateFormat;
	temp = temp.replace("YYYY", date.getFullYear().toString(10));
	temp = temp.replace("YY", date.getYear().toString(10, 2));
	temp = temp.replace("MM", (date.getMonth()-1).toString(10, 2));
	temp = temp.replace("DD", date.getDate().toString(10, 2));
	temp = temp.replace("hh", date.getHours().toString(10, 2));
	temp = temp.replace("mm", date.getMinutes().toString(10, 2));
	temp = temp.replace("ss", date.getSeconds().toString(10, 2));
	return temp;
}

// __getGenericType()
function __getGenericType(type) {
	switch (type) {
		case ADO_TYPE_INT :
		case ADO_TYPE_FLOAT :
		case ADO_TYPE_MONEY : return BASE_TYPE_NUMBER;
		case ADO_TYPE_DATETIME : return BASE_TYPE_DATE;
		case ADO_TYPE_BIT : return BASE_TYPE_BOOLEAN;
		case ADO_TYPE_VARCHAR :
		case ADO_TYPE_TEXT :
		case ADO_TYPE_NVARCHAR :
		case ADO_TYPE_NTEXT : return BASE_TYPE_TEXT;
		default : return BASE_TYPE_UNKNOWN;
	}
}

// __getGenericTypeName()
function __getGenericTypeName(type) {
	switch (type) {
		case ADO_TYPE_INT :
		case ADO_TYPE_FLOAT :
		case ADO_TYPE_MONEY : return "number";
		case ADO_TYPE_DATETIME : return "date";
		case ADO_TYPE_BIT : return "boolean";
		case ADO_TYPE_VARCHAR :
		case ADO_TYPE_TEXT :
		case ADO_TYPE_NVARCHAR :
		case ADO_TYPE_NTEXT : return "string";
		default : return "unknown";
	}
}

// __getValueByType()
function __getValueByType(type, value) {
	if (typeof value === "undefined") return null;
	switch (type) {
		case "string" :
		case BASE_TYPE_TEXT :
			return parseString(value);
		case "number" :
		case BASE_TYPE_NUMBER :
			return parseNumber(value);
		case "boolean" :
		case BASE_TYPE_BOOLEAN :
			return parseBoolean(value);
		case "date" :
		case "object" :
		case BASE_TYPE_DATE :
			var temp = parseDate(value);
			if (temp) return value;
		default :
			// should never get here
			return value;
	}
}

// __showDatabaseError()
function __showDatabaseError(e, conn, sql, msg) {
	Response.Write("<div style=\"border:1px solid #CC9;background-color:#FFC;padding:1em;font-family:Arial;font-size:10pt;\">\n");
	if (msg) Response.Write("<strong>Operation:</strong> <span style=\"color:blue\">" + msg + "</span><br>\n");
	if (sql) Response.Write("<strong>Error in SQL statement:</strong> <span style=\"color:blue\">" + sql + "</span><br>\n");
	Response.Write("<strong>Error code: <span style=\"color:red\">" + e.number.uint().toString(16) + "</span></strong><br>");
	Response.write("<strong>Description:</strong> " + e.description + "<br>\n");
	if (conn) Response.Write("<strong>Connection string:</strong><br>\n");
	if (conn && conn.ADODBConnection) {
		var connstr = conn.ADODBConnection.ConnectionString;
		var i = connstr.toLowerCase().indexOf("password=");
		var j = connstr.indexOf(";", i);
		connstr = connstr.substring(0, i+9) + connstr.substring(j);
		Response.Write(
			"<span style=\"color:green\">" +
			connstr.replace(/;/g,";<br>\n") +
			"</span><br>\n"
		);
	}
	Response.Write("</div>\n");
	Response.End();
}

function __getHtmlTable(rs, settings) {
	if (!rs) return;
	var html = [];
	var fields = [];
	var settings = settings || {};
	if (typeof settings.showLabels === "undefined") settings.showLabels = true;
	if (typeof settings.showRowNumbers === "undefined") settings.showRowNumbers = true;
	for (var e = new Enumerator(rs.Fields); !e.atEnd(); e.moveNext()) {
		fields.push(e.item());
	}
	html.push("<table cellspacing=\"0\" cellpadding=\"0\">\n");
	if (settings.caption) {
		html.push("<caption>" + caption + "</caption>");
	}
	if (settings.showLabels) {
		html.push("<thead>\n");
		html.push("<tr>");
		if (settings.showRowNumbers) {
			html.push( "<th>" + "&nbsp;" + "</th>");
		}
		for (var i=0; i<fields.length; i++) {
			var name = fields[i].Name;
			var label = (settings.labels) ? settings.labels[name]||name : name;
			html.push( "<th id=\"__" + name + "\">" + label + "</th>");
		}
		html.push("</tr>\n");
		html.push("</thead>\n");
	}
	if (!rs.Eof) {
		html.push("<tbody>\n");
		var row=1;
		while (!rs.Eof) {
			html.push("<tr>");
			if (settings.showRowNumbers) {
				html.push("<th>" + (row++) + "</th>");
			}
			for (var i=0; i<fields.length; i++) {
				var classNames = [];
				var name = fields[i].Name;
				var value = rs(i).Value;
				var type = typeof value;

				switch (type) {
					case "object" :
						if (!value || typeof value.getTime !== "function") break;
					case "date" :
						value = new Date(value); break;
					case "string" :
						if (value.length === 0) {
							value = "&nbsp;";
							classNames.push("empty");
						}
						if (value.length > 50) {
							value = Server.HTMLEncode(value.substring(0, 50) + "...");
						}
						break;
					case "unknown" :
						value = "binary"; break;
				}
				classNames.push(type);

				if (value && name.toLowerCase().indexOf("url") !== -1) value = "<a href=\"" + new URL(value) + "\">" + value + "</a>";

				html.push("<td class=\"" + classNames.join(" ") + "\">" + value + "</td>");
			}
			html.push("</tr>\n");
			rs.MoveNext();
		}
		html.push("</tbody>\n");
	}
	html.push("</table>\n");
	return html.join("");
}

//--------------------------------------------------------------------
// asplib.database.Connection()
//--------------------------------------------------------------------

asplib.database.Connection = function(connectionString) {
	this.__name = "asplib.database.Connection";
	this.dateDelimiter = "'";
	if (typeof connectionString !== "undefined") this.connectionString = connectionString;
};

// Open()
asplib.database.Connection.prototype.Open = function(connectionString) {
	trace("asplib.database.Connection.Open");
	if (typeof connectionString !== "undefined") this.connectionString = connectionString;
	try {
		this.ADODBConnection = new ActiveXObject("ADODB.Connection");
		this.ADODBConnection.ConnectionString = this.connectionString;
		this.ADODBConnection.Open();
		this.isOpen = true;
	} catch(e) {
		__showDatabaseError(e, this, null, "Open");
	}
};

// Close()
asplib.database.Connection.prototype.Close = function() {
	trace("asplib.database.Connection.Close");
	try {
		this.ADODBConnection.Close();
		this.isOpen = false;
	} catch(e) {
		__showDatabaseError(e, this, null, "Close");
	}
};

// Execute()
asplib.database.Connection.prototype.Execute = function(sql) {
	trace("asplib.database.Connection.Execute");
	try {
	//	Response.Write(sql+"<br>\n");
		return this.ADODBConnection.Execute(sql);
	} catch(e) {
		__showDatabaseError(e, this, sql, "Execute");
	}
};

// EnumTables()
asplib.database.Connection.prototype.EnumTables = function(schemaName) {
	trace("asplib.database.Connection.EnumTables");
	try {
		var names = [];
		var constraints = [];
		if (schemaName) constraints[1] = schemaName;
		constraints[3] = "TABLE";
		var rs = this.ADODBConnection.OpenSchema(20, constraints); // adSchemaTables
		while (!rs.Eof) {
			names.push(rs(2).Value);
			rs.MoveNext();
		}
		return names;
	} catch(e) {
		__showDatabaseError(e, this, null, "OpenSchema");
	}
};

// EnumViews()
asplib.database.Connection.prototype.EnumViews = function(schemaName) {
	trace("asplib.database.Connection.EnumViews");
	try {
		var names = [];
		var constraints = [];
		if (schemaName) constraints[1] = schemaName;
		constraints[3] = "VIEW";
		var rs = this.ADODBConnection.OpenSchema(20, constraints); // adSchemaTables
		while (!rs.Eof) {
			names.push(rs(2).Value);
			rs.MoveNext();
		}
		return names;
	} catch(e) {
		__showDatabaseError(e, this, null, "OpenSchema");
	}
};

// CompactAndRepair()
asplib.database.Connection.prototype.CompactAndRepair = function() {
	trace("asplib.database.Connection.CompactAndRepair");
};

// AsText()
asplib.database.Connection.prototype.AsText = function(value, maxlength) {
	var temp = parseString(value);
	if (temp == null) {
		return this.nullString;
	} else {
		if (maxlength) temp = temp.substring(0, maxlength);
		return "\'" + temp.replace(/\'/g, "\'\'") + "\'";

	}
};

// AsNumber()
asplib.database.Connection.prototype.AsNumber = function(value) {
	var temp = parseNumber(value);
	if (temp == null) {
		return this.nullString;
	} else {
		return temp.toString();
	}
s};

// AsDate()
asplib.database.Connection.prototype.AsDate = function(value) {
	var temp = parseDate(value);
	if (temp == null) {
		return this.nullString;
	} else {
		return this.dateDelimiter + temp.toISOString() + this.dateDelimiter;
	}
};

// AsBoolean()
asplib.database.Connection.prototype.AsBoolean = function(value) {
	var temp = parseBoolean(value);
	if (temp == null) {
		return this.nullString;
	} else {
		return temp.toYesNoString(this.trueString, this.falseString);
	}
};

// AsSearchText()
asplib.database.Connection.prototype.AsSearchText = function(value) {
	var temp = parseString(value);
	if (temp == null) {
		return this.nullString;
	} else {
		return this.AsText(temp.replace(/%/g, "[%]").replace(/_/g, "[_]").replace(/\*/g, "%").replace(/\?/g, "_"));
	}
};

// As()
asplib.database.Connection.prototype.As = function(type, value, maxlength) {
	switch (type) {
		case BASE_TYPE_TEXT :
			return this.AsText(value, maxlength);
		case BASE_TYPE_NUMBER :
			return this.AsNumber(value);
		case BASE_TYPE_DATE :
			return this.AsDate(value);
		case BASE_TYPE_BOOLEAN :
			return this.AsBoolean(value);
	}
};

// TypeName()
asplib.database.Connection.prototype.TypeName = function(type) {
	return __getAdoTypeName(type);
};

// GetHtmlTable()
asplib.database.Connection.prototype.GetHtmlTable = function(sql, settings) {
	var rs = this.Execute(sql);
	var html = __getHtmlTable(rs, settings);
	rs = null;
	return html;
};

// GetDate()
asplib.database.Connection.prototype.GetDate = function() {
	return this.AsDate(new Date);
};

%>