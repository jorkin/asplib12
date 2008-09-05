<%

/* type.asp - by Thomas Kjoernes <thomas@ipv.no> */

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
	var ADO_TYPE_DATETIME2		= 135;
	var ADO_TYPE_VARCHAR		= 200;
	var ADO_TYPE_TEXT			= 201;
	var ADO_TYPE_NVARCHAR		= 202;
	var ADO_TYPE_NTEXT			= 203;

var Type = {

	baseToName: function(type) {
		switch (type) {
			case BASE_TYPE_NUMBER: return "number";
			case BASE_TYPE_DATE : return "date";
			case BASE_TYPE_BOOLEAN : return "boolean";
			case BASE_TYPE_TEXT : return "string";
			case BASE_TYPE_UNKNOWN : return "unknown" + "(" + type + ")";
			default : return "invalid";
		}
	},

	nameToBase: function(type) {
		switch (type) {
			case "number" : return BASE_TYPE_NUMBER;
			case "date" : return BASE_TYPE_DATE;
			case "boolean" : return BASE_TYPE_BOOLEAN;
			case "string" : return BASE_TYPE_TEXT;
			default : return BASE_TYPE_UNKNOWN;
		}
	},

	adoToName: function(type) {
		switch (type) {
			case ADO_TYPE_INT : return "int";
			case ADO_TYPE_FLOAT : return "float";
			case ADO_TYPE_MONEY : return "money";
			case ADO_TYPE_DATETIME :
			case ADO_TYPE_DATETIME2 : return "datetime";
			case ADO_TYPE_BIT : return "bit";
			case ADO_TYPE_VARCHAR : return "varchar";
			case ADO_TYPE_NVARCHAR : return "nvarchar";
			case ADO_TYPE_TEXT : return "text";
			case ADO_TYPE_NTEXT : return "ntext";
			default : return "unknown" + "(" + type + ")";
		}
	},

	adoToBase: function(type) {
		switch (type) {
			case ADO_TYPE_INT :
			case ADO_TYPE_FLOAT :
			case ADO_TYPE_MONEY : return BASE_TYPE_NUMBER;
			case ADO_TYPE_DATETIME :
			case ADO_TYPE_DATETIME2 : return BASE_TYPE_DATE;
			case ADO_TYPE_BIT : return BASE_TYPE_BOOLEAN;
			case ADO_TYPE_VARCHAR :
			case ADO_TYPE_NVARCHAR :
			case ADO_TYPE_TEXT :
			case ADO_TYPE_NTEXT : BASE_TYPE_TEXT;
			default : return BASE_TYPE_UNKNOWN;
		}
	},

	adoToBaseName: function(type) {
		return this.baseToName(this.adoToBase(type));
	},

	adoToXmlName: function(type) {
		switch (type) {
			case ADO_TYPE_INT : return "int";
			case ADO_TYPE_FLOAT : return "float";
			case ADO_TYPE_MONEY : return "decimal";
			case ADO_TYPE_DATETIME :
			case ADO_TYPE_DATETIME2 : return "date";
			case ADO_TYPE_BIT : return "bool";
			case ADO_TYPE_VARCHAR :
			case ADO_TYPE_NVARCHAR : return "string";
			case ADO_TYPE_TEXT :
			case ADO_TYPE_NTEXT : return "object";
			default : return "unknown" + "(" + type + ")";
		}
	},

	xmlNameToAdo: function(type, unicode) {
		switch (type) {
			case "int" : return ADO_TYPE_INT;
			case "float" : return ADO_TYPE_FLOAT;
			case "decimal" : return ADO_TYPE_MONEY;
			case "date" : return ADO_TYPE_DATETIME;
			case "bool" : return ADO_TYPE_BOOLEAN;
			case "string" : return (unicode ? ADO_TYPE_NVARCHAR : ADO_TYPE_VARCHAR);
			case "object" : return (unicode ?  ADO_TYPE_NTEXT : ADO_TYPE_TEXT);
			default : return "unknown" + "(" + type + ")";
		}
	},

	xmlNameToBase: function(type) {
		switch (type) {
			case "int" :
			case "float" :
			case "decimal" : return BASE_TYPE_NUMBER;
			case "date" : return BASE_TYPE_DATE;
			case "bool" : return BASE_TYPE_BOOLEAN;
			case "string" :
			case "object" : return BASE_TYPE_TEXT;
			default : return BASE_TYPE_UNKNOWN;
		}
	}

};

%>