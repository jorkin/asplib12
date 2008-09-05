<%

/* excel.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.database) asplib.database = {};

//--------------------------------------------------------------------
// asplib.database.ExcelConnection()
//--------------------------------------------------------------------

asplib.database.ExcelConnection = function(dataSource, header) {

	this.__name = "asplib.database.ExcelConnection";
	trace("asplib.database.ExcelConnection")

	// Properties

	this.nullString = "NULL";
	this.trueString = "True";
	this.falseString = "False";
	this.dateDelimiter = "#";
	this.provider = "Microsoft.Jet.OLEDB.4.0";
	this.dataSource = "";
	this.header = (header) ? true : false;

	if (typeof dataSource !== "undefined") this.dataSource = dataSource;

};

// Derived from: asplib.hyperweb.Connection
asplib.database.ExcelConnection.prototype = new asplib.database.Connection;
asplib.database.ExcelConnection.prototype.constructor = asplib.database.ExcelConnection;

// Open()
asplib.database.ExcelConnection.prototype.Open = function(dataSource) {
	trace("asplib.database.ExcelConnection.Open");
	if (typeof dataSource != "undefined") this.dataSource = dataSource;
	if (!this.connectionString) {
		if (!this.dataSource) this.dataSource = Server.MapPath("/") + "_data/";
		this.connectionString =
		"Provider=" + this.provider + "; " +
		"Data source=" + this.dataSource + "; " +
		"User Id=admin; " +
		"Password=; " +
		"Extended Properties=\"Excel 8.0;HDR=" + (this.header ? "Yes" : "No")+ "\";";
	}
	asplib.database.Connection.prototype.Open.call(this);
};

// CompactAndRepair()
asplib.database.ExcelConnection.prototype.CompactAndRepair = function() {
	trace("asplib.database.ExcelConnection.CompactAndRepair()")
};

%>