<%

/* dbase.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.database) asplib.database = {};

//--------------------------------------------------------------------
// asplib.database.DbaseConnection()
//--------------------------------------------------------------------

asplib.database.DbaseConnection = function(dataSource) {

	this.__name = "asplib.database.DbaseConnection";
	trace("asplib.database.DbaseConnection")

	// Properties

	this.nullString = "NULL";
	this.trueString = "True";
	this.falseString = "False";
	this.dateDelimiter = "#";
	this.provider = "Microsoft.Jet.OLEDB.4.0";
	this.dataSource = "";

	if (typeof dataSource !== "undefined") this.dataSource = dataSource;

};

// Derived from: asplib.hyperweb.Connection
asplib.database.DbaseConnection.prototype = new asplib.database.Connection;
asplib.database.DbaseConnection.prototype.constructor = asplib.database.DbaseConnection;

// Open()
asplib.database.DbaseConnection.prototype.Open = function(dataSource) {
	trace("asplib.database.DbaseConnection.Open");
	if (typeof dataSource != "undefined") this.dataSource = dataSource;
	if (!this.connectionString) {
		if (!this.dataSource) this.dataSource = Server.MapPath("/") + "_data/";
		this.connectionString =
		"Provider=" + this.provider + "; " +
		"Data source=" + this.dataSource + "; " +
		"User Id=admin; " +
		"Password=; " +
		"Extended Properties=\"DBASE 5.0;\"";
	}
	asplib.database.Connection.prototype.Open.call(this);
};

// CompactAndRepair()
asplib.database.DbaseConnection.prototype.CompactAndRepair = function() {
	trace("asplib.database.DbaseConnection.CompactAndRepair()")
};

%>

