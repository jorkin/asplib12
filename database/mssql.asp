<%

/* mssql.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.database) asplib.database = {};

//--------------------------------------------------------------------
// asplib.database.MsSqlConnection()
//--------------------------------------------------------------------

asplib.database.MsSqlConnection = function(dataSource, initialCatalog, userId, password) {
	this.__name = "asplib.database.MsSqlConnection";
	trace("asplib.database.MsSqlConnection");
	// Properties
	this.nullString = "NULL";
	this.trueString = "1";
	this.falseString = "0";
	this.dateDelimiter = "'";
	this.provider = "SQLOLEDB";
	this.dataSource = "localhost";
	this.initialCatalog = "";
	this.userId = "sa";
	this.password = "";
	if (dataSource) this.dataSource = dataSource;
	if (initialCatalog) this.initialCatalog = initialCatalog;
	if (userId) this.userId	= userId;
	if (password) this.password = password;
};

// Derived from: asplib.hyperweb.Connection
asplib.database.MsSqlConnection.prototype = new asplib.database.Connection;
asplib.database.MsSqlConnection.prototype.constructor = asplib.database.MsSqlConnection;

// Open()
asplib.database.MsSqlConnection.prototype.Open = function(dataSource, initialCatalog, userId, password) {
	trace("asplib.database.MsSqlConnection.Open");
	if (dataSource) this.dataSource = dataSource;
	if (initialCatalog) this.initialCatalog = initialCatalog;
	if (userId) this.userId	= userId;
	if (password) this.password = password;
	this.connectionString =
	"Provider=" + this.provider + "; " +
	"Data Source=" + this.dataSource + "; " +
	"Initial Catalog=" + this.initialCatalog + "; " +
	"User Id=" + this.userId + "; " +
	"Password=" + this.password + ";";
	asplib.database.Connection.prototype.Open.call(this);
};

// GetDate()
asplib.database.MsSqlConnection.prototype.GetDate = function() {
	trace("asplib.database.MsSqlConnection.GetDate");
	return "GETDATE()";
};

%>

