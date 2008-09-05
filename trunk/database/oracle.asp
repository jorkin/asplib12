<%

/* oracle.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.database) asplib.database = {};

//--------------------------------------------------------------------
// asplib.database.OracleConnection()
//--------------------------------------------------------------------

asplib.database.OracleConnection = function(dataSource, userId, password) {
	this.__name = "asplib.database.OracleConnection";
	trace("asplib.database.OracleConnection");
	// Properties
	this.nullString = "NULL";
	this.trueString = "1";
	this.falseString = "0";
	this.dateDelimiter = "'";
	this.provider = "ORAOLEDB.Oracle";
	this.dataSource = "localhost";
	this.initialCatalog = "";
	this.userId = "system";
	this.password = "";
	if (dataSource) this.dataSource = dataSource;
	if (userId) this.userId	= userId;
	if (password) this.password = password;
};

// Derived from: asplib.hyperweb.Connection
asplib.database.OracleConnection.prototype = new asplib.database.Connection;
asplib.database.OracleConnection.prototype.constructor = asplib.database.OracleConnection;

// Open()
asplib.database.OracleConnection.prototype.Open = function(dataSource, userId, password) {
	trace("asplib.database.OracleConnection.Open");
	if (dataSource) this.dataSource = dataSource;
	if (userId) this.userId	= userId;
	if (password) this.password = password;
	this.connectionString =
	"Provider=" + this.provider + "; " +
	"Data Source=" + this.dataSource + "; " +
	"User Id=" + this.userId + "; " +
	"Password=" + this.password + ";";
	asplib.database.Connection.prototype.Open.call(this);
};

// GetDate()
asplib.database.OracleConnection.prototype.GetDate = function() {
	trace("asplib.database.OracleConnection.GetDate");
	return "GETDATE()";
};

%>

