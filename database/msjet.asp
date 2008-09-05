<%

/* msjet.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.database) asplib.database = {};

//--------------------------------------------------------------------
// asplib.database.MsJetConnection()
//--------------------------------------------------------------------

asplib.database.MsJetConnection = function(dataSource) {

	this.__name = "asplib.database.MsJetConnection";
	trace("asplib.database.MsJetConnection")

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
asplib.database.MsJetConnection.prototype = new asplib.database.Connection;
asplib.database.MsJetConnection.prototype.constructor = asplib.database.MsJetConnection;

// Open()
asplib.database.MsJetConnection.prototype.Open = function(dataSource) {
	trace("asplib.database.MsJetConnection.Open");
	if (typeof dataSource != "undefined") this.dataSource = dataSource;
	if (!this.connectionString) {
		if (!this.dataSource) this.dataSource = Server.MapPath("/") + "_data/data.mdb";
		this.connectionString =
		"Provider=" + this.provider + "; " +
		"Data source=" + this.dataSource + "; " +
		"User Id=admin; " +
		"Password=;";
	}
	asplib.database.Connection.prototype.Open.call(this);
};

// CompactAndRepair()
asplib.database.MsJetConnection.prototype.CompactAndRepair = function() {
	trace("asplib.database.MsJetConnection.CompactAndRepair");
	var engine = new ActiveXObject("JRO.JetEngine");
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var mdb = this.dataSource;
	var tmp = mdb + ".tmp";
	var bak = mdb + ".bak";
	fso.CopyFile(mdb, bak);
	if (fso.FileExists(tmp)) fso.DeleteFile(tmp);
	engine.CompactDatabase(
	"Provider=" + this.provider + "; Data source=" + mdb + "; User Id=admin; Password=;",
	"Provider=" + this.provider + "; Data source=" + tmp + "; User Id=admin; Password=;");
	if (fso.FileExists(mdb)) fso.DeleteFile(mdb);
	fso.MoveFile(tmp, mdb);
  	fso = null;
  	engine = null;
};

// GetDate()
asplib.database.MsJetConnection.prototype.GetDate = function() {
	trace("asplib.database.MsJetConnection.GetDate");
	return "Now()";
};

%>

