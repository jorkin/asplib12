<%

/* folderhelper.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.util) asplib.util = {};

// FolderHelper()
asplib.util.FolderHelper = function(path, wildcard, onfile, onfolder, recurse, invert) {
	this.__name = "asplib.util.FolderHelper";
	this.path = (path) ? path : Server.MapPath("/");
	this.wildcard = (wildcard) ? wildcard : "*";
	this.onfile = onfile;
	this.onfolder = onfolder;
	this.recurse = recurse;
	this.invert = invert;
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	// Open()
	this.Open = function(path, wildcard, onfile, onfolder) {
		if (path) this.path = path;
		if (wildcard) this.wildcard = wildcard;
		if (onfile) this.onfile = onfile;
		if (onfolder) this.onfolder = onfolder;
		this.rxp = new RegExp("^" + this.wildcard.replace(/([\\\^\$+[\]{}.=!:(|)])/g, "\\$1").replace(/\*/g, ".*").replace(/\?/g, ".") + "$");
		if (fso.FolderExists(this.path)) {
			var folder = fso.GetFolder(this.path);
			this.__search(folder);
		}
	};
	// Close()
	this.Close = function() {
		fso = null;
	};
	// __Search()
	this.__search = function(folder) {
		for (var e = new Enumerator(folder.SubFolders); !e.atEnd(); e.moveNext()) {
			var item = e.item();
			if (typeof this.onfolder === "function") this.onfolder(item, fso);
			if (this.recurse) this.__search(item);
		}
		for (var e = new Enumerator(folder.Files); !e.atEnd(); e.moveNext()) {
			var item = e.item();
			var name = item.Name;
			var match = (this.invert) ? !name.match(this.rxp) : name.match(this.rxp);
			if (typeof this.onfile === "function" && match) this.onfile(item, fso);
		}
	}

}

%>