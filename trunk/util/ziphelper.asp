<%

/* ziphelper.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.util) asplib.util = {};

	var PROVIDER_XSTANDARD_ZIP = 1;
	var PROVIDER_SOFTCOMPLEX_ZIP = 2;

// ZipHelper()
asplib.util.ZipHelper = function(provider) {
	this.__name = "asplib.util.ZipHelper";
	this.activeX = null;
	this.providers = [];
	this.features = [];
	this.providers[PROVIDER_XSTANDARD_ZIP] = { progId: "XStandard.Zip", features: [] };
	this.providers[PROVIDER_SOFTCOMPLEX_ZIP] = { progId: "SoftComplex.Zip", features: [] };
	for (var i in this.providers) {
		if (!provider || i==provider) {
			var progId = this.providers[i].progId;
			var features = this.providers[i].features;
			try  {
				this.activeX = new ActiveXObject(progId);
				this.progId = progId;
				this.features = this.features.concat(features);
				this.provider = parseInt(i);
				break;
			} catch(e) {
				Response.Write(e.description);
			}
		}
	}
	return this.provider;
}

// Open()
asplib.util.ZipHelper.prototype.LoadFromFile = function(path) {
	this.folders = [];
	this.length = 0;
	switch (this.provider) {
		case PROVIDER_XSTANDARD_ZIP : {
			var list = this.activeX.Contents(path);
			for (var i=0; i<list.Count; i++) {
				var item = list.Item(i);
				switch (item.Type) {
					case 1 : {
						this.folders.push(item.Path + item.Name);
						break;
					}
					case 2 : {
						this[this.length++] = {
							fileName: item.Name,
							fileSize: item.Size,
							path: item.Path,
							dateCreated: parseDate(item.Date),
							toString: function() {
								return this.fileName;
							}
						};
						break;
					}
				}
			}
			break;
		}
		case PROVIDER_SOFTCOMPLEX_ZIP : {
			this.activeX.Open(path);
			this.activeX.Read();
			var list = this.activeX.FileList;
			for (var i=0; i<this.activeX.Count; i++) {
				var name = this.activeX.FileName(i);
				var size = this.activeX.FileSize(i);
				var path = this.activeX.FilePathName(i);
				if (name) {
					this[this.length++] = {
						fileName: name,
						fileSize: size,
						path: path,
						dateCreated: parseDate(this.activeX.FileDateTime(i)),
						toString: function() {
							return this.fileName;
						}
					};
				} else {
					this.folders.push(path);
				}
			}
			break;
		}
	}
	this.folders.sort();
};

// New()
asplib.util.ZipHelper.prototype.New = function(path) {
	switch (this.provider) {
		case PROVIDER_XSTANDARD_ZIP : {
			break;
		}
		case PROVIDER_SOFTCOMPLEX_ZIP : {
			break;
		}
	}
};

// Add()
asplib.util.ZipHelper.prototype.Add = function(path) {
	switch (this.provider) {
		case PROVIDER_XSTANDARD_ZIP : {
			break;
		}
		case PROVIDER_SOFTCOMPLEX_ZIP : {
			break;
		}
	}
};

// Remove()
asplib.util.ZipHelper.prototype.Remove = function(path) {
	switch (this.provider) {
		case PROVIDER_XSTANDARD_ZIP : {
			break;
		}
		case PROVIDER_SOFTCOMPLEX_ZIP : {
			break;
		}
	}
};

%>
