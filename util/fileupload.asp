<%

/* fileupload.asp - by Thomas Kjoernes <thomas@ipv.no> */

	global = this;

	if (!this.asplib) asplib = {};
	if (!this.asplib.util) asplib.util = {};

// FileUpload()
asplib.util.FileUpload = function(folder) {
	this.__name = "asplib.util.FileUpload";
	var convertTypes = {
		"image/pjpeg" : "image/jpeg",
		"application/msword" : "application/vnd.ms-word",
		"application/msaccess" : "application/vnd.ms-access"
	};
	var vbCrLf = String.fromCharCode(13,10);
	this.utf8 = (Response.Charset === "UTF-8");
	this.scriptTimeout = 900; // 15 minutes
	this.count = 0;
	this.folder = (folder) ? folder.toString().replace(/\//g, "\\") : Server.MapPath("/");
	this.items = [];
	this.files = [];
	this.hasItems = false;
	this.hasFiles = false;
	this.onSave = function() { };
	this.onSaveComplete = function() { };
	this.onSaveError = function() { };
	// getBitmapInfo()
	function getBitmapInfo(formData, contentStart, contentLength) {
		var info = {};
		var unicodeToAnsi = {
			8364:	128,	129:	129,	8218:	130,	402:	131,
			8222:	132,	8230:	133,	8224:	134,	8225:	135,
			710:	136,	8240:	137,	352:	138,	8249:	139,
			338:	140,	141:	141,	381:	142,	143:	143,
			144:	144,	8216:	145,	8217:	146,	8220:	147,
			8221:	148,	8226:	149,	8211:	150,	8212:	151,
			732:	152,	8482:	153,	353:	154,	8250:	155,
			339:	156,	157:	157,	382:	158,	376:	159
		};
		// getString
		function getString(position, length) {
			return formData.substr(contentStart + position, length);
		}
		// getByte
		function getByte(position) {
			var charCode = formData.charCodeAt(contentStart + position);
			if (charCode > 255 && unicodeToAnsi[charCode]) charCode = unicodeToAnsi[charCode];
			return charCode;
		}
		// getInt
		function getInt(position) {
			return getByte(position) + (getByte(position+1)<<8);
		}
		// getLong
		function getLong(position) {
			return getByte(position) + (getByte(position+1)<<8) + (getByte(position+2)<<16) + (getByte(position+3)<<24);
		}
		// getIntBigEndian
		function getIntBigEndian(position) {
			return getByte(position+1) + (getByte(position)<<8);
		}
		// getLongBigEndian
		function getLongBigEndian(position) {
			return (getByte(position+3)<<24) + (getByte(position+2)<<16) + (getByte(position+1)<<8) + getByte(position);
		}
		// JPEG
		if (getString(0, 2) === "\xFF\xD8" && getString(6, 4) === "JFIF" || getString(6, 4) === "Exif") {
			var i = 2;
			info.mimeType = "image/jpeg";
			while (i < contentLength) {
				var c = getByte(i+1);
				var l = getIntBigEndian(i+2) + 2;
				switch (c) {
					case 0xC0 :   // SOF Baseline
					case 0xC2 : { // SOF Progressive
						info.progressive = (c === 194);
						info.width = getIntBigEndian(i+7);
						info.height = getIntBigEndian(i+5);
						break;
					}
					case 0xE0 : { // APP0 JFIF
						info.signature = getString(i+4, 4);
						break;
					}
					case 0xE1 : { // APP1 Exif Attribute Information
						info.exif = {};
						info.signature = getString(i+4, 4); // skip last 2 bytes
						var j = i + 10; // TIFF header position
						var k = j;
						if (getString(j, 4) === "II*\x00" || getString(j, 4) === "MM\x00*") {
							var littleEndian =  (getString(j, 1) === "I");
							var p = (littleEndian) ? getLong(j+4) : getLongBigEndian(j+4);
							var n = (littleEndian) ? getInt(j+p) : getIntBigEndian(j+p);
							j = j + p + 2;
							while (n > 0 && i < contentLength) {
								var tag = (littleEndian) ? getInt(j) : getIntBigEndian(j);
								var type = (littleEndian) ? getInt(j+2) : getIntBigEndian(j+2);
								var length = (littleEndian) ? getInt(j+4) : getIntBigEndian(j+4);
								var value = (littleEndian) ? getLong(j+8) : getLongBigEndian(j+8);
								// all offsets are relative to TIFF header (k)
								switch (tag) {
									case 0x010e : info.exif["ImageDescription"] = getString(k+value, length); break;
									case 0x010f : info.exif["Make"] = getString(k+value, length); break;
									case 0x0110 : info.exif["Model"] = getString(k+value, length); break;
									case 0x0112 : info.exif["Orientation"] = value; break;
									case 0x011a : info.exif["XResolution"] = length + "/" + value; break;
									case 0x011b : info.exif["YResolution"] = length + "/" + value; break;
									case 0x0128 : info.exif["ResolutionUnit"] = [null, "in", "cm"][value]; break;
									case 0x0131 : info.exif["Software"] = getString(k+value, length); break;
									case 0x0132 : info.exif["DateTime"] = getString(k+value, length).replace(/(\d{4})\:(\d{2})\:(\d{2})/, '$1-$2-$3'); break;
									case 0x013b : info.exif["Artist"] = getString(k+value, length); break;
									case 0x013c : info.exif["HostComputer"] = getString(k+value, length); break;
									case 0x013e : info.exif["WhitePoint"] = length + "/" + value; break;
									case 0x013f : info.exif["PrimaryChromaticities "] = value; break;
									case 0x0211 : info.exif["YCbCrCoefficients"] = value; break;
									case 0x0213 : info.exif["YCbCrPositioning"] = value; break;
									case 0x0214 : info.exif["ReferenceBlackWhite"] = value; break;
									case 0x8298 : info.exif["Copyright"] = getString(k+value, length); break;
									case 0x829a : info.exif["ExposureTime"] = value; break;
									case 0x829d : info.exif["FNumber"] = value; break;
									case 0x8822 : info.exif["ExposureProgram"] = value; break;
									case 0x8827 : info.exif["ISOSpeedRatings"] = value; break;
									case 0x9000 : info.exif["ExifVersion"] = value; break;
									case 0x9003 : info.exif["DateTimeOriginal"] = getString(k+value, length).replace(/(\d{4})\:(\d{2})\:(\d{2})/, '$1-$2-$3');; break;
									case 0x9004 : info.exif["DateTimeDigitized"] = getString(k+value, length).replace(/(\d{4})\:(\d{2})\:(\d{2})/, '$1-$2-$3');; break;
									case 0x9101 : info.exif["ComponentsConfiguration"] = value; break;
									case 0x9102 : info.exif["CompressedBitsPerPixel"] = value; break;
									case 0x9201 : info.exif["ShutterSpeedValue"] = value; break;
									case 0x9202 : info.exif["ApertureValue"] = value; break;
									case 0x9204 : info.exif["ExposureBiasValue"] = value; break;
									case 0x9205 : info.exif["MaxApertureValue"] = value; break;
									case 0x9206 : info.exif["SubjectDistance"] = length + "/" + value; break;
									case 0x9207 : info.exif["MeteringMode"] = value; break;
									case 0x9208 : info.exif["LightSource"] = value; break;
									case 0x9209 : info.exif["Flash"] = value; break;
									case 0x920a : info.exif["FocalLength"] = length + "/" + value; break;
									case 0x927c : info.exif["MakerNote"] = getString(k+value, length); break;
									case 0x9290 : info.exif["SubsecTime"] = value; break;
									case 0x9291 : info.exif["SubsecTimeOriginal"] = value; break;
									case 0x9292 : info.exif["SubsecTimeDigitized"] = value; break;
									case 0xa000 : info.exif["FlashPixVersion"] = value; break;
									case 0xa001 : info.exif["ColorSpace"] = value; break;
									case 0xa002 : info.exif["ExifImageWidth"] = value; break;
									case 0xa003 : info.exif["ExifImageHeight"] = value; break;
									case 0xa004 : info.exif["RelatedSoundFile"] = value; break;
									case 0xa005 : info.exif["ExifInteroperabilityOffset"] = value; break;
									case 0xa20e : info.exif["FocalPlaneXResolution"] = length + "/" + value; break;
									case 0xa20f : info.exif["FocalPlaneYResolution"] = length + "/" + value; break;
									case 0xa215 : info.exif["ExposureIndex"] = value; break;
									case 0xa217 : info.exif["SensingMethod"] = value; break;
									case 0xa300 : info.exif["FileSource"] = getString(k+value, length); break;
									case 0xa301 : info.exif["SceneType"] = value; break;
									case 0xa302 : info.exif["CFAPattern"] = value; break;
								}
								j += 12; // IFD entry size
								n--;
							}
						}
						break;
					}
					case 0xE2 : { // APP2 FlashPix Extension Data
						info.signature = getString(i+4, 4);
						break;
					}
					case 0xFE : { // Comment
						info.comment = getString(i+4, l);
						break;
					}
				}
				i += l;
			}
		}
		// TIFF
		if (getString(0, 4) === "II*\x00" || getString(0, 4) === "MM\x00*") {
			info.mimeType = "image/tiff";
			var littleEndian =  (getString(0, 1) === "I");
			var i = (littleEndian) ? getLong(4) : getLongBigEndian(4);
			var n = (littleEndian) ? getInt(i) : getIntBigEndian(i);
			while (n > 0 && i < contentLength) {
				var tag = (littleEndian) ? getInt(i+2) : getIntBigEndian(i+2);
				var value = (littleEndian) ? getLong(i+10) : getLongBigEndian(i+10);
				switch (tag) {
					case 256: info.width = value; break;
					case 257: info.height = value; break;
				}
				i += 12; // IFD entry size
				n--;
			}
		}
		// PNG
		if (getString(1, 3) === "PNG") {
			info.mimeType = "image/png";
			info.width = getIntBigEndian(18);
			info.height = getIntBigEndian(22);
		}
		// GIF
		if (getString(0, 3) === "GIF") {
			info.mimeType = "image/gif";
			info.width = getInt(6);
			info.height = getInt(8);
		}
		// BMP
		if (getString(0, 2) === "BM") {
			info.mimeType = "image/x-bmp";
			info.width = getInt(18);
			info.height = getInt(22);
		}
		return info;
	}
	// Open()
	this.Open = function() {
		Server.ScriptTimeout = this.scriptTimeout;
		var requestMethod = Request.ServerVariables("REQUEST_METHOD").Item();
		if (requestMethod.toUpperCase() !== "POST") return false;
		if (this.folder.charAt(this.folder.length-1) !== "\\") this.folder += "\\";
		var binaryData = Request.BinaryRead(Request.TotalBytes);
		this.streamForm = new ActiveXObject("ADODB.Stream");
		this.streamForm.Type = 1; // adTypeBinary=1
		this.streamForm.Open();
		this.streamForm.Write(binaryData);
		this.streamForm.Position = 0;
		this.streamForm.Type = 2; // adTypeText=2
		this.streamForm.Charset = "ISO-8859-1";
		var formData = this.streamForm.ReadText();
		var boundaryStart = formData.indexOf(vbCrLf)+2;
		var boundaryTag = formData.substring(0, boundaryStart-2);
		var boundaryMax = formData.indexOf(boundaryTag + "--");
		var boundaryEnd = 0;
		// check each message part
		while (boundaryEnd < boundaryMax) {
			boundaryEnd = formData.indexOf(boundaryTag, boundaryStart);
			var contentStart = formData.indexOf(vbCrLf + vbCrLf, boundaryStart)+4;
			var contentEnd = boundaryEnd-2;
			var contentLength = contentEnd - contentStart;
			var contentHeader = formData.substring(boundaryStart, contentStart-4).split(vbCrLf);
			// any content present?
			if (contentLength>0) {
				this.count++;
				var contentType = null;
				var contentName = null;
				var contentFileName = null;
				// check headers
				for (var i=0; i<contentHeader.length; i++) {
					var headerName = contentHeader[i].substring(0, contentHeader[i].indexOf(":")).toLowerCase();
					var headerContent = contentHeader[i].substring(contentHeader[i].indexOf(":")+2).split(";");
					switch (headerName) {
						case "content-type" : {
							// content-type header (MIME type)
							contentType = headerContent[0];
							break;
						}
						case "content-disposition" : {
							// content-disposition header
							for (var j=0; j<headerContent.length; j++) {
								var headerParam = headerContent[j].split("=");
								switch (headerParam[0].trim()) {
									// name param?
									case "name" : {
										if (headerParam[1]) contentName = headerParam[1].replace(/\"/g, "").toLowerCase(); // "
										break;
									}
									// fileName?
									case "filename" : {
										if (headerParam[1]) contentFileName = headerParam[1].replace(/\"/g, ""); // "
										if (this.utf8 && global.Utf8) contentFileName = Utf8.decode(contentFileName);
										break;
									}
								}
							}
							break;
						}
					}
				}
				if (contentFileName) {
					// File
					var info = getBitmapInfo(formData, contentStart, contentLength);
					var i = contentFileName.lastIndexOf("\\");
					var path = (i>1) ? contentFileName.substring(0,i+1) : "";
					var fileName = (i>1) ? contentFileName.substring(i+1) : contentFileName;
					this.files[contentName] = {
						position: contentStart,
						fileSize: contentLength,
						fileName: decodeURIComponent(fileName||""),
						originalFileName: fileName,
						remotePath: path,
						remoteMimeType: convertTypes[contentType] || contentType
					};
					for (var o in info) {
						this.files[contentName][o] = info[o];
					}
					this.hasFiles = true;
				} else {
					// Item
					var contentData = formData.substring(contentStart, contentEnd);
					if (this.items[contentName]) {
						this.items[contentName].value += ", " + contentData;
					} else {
						this.items[contentName] = {
							value: contentData,
							toString: function() { return this.value.toString(); }
						};
					}
					this.hasItems = true;
				}
			}
			// find next message part
			boundaryStart = formData.indexOf(vbCrLf, boundaryEnd)+2;
		}
		return this.count;
	};
	// Close()
	this.Close = function() {
		if (this.streamForm) {
			this.streamForm.Close();
			this.streamForm = null;
		}
	};
	// Form()
	this.Form = function(name) {
		var name = name.toLowerCase();
		return this.items[name];
	};
	// File()
	this.File = function(name) {
		var name = name.toLowerCase();
		return this.files[name];
	};
	// Save()
	this.Save = function(name, path, fileName) {
		if (!name) return;
		var name = name.toLowerCase();
		var item = this.files[name];
		if (item) {
			item.path = (path) ? path : this.folder;
			if (item.path.charAt(item.path.length-1) !== "\\") item.path += "\\";
			if (fileName) item.fileName = fileName;
			if (typeof this.onSave === "function" && this.onSave(item, this) !== false) {
				try {
					var streamFile = new ActiveXObject("ADODB.Stream");
					streamFile.Type = 1; // adTypeBinary=1
					streamFile.Open();
				//	streamFile.Charset = this.streamForm.Charset;
					this.streamForm.Position = item.position;
					this.streamForm.CopyTo(streamFile, item.fileSize);
					streamFile.SaveToFile(item.path + item.fileName, 2); // adSaveCreateOverWrite=2
					streamFile.Close();
					streamFile = null;
					if (typeof this.onSaveComplete === "function") this.onSaveComplete(item, this);
					return true;
				} catch(e) {
					if (typeof this.onSaveError === "function") this.onSaveError(item, this, e);
				}
			}
		} else {
			return false;
		}
	};
	// SaveAll();
	this.SaveAll = function() {
		for (var i in this.files) {
			this.Save(i);
		}
	};
	// MapPath()
	this.MapPath = function(url) {
		this.folder = Server.MapPath(url);
	};
}

%>