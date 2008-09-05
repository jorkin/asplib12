<%

/* imagehelper.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.util) asplib.util = {};

	var PROVIDER_IMGWRITER = 1;
	var PROVIDER_ASPIMAGE = 2;
	var PROVIDER_ASPJPEG = 3;

// ImageHelper()
asplib.util.ImageHelper = function(provider) {
	this.__name = "asplib.util.ImageHelper";
	this.provider = 0;
	this.activeX = null;
	this.imageWidth = 0;
	this.imageHeight = 0;
	this.quality = 80;
	this.features = [
		"MergeWithFile",
		"SaveThumbToFile",
		"SetFont",
		"AddText",
 		"FlipHorizontal",
		"FlipVertical",
		"RotateLeft",
		"RotateRight",
		"Crop",
		"CropToFit",
		"Resize",
		"ResizeToFit",
		"ResizeByPercentage",
		"GrayScale",
		"Brighten",
		"Darken",
		"Sharpen"
	];
	this.providers = [];
	this.providers[PROVIDER_IMGWRITER] = { progId: "SoftArtisans.ImageGen", features: ["Blur","Rotate"] };
	this.providers[PROVIDER_ASPIMAGE] = { progId: "AspImage.Image", features: ["Blur","Rotate"] };
	this.providers[PROVIDER_ASPJPEG] = { progId: "Persits.Jpeg", features: ["AddTextEx"] };
	for (var i in this.providers) {
		try  {
			if (!provider || i==provider) {
				var progId = this.providers[i].progId;
				var features = this.providers[i].features;
				this.activeX = new ActiveXObject(progId);
				this.name = progId;
				this.progId = progId;
				this.provider = parseInt(i);
				this.features = this.features.concat(features);
				break;
			}
		} catch(e) { }
	}
	return this.provider;
}

asplib.util.ImageHelper.prototype.ParseColor = function(color, reverse) {
	var colors = {
		"black":	"#000000",
		"silver":	"#C0C0C0",
		"gray":		"#808080",
		"white":	"#FFFFFF",
		"maroon":	"#800000",
		"red":		"#FF0000",
		"purple":	"#800080",
		"fuchsia":	"#FF00FF",
		"green":	"#008000",
		"lime":		"#00FF00",
		"olive":	"#808000",
		"yellow":	"#FFFF00",
		"navy":		"#000080",
		"blue":		"#0000FF",
		"teal":		"#008080",
		"aqua":		"#00FFFF",
		"orange":	"#FFA500"
	}
	var temp = color;
	if (typeof color === "undefined") {
		temp = "#000000";
	} else {
		temp = colors[color];
	}
	if (temp.charAt(0) === "#") {
		if (temp.length === 4) {
			var r = parseInt("0x0" + temp.substring(1,2) + temp.substring(1,2), 16);
			var g = parseInt("0x0" + temp.substring(2,3) + temp.substring(2,3), 16);
			var b = parseInt("0x0" + temp.substring(3,4) + temp.substring(3,4), 16);
			return (reverse) ? b * 65536 + g * 256 + r : r * 65536 + g * 256 + b;
		}
		if (temp.length === 7) {
			var r = parseInt("0x0" + temp.substring(1,3), 16);
			var g = parseInt("0x0" + temp.substring(3,5), 16);
			var b = parseInt("0x0" + temp.substring(5,7), 16);
			return (reverse) ? b * 65536 + g * 256 + r : r * 65536 + g * 256 + b;
		}
	}
};

// New()
asplib.util.ImageHelper.prototype.New = function(width, height, color) {
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.New(width, height, this.ParseColor(color));
			this.activeX.Canvas.Font.BkColor = this.ParseColor(color);
			this.open = true;
			break;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.imageWidth = width;
			this.imageHeight = height;
			this.activeX.MaxX = width;
			this.activeX.MaxY = height;
			this.activeX.BackgroundColor = this.ParseColor(color, true);
			this.activeX.FillRect(0, 0, width, height);
			break;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.CreateImage(width, height, this.ParseColor(color, true));
			this.open = true;
			break;
		}
	}
	return this.open;
};

// LoadFromFile()
asplib.util.ImageHelper.prototype.LoadFromFile = function(path, merge, opacity) {
	this.open = false;
	this.path = path || this.path;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			try {
				this.activeX.Open(this.path);
				this.imageWidth = this.activeX.OriginalWidth;
				this.imageHeight = this.activeX.OriginalHeight;
				this.open = true;
			} catch(e) { }
			break;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			var stream = new ActiveXObject("ADODB.Stream");
			stream.Type = 1; // 1=adTypeBinary
			stream.Open();
			try {
				stream.LoadFromFile(this.path);
				stream.Position = 0;
				stream.Type = 2;
				stream.CharSet = "ISO-8859-1";
			} catch(e) { }
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
				if (stream.Size===0) return "";
				stream.Position = position;
				return stream.ReadText(length);
			}
			// getByte
			function getByte(position) {
				if (stream.Size===0) return 0;
				stream.Position = position;
				var charCode = stream.ReadText(1).charCodeAt(0);
				return (charCode>255 && unicodeToAnsi[charCode]) ? unicodeToAnsi[charCode] : charCode;
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
			if (getString(6, 4) === "JFIF" || getString(6, 4) === "Exif") {
				var i = 2;
				while (i < contentLength) {
					var c = getByte(i+1);
					var l = getIntBigEndian(i+2) + 2;
					if (c === 192 || c === 194) {
						info.mimeType = (c === 192) ? "image/jpeg" : "image/pjpeg";
						info.width = getIntBigEndian(i+7);
						info.height = getIntBigEndian(i+5);
					}
					i += l;
				}
			}
			// TIFF
			if (getString(0, 4) === "II*\x00" || getString(0, 4) === "MM\x00*") {
				var littleEndian =  (getString(0, 1) === "I");
				var i = littleEndian ? getLong(4) : getLongBigEndian(4);
				var n = littleEndian ? getInt(i) : getIntBigEndian(i);
				while (n > 0 && i < contentLength) {
					var c = littleEndian ? getInt(i+2) : getIntBigEndian(i+2);
					var v = littleEndian ? getLong(i+10) : getLongBigEndian(i+10);
					switch (c) {
						case 256: this.imageWidth = v; break;
						case 257: this.imageWidth = v; break;
					}
					i += 12; // IFD entry size
					n--;
				}
			}
			// PNG
			if (getString(1, 3) === "PNG") {
				this.imageWidth = getIntBigEndian(18);
				this.imageHeight = getIntBigEndian(22);
			}
			// GIF
			if (getString(0, 3) === "GIF") {
				this.imageWidth = getInt(6);
				this.imageHeight = getInt(8);
			}
			// BMP
			if (getString(0, 2) === "BM") {
				this.imageWidth = getInt(18);
				this.imageHeight = getInt(22);
			}
			try {
				this.activeX.LoadImage(path);
				this.open = true;
			} catch(e) { }
			break;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			try {
				this.activeX.LoadImage(this.path);
				this.imageWidth = this.activeX.Width;
				this.imageHeight = this.activeX.Height;
				this.open = true;
			} catch(e) { }
			break;
		}
	}
	return this.open;
};

// SaveToFile()
asplib.util.ImageHelper.prototype.SaveToFile = function(path, quality) {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Quality = quality || this.quality;
			this.activeX.Save(path||this.path);
			this.imageWidth = this.activeX.OriginalWidth;
			this.imageHeight = this.activeX.OriginalHeight;
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.ImageFormat = 1; // 1 = JPEG
			this.activeX.JPEGQuality = quality || this.quality;
			this.activeX.FileName = path||this.path;
			this.activeX.SaveImage();
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.ImageQuality = quality || this.quality;
			this.activeX.SaveImage(0, 3, path||this.path); // method=0 (saiFile), type=3 (saiJPG)
			this.imageWidth = this.activeX.Width;
			this.imageHeight = this.activeX.Height;
			return true;
		}
	}
};

// MergeWithFile()
asplib.util.ImageHelper.prototype.MergeWithFile = function(path, x, y) {
	if (!this.open) return false;
	if (!x) x = 0;
	if (!y) y = 0;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			var jpegActiveX = new ActiveXObject(this.progId);
			jpegActiveX.Open(path);
			this.activeX.DrawImage(x, y, jpegActiveX);
			jpegActiveX = null;
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.AddImageTransparent(path, x, y, 0x00FF00);
			this.activeX.CropImage(0, 0, this.imageWidth, this.imageHeight);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.MergeWithImage(x, y, path, 1, 0x00FF00);
			return true;
		}
	}
};

// SaveThumbToFile()
asplib.util.ImageHelper.prototype.SaveThumbToFile = function(path, width, height, color) {
	if (!this.open) return false;
	if (!width) width = 128;
	if (!height) height = 128;
	if (!color) color = "#FFFFFF";
	this.ResizeToFit(width, height);
	var mergeX = Math.floor((width - imghlp.imageWidth) / 2);
	var mergeY = Math.floor((height - imghlp.imageHeight) / 2);
	imghlp.SaveToFile(path, 100);
	imghlp.New(width, height, color);
	imghlp.MergeWithFile(path, mergeX, mergeY);
	imghlp.SaveToFile(path, 95);
};

// SetFont()
asplib.util.ImageHelper.prototype.SetFont = function(fontName, fontSize, color, bold, italic, underline) {
	if (typeof fontName === "undefined") fontName = "Arial";
	if (typeof fontSize === "undefined") fontSize = 10;
	if (typeof smooth === "undefined") smooth = false;
	if (typeof bold === "undefined") bold = false;
	if (typeof italic === "undefined") italic = false;
	if (typeof underline === "undefined") underline = false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Canvas.Font.Family = fontName;
			this.activeX.Canvas.Font.Size = fontSize;
			this.activeX.Canvas.Font.Color = this.ParseColor(color);
			this.activeX.Canvas.Font.BkMode = "Opaque";
			this.activeX.Canvas.Font.Quality = 4; // AntiAliased
			if (bold) this.activeX.Canvas.Font.Bold = true;
			if (italic) this.activeX.Canvas.Font.Italic = true;
			if (underline) this.activeX.Canvas.Font.Underlined = true;
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.FontName = fontName;
			this.activeX.FontSize = Math.floor(fontSize/1.63);
			this.activeX.FontColor = this.ParseColor(color, true);
			if (bold) this.activeX.Bold = true;
			if (italic) this.activeX.Italic = true;
			if (underline) this.activeX.Underline = true;
			this.activeX.AntiAliasText = true;
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.Font.Name = fontName;
			this.activeX.Font.Height = fontSize;
			this.activeX.Font.Color = this.ParseColor(color, true);
			this.activeX.AntiAliasFactor = 4; // 4=saiAntiAliasBest
			if (bold) this.activeX.Font.Weight = 700; // 700=bold
			if (italic) this.activeX.Font.Italic = true;
			if (underline) this.activeX.Font.Underline = true;
			return true;
		}
	}
};

// AddText()
asplib.util.ImageHelper.prototype.AddText = function(text, x, y) {
	if (!this.open) return false;
	if (typeof x === "undefined") x = 0;
	if (typeof y === "undefined") y = 0;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Canvas.PrintText(x, y, text);
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.TextOut(text, x, y, false);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			var width = Math.min(this.activeX.Width, this.activeX.TextWidth(text));
			var height = Math.min(this.activeX.Height, this.activeX.TextHeight(text));
			this.activeX.DrawTextOnImage(x, y, width, height, text);
			return true;
		}
	}
};

// AddTextEx()
asplib.util.ImageHelper.prototype.AddTextEx = function(text, x, y, fontFile) {
	if (!this.open) return false;
	if (typeof x === "undefined") x = 0;
	if (typeof y === "undefined") y = 0;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Canvas.PrintTextEx(text, x, y, fontFile);
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			return false;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			return false;
		}
	}
};
// GrayScale ()
asplib.util.ImageHelper.prototype.GrayScale  = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Grayscale(1); // method=1
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.CreateGrayScale();
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.ColorResolution = 1; // 1=saiGrayScale
			return false;
		}
	}
};

// Sharpen()
asplib.util.ImageHelper.prototype.Sharpen = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Sharpen(1, 125); // radius=1px, amount=125%
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.Sharpen(1); // times=1
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.SharpenImage(25); // amount=25
			return true;
		}
	}
};

// Blur()
asplib.util.ImageHelper.prototype.Blur = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			return false;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.Blur(1); // times=1
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.BlurImage(100); // amount=100
			return true;
		}
	}
};

// Brighten()
asplib.util.ImageHelper.prototype.Brighten = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Adjust(1, 0.025); // function=1 (brightness), level=0.025
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.BrightenImage(5); // degree=5
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.ChangeBrightness(1.05); // level=1.05
			return true;
		}
	}
};

// Darken()
asplib.util.ImageHelper.prototype.Darken = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Adjust(1, -0.05); // function=1 (brightness), level=-0.05
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.DarkenImage(20); // degree=20
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.ChangeBrightness(0.80); // level=0.8
			return true;
		}
	}
};

// FlipHorizontal()
asplib.util.ImageHelper.prototype.FlipHorizontal = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.FlipH();
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.FlipImage(1);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.FlipImage(0);
			return true;
		}
	}
};

// FlipVertical()
asplib.util.ImageHelper.prototype.FlipVertical = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.FlipV();
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.FlipImage(2);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.FlipImage(1);
			return true;
		}
	}
};

// Rotate()
asplib.util.ImageHelper.prototype.Rotate = function(degrees) {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			return false;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.RotateImage(360-degrees);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.RotateImage(degrees, 0x00FFFFFF);
			return true;
		}
	}
};

// RotateLeft()
asplib.util.ImageHelper.prototype.RotateLeft = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.RotateL();
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.RotateImage(90);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.RotateImage(270, 0x00FFFFFF);
			return true;
		}
	}
};

// RotateRright()
asplib.util.ImageHelper.prototype.RotateRight = function() {
	if (!this.open) return false;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.RotateR();
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.RotateImage(-90);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.RotateImage(90, 0x00FFFFFF);
			return true;
		}
	}
};

// Crop()
asplib.util.ImageHelper.prototype.Crop = function(x, y, width, height) {
	if (!this.open) return false;
	this.imageWidth = width;
	this.imageHeight = height;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			this.activeX.Crop(x, y, x+width, y+height);
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.CropImage(x, y, width, height);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.CropImage(x, y, width, height);
			return true;
		}
	}
};

// Resize()
asplib.util.ImageHelper.prototype.Resize = function(width, height) {
	if (!this.open) return false;
	this.imageWidth = width;
	this.imageHeight = height;
	switch (this.provider) {
		// AspJpeg
		case PROVIDER_ASPJPEG: {
			if (width) this.activeX.Width = width;
			if (height) this.activeX.Height = height;
			return true;
		}
		// AspImage
		case PROVIDER_ASPIMAGE: {
			this.activeX.ResizeR(width, height);
			return true;
		}
		// ImgWriter
		case PROVIDER_IMGWRITER: {
			this.activeX.ResizeImage(width, height, 3, 0); // algorithm=0 (bilinear), proportinal=0
			return true;
		}
	}
};

// ResizeToFit()
asplib.util.ImageHelper.prototype.ResizeToFit = function(width, height) {
	var targetAspect = width / height;
	var sourceAspect = this.imageWidth / this.imageHeight;
	var newWidth = 0;
	var newHeight = 0;
	var scale = 1;
	if (targetAspect < sourceAspect) {
		scale = this.imageWidth / width;
	} else {
		scale = this.imageHeight / height;
	}
	if (scale > 1) {
		newWidth = Math.floor(this.imageWidth / scale);
		newHeight = Math.floor(this.imageHeight / scale);
		this.Resize(newWidth, newHeight);
	}
};

// ResizeByPercentage()
asplib.util.ImageHelper.prototype.ResizeByPercentage = function(percent) {
	var newWidth = Math.floor(this.imageWidth * percent / 100);
	var newHeight = Math.floor(this.imageHeight * percent / 100);
	this.Resize(newWidth, newHeight);
};

// ResizeAndCropToFit()
asplib.util.ImageHelper.prototype.ResizeAndCropToFit = function(width, height, noResize) {
	var targetAspect = width / height;
	var sourceAspect = this.imageWidth / this.imageHeight;
	var newWidth = 0;
	var newHeight = 0;
	var cropX = 0;
	var cropY = 0;
	if (targetAspect > sourceAspect) {
		newWidth = this.imageWidth;
		newHeight = Math.floor(sourceAspect / targetAspect * this.imageHeight);
		cropX = 0;
		cropY = Math.floor((this.imageHeight - newHeight) / 2);
	}
	if (targetAspect < sourceAspect) {
		newWidth = Math.floor(targetAspect / sourceAspect * this.imageWidth);
		newHeight = this.imageHeight;
		cropX = Math.floor((this.imageWidth - newWidth) / 2);
		cropY = 0;
	}
	if (newHeight || newWidth) {
		this.Crop(cropX, cropY, newWidth, newHeight);
	}
	if (!noResize) {
		this.Resize(width, height);
	}
};

%>
