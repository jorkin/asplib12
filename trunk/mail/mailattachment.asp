<%

asplib.mail.Attachment = function(url, displayName) {
	this.__name = "asplib.mail.Attachment";
	this.url = (url||"").replace("\\", "/");
	if (this.url.indexOf("://") === -1) this.url = "file://" + this.url;
	this.displayName = displayName || url.substring(url.lastIndexOf("/")+1);
};

asplib.mail.Attachment.prototype.toString = function() {
	return this.url.toString();
};

asplib.mail.AttachmentList = function() {
	this.__name = "asplib.mail.AttachmentList";
	this.length = 0;
	for (var i=0; i<arguments.length; i++) {
		this.Add(arguments[i]);
	}
};

asplib.mail.AttachmentList.prototype.Add = function(mailAttachment, displayName) {
	if (typeof mailAttachment === "string") {
		mailAttachment = new asplib.mail.Attachment(mailAttachment, displayName);
	} else {
		if (displayName) mailAttachment.displayName = displayName;
	}
	if (mailAttachment.url) {
		var i = this.length++;
		this[i] = mailAttachment;
	}
}

asplib.mail.AttachmentList.prototype.Clear = function() {
	for (var i=0; i<this.length; i++) {
		delete this[i];
	}
	this.length = 0;
};

asplib.mail.AttachmentList.prototype.toString = function() {
	var temp = [];
	for (var i=0; i<this.length; i++) {
		temp.push(this[i].toString());
	}
	return temp.join("; ");
}


%>