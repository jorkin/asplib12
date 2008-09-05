<%

asplib.mail.Message = function(from, to, subject, textBody, htmlBody, cc, bcc, attachments, smtpClient) {
	this.__name = "asplib.mail.Message";
	this.totalBytes = 0;
	this.usageCount = 0;
	this.subject = subject || "";
	this.textBody = textBody || null;
	this.htmlBody = htmlBody || null;
	this.url = null;
	this.fields = {};
	this.from = from || new asplib.mail.Address();
	this.to = to || new asplib.mail.AddressList();
	this.cc = cc || new asplib.mail.AddressList();
	this.bcc = bcc || new asplib.mail.AddressList();
	this.attachments = attachments || new asplib.mail.AttachmentList();
	if (smtpClient) this.smtpClient = smtpClient;
};

asplib.mail.Message.prototype.__init = function() {
	if (!this.smtpClient) {
		if (!asplib.mail.cache.smtpClient) {
			asplib.mail.cache.smtpClient = new asplib.mail.SmtpClient();
		}
		this.smtpClient = asplib.mail.cache.smtpClient;
	}
	if (typeof this.from === "string") this.from = new asplib.mail.Address(this.from);
	if (typeof this.to === "string") this.to = new asplib.mail.AddressList(this.to);
	if (typeof this.cc === "string") this.cc = new asplib.mail.AddressList(this.cc);
	if (typeof this.bcc === "string") this.bcc = new asplib.mail.AddressList(this.bcc);
};

asplib.mail.Message.prototype.Import = function(url) {
	this.url = url;
}

asplib.mail.Message.prototype.Clear = function() {
	this.subject = "";
	this.textBody = "";
	this.textHtml = "";
	this.url = null;
	this.fields = {};
	this.from.Clear();
	this.to.Clear();
	this.cc.Clear();
	this.bcc.Clear();
	this.attachments.Clear();
};

asplib.mail.Message.prototype.Build = function() {
	this.__init();
	return this.smtpClient.BuildMessage(this);
};

asplib.mail.Message.prototype.Send = function() {
	this.__init();
	return this.smtpClient.SendMessage(this);
};

asplib.mail.Message.prototype.SaveToFile = function(path) {
	this.__init();
	return this.smtpClient.SaveMessage(this, path);
};

%>