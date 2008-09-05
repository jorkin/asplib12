<%
	var SMTP_STATUS_SUCCESS = 0;
	var SMTP_STATUS_INVALID_PROVIDER = -1;
	var SMTP_STATUS_MISSING_PROVIDER = -2;
	var SMTP_STATUS_MISSING_MESSAGE = -3;
	var SMTP_STATUS_MISSING_RECIPIENT = -4;

// SmtpClient()
asplib.mail.SmtpClient = function(config) {
	this.__name = "asplib.mail.SmtpClient";
	this.config = config || asplib.mail.config;
	this.status = 0;
	this.statusText = "";
	if (!this.config.provider) this.config.provider = PROVIDER_CDOSYS;
};

asplib.mail.SmtpClient.prototype.ClearMessage = function(mailMessage) {
	if (mailMessage) mailMessage.Clear();
	this.activeX = null;
};

asplib.mail.SmtpClient.prototype.BuildMessage = function(mailMessage) {
	if (!mailMessage) {
		this.status = SMTP_STATUS_MISSING_MESSAGE;
	} else {
		switch (this.config.provider) {
			case PROVIDER_CDONTS : {
				break;
			}
			case PROVIDER_CDOSYS : {
				if (!this.activeX) {
					try {
						var config = new ActiveXObject("CDO.Configuration");
						this.status = 1;
						if (!this.config.smtpServer) {
							config.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1;
							if (this.config.smtpServerPickupDirectory) {
								config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = this.config.smtpServerPickupDirectory;
							 }
						} else {
							config.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2;
							config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = this.config.smtpServer;
							config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = this.config.smtpServerPort || 25;
							config.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = this.config.smtpServerConnectionTimeout || 60;
						}
						config.Fields.Update();
						this.activeX = new ActiveXObject("CDO.Message");
						this.activeX.Configuration = config;

					} catch(e) {
						this.status = SMTP_STATUS_MISSING_PROVIDER;
						this.statusText = e.description;
					}
				}
				try {
					if (!mailMessage.to) {
						this.status = SMTP_STATUS_MISSING_RECIPIENT;
					} else {
						this.activeX.From = mailMessage.from.toString();
						this.activeX.To = mailMessage.to.toString();
						this.activeX.Cc = mailMessage.cc.toString();
						this.activeX.Bcc = mailMessage.bcc.toString();
						this.activeX.Subject = mailMessage.subject;
						if (mailMessage.attachments) {
							for (var i=0; i<mailMessage.attachments.length; i++) {
								var attachment = mailMessage.attachments[i];
								var bodyPart = this.activeX.AddAttachment(attachment.url);
								bodyPart.Fields("urn:schemas:mailheader:content-disposition") = "attachment; filename=\"" + attachment.displayName + "\"";
								bodyPart.Fields.Update();
							}
						}
						if (mailMessage.htmlBody) this.activeX.HTMLBody = mailMessage.htmlBody;
						if (mailMessage.url) {
							this.activeX.CreateMHTMLBody(mailMessage.url);
						}
						if (typeof mailMessage.fields === "object") {
							var template = new asplib.html.Template(null, mailMessage.fields);
							template.Put(this.activeX.HTMLBody);
							this.activeX.HTMLBody = template.Get();
						}
						if (mailMessage.textBody) this.activeX.TextBody = mailMessage.textBody;
						var stream = this.activeX.GetStream();
						mailMessage.totalBytes = stream.Size;
						mailMessage.usageCount++;
						this.status = SMTP_STATUS_SUCCESS;
					}
				} catch(e) {
					this.status = e.number;
					this.statusText = e.description;
				}
				break;
			}
			default : {
				this.status = SMTP_STATUS_INVALID_PROVIDER;
			}
		}
	}
	mailMessage.status = this.status;
	mailMessage.statusText = this.statusText;
	return (this.status === SMTP_STATUS_SUCCESS);
};

asplib.mail.SmtpClient.prototype.SendMessage = function(mailMessage) {
	if (this.BuildMessage(mailMessage)) {
		try {
			switch (this.config.provider) {
				case PROVIDER_CDONTS : {
					break;
				}
				case PROVIDER_CDOSYS : {
					if (this.activeX) {
						this.activeX.Send();
					}
				}
			}
		} catch(e) {
			this.status = e.number;
		}
	}
	mailMessage.status = this.status;
	mailMessage.statusText = this.statusText;
	return (this.status === SMTP_STATUS_SUCCESS);
};

asplib.mail.SmtpClient.prototype.SaveMessage = function(mailMessage, path) {
	if (this.BuildMessage(mailMessage)) {
		try {
			switch (this.config.provider) {
				case PROVIDER_CDONTS : {
					break;
				}
				case PROVIDER_CDOSYS : {
					var stream = this.activeX.GetStream();
					stream.SaveToFile(path, 2); // adSaveCreateOverwrite
				}
			}
		} catch(e) {
			this.status = e.number;
		}
	}
	mailMessage.status = this.status;
	mailMessage.statusText = this.statusText;
	return (this.status === SMTP_STATUS_SUCCESS);
};

asplib.mail.SmtpClient.prototype.GetMessage = function(mailMessage) {
	if (this.BuildMessage(mailMessage)) {
		try {
			switch (this.config.provider) {
				case PROVIDER_CDONTS : {
					break;
				}
				case PROVIDER_CDOSYS : {
					return this.activeX.GetStream().ReadText();
				}
			}
		} catch(e) { }
	}
	return "";
};

%>