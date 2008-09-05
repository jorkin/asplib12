<%

asplib.mail.Address = function(mailAddress, displayName, mergeData) {
	this.__name = "asplib.mail.Address";
	this.mailAddress = (mailAddress || "").toString();
	this.displayName = (displayName || "").toString();
	this.mergeData = mergeData || {};
	this.isValid = this.IsValid();
};

asplib.mail.Address.prototype.toString = function() {
	return this.mailAddress ? (this.displayName ? (this.displayName + " <" + this.mailAddress.trim() + ">") : this.mailAddress.trim()) : "";
};

asplib.mail.Address.prototype.IsValid = function() {
	return this.mailAddress ? (/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w{2,}$/).test(this.mailAddress) : false;
};

asplib.mail.AddressList = function() {
	this.__name = "asplib.mail.AddressList";
	this.length = 0;
	for (var i=0; i<arguments.length; i++) {
		this.Add(arguments[i]);
	}
};

asplib.mail.AddressList.prototype.Add = function(mailAddress, displayName, mergeData) {
	if (typeof mailAddress === "string") mailAddress = new asplib.mail.Address(mailAddress, displayName, mergeData);
	if (mailAddress && mailAddress.IsValid && mailAddress.IsValid()) {
		var i = this.length++;
		this[i] = mailAddress;
	}
};

asplib.mail.AddressList.prototype.Clear = function() {
	for (var i=0; i<this.length; i++) {
		delete this[i];
	}
	this.length = 0;
};

asplib.mail.AddressList.prototype.toString = function() {
	var temp = [];
	for (var i=0; i<this.length; i++) {
		temp.push(this[i].toString());
	}
	return temp.join("; ");
}


%>