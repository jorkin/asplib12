<%

/* form.asp - by Thomas Kjoernes <thomas@ipv.no> */

	if (!this.asplib) asplib = {};
	if (!this.asplib.html) asplib.html = {};

//--------------------------------------------------------------------
// asplib.html.Form()
//--------------------------------------------------------------------

asplib.html.Form = function(action, method, enctype, classNames, id, title) {
	trace("asplib.html.Form");
	this.__name = "asplib.html.Form";
	this.buttons = true;
	this.noSubmit = false;
	this.noReset = false;
	this.Init(action, method, enctype, classNames, id, title);
};

asplib.html.Form.prototype.Init = function(action, method, enctype, classNames, id, name, title) {
	trace("asplib.html.Form.Init");
	this.id = (id) ? id : "frm_" + new Date().valueOf();
	this.name = (name) ? name : this.id;
	this.iframeId = null;
	this.action = (action) ? action.toString() : "";
	this.method = (method) ? method.toString().toUpperCase() : "GET";
	this.enctype = (enctype) ? enctype.toString().toLowerCase() : "";
	this.classNames = (classNames) ? classNames : [];
	this.title = (title) ? title : "";
	this.document = {};
	this.document.__xml = function() { return (this.xml) ? " /" : ""; };
	this.document.__attr = function(name, value) { return (name && value != null && value !== "") ? (" " + name + "=\"" + Server.HTMLEncode(value) + "\"") : ""; };
	this.xml = false;
	this.params = [];
	this.tabs = [];
	this.html = [];
};

asplib.html.Form.prototype.__render = function() {
	var html = [];
	var tabs = [];
	for (var i in this.tabs) {
		var tab = this.tabs[i];
		var groups = [];
		for (var j in tab.groups) {
			var group = tab.groups[j];
			groups.push(group);
		}
		groups.sort(
			function(a,b) { return a.sortOrder-b.sortOrder;	}
		);
		tab.groups = groups;
		tabs.push(tab);
	}
	tabs.sort(
		function(a,b) {	return a.sortOrder-b.sortOrder; }
	);
	this.tabs = tabs;
	if (this.tabs[0]) {
		this.firstTab = this.tabs[0].id;
	}
	if (!this.selectedTab) {
		this.selectedTab = this.firstTab;
	}
	if (typeof this.onRender === "function") this.onRender();
	html.push(
		"<form" +
		this.document.__attr("id", this.id) +
		this.document.__attr("name", this.name) +
		this.document.__attr("class", this.classNames.join(" ")) +
		this.document.__attr("title", this.title) +
		this.document.__attr("action", this.action) +
		this.document.__attr("method", this.method) +
		this.document.__attr("enctype", this.enctype) +
		this.document.__attr("target", this.iframeId) +
		">\n"
	);
	for (var i in this.params) {
		html.push(
			"<input" +
			this.document.__attr("type", "hidden") +
			this.document.__attr("name", this.params[i].name) +
			this.document.__attr("value", this.params[i].value) +
			this.document.__xml() +
			">\n"
		);
	}
	if (this.buttonsOnTop) html.push(this.__buttons());
	if (tabs.length) {
		html.push("<ul" + this.document.__attr("class", "tabs") + ">\n");
		for (var i in this.tabs) {
			var tab = this.tabs[i];
			if (!tab.id) tab.id = i;
			if (!tab.title) tab.title = tab.id;
			var classNames = [];
			if (this.selectedTab === tab.id) classNames.push("active");
			html.push(
				"<li" +	this.document.__attr("class", classNames.join(" ")) + ">" +
				"<a" + this.document.__attr("href", "#" + "tab_" + tab.id) +
				">" + tab.title + "</a></li>\n"
			);
		}
		html.push("</ul>\n");
		for (var i in this.tabs) {
			var tab = this.tabs[i];
			var classNames = ["sheet"];
			if (this.selectedTab === tab.id) classNames.push("active");
			html.push(
				"<div" +
				this.document.__attr("id", "tab_" + tab.id) +
				this.document.__attr("class", classNames.join(" ")) +
				">\n"
			);
			if (tab.groups) {
				for (var j in tab.groups) {
					var group = tab.groups[j];
					if (!group.id) group.id = j;
					if (!group.title) group.title = group.id;
					group.id = "grp_" + group.id + "_" + i;
					html.push(
						"<fieldset" + this.document.__attr("id", group.id) + ">\n" +
						"<legend>" + group.title + "</legend>\n"
					);
					if (group.controls) {
						var controls = [];
						for (var k in group.controls) {
							var control = group.controls[k];
							if (control) controls.push(control);
						}
						controls.sort(
							function(a,b) {	return a.sortOrder-b.sortOrder; }
						);
						for (var k in controls) {
							html.push(controls[k]);
						}
					}
					html.push(
					"</fieldset>\n"
					);
				}
			}
			html.push(
			"</div>\n"
			);
			if (typeof this.controls !== "undefined") {
				if (typeof this.controls.join === "function") {
					html.push("<fieldset>\n");
					html.push("<legend></legend>\n");
					for (var i in this.controls) {
						html.push(this.controls[i]);
					}
					html.push("</fieldset>\n");
				} else {
					html.push(this.controls);
				}

			}
		}
	}
	html.push(this.html.join(""));
	if (!this.buttonsOnTop) html.push(this.__buttons());
	if (this.iframeId) {
		html.push(
			"<iframe" +
			this.document.__attr("id", this.iframeId) +
			this.document.__attr("name", this.iframeId) +
			this.document.__attr("class", "hidden") +
			"></iframe>\n"
		);
	}
	html.push("</form>\n");
	return html.join("");
};

asplib.html.Form.prototype.__buttons = function() {
	var html = [];
	if (this.buttons) {
		html.push("<div" + this.document.__attr("id", this.id + "_buttons") + this.document.__attr("class", "buttons") + ">\n");
		if (!this.noSubmit) {
			html.push(
			"<input" +
			this.document.__attr("type", "submit") +
			this.document.__attr("class", "button") +
			this.document.__attr("value", this.submitLabel) +
			this.document.__xml() + ">\n");
		}
		if (!this.noReset) {
			html.push(
			"<input" +
			this.document.__attr("type", "reset") +
			this.document.__attr("class", "button") +
			this.document.__attr("value", this.resetLabel) +
			this.document.__xml() + ">\n");
		}
		html.push("</div>\n");
	}
	return html.join("");
}

asplib.html.Form.prototype.AddParam = function(name, value) {
	trace("asplib.html.Form.AddParam");
	if (typeof name !== "undefined") {
		this.params.push({ name: name, value: value });
	}
};

asplib.html.Form.prototype.GetParam = function(name) {
	trace("asplib.html.Form.GetParam");
	if (typeof name !== "undefined") {
		for (var i=0; i<this.params.length; i++) {
			if (this.params[i].name === name) return this.params[i].value;
		}
	}
};

asplib.html.Form.prototype.AddTab = function(tabId, tabTitle, groups, sortOrder) {
	trace("asplib.html.Form.AddTab");
	if (!tabId) tabId = "default";
	if (!tabTitle) tabTitle = tabId;
	if (!this.tabs[tabId]) this.tabs[tabId] = { id: tabId, title: tabTitle, sortOrder: sortOrder, groups: [] };
	for (var i in groups) {
		this.AddGroup(tabId, groups[i].id, groups[i].title);
	}
	return tabId;
};

asplib.html.Form.prototype.AddGroup = function(tabId, groupId, groupTitle, controls, sortOrder) {
	trace("asplib.html.Form.AddGroup");
	if (!tabId) tabId = "default";
	if (!groupId) groupId = "default";
	if (!groupTitle) groupTitle = groupId;
	if (!this.tabs[tabId]) this.AddTab(tabId);
	if (!this.tabs[tabId].groups[groupId]) this.tabs[tabId].groups[groupId] = { id: groupId, title: groupTitle, sortOrder: sortOrder, controls: [] };
	for (var i in controls) {
		this.AddControl(tabId, groupId, controls[i]);
	}
	return groupId;
};

asplib.html.Form.prototype.AddControl = function(tabId, groupId, control) {
	trace("asplib.html.Form.AddControl");
	if (!tabId) tabId = "default";
	if (!groupId) groupId = "default";
	if (!this.tabs[tabId]) this.AddTab(tabId);
	if (!this.tabs[tabId].groups[groupId]) this.AddGroup(tabId, groupId);
	this.tabs[tabId].groups[groupId].controls.push(control);
};

asplib.html.Form.prototype.Put = function(text) {
	trace("asplib.html.Form.Put");
	this.html.push(text);
};

asplib.html.Form.prototype.Render = function() {
	trace("asplib.html.Form.Render");
	if (this.document && this.document.Put) {
		this.document.Put(this.__render());
	}
};

asplib.html.Form.prototype.toString = function() {
	return this.__render();
};

%>