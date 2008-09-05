
/* dom.js - by Thomas Kjoernes <thomas@ipv.no> */

if (!window.trace) window.trace = function() {};

function openDialog(url, args, features, name) {
	var winref = null;
	var availTop = (screen.availTop) ? screen.availTop : 0;
	var availLeft = (screen.availLeft) ? screen.availLeft : 0;
	if (!features) features = {};
	if (!features.top) features.top = 0;
	if (!features.left) features.left = 0;
	if (!features.width) features.width = 400;
	if (!features.height) features.height = 300;
	if (typeof features.center == "undefined") features.center = true;
	if (window.showModalDialog) {
		var features =
			(features.top ? ("dialogTop:" + availTop + features.top + ";") : "") +
			(features.left ? ("dialogLeft:" + availLeft + features.left + ";") : "") +
			"dialogWidth:" + features.width + "px;" +
			"dialogHeight:" + features.height + "px;" +
			"center:" + (features.center ? "yes" : "no") + ";" +
			"status:" + (features.status ? "yes" : "no") + ";" +
			"resizable:" + (features.resizable ? "yes" : "no") + ";" +
			"scroll:" + (features.scroll ? "yes" : "no");
		trace("window.showModalDialog", 1);
		showModalDialog(url, args, features);
	} else {
		if (features.center) {
			features.top = Math.floor((screen.availHeight-features.height)/2);
			features.left = Math.floor((screen.availWidth-features.width)/2);
		}
		features =
			(features.top ? ("top=" + availTop + features.top + ",") : "") +
			(features.left ? ("left=" + availLeft + features.left + ",") : "") +
			"width=" + features.width + "," +
			"height=" + features.height + "," +
			"menubar=no,toolbar=no,location=no," +
			"status=" + (features.status ? "yes" : "no") + "," +
			"resizable=" + (features.resizable ? "yes" : "no") + "," +
			"scrollbars=" + (features.scroll ? "yes" : "no") + "," +
			"modal=yes";
		trace("window.open", 1);
		winref = open(url, name, features);
		if (winref) winref.dialogArguments = args;
	}
	return winref;
}

function preventDefault(e) {
	if (!e) e = window.event;
	if (e.preventDefault) e.preventDefault();
	e.returnValue = false;
}

function stopPropagation(e) {
	if (!e) e = window.event;
	if (e.stopPropagation) e.stopPropagation();
	e.cancelBubble = true;
}

function insertChild(element, parentNode) {
	if (element && parentNode) {
		if (parentNode.firstChild) {
			parentNode.insertBefore(element, parentNode.firstChild);
		} else {
			parentNode.appendChild(element);
		}
	}
}

function insertAfter(element, siblingNode) {
	if (element && siblingNode) {
		if (siblingNode.nextSibling) {
			siblingNode.parentNode.insertBefore(element, siblingNode.nextSibling);
		} else {
			siblingNode.parentNode.appendChild(element);
		}
	}
}

function getViewportProperties() {
	var viewport = {};
	if (self.innerHeight) {
		viewport.innerWidth = self.innerWidth;
		viewport.innerHeight = self.innerHeight;
	} else if (document.documentElement && document.documentElement.clientHeight) {
		viewport.innerWidth = document.documentElement.clientWidth;
		viewport.innerHeight = document.documentElement.clientHeight;
	} else if (document.body) {
		viewport.innerWidth = document.body.clientWidth;
		viewport.innerHeight = document.body.clientHeight;
	}
	if (self.pageYOffset) {
		viewport.scrollLeft = self.pageXOffset;
		viewport.scrollTop = self.pageYOffset;
	} else if (document.documentElement && document.documentElement.scrollTop) {
		viewport.scrollLeft = document.documentElement.scrollLeft;
		viewport.scrollTop = document.documentElement.scrollTop;
	} else if (document.body)	{
		viewport.scrollLeft = document.body.scrollLeft;
		viewport.scrollTop = document.body.scrollTop;
	}
	if (document.body.scrollHeight > document.body.offsetHeight) {
		viewport.scrollWidth = document.body.scrollWidth;
		viewport.scrollHeight = document.body.scrollHeight;
	} else {
		viewport.scrollWidth = document.body.offsetWidth;
		viewport.scrollHeight = document.body.offsetHeight;
	}
	return viewport;
}

function hasClass(element, className) {
	if (element && className) {
		var classNames = (element.className) ? element.className.split(" ") : [];
		var l = classNames.length;
		for (var i=0; i<l; i++) {
			if (classNames[i] === className) return true;
		}
	}
}

function addClass(element, className) {
	if (element && className) {
		var classNames = (element.className) ? element.className.split(" ") : [];
		if (!hasClass(element, className)) classNames.push(className);
		element.className = classNames.join(" ");
		return element.className;
	}
}

function setClass(element, className) {
	if (element) {
		var oldClass = element.className;
		element.className = (className ? className : "");
		return oldClass;
	}
}

function replaceClass(element, className, newClassName) {
	if (element && className) {
		if (!newClassName) newClassName="";
		var replaced = false;
		var classNames = (element.className) ? element.className.split(" ") : [];
		var l = classNames.length;
		for (var i=0; i<l; i++) {
			if (classNames[i] === className) {
				classNames[i]=newClassName;
				replaced = true;
			}
		}
		if (!replaced) classNames.push(newClassName);
		element.className = classNames.join(" ");
		return element.className;
	}
}

function removeClass(element, className) {
	if (element && className) {
		var classNames = (element.className) ? element.className.split(" ") : [];
		var l = classNames.length;
		for (var i=0; i<l; i++) {
			if (classNames[i] === className) {
				classNames.splice(i,1);
				element.className = classNames.join(" ");
				break;
			}
		}
		return element.className;
	}
}
