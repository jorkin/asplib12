
/* tree.js - by Thomas Kjoernes <thomas@ipv.no> */

//	onInitStart()
//	onInitComplete()
//	onInitNodeStart(itemNode)
//	onInitNodeComplete(itemNode)
//	onFocus()
//	onBlur()
//	onSelect(itemNode)
//	onExecute(itemNode)
//	onDeselect(itemNode)
//	onRemoveSelection()
//	onExpand(itemNode)
//	onCollapse(itemNode)
//	onRename(itemNode)
//	onSaveChanges(itemNode)
//	onDiscardChanges(itemNode)
//	onContextMenu(itemNode, contextMenu, multiple)
//	onContextMenuExecute(itemNode, contextMenuItemNode)

if (!window.trace) window.trace = function() {};
if (!window.nextUniqueId) window.nextUniqueId = 0;
if (!window.config) config = {};
if (!config.Tree) config.Tree = {
	autoInit: true,
	debug: true,
	useRoot: true,
	focusOnEdit: true,
	focusOnSelect: true,
	selectTextOnEdit: false,
	imageBase: "icons/",
	loadingImage: "loading.gif",
	blankImage: "blank.gif",
	defaultType: "application/x-folder",
	defaultTarget: "",
	defaultImage: "folder.gif",
	defaultCollapsed: true,
	defaultTagged: false,
	defaultEditable: true,
	createUniqueIds: false
};

autoInitTree();

function autoInitTree() {
	var windowOnLoad = window.onload;
	window.onload = function() {
	//	trace("Tree:window.onload", 1);
		if (config.Tree.autoInit) {
			var els = document.getElementsByTagName("UL");
			for (var i=0; i<els.length; i++) {
				if (hasClass(els[i], "tree")) {
					els[i].Tree = new Tree(els[i]);
				}
			}
		}
		if (typeof windowOnLoad === "function") windowOnLoad();
	};
	var documentOnKeyDown = document.onkeydown;
	document.onkeydown = function(e) {
		var e = e || event;
		var keyCode = e.which || e.keyCode;
	//	trace("Tree:document.onkeydown:" + keyCode, 1);
		var currentUIControl = window.top.currentUIControl;
		if (currentUIControl && currentUIControl.hasFocus && !currentUIControl.editing) {
			if (window.top.currentUIControl.HandleKeyEvent(keyCode) === false) preventDefault(e);
		}
		if (typeof documentOnKeyDown === "function") documentOnKeyDown(e);
	};
	var documentOnClick = document.onclick;
	document.onclick = function(e) {
	//	trace("Tree:document.onclick", 1);
		if (window.top.cancelPendingClick) {
			window.top.cancelPendingClick = false;
			preventDefault(e);
		} else {
			if (window.top.currentUIControl) window.top.currentUIControl.Blur();
			if (typeof documentOnClick === "function") documentOnClick(e);
		}
	};
	var documentOnContextMenu = document.oncontextmenu;
	document.oncontextmenu = function(e) {
	//	trace("Tree:document.oncontextmenu", 1);
		if (window.top.cancelPendingClick) {
			window.top.cancelPendingClick = false;
			preventDefault(e);
		} else {
			if (window.top.currentUIControl) window.top.currentUIControl.Blur();
			if (typeof documentOnContextMenu === "function") documentOnContextMenu(e);
		}
	};
}

function Tree(root, target, type, collapsed, tagged, editable, image, loadingImage, imageBase, useRoot) {
	this.imageBase = (imageBase) ? imageBase : config.Tree.imageBase;
	this.loadingImage = (loadingImage) ? loadingImage : config.Tree.loadingImage;
	this.defaultImage = (image) ? image : config.Tree.defaultImage;
	this.defaultType = (type) ? type : config.Tree.defaultType;
	this.defaultTarget = (target) ? target : config.Tree.defaultTarget;
	this.defaultCollapsed = (typeof collapsed !== "undefined") ? collapsed : config.Tree.defaultCollapsed;
	this.defaultTagged = (typeof tagged !== "undefined") ? tagged : config.Tree.defaultTagged;
	this.defaultEditable = (typeof editable !== "undefined") ? editable : config.Tree.defaultEditable;
	this.useRoot = (typeof useRoot !== "undefined") ? useRoot : config.Tree.useRoot;
	this.focusOnEdit = config.Tree.focusOnEdit;
	this.focusOnSelect = config.Tree.focusOnSelect;
	this.selectTextOnEdit = config.Tree.selectTextOnEdit;
	this.Init(root);
}

Tree.prototype.Init = function(root) {
	this.rootNode = (typeof root === "string") ? document.getElementById(root) : root;
	this.selection = [];
	this.selectedNode = null;
	this.hasFocus = false;
	this.minInputWidth = 120;
	if (!this.rootNode) return;
	if (!this.rootNode.Tree) this.rootNode.Tree = this;
	if (typeof this.onInitStart === "function") this.onInitStart();
	addClass(this.rootNode, "tree");
	addClass(this.rootNode, "js");
	this.InitChildNodes(this.rootNode);
	this.rootNode.onclick = function(e) { this.Tree.RemoveSelection(); };
	this.rootNode.ondblclick = preventDefault;
	this.rootNode.onselectstart = function(e) {
		if (!this.Tree.editing) preventDefault(e);
	}
	if (hasClass(this.rootNode, "buttons")) {
		removeClass(this.rootNode, "buttons");
		var container = document.createElement("DIV");
		container.style.width = "200px";
		container.style.border = "1px solid buttonshadow";
		container.style.padding = "3px";
		container.style.marginBottom = "8px";
		container.style.backgroundColor = "buttonface";
		var button1 = document.createElement("BUTTON");
		button1.style.width = "90px";
		button1.style.marginLeft = "6px";
		button1.appendChild(document.createTextNode("Expand"));
		button1.Tree = this;
		button1.onclick = function(e) {
			this.Tree.ExpandAll();
		};
		var button2 = document.createElement("BUTTON");
		button2.style.width = "90px";
		button2.style.marginLeft = "6px";
		button2.appendChild(document.createTextNode("Collapse"));
		button2.Tree = this;
		button2.onclick = function(e) {
			this.Tree.CollapseAll();
		};
		container.appendChild(button1);
		container.appendChild(button2);
		this.rootNode.parentNode.insertBefore(container, this.rootNode);
	}
	if (typeof this.onInitComplete === "function") this.onInitComplete();
};

Tree.prototype.InitChildNodes = function(listNode) {
	if (!listNode) return;
	var itemNode = null;
	var firstItemNode = null;
	listNode.itemCount = 0;
	removeClass(listNode.lastItemNode, "last");
	for (var i=0; i<listNode.childNodes.length; i++) {
		if (listNode.childNodes[i].nodeName.toUpperCase() === "LI") {
			itemNode = listNode.childNodes[i];
			itemNode.itemIndex = listNode.itemCount++;
		//	removeClass(itemNode, "last");
			if (!itemNode.isInitialized) {
				itemNode.isInitialized = true;
				itemNode.isExpanded = !hasClass(itemNode, "collapsed");
				if (itemNode.itemIndex === 0) firstItemNode = itemNode;
				var divA = document.createElement("DIV");
				var divB = document.createElement("DIV");
				var divC = document.createElement("DIV");
				itemNode.ondblclick = preventDefault;
				itemNode.symbolNode = divA;
				if (typeof this.onInitNodeStart === "function") this.onInitNodeStart(itemNode);
				divA.appendChild(divB);
				divB.appendChild(divC);
				divA.className = "a";
				divB.className = "b";
				divC.className = "c";
				divB.Tree = this;
				divC.Tree = this;
				divB.itemNode = itemNode;
				divB.onclick = function(e) { stopPropagation(e); this.Tree.Toggle(this.itemNode); };
				divC.onclick = function(e) { stopPropagation(e); this.Tree.RemoveSelection(); };
				var childListNode = null;
				var l = itemNode.childNodes.length;
				for (var j=0; j<l; j++) {
					var tempNode = itemNode.childNodes[j];
					switch (tempNode.nodeName.toUpperCase()) {
						case "UL": {
							itemNode.childListNode = tempNode;
							childListNode = tempNode;
							addClass(divA, "children");
							l=j;
							break;
						}
						case "IMG": {
							itemNode.imageNode = tempNode;
							tempNode.Tree = this;
							tempNode.itemNode = itemNode;
							tempNode.onclick = function(e) {
								stopPropagation(e);
								if (this.itemNode.cancelPendingClick === true) {
									this.itemNode.cancelPendingClick = false;
									preventDefault(e);
								} else {
									var e = e || event;
									this.Tree.Select(this.itemNode, e.ctrlKey);
									if (e.ctrlKey) preventDefault(e);
								}
							};
							tempNode.ondblclick = function(e) { this.Tree.Toggle(this.itemNode); };
							tempNode.oncontextmenu = function(e) {
								var e = e || event;
								var x = e.pageX || e.clientX
								var y = e.pageY || e.clientY;
								if (!e.shiftKey) {
									this.Tree.ContextMenu(this.itemNode, x, y);
									stopPropagation(e);
									preventDefault(e);
								}
							};
							if (!tempNode.src) tempNode.src = "blank.gif";
							break;
						}
						case "A": {
							itemNode.anchorNode = tempNode;
							if (config.Tree.createUniqueIds) {
								tempNode.id = "a" + (++window.nextUniqueId) + "_" + (tempNode.id ? tempNode.id : "");
							}
						//	tempNode.title = tempNode.id;
							tempNode.Tree = this;
							tempNode.itemNode = itemNode;
							tempNode.previousOnClick = tempNode.onclick;
							tempNode.onclick = function(e) {
								stopPropagation(e);
								var e = e || event;
								this.Tree.Select(this.itemNode, e.ctrlKey);
								if (e.ctrlKey) {
									preventDefault(e);
								} else {
									if (this.itemNode.cancelPendingClick === true) {
										this.itemNode.cancelPendingClick = false;
										preventDefault(e);
									} else {
										if (this.Tree.editing || this.Tree.Execute(this.itemNode) === false) {
											preventDefault(e);
										} else {
											if (typeof this.previousOnClick	=== "function" &&
												this.previousOnClick(e) === false) preventDefault(e);
											if (!this.href || this.href == "#") preventDefault(e);
										}
									}
								}
							};
							tempNode.ondblclick = function(e) {
								this.Tree.Toggle(this.itemNode);
							};
							tempNode.oncontextmenu = function(e) {
								var e = e || event;
								var x = e.pageX || e.clientX
								var y = e.pageY || e.clientY;
								if (!e.shiftKey) {
									this.Tree.ContextMenu(this.itemNode, x, y);
									stopPropagation(e);
									preventDefault(e);
								}
							};
							tempNode.onkeydown = function(e) {
								var e = e || event;
								var keyCode = e.which || e.keyCode;
								switch (keyCode) {
									case 13: {
										if (!this.isSelected) this.Tree.Select(this.itemNode);
										this.Tree.Execute(this.itemNode);
										preventDefault(e);
										break;
									}
									case 32: {
										var x = e.pageX || e.clientX
										var y = e.pageY || e.clientY;
										this.Tree.ContextMenu(this.itemNode, x, y);
										preventDefault(e);
										break;
									}
								}
							};
							if (!tempNode.href) tempNode.href = "#" + tempNode.id;
							if (!tempNode.type) tempNode.type = this.defaultType;
							if (!tempNode.target) tempNode.target = this.defaultTarget;
							if (hasClass(tempNode, "active")) {
								tempNode.isSelected = true;
								this.selection.push(tempNode);
							}
							if (tempNode.firstChild.nodeType === 3) {
								var span = document.createElement("SPAN");
								span.appendChild(tempNode.firstChild);
								tempNode.appendChild(span);
								itemNode.anchorTextNode = span.firstChild;
							}
							break;
						}
						case "B": {
							itemNode.commentNode = tempNode;
							if (tempNode.firstChild.nodeType === 3) {
								itemNode.commentTextNode = tempNode.firstChild;
							}
							break;
						}
					}
				}
				if (!itemNode.anchorNode) {
					for (var j=0; j<l; j++) {
						var tempNode = itemNode.childNodes[j];
						if (tempNode.nodeType === 3) {
							itemNode.textNode; break;
						}
					}
				}
				if (!itemNode.imageNode) {
					addClass(itemNode, "text");
				}
				for (var j=0; j<l; j++) {
					divC.appendChild(itemNode.childNodes[0]);
				}
				if (itemNode.childNodes[0]) {
					itemNode.insertBefore(divA, itemNode.childNodes[0]);
				} else {
					itemNode.appendChild(divA);
				}
				if (typeof this.onInitNodeComplete === "function") this.onInitNodeComplete(itemNode);
			}
			this.InitChildNodes(childListNode);
		}
	}
	if (itemNode) {
		if (this.useRoot && this.rootNode === listNode && itemNode === firstItemNode) {
			addClass(firstItemNode, "root");
		}
		addClass(itemNode, "last");
	}
	listNode.firstItemNode = firstItemNode;
	listNode.lastItemNode = itemNode;
};

Tree.prototype.HandleKeyEvent = function(keyCode) {
	if (this.editing) return;
	switch (keyCode) {
		case 35: {
			this.SelectBottom();
			 return false;
		}
		case 36: {
			this.SelectTop();
			 return false;
		}
		case 37: {
			if (!this.Collapse()) this.SelectPrevious();
			 return false;
		}
		case 38: {
			this.SelectPrevious();
			 return false;
		}
		case 39: {
			if (!this.Expand()) this.SelectNext();
			 return false;
		}
		case 40: {
			this.SelectNext();
			 return false;
		}
		case 113: {
			this.Rename();
			return false;
		}
	}
};

Tree.prototype.Focus = function() {
	if (window.top.currentContextMenu) window.top.currentContextMenu.Hide();
	if (window.top.currentUIControl == this) return;
	if (window.top.currentUIControl) window.top.currentUIControl.Blur();
///	trace("Tree.Focus", 2);
	window.top.currentUIControl = this;
	if (typeof this.onFocus === "function") this.onFocus();
	this.hasFocus = true;
	addClass(this.rootNode, "active");
};

Tree.prototype.Blur = function() {
	if (window.top.currentContextMenu) window.top.currentContextMenu.Hide();
	if (window.top.currentUIControl == null) return;
	window.top.currentUIControl = null;
///	trace("Tree.Blur", 2);
	if (typeof this.onBlur === "function") this.onBlur();
	this.hasFocus = false;
	removeClass(this.rootNode, "active");
};

Tree.prototype.Execute = function(itemNode) {
///	trace("Tree.Execute", 2);
	if (typeof this.onExecute === "function") {
		return this.onExecute(itemNode);
	}
};

Tree.prototype.Select = function(itemNode, multiple) {
	this.Focus();
	if (!itemNode) return;
	if (!itemNode.anchorNode) return;
	if (!multiple && itemNode.anchorNode.isSelected && this.selection.length==1) {
		this.Rename();
	} else {
		if (multiple && itemNode.anchorNode.isSelected) {
			// remove/deselect item
			removeClass(itemNode.anchorNode, "active");
			itemNode.anchorNode.isSelected = false;
			if (this.focusOnSelect) itemNode.anchorNode.blur();
			var selection = [];
			for (var i=0; i<this.selection; i++) {
				if (this.selection[i] != itemNode) selection.push(this.selection[i]);
			}
			this.selection = selection;
			if (itemNode == this.selectedNode) this.selectedNode = this.selection[0];
		///	trace("Tree.Deselect", 2);
			if (typeof this.onDeselect === "function") this.onDeselect(itemNode);
		} else {
			// add/select item
			if (!multiple) this.RemoveSelection();
			addClass(itemNode.anchorNode, "active");
			itemNode.anchorNode.isSelected = true;
			if (this.focusOnSelect) itemNode.anchorNode.focus();
			this.selection.push(itemNode);
			this.selectedNode = itemNode;
		///	trace("Tree.Select", 2);
			if (typeof this.onSelect === "function") this.onSelect(itemNode);
		}
	}
};

Tree.prototype.SelectParent = function() {
	if (this.selection.length!=1) return;
	if (!this.selectedNode.parentNode) return;
	this.Select(this.selectedNode.parentNode.parentNode);
};

Tree.prototype.GetTop = function() {
	var itemNode = this.rootNode.firstChild;
	while (itemNode && itemNode.nodeName.toUpperCase() !== "LI") {
		itemNode = itemNode.nextSibling;
	}
	return itemNode;
};

Tree.prototype.GetBottom = function() {
	var itemNode = this.rootNode.lastChild;
	while (itemNode && itemNode.nodeName.toUpperCase() !== "LI") {
		itemNode = itemNode.previousSibling;
	}
	return this.GetPrevious(itemNode);
};

Tree.prototype.GetPrevious = function(itemNode) {
	if (itemNode && itemNode.childListNode && itemNode.isExpanded) {
		var tempNode = itemNode.childListNode.lastChild;
		while (tempNode && tempNode.nodeName.toUpperCase() !== "LI") {
			tempNode = tempNode.previousSibling;
		}
		if (tempNode) {
			return this.GetPrevious(tempNode);
		} else {
			return itemNode;
		}
	} else {
		return itemNode;
	};
};

Tree.prototype.GetNext = function(itemNode) {
	if (!itemNode) return;
	var parentNode = itemNode.parentNode;
	itemNode = itemNode.nextSibling;
	while (itemNode && itemNode.nodeName.toUpperCase() !== "LI") {
		itemNode = itemNode.nextSibling;
	}
	if (!itemNode && parentNode) {
		itemNode = this.GetNext(parentNode.parentNode);
	}
	return itemNode;
}

Tree.prototype.SelectTop = function() {
	this.Select(this.GetTop());
};

Tree.prototype.SelectBottom = function() {
	this.Select(this.GetBottom());
};

Tree.prototype.SelectPrevious = function() {
	var itemNode = this.selectedNode;
	if (!itemNode) return;
	var tempNode = itemNode.previousSibling;
	while (tempNode && tempNode.nodeName.toUpperCase() !== "LI") {
		tempNode = tempNode.previousSibling;
	}
	if (tempNode) {
		itemNode = this.GetPrevious(tempNode);
	} else {
		if (itemNode.parentNode) itemNode = itemNode.parentNode.parentNode;
	}
	this.Select(itemNode);
};

Tree.prototype.SelectNext = function() {
	var itemNode = this.selectedNode;
	if (!itemNode) return;
	var parentNode = itemNode.parentNode;
	if (itemNode.childListNode && itemNode.isExpanded) {
		itemNode = itemNode.childListNode.firstChild;
	} else {
		itemNode = itemNode.nextSibling;
	}
	while (itemNode && itemNode.nodeName.toUpperCase() !== "LI") {
		itemNode = itemNode.nextSibling;
	}
	if (!itemNode) {
		itemNode = this.GetNext(parentNode.parentNode);
	}
	this.Select(itemNode);
};

Tree.prototype.RemoveSelection = function() {
//	this.Focus();
	if (this.selection.length===0) return;
///	trace("Tree.RemoveSelection", 2);
	var l = this.selection.length;
	for (var i=0; i<l; i++) {
		var tempNode = this.selection[i].anchorNode;
		removeClass(tempNode, "active");
		tempNode.isSelected = false;
		tempNode.blur();
	}
	this.selection = [];
	this.selectedNode = null;
	if (typeof this.onRemoveSelection === "function") this.onRemoveSelection();
};

Tree.prototype.Expand = function(itemNode) {
//	this.Focus();
	if (!itemNode) itemNode = this.selectedNode;
	if (!itemNode) return false;
	if (!itemNode.childListNode) return false;
	if (itemNode.isExpanded) return false;
	itemNode.isExpanded = true;
	removeClass(itemNode, "collapsed");
///	trace("Tree.Expand", 2);
	if (typeof this.onExpand === "function") this.onExpand(itemNode);
	return true;
};

Tree.prototype.Collapse = function(itemNode) {
//	this.Focus();
	if (!itemNode) itemNode = this.selectedNode;
	if (!itemNode) return false;
	if (!itemNode.childListNode) return false;
	if (!itemNode.isExpanded) return false;
	itemNode.isExpanded = false;
	addClass(itemNode, "collapsed");
///	trace("Tree.Collapse", 2);
	if (typeof this.onCollapse === "function") this.onCollapse(itemNode);
	return true;
};

Tree.prototype.ExpandAll = function() {
	var temp = this.rootNode.getElementsByTagName("LI");
	for (var i=0; i<temp.length; i++) {
		removeClass(temp[i], "collapsed");
		if (temp[i].childListNode) temp[i].isExpanded = true;
	}
};

Tree.prototype.CollapseAll = function() {
	var temp = this.rootNode.getElementsByTagName("LI");
	for (var i=0; i<temp.length; i++) {
		if (temp[i].childListNode) {
			temp[i].isExpanded = false;
			addClass(temp[i], "collapsed");
		}
	}
};

Tree.prototype.Toggle = function(itemNode) {
	this.Focus();
	if (itemNode.isExpanded) {
		this.Collapse(itemNode);
	} else {
		this.Expand(itemNode);
	}
};

Tree.prototype.Tag = function(itemNode) {
	var selection = (itemNode) ? [itemNode] : this.selection;
	for (var i=0; i<selection.length; i++) {
		var itemNode = selection[i];
		if (!itemNode) continue;
		if (!itemNode.anchorNode) continue;
		addClass(itemNode.anchorNode, "tagged");
	///	trace("Tree.Tag", 2);
		itemNode.anchorNode.tagged = true;
	}
};

Tree.prototype.UnTag = function(itemNode) {
	var selection = (itemNode) ? [itemNode] : this.selection;
	for (var i=0; i<selection.length; i++) {
		var itemNode = selection[i];
		if (!itemNode) continue;
		if (!itemNode.anchorNode) continue;
		removeClass(itemNode.anchorNode, "tagged");
	///	trace("Tree.UnTag", 2);
		itemNode.anchorNode.tagged = false;
	}
};

Tree.prototype.Rename = function() {
	if (this.selection.length!=1) return;
	if (this.editing) return;
	if (window.top.currentContextMenu) window.top.currentContextMenu.Hide();
	var itemNode = this.selectedNode;
	if (!itemNode) return;
	if (!itemNode.anchorNode) return;
	if (!itemNode.anchorTextNode) return;
	if (!hasClass(itemNode.anchorNode, "editable")) return;
	if (typeof this.onRename === "function" && this.onRename(itemNode) === false) return;
///	trace("Tree.Rename", 2);
	this.editing = itemNode;
	var inputNode = document.createElement("INPUT");
	inputNode.Tree = this;
	inputNode.itemNode = itemNode;
	inputNode.type = "text";
	inputNode.value = itemNode.anchorTextNode.nodeValue;
	inputNode.style.width = Math.max(itemNode.anchorNode.offsetWidth, this.minInputWidth) + "px";
	inputNode.onblur = function(e) {
		this.Tree.SaveChanges();
		this.itemNode.cancelPendingClick = true;
	};
	inputNode.onclick = stopPropagation;
	inputNode.onkeydown = function(e) {
		var e = e || event;
		var keyCode = e.which || e.keyCode;
		switch (keyCode) {
			case 13 : {
				this.Tree.SaveChanges();
				preventDefault(e);
				break;
			}
			case 27 : {
				this.Tree.DiscardChanges();
				preventDefault(e);
				break;
			 }
		}
	};
	itemNode.editInputNode = inputNode;
	itemNode.anchorNode.blur();
	itemNode.anchorNode.parentNode.insertBefore(inputNode, itemNode.anchorNode);
	itemNode.anchorNode.style.display = "none";
	if (itemNode.commentNode) itemNode.commentNode.style.display = "none";
	if (this.focusOnEdit) inputNode.focus();
	if (this.selectTextOnEdit) inputNode.select();
};

Tree.prototype.SaveChanges = function() {
	if (!this.editing) return;
	var itemNode = this.editing;
	itemNode.anchorNode.parentNode.removeChild(itemNode.editInputNode);
	itemNode.anchorNode.style.display = "";
	if (this.focusOnSelect) itemNode.anchorNode.focus();
	if (itemNode.commentNode) itemNode.commentNode.style.display = "";
	itemNode.anchorTextNode.nodeValue = itemNode.editInputNode.value;
	itemNode.editInputNode = null;
	this.editing = null;
///	trace("Tree.SaveChanges", 2);
	if (typeof this.onSaveChanges === "function") this.onSaveChanges(itemNode);
};

Tree.prototype.DiscardChanges = function() {
	if (!this.editing) return;
	var itemNode = this.editing;
	itemNode.anchorNode.parentNode.removeChild(itemNode.editInputNode);
	itemNode.anchorNode.style.display = "";
	if (this.focusOnSelect) itemNode.anchorNode.focus();
	if (itemNode.commentNode) itemNode.commentNode.style.display = "";
	itemNode.editInputNode = null;
	this.editing = null;
///	trace("Tree.DiscardChanges", 2);
	if (typeof this.onDiscardChanges === "function") this.onDiscardChanges(itemNode);
};

Tree.prototype.ContextMenu = function(itemNode, x, y) {
	this.Focus();
	if (!itemNode) return;
	if (!itemNode.anchorNode) return;
	if (!itemNode.anchorNode.isSelected) this.Select(itemNode);
	if (window.Menu) {
	///	trace("Tree.ContextMenu", 2);
		if (!this.contextMenu) {
			this.contextMenu = new Menu;
			this.contextMenu.Tree = this;
			this.contextMenu.onExecute = function(contextMenuItemNode, x, y) {
			///	trace("Tree.ContextMenuExecute", 2);
				if (typeof this.Tree.onContextMenuExecute === "function")
					return this.Tree.onContextMenuExecute(this.Tree.selectedNode, contextMenuItemNode, this.Tree.selection.length>1, x, y);
			};
		}
		if (typeof this.onContextMenu === "function") {
			if (this.onContextMenu(itemNode, this.contextMenu, this.selection.length>1) === false) {
			} else {
				this.contextMenu.Show(x, y);
			}
		}
	}
};

Tree.prototype.Import = function(source, target, clear) {
///	trace("Tree.Import", 2);
	if (!this.importNode) this.importNode = document.createElement("DIV");
	this.importNode.innerHTML = (source && source.appendChild) ? source.innerHTML : source;
	var targetNode = (target && target.appendChild) ? target : document.getElementById(target);
	var sourceNode = this.importNode.firstChild;
	while (sourceNode && sourceNode.nodeName.toUpperCase() !== "UL") {
		sourceNode = sourceNode.nextSibling;
	}
	var listNode = sourceNode;
	if (sourceNode && targetNode) {
		if (targetNode.childListNode && clear) {
			targetNode.removeChild(targetNode.childListNode);
			targetNode.childListNode = null;
			if (targetNode.symbolNode) removeClass(targetNode.symbolNode, "children");
		}
		if (targetNode.childListNode) {
			listNode = targetNode.childListNode;
			for (var i=0; i<sourceNode.childNodes.length; i++) {
				var tempNode = sourceNode.childNodes[i];
				if (tempNode.nodeName.toUpperCase() === "LI") {
					listNode.appendChild(tempNode.cloneNode(true));
				}
			}
			if (sourceNode.parentNode) sourceNode.parentNode.removeChild(sourceNode);
		} else {
			targetNode.appendChild(sourceNode);
			targetNode.childListNode = sourceNode;
			if (targetNode.symbolNode) addClass(targetNode.symbolNode, "children");
		}
	} else {
		return sourceNode;
	}
	if (targetNode.isInitialized) {
		this.InitChildNodes(listNode);
	} else {
		this.Init(listNode);
	}
};

Tree.prototype.CreateNode = function(parentNode, text, href, type, target, image, comment, editable, tagged, collapsed) {
	var type = (typeof type !== "undefined") ? type : this.defaultType;
	var image = (typeof image !== "undefined") ? image : this.defaultImage;
	var target = (typeof target !== "undefined") ? target : this.defaultTarget;
	var collapsed = (typeof collapsed !== "undefined") ? collapsed : this.defaultCollapsed;
	var tagged = (typeof tagged !== "undefined") ? tagged : this.defaultTagged;
	var editable = (typeof editable !== "undefined") ? editable : this.defaultEditable;
	if (!parentNode && this.selection.length==1) parentNode = this.selectedNode;
	var listNode = (parentNode) ? parentNode.childListNode : this.rootNode;
	if (!listNode) {
		listNode = document.createElement("UL");
		listNode.itemCount = 0;
		parentNode.appendChild(listNode);
		parentNode.childListNode = listNode;
	} else {
		var tempNode = listNode.firstChild;
		while (tempNode && tempNode.nodeName.toUpperCase() !== "LI") {
			tempNode = tempNode.nextSibling;
		}
		removeClass(tempNode, "root");
	}
	var itemNode = document.createElement("LI");
	if (collapsed) addClass(itemNode, "collapsed");
	if (image != null) {
		var imageNode = document.createElement("IMG");
		var imageBase = (image.indexOf("/")!=0) ? this.imageBase : "";
		imageNode.src = (image) ? imageBase + image : config.Tree.blankImage;
		itemNode.appendChild(imageNode);
	}
	var anchorNode = document.createElement("A");
	if (href) anchorNode.href = href;
	if (type) anchorNode.type = type;
	if (target) anchorNode.target = target;
	if (editable) addClass(anchorNode, "editable");
	if (tagged) addClass(anchorNode, "tagged");
	if (text) anchorNode.appendChild(document.createTextNode(text));
	itemNode.appendChild(anchorNode);
	if (comment) {
		var commentNode = document.createElement("B");
		commentNode.appendChild(document.createTextNode(comment));
		itemNode.appendChild(commentNode);
	}
	listNode.appendChild(itemNode);
	listNode.itemCount++;
	if (parentNode && parentNode.symbolNode) addClass(parentNode.symbolNode, "children");
	this.InitChildNodes(listNode);
///	trace("Tree.CreateNode", 2);
	return itemNode;
};

Tree.prototype.RemoveNode = function(itemNode) {
	var selection = (itemNode) ? [itemNode] : this.selection;
	for (var i=0; i<selection.length; i++) {
		var tempNode = selection[i];
		var listNode = tempNode.parentNode;
		if (!listNode) continue;
		var parentItemNode = listNode.parentNode;
		listNode.removeChild(tempNode);
		listNode.itemCount--;
		var tempNode = null;
		var l = listNode.childNodes.length-1;
		for (var j=l; j>=0; j--) {
			if (listNode.childNodes[j].nodeName.toUpperCase() === "LI") {
				tempNode = listNode.childNodes[j];
				break;
			}
		}
		if (parentItemNode) {
			if (!tempNode) {
				removeClass(parentItemNode.symbolNode, "children");
				parentItemNode.removeChild(listNode);
				parentItemNode.childListNode = null;
				parentItemNode.isExpanded = false;
			} else {
				addClass(tempNode, "last");
			}
		}
	}
///	trace("Tree.RemoveNode", 2);
	this.RemoveSelection();
};