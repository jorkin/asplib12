<%

/* boolean.asp - by Thomas Kjoernes <thomas@ipv.no> */

// toYesNoString()
Boolean.prototype.toYesNoString = function(yes, no) {
	return (this.valueOf()) ? yes : no;
};

%>