<%

/* date.asp - by Thomas Kjoernes <thomas@ipv.no> */

// getWeek()
Date.prototype.getWeek = function() {
    var a = Math.floor((13-(this.getMonth()))/12);
    var y = this.getFullYear()+4800-a;
    var m = (this.getMonth())+(12*a)-2;
    var jd = this.getDate() + Math.floor(((153*m)+2)/5) + (365*y) + Math.floor(y/4) - Math.floor(y/100) + Math.floor(y/400) - 32045;
    var d4 = (jd+31741-(jd%7))%146097%36524%1461;
    var L = Math.floor(d4/1460);
    var d1 = ((d4-L)%365)+L;
    return Math.floor(d1/7)+1;
};

// isLeepYear()
Date.prototype.isLeepYear = function(year) {
    var y = year || this.getFullYear();
    return (y%4===0 && (y%100!==0 || y%400===0)) ? true : false;
};

// toDateFormat()
Date.prototype.format = function(format, lang) {
	if (format != null) {
		var temp = format.toString();
		temp = temp.replace("MMMM", this.getMonthName(null, lang));
		temp = temp.replace("MMM", this.getMonthName(null, lang).substring(0,3));
		temp = temp.replace("DDDD", this.getDayName(null, lang));
		temp = temp.replace("DDD", this.getDayName(null, lang).substring(0,3));
		temp = temp.replace("YYYY", this.getFullYear().toPaddedString(4));
		temp = temp.replace("YY", this.getYear().toPaddedString(2));
		temp = temp.replace("MM", this.getMonth().inc().toPaddedString(2));
		temp = temp.replace("M", this.getMonth().inc().toString());
		temp = temp.replace("DD", this.getDate().toPaddedString(2));
		temp = temp.replace("D", this.getDate().toString());
		temp = temp.replace("W", this.getWeek());
		temp = temp.replace("hh", this.getHours().toPaddedString(2));
		temp = temp.replace("h", this.getHours().toString());
		temp = temp.replace("mm", this.getMinutes().toPaddedString(2));
		temp = temp.replace("ss", this.getSeconds().toPaddedString(2));
		temp = temp.replace("f", this.getMilliseconds().toString());
		return temp;
	}
};

// toString()
Date.prototype.toString = function() {
	return this.toUTCString();
};

// toISOString()
Date.prototype.toISOString = function() {
	return this.format("YYYY-MM-DD hh:mm:ss");
};

%>