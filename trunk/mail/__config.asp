<%

	asplib.mail.config["provider"] = PROVIDER_CDOSYS;
	asplib.mail.config["smtpServer"] = null; // use pickup-directory
	asplib.mail.config["smtpServerPort"] = 25;
	asplib.mail.config["smtpServerConnectionTimeout"] = 60;
	asplib.mail.config["smtpServerPickupDirectory"] = null;

%>