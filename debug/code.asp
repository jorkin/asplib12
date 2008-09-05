<%

/* code.asp - by Thomas Kjoernes <thomas@ipv.no> */

function code(text) {

	var TABLEN			= 4;
	var WORDSEP			= RegExp().compile("(\\.|\\s|\\u00a0)");
	var WHITESPACE		= RegExp().compile("(\\s)");
	var WORDCHR			= RegExp().compile("\\w");
	var MARKUPCHR		= RegExp().compile("[a-z0-6\\-\\:\\!]", "i");
	var MARKUPLANG		= RegExp().compile("[a-z0-6\\s\\-\\:\"'\\<\\>/=\\!]", "i");
	var RE_BRACKET		= RegExp().compile("(\\[|\\]|\\{|\\}|\\(|\\))");
	var RE_KEYWORDS		= RegExp().compile("^(Array|Boolean|Date|Image|Math|Object|RegExp|ScriptEngine|ScriptEngineBuildVersion|ScriptEngineMajorVersion|ScriptEngineMinorVersion|String|UTC|abs|acos|addParameter|alert|all|anchor|appendChild|appendData|arguments|asin|async|atEnd|atan|atan2|attributes|back|backgroundColor|big|blink|blur|body|bold|break|byteToString|callee|caller|captureEvents|case|catch|ceil|charAt|charCodeAt|childNodes|children|clearInterval|clearTimeout|click|cloneNode|close|closed|compile|concat|confirm|contains|continue|cos|createAttribute|createComment|createDocument|createDocumentFragment|createElement|createProcessingInstruction|createProcessor|createTextNode|cursor|data|decodeURI|decodeURIComponent|default|delete|deleteData|description|disableExternalCapture|display|do|document|documentElement|elements|else|enableExternalCapture|encodeURI|encodeURIComponent|errorCode|escape|eval|event|exp|export|false|files|finally|find|firstChild|fixed|floor|focus|fontcolor|fontsize|for|forms|forward|fromCharCode|function|getAttribute|getAttributeNode|getDate|getDay|getElementById|getElementsByName|getElementsByTagName|getFullYear|getHours|getMilliseconds|getMinutes|getMonth|getOptionValue|getOptionValueCount|getSeconds|getSelection|getTime|getTimezoneOffset|getUTCDate|getUTCDay|getUTCFullYear|getUTCHours|getUTCMilliseconds|getUTCMinutes|getUTCMonth|getUTCSeconds|getYear|go|handleEvent|hasAttribute|hasAttributes|hasChildNodes|hasFeature|home|if|import|in|indexOf|innerHTML|input|insertBefore|insertData|instanceof|isNaN|italics|item|javaEnabled|join|lastChild|lastIndexOf|length|line|link|load|loadXML|location|log|match|max|mimeTypes|min|moveAbove|moveBelow|moveBy|moveNext|moveTo|moveToAbsolute|NaN|name|navigator|new|nextSibling|nodeName|nodeType|nodeValue|normalize|null|number|open|options|output|ownerDocument|parentNode|parse|parseError|parseFloat|parseInt|plugins|pop|pow|preference|previousSibling|print|prompt|prototype|push|random|readyState|reason|refresh|releaseEvents|reload|removeAttribute|removeAttributeNode|removeChild|replace|replaceChild|replaceData|reset|resizeBy|resizeTo|resolveExternals|return|reverse|round|routeEvent|screen|scroll|scrollBy|scrollTo|search|select|selectNodes|setAttribute|setAttributeNode|setDate|setFullYear|setHours|setInterval|setMinutes|setMonth|setSeconds|setTime|setTimeout|setUTCDate|setUTCDay|setUTCFullYear|setUTCHours|setUTCMilliseconds|setUTCMinutes|setUTCMonth|setUTCSeconds|setYear|shift|sin|slice|small|sort|sourceIndex|specified|splice|split|splitText|sqrt|srcElement|stop|strike|style|stylesheet|sub|submit|substr|substring|substringData|sup|switch|tagName|taintEnabled|tan|test|this|throw|toGMTString|toLocaleString|toLowerCase|toString|toUTCString|toUpperCase|transform|transformNode|true|try|typeof|undefined|unescape|unit|unshift|unwatch|validateOnParse|valueOf|var|void|watch|while|window|with|write|writeln|xml)$"); // words that make up language keywords
	var RE_OBJECT		= RegExp().compile("^(#include|Abandon|AbsolutePage|AbsolutePosition|ActiveConnection|ActiveXObject|ActualSize|AddHeader|AddNew|Append|AppendChunk|AppendToLog|Application|Application_OnEnd|Application_OnStart|AtEnd|Attributes|BOF|BeginTrans|BinaryRead|BinaryWrite|Bookmark|Buffer|CacheControl|CacheSize|Cancel|CancelBatch|CancelUpdate|CharSet|Clear|ClientCertificate|Clone|Close|CodePage|Command|CommandText|CommandTimeout|CommandType|CommitTrans|Connection|ConnectionString|ConnectionTimeout|ContentType|Contents|Cookies|CopyFile|CopyFolder|CopyTo|Count|CreateFile|CreateFolder|CreateObject|CreateParameter|CreateTextFile|CursorLocation|CursorType|DateCreated|DateLastModified|DefaultDatabase|DefinedSize|Delete|DeleteFile|Description|Direction|Domain|EOF|EOS|EditMode|EnableSessionState|End|Enumerator|Error|Errors|Execute|Expires|ExpiresAbsolute|Field|Fields|FileExists|Files|Filter|Flush|FolderExists|ForAppending|ForReading|ForWriting|Form|GetChunk|GetFile|GetFolder|GetLastError|GetObject|HasKeys|HTMLEncode|HelpContext|HelpFile|Index|IsClientConnected|IsolationLevel|Item|Key|LCID|Language|LineSeparator|LoadFromFile|Lock|LockType|MapPath|MarshalOptions|Mode|Move|MoveFirst|MoveLast|MoveNext|MovePrevious|Name|NamedParameters|NativeError|Number|NumericScale|ObjectContext|OnTransactionAbort|OnTransactionCommit|Open|OpenAsTextStream|OpenSchema|OpenTextFile|OriginalValue|PageCount|PageSize|Parameter|Parameters|Path|Pics|Position|Precision|Prepared|Properties|Property|Provider|QueryString|Read|ReadLine|ReadText|RecordCount|Recordset|Redirect|Requery|Request|Response|Resync|RollbackTrans|SQLState|Save|SaveToFile|ScriptTimeout|Seek|Server|ServerVariables|Session|SessionID|Session_OnEnd|Session_OnStart|SetAbort|SetComplete|Size|Sort|Source|State|StaticObjects|Status|SubFolders|Supports|Timeout|TotalBytes|Transfer|TristateFalse|TristateTrue|TristateUseDefault|Type|URLEncode|URLPathEncode|UnderlyingValue|Unlock|Update|UpdateBatch|Value|Version|Write|WriteLine|WriteText|adAddNew|adAffectAll|adAffectAllChapter|adAffectCurrent|adAffectGroup|adApproxPosition|adAsyncConnect|adAsyncExecute|adAsyncFetch|adAsyncFetchNonBlocking|adBSTR|adBigInt|adBinary|adBinary|adBookmark|adBookmarkCurrent|adBookmarkFirst|adBookmarkLast|adBoolean|adChapter|adChar|adChar|adClipString|adCmdFile|adCmdStoredProc|adCmdTable|adCmdTableDirect|adCmdText|adCmdUnknown|adCompareEqual|adCompareGreaterThan|adCompareLessThan|adCompareNotComparable|adCompareNotEqual|adCriteriaAllCol|adCriteriaKey|adCriteriaTimeStamp|adCriteriaUpdCol|adCurrency|adDBDate|adDBFileTime|adDBTime|adDBTimeStamp|adDate|adDecimal|adDelete|adDouble|adEditAdd|adEditDelete|adEditInProgre|adEditNone|adEmpty|adErrBoundToCommand|adErrDataConversion|adErrFeatureNotAvailable|adErrIllegalOperation|adErrInTransaction|adErrInvalidArgument|adErrInvalidConnection|adErrInvalidParamInfo|adErrItemNotFound|adErrNoCurrentRecord|adErrNotExecuting|adErrNotReentrant|adErrObjectClosed|adErrObjectInCollection|adErrObjectNotSet|adErrObjectOpen|adErrOperationCancelled|adErrProviderNotFound|adErrStillConnecting|adErrStillExecuting|adErrUnsafeOperation|adError|adExecuteNoRecords|adFileTime|adFilterAffectedRecord|adFilterConflictingRecord|adFilterFetchedRecord|adFilterNone|adFilterPendingRecord|adFilterPredicate|adFind|adFldCacheDeferred|adFldFixed|adFldIsNullable|adFldKeyColumn|adFldLong|adFldMayBeNull|adFldMayDefer|adFldRowID|adFldRowVersion|adFldUnknownUpdatable|adFldUpdatable|adGUID|adGetRowsRest|adHoldRecord|adHoldRecords|adIDispatch|adIUnknown|adIndex|adInteger|adLockBatchOptimistic|adLockOptimistic|adLockPessimistic|adLockReadOnly|adLongBinary|adLongChar|adLongVarBinary|adLongVarChar|adLongVarWChar|adLongWChar|adMarshalAll|adMarshalModifiedOnly|adModeRead|adModeReadWrite|adModeShareDenyNone|adModeShareDenyRead|adModeShareDenyWrite|adModeShareExclusive|adModeUnknown|adModeWrite|adMovePrevious|adNotify|adNumeric|adNumeric|adOpenDynamic|adOpenForwardOnly|adOpenKeyset|adOpenStatic|adParamInput|adParamInputOutput|adParamLong|adParamNullable|adParamOutput|adParamReturnValue|adParamSigned|adParamUnknown|adPersistADTG|adPersistXML|adPosBOF|adPosEOF|adPosUnknown|adPriorityAboveNormal|adPriorityBelowNormal|adPriorityHighest|adPriorityLowest|adPriorityNormal|adPromptAlway|adPromptComplete|adPromptCompleteRequired|adPromptNever|adPropNotSupported|adPropOptional|adPropRead|adPropRequired|adPropWrite|adPropiant|adPropVariant|adReadAll|adReadLine|adRecCanceled|adRecCantRelease|adRecConcurrencyViolation|adRecDBDeleted|adRecDeleted|adRecIntegrityViolation|adRecInvalid|adRecMaxChangesExceeded|adRecModified|adRecMultipleChange|adRecNew|adRecOK|adRecObjectOpen|adRecOutOfMemory|adRecPendingChange|adRecPermissionDenied|adRecSchemaViolation|adRecUnmodified|adRecalcAlway|adRecalcUpFront|adResync|adResyncAll|adResyncAllValue|adResyncAutoIncrement|adResyncConflict|adResyncInsert|adResyncNone|adResyncUnderlyingValue|adResyncUpdate|adRsnAddNew|adRsnClose|adRsnDelete|adRsnFirstChange|adRsnMove|adRsnMoveFirst|adRsnMoveLast|adRsnMoveNext|adRsnMovePreviou|adRsnRequery|adRsnResynch|adRsnUndoAddNew|adRsnUndoDelete|adRsnUndoUpdate|adRsnUpdate|adRunAsync|adSaveCreateNotExist|adSaveCreateOverWrite|adSchemaAssert|adSchemaCatalog|adSchemaCharacterSet|adSchemaCheckConstraint|adSchemaCollation|adSchemaColumn|adSchemaColumnPrivilege|adSchemaColumnsDomainUsage|adSchemaConstraintColumnUsage|adSchemaConstraintTableUsage|adSchemaCube|adSchemaDBInfoKeyword|adSchemaDBInfoLiteral|adSchemaDimension|adSchemaForeignKey|adSchemaHierarchie|adSchemaIndexe|adSchemaKeyColumnUsage|adSchemaLevel|adSchemaMeasure|adSchemaMember|adSchemaPrimaryKey|adSchemaProcedure|adSchemaProcedureColumn|adSchemaProcedureParameter|adSchemaPropertie|adSchemaProviderSpecific|adSchemaProviderType|adSchemaReferentialConstraint|adSchemaSQLLanguage|adSchemaSchemata|adSchemaStatistic|adSchemaTable|adSchemaTableConstraint|adSchemaTablePrivilege|adSchemaTranslation|adSchemaUsagePrivilege|adSchemaView|adSchemaViewColumnUsage|adSchemaViewTableUsage|adSearchBackward|adSearchForward|adSeek|adSeekAfter|adSeekAfterEQ|adSeekBefore|adSeekBeforeEQ|adSeekFirstEQ|adSeekLastEQ|adSingle|adSmallInt|adStateClosed|adStateConnecting|adStateExecuting|adStateFetching|adStateOpen|adStatusCancel|adStatusCantDeny|adStatusErrorsOccurred|adStatusOK|adStatusUnwantedEvent|adStringHTML|adStringXML|adTinyInt|adTypeBinary|adTypeText|adUnsignedBigInt|adUnsignedInt|adUnsignedSmallInt|adUnsignedTinyInt|adUpdate|adUpdateBatch|adUseClient|adUseServer|adUserDefined|adVarBinary|adVarchar|adVarChar|adVariant|adVarWChar|adVarNumeric|adWChar|adWChar|adXactAbortRetaining|adXactBrowse|adXactChao|adXactCommitRetaining|adXactCursorStability|adXactIsolated|adXactReadCommitted|adXactReadUncommitted|adXactRepeatableRead|adXactSerializable|adXactUnspecified|adiant|getAllResponseHeaders|getResponseHeader|open|responseText|responseXML|send|setRequestHeader|setTimeouts|status|virtual)$"); // words that make up native language objects and functions
	var RE_MARKUP		= RegExp().compile("^(a|abbr|acronym|address|applet|area|b|base|basefont|bdo|bgsound|big|blink|blockquote|body|br|button|caption|center|cite|code|col|colgroup|comment|dd|del|dfn|dir|div|dl|dt|em|embed|fieldset|font|form|frame|frameset|h|h1|h2|h3|h4|h5|h6|head|hr|hta:application|html|i|iframe|img|input|ins|isindex|kbd|label|legend|li|link|listing|map|marquee|menu|meta|multicol|nextid|nobr|noframes|noscript|object|ol|optgroup|option|p|param|plaintext|pre|q|s|samp|script|select|server|small|sound|spacer|span|strike|strong|style|sub|sup|table|tbody|td|textarea|textflow|tfoot|th|thead|title|tr|tt|u|ul|var|wbr|xmp|!DOCTYPE)$", "i"); // names of HTML elements
	var RE_MARKUP_ATTR	= RegExp().compile("^(abbr|accept-charset|accept|accesskey|action|addEventListener|align|alink|alt|applicationname|attachEvent|archive|autoFlush|axis|background|behavior|bgcolor|bgproperties|border|bordercolor|bordercolordark|bordercolorlight|borderstyle|buffer|caption|cellpadding|cellspacing|char|charoff|charset|checked|cite|class|className|classid|clear|code|codebase|codetype|color|cols|colspan|compact|content|contentType|coords|data|datetime|declare|defer|dir|direction|disabled|dynsrc|encoding|enctype|errorPage|extends|face|file|flush|for|frame|frameborder|framespacing|gutter|headers|height|href|hreflang|hspace|http-equiv|icon|id|import|info|isErrorPage|ismap|isThreadSafe|label|language|lang|leftmargin|link|longdesc|loop|lowsrc|marginheight|marginwidth|maximizebutton|maxlength|media|method|methods|minimizebutton|multiple|nohref|noresize|noshade|nowrap|object|onabort|onAbort|onblur|onBlur|onchange|onChange|onclick|onClick|ondblclick|onDblClick|onerror|onError|onfocus|onFocus|onkeydown|onKeyDown|onkeypress|onKeyPress|onkeyup|onKeyUp|onload|onLoad|onmousedown|onMouseDown|onmousemove|onMouseMove|onmouseout|onMouseOut|onmouseover|onMouseOver|onmouseup|onMouseUp|onreset|onReset|onselect|onSelect|onsubmit|onSubmit|onunload|onUnload|page|param|profile|prompt|property|readonly|rel|rev|rows|rowspan|rules|runat|scheme|scope|scrollamount|scrolldelay|scrolling|selected|session|shape|showintaskbar|singleinstance|size|span|src|standby|start|style|summary|sysmenu|tabindex|target|text|title|topmargin|type|urn|usemap|valign|value|valuetype|version|vlink|vrml|vspace|width|windowstate|wrap|xmlns|xmlns:jsp|xml:lang)$", "i"); // names of HTML attributes
	var html			= [];
	var text			= (text||"").replace(/\r\n|\r/g, "\n");
	var SOURCELEN		= text.length;
	var markuptmp		= [];
	var c				= -1;
	var wordtmp			= "";
	var FULLTAB			= "";
	var LINEFEED		= "\n";
	var TABCHAR			= "\t";
	var SPACE			= " ";
	var RESETCOL		= -1;
	var COMMENT_C		= 1;
	var COMMENT_CPP		= 2;
	var STRING_DBL		= 3;
	var STRING_SNGL		= 4;
	var WORD			= 5;
	var REGEXP			= 6;
	var HTML			= 7;
	var HTML_ATTR		= 8;
	var COMMENT_HTML	= 9;
	var NORMAL			= null;
	var STATE			= NORMAL;
	var SUBSTATE		= NORMAL;
	var colCount		= RESETCOL;

	for (var i=0; i<TABLEN; ++i) FULLTAB += "\u00a0";

//	html.push("<link rel=\"stylesheet\" type=\"text/css\" href=\"/asplib1.2/debug/styles/code.css\">\n");
//	html.push("<samp>");

//	try {
		while (++c < SOURCELEN) {
			++colCount;
			var chr		= text.charAt(c);
			var nChr	= (c+1 < SOURCELEN	? text.charAt(c+1) : null);
			var pChr	= (c-1 > 0			? text.charAt(c-1) : null);
			var p2Chr	= (c-2 > 0			? text.charAt(c-2) : null);
			if (STATE == COMMENT_C) {
				if (chr == "*" && nChr == "/") {
					// Terminate C comment
					html.push("*/</i>");
					++c;
					++colCount;
					STATE = NORMAL;
				}
				else if (chr == LINEFEED) {
					// Linebreak in C comment
					html.push("<br>");
					colCount = RESETCOL;
				}
				else if (chr == TABCHAR) {
					// Add tab equivalents that snap to a column grid of 'TABLEN' characters wide.
					var equivTabLen = colCount % TABLEN;
					html.push(FULLTAB.substr(equivTabLen));
					colCount += (TABLEN - equivTabLen - 1);
				}
				else if (chr == SPACE) {
					html.push((pChr == SPACE) ? "\u00a0" : SPACE); // Alternate &nbsp; with space character so that we have the ability to line-wrap, but don't lose the specified number of spaces in the string
				}
				else {
					// Continuation of C comment
					html.push(Server.HTMLEncode(chr));
				}
				continue;
			}
			else if (STATE == NORMAL && chr == "/" && nChr == "*") {
				// Start of C-style comment
				STATE = COMMENT_C;
				html.push("<i>/*");
				++c;
				++colCount;
				continue;
			}
			else if (STATE == COMMENT_CPP) {
				if (chr == LINEFEED) {
					// Terminate C++ comment
					html.push("</i><br>");
					STATE = NORMAL;
					colCount = RESETCOL;
				}
				else if (chr == TABCHAR) {
					// Add tab equivalents that snap to a column grid of 'TABLEN' characters wide.
					var equivTabLen = colCount % TABLEN;
					html.push(FULLTAB.substr(equivTabLen));
					colCount += (TABLEN - equivTabLen - 1);
				}
				else if (chr == SPACE) {
					html.push((pChr == SPACE) ? "\u00a0" : SPACE); // Alternate &nbsp; with space character so that we have the ability to line-wrap, but don't lose the specified number of spaces in the string
				}
				else {
					// Continuation of C comment
					html.push(Server.HTMLEncode(chr));
				}
				continue;
			}
			else if (STATE == NORMAL && chr == "/" && nChr == "/") {
				// Start of C++ style comment
				STATE = COMMENT_CPP;
				html.push("<i>//");
				++c;
				++colCount;
				continue;
			}
			else if (STATE == NORMAL && chr.match(RE_BRACKET)) {
				// Bracket
				html.push("<b>" + RegExp.$1 + "</b>");
				continue;
			}
			else if (STATE == STRING_DBL) {
				if ((chr == "\"" && (pChr != "\\" || (pChr == "\\" && p2Chr == "\\"))) || nChr == LINEFEED) {
					// Terminate " string
					if (nChr == LINEFEED) {
						html.push(Server.HTMLEncode(chr) + "</span>");
					}
					else {
						html.push("\"</span>");
					}
					STATE = NORMAL;
				}
				else if (chr == SPACE) {
					html.push((pChr == SPACE) ? "\u00a0" : SPACE); // Alternate &nbsp; with space character so that we have the ability to line-wrap, but don't lose the specified number of spaces in the string
				}
				else {
					// Continuation " string
					html.push(Server.HTMLEncode(chr));
				}
				continue;
			}
			else if (STATE == NORMAL && chr == "\"" && pChr != "\\") {
				// Start of " string
				STATE = STRING_DBL;
				html.push("<span>\"");
				continue;
			}
			else if (STATE == STRING_SNGL) {
				if ((chr == "'" && (pChr != "\\" || (pChr == "\\" && p2Chr == "\\"))) || nChr == LINEFEED) {
					// Terminate ' string
					html.push("'</span>");
					STATE = NORMAL;
				}
				else if (chr == SPACE) {
					html.push((pChr == SPACE) ? "\u00a0" : SPACE); // Alternate &nbsp; with space character so that we have the ability to line-wrap, but don't lose the specified number of spaces in the string
				}
				else {
					// Continuation ' string
					html.push(Server.HTMLEncode(chr));
				}
				continue;
			}
			else if (STATE == NORMAL && chr == "'" && pChr != "\\") {
				// Start of ' string
				STATE = STRING_SNGL;
				html.push("<span>'");
				continue;
			}
			else if (STATE == WORD) {
				// Continuation of word
				var blnValidChar = WORDCHR.test(chr);
				if (blnValidChar) {
					wordtmp += chr;
				}

				if (WORDSEP.test(nChr) || nChr == LINEFEED || !WORDCHR.test(nChr) || !blnValidChar) {
					// end of word, check if exists in dictionary
					var encWord = Server.HTMLEncode(wordtmp);
					if (RE_KEYWORDS.test(wordtmp)) {
						html.push("<u>" + encWord + "</u>");
					}
					else if (RE_OBJECT.test(wordtmp)) {
						html.push("<q>" + encWord + "</q>");
					}
					else {
						html.push(encWord);
					}
					STATE = NORMAL;
					wordtmp = "";
					if (!blnValidChar) {
						// Character that came in wasn't a valid word character, so we didn't add it to the word earlier in the statement block so we'll add it now.
						html.push(Server.HTMLEncode(chr));
					}
				}
				continue;
			}
			else if (STATE == NORMAL && WORDCHR.test(chr)) {
				// Start of word
				if (!WORDCHR.test(nChr)) {
					// If only a 1 letter word, don't start the WORD STATE as we can sometimes fail to process following brackets
					var encChr = Server.HTMLEncode(chr);
					if (RE_KEYWORDS.test(chr)) {
						html.push("<u>" + encChr + "</u>");
					}
					else if (RE_OBJECT.test(chr)) {
						html.push("<q>" + encChr + "</q>");
					}
					else {
						html.push(encChr);
					}
				}
				else {
					STATE = WORD;
					wordtmp = chr;
				}
				continue;
			}
			else if (STATE == REGEXP) {
				if (chr == "/" && pChr != "\\") {
					// Terminate Regular Expression
					html.push("/");
					STATE = NORMAL;
				}
				else {
					// Continuation of Regular Expression
					html.push(Server.HTMLEncode(chr));
				}
				continue;
			}
			else if (STATE == NORMAL && chr == "/" && nChr != "/" && pChr == "(") {
				// Start of Regular Expression
				STATE = REGEXP;
				html.push("/");
				continue;
			}
			else if (STATE == HTML) {
				if (chr == TABCHAR) {
					var equivTabLen = colCount % TABLEN;
					markuptmp.push(FULLTAB.substr(equivTabLen));
					colCount += (TABLEN - equivTabLen - 1);
				}
				else if (SUBSTATE == NORMAL && (chr == ">" || !MARKUPLANG.test(chr))) {
					// End of HTML
					STATE		= NORMAL;
					SUBSTATE	= NORMAL;
					if (chr == ">") {
						html.push("<span class=\"html\">" + markuptmp.join("") + Server.HTMLEncode(wordtmp + chr) + "</span>");
					}
					else {
						html.push(markuptmp.join("") + Server.HTMLEncode(wordtmp + chr)); // Content turned out not to be valid markup - say an opening < without valid markup inside
					}
					wordtmp		= "";
					markuptmp	= [];
				}
				else {
					var blnValidChar = MARKUPCHR.test(chr);
					if (blnValidChar) {
						wordtmp += chr;
					}
					if (SUBSTATE == STRING_DBL) {
						if ((chr == "\"" && (pChr != "\\" || (pChr == "\\" && p2Chr == "\\"))) || nChr == LINEFEED) {
							// Terminate " string
							if (nChr == LINEFEED) {
								markuptmp.push(Server.HTMLEncode(chr) + "</i>");
							}
							else {
								markuptmp.push("\"</i>");
							}
							SUBSTATE = NORMAL;
						}
						else {
							// Continuation " string
							markuptmp.push(Server.HTMLEncode(chr));
							wordtmp = "";
						}
						continue;
					}
					else if (SUBSTATE == NORMAL && chr == "\"" && pChr != "\\") {
						// Start of " string
						SUBSTATE = STRING_DBL;
						markuptmp.push("<i>\"");
						continue;
					}
					else if (SUBSTATE == COMMENT_HTML) {
						if (chr == "-" && nChr == "-") {
							// Terminate HTML comment
							markuptmp.push("--</i>");
							++c;
							++colCount;
							SUBSTATE = NORMAL;
						}
						else if (chr == LINEFEED) {
							// Linebreak in HTML comment
							markuptmp.push("<br>");
							colCount = RESETCOL;
						}
						else if (chr == TABCHAR) {
							// Add tab equivalents that snap to a column grid of 'TABLEN' characters wide.
							var equivTabLen = colCount % TABLEN;
							html.push(FULLTAB.substr(equivTabLen));
							colCount += (TABLEN - equivTabLen - 1);
						}
						else if (chr == SPACE) {
							html.push((pChr == SPACE) ? "\u00a0" : SPACE); // Alternate &nbsp; with space character so that we have the ability to line-wrap, but don't lose the specified number of spaces in the string
						}
						else {
							// Continuation of HTML comment
							markuptmp.push(Server.HTMLEncode(chr));
						}
						wordtmp = "";
						continue;
					}
					else if (SUBSTATE == NORMAL && chr == "-" && nChr == "-" && pChr == "!" && p2Chr == "<") {
						// Start of HTML comment
						SUBSTATE = COMMENT_HTML;
						markuptmp.push("<i class=\"html\">" + wordtmp);
						wordtmp = "";
						continue;
					}
					else if (SUBSTATE == STRING_SNGL) {
						if ((chr == "'" && (pChr != "\\" || (pChr == "\\" && p2Chr == "\\"))) || nChr == LINEFEED) {
							// Terminate ' string
							markuptmp.push("'</i>");
							SUBSTATE = NORMAL;
						}
						else {
							// Continuation ' string
							markuptmp.push(Server.HTMLEncode(chr));
							wordtmp = "";
						}
						continue;
					}
					else if (SUBSTATE == NORMAL && chr == "'" && pChr != "\\") {
						// Start of ' string
						SUBSTATE = STRING_SNGL;
						markuptmp.push("<i>'");
						continue;
					}
					else if (SUBSTATE == NORMAL && (!MARKUPCHR.test(nChr) || nChr == LINEFEED || !blnValidChar)) {
						// end of element, check if exists in dictionary
						var encTag = Server.HTMLEncode(wordtmp);
						if (RE_MARKUP.test(wordtmp)) {
							markuptmp.push("<b class=\"html\">" + encTag + "</b>");
						}
						else if (RE_MARKUP_ATTR.test(wordtmp)) {
							markuptmp.push("<b class=\"attr\">" + encTag + "</b>");
						}
						else {
							markuptmp.push(encTag);
						}
						wordtmp	= "";
						if (!blnValidChar) {
							// Character that came in wasn't a valid word character, so we didn't add it to the word earlier in the statement block so we'll add it now.
							markuptmp.push(Server.HTMLEncode(chr));
						}
					}
				}
				continue;
			}
			else if (STATE == NORMAL && chr == "<" && !WHITESPACE.test(nChr)) {
				// Start of HTML
				STATE		= HTML;
				SUBSTATE	= NORMAL;
				wordtmp		= "";
				markuptmp.push("&lt;");
				continue;
			}
			else {
				switch (chr) {
					case LINEFEED	: html.push("<br>"); colCount = RESETCOL; break;
					case TABCHAR	: {
						// Add tab equivalents that snap to a column grid of 'TABLEN' characters wide.
						var equivTabLen = colCount % TABLEN;
						html.push(FULLTAB.substr(equivTabLen));
						colCount += (TABLEN - equivTabLen - 1);
						break;
					}
					case SPACE		: html.push((pChr == SPACE) ? "\u00a0" : SPACE);  break; // Alternate &nbsp; with space character so that we have the ability to line-wrap, but don't lose the specified number of spaces in the string
					default			: html.push(Server.HTMLEncode(chr));
				}
			}
		}
//		html.push("</samp>");
		return html.join("");
//	} catch (e) { }
}

%>