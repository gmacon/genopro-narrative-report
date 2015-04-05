<%[@ IncludeFile "Code/Lang.vbs" ]%>
<%[@ IncludeFile "Code/Util.vbs" ]%>
 		function tocSetToggle(close) {
				var tocOpen = "@[Report.Write StrDicExt("AltTOCToggleOpen", "", "This frame will stay open after an entry is selected. Click to change","", "2011.02.04")]@";
				var tocClose = "@[Report.Write StrDicExt("AltTOCToggleClose", "", "This frame will close after an entry is selected. Click to change","", "2011.02.04")]@"
				var toggle = document.images["tocStateButton"];
				toggle.src = (close ? "images/toc_close.gif" : "images/toc_open.gif");
				toggle.alt =  (close ? tocClose : tocOpen);
				toggle.title =  (close ? tocClose : tocOpen);
		}

		function tocToggle() {
				 var close = $.cookie('tocStateToggle') == 'Close';
				 $.cookie('tocStateToggle', (close ? 'Open' : 'Close'));
				 tocSetToggle(!close);
		}
		function doResize() {
                 document.getElementById('toc').style.height = parseInt(getInnerHeight() - 110) + 'px';
        };
        $(function () {
					var match,
						pl     = /\+/g,  // Regex for replacing addition symbol with a space
						search = /([^&=]+)=?([^&]*)/g,
						decode = function (s) { return decodeURIComponent(s.replace(pl, " ")); },
						query  = window.location.search.substring(1);

					var urlParams = {};
					while (match = search.exec(query))
					   urlParams[decode(match[1])] = decode(match[2]);
					if ('open' in urlParams) explorerTreeOpenTo(window, "names",urlParams['open'], 0, 1, "2");
					$.cookie('mytop',mytop);
					tocSetToggle($.cookie('tocStateToggle') == 'Close');
					hidePopUpFrame();
					PageInit(@[Report.Write Util.IfElse(Session("ForceFrames"), "true", "false")]@, '','names');
					doResize();
					window.onresize=doResize;
				});
