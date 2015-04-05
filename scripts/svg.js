//	Scripts used by the SVG graphics
//
//	HISTORY
//	23-Nov-2005	Ron Prior       	Created
//  12-Oct-2010 Ron Prior        Major changes to zoom & pan for Google Chrome and other webkit browsers
//	Aug 2006 	Ron Prior		Added 'tooltip' code with acknowledgements to
//						Doug Schepers at www.svg-whiz.com
//
	var r, loop, genomap;
	var myparent = window.parent||window.__parent__		// work-around for 'parent' being broken in ASV6

	var SVGDocument = null;
	var root = null;
	var SVGViewBox = null;
	var svgns = 'http://www.w3.org/2000/svg';
	var xlinkns = 'http://www.w3.org/1999/xlink';
	var toolTip = null;
	var TrueCoords = null;
	var tipBox = null;
	var tipText = null;
	var tipTitle = null;
	var tipDesc = null;

	var lastElement = null;
	var titleText = '';
	var titleDesc = '';
	var tipping = false;
	var genopro = null;

	var map;
  
	var loaded = true;
	
	function doResize() {
		document.getElementById('svgCanvas').setAttributeNS(null,"height", parseInt(getInnerHeight() - 73) + 'px');
		if (svgwin && mapInfo.Toggle=="SVG") svgwin.setAll();
	}



function svgInit(info)
{
  if (! myparent) alert('svg init: orphaned!');
	map = info;
	root = document.getElementById('svgCanvas');

	 genomap=document.getElementById("viewport")

  toolTip = document.getElementById('ToolTip');
  tipBox =  document.getElementById('tipbox');
  tipText = document.getElementById('tipText');
  tipTemp = document.getElementById('tipTemp');
  genopro = document.getElementById('GenoPro');
 
  centermap(map.X, map.Y, map.Highlight);
  if (!map.Highlight && map.ZoomExtent) zoomextent(); 
 	setAll();
}
function centermap(x,y,hlight)
{
	var w=getInnerWidth();
	var h=getInnerHeight();
	var matrix = genomap.getCTM();
	setCTM(genomap, matrix.translate(-x + w / 2,-y + h / 2));
	var hl=document.getElementById("highlight");
  if (hlight) {
  	hl.setAttribute("visibility","visible");
  	hl.setAttribute("cx", x);
  	hl.setAttribute("cy", y);
	} else {
	  hl.setAttribute("visibility","hidden");
	}
	
}
function displaylink(href,target)
{
	try {
		window.open(href,target);
	} catch(e) {
		browserEval('window.open("'+href+'",target="'+target+'")');
	}
}
function panmap(evt,x,y)
{
		pan(x,y,5);
}
function pan(x,y,speed)
{
	var matrix = genomap.getCTM();
  setCTM(genomap, matrix.translate(x/matrix.a,y/matrix.d));
}
function ZoomMap(factor) {
  var p=root.createSVGPoint();
  p.x = getInnerWidth()/2;
  p.y = getInnerHeight()/2;
 	p = p.matrixTransform(genomap.getCTM().inverse());
	var k = root.createSVGMatrix().translate(p.x, p.y).scale(factor).translate(-p.x, -p.y);
  setCTM(genomap, genomap.getCTM().multiply(k));
}
function ZoomIn() {
         ZoomMap(2)
}
function ZoomOut() {
         ZoomMap(0.5)
}

function ZoomPanReset()
{
	genomap.currentScale = 1;
	genomap.currentTranslate.x = initX;
	genomap.currentTranslate.y = initY;
	setAll();
}
function zoomextent()
{
	var s;
	w=getInnerWidth();
	h=getInnerHeight();
	wImg=map.Bounds[2]-map.Bounds[0];
	hImg=map.Bounds[3]-map.Bounds[1];
	if(w/wImg < h/hImg) {
		s = w / wImg;
	} else {
		s = h / hImg;
	}
	ZoomMap(s);
}
	
function setStatic(id, xoffset, yoffset)
{
// undo effects of any pan, zoom or resize on element specified in 'id'
	var e=document.getElementById(id);
//	var s = 1/root.currentScale;
//	e.setAttribute( "transform", "scale(" + s + "," + s + ")" );
	e.setAttribute( "x", -root.currentTranslate.x + xoffset );
	e.setAttribute( "y", -root.currentTranslate.y + yoffset );
}
function setAll()
{
//
// negate effects of pan, zoom or resize on panning arrows
//
	w=getInnerWidth();
	h=getInnerHeight();
	setStatic("controls", 10, 10)
	setStatic("GenoPro", 40, 60);
	if (root.currentTranslate.x==0 && root.currentTranslate.y==0 && root.currentScale == 1) {
		//centermap(map.X, map.Y, map.Highlight);
	}
}	
function getInnerHeight() {
	if (window.innerHeight) { return window.innerHeight; } // netscape behavior
	else if (document.body && document.body.offsetHeight) { return  document.body.offsetHeight; } // IE behavior
	else if (document.documentElement.offsetHeight) { return  document.documentElement.offsetHeight; } // IE behavior
	else { return null; }
}

function getInnerWidth() {
	if (window.innerWidth) { return window.innerWidth; } // netscape behavior
	else if (document.body && document.body.offsetWidth) { return  document.body.offsetWidth; } // IE behavior
	else if (document.documentElement.offsetWidth) { return  document.documentElement.offsetWidth; } // IE behavior
	else { return null; }
}

      
    // Following is from Holger Will since ASV3 and O9 do not support getScreenTCM()
    // See http://groups.yahoo.com/group/svg-developers/message/50789
    function getScreenCTM(doc){
        if(doc.getScreenCTM) { return doc.getScreenCTM(); }
        
        var root=doc
        var sCTM= root.createSVGMatrix()
 
        var tr= root.createSVGMatrix()
        var par=root.getAttribute("preserveAspectRatio")
        if (par==null || par=="") par="xMidYMid meet"//setting to default value
        parX=par.substring(0,4) //xMin;xMid;xMax
        parY=par.substring(4,8)//YMin;YMid;YMax;
        ma=par.split(" ")
        mos=ma[1] //meet;slice
 
        //get dimensions of the viewport
        sCTM.a= 1
        sCTM.d=1
        sCTM.e= 0
        sCTM.f=0
 
 
        w=root.getAttribute("width")
        if (w==null || w=="") w=innerWidth
 
        h=root.getAttribute("height")
        if (h==null || h=="") h=innerHeight
 
        // Jeff Schiller:  Modified to account for percentages - I'm not 
        // absolutely certain this is correct but it works for 100%/100%
        if(w.substr(w.length-1, 1) == "%") {
            w = (parseFloat(w.substr(0,w.length-1)) / 100.0) * innerWidth;
        }
        if(h.substr(h.length-1, 1) == "%") {
            h = (parseFloat(h.substr(0,h.length-1)) / 100.0) * innerHeight;
        }
 
        // get the ViewBox
        vba=root.getAttribute("viewBox")
        if(vba==null) vba="0 0 "+w+" "+h
        var vb=vba.split(" ")//get the viewBox into an array
 
        //--------------------------------------------------------------------------
        //create a matrix with current user transformation
        tr.a= root.currentScale
        tr.d=root.currentScale
        tr.e= root.currentTranslate.x
        tr.f=root.currentTranslate.y
 
 
        //scale factors
        sx=w/vb[2]
        sy=h/vb[3]
 
 
        //meetOrSlice
        if(mos=="slice"){
        s=(sx>sy ? sx:sy)
        }else{
        s=(sx<sy ? sx:sy)
        }
 
        //preserveAspectRatio="none"
        if (par=="none"){
            sCTM.a=sx//scaleX
            sCTM.d=sy//scaleY
            sCTM.e=- vb[0]*sx //translateX
            sCTM.f=- vb[0]*sy //translateY
            sCTM=tr.multiply(sCTM)//taking user transformations into acount
 
            return sCTM
        }
 
 
        sCTM.a=s //scaleX
        sCTM.d=s//scaleY
        //-------------------------------------------------------
        switch(parX){
        case "xMid":
        sCTM.e=((w-vb[2]*s)/2) - vb[0]*s //translateX
 
        break;
        case "xMin":
        sCTM.e=- vb[0]*s//translateX
        break;
        case "xMax":
        sCTM.e=(w-vb[2]*s)- vb[0]*s //translateX
        break;
        }
        //------------------------------------------------------------
        switch(parY){
        case "YMid":
        sCTM.f=(h-vb[3]*s)/2 - vb[1]*s //translateY
        break;
        case "YMin":
        sCTM.f=- vb[1]*s//translateY
        break;
        case "YMax":
        sCTM.f=(h-vb[3]*s) - vb[1]*s //translateY
        break;
        }
        sCTM=tr.multiply(sCTM)//taking user transformations into acount
 
        return sCTM
    }


      function HideTooltip( evt )
      {
         if (toolTip) toolTip.setAttributeNS(null, 'visibility', 'hidden');
         children = tipText.childNodes;
         for (var i=0; i<children.length; i++) {
             tipText.removeChild(children.item(i));
         }
         return ! tipText.hasChildNodes;
       };


      function ShowTooltip( evt )
      {
          // there is a problem with Firefox in that occaionally not all tspan elements are removed
          // this hack seems to fix it!
          	 var isEmpty = HideTooltip(evt);
          if ( ! isEmpty ) {
          	isEmpty = HideTooltip(evt);
          }
          if ( ! isEmpty ) {
          	isEmpty = HideTooltip(evt);
          }
          if ( ! isEmpty ) {
          	isEmpty = HideTooltip(evt);
          }
          var tipScale = 1/root.currentScale;

          tipBox.setAttributeNS(null, 'transform', 'scale(' + tipScale + ',' + tipScale + ')' );
          tipText.setAttributeNS(null, 'transform', 'scale(' + tipScale + ',' + tipScale + ')' );


          var titleValue = '';
          var descValue = '';

	        var box = new Object();
    	    box.maxWidth = getInnerWidth();
    	    box.textWidth = 0;
    	    box.lineHeight = 10;

         var targetElement = evt.target;
         if ( lastElement != targetElement )
         {
            var targetTitle = targetElement.getElementsByTagName('title').item(0);
            if ( ! targetTitle ) targetTitle = targetElement.parentNode.getElementsByTagName('title').item(0);

            if ( targetTitle )
            {
               // if there is a 'title' element, use its contents for the tooltip title
               if (targetTitle.hasChildNodes()) titleValue = targetTitle.firstChild.nodeValue;
            }

            var targetDesc = targetElement.getElementsByTagName('desc').item(0);
            if ( ! targetDesc) targetDesc = targetElement.parentNode.getElementsByTagName('desc').item(0);
            if ( targetDesc )
            {
               // if there is a 'desc' element, use its contents for the tooltip desc
               if (targetDesc.hasChildNodes()) descValue = targetDesc.firstChild.nodeValue;

               if ( '' == titleValue )
               {
                  // if there is no 'title' element, use the contents of the 'desc' element for the tooltip title instead
 //                 titleValue = descValue;
 //                 descValue = '';
               }
            }

            // if there is still no 'title' element, use the contents of the 'id' attribute for the tooltip title
            if ( '' == titleValue )
            {
//               titleValue = targetElement.getAttributeNS(null, 'id');
            }

            // selectively assign the tooltip title and desc the proper values,
            //   and hide those which don't have text values
            //

            if ( '' != titleValue ) AddTipText(tipText, titleValue, 'black', box);

            if ( '' != descValue ) AddTipText(tipText, descValue, 'blue', box);
         }

         // if there are tooltip contents to be displayed, adjust the size and position of the box
         if ( tipText.hasChildNodes )
         {
           var p=root.createSVGPoint();
           p.x = evt.clientX;
           p.y = evt.clientY;
 	         p = p.matrixTransform(root.getCTM().inverse());
           var xPos = p.x + (10 * tipScale);
           var yPos = p.y + (10 * tipScale);


            //return rectangle around text as SVGRect object
	          // but getBBox() seems to be broken with Adobe so...
	          box.Height = tipText.childNodes.length * box.lineHeight

          	if (box.Height > 0 && box.textWidth > 0 ) {   
          	    if ((evt.clientX + box.textWidth + 10 ) > getInnerWidth())  xPos = ((getInnerWidth()  - box.textWidth -10)  - root.currentTranslate.x) * tipScale;
          	    if ((evt.clientY + box.Height + 20 ) > getInnerHeight()) yPos = ((getInnerHeight() - box.Height - 10) - root.currentTranslate.y) * tipScale;
          
                      tipBox.setAttributeNS(null, 'width', Number(box.textWidth + 10));
                      tipBox.setAttributeNS(null, 'height', Number(box.Height + 5));
          
                      // update position
                      toolTip.setAttributeNS(null, 'transform', 'translate(' + xPos + ',' + yPos + ')');
                      toolTip.setAttributeNS(null, 'visibility', 'visible');
          	}
         }
      };

function AddTipText(textNode, tip, col, tipbox)
{
	var tipLines = tip.split('\n'), tspanNode, newNode, childNode, len;
	for (var i=0; i<tipLines.length; i++) {
		newNode = document.createElementNS(svgns, 'tspan');
		tspanNode = textNode.appendChild(newNode);
		tspanNode.setAttributeNS(null, 'style', 'fill:'+col+';stroke:none;');
		tspanNode.setAttributeNS(null, 'x', '5');
		tspanNode.setAttributeNS(null, 'dy', Number(tipbox.lineHeight));
		childNode = document.createTextNode(tipLines[i] != ''?tipLines[i]:' ');
// Adobe's getComputedTextLength() returns total length of all tspan elements , so determine max length from each tspan
		newNode = tipTemp.appendChild(childNode);
		len = tipTemp.getComputedTextLength();
		childNode = tipTemp.removeChild(newNode);
		if (len > tipbox.textWidth ) tipbox.textWidth = len;
		newNode = tspanNode.appendChild(childNode);
	}
}
	/** 
 *  SVGPan library 1.2.2 as amended in issue 12 by sara & Callofdu
 * ======================
 *
 * Given an unique existing element with id "viewport" (or when missing, the 
 * first g-element), including the the library into any SVG adds the following 
 * capabilities:
 *
 *  - Mouse panning
 *  - Mouse zooming (using the wheel)
 *  - Object dragging
 *
 * You can configure the behaviour of the pan/zoom/drag with the variables
 * listed in the CONFIGURATION section of this file.
 *
 * Known issues:
 *
 *  - Zooming (while panning) on Safari has still some issues
 *
 * Releases:
 *
 * 1.2.2, Tue Aug 30 17:21:56 CEST 2011, Andrea Leofreddi
 *      - Fixed viewBox on root tag (#7)
 *      - Improved zoom speed (#2)
 *
 * 1.2.1, Mon Jul  4 00:33:18 CEST 2011, Andrea Leofreddi
 *      - Fixed a regression with mouse wheel (now working on Firefox 5)
 *      - Working with viewBox attribute (#4)
 *      - Added "use strict;" and fixed resulting warnings (#5)
 *      - Added configuration variables, dragging is disabled by default (#3)
 *
 * 1.2, Sat Mar 20 08:42:50 GMT 2010, Zeng Xiaohui
 *      Fixed a bug with browser mouse handler interaction
 *
 * 1.1, Wed Feb  3 17:39:33 GMT 2010, Zeng Xiaohui
 *      Updated the zoom code to support the mouse wheel on Safari/Chrome
 *
 * 1.0, Andrea Leofreddi
 *      First release
 *
 * This code is licensed under the following BSD license:
 *
 * Copyright 2009-2010 Andrea Leofreddi <a.leofreddi@itcharm.com>. All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 * 
 *    1. Redistributions of source code must retain the above copyright notice, this list of
 *       conditions and the following disclaimer.
 * 
 *    2. Redistributions in binary form must reproduce the above copyright notice, this list
 *       of conditions and the following disclaimer in the documentation and/or other materials
 *       provided with the distribution.
 * 
 * THIS SOFTWARE IS PROVIDED BY Andrea Leofreddi ``AS IS'' AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
 * FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL Andrea Leofreddi OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * 
 * The views and conclusions contained in the software and documentation are those of the
 * authors and should not be interpreted as representing official policies, either expressed
 * or implied, of Andrea Leofreddi.
 */

"use strict";
/// CONFIGURATION 
/// ====>

var enablePan = 1; // 1 or 0: enable or disable panning (default enabled)
var enableZoom = 1; // 1 or 0: enable or disable zooming (default enabled)
var enableDrag = 0; // 1 or 0: enable or disable dragging (default disabled)
var zoomScale = 0.8; // Zoom sensitivity

/// <====
/// END OF CONFIGURATION 

var state = 'none', svgRoot = null, stateTarget, stateOrigin, stateTf;

/**
 * Register handlers
 */
function setupHandlers(){


        setAttributes(root, {
                "onmouseup" : "handleMouseUp(evt)",
                "onmousedown" : "handleMouseDown(evt)",
                "onmousemove" : "handleMouseMove(evt)",
                //"onmouseout" : "handleMouseUp(evt)", // Decomment this to stop the pan functionality when dragging out of the SVG element
        });

/* <<  per issue 33     
        if(navigator.userAgent.toLowerCase().indexOf('webkit') >= 0 )
                root.addEventListener('mousewheel', handleMouseWheel, false); // Chrome/Safari
   << */
// >>
        if(navigator.userAgent.toLowerCase().indexOf('webkit') >= 0 || !!navigator.userAgent.match(/Trident\/7\./))
                root.addEventListener('mousewheel', handleMouseWheel, false); // Chrome/Safari/ie11
// >> end issue 33
				else
                root.addEventListener('DOMMouseScroll', handleMouseWheel, false); // Others
}

/**
 * Retrieves the root element for SVG manipulation. The element is then cached into the svgRoot global variable.
 */
function getRoot(root) {
        if(svgRoot == null) {
                var r = root.getElementById("viewport") ? root.getElementById("viewport") : root.documentElement, t = r;

                while(t != root) {
                        if(t.getAttribute("viewBox")) {
                                setCTM(r, t.getCTM());

                                t.removeAttribute("viewBox");
                        }

                        t = t.parentNode;
                }

                svgRoot = r;
        }

        return svgRoot;
}

/**
 * Instance an SVGPoint object with given event coordinates.
 */
/* from 1.2.2 
function getEventPoint(evt) {
        var p = root.createSVGPoint();

        p.x = evt.clientX;
        p.y = evt.clientY;

        return p;
} */
// from issue 14
function getEventPoint(evt) {
	var s = evt.target.nearestViewportElement;
	if (!s) s = evt.target; // in the unlikely event that event targets the svg element itself
	var p = s.createSVGPoint();
	var ctm = s.getScreenCTM();
	p.x = evt.clientX - ctm.e;
	p.y = evt.clientY - ctm.f;
	return p;
}

/**
 * Sets the current transform matrix of an element.
 */
function setCTM(element, matrix) {
// << per issue 30
        var s = "matrix(" + matrix.a + "," + matrix.b + "," + matrix.c + "," + matrix.d + "," + matrix.e + "," + matrix.f + ")";

        element.setAttribute("transform", s);
    element.transform.baseVal.consolidate();
 //  << */
 /* >>
    // The goal is to set element.transform to some matrix T='mat2' 
    // such that the new CTM is equal to the given input matrix.  The 
    // expression we need for 'mat2' is 
    //   T=(newCTM)*(inverse(oldCTM))*oldT .
    var mat2 = element.transform.baseVal.getItem(0).matrix.multiply(element.getCTM().inverse()).multiply(matrix);
                    element.transform.baseVal.replaceItem(element.transform.baseVal.createSVGTransformFromMatrix(mat2), 0);
// >> end issue 30 */
}

/**
 * Dumps a matrix to a string (useful for debug).
 */
function dumpMatrix(matrix) {
        var s = "[ " + matrix.a + ", " + matrix.c + ", " + matrix.e + "\n  " + matrix.b + ", " + matrix.d + ", " + matrix.f + "\n  0, 0, 1 ]";

        return s;
}

/**
 * Sets attributes of an element.
 */
function setAttributes(element, attributes){
        for (var i in attributes)
                element.setAttributeNS(null, i, attributes[i]);
}

/**
 * Handle mouse wheel event.
 */
function handleMouseWheel(evt) {
        if(!enableZoom)
                return;

        if(evt.preventDefault)
                evt.preventDefault();

        evt.returnValue = false;

        var svgDoc = evt.target.ownerDocument;

        var delta;

        if(evt.wheelDelta)
                delta = evt.wheelDelta / 360; // Chrome/Safari
        else
                delta = evt.detail / -9; // Mozilla

        var z = Math.pow(1 + zoomScale, delta);

        var g = getRoot(svgDoc);
        
        var p = getEventPoint(evt);

        p = p.matrixTransform(g.getCTM().inverse());

        // Compute new scale matrix in current mouse position
        var k = root.createSVGMatrix().translate(p.x, p.y).scale(z).translate(-p.x, -p.y);

        setCTM(g, g.getCTM().multiply(k));

        if(typeof(stateTf) == "undefined")
                stateTf = g.getCTM().inverse();

        stateTf = stateTf.multiply(k.inverse());
}

/**
 * Handle mouse move event.
 */
function handleMouseMove(evt) {
        if(evt.preventDefault)
                evt.preventDefault();

        evt.returnValue = false;

        var svgDoc = evt.target.ownerDocument;

        var g = getRoot(svgDoc);

        if(state == 'pan' && enablePan) {
                // Pan mode
                var p = getEventPoint(evt).matrixTransform(stateTf);

                setCTM(g, stateTf.inverse().translate(p.x - stateOrigin.x, p.y - stateOrigin.y));
        } else if(state == 'drag' && enableDrag) {
                // Drag mode
                var p = getEventPoint(evt).matrixTransform(g.getCTM().inverse());

                setCTM(stateTarget, root.createSVGMatrix().translate(p.x - stateOrigin.x, p.y - stateOrigin.y).multiply(g.getCTM().inverse()).multiply(stateTarget.getCTM()));

                stateOrigin = p;
        }
}

/**
 * Handle click event.
 */
function handleMouseDown(evt) {
        if(evt.preventDefault)
                evt.preventDefault();

        evt.returnValue = false;

        var svgDoc = evt.target.ownerDocument;

        var g = getRoot(svgDoc);

        if(
                evt.target.tagName == "svg" 
                || !enableDrag // Pan anyway when drag is disabled and the user clicked on an element 
        ) {
                // Pan mode
                state = 'pan';

                stateTf = g.getCTM().inverse();

                stateOrigin = getEventPoint(evt).matrixTransform(stateTf);
        } else {
                // Drag mode
                state = 'drag';

                stateTarget = evt.target;

                stateTf = g.getCTM().inverse();

                stateOrigin = getEventPoint(evt).matrixTransform(stateTf);
        }
}

/**
 * Handle mouse button release event.
 */
function handleMouseUp(evt) {
        if(evt.preventDefault)
                evt.preventDefault();

        evt.returnValue = false;

        var svgDoc = evt.target.ownerDocument;

        if(state == 'pan' || state == 'drag') {
                // Quit pan mode
                state = '';
        }
}

