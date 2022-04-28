!function(){var e={9666:function(e,t,n){var r,o,i;function a(e){return a="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},a(e)}"undefined"!=typeof self?self:"undefined"!=typeof window?window:void 0!==n.g&&n.g,i=function(){"use strict";var e,t="function"==typeof atob,n="function"==typeof btoa,r="function"==typeof Buffer,o="function"==typeof TextDecoder?new TextDecoder:void 0,i="function"==typeof TextEncoder?new TextEncoder:void 0,a=Array.prototype.slice.call("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="),l=(e={},a.forEach((function(t,n){return e[t]=n})),e),u=/^(?:[A-Za-z\d+\/]{4})*?(?:[A-Za-z\d+\/]{2}(?:==)?|[A-Za-z\d+\/]{3}=?)?$/,c=String.fromCharCode.bind(String),f="function"==typeof Uint8Array.from?Uint8Array.from.bind(Uint8Array):function(e,t){return void 0===t&&(t=function(e){return e}),new Uint8Array(Array.prototype.slice.call(e,0).map(t))},s=function(e){return e.replace(/=/g,"").replace(/[+\/]/g,(function(e){return"+"==e?"-":"_"}))},d=function(e){return e.replace(/[^A-Za-z0-9\+\/]/g,"")},p=function(e){for(var t,n,r,o,i="",l=e.length%3,u=0;u<e.length;){if((n=e.charCodeAt(u++))>255||(r=e.charCodeAt(u++))>255||(o=e.charCodeAt(u++))>255)throw new TypeError("invalid character found");i+=a[(t=n<<16|r<<8|o)>>18&63]+a[t>>12&63]+a[t>>6&63]+a[63&t]}return l?i.slice(0,l-3)+"===".substring(l):i},h=n?function(e){return btoa(e)}:r?function(e){return Buffer.from(e,"binary").toString("base64")}:p,g=r?function(e){return Buffer.from(e).toString("base64")}:function(e){for(var t=[],n=0,r=e.length;n<r;n+=4096)t.push(c.apply(null,e.subarray(n,n+4096)));return h(t.join(""))},y=function(e,t){return void 0===t&&(t=!1),t?s(g(e)):g(e)},m=function(e){if(e.length<2)return(t=e.charCodeAt(0))<128?e:t<2048?c(192|t>>>6)+c(128|63&t):c(224|t>>>12&15)+c(128|t>>>6&63)+c(128|63&t);var t=65536+1024*(e.charCodeAt(0)-55296)+(e.charCodeAt(1)-56320);return c(240|t>>>18&7)+c(128|t>>>12&63)+c(128|t>>>6&63)+c(128|63&t)},b=/[\uD800-\uDBFF][\uDC00-\uDFFFF]|[^\x00-\x7F]/g,v=function(e){return e.replace(b,m)},x=r?function(e){return Buffer.from(e,"utf8").toString("base64")}:i?function(e){return g(i.encode(e))}:function(e){return h(v(e))},w=function(e,t){return void 0===t&&(t=!1),t?s(x(e)):x(e)},E=function(e){return w(e,!0)},B=/[\xC0-\xDF][\x80-\xBF]|[\xE0-\xEF][\x80-\xBF]{2}|[\xF0-\xF7][\x80-\xBF]{3}/g,A=function(e){switch(e.length){case 4:var t=((7&e.charCodeAt(0))<<18|(63&e.charCodeAt(1))<<12|(63&e.charCodeAt(2))<<6|63&e.charCodeAt(3))-65536;return c(55296+(t>>>10))+c(56320+(1023&t));case 3:return c((15&e.charCodeAt(0))<<12|(63&e.charCodeAt(1))<<6|63&e.charCodeAt(2));default:return c((31&e.charCodeAt(0))<<6|63&e.charCodeAt(1))}},C=function(e){return e.replace(B,A)},T=function(e){if(e=e.replace(/\s+/g,""),!u.test(e))throw new TypeError("malformed base64.");e+="==".slice(2-(3&e.length));for(var t,n,r,o="",i=0;i<e.length;)t=l[e.charAt(i++)]<<18|l[e.charAt(i++)]<<12|(n=l[e.charAt(i++)])<<6|(r=l[e.charAt(i++)]),o+=64===n?c(t>>16&255):64===r?c(t>>16&255,t>>8&255):c(t>>16&255,t>>8&255,255&t);return o},I=t?function(e){return atob(d(e))}:r?function(e){return Buffer.from(e,"base64").toString("binary")}:T,O=r?function(e){return f(Buffer.from(e,"base64"))}:function(e){return f(I(e),(function(e){return e.charCodeAt(0)}))},k=function(e){return O(R(e))},F=r?function(e){return Buffer.from(e,"base64").toString("utf8")}:o?function(e){return o.decode(O(e))}:function(e){return C(I(e))},R=function(e){return d(e.replace(/[-_]/g,(function(e){return"-"==e?"+":"/"})))},N=function(e){return F(R(e))},S=function(e){return{value:e,enumerable:!1,writable:!0,configurable:!0}},H=function(){var e=function(e,t){return Object.defineProperty(String.prototype,e,S(t))};e("fromBase64",(function(){return N(this)})),e("toBase64",(function(e){return w(this,e)})),e("toBase64URI",(function(){return w(this,!0)})),e("toBase64URL",(function(){return w(this,!0)})),e("toUint8Array",(function(){return k(this)}))},_=function(){var e=function(e,t){return Object.defineProperty(Uint8Array.prototype,e,S(t))};e("toBase64",(function(e){return y(this,e)})),e("toBase64URI",(function(){return y(this,!0)})),e("toBase64URL",(function(){return y(this,!0)}))},L={version:"3.7.2",VERSION:"3.7.2",atob:I,atobPolyfill:T,btoa:h,btoaPolyfill:p,fromBase64:N,toBase64:w,encode:w,encodeURI:E,encodeURL:E,utob:v,btou:C,decode:N,isValid:function(e){if("string"!=typeof e)return!1;var t=e.replace(/\s+/g,"").replace(/={0,2}$/,"");return!/[^\s0-9a-zA-Z\+/]/.test(t)||!/[^\s0-9a-zA-Z\-_]/.test(t)},fromUint8Array:y,toUint8Array:k,extendString:H,extendUint8Array:_,extendBuiltins:function(){H(),_()},Base64:{}};return Object.keys(L).forEach((function(e){return L.Base64[e]=L[e]})),L},"object"===a(t)?e.exports=i():void 0===(o="function"==typeof(r=i)?r.call(t,n,t,e):r)||(e.exports=o)}},t={};function n(r){var o=t[r];if(void 0!==o)return o.exports;var i=t[r]={exports:{}};return e[r].call(i.exports,i,i.exports,n),i.exports}n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,{a:t}),t},n.d=function(e,t){for(var r in t)n.o(t,r)&&!n.o(e,r)&&Object.defineProperty(e,r,{enumerable:!0,get:t[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){"use strict";var e=n(9666);function t(e,t){var n;return"hero"==e&&(n="https://hero.epa.gov/hero/index.cfm/reference/details/reference_id/"+t),"heronet"==e&&(n="http://heronet.epa.gov/heronet/index.cfm/reference/details/reference_id/"+t),n}function r(e,t){var n,r="https://hero.epa.gov/hero/index.cfm/reference/details/reference_id/",o="http://heronet.epa.gov/heronet/index.cfm/reference/details/reference_id/";return t.includes(r)||t.includes(o)?("hero"==e&&(n=t.replace(o,r)),"heronet"==e&&(n=t.replace(r,o)),n):null}function o(n){Word.run((function(o){document.getElementById("error-box").innerHTML='<p style="">Errors and Warnings</p>',document.getElementById("loader").style.display="flex",document.getElementById("app-body").style.display="none",document.getElementById("progress-text").innerHTML="Changing links...";var i=o.document.body,l=i.getOoxml(),u=i.getRange("Content").getHyperlinkRanges();u.load("items, hyperlink, font, text");var c,f,s,d,p,h,g,y=new DOMParser;return o.sync().then((function(){var f=y.parseFromString(l.value,"text/xml");return d=function(t,n){var r=[],o=[],i=n.getElementsByTagName("w:instrText"),l=n.getElementsByTagName("w:fldData");null===l&&(l=[]),l.length>0&&r.push(l[0].textContent);for(var u=1,c=1;c<l.length;c++)l[c-1].textContent!=l[c].textContent||u>1?(r.push(l[c].textContent),u=0):u+=1;for(var f=0,s=0;s<i.length;s++){var d=i[s].textContent;d.includes(" ADDIN EN.CITE ")&&" ADDIN EN.CITE "!=d&&o.push(d),l.length>0&&" ADDIN EN.CITE.DATA "==d&&(o.push("DECODE"),f+=1)}var p=!0;f!=r.length&&(document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-err">Error decoding citations; could not citations to fields.</span></p>',p=!1);for(var h=[],g=0,y=0;y<o.length;y++){var m;if("DECODE"==o[y]){if(m="",p)for(var b=r[g].split("\n"),v=0;v<b.length;v++){var x=b[v].replace("\r","");x.length>0&&(m+=e.Base64.decode(x))}g+=1}else m=o[y];m.length>0&&(h=a(0,h,m))}return h}(0,f),h=function(e,t){for(var n=new Object,r=t.getElementsByTagName("w:bookmarkStart"),o=0;o<r.length;o++){var i=r[o];if(i.hasAttribute("w:name")){var a=i.getAttribute("w:name");if(a.includes("_ENREF_")){for(var l,u=i.nextSibling,c=!1;!c&&null!==u;){if("w:r"==u.nodeName){var f=u.firstChild;if(null!==f&&"w:t"==f.nodeName){var s=f.textContent;if(null!==s&&s.length>40){c=!0,l=s;continue}}}u=u.nextSibling}c&&(n["#"+a]=l)}}}return n}(0,f),p=function(e,t,n){var r=new Object,o=[];for(var i in n)if(Object.prototype.hasOwnProperty.call(n,i)){for(var a=-1,l=n[i],u=-1,c=0;c<t.length;c++)if(!o.includes(c)){var f=t[c],s=0;if(""!=f.author&&l.includes(f.author)&&(s+=2),""!=f.year&&l.includes(f.year)&&(s+=1),""!=f.title&&l.includes(f.title)&&(s+=4),7==s){a=c,u=s;break}s<4||(-1!=a?s>u&&(a=c,u=s):(a=c,u=s))}-1==a?document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-err">Could not find a matching citation for bib entry "'+l.split(".")[0]+'...".</span></p>':(u<7&&(document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-warn">Matched citation for bib entry "'+l.split(".")[0]+'...", but it was not a full match.</span></p>'),r[i]=a,o.push(a))}return r}(0,d,h),function(e,n,o,i,a){for(var l=0;l<o.items.length;l++){var u=o.items[l].hyperlink,c=o.items[l].text;if(u in i&&u!=c){var f=t(n,a[i[u]].label);o.items[l].hyperlink=f}else{var s=r(n,u);if(null!==s)c==u&&o.items[l].insertText(s,"Replace"),o.items[l].hyperlink=s;else if(c!=u){var d="info";u.includes("_ENREF_")&&(o.items[l].font.highlightColor="#FFFF00",d="err"),document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-'+d+'">Hyperlink "'+c+'" ("'+u+'") not changed.</span></p>'}}}}(0,n,u,p,d),c=function(e,t,n){var r=new Object;for(var o in n)if(Object.prototype.hasOwnProperty.call(n,o)&&"#_ENREF_1"!=o){var i=t.search(n[o].substring(0,255),{matchCase:!0});i.load("items, text"),r[o]=i}return r}(0,i,h),document.getElementById("progress-text").innerHTML="Adding links in bibliography...",o.sync()})).then((function(){return f=function(e,t){var n=new Object;for(var r in t)if(Object.prototype.hasOwnProperty.call(t,r)){for(var o=[],i=t[r],a=0;a<i.items.length;a++){var l=i.items[a].getTextRanges(["."]);l.load("text"),o.push(l)}n[r]=o}return n}(0,c),o.sync()})).then((function(){return s=function(e,t){var n=new Object;for(var r in t)if(Object.prototype.hasOwnProperty.call(t,r)){for(var o=t[r],i=[],a=0;a<o.length;a++){for(var l=o[a],u=0,c=0,f=0;f<l.items.length;f++){var s=l.items[f];if(s.text.includes("(")&&(c+=1),s.text.includes(")")&&(c-=1),u+=s.text.length,!(s.text.length<4)){if(s.text.match(/[(][^)]{4,}[)][.]/))break;if(f>4)break;if(u>40&&0==c)break}}if(u>0){var d=l.items[0].getRange().expandTo(l.items[f].getRange()).getRange();d.load("text, hyperlink"),i.push(d)}else document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-err">Could find text to add ref '+r+" to a bibliography entry.</span></p>"}n[r]=i}return n}(0,f),o.sync()})).then((function(){!function(e,n,r,o,i){for(var a in r)if(Object.prototype.hasOwnProperty.call(r,a))for(var l=i[o[a]].label,u=r[a],c=0;c<u.length;c++)u[c].hyperlink=t(n,l)}(0,n,s,p,d);var e=o.document.body;return g=e.getOoxml(),document.getElementById("progress-text").innerHTML="Changing first bibliography entry...",o.sync()})).then((function(){var e=o.document.body,r=function(e,n,r,o,i,a,l){var u;if(u=function(e,n,r,o,i,a,l){var u="_ENREF_1";if(!o.includes(u))return null;if(!("#_ENREF_1"in i))return null;for(var c=t(n,l[a["#_ENREF_1"]].label),f=i["#_ENREF_1"],s=f.split("."),d=0,p=0,h=0;h<s.length;h++){var g=s[h];if(g.includes("(")&&(p+=1),g.includes(")")&&(p-=1),d+=g.length,!(g.length<4)){if(g.match(/[(][^)]{4,}[a-z]?[)][.]?/))break;if(h>4)break;if(d>40&&0==p)break}}var y="";if(d>0)for(var m=0;m<h+1;m++){var b="";m<s.length-1&&(b="."),y=y+s[m]+b}if(0==y.length)return document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-err">Could not find any text to link in the first citation.</span></p>',null;for(var v="",x=r.parseFromString(o,"text/xml").getElementsByTagName("w:bookmarkStart"),w=0;w<x.length;w++){var E=x[w];if(E.hasAttribute("w:name")&&"_ENREF_1"==E.getAttribute("w:name")){var B=E.parentNode;if("w:p"!=B.nodeName||!B.hasAttribute("w:rsidRPr"))continue;var A='<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "'+c+'" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidRPr="'+B.getAttribute("w:rsidRPr")+'"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>'+y+'</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',C=o.match(/<w:bookmarkStart w:id="\d" w:name="_ENREF_1"\/>/);if(!C){document.getElementById("error-box").innerHTML+="OH NO";continue}var T=C[0],I=o.split(T);v=I[0]+T+A+I[1];break}}return 0==v.length?(document.getElementById("error-box").innerHTML+='<p class="p-warn"><span class="style-err">Could not add a link to the first bibliography entry.</span></p>',null):v.replace(f,f.replace(y,""))}(0,n,r,o,i,a,l),null!==u){for(var c in i)if(Object.prototype.hasOwnProperty.call(i,c)){var f=new RegExp('<w:bookmarkStart w:id="\\d+" w:name="'+c.replace("#","")+'"/>',"g");u=u.replace(f,"")}}else u=o;return u}(0,n,y,g.value,h,p,d);return e.insertOoxml(r,"Replace"),"Errors and Warnings"==document.getElementById("error-box").textContent&&(document.getElementById("error-box").innerHTML='<p style="">Errors and Warnings: None</p>'),document.getElementById("loader").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("progress-text").innerHTML="",o.sync()}))})).catch((function(e){document.getElementById("loader").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("error-box").innerHTML+="Fatal error, script exiting...<br/>",document.getElementById("error-box").innerHTML+="Error: "+e+"<br/>",document.getElementById("progress-text").innerHTML="Fatal Error",e instanceof OfficeExtension.Error&&(document.getElementById("error-box").innerHTML+="Debug info: "+JSON.stringify(e.debugInfo)+"<br/>")}))}function i(e){return e.replace(/(&quot;|&lt;|&gt;|&amp;|&apos;)/g,(function(e){switch(e){case"&lt;":return"<";case"&gt;":return">";case"&amp;":return"&";case"&apos;":return"'";case"&quot;":return'"'}}))}function a(e,t,n){if(!n.includes("<EndNote><Cite>"))return t;for(var r=new Object,o=n.split("<Cite>"),a=0;a<o.length;a++){var l="",u="",c="",f="",s=o[a].match(/<label>(\d{1,})<\/label>/g);if(s)if(1==s.length)c=s[0];else{for(var d=[],p=0;p<s.length;p++)d.includes(s[p])||d.push(s[p]);1==d.length&&(c=d[0])}if(""!=c){c=c.match(/<label>(\d{1,})<\/label>/)[1];var h=o[a].match(/<Author>([^<>/]{1,})<\/Author>/),g=o[a].match(/<Author>([^<>/]{1,})<\/Author>/g);h&&g&&1==g.length&&(l=i(h[1]));var y=o[a].match(/<Year>(\d{4})<\/Year>/),m=o[a].match(/<Year>(\d{4})<\/Year>/g);y&&m&&1==m.length&&(u=y[1]);var b=o[a].match(/<title>([^<>]{1,})<\/title>/),v=o[a].match(/<title>([^<>]{1,})<\/title>/g);b&&v&&1==v.length&&(f=i(b[1]));var x=l+u+f;if(c in r){if(r[c].includes(x))continue;r[c].push(x)}else r[c]=[x];var w={label:c,author:l,year:u,title:f};t.push(w)}}return t}Office.onReady((function(e){e.host===Office.HostType.Word&&(Office.context.requirements.isSetSupported("WordApi","1.3")||(document.getElementById("error-box").innerHTML+="Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.<br/>"),OfficeExtension.config.extendedErrorLogging=!0,document.getElementById("link-hero").onclick=function(){o("hero")},document.getElementById("link-heronet").onclick=function(){o("heronet")},document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("progress-text").innerHTML="Initializing...")}))}()}();
//# sourceMappingURL=taskpane.js.map