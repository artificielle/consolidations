(this.webpackJsonpconsolidations=this.webpackJsonpconsolidations||[]).push([[0],{26:function(e,t,n){},29:function(e,t,n){},30:function(e,t,n){},31:function(e,t,n){"use strict";n.r(t);var r=n(0),c=n.n(r),a=n(15),o=n.n(a),u=n(3),s=n(5),i=n(16),l=n(10),j=n(4),f=n.n(j),b=n(8),d=n(11),p=n(13),h=n(7),x=n.n(h),O=(n(26),n(1)),v=function(e){var t=e.worksheet,n=k(t),r=n.table,c=n.size;return Object(O.jsx)("div",{className:"worksheet-container",children:Object(O.jsx)("table",{children:Object(O.jsx)("tbody",{children:w(c[0]).map((function(e){return Object(O.jsx)("tr",{children:w(c[1]).map((function(t){var n;return Object(O.jsx)("td",{children:Object(O.jsx)("div",{children:m(null===r||void 0===r||null===(n=r[e])||void 0===n?void 0:n[t])})},t)}))},e)}))})})})},m=function(e){if(e&&e.master===e)return e.type===x.a.ValueType.Formula?B(e):e.text},B=function(e){var t,n,r=null===e||void 0===e||null===(t=e.value)||void 0===t?void 0:t.result;return null!==(n=null===r||void 0===r?void 0:r.error)&&void 0!==n?n:r},w=function(e){return Array.from({length:e}).map((function(e,t){return t+1}))},k=function(e){var t=[],n=[0,0];return e.eachRow((function(e,r){n[0]=Math.max(n[0],r),e.eachCell((function(e,c){var a;n[1]=Math.max(n[1],c),null!==(a=t[r])&&void 0!==a||(t[r]=[]),t[r][c]=e}))})),{table:t,size:n}},C=n(18),E=(n(28),n(29),function(e){var t=e.topic,n=Object(r.useState)(),c=Object(d.a)(n,2),a=c[0],o=c[1],j=Object(r.useState)([]),h=Object(d.a)(j,2),m=h[0],B=h[1],w=Object(r.useState)([]),k=Object(d.a)(w,2),E=k[0],A=k[1],S=Object(r.useState)([]),P=Object(d.a)(S,2),M=P[0],N=P[1];Object(r.useEffect)((function(){(function(){var e=Object(b.a)(f.a.mark((function e(){var n,r,c,a;return f.a.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,fetch("/xlsx-templates/20XX"+t+"\uff08\u6a21\u677f\uff09.xlsx");case 3:return n=e.sent,e.next=6,n.arrayBuffer();case 6:return r=e.sent,c=new x.a.Workbook,e.next=10,c.xlsx.load(r);case 10:a=[c.worksheets[0],c.worksheets[6]],"\u5408\u5e76\u73b0\u91d1\u6d41\u91cf\u8868"===t&&(a[1]=c.worksheets[7]),o(c),B(a),console.log("load template completed",c),e.next=21;break;case 17:e.prev=17,e.t0=e.catch(0),console.error(e.t0),alert("\u6a21\u677f\u6587\u4ef6\u9519\u8bef");case 21:case"end":return e.stop()}}),e,null,[[0,17]])})));return function(){return e.apply(this,arguments)}})()()}),[t]);var I=function(e){return function(){var t=Object(b.a)(f.a.mark((function t(n){var r;return f.a.wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,Promise.all(n.map(function(){var e=Object(b.a)(f.a.mark((function e(t){var n,r;return f.a.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,g(t);case 2:return n=e.sent,r=new x.a.Workbook,e.next=6,r.xlsx.load(n);case 6:return e.abrupt("return",r.worksheets[0]);case 7:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}()));case 3:r=t.sent,e((function(e){return[].concat(Object(l.a)(e),Object(l.a)(r))})),t.next=11;break;case 7:t.prev=7,t.t0=t.catch(0),console.error("read xlsx file",t.t0),alert("\u6587\u4ef6\u683c\u5f0f\u9519\u8bef");case 11:case"end":return t.stop()}}),t,null,[[0,7]])})));return function(e){return t.apply(this,arguments)}}()},R=Object(p.a)({onDrop:I(A)}),z=Object(p.a)({onDrop:I(N)}),F=function(){var e=Object(b.a)(f.a.mark((function e(){var n;return f.a.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(e.prev=0,a){e.next=3;break}throw new Error;case 3:return e.next=5,a.xlsx.writeBuffer();case 5:n=e.sent,Object(C.saveAs)(new Blob([n]),t+".xlsx"),e.next=13;break;case 9:e.prev=9,e.t0=e.catch(0),console.error("write xlsx file",e.t0),alert("\u6a21\u677f\u6587\u4ef6\u9519\u8bef");case 13:case"end":return e.stop()}}),e,null,[[0,9]])})));return function(){return e.apply(this,arguments)}}(),D=Object(r.useCallback)((function(){var e=function(e,t,n){var r,c=Object(i.a)(y(n));try{var a=function(){var n=r.value;e.getCell(n).value=t.map((function(e){return function(e){if(e.type!==x.a.ValueType.Number)throw new Error("".concat(e.worksheet.name," ").concat(e.address));return e}(e.getCell(n)).value})).reduce((function(e,t){return e+t}),0)};for(c.s();!(r=c.n()).done;)a()}catch(o){c.e(o)}finally{c.f()}},n=function(t){M.length&&e(m[1],M,t)},r=function(t){E.length&&e(m[0],[m[1]].concat(Object(l.a)(E)),t)};try{"\u5408\u5e76\u5229\u6da6\u8868"===t&&(n(["B5:C21","B23:C24","B26:C26","B28:C33"]),r(["B5:C21","B23:C24","B26:C26"])),"\u5408\u5e76\u73b0\u91d1\u6d41\u91cf\u8868"===t&&(n(["B6:C8","B10:C13","B17:C21","B23:C26","B30:C32","B34:C36","B41"]),r(["B6:C8","B10:C13","B17:C21","B23:C26","B30:C32","B34:C36","B41"])),"\u5408\u5e76\u8d44\u4ea7\u8d1f\u503a\u8868"===t&&(n(["B6:B10","B11:B20","B23:B30","B31:B40"]),n(["E6:E10","E11:E20","E23:E30","E31:E40","E41:E47"]),r(["B6:B10","B11:B20","B23:B30","B31:B40"]),r(["E6:E10","E11:E20","E23:E30","E31:E40","E41:E47"]))}catch(c){console.error("consolidate",c),alert("\u6587\u4ef6\u683c\u5f0f\u9519\u8bef: "+c)}B(Object(l.a)(m)),console.log("consolidate completed",m,E,M)}),[t,m,E,M]);return Object(O.jsxs)("div",{children:[Object(O.jsxs)("div",{className:"workspace-dropzone-container",children:[Object(O.jsxs)("div",Object(s.a)(Object(s.a)({},R.getRootProps()),{},{children:[Object(O.jsx)("input",Object(s.a)({},R.getInputProps())),Object(O.jsx)("span",{children:"\u6dfb\u52a0\u5206\u516c\u53f8"})]})),Object(O.jsxs)("div",Object(s.a)(Object(s.a)({},z.getRootProps()),{},{children:[Object(O.jsx)("input",Object(s.a)({},z.getInputProps())),Object(O.jsx)("span",{children:"\u6dfb\u52a0\u5b50\u516c\u53f8"})]})),Object(O.jsx)("div",{onClick:F,children:Object(O.jsx)("span",{children:"\u4e0b\u8f7d\u5408\u5e76\u8868"})})]}),Object(O.jsxs)(u.d,{onSelect:D,children:[Object(O.jsxs)(u.b,{children:[m.map((function(e,n){return Object(O.jsx)(u.a,{children:[t,t.replace("\u5408\u5e76","\u5408\u5e76\u5206\u516c\u53f8")][n]},n)})),E.map((function(e,t){return Object(O.jsxs)(u.a,{children:["\u5b50\u516c\u53f8 ",t+1]},t)})),M.map((function(e,t){return Object(O.jsxs)(u.a,{children:["\u5206\u516c\u53f8 ",t+1]},t)}))]}),m.map((function(e,t){return Object(O.jsx)(u.c,{children:Object(O.jsx)(v,{worksheet:e})},t)})),E.map((function(e,t){return Object(O.jsx)(u.c,{children:Object(O.jsx)(v,{worksheet:e})},t)})),M.map((function(e,t){return Object(O.jsx)(u.c,{children:Object(O.jsx)(v,{worksheet:e})},t)}))]})]})}),g=function(){var e=Object(b.a)(f.a.mark((function e(t){return f.a.wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.abrupt("return",new Promise((function(e,n){var r=new FileReader;r.onload=function(t){try{e(t.target.result)}catch(r){n(r)}},r.readAsArrayBuffer(t)})));case 1:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}(),y=function(e){return e.flatMap(A)},A=function(e){if(!e.includes(":"))return[e];var t=e.match(/^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/);if(!t||t.length<5)throw new Error("Invalid range: "+e);var n=Number(t[2]),r=Number(t[4]),c=t[1].charCodeAt(0),a=t[3].charCodeAt(0),o=function(e,t){return Array.from({length:t-e+1}).map((function(t,n){return e+n}))};return o(c,a).flatMap((function(e){return o(n,r).map((function(t){return String.fromCharCode(e)+String(t)}))}))},S=function(){var e=["\u5408\u5e76\u5229\u6da6\u8868","\u5408\u5e76\u73b0\u91d1\u6d41\u91cf\u8868","\u5408\u5e76\u8d44\u4ea7\u8d1f\u503a\u8868"];return Object(O.jsx)("div",{children:Object(O.jsxs)(u.d,{children:[Object(O.jsx)(u.b,{children:e.map((function(e,t){return Object(O.jsx)(u.a,{children:e},t)}))}),e.map((function(e,t){return Object(O.jsx)(u.c,{children:Object(O.jsx)(E,{topic:e})},t)}))]})})},P=function(){return Object(O.jsx)(O.Fragment,{children:Object(O.jsx)(S,{})})};n(30);o.a.render(Object(O.jsx)(c.a.StrictMode,{children:Object(O.jsx)(P,{})}),document.getElementById("root"))}},[[31,1,2]]]);
//# sourceMappingURL=main.86187e59.chunk.js.map