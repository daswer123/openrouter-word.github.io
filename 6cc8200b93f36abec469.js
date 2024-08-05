/*! For license information please see 6cc8200b93f36abec469.js.LICENSE.txt */
function _typeof(t){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},_typeof(t)}function _createForOfIteratorHelper(t,e){var n="undefined"!=typeof Symbol&&t[Symbol.iterator]||t["@@iterator"];if(!n){if(Array.isArray(t)||(n=_unsupportedIterableToArray(t))||e&&t&&"number"==typeof t.length){n&&(t=n);var r=0,o=function(){};return{s:o,n:function(){return r>=t.length?{done:!0}:{done:!1,value:t[r++]}},e:function(t){throw t},f:o}}throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}var a,i=!0,c=!1;return{s:function(){n=n.call(t)},n:function(){var t=n.next();return i=t.done,t},e:function(t){c=!0,a=t},f:function(){try{i||null==n.return||n.return()}finally{if(c)throw a}}}}function _unsupportedIterableToArray(t,e){if(t){if("string"==typeof t)return _arrayLikeToArray(t,e);var n={}.toString.call(t).slice(8,-1);return"Object"===n&&t.constructor&&(n=t.constructor.name),"Map"===n||"Set"===n?Array.from(t):"Arguments"===n||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)?_arrayLikeToArray(t,e):void 0}}function _arrayLikeToArray(t,e){(null==e||e>t.length)&&(e=t.length);for(var n=0,r=Array(e);n<e;n++)r[n]=t[n];return r}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return e};var t,e={},n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(t,e,n){t[e]=n.value},a="function"==typeof Symbol?Symbol:{},i=a.iterator||"@@iterator",c=a.asyncIterator||"@@asyncIterator",l=a.toStringTag||"@@toStringTag";function s(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{s({},"")}catch(t){s=function(t,e,n){return t[e]=n}}function u(t,e,n,r){var a=e&&e.prototype instanceof y?e:y,i=Object.create(a.prototype),c=new C(r||[]);return o(i,"_invoke",{value:_(t,n,c)}),i}function p(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}e.wrap=u;var d="suspendedStart",m="suspendedYield",f="executing",h="completed",g={};function y(){}function x(){}function v(){}var b={};s(b,i,(function(){return this}));var S=Object.getPrototypeOf,w=S&&S(S(B([])));w&&w!==n&&r.call(w,i)&&(b=w);var k=v.prototype=y.prototype=Object.create(b);function I(t){["next","throw","return"].forEach((function(e){s(t,e,(function(t){return this._invoke(e,t)}))}))}function T(t,e){function n(o,a,i,c){var l=p(t[o],t,a);if("throw"!==l.type){var s=l.arg,u=s.value;return u&&"object"==_typeof(u)&&r.call(u,"__await")?e.resolve(u.__await).then((function(t){n("next",t,i,c)}),(function(t){n("throw",t,i,c)})):e.resolve(u).then((function(t){s.value=t,i(s)}),(function(t){return n("throw",t,i,c)}))}c(l.arg)}var a;o(this,"_invoke",{value:function(t,r){function o(){return new e((function(e,o){n(t,r,e,o)}))}return a=a?a.then(o,o):o()}})}function _(e,n,r){var o=d;return function(a,i){if(o===f)throw Error("Generator is already running");if(o===h){if("throw"===a)throw i;return{value:t,done:!0}}for(r.method=a,r.arg=i;;){var c=r.delegate;if(c){var l=E(c,r);if(l){if(l===g)continue;return l}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===d)throw o=h,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=f;var s=p(e,n,r);if("normal"===s.type){if(o=r.done?h:m,s.arg===g)continue;return{value:s.arg,done:r.done}}"throw"===s.type&&(o=h,r.method="throw",r.arg=s.arg)}}}function E(e,n){var r=n.method,o=e.iterator[r];if(o===t)return n.delegate=null,"throw"===r&&e.iterator.return&&(n.method="return",n.arg=t,E(e,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),g;var a=p(o,e.iterator,n.arg);if("throw"===a.type)return n.method="throw",n.arg=a.arg,n.delegate=null,g;var i=a.arg;return i?i.done?(n[e.resultName]=i.value,n.next=e.nextLoc,"return"!==n.method&&(n.method="next",n.arg=t),n.delegate=null,g):i:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,g)}function P(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function L(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function C(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(P,this),this.reset(!0)}function B(e){if(e||""===e){var n=e[i];if(n)return n.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,a=function n(){for(;++o<e.length;)if(r.call(e,o))return n.value=e[o],n.done=!1,n;return n.value=t,n.done=!0,n};return a.next=a}}throw new TypeError(_typeof(e)+" is not iterable")}return x.prototype=v,o(k,"constructor",{value:v,configurable:!0}),o(v,"constructor",{value:x,configurable:!0}),x.displayName=s(v,l,"GeneratorFunction"),e.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===x||"GeneratorFunction"===(e.displayName||e.name))},e.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,v):(t.__proto__=v,s(t,l,"GeneratorFunction")),t.prototype=Object.create(k),t},e.awrap=function(t){return{__await:t}},I(T.prototype),s(T.prototype,c,(function(){return this})),e.AsyncIterator=T,e.async=function(t,n,r,o,a){void 0===a&&(a=Promise);var i=new T(u(t,n,r,o),a);return e.isGeneratorFunction(n)?i:i.next().then((function(t){return t.done?t.value:i.next()}))},I(k),s(k,l,"Generator"),s(k,i,(function(){return this})),s(k,"toString",(function(){return"[object Generator]"})),e.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},e.values=B,C.prototype={constructor:C,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=t,this.done=!1,this.delegate=null,this.method="next",this.arg=t,this.tryEntries.forEach(L),!e)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=t)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var n=this;function o(r,o){return c.type="throw",c.arg=e,n.next=r,o&&(n.method="next",n.arg=t),!!o}for(var a=this.tryEntries.length-1;a>=0;--a){var i=this.tryEntries[a],c=i.completion;if("root"===i.tryLoc)return o("end");if(i.tryLoc<=this.prev){var l=r.call(i,"catchLoc"),s=r.call(i,"finallyLoc");if(l&&s){if(this.prev<i.catchLoc)return o(i.catchLoc,!0);if(this.prev<i.finallyLoc)return o(i.finallyLoc)}else if(l){if(this.prev<i.catchLoc)return o(i.catchLoc,!0)}else{if(!s)throw Error("try statement without catch or finally");if(this.prev<i.finallyLoc)return o(i.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var a=o;break}}a&&("break"===t||"continue"===t)&&a.tryLoc<=e&&e<=a.finallyLoc&&(a=null);var i=a?a.completion:{};return i.type=t,i.arg=e,a?(this.method="next",this.next=a.finallyLoc,g):this.complete(i)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),L(n),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;L(n)}return o}}throw Error("illegal catch attempt")},delegateYield:function(e,n,r){return this.delegate={iterator:B(e),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=t),g}},e}function asyncGeneratorStep(t,e,n,r,o,a,i){try{var c=t[a](i),l=c.value}catch(t){return void n(t)}c.done?e(l):Promise.resolve(l).then(r,o)}function _asyncToGenerator(t){return function(){var e=this,n=arguments;return new Promise((function(r,o){var a=t.apply(e,n);function i(t){asyncGeneratorStep(a,r,o,i,c,"next",t)}function c(t){asyncGeneratorStep(a,r,o,i,c,"throw",t)}i(void 0)}))}}function getPrompts(){return JSON.parse(localStorage.getItem("prompts")||"[]")}function savePrompts(t){localStorage.setItem("prompts",JSON.stringify(t))}function getLastSelectedPromptIndex(){return parseInt(localStorage.getItem("lastSelectedPromptIndex")||"-1",10)}function saveLastSelectedPromptIndex(t){localStorage.setItem("lastSelectedPromptIndex",t.toString())}function initSettings(){document.getElementById("settings").innerHTML='\n        <div class="mt-3">\n            <label for="api-key" class="form-label">API Ключ:</label>\n            <input type="password" id="api-key" class="form-control">\n        </div>\n        <div class="mt-3" style="margin-left: 20px;">\n            <label for="apply-formatting" class="form-check-label">\n                <input type="checkbox" id="apply-formatting" class="form-check-input">\n                Применять форматирование\n            </label>\n        </div>\n    ',document.getElementById("api-key").onchange=saveApiKey,document.getElementById("apply-formatting").onchange=saveApplyFormatting,loadModels(),loadSettings()}function saveApplyFormatting(){var t=document.getElementById("apply-formatting").checked;localStorage.setItem("applyFormatting",t)}function loadSettings(){var t=localStorage.getItem("apiKey")||"";document.getElementById("api-key").value=t;var e="true"===localStorage.getItem("applyFormatting");document.getElementById("apply-formatting").checked=e}function populateModelSelect_1(t){var e=document.getElementById("prompt-model");e.innerHTML="",t.forEach((function(t){var n=document.createElement("option");n.value=t.id,n.textContent=t.name,e.appendChild(n)}))}function saveApiKey(){var t=document.getElementById("api-key").value;localStorage.setItem("apiKey",t)}var OPENROUTER_BASE_URL="https://openrouter.ai/api/v1";function loadModels(){return _loadModels.apply(this,arguments)}function _loadModels(){return(_loadModels=_asyncToGenerator(_regeneratorRuntime().mark((function t(){var e,n;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,axios.get("".concat(OPENROUTER_BASE_URL,"/models"));case 3:e=t.sent,n=e.data.data,localStorage.setItem("models",JSON.stringify(n)),populateModelSelect_1(n),t.next=12;break;case 9:t.prev=9,t.t0=t.catch(0),console.error("Error loading models:",t.t0);case 12:case"end":return t.stop()}}),t,null,[[0,9]])})))).apply(this,arguments)}function generateTextWithAI(t,e,n,r){return _generateTextWithAI.apply(this,arguments)}function _generateTextWithAI(){return(_generateTextWithAI=_asyncToGenerator(_regeneratorRuntime().mark((function t(e,n,r,o){var a,i,c;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return a=localStorage.getItem("apiKey"),t.prev=1,t.next=4,fetch("".concat(OPENROUTER_BASE_URL,"/chat/completions"),{method:"POST",signal:o,headers:{Authorization:"Bearer ".concat(a),"Content-Type":"application/json","HTTP-Referer":"https://www.office.com","X-Title":"Word AI Assistant"},body:JSON.stringify({model:r,messages:[{role:"system",content:e},{role:"user",content:n}]})});case 4:if((i=t.sent).ok){t.next=7;break}throw new Error("HTTP error! status: ".concat(i.status));case 7:return t.next=9,i.json();case 9:return c=t.sent,t.abrupt("return",c.choices[0].message.content.trim());case 13:throw t.prev=13,t.t0=t.catch(1),console.error("Error in generateTextWithAI:",t.t0),t.t0;case 17:case"end":return t.stop()}}),t,null,[[1,13]])})))).apply(this,arguments)}var originalText="",currentController=null,lastSelectedPrompt=null;function initHome(){document.getElementById("home").innerHTML='\n      <select id="prompt-select" class="form-select mt-3"></select>\n      <div id="selected-model" class="mt-2"></div>\n      <button id="replace-text" class="btn btn-primary mt-3" disabled>Заменить выделенный текст</button>\n      <button id="cancel-request" class="btn btn-secondary mt-3" style="display: none;">Отменить запрос</button>\n      <button id="show-selected-text" class="btn btn-info mt-3">Показать информацию о выделенном тексте</button>\n      \x3c!-- <button id="apply-tricks" class="btn btn-warning mt-3">Применить приколы</button>\n      <button id="repeat-styles" class="btn btn-primary mt-3">Повтор стилей</button> --\x3e\n      <div id="selected-text-info" class="mt-3"></div>\n  ',$("#prompt-select").select2({width:"100%",placeholder:"Выберите промпт",templateResult:formatPrompt,templateSelection:formatPromptSelection}),document.getElementById("replace-text").onclick=replaceSelectedText,document.getElementById("cancel-request").onclick=cancelRequest,document.getElementById("show-selected-text").onclick=showSelectedText,updatePromptSelect(),setInterval(checkSelection,1e3)}function formatPrompt(t){if(!t.id)return t.text;var e=getPrompts()[t.id];return $("<span>".concat(e.name," (").concat(e.model,")</span>"))}function formatPromptSelection(t){return t.id?getPrompts()[t.id].name:t.text}function updatePromptSelect(){var t=$("#prompt-select"),e=document.getElementById("selected-model"),n=getPrompts();t.empty(),t.append(new Option("Выберите промпт","",!0,!0));var r=getLastSelectedPromptIndex();n.forEach((function(e,n){var o=new Option(e.name,n,!1,n===r);t.append(o)})),t.trigger("change");var o=r>=0&&r<n.length?r:null;lastSelectedPrompt=null!==o?n[o]:null,e.textContent=lastSelectedPrompt?"Модель: ".concat(lastSelectedPrompt.model):"",document.getElementById("replace-text").disabled=!lastSelectedPrompt,t.on("change",(function(){var t=$(this).val();lastSelectedPrompt=t?n[t]:null,saveLastSelectedPromptIndex(t||-1),e.textContent=lastSelectedPrompt?"Модель: ".concat(lastSelectedPrompt.model):"",document.getElementById("replace-text").disabled=!lastSelectedPrompt}))}function checkSelection(){Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var n,r;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(n=e.document.getSelection()).load("text"),t.next=4,e.sync();case 4:r=document.getElementById("replace-text"),n.text.trim().length>0?r.disabled=!1:r.disabled=!0;case 6:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}())}function replaceSelectedText(){return _replaceSelectedText.apply(this,arguments)}function _replaceSelectedText(){return _replaceSelectedText=_asyncToGenerator(_regeneratorRuntime().mark((function t(){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var n,r,o,a,i,c,l,s;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(n=e.document.getSelection()).load("text"),t.next=4,e.sync();case 4:if(originalText=n.text,lastSelectedPrompt){t.next=7;break}return t.abrupt("return");case 7:if(document.getElementById("replace-text").style.display="none",document.getElementById("cancel-request").style.display="inline-block",t.prev=9,currentController=new AbortController,r="true"===localStorage.getItem("applyFormatting"),o=[],!r){t.next=17;break}return t.next=16,getWordsInfo(e,n);case 16:o=t.sent;case 17:return a="".concat(lastSelectedPrompt.text,'\n\nUser provided the following text: "').concat(originalText,'"').concat(r?"\n\nFormatting information: ".concat(JSON.stringify(o)):""),t.next=21,generateTextWithAI("You are an AI assistant specialized in text processing and formatting. Your task is to generate text based on user prompts and, when required, update formatting information for the generated text. Always provide responses in the requested format.",a,lastSelectedPrompt.model,currentController.signal);case 21:if(i=t.sent,!r){t.next=43;break}return c='\nOriginal text: "'.concat(originalText,'"\nModified text: "').concat(i,'"\nOriginal formatting information: ').concat(JSON.stringify(o),'\n\nTask: Update the formatting information to match the new text while preserving the original styles and formats. Return ONLY the updated formatting information array in valid JSON format.\n\nInstructions:\n1. Analyze the original and modified texts.\n2. Create a new array of formatting information for the modified text.\n3. Preserve the original formatting styles as much as possible.\n4. Ensure the number of elements in the array matches the number of words in the modified text.\n5. Return ONLY the JSON array, without any additional text or explanation.\n\nThe JSON array should contain objects with the following possible properties:\n- text: string (required) - the word from the modified text\n- bold: boolean\n- italic: boolean\n- underline: boolean\n- color: string (e.g., "#FF0000" for red)\n- highlightColor: string (e.g., "yellow")\n- size: number\n- fontName: string\n- hyperlink: string (URL)\n- footnoteText: string\n\nExample of a valid response:\n[\n{"text": "This", "bold": true, "color": "#0000FF"},\n{"text": "is", "italic": true},\n{"text": "an", "underline": true},\n{"text": "example", "size": 14, "fontName": "Arial"},\n{"text": "with", "highlightColor": "yellow"},\n{"text": "formatting", "hyperlink": "https://example.com"},\n{"text": "applied", "footnoteText": "This is a footnote example"}\n]\n\nEnsure your response is a valid JSON array and includes formatting information for all words in the modified text.'),t.prev=25,t.next=28,generateTextWithAI("You are an AI assistant specialized in text formatting. Your task is to update formatting information for modified text while preserving the original styles as much as possible. Always return your response as a valid JSON array.",c,lastSelectedPrompt.model,currentController.signal);case 28:if(s=t.sent,l=JSON.parse(s),Array.isArray(l)){t.next=32;break}throw new Error("Response is not an array");case 32:t.next=39;break;case 34:return t.prev=34,t.t0=t.catch(25),console.error("Error parsing formatting information:",t.t0),n.insertText(i,Word.InsertLocation.replace),t.abrupt("return");case 39:return t.next=41,applyFormattedText(e,n,i,l);case 41:t.next=44;break;case 43:n.insertText(i,Word.InsertLocation.replace);case 44:t.next=49;break;case 46:t.prev=46,t.t1=t.catch(9),"AbortError"!==t.t1.name&&console.error("Error:",t.t1);case 49:return t.prev=49,document.getElementById("cancel-request").style.display="none",document.getElementById("replace-text").style.display="inline-block",currentController=null,t.finish(49);case 54:case"end":return t.stop()}}),t,null,[[9,46,49,54],[25,34]])})));return function(e){return t.apply(this,arguments)}}());case 1:case"end":return t.stop()}}),t)}))),_replaceSelectedText.apply(this,arguments)}function applyFormattedText(t,e,n,r){return _applyFormattedText.apply(this,arguments)}function _applyFormattedText(){return(_applyFormattedText=_asyncToGenerator(_regeneratorRuntime().mark((function t(e,n,r,o){var a,i,c,l,s,u;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n.clear(),t.next=3,e.sync();case 3:a=r.split(/\s+/),i=0,c=0;case 6:if(!(c<a.length)){t.next=16;break}return l=a[c],s=n.insertText(l,Word.InsertLocation.end),i<o.length&&(void 0!==(u=o[i]).bold&&(s.font.bold=u.bold),void 0!==u.italic&&(s.font.italic=u.italic),u.color&&(s.font.color=u.color),u.highlightColor&&(s.font.highlightColor=u.highlightColor),u.size&&(s.font.size=u.size),u.hyperlink&&(s.hyperlink=u.hyperlink),u.footnoteText&&s.insertFootnote().body.insertText(u.footnoteText,Word.InsertLocation.start),i++),c<a.length-1&&n.insertText(" ",Word.InsertLocation.end),t.next=13,e.sync();case 13:c++,t.next=6;break;case 16:case"end":return t.stop()}}),t)})))).apply(this,arguments)}function cancelRequest(){currentController&&(currentController.abort(),document.getElementById("cancel-request").style.display="none",document.getElementById("replace-text").style.display="inline-block")}function applyTricks(){return _applyTricks.apply(this,arguments)}function _applyTricks(){return _applyTricks=_asyncToGenerator(_regeneratorRuntime().mark((function t(){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var n,r,o,a,i,c;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(n=e.document.getSelection()).load("text"),t.next=4,e.sync();case 4:if(!((r=n.text.trim().split(/\s+/)).length<4)){t.next=8;break}return console.log("Выделите минимум 4 слова"),t.abrupt("return");case 8:return(o=n.getRange()).clear(),t.next=12,e.sync();case 12:a=0;case 13:if(!(a<r.length)){t.next=33;break}return i=r[a],(c=o.insertText(i+" ",Word.InsertLocation.end)).font.bold=!1,c.font.italic=!1,c.font.color="black",c.font.highlightColor="white",c.font.size=11,0===a&&(c.font.bold=!0),1===a&&(c.font.italic=!0),2===a&&(c.hyperlink="https://www.example.com"),3===a&&(c.font.color="red"),4===a&&(c.font.highlightColor="yellow"),5===a&&(c.font.size=16),a===r.length-1&&c.insertFootnote().body.insertText("Это пример сноски",Word.InsertLocation.start),t.next=30,e.sync();case 30:a++,t.next=13;break;case 33:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),_applyTricks.apply(this,arguments)}function formatSpecialWords(t,e,n,r){var o=[],a=[];return t.forEach((function(t,i){var c={text:t.text},l=!1,s={};!0!==t.bold&&null!==t.bold||(s.bold=!0,l=!0),t.italic&&(s.italic=!0,l=!0),t.underline&&(s.underline=!0,l=!0),t.color!==e&&""!==t.color&&(s.color=t.color,l=!0),t.highlightColor&&null!==t.highlightColor&&(s.highlight=t.highlightColor,l=!0),t.fontName!==n&&""!==t.fontName&&(s.font=t.fontName,l=!0),t.fontSize!=r&&""!==t.fontSize&&null!==t.fontSize&&(s.size=t.fontSize,l=!0),t.hyperlink&&(s.hyperlink=t.hyperlink,l=!0),Object.keys(s).length>0&&(c.formatting=s),t.hasFootnote&&(c.footnoteId=i+1,a.push({id:i+1,text:t.footnoteText,relatedWordIndex:o.length}),l=!0),l&&o.push(c)})),{formattedText:o,footnotes:a,baseFormatting:{color:e,font:n,size:r}}}function showSelectedText(){return _showSelectedText.apply(this,arguments)}function _showSelectedText(){return _showSelectedText=_asyncToGenerator(_regeneratorRuntime().mark((function t(){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var n,r,o,a,i,c,l,s,u,p,d,m,f,h,g,y,x,v;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return n=e.document.getSelection(),(r=n.getTextRanges([" "],!1)).load("items"),t.next=5,e.sync();case 5:o=[],a={},i={},c={},l=0;case 10:if(!(l<r.items.length)){t.next=36;break}return(s=r.items[l]).load("text, hyperlink"),s.font.load("bold, italic, underline, color, highlightColor, name, size"),(u=s.footnotes).load("items"),t.next=18,e.sync();case 18:if(!s.text.trim()){t.next=33;break}if(p=s.font.color||"default",d=s.font.name||"default",m=s.font.size||"default",a[p]=(a[p]||0)+1,i[d]=(i[d]||0)+1,c[m]=(c[m]||0)+1,f="",!(u.items.length>0)){t.next=32;break}return(h=u.items[0]).body.load("text"),t.next=31,e.sync();case 31:f=h.body.text;case 32:o.push({text:s.text.trim(),bold:s.font.bold,italic:s.font.italic,underline:"None"!==s.font.underline,color:p,highlightColor:s.font.highlightColor,hasFootnote:u.items.length>0,footnoteText:f,fontName:d,fontSize:m,hyperlink:s.hyperlink?s.hyperlink.address:null,isFirstWord:0===l,isSecondWord:1===l,isThirdWord:2===l,isLastWord:l===r.items.length-1});case 33:l++,t.next=10;break;case 36:g=Object.entries(a).reduce((function(t,e){return t[1]>e[1]?t:e}))[0],y=Object.entries(i).reduce((function(t,e){return t[1]>e[1]?t:e}))[0],x=Object.entries(c).reduce((function(t,e){return t[1]>e[1]?t:e}))[0],displayWordsInfo(o,g,y,x),v=formatSpecialWords(o,g,y,x),console.log(v);case 42:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),_showSelectedText.apply(this,arguments)}function displayWordsInfo(t,e,n,r){var o=document.getElementById("selected-text-info"),a="<h3>Анализ выделенного текста:</h3>";t.forEach((function(t,o){a+="<h4>Слово ".concat(o+1,": ").concat(t.text,"</h4>"),a+="<p>\n      ".concat(t.bold?"✅":"❌"," Жирный").concat(t.isFirstWord?" (Первое слово)":"","<br>\n      ").concat(t.italic?"✅":"❌"," Курсив").concat(t.isSecondWord?" (Второе слово)":"","<br>\n      ").concat(t.underline?"✅":"❌"," Подчеркнутый<br>\n      ").concat(t.color!==e&&"default"!==t.color?"✅ Особый цвет: ".concat(t.color):"❌ Основной цвет","<br>\n      ").concat(t.highlightColor?"✅ Выделение: ".concat(t.highlightColor):"❌ Без выделения","<br>\n      ").concat(t.hasFootnote?"✅ Есть сноска: ".concat(t.footnoteText).concat(t.isLastWord?" (Последнее слово)":""):"❌ Без сноски","<br>\n      ").concat(t.fontName!==n&&"default"!==t.fontName?"✅ Особый шрифт: ".concat(t.fontName):"❌ Основной шрифт: ".concat(n),"<br>\n      ").concat(t.fontSize!=r&&"default"!==t.fontSize?"✅ Особый размер: ".concat(t.fontSize):"❌ Основной размер: ".concat(r),"<br>\n      ").concat(t.hyperlink?"✅ Ссылка: ".concat(t.hyperlink).concat(t.isThirdWord?" (Третье слово)":""):"❌ Без ссылки","\n    </p>")})),o.innerHTML=a}function repeatStyles(){return _repeatStyles.apply(this,arguments)}function _repeatStyles(){return _repeatStyles=_asyncToGenerator(_regeneratorRuntime().mark((function t(){return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.next=2,Word.run(function(){var t=_asyncToGenerator(_regeneratorRuntime().mark((function t(e){var n,r,o,a,i,c,l,s,u,p;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(n=e.document.getSelection()).load("text"),t.next=4,e.sync();case 4:return r=n.text,t.next=7,getWordsInfo(e,n);case 7:return o=t.sent,a=getBaseFormatting(o),o=(o=optimizeWordsInfo(o,a)).filter((function(t){return Object.keys(t).length>1})),console.log(o),console.log("Base formatting:",a),n.clear(),t.next=16,e.sync();case 16:i=r.split(/\s+/),c=0,l=0;case 19:if(!(l<i.length)){t.next=34;break}return s=i[l],(u=n.insertText(s,Word.InsertLocation.end)).font.bold=a.bold,u.font.italic=a.italic,u.font.color=a.color,u.font.highlightColor=a.highlightColor,u.font.size=a.size,c<o.length&&o[c].text===s&&(void 0!==(p=o[c]).bold&&(u.font.bold=p.bold),void 0!==p.italic&&(u.font.italic=p.italic),p.color&&(u.font.color=p.color),p.highlightColor&&(u.font.highlightColor=p.highlightColor),p.size&&(u.font.size=p.size),p.hyperlink&&(u.hyperlink=p.hyperlink),p.footnoteText&&(u.insertFootnote(p.footnoteText.replace(""," "),Word.InsertLocation.start),s=s.replace(""," ")),c++),l<i.length-1&&n.insertText(" ",Word.InsertLocation.end),t.next=31,e.sync();case 31:l++,t.next=19;break;case 34:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}());case 2:case"end":return t.stop()}}),t)}))),_repeatStyles.apply(this,arguments)}function getBaseFormatting(t){return{bold:getMostCommonValue(t,"bold"),italic:getMostCommonValue(t,"italic"),color:getMostCommonValue(t,"color"),highlightColor:getMostCommonValue(t,"highlightColor"),size:getMostCommonValue(t,"size")}}function optimizeWordsInfo(t,e){return t.map((function(t){var n={text:t.text};return t.bold!==e.bold&&(n.bold=t.bold),t.italic!==e.italic&&(n.italic=t.italic),t.color!==e.color&&(n.color=t.color),t.highlightColor!==e.highlightColor&&(n.highlightColor=t.highlightColor),t.size!==e.size&&(n.size=t.size),t.hasFootnote&&(n.footnoteText=t.footnoteText),t.hyperlink&&(n.hyperlink=t.hyperlink),n}))}function getWordsInfo(t,e){return _getWordsInfo.apply(this,arguments)}function _getWordsInfo(){return(_getWordsInfo=_asyncToGenerator(_regeneratorRuntime().mark((function t(e,n){var r,o,a,i,c,l,s,u;return _regeneratorRuntime().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(r=n.getTextRanges([" "],!1)).load("items"),t.next=4,e.sync();case 4:o=[],a=_createForOfIteratorHelper(r.items),t.prev=6,a.s();case 8:if((i=a.n()).done){t.next=27;break}return(c=i.value).load("text, hyperlink"),c.font.load("bold, italic, color, highlightColor, size"),(l=c.footnotes).load("items"),t.next=16,e.sync();case 16:if(!c.text.trim()){t.next=25;break}if(s="",!(l.items.length>0)){t.next=24;break}return(u=l.items[0]).body.load("text"),t.next=23,e.sync();case 23:s=u.body.text;case 24:o.push({text:c.text.trim(),bold:c.font.bold,italic:c.font.italic,color:c.font.color||"black",highlightColor:c.font.highlightColor||"white",size:c.font.size||"11",hasFootnote:l.items.length>0,footnoteText:s,hyperlink:c.hyperlink?c.hyperlink:null});case 25:t.next=8;break;case 27:t.next=32;break;case 29:t.prev=29,t.t0=t.catch(6),a.e(t.t0);case 32:return t.prev=32,a.f(),t.finish(32);case 35:return t.abrupt("return",o);case 36:case"end":return t.stop()}}),t,null,[[6,29,32,35]])})))).apply(this,arguments)}function getMostCommonValue(t,e){var n,r={},o=0;return t.forEach((function(t){var a=t[e];r[a]=(r[a]||0)+1,r[a]>o&&(o=r[a],n=a)})),n}function initPrompts(){document.getElementById("prompts").innerHTML='\n        <div class="mt-3">\n            <input type="text" id="prompt-name" class="form-control" placeholder="Название промпта">\n            <select id="prompt-model" class="form-select mt-2"></select>\n            <textarea id="prompt-text" class="form-control mt-2" rows="5" placeholder="Текст промпта"></textarea>\n            <button id="save-prompt" class="btn btn-primary mt-2">Сохранить промпт</button>\n        </div>\n        <div id="prompt-list" class="mt-3"></div>\n    ',$("#prompt-model").select2({width:"100%",placeholder:"Выберите модель"}).on("select2:opening",(function(){setTimeout((function(){$(".select2-search__field").get(0).focus()}),0)})),document.getElementById("save-prompt").onclick=savePrompt,updatePromptList(),populateModelSelect()}function populateModelSelect(){var t=$("#prompt-model");JSON.parse(localStorage.getItem("models")||"[]").forEach((function(e){var n=new Option(e.name,e.id,!1,!1);t.append(n)})),t.trigger("change")}function savePrompt(){var t=document.getElementById("prompt-name").value,e=$("#prompt-model").val(),n=document.getElementById("prompt-text").value;if(t&&e&&n){var r=getPrompts();r.push({name:t,model:e,text:n}),savePrompts(r),updatePromptList(),updatePromptSelect(),document.getElementById("prompt-name").value="",$("#prompt-model").val(null).trigger("change"),document.getElementById("prompt-text").value=""}}function updatePromptList(){var t=document.getElementById("prompt-list"),e=getPrompts();t.innerHTML="",e.forEach((function(e,n){var r=document.createElement("div");r.className="mt-2",r.innerHTML="\n            <strong>".concat(e.name,"</strong> (").concat(e.model,')\n            <button class="btn btn-sm btn-primary edit-prompt" data-index="').concat(n,'">Редактировать</button>\n            <button class="btn btn-sm btn-info duplicate-prompt" data-index="').concat(n,'">Дубликат</button>\n            <button class="btn btn-sm btn-danger delete-prompt" data-index="').concat(n,'">Удалить</button>\n        '),t.appendChild(r)})),t.querySelectorAll(".delete-prompt").forEach((function(t){t.onclick=function(){return deletePrompt(t.dataset.index)}})),t.querySelectorAll(".edit-prompt").forEach((function(t){t.onclick=function(){return editPrompt(t.dataset.index)}})),t.querySelectorAll(".duplicate-prompt").forEach((function(t){t.onclick=function(){return duplicatePrompt(t.dataset.index)}}))}function deletePrompt(t){var e=getPrompts();e.splice(t,1),savePrompts(e),updatePromptList(),updatePromptSelect()}function editPrompt(t){var e=getPrompts()[t];document.getElementById("prompt-name").value=e.name,$("#prompt-model").val(e.model).trigger("change"),document.getElementById("prompt-text").value=e.text,document.getElementById("save-prompt").textContent="Обновить промпт",document.getElementById("save-prompt").onclick=function(){updatePrompt(t),document.getElementById("save-prompt").textContent="Сохранить промпт",document.getElementById("save-prompt").onclick=savePrompt}}function updatePrompt(t){var e=document.getElementById("prompt-name").value,n=$("#prompt-model").val(),r=document.getElementById("prompt-text").value;if(e&&n&&r){var o=getPrompts();o[t]={name:e,model:n,text:r},savePrompts(o),updatePromptList(),updatePromptSelect(),document.getElementById("prompt-name").value="",$("#prompt-model").val(null).trigger("change"),document.getElementById("prompt-text").value=""}}function duplicatePrompt(t){var e=getPrompts()[t];document.getElementById("prompt-name").value=e.name+" (копия)",$("#prompt-model").val(null).trigger("change"),document.getElementById("prompt-text").value=e.text,document.getElementById("save-prompt").textContent="Сохранить промпт",document.getElementById("save-prompt").onclick=savePrompt}Office.onReady((function(t){t.host===Office.HostType.Word&&(initHome(),initPrompts(),initSettings())}));