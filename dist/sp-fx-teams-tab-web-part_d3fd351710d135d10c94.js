define("93df0ba1-9407-4b9b-b8b7-a764f4447cf0_0.0.1",["@microsoft/sp-property-pane","@microsoft/sp-lodash-subset","@microsoft/sp-core-library","@microsoft/sp-webpart-base","react","SpFxTeamsTabWebPartStrings","react-dom"],function(n,a,i,r,o,s,c){return function(e){var t={};function n(a){if(t[a])return t[a].exports;var i=t[a]={i:a,l:!1,exports:{}};return e[a].call(i.exports,i,i.exports,n),i.l=!0,i.exports}return n.m=e,n.c=t,n.d=function(e,t,a){n.o(e,t)||Object.defineProperty(e,t,{enumerable:!0,get:a})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,t){if(1&t&&(e=n(e)),8&t)return e;if(4&t&&"object"==typeof e&&e&&e.__esModule)return e;var a=Object.create(null);if(n.r(a),Object.defineProperty(a,"default",{enumerable:!0,value:e}),2&t&&"string"!=typeof e)for(var i in e)n.d(a,i,function(t){return e[t]}.bind(null,i));return a},n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,"a",t),t},n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},n.p="",function(){var e,t=document.getElementsByTagName("script"),a=/sp-fx-teams-tab-web-part_d3fd351710d135d10c94\.js/i;if(t&&t.length)for(var i=0;i<t.length;i++)if(t[i]){var r=t[i].getAttribute("src");if(r&&r.match(a)){e=r.substring(0,r.lastIndexOf("/")+1);break}}if(!e)for(var o in window.__setWebpackPublicPathLoaderSrcRegistry__)if(o&&o.match(a)){e=o.substring(0,o.lastIndexOf("/")+1);break}n.p=e}(),n(n.s="m0/m")}({"26ea":function(e,t){e.exports=n},JPst:function(e,t,n){"use strict";e.exports=function(e){var t=[];return t.toString=function(){return this.map(function(t){var n=function(e,t){var n,a,i,r=e[1]||"",o=e[3];if(!o)return r;if(t&&"function"==typeof btoa){var s=(n=o,a=btoa(unescape(encodeURIComponent(JSON.stringify(n)))),i="sourceMappingURL=data:application/json;charset=utf-8;base64,".concat(a),"/*# ".concat(i," */")),c=o.sources.map(function(e){return"/*# sourceURL=".concat(o.sourceRoot).concat(e," */")});return[r].concat(c).concat([s]).join("\n")}return[r].join("\n")}(t,e);return t[2]?"@media ".concat(t[2],"{").concat(n,"}"):n}).join("")},t.i=function(e,n){"string"==typeof e&&(e=[[null,e,""]]);for(var a={},i=0;i<this.length;i++){var r=this[i][0];null!=r&&(a[r]=!0)}for(var o=0;o<e.length;o++){var s=e[o];null!=s[0]&&a[s[0]]||(n&&!s[2]?s[2]=n:n&&(s[2]="(".concat(s[2],") and (").concat(n,")")),t.push(s))}},t}},ODIj:function(e,t,n){e.exports=n.p+"welcome-light_a2dcb0d64c8d6e80cf49f607eeb17723.png"},OhSF:function(e,t,n){var a=n("x15E"),i=n("ruv1");"string"==typeof a&&(a=[[e.i,a]]);for(var r=0;r<a.length;r++)i.loadStyles(a[r][1],!0);a.locals&&(e.exports=a.locals)},Pk8u:function(e,t){e.exports=a},PqoK:function(e,t,n){e.exports=n.p+"welcome-dark_bc81978d2f17e05985eed81d09531178.png"},UWqr:function(e,t){e.exports=i},br4S:function(e,t){e.exports=r},cDcd:function(e,t){e.exports=o},"e/AI":function(e,t){e.exports=s},faye:function(e,t){e.exports=c},"m0/m":function(e,t,n){"use strict";n.r(t);var a=n("cDcd"),i=n("faye"),r=n("UWqr"),o=n("26ea"),s=n("br4S"),c=n("e/AI");n("OhSF");var d,l=n("Pk8u"),u=(d=function(e,t){return d=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])},d(e,t)},function(e,t){function n(){this.constructor=e}d(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),f=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return u(t,e),t.prototype.onInit=function(){var e=this;return new Promise(function(t,n){e.context.microsoftTeams?e.context.microsoftTeams.getContext(function(n){e.teamsContext=n,t()}):t()})},t.prototype.render=function(){var e=this.teamsContext?"Teams":"SharePoint",t=this.teamsContext?"Teams "+this.teamsContext.teamName:"SharePoint workbench",i=this.props,r=i.description,o=i.isDarkTheme,s=i.environmentMessage,c=i.hasTeamsContext,d=i.userDisplayName;return a.createElement("section",{className:"spFxTeamsTab_508c780f "+(c?"teams_508c780f":"")},a.createElement("div",{className:"welcome_508c780f"},a.createElement("img",{alt:"",src:n(o?"PqoK":"ODIj"),className:"welcomeImage_508c780f"}),a.createElement("h2",null,"Well done, ",Object(l.escape)(d),"!"),a.createElement("div",null,s),a.createElement("div",null,"Web part property value: ",a.createElement("strong",null,Object(l.escape)(r)))),a.createElement("div",null,a.createElement("h3",null,"Welcome to ",e),a.createElement("p",null,"Currently in the context of following ",t),a.createElement("h4",null,"Learn more about SPFx development:"),a.createElement("ul",{className:"links_508c780f"},a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/spfx",target:"_blank"},"SharePoint Framework Overview")),a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/spfx-yeoman-graph",target:"_blank"},"Use Microsoft Graph in your solution")),a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/spfx-yeoman-teams",target:"_blank"},"Build for Microsoft Teams using SharePoint Framework")),a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/spfx-yeoman-viva",target:"_blank"},"Build for Microsoft Viva Connections using SharePoint Framework")),a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/spfx-yeoman-store",target:"_blank"},"Publish SharePoint Framework applications to the marketplace")),a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/spfx-yeoman-api",target:"_blank"},"SharePoint Framework API reference")),a.createElement("li",null,a.createElement("a",{href:"https://aka.ms/m365pnp",target:"_blank"},"Microsoft 365 Developer Community")))))},t}(a.Component),p=f,m=function(){var e=function(t,n){return e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)t.hasOwnProperty(n)&&(e[n]=t[n])},e(t,n)};return function(t,n){function a(){this.constructor=t}e(t,n),t.prototype=null===n?Object.create(n):(a.prototype=n.prototype,new a)}}(),_=function(e){function t(){var t=null!==e&&e.apply(this,arguments)||this;return t._isDarkTheme=!1,t._environmentMessage="",t}return m(t,e),t.prototype.onInit=function(){return this._environmentMessage=this._getEnvironmentMessage(),e.prototype.onInit.call(this)},t.prototype.render=function(){var e=a.createElement(p,{description:this.properties.description,isDarkTheme:this._isDarkTheme,environmentMessage:this._environmentMessage,hasTeamsContext:!!this.context.sdks.microsoftTeams,userDisplayName:this.context.pageContext.user.displayName});i.render(e,this.domElement)},t.prototype._getEnvironmentMessage=function(){return this.context.sdks.microsoftTeams?this.context.isServedFromLocalhost?c.AppLocalEnvironmentTeams:c.AppTeamsTabEnvironment:this.context.isServedFromLocalhost?c.AppLocalEnvironmentSharePoint:c.AppSharePointEnvironment},t.prototype.onThemeChanged=function(e){if(e){this._isDarkTheme=!!e.isInverted;var t=e.semanticColors;this.domElement.style.setProperty("--bodyText",t.bodyText),this.domElement.style.setProperty("--link",t.link),this.domElement.style.setProperty("--linkHovered",t.linkHovered)}},t.prototype.onDispose=function(){i.unmountComponentAtNode(this.domElement)},Object.defineProperty(t.prototype,"dataVersion",{get:function(){return r.Version.parse("1.0")},enumerable:!1,configurable:!0}),t.prototype.getPropertyPaneConfiguration=function(){return{pages:[{header:{description:c.PropertyPaneDescription},groups:[{groupName:c.BasicGroupName,groupFields:[Object(o.PropertyPaneTextField)("description",{label:c.DescriptionFieldLabel})]}]}]}},t}(s.BaseClientSideWebPart);t.default=_},ruv1:function(e,t,n){"use strict";(function(e){var n=this&&this.__assign||function(){return n=Object.assign||function(e){for(var t,n=1,a=arguments.length;n<a;n++)for(var i in t=arguments[n])Object.prototype.hasOwnProperty.call(t,i)&&(e[i]=t[i]);return e},n.apply(this,arguments)};Object.defineProperty(t,"__esModule",{value:!0}),t.splitStyles=t.detokenize=t.clearStyles=t.loadTheme=t.flush=t.configureRunMode=t.configureLoadStyles=t.loadStyles=void 0;var a,i="undefined"==typeof window?e:window,r=i&&i.CSPSettings&&i.CSPSettings.nonce,o=((a=i.__themeState__||{theme:void 0,lastStyleElement:void 0,registeredStyles:[]}).runState||(a=n(n({},a),{perf:{count:0,duration:0},runState:{flushTimer:0,mode:0,buffer:[]}})),a.registeredThemableStyles||(a=n(n({},a),{registeredThemableStyles:[]})),i.__themeState__=a,a),s=/[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g,c=function(){return"undefined"!=typeof performance&&performance.now?performance.now():Date.now()};function d(e){var t=c();e();var n=c();o.perf.duration+=n-t}function l(){d(function(){var e=o.runState.buffer.slice();o.runState.buffer=[];var t=[].concat.apply([],e);t.length>0&&u(t)})}function u(e,t){o.loadStyles?o.loadStyles(m(e).styleString,e):function(e){if("undefined"!=typeof document){var t=document.getElementsByTagName("head")[0],n=document.createElement("style"),a=m(e),i=a.styleString,s=a.themable;n.setAttribute("data-load-themed-styles","true"),r&&n.setAttribute("nonce",r),n.appendChild(document.createTextNode(i)),o.perf.count++,t.appendChild(n);var c=document.createEvent("HTMLEvents");c.initEvent("styleinsert",!0,!1),c.args={newStyle:n},document.dispatchEvent(c);var d={styleElement:n,themableStyle:e};s?o.registeredThemableStyles.push(d):o.registeredStyles.push(d)}}(e)}function f(e){void 0===e&&(e=3),3!==e&&2!==e||(p(o.registeredStyles),o.registeredStyles=[]),3!==e&&1!==e||(p(o.registeredThemableStyles),o.registeredThemableStyles=[])}function p(e){e.forEach(function(e){var t=e&&e.styleElement;t&&t.parentElement&&t.parentElement.removeChild(t)})}function m(e){var t=o.theme,n=!1;return{styleString:(e||[]).map(function(e){var a=e.theme;if(a){n=!0;var i=t?t[a]:void 0,r=e.defaultValue||"inherit";return t&&!i&&console,i||r}return e.rawString}).join(""),themable:n}}function _(e){var t=[];if(e){for(var n=0,a=void 0;a=s.exec(e);){var i=a.index;i>n&&t.push({rawString:e.substring(n,i)}),t.push({theme:a[1],defaultValue:a[2]}),n=s.lastIndex}t.push({rawString:e.substring(n)})}return t}t.loadStyles=function(e,t){void 0===t&&(t=!1),d(function(){var n=Array.isArray(e)?e:_(e),a=o.runState,i=a.mode,r=a.buffer,s=a.flushTimer;t||1===i?(r.push(n),s||(o.runState.flushTimer=setTimeout(function(){o.runState.flushTimer=0,l()},0))):u(n)})},t.configureLoadStyles=function(e){o.loadStyles=e},t.configureRunMode=function(e){o.runState.mode=e},t.flush=l,t.loadTheme=function(e){o.theme=e,function(){if(o.theme){for(var e=[],t=0,n=o.registeredThemableStyles;t<n.length;t++){var a=n[t];e.push(a.themableStyle)}e.length>0&&(f(1),u([].concat.apply([],e)))}}()},t.clearStyles=f,t.detokenize=function(e){return e&&(e=m(_(e)).styleString),e},t.splitStyles=_}).call(this,n("yLpj"))},x15E:function(e,t,n){(e.exports=n("JPst")(!1)).push([e.i,'.spFxTeamsTab_508c780f{overflow:hidden;padding:1em;color:"[theme:bodyText, default: #323130]";color:var(--bodyText)}.spFxTeamsTab_508c780f.teams_508c780f{font-family:Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif}.welcome_508c780f{text-align:center}.welcomeImage_508c780f{width:100%;max-width:420px}.links_508c780f a{text-decoration:none;color:"[theme:link, default:#03787c]";color:var(--link)}.links_508c780f a:hover{text-decoration:underline;color:"[theme:linkHovered, default: #014446]";color:var(--linkHovered)}',""])},yLpj:function(e,t){var n;n=function(){return this}();try{n=n||new Function("return this")()}catch(e){"object"==typeof window&&(n=window)}e.exports=n}})});