"use strict";(self.webpackChunkreact_contoso=self.webpackChunkreact_contoso||[]).push([[622],{768:function(n,e,t){t.d(e,{Ad:function(){return X}});var i=t(9249),o=t(7371),a=t(753),r=t(3069),s=t(5058),d=t(5754),c=t(6906),u=t(6647),l=t(3369),p=t(9763),f=t(9604),m=t(7450),h=["input","select","textarea","a[href]","button","[tabindex]:not(slot)","audio[controls]","video[controls]",'[contenteditable]:not([contenteditable="false"])',"details>summary:first-of-type","details"],b=h.join(","),g="undefined"===typeof Element,v=g?function(){}:Element.prototype.matches||Element.prototype.msMatchesSelector||Element.prototype.webkitMatchesSelector,y=!g&&Element.prototype.getRootNode?function(n){return n.getRootNode()}:function(n){return n.ownerDocument},w=function(n,e){return n.tabIndex<0&&(e||/^(AUDIO|VIDEO|DETAILS)$/.test(n.tagName)||n.isContentEditable)&&isNaN(parseInt(n.getAttribute("tabindex"),10))?0:n.tabIndex},k=function(n){return"INPUT"===n.tagName},F=function(n){return function(n){return k(n)&&"radio"===n.type}(n)&&!function(n){if(!n.name)return!0;var e,t=n.form||y(n),i=function(n){return t.querySelectorAll('input[type="radio"][name="'+n+'"]')};if("undefined"!==typeof window&&"undefined"!==typeof window.CSS&&"function"===typeof window.CSS.escape)e=i(window.CSS.escape(n.name));else try{e=i(n.name)}catch(a){return console.error("Looks like you have a radio button with a name attribute containing invalid CSS selector characters and need the CSS.escape polyfill: %s",a.message),!1}var o=function(n,e){for(var t=0;t<n.length;t++)if(n[t].checked&&n[t].form===e)return n[t]}(e,n.form);return!o||o===n}(n)},x=function(n){var e=n.getBoundingClientRect(),t=e.width,i=e.height;return 0===t&&0===i},T=function(n,e){return!(e.disabled||function(n){return k(n)&&"hidden"===n.type}(e)||function(n,e){var t=e.displayCheck,i=e.getShadowRoot;if("hidden"===getComputedStyle(n).visibility)return!0;var o=v.call(n,"details>summary:first-of-type")?n.parentElement:n;if(v.call(o,"details:not([open]) *"))return!0;var a=y(n).host,r=(null===a||void 0===a?void 0:a.ownerDocument.contains(a))||n.ownerDocument.contains(n);if(t&&"full"!==t){if("non-zero-area"===t)return x(n)}else{if("function"===typeof i){for(var s=n;n;){var d=n.parentElement,c=y(n);if(d&&!d.shadowRoot&&!0===i(d))return x(n);n=n.assignedSlot?n.assignedSlot:d||c===n.ownerDocument?d:c.host}n=s}if(r)return!n.getClientRects().length}return!1}(e,n)||function(n){return"DETAILS"===n.tagName&&Array.prototype.slice.apply(n.children).some((function(n){return"SUMMARY"===n.tagName}))}(e)||function(n){if(/^(INPUT|BUTTON|SELECT|TEXTAREA)$/.test(n.tagName))for(var e=n.parentElement;e;){if("FIELDSET"===e.tagName&&e.disabled){for(var t=0;t<e.children.length;t++){var i=e.children.item(t);if("LEGEND"===i.tagName)return!!v.call(e,"fieldset[disabled] *")||!i.contains(n)}return!0}e=e.parentElement}return!1}(e))},E=function(n,e){return!(F(e)||w(e)<0||!T(n,e))},S=function(n,e){if(e=e||{},!n)throw new Error("No node provided");return!1!==v.call(n,b)&&E(e,n)},C=function(n){(0,d.Z)(t,n);var e=(0,c.Z)(t);function t(){var n;return(0,i.Z)(this,t),(n=e.apply(this,arguments)).modal=!0,n.hidden=!1,n.trapFocus=!0,n.trapFocusChanged=function(){n.$fastController.isConnected&&n.updateTrapFocus()},n.isTrappingFocus=!1,n.handleDocumentKeydown=function(e){if(!e.defaultPrevented&&!n.hidden)switch(e.key){case m.CX:n.dismiss(),e.preventDefault();break;case m.oM:n.handleTabKeyDown(e)}},n.handleDocumentFocus=function(e){!e.defaultPrevented&&n.shouldForceFocus(e.target)&&(n.focusFirstElement(),e.preventDefault())},n.handleTabKeyDown=function(e){if(n.trapFocus&&!n.hidden){var t=n.getTabQueueBounds();if(0!==t.length)return 1===t.length?(t[0].focus(),void e.preventDefault()):void(e.shiftKey&&e.target===t[0]?(t[t.length-1].focus(),e.preventDefault()):e.shiftKey||e.target!==t[t.length-1]||(t[0].focus(),e.preventDefault()))}},n.getTabQueueBounds=function(){return t.reduceTabbableItems([],(0,a.Z)(n))},n.focusFirstElement=function(){var e=n.getTabQueueBounds();e.length>0?e[0].focus():n.dialog instanceof HTMLElement&&n.dialog.focus()},n.shouldForceFocus=function(e){return n.isTrappingFocus&&!n.contains(e)},n.shouldTrapFocus=function(){return n.trapFocus&&!n.hidden},n.updateTrapFocus=function(e){var t=void 0===e?n.shouldTrapFocus():e;t&&!n.isTrappingFocus?(n.isTrappingFocus=!0,document.addEventListener("focusin",n.handleDocumentFocus),l.SO.queueUpdate((function(){n.shouldForceFocus(document.activeElement)&&n.focusFirstElement()}))):!t&&n.isTrappingFocus&&(n.isTrappingFocus=!1,document.removeEventListener("focusin",n.handleDocumentFocus))},n}return(0,o.Z)(t,[{key:"dismiss",value:function(){this.$emit("dismiss"),this.$emit("cancel")}},{key:"show",value:function(){this.hidden=!1}},{key:"hide",value:function(){this.hidden=!0,this.$emit("close")}},{key:"connectedCallback",value:function(){(0,r.Z)((0,s.Z)(t.prototype),"connectedCallback",this).call(this),document.addEventListener("keydown",this.handleDocumentKeydown),this.notifier=p.y$.getNotifier(this),this.notifier.subscribe(this,"hidden"),this.updateTrapFocus()}},{key:"disconnectedCallback",value:function(){(0,r.Z)((0,s.Z)(t.prototype),"disconnectedCallback",this).call(this),document.removeEventListener("keydown",this.handleDocumentKeydown),this.updateTrapFocus(!1),this.notifier.unsubscribe(this,"hidden")}},{key:"handleChange",value:function(n,e){if("hidden"===e)this.updateTrapFocus()}}],[{key:"reduceTabbableItems",value:function(n,e){return"-1"===e.getAttribute("tabindex")?n:S(e)||t.isFocusableFastElement(e)&&t.hasTabbableShadow(e)?(n.push(e),n):e.childElementCount?n.concat(Array.from(e.children).reduce(t.reduceTabbableItems,[])):n}},{key:"isFocusableFastElement",value:function(n){var e,t;return!!(null===(t=null===(e=n.$fastController)||void 0===e?void 0:e.definition.shadowOptions)||void 0===t?void 0:t.delegatesFocus)}},{key:"hasTabbableShadow",value:function(n){var e,t;return Array.from(null!==(t=null===(e=n.shadowRoot)||void 0===e?void 0:e.querySelectorAll("*"))&&void 0!==t?t:[]).some((function(n){return S(n)}))}}]),t}(t(9350).I);(0,u.gn)([(0,f.Lj)({mode:"boolean"})],C.prototype,"modal",void 0),(0,u.gn)([(0,f.Lj)({mode:"boolean"})],C.prototype,"hidden",void 0),(0,u.gn)([(0,f.Lj)({attribute:"trap-focus",mode:"boolean"})],C.prototype,"trapFocus",void 0),(0,u.gn)([(0,f.Lj)({attribute:"aria-describedby"})],C.prototype,"ariaDescribedby",void 0),(0,u.gn)([(0,f.Lj)({attribute:"aria-labelledby"})],C.prototype,"ariaLabelledby",void 0),(0,u.gn)([(0,f.Lj)({attribute:"aria-label"})],C.prototype,"ariaLabel",void 0);var D,L,Z,I=t(1171),N=t(982),A=t(7376),j=t(3025),R=t(3032),B=t(4901),H=t(2132),X=C.compose({baseName:"dialog",template:function(n,e){return(0,N.d)(D||(D=(0,I.Z)(['\n    <div class="positioning-region" part="positioning-region">\n        ','\n        <div\n            role="dialog"\n            tabindex="-1"\n            class="control"\n            part="control"\n            aria-modal="','"\n            aria-describedby="','"\n            aria-labelledby="','"\n            aria-label="','"\n            ',"\n        >\n            <slot></slot>\n        </div>\n    </div>\n"])),(0,A.g)((function(n){return n.modal}),(0,N.d)(L||(L=(0,I.Z)(['\n                <div\n                    class="overlay"\n                    part="overlay"\n                    role="presentation"\n                    @click="','"\n                ></div>\n            '])),(function(n){return n.dismiss()}))),(function(n){return n.modal}),(function(n){return n.ariaDescribedby}),(function(n){return n.ariaLabelledby}),(function(n){return n.ariaLabel}),(0,j.i)("dialog"))},styles:function(n,e){return(0,R.i)(Z||(Z=(0,I.Z)(["\n  :host([hidden]) {\n    display: none;\n  }\n\n  :host {\n    --dialog-height: 480px;\n    --dialog-width: 640px;\n    display: block;\n  }\n\n  .overlay {\n    position: fixed;\n    top: 0;\n    left: 0;\n    right: 0;\n    bottom: 0;\n    background: rgba(0, 0, 0, 0.3);\n    touch-action: none;\n  }\n\n  .positioning-region {\n    display: flex;\n    justify-content: center;\n    position: fixed;\n    top: 0;\n    bottom: 0;\n    left: 0;\n    right: 0;\n    overflow: auto;\n  }\n\n  .control {\n    box-shadow: ",";\n    margin-top: auto;\n    margin-bottom: auto;\n    border-radius: calc("," * 1px);\n    width: var(--dialog-width);\n    height: var(--dialog-height);\n    background: ",";\n    z-index: 1;\n    border: calc("," * 1px) solid transparent;\n  }\n"])),B.CJ,H.rSr,H.IfY,H.Han)}})},4622:function(n,e,t){t.d(e,{Jh:function(){return F}});var i,o,a,r,s,d=t(7371),c=t(9249),u=t(5754),l=t(6906),p=t(2171),f=t(1171),m=t(982),h=t(7376),b=t(3032),g=t(287),v=t(4047),y=t(4101),w=t(2132),k=function(n){(0,u.Z)(t,n);var e=(0,l.Z)(t);function t(){return(0,c.Z)(this,t),e.apply(this,arguments)}return(0,d.Z)(t)}(p.B),F=k.compose({baseName:"progress",template:function(n,e){return(0,m.d)(i||(i=(0,f.Z)(['\n    <template\n        role="progressbar"\n        aria-valuenow="','"\n        aria-valuemin="','"\n        aria-valuemax="','"\n        class="','"\n    >\n        ',"\n    </template>\n"])),(function(n){return n.value}),(function(n){return n.min}),(function(n){return n.max}),(function(n){return n.paused?"paused":""}),(0,h.g)((function(n){return"number"===typeof n.value}),(0,m.d)(o||(o=(0,f.Z)(['\n                <div class="progress" part="progress" slot="determinate">\n                    <div\n                        class="determinate"\n                        part="determinate"\n                        style="width: ','%"\n                    ></div>\n                </div>\n            '])),(function(n){return n.percentComplete})),(0,m.d)(a||(a=(0,f.Z)(['\n                <div class="progress" part="progress" slot="indeterminate">\n                    <slot class="indeterminate" name="indeterminate">\n                        ',"\n                        ","\n                    </slot>\n                </div>\n            "])),e.indeterminateIndicator1||"",e.indeterminateIndicator2||"")))},styles:function(n,e){return(0,b.i)(r||(r=(0,f.Z)(["\n    "," :host {\n      align-items: center;\n      height: calc(("," * 3) * 1px);\n    }\n\n    .progress {\n      background-color: ",";\n      border-radius: calc("," * 1px);\n      width: 100%;\n      height: calc("," * 1px);\n      display: flex;\n      align-items: center;\n      position: relative;\n    }\n\n    .determinate {\n      background-color: ",";\n      border-radius: calc("," * 1px);\n      height: calc(("," * 3) * 1px);\n      transition: all 0.2s ease-in-out;\n      display: flex;\n    }\n\n    .indeterminate {\n      height: calc(("," * 3) * 1px);\n      border-radius: calc("," * 1px);\n      display: flex;\n      width: 100%;\n      position: relative;\n      overflow: hidden;\n    }\n\n    .indeterminate-indicator-1 {\n      position: absolute;\n      opacity: 0;\n      height: 100%;\n      background-color: ",";\n      border-radius: calc("," * 1px);\n      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);\n      width: 40%;\n      animation: indeterminate-1 2s infinite;\n    }\n\n    .indeterminate-indicator-2 {\n      position: absolute;\n      opacity: 0;\n      height: 100%;\n      background-color: ",";\n      border-radius: calc("," * 1px);\n      animation-timing-function: cubic-bezier(0.4, 0, 0.6, 1);\n      width: 60%;\n      animation: indeterminate-2 2s infinite;\n    }\n\n    :host(.paused) .indeterminate-indicator-1,\n    :host(.paused) .indeterminate-indicator-2 {\n      animation: none;\n      background-color: ",";\n      width: 100%;\n      opacity: 1;\n    }\n\n    :host(.paused) .determinate {\n      background-color: ",";\n    }\n\n    @keyframes indeterminate-1 {\n      0% {\n        opacity: 1;\n        transform: translateX(-100%);\n      }\n      70% {\n        opacity: 1;\n        transform: translateX(300%);\n      }\n      70.01% {\n        opacity: 0;\n      }\n      100% {\n        opacity: 0;\n        transform: translateX(300%);\n      }\n    }\n\n    @keyframes indeterminate-2 {\n      0% {\n        opacity: 0;\n        transform: translateX(-150%);\n      }\n      29.99% {\n        opacity: 0;\n      }\n      30% {\n        opacity: 1;\n        transform: translateX(-150%);\n      }\n      100% {\n        transform: translateX(166.66%);\n        opacity: 1;\n      }\n    }\n  "])),(0,v.j)("flex"),w.Han,w.rU8,w._5n,w.Han,w.Avx,w._5n,w.Han,w.Han,w._5n,w.Avx,w._5n,w.Avx,w._5n,w.Q5n,w.Q5n).withBehaviors((0,y.vF)((0,b.i)(s||(s=(0,f.Z)(["\n        .indeterminate-indicator-1,\n        .indeterminate-indicator-2,\n        .determinate,\n        .progress {\n          background-color: ",";\n        }\n        :host(.paused) .indeterminate-indicator-1,\n        :host(.paused) .indeterminate-indicator-2,\n        :host(.paused) .determinate {\n          background-color: ",";\n        }\n      "])),g.H.ButtonText,g.H.GrayText)))},indeterminateIndicator1:'\n    <span class="indeterminate-indicator-1" part="indeterminate-indicator-1"></span>\n  ',indeterminateIndicator2:'\n    <span class="indeterminate-indicator-2" part="indeterminate-indicator-2"></span>\n  '})}}]);
//# sourceMappingURL=622.c4ce0550.chunk.js.map