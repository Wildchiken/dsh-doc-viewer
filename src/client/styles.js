/**
 * Page styles. Every surface, label, radius and shadow is one of the Harness's
 * own tokens, so the page follows its light and dark themes and any installed
 * skin. The language is deliberately quiet: hairlines rather than boxes, one
 * filled control at a time, and materials instead of opaque popovers.
 */
const FONT = '-apple-system,BlinkMacSystemFont,"SF Pro Text","PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif';

export const CSS = `
.dvd-root,.dvd-root *{box-sizing:border-box}
.dvd-root{--dvd-hairline:var(--dsw-alias-border-l2);--dvd-hover:var(--dsw-alias-interactive-bg-hover);--dvd-select:var(--dsw-specific-sidebar-nav-item-active-accent);--dvd-accent:#4d6bfe;--dvd-canvas:var(--dsw-specific-sidebar-fill);--dvd-material:color-mix(in srgb,var(--dsw-alias-bg-layer-1) 84%,transparent);--dvd-ease:cubic-bezier(.2,.8,.2,1);--dvd-ease-out:cubic-bezier(.4,0,1,1);--dvd-fast:.14s;--dvd-mid:.2s;position:relative;display:flex;flex-direction:column;flex:1;min-width:0;min-height:0;background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-family:${FONT};font-size:13px;-webkit-font-smoothing:antialiased}
.dvd-root button{font:inherit;color:inherit}
.dvd-root{container-type:inline-size}

.dvd-icon{flex:none;display:block}

/* Top bar: documents on the left, what you can do with this one on the right. */
.dvd-bar{display:flex;align-items:center;gap:8px;height:52px;padding:0 14px 0 16px;flex:none}
.dvd-tabs{display:flex;align-items:center;gap:2px;flex:0 1 auto;min-width:0;overflow-x:auto;scrollbar-width:none}
.dvd-tabs::-webkit-scrollbar{display:none}
.dvd-tab{position:relative;display:flex;align-items:center;gap:6px;flex:0 1 auto;min-width:92px;max-width:200px;height:30px;padding:0 6px 0 10px;border-radius:var(--dsw-radius-sm);color:var(--dsw-alias-label-secondary);cursor:default;user-select:none;transition:background-color var(--dvd-fast) ease,color var(--dvd-fast) ease;animation:dvd-tab-in var(--dvd-mid) var(--dvd-ease)}
.dvd-tab.closing{pointer-events:none;animation:dvd-tab-out var(--dvd-fast) var(--dvd-ease-out) forwards}
@keyframes dvd-tab-in{from{opacity:0;transform:translateY(3px) scale(.97)}to{opacity:1;transform:none}}
@keyframes dvd-tab-out{to{opacity:0;transform:scale(.94);max-width:0;min-width:0;padding:0;margin:0}}
.dvd-tab:hover{background:var(--dvd-hover)}
.dvd-tab[aria-selected="true"]{background:var(--dvd-hover);color:var(--dsw-alias-label-primary);font-weight:500}
.dvd-tab:focus-visible,.dvd-btn:focus-visible,.dvd-iconbtn:focus-visible,.dvd-row-btn:focus-visible,.dvd-seg button:focus-visible{outline:2px solid var(--dvd-accent);outline-offset:1px}
.dvd-tab-name{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dvd-tab-close{position:relative;display:inline-flex;align-items:center;justify-content:center;flex:none;width:18px;height:18px;padding:0;border:0;border-radius:var(--dsw-radius-xs);background:transparent;color:var(--dsw-alias-label-tertiary);opacity:0;cursor:pointer;transition:opacity var(--dvd-fast) ease}
.dvd-tab:hover .dvd-tab-close,.dvd-tab[aria-selected="true"] .dvd-tab-close,.dvd-tab-close:focus-visible{opacity:1}
.dvd-tab-close:hover{background:var(--dvd-hover);color:var(--dsw-alias-label-primary)}
.dvd-dirty{position:absolute;width:6px;height:6px;border-radius:50%;background:var(--dsw-alias-label-secondary)}
.dvd-tab:hover .dvd-dirty{display:none}
.dvd-actions{display:flex;align-items:center;gap:2px;flex:none;margin-left:auto}
.dvd-actions-slot{display:contents}
.dvd-ask{gap:5px;margin-right:4px;padding:0 10px}
.dvd-toast button+button{margin-left:2px}
.dvd-toast{white-space:nowrap}
.dvd-progress{position:absolute;left:0;right:0;top:51px;height:1.5px;overflow:hidden;z-index:5;pointer-events:none}
.dvd-progress::after{content:"";position:absolute;inset:0 auto 0 0;width:32%;background:var(--dvd-accent);border-radius:1px;animation:dvd-slide 1.15s cubic-bezier(.4,0,.2,1) infinite}
@keyframes dvd-slide{from{transform:translateX(-110%)}to{transform:translateX(320%)}}

/* Controls */
.dvd-btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;height:30px;padding:0 12px;border:0;border-radius:var(--dsw-radius-sm);background:var(--dvd-hover);font-weight:500;white-space:nowrap;cursor:pointer;transition:background-color var(--dvd-fast) ease,opacity var(--dvd-fast) ease}
.dvd-btn:hover:not(:disabled){background:var(--dsw-alias-border-l2)}
.dvd-btn.primary{background:var(--dsw-alias-button-primary-fill);color:var(--dsw-alias-label-primary-foreground)}
.dvd-btn.primary:hover:not(:disabled){background:var(--dsw-alias-button-primary-hover)}
.dvd-btn.plain{background:transparent;color:var(--dsw-alias-label-secondary);font-weight:400}
.dvd-btn.plain:hover:not(:disabled){background:var(--dvd-hover);color:var(--dsw-alias-label-primary)}
.dvd-btn:disabled,.dvd-iconbtn:disabled{opacity:.38;cursor:default}
.dvd-iconbtn{display:inline-flex;align-items:center;justify-content:center;flex:none;width:30px;height:30px;padding:0;border:0;border-radius:var(--dsw-radius-sm);background:transparent;color:var(--dsw-alias-label-secondary);cursor:pointer;transition:background-color var(--dvd-fast) ease,color var(--dvd-fast) ease}
.dvd-iconbtn:hover:not(:disabled){background:var(--dvd-hover);color:var(--dsw-alias-label-primary)}
.dvd-iconbtn[aria-pressed="true"]{background:var(--dvd-hover);color:var(--dsw-alias-label-primary)}
.dvd-iconbtn.small{width:24px;height:24px}
.dvd-input{width:100%;height:30px;padding:0 10px;border:0;border-radius:var(--dsw-radius-sm);background:var(--dvd-hover);color:var(--dsw-alias-label-primary);font:inherit;outline:none;transition:box-shadow var(--dvd-fast) ease}
.dvd-input::placeholder{color:var(--dsw-alias-label-tertiary)}
.dvd-input:focus{box-shadow:0 0 0 2px color-mix(in srgb,var(--dvd-accent) 55%,transparent)}
.dvd-seg{display:inline-flex;gap:2px;padding:2px;border-radius:var(--dsw-radius-sm);background:var(--dvd-hover);align-self:flex-start;max-width:100%;overflow-x:auto}
.dvd-seg button{height:26px;padding:0 12px;border:0;border-radius:6px;background:transparent;color:var(--dsw-alias-label-secondary);white-space:nowrap;cursor:pointer;transition:background-color var(--dvd-fast) ease,color var(--dvd-fast) ease}
.dvd-seg button[aria-pressed="true"]{background:var(--dsw-alias-bg-base);color:var(--dsw-alias-label-primary);font-weight:500;box-shadow:0 0 0 .5px var(--dvd-hairline),0 1px 2px rgb(0 0 0 / .08)}
.dvd-seg button:disabled{opacity:.35;cursor:default}
.dvd-select{height:30px;padding:0 8px;border:0;border-radius:var(--dsw-radius-sm);background:var(--dvd-hover);color:var(--dsw-alias-label-primary);font:inherit;outline:none}
.dvd-label{font-size:12px;color:var(--dsw-alias-label-secondary)}
.dvd-field{display:flex;flex-direction:column;gap:6px;min-width:0}
.dvd-fieldrow{display:flex;gap:12px;flex-wrap:wrap}
.dvd-fieldrow .dvd-field{flex:1 1 120px}
.dvd-check{display:inline-flex;align-items:center;gap:7px;color:var(--dsw-alias-label-secondary);cursor:pointer}
.dvd-check input{margin:0;accent-color:var(--dvd-accent)}

/* Menus use the Host's material */
.dvd-menuwrap{position:relative;display:inline-flex}
.dvd-hint{position:absolute;right:-4px;top:calc(100% + 8px);z-index:45;display:flex;align-items:center;gap:10px;padding:7px 8px 7px 12px;border-radius:var(--dsw-radius-md);background:var(--dvd-accent);color:#fff;font-size:12.5px;white-space:nowrap;box-shadow:0 8px 24px rgb(77 107 254 / .28);transform-origin:top right;animation:dvd-pop var(--dvd-mid) var(--dvd-ease)}
.dvd-hint::before{content:"";position:absolute;right:14px;top:-4px;width:8px;height:8px;border-radius:1px;background:inherit;transform:rotate(45deg)}
.dvd-hint button{height:22px;padding:0 8px;border:0;border-radius:6px;background:rgb(255 255 255 / .2);color:#fff;font:inherit;font-weight:500;cursor:pointer}
.dvd-hint button:hover{background:rgb(255 255 255 / .3)}
.dvd-menu{position:absolute;right:0;top:calc(100% + 6px);z-index:40;min-width:190px;padding:5px;border-radius:var(--dsw-radius-md);background:var(--dvd-material);-webkit-backdrop-filter:blur(24px) saturate(180%);backdrop-filter:blur(24px) saturate(180%);box-shadow:var(--dsw-elevation-prominent,0 12px 36px rgb(0 0 0 / .18));transform-origin:top right;animation:dvd-pop var(--dvd-fast) var(--dvd-ease)}
@keyframes dvd-pop{from{opacity:0;transform:scale(.97) translateY(-2px)}to{opacity:1;transform:none}}
.dvd-menu button{display:flex;align-items:center;gap:9px;width:100%;height:30px;padding:0 9px;border:0;border-radius:var(--dsw-radius-sm);background:transparent;text-align:left;white-space:nowrap;cursor:pointer}
.dvd-menu button:hover:not(:disabled){background:var(--dvd-hover)}
.dvd-menu button:disabled{opacity:.38;cursor:default}
.dvd-menu button .dvd-icon{color:var(--dsw-alias-label-secondary)}
.dvd-changes-chip{gap:5px;color:var(--dvd-accent)!important;font-weight:500!important;animation:dvd-pop var(--dvd-mid) var(--dvd-ease)}
.dvd-count{display:inline-block;font-variant-numeric:tabular-nums;animation:dvd-bump .32s var(--dvd-ease)}
@keyframes dvd-bump{0%{transform:scale(1)}40%{transform:scale(1.22)}100%{transform:scale(1)}}
.dvd-changes{width:320px;max-height:380px;overflow:auto;white-space:normal}
.dvd-change{padding:6px 8px 8px}
.dvd-change+.dvd-change{border-top:.5px solid var(--dvd-hairline)}
.dvd-change-head{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--dsw-alias-label-secondary)}
.dvd-change-tag{flex:1;min-width:0}
.dvd-menu .dvd-change-undo{width:auto;height:24px;padding:0 8px;color:var(--dsw-alias-label-primary)}
.dvd-change-cells{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
.dvd-menu .dvd-change-cell{width:auto;height:22px;padding:0 7px;border-radius:5px;background:var(--dvd-hover);font-family:ui-monospace,"SF Mono",Menlo,monospace;font-size:11.5px}
.dvd-change-more,.dvd-change-note{font-size:12px;color:var(--dsw-alias-label-tertiary);align-self:center}
.dvd-change-note{margin-top:4px}
.dvd-menu.left{left:0;right:auto;transform-origin:top left}
.dvd-menu .dvd-sep{height:1px;margin:5px 9px;padding:0;border:0;background:color-mix(in srgb,var(--dsw-alias-label-primary) 12%,transparent)}

/* Panes: all documents stay alive; only the selected one is visible. */
.dvd-body{position:relative;flex:1;min-height:0}
.dvd-pane{position:absolute;inset:0;display:flex;flex-direction:column}
.dvd-pane:not(.active){visibility:hidden;pointer-events:none}
.dvd-pane.active{animation:dvd-fade var(--dvd-mid) ease}
.dvd-notice{display:flex;align-items:flex-start;gap:8px;padding:7px 16px;flex:none;border-top:.5px solid var(--dvd-hairline);color:var(--dsw-alias-label-secondary);font-size:12px;line-height:1.5}
.dvd-notice .dvd-icon{margin-top:2px;color:var(--dsw-alias-state-warn-primary)}
.dvd-notice span{flex:1}

/* Empty state */
.dvd-empty{display:flex;flex-direction:column;align-items:center;flex:1;min-height:0;overflow:auto;padding:12vh 24px 40px}
.dvd-empty-inner{display:flex;flex-direction:column;width:min(520px,100%)}
.dvd-empty h1{margin:0;font-size:28px;font-weight:600;letter-spacing:-.02em;line-height:1.2;text-align:center}
.dvd-empty p{margin:10px 0 0;font-size:14px;line-height:1.6;color:var(--dsw-alias-label-secondary);text-align:center}
.dvd-empty-actions{display:flex;justify-content:center;gap:8px;margin-top:24px}
.dvd-empty-actions .dvd-btn{height:34px;padding:0 16px}
.dvd-empty-inner>*{animation:dvd-rise var(--dvd-mid) var(--dvd-ease) backwards}
.dvd-empty-inner>:nth-child(2){animation-delay:30ms}.dvd-empty-inner>:nth-child(3){animation-delay:60ms}.dvd-empty-inner>:nth-child(n+4){animation-delay:90ms}
.dvd-empty-tips{display:flex;flex-direction:column;gap:2px;margin:28px auto 0;padding:0;list-style:none;width:min(400px,100%)}
.dvd-empty-tips li{display:flex;align-items:flex-start;gap:10px;padding:6px 10px;color:var(--dsw-alias-label-secondary);font-size:12.5px;line-height:1.5}
.dvd-empty-tips .dvd-icon{flex:none;margin-top:2px;color:var(--dsw-alias-label-tertiary)}
.dvd-empty-tips b{margin-right:8px;font-weight:500;color:var(--dsw-alias-label-primary)}
.dvd-section-title{margin:44px 0 6px;padding:0 10px;font-size:12px;font-weight:500;color:var(--dsw-alias-label-tertiary)}
.dvd-row-btn{display:flex;align-items:center;gap:10px;width:100%;min-height:36px;padding:6px 10px;border:0;border-radius:var(--dsw-radius-sm);background:transparent;text-align:left;cursor:pointer;transition:background-color var(--dvd-fast) ease}
.dvd-row-btn:hover:not(:disabled){background:var(--dvd-hover)}
.dvd-row-btn:disabled{opacity:.45;cursor:default}
.dvd-row-btn.selected{background:var(--dvd-select)}
.dvd-row-btn .dvd-icon{color:var(--dsw-alias-label-secondary)}
.dvd-row-name{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dvd-row-meta{flex:none;font-size:12px;color:var(--dsw-alias-label-tertiary)}
.dvd-drop{position:absolute;inset:0;z-index:60;display:flex;align-items:center;justify-content:center;background:color-mix(in srgb,var(--dsw-alias-bg-base) 72%,transparent);-webkit-backdrop-filter:blur(12px);backdrop-filter:blur(12px);font-size:17px;font-weight:500;letter-spacing:-.01em;pointer-events:none;animation:dvd-fade var(--dvd-fast) ease}
@keyframes dvd-fade{from{opacity:0}to{opacity:1}}

/* Loading and errors */
.dvd-center{display:flex;flex:1;flex-direction:column;align-items:center;justify-content:center;gap:12px;padding:32px;text-align:center;color:var(--dsw-alias-label-secondary)}
.dvd-center h3{margin:0;font-size:15px;font-weight:600;color:var(--dsw-alias-label-primary)}
.dvd-center p{margin:0;max-width:440px;line-height:1.6}
.dvd-spinner{width:20px;height:20px;border:2px solid var(--dvd-hairline);border-top-color:var(--dsw-alias-label-secondary);border-radius:50%;animation:dvd-spin .8s linear infinite}
@keyframes dvd-spin{to{transform:rotate(360deg)}}
.dvd-skeleton{flex:1;min-height:0;overflow:hidden;animation:dvd-fade var(--dvd-mid) ease .12s backwards}
.dvd-skeleton span{display:block;border-radius:4px;background:linear-gradient(90deg,var(--dvd-hover) 0%,color-mix(in srgb,var(--dvd-hover) 40%,transparent) 50%,var(--dvd-hover) 100%);background-size:200% 100%;animation:dvd-shimmer 1.4s ease-in-out infinite}
@keyframes dvd-shimmer{from{background-position:100% 0}to{background-position:-100% 0}}
.dvd-skeleton.grid{display:flex;flex-direction:column}
.dvd-skel-row{display:grid;grid-template-columns:44px repeat(6,minmax(0,1fr));border-bottom:.5px solid var(--dvd-hairline)}
.dvd-skel-row span{height:10px;margin:8px 10px;border-radius:3px}
.dvd-skel-row:first-child span,.dvd-skel-row span.head{opacity:.6}
.dvd-skeleton.page,.dvd-skeleton.text{display:flex;justify-content:center;padding:28px 24px;background:var(--dvd-canvas)}
.dvd-skeleton.text{background:transparent}
.dvd-skel-sheet{display:flex;flex-direction:column;gap:11px;width:min(620px,100%);height:max-content;padding:44px 48px;border-radius:4px;background:var(--dsw-alias-bg-layer-1);box-shadow:0 1px 3px rgb(0 0 0 / .08)}
.dvd-skeleton.text .dvd-skel-sheet{padding:8px 0;background:transparent;box-shadow:none}
.dvd-skel-sheet span{height:9px}
.dvd-skel-sheet span.title{width:46%;height:16px;margin-bottom:10px}

/* Spreadsheet: FortuneSheet, dressed to sit quietly in the page */
.dvd-sheet{position:relative;flex:1;min-height:0;display:flex;flex-direction:column;border-top:.5px solid var(--dvd-hairline)}
.dvd-sheet .fortune-presence-selection{border-width:2px!important;border-radius:2px;background:color-mix(in srgb,var(--dvd-accent) 14%,transparent);opacity:1;animation:dvd-mark .9s cubic-bezier(.2,.8,.2,1);transition:opacity 1.2s ease}
.dvd-sheet .fortune-presence-username{border-radius:4px 4px 4px 0;font-size:11px;font-weight:500;animation:dvd-fade .2s ease}
.dvd-sheet .fortune-presence-username:empty{display:none}
.dvd-sheet[data-marks-fading] .fortune-presence-selection{opacity:0}
@keyframes dvd-mark{0%{box-shadow:0 0 0 0 color-mix(in srgb,var(--dvd-accent) 55%,transparent)}100%{box-shadow:0 0 0 8px transparent}}
.dvd-sheet .fortune-container{flex:1;min-height:0;font-family:${FONT}}
.dvd-sheet .fortune-toolbar{padding:4px 12px;border:0;background:transparent;box-shadow:none}
.dvd-sheet .fortune-toolbar-divider{opacity:.5}
.dvd-sheet .fortune-fx-editor{border-top:.5px solid var(--dvd-hairline);border-bottom:.5px solid var(--dvd-hairline)}
.dvd-sheet .luckysheet-sheet-area{border-top:.5px solid var(--dvd-hairline);background:transparent}
:root[style*="color-scheme: dark"] .dvd-sheet .fortune-container{filter:invert(.91) hue-rotate(180deg)}

/* PDF */
.dvd-pdf{position:relative;display:flex;flex:1;min-height:0;border-top:.5px solid var(--dvd-hairline);background:var(--dvd-canvas)}
.dvd-pdf-thumbs{width:136px;flex:none;display:flex;flex-direction:column;gap:6px;padding:12px 16px;overflow:auto;border-right:.5px solid var(--dvd-hairline);scrollbar-width:thin}
.dvd-thumb{display:flex;flex-direction:column;align-items:center;gap:5px;padding:0;border:0;background:transparent;color:var(--dsw-alias-label-tertiary);font-size:11px;cursor:pointer}
.dvd-thumb canvas{display:block;width:100%;height:auto;min-height:60px;background:#fff;border-radius:3px;box-shadow:0 0 0 .5px var(--dvd-hairline),0 1px 3px rgb(0 0 0 / .1)}
.dvd-thumb.active canvas{box-shadow:0 0 0 2px var(--dvd-accent)}
.dvd-thumb.active{color:var(--dsw-alias-label-primary);font-weight:500}
.dvd-pdf-main{position:relative;display:flex;flex:1;min-width:0;flex-direction:column}
.dvd-pdf-scroll{flex:1;min-height:0;overflow:auto;padding:20px 16px 76px}
.dvd-pdf-pages{display:flex;flex-direction:column;align-items:center;gap:14px;width:fit-content;min-width:100%}
.dvd-pdf-page{position:relative;flex:none;background:#fff;border-radius:2px;box-shadow:0 0 0 .5px rgb(0 0 0 / .12),0 2px 10px rgb(0 0 0 / .08)}
.dvd-pdf-page canvas{display:block;border-radius:2px}
.dvd-hud{position:absolute;left:50%;bottom:18px;z-index:10;display:flex;align-items:center;gap:1px;height:38px;padding:0 5px;border-radius:999px;background:var(--dvd-material);-webkit-backdrop-filter:blur(24px) saturate(180%);backdrop-filter:blur(24px) saturate(180%);box-shadow:var(--dsw-elevation-prominent,0 8px 28px rgb(0 0 0 / .16));transform:translateX(-50%);white-space:nowrap}
.dvd-hud .dvd-iconbtn{border-radius:50%;width:28px;height:28px}
.dvd-hud-sep{width:.5px;height:16px;margin:0 5px;background:var(--dvd-hairline)}
.dvd-pageinput{width:30px;height:24px;padding:0;border:0;border-radius:6px;background:transparent;color:var(--dsw-alias-label-primary);font:inherit;font-variant-numeric:tabular-nums;text-align:right;outline:none}
.dvd-pageinput:focus{background:var(--dvd-hover);text-align:center}
.dvd-hud-text{min-width:40px;padding:0 3px;font-variant-numeric:tabular-nums;text-align:center;color:var(--dsw-alias-label-secondary)}
.dvd-root .textLayer{position:absolute;inset:0;overflow:clip;opacity:1;line-height:1;text-align:initial;text-size-adjust:none;forced-color-adjust:none;transform-origin:0 0;caret-color:CanvasText;z-index:0;--min-font-size:1;--text-scale-factor:calc(var(--total-scale-factor)*var(--min-font-size));--min-font-size-inv:calc(1/var(--min-font-size))}
.dvd-root .textLayer :is(span,br){position:absolute;color:transparent;white-space:pre;cursor:text;transform-origin:0% 0%}
.dvd-root .textLayer>:not(.markedContent),.dvd-root .textLayer .markedContent span:not(.markedContent){z-index:1;--font-height:0;font-size:calc(var(--text-scale-factor)*var(--font-height));--scale-x:1;--rotate:0deg;transform:rotate(var(--rotate)) scaleX(var(--scale-x)) scale(var(--min-font-size-inv))}
.dvd-root .textLayer .markedContent{display:contents}
.dvd-root .textLayer ::selection{background:color-mix(in srgb,var(--dvd-accent) 30%,transparent)}
.dvd-root .textLayer mark{position:static;padding:0;border-radius:2px;background:rgb(255 200 0 / .5);color:transparent}

/* Word and text */
.dvd-paper{display:flex;flex:1;min-height:0;padding:16px 16px 0;border-top:.5px solid var(--dvd-hairline);background:var(--dvd-canvas)}
.dvd-paper iframe{display:block;width:min(860px,100%);height:100%;margin:0 auto;border:0;border-radius:4px 4px 0 0;background:#fff;box-shadow:0 0 0 .5px rgb(0 0 0 / .12),0 2px 10px rgb(0 0 0 / .08)}
.dvd-text{flex:1;min-height:0;overflow:auto;border-top:.5px solid var(--dvd-hairline)}
.dvd-text pre{width:min(860px,100%);margin:0 auto;padding:24px 28px 60px;font-family:ui-monospace,"SF Mono",SFMono-Regular,Menlo,monospace;font-size:12.5px;line-height:1.7;white-space:pre-wrap;word-break:break-word}
.dvd-text.editing{display:flex;overflow:hidden}
.dvd-text.editing textarea{flex:1;min-width:0;margin:0;padding:24px max(28px,calc((100% - 804px) / 2)) 60px;border:0;border-radius:0;outline:0;box-shadow:none;resize:none;background:transparent;color:inherit;font-family:ui-monospace,"SF Mono",SFMono-Regular,Menlo,monospace;font-size:12.5px;line-height:1.7;white-space:pre-wrap;overflow-wrap:anywhere}
.dvd-text mark{border-radius:2px;background:rgb(255 200 0 / .4);color:inherit}
.dvd-text mark.current{background:rgb(255 170 0 / .9)}

/* Find bar */
.dvd-find{position:absolute;right:14px;top:10px;z-index:30;display:flex;flex-direction:column;gap:6px;width:min(340px,calc(100% - 28px));padding:7px;border-radius:var(--dsw-radius-md);background:var(--dvd-material);-webkit-backdrop-filter:blur(24px) saturate(180%);backdrop-filter:blur(24px) saturate(180%);box-shadow:var(--dsw-elevation-prominent,0 8px 28px rgb(0 0 0 / .16));animation:dvd-pop var(--dvd-fast) var(--dvd-ease)}
.dvd-find-row{display:flex;align-items:center;gap:2px}
.dvd-find .dvd-input{flex:1;min-width:0}
.dvd-find-count{flex:none;min-width:42px;padding:0 6px;font-size:12px;font-variant-numeric:tabular-nums;text-align:center;color:var(--dsw-alias-label-tertiary)}

/* Sheets: modal cards inside the page, not over the whole app */
.dvd-scrim{position:absolute;inset:0;z-index:50;display:flex;align-items:center;justify-content:center;padding:24px;background:color-mix(in srgb,var(--dsw-alias-bg-base) 45%,rgb(0 0 0 / .18));-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px);animation:dvd-fade var(--dvd-fast) ease}
.dvd-card{display:flex;flex-direction:column;width:min(480px,100%);max-height:100%;overflow:hidden;border-radius:var(--dsw-radius-lg);background:var(--dsw-alias-bg-layer-1);box-shadow:var(--dsw-elevation-prominent,0 24px 64px rgb(0 0 0 / .24));animation:dvd-rise var(--dvd-mid) var(--dvd-ease)}
.dvd-card.wide{width:min(680px,100%)}
@keyframes dvd-rise{from{opacity:0;transform:scale(.98) translateY(6px)}to{opacity:1;transform:none}}
.dvd-card-head{display:flex;align-items:center;gap:8px;padding:18px 16px 4px 22px}
.dvd-card-head h3{flex:1;min-width:0;margin:0;font-size:16px;font-weight:600;letter-spacing:-.01em}
.dvd-card-body{display:flex;flex-direction:column;gap:14px;padding:12px 22px 18px;overflow:auto}
.dvd-card-body p{margin:0;line-height:1.6;color:var(--dsw-alias-label-secondary)}
.dvd-card-foot{display:flex;justify-content:flex-end;gap:8px;padding:0 22px 20px}
.dvd-list{display:flex;flex-direction:column;min-height:48px;max-height:280px;overflow:auto;margin:0 -10px;scrollbar-width:thin}
.dvd-empty-note{padding:18px 10px;text-align:center;color:var(--dsw-alias-label-tertiary)}
.dvd-crumb{display:flex;align-items:center;gap:4px;min-width:0}
.dvd-crumb-path{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--dsw-alias-label-secondary)}
.dvd-callout{padding:9px 12px;border-radius:var(--dsw-radius-sm);background:var(--dvd-hover);font-size:12px;line-height:1.6;color:var(--dsw-alias-label-secondary)}
.dvd-result{display:flex;align-items:center;gap:8px;min-height:32px}
.dvd-result .dvd-row-name{color:var(--dsw-alias-label-secondary)}
.dvd-ok{color:var(--dsw-alias-state-success-primary)}
.dvd-bad{color:var(--dsw-alias-state-error-primary)}

/* Statistics */
.dvd-tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;overflow:hidden;border-radius:var(--dsw-radius-md);background:var(--dvd-hairline)}
.dvd-tile{padding:11px 13px;background:var(--dsw-alias-bg-layer-1)}
.dvd-tile span{display:block;font-size:12px;color:var(--dsw-alias-label-secondary)}
.dvd-tile strong{display:block;margin-top:3px;overflow:hidden;font-size:18px;font-weight:600;letter-spacing:-.01em;font-variant-numeric:tabular-nums;text-overflow:ellipsis;white-space:nowrap}
.dvd-chart{overflow:hidden;border-radius:var(--dsw-radius-md);box-shadow:0 0 0 .5px var(--dvd-hairline)}
.dvd-chart svg{display:block;width:100%;height:auto}

.dvd-toast{position:absolute;left:50%;bottom:24px;z-index:70;display:flex;align-items:center;gap:10px;max-width:min(620px,calc(100% - 40px));min-height:36px;padding:7px 15px;border-radius:999px;background:var(--dvd-material);-webkit-backdrop-filter:blur(24px) saturate(180%);backdrop-filter:blur(24px) saturate(180%);box-shadow:var(--dsw-elevation-prominent,0 8px 28px rgb(0 0 0 / .16));line-height:1.45;transform:translateX(-50%);animation:dvd-toast var(--dvd-mid) var(--dvd-ease)}
@keyframes dvd-toast{from{opacity:0;transform:translate(-50%,8px)}to{opacity:1;transform:translate(-50%,0)}}
.dvd-toast.leaving{pointer-events:none;animation:dvd-toast-out var(--dvd-fast) var(--dvd-ease-out) forwards}
@keyframes dvd-toast-out{to{opacity:0;transform:translate(-50%,6px) scale(.98)}}
.dvd-toast span{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dvd-toast.error{color:var(--dsw-alias-state-error-primary)}
.dvd-toast button{flex:none;padding:0;border:0;background:none;color:var(--dvd-accent);font-weight:500;cursor:pointer}
@media (prefers-reduced-motion:reduce){.dvd-root *,.dvd-root *::after{animation-duration:.01ms!important;transition-duration:.01ms!important}.dvd-skeleton span{animation:none!important}}
/* As a file preview the Harness frames one document: the others stay in their tabs elsewhere. */
.dvd-host[data-placement^="preview"] .dvd-tab{display:none}
/* Narrow, as beside the conversation: labels give way to icons, and the document in front keeps its name. */
@container (max-width:720px){
.dvd-bar{gap:4px;height:44px;padding:0 8px 0 10px}
.dvd-progress{top:43px}
.dvd-wide{display:none}
.dvd-ask{margin-right:0;padding:0 7px}
.dvd-changes-chip{padding:0 8px}
.dvd-tab{min-width:34px;padding:0 4px 0 9px}
.dvd-tab[aria-selected="true"]{flex:0 1 auto;min-width:96px;max-width:62%}
.dvd-tab:not([aria-selected="true"]) .dvd-tab-name,.dvd-tab:not([aria-selected="true"]) .dvd-tab-close{display:none}
.dvd-tab:not([aria-selected="true"]){padding:0 9px}
.dvd-seg button{padding:0 8px}
.dvd-toast{white-space:normal;max-width:calc(100% - 24px)}
.dvd-tab-ext{display:none}
}
`;
