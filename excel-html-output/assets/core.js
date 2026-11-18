/* Excel HTML Output — original implementation, no runtime dependencies. Formulas are read and computed by formulas.js. */
(function(root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./toon.js'), require('./formulas.js'));
  else root.ExcelHTMLCore = factory(root.ExcelHTMLToon, root.ExcelHTMLFormulas);
})(typeof globalThis !== 'undefined' ? globalThis : this, function(TOON, F) {
  'use strict';
  const SCHEMA = 'excel-html-workbook/v1';
  // Column widths and row heights as [smallest, largest], in CSS pixels. The largest are Excel's: the widest column
  // and the highest row that Excel can set.
  const SIZE_LIMITS = {colWidth: [24, 2044], rowHeight: [16, 546]};
  const FORMATS = ['General','@','0','0.00','#,##0','#,##0.00','0%','0.00%','¥#,##0','yyyy/mm/dd'];
  // Excel's own clipboard spelling of each format (measured on Excel 2021 copy output).
  const MSO_FORMAT = {'General':'General','@':'"\\@"','0':'0','0.00':'Fixed','#,##0':'"\\#\\,\\#\\#0"','#,##0.00':'Standard',
    '0%':'0%','0.00%':'Percent','¥#,##0':'"\\0022¥\\0022\\#\\,\\#\\#0"','yyyy/mm/dd':'"yyyy\\/mm\\/dd"'};
  // Japanese Excel defaults: no fill, vertically centred (the Normal style of ja-JP workbooks).
  const DEFAULT = {font:'Yu Gothic',size:11,bold:false,italic:false,underline:false,strike:false,
    color:'#000000',fill:'none',align:'general',valign:'middle',wrap:false,numFmt:'General',
    borderTop:'none',borderRight:'none',borderBottom:'none',borderLeft:'none',indent:0};
  // Border kinds that Excel can represent, with the clipboard CSS Excel itself writes for them.
  const BORDER_XL = {'0.5px solid':'.5pt hairline','1px solid':'.5pt solid','2px solid':'1.0pt solid','3px solid':'1.5pt solid',
    '1px dashed':'.5pt dashed','2px dashed':'1.0pt dashed','1px dotted':'.5pt dotted','3px double':'2.0pt double'};
  // What a book can hold, besides the style keys of DEFAULT and the formats and borders above. validate takes nothing
  // else and the formulas compute only FUNCTIONS, so a new feature has to be added to these lists, and then
  // tests/xlsx-features.test.mjs asks for its samples written to xlsx and read back (SKILL.md, 実装を変えるとき).
  const CHOICES={align:['general','left','center','right'],valign:['top','middle','bottom']};
  const CELL_TYPES=['s','n','b','f'];
  // The functions the formulas compute (formulas.js): a function taken out of this list is #NAME?.
  const FUNCTIONS=F.FUNCTIONS;
  const FIELDS={book:['schema','id','revision','title','activeSheet','defaultStyle','view','assets','sheets'],
    sheet:['id','name','rows','cols','cells','merges','colWidths','rowHeights','freeze','gridlines','images','defaultStyle'],
    cell:['t','v','f','s','link'],view:['zoom','active','selection','scrollTop','scrollLeft'],freeze:['rows','cols'],
    range:['r1','c1','r2','c2'],image:['id','asset','r','c','x','y','w','h']};
  const clone = value => JSON.parse(JSON.stringify(value));
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const colName = n => { let out=''; for(n++;n>0;n=Math.floor((n-1)/26)) out=String.fromCharCode(65+(n-1)%26)+out; return out; };
  const addr = (r,c) => colName(c)+(r+1);
  function coord(a) { const m=/^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(a); if(!m) throw Error('セル番地が不正です: '+a); return [+m[2]-1,[...m[1].toUpperCase()].reduce((n,x)=>n*26+x.charCodeAt(0)-64,0)-1]; }
  const rect = (a,b=a) => ({r1:Math.min(a[0],b[0]),c1:Math.min(a[1],b[1]),r2:Math.max(a[0],b[0]),c2:Math.max(a[1],b[1])});
  function range(s) { const parts=s.toUpperCase().split(':'); if(parts.length>2) throw Error('範囲が不正です'); return rect(coord(parts[0]),coord(parts[1]||parts[0])); }
  const within=(r,c,q)=>r>=q.r1&&r<=q.r2&&c>=q.c1&&c<=q.c2;
  const overlap=(a,b)=>a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;
  const same=(a,b)=>a.r1===b.r1&&a.r2===b.r2&&a.c1===b.c1&&a.c2===b.c2;
  const each=(q,fn)=>{for(let r=q.r1;r<=q.r2;r++)for(let c=q.c1;c<=q.c2;c++)fn(r,c);};
  const get=(s,r,c)=>s.cells[addr(r,c)] || {t:'s',v:''};
  const style=(w,s,cell)=>({...DEFAULT,...w.defaultStyle,...s.defaultStyle,...cell.s});
  function mergeAt(s,r,c){return s.merges.find(m=>within(r,c,m));}
  function expanded(s,q){q={...q};let changed=true;while(changed){changed=false;for(const m of s.merges) if(overlap(q,m)){const n=rect([Math.min(q.r1,m.r1),Math.min(q.c1,m.c1)],[Math.max(q.r2,m.r2),Math.max(q.c2,m.c2)]);if(!same(n,q)){q=n;changed=true;}}}return q;}
  // The area a selection covers. As in Excel, whole rows or whole columns (chosen with the headers) stay as chosen:
  // clicking column B selects B alone even where A11:C12 is merged. Any other range takes in every merged cell it
  // touches. A selection of whole rows or columns can therefore hold only part of a merged cell (see splitsMerge).
  const selectionArea=(s,q)=>q.r1===0&&q.r2===s.rows-1||q.c1===0&&q.c2===s.cols-1?q:expanded(s,q);
  const inside=(m,q)=>within(m.r1,m.c1,q)&&within(m.r2,m.c2,q);
  // Whether q holds only part of some merged cell. Excel refuses to cut or clear such a range ("この操作は結合した
  // セルには行えません。", desktop Excel 16.0), and copies it as separate cells.
  const splitsMerge=(s,q)=>s.merges.some(m=>overlap(m,q)&&!inside(m,q));
  // The merged cell at (r,c), if it lies wholly inside q; a merged cell only partly inside q is copied as separate cells.
  const mergeIn=(s,r,c,q)=>{const m=mergeAt(s,r,c);return m&&inside(m,q)?m:undefined;};
  const PICTURE_DATA=/^data:image\/(?:png|jpeg|gif|webp);base64,[A-Za-z0-9+/]+={0,2}$/;
  // assetsDetached: the editor keeps picture bytes outside the book while editing, so references are not checked then.
  function validate(w,{assetsDetached=false}={}) {
    const fail=m=>{throw Error('ブック形式エラー: '+m);};
    const integer=(n,min,max)=>Number.isInteger(n)&&n>=min&&n<=max;
    // A field the page does not have is refused, as an unknown style key is, rather than kept unseen and lost later.
    const fields=(kind,o,where)=>{for(const k of Object.keys(o))if(!FIELDS[kind].includes(k))fail('未対応の項目 '+where+k);};
    if(!w||w.schema!==SCHEMA||typeof w.id!=='string'||!w.id||typeof w.revision!=='string'||!w.revision) fail('schema / id / revision');
    fields('book',w,'');
    if(typeof w.title!=='string'||!Array.isArray(w.sheets)||!w.sheets.length||w.sheets.length>30) fail('title / sheets');
    if(w.assets!==undefined){
      if(!w.assets||typeof w.assets!=='object'||Array.isArray(w.assets)) fail('assets');
      for(const [k,v] of Object.entries(w.assets)) if(!k||typeof v!=='string'||v.length>15e6||!PICTURE_DATA.test(v)) fail('画像データ '+k+'（png / jpeg / gif / webp の data URL）');
    }
    const ids=new Set(), names=new Set();
    const vs = st => { if(!st) return; if(typeof st!=='object'||Array.isArray(st)) fail('style');
      for(const [k,v] of Object.entries(st)) {
        if(!(k in DEFAULT)) fail('未対応の書式 '+k);
        if(k==='color'&&!/^#[\da-f]{6}$/i.test(v)) fail(k);
        if(k==='fill'&&v!=='none'&&!/^#[\da-f]{6}$/i.test(v)) fail(k);
        if(k==='font'&&(typeof v!=='string'||!v||v.length>100||/[<>;{}"\\]/.test(v))) fail('font');
        if(k==='size'&&!(Number.isFinite(v)&&v>=6&&v<=96)) fail('size');
        if(k==='indent'&&!integer(v,0,15)) fail('indent');
        if(['bold','italic','underline','strike','wrap'].includes(k)&&typeof v!=='boolean') fail(k);
        if(k==='align'&&!CHOICES.align.includes(v)) fail(k);
        if(k==='valign'&&!CHOICES.valign.includes(v)) fail(k);
        if(k==='numFmt'&&!FORMATS.includes(v)) fail('表示形式 '+v);
        if(k.startsWith('border')&&v!=='none'&&!(typeof v==='string'&&/^\S+ \S+ #[\da-f]{6}$/i.test(v)&&BORDER_XL[v.split(' ').slice(0,2).join(' ')])) fail(k+'（Excelで表せる線: '+Object.keys(BORDER_XL).join(' / ')+'）');
      }
    };
    vs(w.defaultStyle);
    for(const [si,s] of w.sheets.entries()) {
      const at='sheets['+si+'].';
      if(typeof s.id!=='string'||!s.id||ids.has(s.id)) fail('重複したsheet id'); ids.add(s.id);
      fields('sheet',s,at);
      if(typeof s.name!=='string'||!s.name||s.name.length>31||/[\\/?*\[\]:]/.test(s.name)||names.has(s.name.toLowerCase())) fail('シート名'); names.add(s.name.toLowerCase());
      if(!integer(s.rows,1,500)||!integer(s.cols,1,100)) fail('上限500行 / 100列');
      if(!s.cells||typeof s.cells!=='object'||Array.isArray(s.cells)||!Array.isArray(s.merges)) fail('cells / merges');
      for(const [key,cell] of Object.entries(s.cells)) {
        const [r,c]=coord(key); if(key!==addr(r,c)||r>=s.rows||c>=s.cols) fail('範囲外のセル '+key);
        if(!cell||!CELL_TYPES.includes(cell.t)) fail('型 '+key);
        fields('cell',cell,at+'cells.'+key+'.');
        if(cell.t==='s'&&typeof cell.v!=='string'||cell.t==='n'&&!Number.isFinite(cell.v)||cell.t==='b'&&typeof cell.v!=='boolean') fail('値 '+key);
        if(cell.t==='f'&&(typeof cell.f!=='string'||!cell.f.startsWith('='))) fail('数式 '+key);
        if(cell.link && (typeof cell.link!=='string'||!/^https?:\/\//i.test(cell.link))) fail('リンクはhttp(s)のみ');
        vs(cell.s);
      }
      vs(s.defaultStyle);
      for(const [field,maxLen,min,max] of [['colWidths',s.cols,...SIZE_LIMITS.colWidth],['rowHeights',s.rows,...SIZE_LIMITS.rowHeight]]) {
        if(!Array.isArray(s[field])||s[field].length>maxLen||s[field].some(v=>v!==null&&(!Number.isFinite(v)||v<min||v>max))) fail(field);
      }
      for(let i=0;i<s.merges.length;i++){
        const m=s.merges[i]; if(!m||!integer(m.r1,0,s.rows-1)||!integer(m.r2,m.r1,s.rows-1)||!integer(m.c1,0,s.cols-1)||!integer(m.c2,m.c1,s.cols-1)) fail('結合範囲');
        fields('range',m,at+'merges['+i+'].');
        if(s.merges.slice(0,i).some(x=>overlap(x,m))) fail('結合の重複');
        each(m,(r,c)=>{if(r===m.r1&&c===m.c1)return;const cell=get(s,r,c);if(cell.t==='f'||cell.v!==''&&cell.v!==undefined)fail('結合セル内の隠れた値 '+addr(r,c));});
      }
      if(s.freeze&&(!integer(s.freeze.rows,0,s.rows)||!integer(s.freeze.cols,0,s.cols))) fail('固定範囲');
      if(s.freeze&&s.merges.some(m=>m.r1<s.freeze.rows&&m.r2>=s.freeze.rows||m.c1<s.freeze.cols&&m.c2>=s.freeze.cols)) fail('結合をまたぐ固定は未対応');
      if(s.freeze) fields('freeze',s.freeze,at+'freeze.');
      if(s.gridlines!==undefined&&typeof s.gridlines!=='boolean') fail('gridlines（true / false）');
      if(s.images!==undefined){
        if(!Array.isArray(s.images)||s.images.length>100) fail('images（1シート100枚まで）');
        const seen=new Set();
        for(const [pi,p] of s.images.entries()){
          if(!p||typeof p.id!=='string'||!p.id||seen.has(p.id)||typeof p.asset!=='string'||!p.asset||!integer(p.r,0,s.rows-1)||!integer(p.c,0,s.cols-1)||![p.x,p.y].every(v=>Number.isFinite(v)&&v>=0&&v<=20000)||![p.w,p.h].every(v=>Number.isFinite(v)&&v>=4&&v<=10000)) fail('画像 '+(p&&p.id||''));
          seen.add(p.id);
          fields('image',p,at+'images['+pi+'].');
          if(!assetsDetached&&!(w.assets&&Object.prototype.hasOwnProperty.call(w.assets,p.asset))) fail('画像の中身がありません '+p.id);
        }
      }
    }
    if(!ids.has(w.activeSheet)) fail('activeSheet');
    if(w.view){
      if(typeof w.view!=='object'||Array.isArray(w.view)||!Number.isFinite(w.view.zoom)||w.view.zoom<50||w.view.zoom>200)fail('zoom');
      const s=w.sheets.find(s=>s.id===w.activeSheet),v=w.view;
      fields('view',v,'view.');
      if(v.active!==undefined&&(!Array.isArray(v.active)||v.active.length!==2||!integer(v.active[0],0,s.rows-1)||!integer(v.active[1],0,s.cols-1)))fail('view.active');
      if(v.selection!==undefined){const q=v.selection;if(!q||!integer(q.r1,0,s.rows-1)||!integer(q.r2,q.r1,s.rows-1)||!integer(q.c1,0,s.cols-1)||!integer(q.c2,q.c1,s.cols-1))fail('view.selection');fields('range',q,'view.selection.');}
      for(const key of ['scrollTop','scrollLeft'])if(v[key]!==undefined&&(!Number.isFinite(v[key])||v[key]<0))fail('view.'+key);
    }
    return w;
  }
  const width=(s,c)=>s.colWidths[c]??96;
  const height=(s,r)=>s.rowHeights[r]??22;
  function infer(text, old) {
    text=String(text).replace(/\r\n?/g,'\n'); const st=old?.s ? {s:clone(old.s)} : {};
    if(text.startsWith("'")) return {...st,t:'s',v:text.slice(1)};
    if(old?.s?.numFmt==='@') return {...st,t:'s',v:text};
    if(text.startsWith('=')) return {...st,t:'f',f:text};
    if(/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(text)&&text.replace(/\D/g,'').length<=15&&Number.isFinite(+text)) return {...st,t:'n',v:+text};
    if(/^(TRUE|FALSE)$/.test(text)) return {...st,t:'b',v:text==='TRUE'};
    return {...st,t:'s',v:text};
  }
  const raw=cell=>cell.t==='f'?cell.f:cell.t==='b'?(cell.v?'TRUE':'FALSE'):String(cell.v??'');
  const isErr=F.isErr,BLANK=F.BLANK;
  // The value of the cell at (r, c): a formula computed (each cell once per cache), or the value written in the cell. An
  // empty cell is "" here; a formula reads it as F.BLANK (0, "" or FALSE, by what it is used for).
  function calculate(w, sheet, r, c, stack=new Set(), cache=new Map()) {
    const key=sheet.id+'!'+addr(r,c); if(cache.has(key))return cache.get(key); if(stack.has(key))return '#CIRC!';
    const cell=get(sheet,r,c);if(cell.t!=='f')return cell.v??'';
    stack.add(key);let value;
    // What a formula reads from a cell: BLANK for an empty cell, which ISBLANK and COUNTA tell from a formula that gives "".
    const read=(rr,cc)=>{const x=get(sheet,rr,cc);return x.t!=='f'&&(x.v===''||x.v===undefined)?BLANK:calculate(w,sheet,rr,cc,stack,cache);};
    // A range gives the values of its cells row by row, with the range itself (for the implicit intersection).
    const ref=(a,b)=>{const [rr,cc]=coord(a);if(rr>=sheet.rows||cc>=sheet.cols)return '#REF!';if(!b)return read(rr,cc);const q=range(a+':'+b),v=[];if(q.r2>=sheet.rows||q.c2>=sheet.cols)return '#REF!';each(q,(r1,c1)=>v.push(read(r1,c1)));v.range=q;return v;};
    const formulaAt=(rr,cc)=>{const x=get(sheet,rr,cc);return x.t==='f'?x.f:null;};
    try {value=F.evalFormula(cell.f,ref,false,{row:r,col:c,formulaAt});}
    catch(e){value=isErr(e.message)?e.message:'#VALUE!';}
    stack.delete(key); if(typeof value==='number'&&!Number.isFinite(value))value='#NUM!';cache.set(key,value); return value;
  }
  const evalFormula=F.evalFormula;
  function display(value,fmt='General') {
    if(isErr(value)||typeof value==='string')return value;if(typeof value==='boolean')return value?'TRUE':'FALSE';
    if(value==null)return ''; if(fmt==='yyyy/mm/dd'){try{return F.formatValue(value,'yyyy/mm/dd');}catch(e){return '#VALUE!';}}
    if(fmt==='General'||fmt==='@')return String(Number(value.toPrecision(15)));
    const percent=fmt.includes('%'),decimals=fmt.includes('.00')?2:0;
    const yen=fmt.startsWith('¥'),neg=yen&&value<0;
    const text=((neg?-value:value)*(percent?100:1)).toLocaleString('en-US',{minimumFractionDigits:decimals,maximumFractionDigits:decimals,useGrouping:fmt.includes(',')});
    return (neg?'-':'')+(yen?'¥':'')+text+(percent?'%':'');
  }
  // Horizontal alignment Excel shows for "general": numbers and errors right, booleans centre, text left.
  const generalAlign=value=>typeof value==='number'||isErr(value)?'right':typeof value==='boolean'?'center':'left';
  // Screen CSS. Indent is 1 character (about 9px at 11pt) per level on the aligned side, as in Excel.
  function css(st,value) {
    const align=st.align==='general'?generalAlign(value):st.align,indent=st.indent*9;
    const out={ 'font-family':'"'+st.font.replace(/["\\]/g,'')+'", sans-serif','font-size':st.size+'pt','font-weight':st.bold?'700':'400','font-style':st.italic?'italic':'normal','text-decoration':[(st.underline?'underline':''),(st.strike?'line-through':'')].filter(Boolean).join(' ')||'none','color':st.color,'background-color':st.fill==='none'?'transparent':st.fill,'text-align':align,'vertical-align':st.valign,'white-space':st.wrap?'pre-wrap':'pre','padding-left':(3+(align==='right'?0:indent))+'px','padding-right':(3+(align==='right'?indent:0))+'px' };
    for(const side of ['Top','Right','Bottom','Left']) {const v=st['border'+side];if(v!=='none')out['border-'+side.toLowerCase()]=v;}
    return Object.entries(out).map(([k,v])=>k+':'+v).join(';');
  }
  const JP_FONT=/^(?:yu |ms |biz ud|ud digi|meiryo|hg|noto (?:sans|serif) (?:jp|cjk)|hiragino|source han|ipa|takao)/i;
  const fontCharset=name=>/[^\x00-\x7f]/.test(name)||JP_FONT.test(name)?128:0;
  // Clipboard CSS in the dialect Excel writes and reads (pt borders, mso-* properties, explicit nowrap).
  function xlCss(st,value) {
    const align=st.indent>0&&st.align==='general'?'left':st.align,out=[
      'font-family:"'+st.font.replace(/["\\]/g,'')+'"','mso-font-charset:'+fontCharset(st.font),'font-size:'+st.size+'pt',
      'font-weight:'+(st.bold?700:400),'font-style:'+(st.italic?'italic':'normal'),
      'text-decoration:'+([(st.underline?'underline':''),(st.strike?'line-through':'')].filter(Boolean).join(' ')||'none'),
      'color:'+st.color.toUpperCase(),'text-align:'+align,'vertical-align:'+st.valign,'white-space:'+(st.wrap?'normal':'nowrap'),
      'mso-number-format:'+MSO_FORMAT[st.numFmt]];
    if(st.underline)out.push('text-underline-style:single');
    if(st.fill!=='none')out.push('background:'+st.fill.toUpperCase(),'mso-pattern:black none');
    if(st.indent>0)out.push('mso-char-indent-count:'+st.indent,(align==='right'?'padding-right:':'padding-left:')+(st.indent*15)+'px');
    for(const side of ['Top','Right','Bottom','Left']){const v=st['border'+side];out.push('border-'+side.toLowerCase()+':'+(v==='none'?'none':BORDER_XL[v.split(' ').slice(0,2).join(' ')]+' '+v.split(' ')[2].toUpperCase()));}
    return out.join(';');
  }
  // Strings Excel would turn into a number, date, time, boolean, error or formula when it reads the text.
  function excelWouldConvert(text) {
    const s=String(text).trim();if(s!==String(text)&&s==='')return false;
    if(/^[=+\-@]/.test(s)&&s.length>1||/^'/.test(String(text)))return true;
    if(/^(?:TRUE|FALSE)$/i.test(s)||/^#(?:N\/A|NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|GETTING_DATA)$/i.test(s))return true;
    if(/^[(]?[-+]?[¥$€£]?\s*(?:\d[\d,]*(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?\s*%?[)]?$/i.test(s))return true;
    if(/^\d{1,4}[/\-.]\d{1,2}(?:[/\-.]\d{1,4})?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/.test(s))return true;
    if(/^\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s*[AP]M?)?$/i.test(s))return true;
    if(/^(?:\d{1,4}年)?\d{1,2}月(?:\d{1,2}日)?$|^\d{1,2}日$|^(?:明治|大正|昭和|平成|令和|[MTSHR])\s*\d{1,2}[年./-]/.test(s))return true;
    if(/^\d{1,2}[\s-](?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*(?:[\s-]\d{2,4})?$/i.test(s)||/^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s-]\d{1,4}$/i.test(s))return true;
    return false;
  }
  // x:fmla is read relative to the pasted block (even $-references), so rewrite references against the
  // copied range's top-left. Returns null when any reference points outside the copied range.
  function blockFormula(f,q) {
    let outside=false;
    const out=F.englishNames(f).replace(/"(?:[^"]|"")*"|\$?[A-Z]{1,3}\$?[1-9]\d*(?![A-Z0-9_(])/gi,m=>{if(m[0]==='"')return m;const p=/^(\$?)([A-Z]+)(\$?)(\d+)$/i.exec(m),[r,c]=coord(m);if(!within(r,c,q)){outside=true;return m;}return p[1]+colName(c-q.c1)+p[3]+(r-q.r1+1);});
    return outside?null:out;
  }
  const attr=v=>esc(v).replace(/\n/g,'&#10;').replace(/\r/g,'&#13;').replace(/\t/g,'&#9;');

  /* ---- Reading HTML that Excel (desktop or web) put on the clipboard. Pure functions; the DOM walk is in app.js. ---- */
  const cssUnescape=v=>String(v).replace(/\\([0-9a-f]{1,6})\s?/gi,(_,h)=>String.fromCodePoint(parseInt(h,16))).replace(/\\(.)/g,'$1');
  function parseDecls(text) {
    const out={};let buf='',quote='';
    const flush=()=>{const i=buf.indexOf(':');if(i>0)out[buf.slice(0,i).trim().toLowerCase()]=buf.slice(i+1).trim();buf='';};
    for(let i=0;i<String(text||'').length;i++){const ch=text[i];if(quote){buf+=ch;if(ch==='\\'){buf+=text[++i]||'';}else if(ch===quote)quote='';}else if(ch==='"'||ch==="'"){quote=ch;buf+=ch;}else if(ch===';')flush();else buf+=ch;}
    flush();return out;
  }
  // Rules of <style> blocks as [selector, declarations]; Excel writes "td{...}" and ".xl65{...}".
  function parseRules(text) {
    const rules=[],src=String(text||'').replace(/<!--|-->/g,'').replace(/\/\*[\s\S]*?\*\//g,'');
    for(const m of src.matchAll(/([^{}]+)\{([^{}]*)\}/g))for(const sel of m[1].split(','))rules.push([sel.trim().toLowerCase(),parseDecls(m[2])]);
    return rules;
  }
  const NAMED={black:'#000000',windowtext:'#000000',white:'#ffffff',window:'#ffffff',red:'#ff0000',green:'#008000',blue:'#0000ff',yellow:'#ffff00',gray:'#808080',grey:'#808080',silver:'#c0c0c0',maroon:'#800000',navy:'#000080',purple:'#800080',teal:'#008080',olive:'#808000',lime:'#00ff00',aqua:'#00ffff',fuchsia:'#ff00ff',orange:'#ffa500'};
  function cssHex(v) {
    v=String(v||'').trim().toLowerCase();if(NAMED[v])return NAMED[v];
    let m=/^#([\da-f]{6})$/.exec(v);if(m)return '#'+m[1];m=/^#([\da-f])([\da-f])([\da-f])$/.exec(v);if(m)return '#'+m[1]+m[1]+m[2]+m[2]+m[3]+m[3];
    m=/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(v);if(m)return '#'+m.slice(1,4).map(n=>(+n).toString(16).padStart(2,'0')).join('');
    return null;
  }
  // "Standard", "\0022¥\0022\#\,\#\#0", "yyyy\/mm\/dd" ... -> one of FORMATS, or null when not representable.
  function formatFromMso(v) {
    let f=cssUnescape(String(v||'').trim().replace(/^(['"])([\s\S]*)\1$/,'$2'));
    const named={general:'General',standard:'#,##0.00',fixed:'0.00',percent:'0.00%','short date':'yyyy/mm/dd','\\@':'@'};
    if(named[f.toLowerCase()])return named[f.toLowerCase()];
    f=f.replace(/"([^"]*)"/g,'$1').replace(/\[\$([^\]-]*)(?:-[^\]]*)?\]/g,'$1').replace(/_.|\*./g,'').replace(/\\(.)/g,'$1').trim();
    if(/^[\\¥￥]#,##0$/.test(f))f='¥#,##0';
    if(/^yyyy\/m{1,2}\/d{1,2}$/i.test(f))f='yyyy/mm/dd';
    return FORMATS.includes(f)?f:null;
  }
  function borderFromCss(v) {
    v=String(v||'').trim().toLowerCase();if(!v||/^none\b|^0(?:pt|px)?\b|hidden/.test(v))return 'none';
    const width=/([\d.]+)(pt|px)/.exec(v),pts=width?(+width[1])*(width[2]==='px'?.75:1):.5,color=cssHex((v.match(/#[\da-f]{3,6}\b|rgba?\([^)]*\)|\b[a-z]+\b(?=\s*$)/g)||[]).pop())||'#000000';
    if(v.includes('hairline'))return '0.5px solid '+color;
    if(v.includes('double'))return '3px double '+color;
    if(/dot-dash|dash|dot-dot/.test(v)&&!/\bdotted\b/.test(v))return (pts>=1?'2px':'1px')+' dashed '+color;
    if(v.includes('dotted'))return '1px dotted '+color;
    return (pts>=1.5?'3px':pts>=1?'2px':'1px')+' solid '+color;
  }
  // Japanese display names Excel writes for fonts this workbook names in English.
  const FONT_ALIAS={'游ゴシック':'Yu Gothic','游明朝':'Yu Mincho','メイリオ':'Meiryo','ＭＳ ゴシック':'MS Gothic','ＭＳ Ｐゴシック':'MS PGothic','ＭＳ 明朝':'MS Mincho','ＭＳ Ｐ明朝':'MS PMincho'};
  // Excel-dialect declarations -> this workbook's style keys (only those the source states).
  function styleFromDecls(d) {
    const st={};
    if(d['font-family']){let f=cssUnescape(d['font-family'].split(',')[0]).replace(/["']/g,'').trim();f=FONT_ALIAS[f]||f;if(f&&f.length<=100&&!/[<>;{}"\\]/.test(f))st.font=f;}
    if(d['font-size']){const m=/([\d.]+)\s*(pt|px)?/.exec(d['font-size']);if(m){const p=+m[1]*(m[2]==='px'?.75:1);if(p>=6&&p<=96)st.size=Math.round(p*2)/2;}}
    if(d['font-weight'])st.bold=/bold|[6-9]00/.test(d['font-weight']);
    if(d['font-style'])st.italic=/italic|oblique/.test(d['font-style']);
    if(d['text-decoration']!==undefined||d['text-underline-style']!==undefined){const t=(d['text-decoration']||'')+' '+(d['text-underline-style']||'');st.underline=/underline|single|double/.test(t)&&!/none\s*$/.test(d['text-underline-style']||'');st.strike=/line-through/.test(t);}
    if(d.color){const c=cssHex(d.color);if(c)st.color=c;}
    const bg=d['background-color']??d.background;if(bg!==undefined)st.fill=cssHex(bg.split(/\s+/)[0])||'none';
    if(d['text-align']){const a=d['text-align'].toLowerCase();st.align=['left','center','right'].includes(a)?a:a==='justify'||a==='fill'?'left':'general';}
    if(d['vertical-align']){const a=d['vertical-align'].toLowerCase();st.valign=a==='top'?'top':a==='bottom'?'bottom':'middle';}
    if(d['white-space'])st.wrap=/normal|pre-wrap|pre-line/.test(d['white-space']);
    if(d['mso-char-indent-count']){const n=parseInt(d['mso-char-indent-count'],10);if(n>=0&&n<=15)st.indent=n;}
    if(d['mso-number-format']){const f=formatFromMso(d['mso-number-format']);st.numFmt=f||'General';}
    const all=d.border!==undefined?borderFromCss(d.border):undefined;
    for(const side of ['Top','Right','Bottom','Left']){const v=d['border-'+side.toLowerCase()];if(v!==undefined)st['border'+side]=borderFromCss(v);else if(all!==undefined)st['border'+side]=all;}
    return st;
  }
  // Number shown with a format ("¥1,200", "12.35%", "2025/12/09") back to its value; null when it is not a number.
  function valueFromDisplay(text,fmt) {
    const s=String(text).trim();if(!s)return null;
    let m=/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(s);
    if(m){const t=Date.UTC(+m[1],+m[2]-1,+m[3]);return Math.round((t-Date.UTC(1899,11,30))/86400000);}
    m=/^([(-])?\s*[¥￥$]?\s*([\d,]*\.?\d+(?:e[-+]?\d+)?)\s*(%)?\)?$/i.exec(s);
    if(!m||fmt==='@')return null;const n=+m[2].replace(/,/g,'');if(!Number.isFinite(n))return null;
    return (m[1]?-1:1)*n/(m[3]?100:1);
  }
  function shiftFormula(f,dr,dc){return f.replace(/"(?:[^"]|"")*"|\$?[A-Z]{1,3}\$?[1-9]\d*/gi,m=>{if(m[0]==='"')return m;const p=/^(\$?)([A-Z]+)(\$?)(\d+)$/i.exec(m),[r,c]=coord(m),nr=r+(p[3]?0:dr),nc=c+(p[1]?0:dc);return nr<0||nc<0||nr>=500||nc>=100?'#REF!':p[1]+colName(nc)+p[3]+(nr+1);});}
  // rangeData, clipboard and markdown copy exactly the range q, as Excel does: a merged cell wholly inside q stays
  // merged, and one only partly inside q is copied as separate cells (its value goes with its top-left cell).
  // A selection of cells always takes in whole merged cells (selectionArea), so only whole rows or columns differ.
  function rangeData(w,s,q) {
    const cells={},merges=s.merges.filter(m=>inside(m,q)).map(m=>({r1:m.r1-q.r1,r2:m.r2-q.r1,c1:m.c1-q.c1,c2:m.c2-q.c1}));
    each(q,(r,c)=>{const cell=clone(get(s,r,c));cell.s=style(w,s,cell);cells[addr(r-q.r1,c-q.c1)]=cell;});
    return {version:1,source:{book:w.id,sheet:s.id,r:q.r1,c:q.c1},rows:q.r2-q.r1+1,cols:q.c2-q.c1+1,cells,merges,colWidths:Array.from({length:q.c2-q.c1+1},(_,i)=>width(s,q.c1+i)),rowHeights:Array.from({length:q.r2-q.r1+1},(_,i)=>height(s,q.r1+i))};
  }
  const tsvQuote=v=>/[\t\r\n"]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;
  // Clipboard HTML in Excel's own dialect. Measured against desktop Excel 2021 (see references/excel-compatibility.md):
  // x:str carries exact text (leading zeros, spaces, line breaks), x:num full-precision numbers, x:fmla formulas
  // whose references stay inside the copied range. Strings Excel would convert are also marked "\@" so that a
  // reader which ignores x:str (possibly Excel for the web) still keeps the text.
  function clipboard(w,s,q) {
    const data=rangeData(w,s,q),cache=new Map(),plain=[];
    const pt=px=>+(px*.75).toFixed(2),totalW=Array.from({length:q.c2-q.c1+1},(_,i)=>width(s,q.c1+i)).reduce((a,b)=>a+b,0);
    let html='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">'+
      '<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"><meta name="ProgId" content="Excel.Sheet"></head><body>'+
      '<table border="0" cellpadding="0" cellspacing="0" width="'+totalW+'" style="border-collapse:collapse;table-layout:fixed;width:'+pt(totalW)+'pt" data-excel-html="'+esc(JSON.stringify(data))+'">';
    for(let c=q.c1;c<=q.c2;c++)html+='<col width="'+width(s,c)+'" style="mso-width-source:userset;width:'+pt(width(s,c))+'pt">';
    for(let r=q.r1;r<=q.r2;r++){
      html+='<tr height="'+height(s,r)+'" style="mso-height-source:userset;height:'+pt(height(s,r))+'pt">';const row=[];
      for(let c=q.c1;c<=q.c2;c++) {
        const merge=mergeIn(s,r,c,q),hidden=merge&&(r!==merge.r1||c!==merge.c1),cell=get(s,r,c),value=calculate(w,s,r,c,new Set(),cache),st=style(w,s,cell),label=display(value,st.numFmt);row.push(tsvQuote(hidden?'':label));if(hidden)continue;
        const colspan=merge?merge.c2-merge.c1+1:1,rowspan=merge?merge.r2-merge.r1+1:1,cw=Array.from({length:colspan},(_,i)=>width(s,c+i)).reduce((a,b)=>a+b,0);
        let attrs='',cst=st;
        if(cell.t==='f'){const f=blockFormula(cell.f,q);if(f)attrs+=' x:fmla="'+attr(f)+'"';}
        if(typeof value==='number')attrs+=' x:num="'+value+'"';
        else if(typeof value==='boolean')attrs+=' x:bool="'+(value?'TRUE':'FALSE')+'"';
        else if(isErr(value))attrs+=' x:err="'+esc(value)+'"';
        else if(value!==''){attrs+=' x:str="'+attr(value.startsWith("'")?"'"+value:value)+'"';if(st.numFmt!=='@'&&excelWouldConvert(value)&&cell.t!=='f')cst={...st,numFmt:'@'};}
        html+='<td'+(colspan>1?' colspan="'+colspan+'"':'')+(rowspan>1?' rowspan="'+rowspan+'"':'')+' width="'+cw+'" height="'+height(s,r)+'"'+attrs+' style="'+esc(xlCss(cst,value)+';width:'+pt(cw)+'pt;height:'+pt(height(s,r))+'pt')+'">';
        const content=esc(label).replace(/\n/g,'<br style="mso-data-placement:same-cell">');
        html+=(cell.link?'<a href="'+esc(cell.link)+'">'+content+'</a>':content)+'</td>';
      }
      plain.push(row.join('\t'));html+='</tr>';
    }
    html+='</table></body></html>';return {html,text:plain.join('\r\n'),data};
  }
  // Markdown (GFM) table of what the sheet shows: formatted values, the first row as the header.
  // Pipes are escaped and in-cell line breaks become <br>; a merged area keeps its value in its
  // top-left cell; trailing rows and columns without any value are left out; cells are padded
  // (East Asian wide characters count as two) so that the text also lines up in plain editors.
  function markdown(w,s,q) {
    const cache=new Map(),grid=[];
    for(let r=q.r1;r<=q.r2;r++){
      const row=[];
      for(let c=q.c1;c<=q.c2;c++){
        const merge=mergeIn(s,r,c,q);if(merge&&(r!==merge.r1||c!==merge.c1)){row.push({text:'',align:''});continue;}
        const cell=get(s,r,c),value=calculate(w,s,r,c,new Set(),cache),st=style(w,s,cell),label=String(display(value,st.numFmt));
        let text=label.replace(/\r\n?/g,'\n').replace(/\|/g,'\\|').replace(/\n/g,'<br>');
        if(cell.link&&text)text='['+text.replace(/[[\]]/g,'\\$&')+']('+cell.link.replace(/[ ()|]/g,ch=>'%'+ch.charCodeAt(0).toString(16).toUpperCase())+')';
        row.push({text,align:label===''?'':st.align==='general'?generalAlign(value):st.align});
      }
      grid.push(row);
    }
    while(grid.length>1&&grid[grid.length-1].every(x=>!x.text))grid.pop();
    let cols=grid[0].length;while(cols>1&&grid.every(row=>!row[cols-1].text))cols--;for(const row of grid)row.length=cols;
    const wide=ch=>ch.codePointAt(0)>0xffff||/[\u1100-\u115f\u2e80-\u303e\u3041-\u33ff\u3400-\u4dbf\u4e00-\u9fff\ua000-\ua4cf\uac00-\ud7a3\uf900-\ufaff\ufe30-\ufe4f\uff00-\uff60\uffe0-\uffe6]/.test(ch);
    const len=t=>[...t].reduce((n,ch)=>n+(wide(ch)?2:1),0),body=grid.length>1?grid.slice(1):grid;
    const aligns=Array.from({length:cols},(_,c)=>{const seen=new Set(body.map(row=>row[c].align).filter(Boolean));return seen.size===1?[...seen][0]:'left';});
    const widths=Array.from({length:cols},(_,c)=>Math.max(3,...grid.map(row=>len(row[c].text))));
    const pad=(t,c)=>{const gap=widths[c]-len(t);return aligns[c]==='right'?' '.repeat(gap)+t:aligns[c]==='center'?' '.repeat(gap>>1)+t+' '.repeat(gap-(gap>>1)):t+' '.repeat(gap);};
    const line=cells=>'| '+cells.join(' | ')+' |';
    const rule=aligns.map((a,c)=>a==='right'?'-'.repeat(widths[c]-1)+':':a==='center'?':'+'-'.repeat(widths[c]-2)+':':'-'.repeat(widths[c]));
    const text=[line(grid[0].map((x,c)=>pad(x.text,c))),line(rule),...grid.slice(1).map(row=>line(row.map((x,c)=>pad(x.text,c))))].join('\n');
    return {text,rows:grid.length,cols,empty:grid.every(row=>row.every(x=>!x.text))};
  }
  function parseTSV(text) {
    const rows=[];let row=[],cell='',quoted=false;
    text=text.replace(/\r\n/g,'\n').replace(/\r/g,'\n');
    for(let i=0;i<text.length;i++){const ch=text[i];if(ch==='"'&&quoted&&text[i+1]==='"'){cell+='"';i++;}else if(ch==='"'&&(quoted||cell===''))quoted=!quoted;else if(ch==='\t'&&!quoted){row.push(cell);cell='';}else if(ch==='\n'&&!quoted){row.push(cell);rows.push(row);row=[];cell='';}else cell+=ch;}
    if(quoted)throw Error('引用符が閉じていない貼り付けデータです');if(cell!==''||row.length||!rows.length){row.push(cell);rows.push(row);}return rows;
  }
  function paste(s,data,r,c,{valuesOnly=false,dimensions=false,relocate=true}={}) {
    if(r+data.rows>s.rows||c+data.cols>s.cols)throw Error('貼り付け先がシート範囲を超えます');
    const q={r1:r,c1:c,r2:r+data.rows-1,c2:c+data.cols-1};
    if(s.merges.some(m=>overlap(q,m)&&!within(m.r1,m.c1,q)||overlap(q,m)&&!within(m.r2,m.c2,q)))throw Error('貼り付け範囲が結合セルの一部に重なります');
    if(valuesOnly&&s.merges.some(m=>overlap(q,m)))throw Error('結合範囲への値のみ貼り付けは、先に結合を解除してください');
    if(!valuesOnly)s.merges=s.merges.filter(m=>!overlap(m,q));
    each({r1:0,c1:0,r2:data.rows-1,c2:data.cols-1},(rr,cc)=>{
      const cell=clone(data.cells[addr(rr,cc)]||{t:'s',v:''});if(cell.t==='f'&&relocate&&data.source)cell.f=shiftFormula(cell.f,r-data.source.r,c-data.source.c);
      if(valuesOnly)cell.s=clone(get(s,r+rr,c+cc).s||{});s.cells[addr(r+rr,c+cc)]=cell;
    });
    if(!valuesOnly)for(const m of data.merges||[])s.merges.push({r1:m.r1+r,r2:m.r2+r,c1:m.c1+c,c2:m.c2+c});
    if(dimensions){data.colWidths?.forEach((v,i)=>{s.colWidths[c+i]=v;});data.rowHeights?.forEach((v,i)=>{s.rowHeights[r+i]=v;});}
    return q;
  }
  // The text that 「AIに渡す」 hands to Claude Code. By default the state is written as TOON (toon.js, specification 4.4),
  // which says the same as the JSON in fewer characters, between the first line and an end line. The end line is what
  // shows that a TOON text was not cut short: TOON has no closing braces, so a text cut between two lines could
  // otherwise still read as a whole. The JSON form stays exactly as it always was, and is read as before.
  const STATE_FIRST = 'EXCEL_HTML_STATE_V1', STATE_TOON = STATE_FIRST + ' TOON', STATE_END = 'EXCEL_HTML_STATE_END';
  function stateText(state, format = 'toon') {
    if (format === 'json') return STATE_FIRST + '\n' + JSON.stringify(state, null, 2);
    return STATE_TOON + '\n' + TOON.encode(state) + '\n' + STATE_END;
  }
  function readStateText(text) {
    const raw = String(text).replace(/^﻿/, '').trim(), nl = raw.indexOf('\n');
    const first = (nl < 0 ? raw : raw.slice(0, nl)).replace(/[ \t\r]+$/, '');
    if (first === STATE_FIRST) {
      try { return JSON.parse(raw.slice(nl + 1)); }
      catch (e) { throw Error('状態テキストのJSONを読み取れません（途中で切れているか、書き換えられています）: ' + e.message); }
    }
    if (first !== STATE_TOON) throw Error('編集状態はEXCEL_HTML_STATE_V1で始まる完全なテキストを渡してください。');
    const end = raw.lastIndexOf('\n');
    if (nl < 0 || raw.slice(end + 1).replace(/[ \t\r]+$/, '') !== STATE_END)
      throw Error('状態テキストの最後の行が ' + STATE_END + ' ではありません。途中で切れているので、コピーし直したテキストか、エクスポートしたファイルを渡してください。');
    return TOON.decode(end > nl ? raw.slice(nl + 1, end) : '', {firstLine: 2});
  }
  return {SCHEMA,SIZE_LIMITS,FORMATS,MSO_FORMAT,BORDER_XL,DEFAULT,FONT_ALIAS,CHOICES,CELL_TYPES,FUNCTIONS,FIELDS,stateText,readStateText,selectionArea,splitsMerge,xlCss,excelWouldConvert,blockFormula,generalAlign,
    fontCharset,cssUnescape,parseDecls,parseRules,cssHex,formatFromMso,borderFromCss,styleFromDecls,valueFromDisplay,clone,esc,colName,addr,coord,rect,range,within,overlap,same,each,get,style,mergeAt,expanded,validate,width,height,infer,raw,isErr,calculate,evalFormula,display,css,shiftFormula,rangeData,clipboard,markdown,parseTSV,paste};
});
