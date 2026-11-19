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
  // A key of the object itself: toString or constructor, which every object has, are not style keys.
  const own = (o,k) => Object.prototype.hasOwnProperty.call(o,k);
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
  // What a style value has to be, or '' when it is one. A wrong value is told with where it was written and what is
  // taken, so that a book written by hand can be mended at once.
  function styleProblem(k,v) {
    const hex=x=>typeof x==='string'&&/^#[\da-f]{6}$/i.test(x);
    if(k==='color') return hex(v)?'':'#rrggbb の6桁の色';
    if(k==='fill') return v==='none'||hex(v)?'':'#rrggbb の6桁の色か none';
    if(k==='font') return typeof v==='string'&&v&&v.length<=100&&!/[<>;{}"\\]/.test(v)?'':'フォント名。100文字まで。< > ; { } " \\ を含まない';
    if(k==='size') return Number.isFinite(v)&&v>=6&&v<=96?'':'6〜96 の数';
    if(k==='indent') return Number.isInteger(v)&&v>=0&&v<=15?'':'0〜15 の整数';
    if(k==='align'||k==='valign') return CHOICES[k].includes(v)?'':CHOICES[k].join(' / ');
    if(k==='numFmt') return FORMATS.includes(v)?'':FORMATS.join(' / ');
    if(k.startsWith('border')) return v==='none'||typeof v==='string'&&/^\S+ \S+ #[\da-f]{6}$/i.test(v)&&own(BORDER_XL,v.split(' ').slice(0,2).join(' '))?'':
      'none か、線と色（例 1px solid #808080）。線は '+Object.keys(BORDER_XL).join(' / ')+'、色は #rrggbb';
    return typeof v==='boolean'?'':'true / false';
  }
  // assetsDetached: the editor keeps picture bytes outside the book while editing, so references are not checked then.
  function validate(w,{assetsDetached=false}={}) {
    const fail=m=>{throw Error('ブック形式エラー: '+m);};
    const integer=(n,min,max)=>Number.isInteger(n)&&n>=min&&n<=max;
    // A field the page does not have is refused, as an unknown style key is, rather than kept unseen and lost later.
    const fields=(kind,o,where)=>{for(const k of Object.keys(o))if(!FIELDS[kind].includes(k))fail('未対応の項目 '+where+k);};
    if(!w||typeof w!=='object'||Array.isArray(w)) fail('ブックのオブジェクト');
    if(w.schema!==SCHEMA) fail('schema（"'+SCHEMA+'"）');
    if(typeof w.id!=='string'||!w.id) fail('id（空でない文字）');
    if(typeof w.revision!=='string'||!w.revision) fail('revision（空でない文字。"draft" でよい）');
    fields('book',w,'');
    if(typeof w.title!=='string'||!Array.isArray(w.sheets)||!w.sheets.length||w.sheets.length>30) fail('title（文字）と sheets（1〜30枚のシートの配列）');
    if(w.assets!==undefined){
      if(!w.assets||typeof w.assets!=='object'||Array.isArray(w.assets)) fail('assets');
      for(const [k,v] of Object.entries(w.assets)) if(!k||typeof v!=='string'||v.length>15e6||!PICTURE_DATA.test(v)) fail('画像データ '+k+'（png / jpeg / gif / webp の data URL）');
    }
    const ids=new Set(), names=new Set();
    const vs = (st,where) => { if(!st) return; if(typeof st!=='object'||Array.isArray(st)) fail(where+'（書式のオブジェクト）');
      for(const [k,v] of Object.entries(st)) {
        if(!own(DEFAULT,k)) fail('未対応の書式 '+where+'.'+k);
        const problem=styleProblem(k,v); if(problem) fail(where+'.'+k+' '+JSON.stringify(v)+'（'+problem+'）');
      }
    };
    vs(w.defaultStyle,'defaultStyle');
    for(const [si,s] of w.sheets.entries()) {
      const at='sheets['+si+'].';
      if(typeof s.id!=='string'||!s.id||ids.has(s.id)) fail(at+'id（シートごとに違う、空でない文字。重複したsheet id は使えない）'); ids.add(s.id);
      fields('sheet',s,at);
      if(typeof s.name!=='string'||!s.name||s.name.length>31||/[\\/?*\[\]:]/.test(s.name)||names.has(s.name.toLowerCase())) fail(at+'name '+JSON.stringify(s.name)+'（シート名は31文字以内で \\ / ? * [ ] : を含まず、ほかのシートと大文字・小文字の違いだけで重ならない）'); names.add(s.name.toLowerCase());
      if(!integer(s.rows,1,500)||!integer(s.cols,1,100)) fail(at+'rows と cols（上限500行 / 100列の、1以上の整数）');
      if(!s.cells||typeof s.cells!=='object'||Array.isArray(s.cells)||!Array.isArray(s.merges)) fail(at+'cells（セル番地をキーにしたオブジェクト）と merges（配列）');
      for(const [key,cell] of Object.entries(s.cells)) {
        let r,c;try{[r,c]=coord(key);}catch(e){fail(at+'cells.'+key+'（"A1" のようなセル番地）');}
        if(key!==addr(r,c)) fail(at+'cells.'+key+'（"'+addr(r,c)+'" のように、大文字で $ のない番地）');
        if(r>=s.rows||c>=s.cols) fail(at+'cells.'+key+' 範囲外のセル（シートは'+s.rows+'行・'+s.cols+'列）');
        if(!cell||!CELL_TYPES.includes(cell.t)) fail('型 '+key);
        fields('cell',cell,at+'cells.'+key+'.');
        if(cell.t==='s'&&typeof cell.v!=='string'||cell.t==='n'&&!Number.isFinite(cell.v)||cell.t==='b'&&typeof cell.v!=='boolean') fail(at+'cells.'+key+'.v（t が s なら文字、n なら有限の数、b なら true / false）');
        if(cell.t==='f'&&(typeof cell.f!=='string'||!cell.f.startsWith('='))) fail('数式 '+key);
        if(cell.link && (typeof cell.link!=='string'||!/^https?:\/\//i.test(cell.link))) fail('リンクはhttp(s)のみ（'+at+'cells.'+key+'.link）');
        vs(cell.s,at+'cells.'+key+'.s');
      }
      vs(s.defaultStyle,at+'defaultStyle');
      for(const [field,maxLen,min,max] of [['colWidths',s.cols,...SIZE_LIMITS.colWidth],['rowHeights',s.rows,...SIZE_LIMITS.rowHeight]]) {
        if(!Array.isArray(s[field])||s[field].length>maxLen||s[field].some(v=>v!==null&&(!Number.isFinite(v)||v<min||v>max))) fail(at+field+'（'+min+'〜'+max+' のピクセルの数か null を、'+(field==='colWidths'?'列':'行')+'の数まで並べる）');
      }
      for(let i=0;i<s.merges.length;i++){
        const m=s.merges[i]; if(!m||typeof m!=='object'||!integer(m.r1,0,s.rows-1)||!integer(m.r2,m.r1,s.rows-1)||!integer(m.c1,0,s.cols-1)||!integer(m.c2,m.c1,s.cols-1))
          fail(at+'merges['+i+']（{"r1":0,"c1":0,"r2":0,"c2":4} のような、0始まりで両端を含む行と列の番号。シートの中に収める）');
        fields('range',m,at+'merges['+i+'].');
        if(s.merges.slice(0,i).some(x=>overlap(x,m))) fail('結合の重複（'+at+'merges['+i+'] '+addr(m.r1,m.c1)+':'+addr(m.r2,m.c2)+'）');
        each(m,(r,c)=>{if(r===m.r1&&c===m.c1)return;const cell=get(s,r,c);if(cell.t==='f'||cell.v!==''&&cell.v!==undefined)fail('結合セル内の隠れた値 '+addr(r,c));});
      }
      if(s.freeze&&(typeof s.freeze!=='object'||!integer(s.freeze.rows,0,s.rows)||!integer(s.freeze.cols,0,s.cols)))
        fail(at+'freeze（rows と cols の両方を、0 から行の数・列の数までの整数で書く。固定しない方は 0）');
      const crossRows=m=>m.r1<s.freeze.rows&&m.r2>=s.freeze.rows,crossing=s.freeze&&s.merges.find(m=>crossRows(m)||m.c1<s.freeze.cols&&m.c2>=s.freeze.cols);
      if(crossing) fail('結合をまたぐ固定は未対応（シート「'+s.name+'」の結合 '+addr(crossing.r1,crossing.c1)+':'+addr(crossing.r2,crossing.c2)+' が、固定した'+
        (crossRows(crossing)?'行（'+s.freeze.rows+'行）':'列（'+s.freeze.cols+'列）')+'の境目をまたぐ）');
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
    if(!ids.has(w.activeSheet)) fail('activeSheet '+JSON.stringify(w.activeSheet)+'（最初に開くシートの id。'+[...ids].join(' / ')+' のどれか）');
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
  // The shorthand of a book written by hand (build.py --input), made into the cells it stands for, as if each had been
  // written in cells. values: rows of values from the cell at. formulas: a formula copied over a range, its relative
  // references moved as Excel moves them. styles: styles laid over ranges in the order written, over the style a cell
  // gives itself; as Excel keeps one line per cell edge, a line laid on one side of an edge is laid on the other side
  // too where that side has another line, and a merged cell is one block. cells, merges, rowHeights and colWidths may be
  // left out, and a merged cell may be written as a range ("A1:E1"). A value written twice, a range outside the sheet or
  // cutting a merged cell, a line laid on a part of the edge of a merged cell that has another line there (its edge is
  // one line), and a wrong entry or style value are refused with where they are; the rest is left for validate.
  // note(text) hears of texts that read as numbers or dates, which SUM leaves out. The book given is not changed.
  function expand(w,note=()=>{}) {
    const fail=m=>{throw Error('ブック形式エラー: '+m);};
    if(!w||typeof w!=='object'||!Array.isArray(w.sheets)) return w;
    const numberLike=[],dateLike=[];
    const out={...w,sheets:w.sheets.map((sheet,si)=>{
      if(!sheet||typeof sheet!=='object'||Array.isArray(sheet)) return sheet;
      const {values,formulas,styles,...s}=sheet,at='sheets['+si+'].',size='（'+s.rows+'行・'+s.cols+'列）';
      if(Array.isArray(s.rows)) fail(at+'rows は行の数（値の行は values の rows に書く）');
      for(const k of ['merges','rowHeights','colWidths']) if(s[k]===undefined) s[k]=[];
      if(s.cells===undefined) s.cells={};
      // A merged cell may be written as a range ("A1:E1"), as the ranges of values, formulas and styles are.
      if(Array.isArray(s.merges)&&s.merges.some(m=>typeof m==='string')) s.merges=s.merges.map((m,i)=>{
        if(typeof m!=='string') return m;
        let q=null;try{q=range(m);}catch(e){}
        if(!q) fail(at+'merges['+i+'] '+JSON.stringify(m)+'（"A1:E1" のような範囲）');
        if(Number.isInteger(s.rows)&&Number.isInteger(s.cols)&&!(q.r2<s.rows&&q.c2<s.cols)) fail(at+'merges['+i+'] '+m+' がシートの外です'+size);
        return q;});
      if(values===undefined&&formulas===undefined&&styles===undefined) return s;
      if(!s.cells||typeof s.cells!=='object'||Array.isArray(s.cells)||!Array.isArray(s.merges)||!Number.isInteger(s.rows)||!Number.isInteger(s.cols)) return s;
      const cells={...s.cells},given=new Map(),texts=[];
      // A cell holding only a style (empty text) can take a value; it keeps its style and link.
      const valueless=cell=>!!cell&&cell.t==='s'&&(cell.v===''||cell.v===undefined)&&cell.f===undefined;
      for(const [a,cell] of Object.entries(cells)) if(!valueless(cell)) given.set(a,'cells');
      const onSheet=(r,c)=>r<s.rows&&c<s.cols;
      const keys=(o,allowed,where)=>{for(const k of Object.keys(o)) if(!allowed.includes(k)) fail('未対応の項目 '+where+'.'+k);};
      const entry=(x,where,shape)=>{if(!x||typeof x!=='object'||Array.isArray(x)) fail(where+'（'+shape+'）');return x;};
      const list=(x,name)=>{if(x===undefined) return [];if(!Array.isArray(x)) fail(at+name+'（配列）');return x;};
      const area=(text,where)=>{let q=null;try{if(typeof text==='string') q=range(text);}catch(e){}
        if(!q) fail(where+'（"A1:C3" のような範囲）');if(!onSheet(q.r2,q.c2)) fail(where+' '+text+' がシートの外です'+size);return q;};
      const put=(a,cell,where)=>{if(given.has(a)) fail('sheets['+si+'] の '+a+' に値を2回書いています（'+given.get(a)+' と '+where+'）');given.set(a,where);
        const old=cells[a];if(!valueless(old)){cells[a]=cell;return;}
        const merged={...cell};if(old.link!==undefined&&merged.link===undefined) merged.link=old.link;
        if(old.s!==undefined) merged.s={...old.s,...cell.s};cells[a]=merged;};
      list(values,'values').forEach((block,i)=>{
        const where=at+'values['+i+']';entry(block,where,'{"at": "A3", "rows": [[…], …]}');keys(block,['at','rows'],where);
        let r0,c0;try{[r0,c0]=coord(String(block.at));}catch(e){fail(where+'.at（"A3" のようなセル番地）');}
        if(!onSheet(r0,c0)) fail(where+'.at '+block.at+' がシートの外です'+size);
        if(!Array.isArray(block.rows)) fail(where+'.rows（行の配列）');
        block.rows.forEach((row,ri)=>{
          if(!Array.isArray(row)) fail(where+'.rows['+ri+']（値の配列）');
          row.forEach((x,ci)=>{
            if(x===null||x==='') return;
            const r=r0+ri,c=c0+ci,a=addr(r,c),here=where+'.rows['+ri+']['+ci+']';
            if(!onSheet(r,c)) fail(here+' '+a+' がシートの外です'+size);
            let cell;
            if(typeof x==='string'){cell=x.startsWith('=')?{t:'f',f:x}:{t:'s',v:x};const kind=cell.t==='s'&&textKind(x);if(kind) texts.push([a,x,kind]);}
            else if(typeof x==='number') cell={t:'n',v:x};
            else if(typeof x==='boolean') cell={t:'b',v:x};
            else if(x&&typeof x==='object'&&!Array.isArray(x)){if(!CELL_TYPES.includes(x.t)) fail(here+'（セルのオブジェクトには t（s / n / b / f）を書く）');cell=clone(x);}
            else fail(here+'（文字・数値・true / false・null・セルのオブジェクト）');
            put(a,cell,'values['+i+']');
          });
        });
      });
      list(formulas,'formulas').forEach((e,i)=>{
        const where=at+'formulas['+i+']';entry(e,where,'{"range": "E4:E20", "f": "=C4*D4"}');keys(e,['range','f'],where);
        const q=area(e.range,where+'.range');
        if(typeof e.f!=='string'||!e.f.startsWith('=')) fail(where+'.f（= で始まる式）');
        each(q,(r,c)=>put(addr(r,c),{t:'f',f:r===q.r1&&c===q.c1?e.f:shiftFormula(e.f,r-q.r1,c-q.c1)},'formulas['+i+']'));
      });
      const entries=list(styles,'styles').map((e,i)=>{
        const where=at+'styles['+i+']';entry(e,where,'{"range": "A3:H20", "s": {…}}');keys(e,['range','s'],where);
        const q=area(e.range,where+'.range');entry(e.s,where+'.s','書式のオブジェクト');
        for(const [k,v] of Object.entries(e.s)){
          if(!own(DEFAULT,k)) fail('未対応の書式 '+where+'.s.'+k);
          const problem=styleProblem(k,v);if(problem) fail(where+'.s.'+k+' '+JSON.stringify(v)+'（'+problem+'）');
        }
        for(const m of s.merges) if(overlap(m,q)&&!inside(m,q))
          fail(where+'.range '+e.range+' は結合セル '+addr(m.r1,m.c1)+':'+addr(m.r2,m.c2)+' の一部だけを含みます。結合セル全体を含む範囲にするか、範囲から外してください');
        return {q,st:e.s,where,text:e.range};
      });
      if(entries.length){
        const blocks=new Map();for(const m of s.merges) each(m,(r,c)=>blocks.set(addr(r,c),m));
        const holder=(r,c)=>{const m=blocks.get(addr(r,c));return m?addr(m.r1,m.c1):addr(r,c);};
        const base={...DEFAULT,...w.defaultStyle,...s.defaultStyle};
        const now=new Map(),styleOf=a=>{if(!now.has(a)) now.set(a,{...(cells[a]&&cells[a].s)});return now.get(a);};
        // Each side, with the side of the neighbouring cell on the same edge, the way to that neighbour, and the names
        // of the two sides.
        const SIDE={borderTop:['borderBottom',-1,0,'上','下'],borderBottom:['borderTop',1,0,'下','上'],borderLeft:['borderRight',0,-1,'左','右'],borderRight:['borderLeft',0,1,'右','左']};
        // The edge of a merged cell is one line. A line laid from a part of it changes it beyond the range, unless the
        // styles lay the same line on the whole edge (in one range or in several) or on the merged cell itself.
        // pending: edges of merged cells changed so, checked once all the styles lie; written: sides the styles give.
        const pending=new Map(),written=new Map();
        for(const {q,st,where,text} of entries){
          const held=[];each(q,(r,c)=>{const m=blocks.get(addr(r,c));if(!m||m.r1===r&&m.c1===c) held.push(m||{r1:r,c1:c,r2:r,c2:c});});
          for(const b of held){
            const a=addr(b.r1,b.c1);Object.assign(styleOf(a),st);
            for(const k of Object.keys(st)) if(own(SIDE,k)){written.set(a+'|'+k,st[k]);pending.delete(a+'|'+k);}
          }
          for(const [k,line] of Object.entries(st)){
            if(!own(SIDE,k)) continue;const [opp,dr,dc,mine,theirs]=SIDE[k];
            for(const b of held){
              const next=[];
              if(dr){const r=dr<0?b.r1-1:b.r2+1;if(r>=0&&r<s.rows) for(let c=b.c1;c<=b.c2;c++) next.push([r,c]);}
              else{const c=dc<0?b.c1-1:b.c2+1;if(c>=0&&c<s.cols) for(let r=b.r1;r<=b.r2;r++) next.push([r,c]);}
              for(const [r,c] of next){
                if(within(r,c,q)&&own(st,opp)) continue; // the entry gives that side itself
                const side=styleOf(holder(r,c)),there=side[opp]??base[opp];
                if(there==='none'||String(there).toLowerCase()===String(line).toLowerCase()) continue;
                const m=blocks.get(addr(r,c));
                if(m){const facing=dr?{r1:dr>0?m.r1-1:m.r2+1,c1:m.c1,r2:dr>0?m.r1-1:m.r2+1,c2:m.c2}:{r1:m.r1,c1:dc>0?m.c1-1:m.c2+1,r2:m.r2,c2:dc>0?m.c1-1:m.c2+1};
                  pending.set(addr(m.r1,m.c1)+'|'+opp,{m,k,opp,facing,where,text,mine,theirs});}
                side[opp]=line;
              }
            }
          }
        }
        for(const {m,k,opp,facing,where,text,mine,theirs} of pending.values()){
          const line=String(styleOf(addr(m.r1,m.c1))[opp]).toLowerCase();let whole=true;
          each(facing,(r,c)=>{if(String(written.get(holder(r,c)+'|'+k)).toLowerCase()!==line)whole=false;});
          if(!whole) fail(where+'.range '+text+' の'+mine+'の線は、結合セル '+addr(m.r1,m.c1)+':'+addr(m.r2,m.c2)+' の'+theirs+
            'の辺の一部にしか接しません。結合セルの辺は1本の線なので、'+addr(facing.r1,facing.c1)+':'+addr(facing.r2,facing.c2)+' のように辺全体に引いてください');
        }
        for(const [a,st] of now){if(!Object.keys(st).length) continue;const cell=cells[a]?{...cells[a]}:{t:'s',v:''};cell.s=st;cells[a]=cell;}
      }
      for(const [a,x,kind] of texts) if(style(w,s,cells[a]).numFmt!=='@') (kind==='date'?dateLike:numberLike).push(s.name+'!'+a+' '+JSON.stringify(x));
      return {...s,cells};
    })};
    const first=list=>list.slice(0,5).join('、')+(list.length>5?'、ほか'+(list.length-5)+'件':'');
    if(numberLike.length) note('数値に見える文字列が、文字のセルになります（SUM などの計算に入りません）: '+first(numberLike)+
      '。計算する値なら引用符を外して数値で書き、番号のように文字のままにするなら numFmt "@" を付けてください。');
    if(dateLike.length) note('日付や時刻に見える文字列が、文字のセルになります: '+first(dateLike)+
      '。日付として計算するなら、シリアル値（2026/1/2 は 46024）に numFmt "yyyy/mm/dd" を付けて書いてください。文字のままにするなら numFmt "@" を付けてください。');
    return out;
  }
  // What a book holds that Excel cannot take, or that the page and Excel show differently, found from the book alone
  // (build.py tells of it, since Excel is not asked each time), each as the cell and what it is:
  //   limits: texts and formulas longer than Excel takes;
  //   unreadable: formulas the page cannot read (a name, another sheet, an unclosed bracket...), #NAME? or #VALUE! on
  //     the screen and a text in an xlsx file;
  //   uncomputable: formulas with a function the page does not have, a reference outside the rows and cols of the
  //     sheet (also where SUMIF or AVERAGEIF stretches its sum range), or a circular reference (#CIRC!): Excel may
  //     compute them, also when IFERROR hides the page's error;
  //   dependent: formulas that are #NAME? or #REF! only because a cell they read is the same error;
  //   errors: the other #NAME? and #REF! formulas (INDEX beyond its range, =#REF!), errors in Excel too;
  //   textFormulas: formulas in cells formatted as text, which become text once edited.
  function cautions(w) {
    const limits=[],unreadable=[],uncomputable=[],dependent=[],errors=[],textFormulas=[],count=n=>n.toLocaleString('en-US');
    const walk=(node,fn)=>{if(!node||typeof node!=='object')return;fn(node);for(const k of ['x','l','r'])walk(node[k],fn);if(node.args)for(const x of node.args)walk(x,fn);};
    const area=n=>range(n.b?n.a+':'+n.b:n.a);
    for(const s of w.sheets){
      const cache=new Map(),value=(r,c)=>calculate(w,s,r,c,new Set(),cache),read=new Map();
      // Whether a cell that the reference node reads (within the sheet) holds the error v; kept for each range.
      const readsError=(n,v)=>{const key=n.a+':'+(n.b||'')+v;if(read.has(key))return read.get(key);const q=area(n);let found=false;
        for(let r=q.r1;r<=Math.min(q.r2,s.rows-1)&&!found;r++)for(let c=q.c1;c<=Math.min(q.c2,s.cols-1)&&!found;c++)if(value(r,c)===v)found=true;
        read.set(key,found);return found;};
      const outsideOf=(r,c)=>r>=s.rows||c>=s.cols;
      for(const [a,cell] of Object.entries(s.cells)){
        const at=s.name+'!'+a;
        if(cell.t==='s'&&cell.v.length>32767) limits.push(at+' の文字（'+count(cell.v.length)+'文字。Excelのセルは32,767文字まで）');
        if(cell.t!=='f') continue;
        if(cell.f.length>8192) limits.push(at+' の数式（'+count(cell.f.length)+'文字。Excelの数式は8,192文字まで）');
        const literal=Math.max(0,...[...cell.f.matchAll(/"((?:[^"]|"")*)"/g)].map(m=>m[1].replace(/""/g,'"').length));
        if(literal>255) limits.push(at+' の数式の中の文字（'+count(literal)+'文字。Excelの数式の中の文字は255文字まで）');
        if(style(w,s,cell).numFmt==='@') textFormulas.push(at);
        let ast;try{ast=F.evalFormula(cell.f,null,true);}catch(e){unreadable.push(at+' '+cell.f);continue;}
        const unknown=new Set(),outside=new Set(),refs=[];
        walk(ast,n=>{if(n.op==='call'&&!FUNCTIONS.includes(n.name))unknown.add(n.name);
          if(n.op==='ref'){refs.push(n);for(const x of [n.a,n.b])if(x){const [rr,cc]=coord(x);if(outsideOf(rr,cc))outside.add(x.replace(/\$/g,''));}}
          // SUMIF and AVERAGEIF sum a range of the size of their criteria range, from the top-left of the one given.
          if(n.op==='call'&&(n.name==='SUMIF'||n.name==='AVERAGEIF')&&n.args[0]?.op==='ref'&&n.args[2]?.op==='ref'){
            const q=area(n.args[0]),p=area(n.args[2]),rr=p.r1+q.r2-q.r1,cc=p.c1+q.c2-q.c1;if(outsideOf(rr,cc))outside.add(addr(rr,cc));}});
        const [r,c]=coord(a),v=value(r,c),shown=typeof v==='string'&&!isErr(v)?JSON.stringify(v):String(v);
        if(unknown.size) uncomputable.push(at+' '+shown+'（'+[...unknown].join('・')+' はHTMLで計算できない関数）');
        else if(outside.size) uncomputable.push(at+' '+shown+'（'+[...outside].join('・')+' はシートの外。シートは'+s.rows+'行・'+s.cols+'列）');
        else if(v==='#CIRC!') uncomputable.push(at+' #CIRC!（循環参照）');
        else if(v==='#NAME?'||v==='#REF!') (refs.some(n=>readsError(n,v))?dependent:errors).push(at+' '+v);
      }
    }
    return {limits,unreadable,uncomputable,dependent,errors,textFormulas};
  }
  // Edges where two cells side by side draw different lines: Excel keeps one line per cell edge, so one of the two is
  // lost there (the page draws the line of the upper or left cell). A merged cell is one block with the lines of its
  // top-left cell. Colours are compared without case. q limits the cells looked at. a and b are the cells on the two
  // sides of the edge; cells names where their lines are written (a merged cell as its range).
  function borderConflicts(w,s,q={r1:0,c1:0,r2:s.rows-1,c2:s.cols-1}) {
    const merged=new Map();for(const m of s.merges) each(m,(r,c)=>merged.set(addr(r,c),m));
    const owner=(r,c)=>{const m=merged.get(addr(r,c));return m?addr(m.r1,m.c1):addr(r,c);};
    const label=a=>{const m=merged.get(a);return m?addr(m.r1,m.c1)+':'+addr(m.r2,m.c2):a;};
    const styles=new Map(),st=a=>{if(!styles.has(a)) styles.set(a,style(w,s,s.cells[a]||{t:'s',v:''}));return styles.get(a);};
    const out=[];
    each(q,(r,c)=>{for(const [dr,dc,mine,theirs,edge] of [[0,1,'borderRight','borderLeft','右/左'],[1,0,'borderBottom','borderTop','下/上']]){
      const rr=r+dr,cc=c+dc;if(rr>q.r2||cc>q.c2) continue;
      const a=owner(r,c),b=owner(rr,cc);if(a===b) continue;
      const x=st(a)[mine],y=st(b)[theirs];
      if(x!=='none'&&y!=='none'&&x.toLowerCase()!==y.toLowerCase()) out.push({a:addr(r,c),b:addr(rr,cc),edge,lines:[x,y],cells:[label(a),label(b)]});}});
    return out;
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
  // toLocaleString with options prepares a new number formatter on every call (about 30µs, the most of what a large
  // sheet spends on showing numbers); the formatter of the same options gives the same text, so each one is kept.
  const NUMBER_FORMATS=new Map();
  function numberFormat(decimals,grouping){const key=decimals+(grouping?',':'');let f=NUMBER_FORMATS.get(key);
    if(!f)NUMBER_FORMATS.set(key,f=new Intl.NumberFormat('en-US',{minimumFractionDigits:decimals,maximumFractionDigits:decimals,useGrouping:grouping}));return f;}
  function display(value,fmt='General') {
    if(isErr(value)||typeof value==='string')return value;if(typeof value==='boolean')return value?'TRUE':'FALSE';
    if(value==null)return ''; if(value===0)value=0; // Excel has no negative zero (=A1*-1 for an A1 of 0 shows 0)
    if(fmt==='yyyy/mm/dd'){try{return F.formatValue(value,'yyyy/mm/dd');}catch(e){return '#VALUE!';}}
    if(fmt==='General'||fmt==='@')return String(Number(value.toPrecision(15)));
    const percent=fmt.includes('%'),decimals=fmt.includes('.00')?2:0;
    const yen=fmt.startsWith('¥'),neg=yen&&value<0;
    const text=numberFormat(decimals,fmt.includes(',')).format((neg?-value:value)*(percent?100:1));
    return (neg?'-':'')+(yen?'¥':'')+text+(percent?'%':'');
  }
  // Horizontal alignment Excel shows for "general": numbers and errors right, booleans centre, text left.
  const generalAlign=value=>typeof value==='number'||isErr(value)?'right':typeof value==='boolean'?'center':'left';
  // Screen CSS. Indent is 1 character (about 9px at 11pt) per level on the aligned side, as in Excel.
  // Excel aligns a cell aligned "general" left once it is indented, whatever its value (xlCss writes it so too).
  function css(st,value) {
    const align=st.align==='general'?(st.indent>0?'left':generalAlign(value)):st.align,indent=st.indent*9;
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
  // What Excel reads a text as when it is typed or pasted without a text format: 'number' (with a sign, ¥, %, thousands
  // separators or an exponent), 'date' (a date or a time), or null.
  function textKind(text) {
    const s=String(text).trim();
    if(/^[(]?[-+]?[¥$€£]?\s*(?:\d[\d,]*(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?\s*%?[)]?$/i.test(s))return 'number';
    if(/^\d{1,4}[/\-.]\d{1,2}(?:[/\-.]\d{1,4})?(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/.test(s))return 'date';
    if(/^\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s*[AP]M?)?$/i.test(s))return 'date';
    if(/^(?:\d{1,4}年)?\d{1,2}月(?:\d{1,2}日)?$|^\d{1,2}日$|^(?:明治|大正|昭和|平成|令和|[MTSHR])\s*\d{1,2}[年./-]/.test(s))return 'date';
    if(/^\d{1,2}[\s-](?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*(?:[\s-]\d{2,4})?$/i.test(s)||/^(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s-]\d{1,4}$/i.test(s))return 'date';
    return null;
  }
  // Strings Excel would turn into a number, date, time, boolean, error or formula when it reads the text.
  function excelWouldConvert(text) {
    const s=String(text).trim();if(s!==String(text)&&s==='')return false;
    if(/^[=+\-@]/.test(s)&&s.length>1||/^'/.test(String(text)))return true;
    if(/^(?:TRUE|FALSE)$/i.test(s)||/^#(?:N\/A|NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|GETTING_DATA)$/i.test(s))return true;
    return textKind(s)!==null;
  }
  // x:fmla is read relative to the pasted block (even $-references), so rewrite references against the
  // copied range's top-left. Returns null when any reference points outside the copied range.
  function blockFormula(f,q) {
    let outside=false;
    const out=mapRefs(F.englishNames(f),(m,c$,col,r$)=>{const [r,c]=coord(m);if(!within(r,c,q)){outside=true;return m;}return c$+colName(c-q.c1)+r$+(r-q.r1+1);});
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
  // The cell references of a formula, outside its texts and sheet names, as formulas.js reads them: a reference stands
  // alone, not inside a name (LOG10( is a function, 1E5 a number, ABCD1 or Sheet2 a name, 'Q1 data' a sheet).
  // fn(m,$col,col,$row,row,sheet) gives what takes the place of the reference m; sheet is the name before it
  // (Sheet2!A1, 'Q1 data'!A1, with its quotes), which stays, and the end of a range has the sheet of its start
  // (Sheet2!A1:B2). Copying (shiftFormula), copying to Excel (blockFormula) and moving rows and columns (app.js) all
  // find references so. A name may hold letters other than ASCII (売上A1 is a name), as a sheet name may.
  const REFERENCE=/"(?:[^"]|"")*"|((?:'(?:[^']|'')+'|(?<![A-Z0-9_.$\u0080-￿])[A-Z_\u0080-￿][A-Z0-9_.\u0080-￿]*)!)?(?<![A-Z0-9_.$\u0080-￿])(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)(?![A-Z0-9_.(!\u0080-￿])|'(?:[^']|'')*'/gi;
  function mapRefs(f,fn){
    let end=-1,last;
    return f.replace(REFERENCE,(m,sheet,c$,col,r$,row,at)=>{
      if(col===undefined) return m;
      const name=sheet?sheet.slice(0,-1):at===end+1&&f[end]===':'?last:undefined;
      end=at+m.length;last=name;
      return (sheet||'')+fn(c$+col+r$+row,c$,col,r$,row,name);
    });
  }
  // The name of a sheet as a formula writes it before a reference: 'Q1 data' in quotes ('' for '), or Sheet2.
  const sheetName=t=>t.startsWith("'")?t.slice(1,-1).replace(/''/g,"'"):t;
  // A formula copied dr rows down and dc columns right, as Excel copies it: a reference that does not move stays as it
  // is written, and one moved above row 1, left of column A or beyond Excel's last row and column is #REF!. A reference
  // moved outside the rows and cols of the sheet stays a reference (the page computes it to #REF!, and build.py tells).
  function shiftFormula(f,dr,dc){return mapRefs(f,(m,c$,col,r$,row)=>{
    const moveRow=!r$&&dr!==0,moveCol=!c$&&dc!==0;if(!moveRow&&!moveCol) return c$+col.toUpperCase()+r$+row;
    const [r,c]=coord(m),nr=r+(moveRow?dr:0),nc=c+(moveCol?dc:0);
    return nr<0||nc<0||nr>=1048576||nc>=16384?'#REF!':c$+colName(nc)+r$+(nr+1);});}
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
    fontCharset,cssUnescape,parseDecls,parseRules,cssHex,formatFromMso,borderFromCss,styleFromDecls,valueFromDisplay,clone,esc,colName,addr,coord,rect,range,within,overlap,same,each,get,style,mergeAt,expanded,validate,expand,cautions,borderConflicts,width,height,infer,raw,isErr,calculate,evalFormula,display,css,mapRefs,sheetName,shiftFormula,rangeData,clipboard,markdown,parseTSV,paste};
});
