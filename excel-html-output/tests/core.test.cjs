// assets/core.js の関数：型、計算、表示、範囲操作、検査、Excel方言のクリップボードHTMLの書き出しと読み取り、Markdown、AIに渡す状態テキスト。
const assert=require('node:assert/strict');
const {test}=require('node:test');
const C=require('../assets/core.js');
const make=()=>({schema:C.SCHEMA,id:'test',revision:'draft',title:'test',activeSheet:'s',sheets:[{id:'s',name:'Sheet1',rows:30,cols:15,cells:{},merges:[],rowHeights:[],colWidths:[]}]});
test('typed values retain zeros, identifiers, quotes and Unicode',()=>{assert.equal(C.infer('00123').t,'s');assert.equal(C.infer('1234567890123456').t,'s');assert.equal(C.infer('-1.25').v,-1.25);assert.equal(C.infer("'=A1").v,'=A1');assert.equal(C.infer('001\n日本語').v,'001\n日本語');assert.equal(C.infer('12',{s:{numFmt:'@'}}).t,'s');});
test('formulas compute, reject code and handle IF lazily',()=>{const w=make(),s=w.sheets[0];s.cells={A1:{t:'n',v:2},A2:{t:'n',v:3},B1:{t:'f',f:'=SUM(A1:A2)*2'},B2:{t:'f',f:'=IF(A1=2,B1,1/0)'},C1:{t:'f',f:'=C1'},D1:{t:'f',f:'=fetch("https://x")'},D2:{t:'f',f:'=ROUND(-1.25,1)'},E1:{t:'f',f:'=A1/0'}};assert.equal(C.calculate(w,s,0,1),10);assert.equal(C.calculate(w,s,1,1),10);assert.equal(C.calculate(w,s,0,2),'#CIRC!');assert.equal(C.calculate(w,s,0,3),'#NAME?');assert.equal(C.calculate(w,s,1,3),-1.3);assert.equal(C.calculate(w,s,0,4),'#DIV/0!');});
test('relative and absolute formula copying preserves literals',()=>{assert.equal(C.shiftFormula('=A1+$B2+C$3+$D$4&"A1"',2,3),'=D3+$B4+F$3+$D$4&"A1"');assert.equal(C.shiftFormula('=A1',-1,0),'=#REF!');});
test('clipboard writes Excel dialect: exact text, full numbers, nowrap, pt borders, no selection UI',()=>{const w=make(),s=w.sheets[0];s.cells={A1:{t:'s',v:'見出し <script> & "',s:{bold:true,fill:'#123456',color:'#ffffff'}},A2:{t:'s',v:'00123',s:{numFmt:'@'}},B2:{t:'n',v:.125,s:{numFmt:'0.00%'}},C2:{t:'s',v:'日本語\n改行',s:{wrap:true,borderBottom:'3px double #ff0000'}},A3:{t:'s',v:'  00456  '},B3:{t:'s',v:'普通の文字'},C3:{t:'s',v:"'先頭"}};s.merges=[{r1:0,c1:0,r2:0,c2:2}];s.colWidths=[130,160,210];s.rowHeights=[36,44];C.validate(w);const clip=C.clipboard(w,s,{r1:0,c1:0,r2:2,c2:2}),h=clip.html;
  assert.ok(h.includes('xmlns:x="urn:schemas-microsoft-com:office:excel"'));assert.ok(h.includes('colspan="3"'));assert.ok(h.includes('<col width="130"'));assert.ok(h.includes('height:33pt'));
  assert.ok(!/x:str=""/.test(h),'empty x:str would blank the text in Excel');assert.ok(h.includes('x:str="見出し &lt;script&gt; &amp; &quot;"'));assert.ok(h.includes('x:num="0.125"'));assert.ok(h.includes('mso-number-format:Percent'));
  assert.ok(h.includes('x:str="日本語&#10;改行"'));assert.ok(h.includes('日本語<br style="mso-data-placement:same-cell">改行'));assert.ok(h.includes('white-space:normal'));assert.ok(h.includes('white-space:nowrap'));
  assert.ok(h.includes('border-bottom:2.0pt double #FF0000'));assert.ok(h.includes('background:#123456;mso-pattern:black none'));assert.ok(h.includes('font-family:&quot;Yu Gothic&quot;;mso-font-charset:128'));
  assert.ok(/x:str="  00456  "[^>]*mso-number-format:&quot;\\@&quot;/.test(h),'number-like text is protected as text');assert.ok(/x:str="普通の文字"[^>]*mso-number-format:General/.test(h));assert.ok(h.includes('x:str="&#39;&#39;先頭"'));
  assert.ok(h.includes('&lt;script&gt;'));assert.ok(!h.includes('<script>'));assert.ok(!h.includes('selected-cell'));assert.equal(clip.data.cells.A2.t,'s');assert.equal(clip.data.cells.C2.s.borderBottom,'3px double #ff0000');assert.deepEqual(C.parseTSV(clip.text),[['見出し <script> & "','',''],['00123','12.50%','日本語\n改行'],['  00456  ','普通の文字',"'先頭"]]);});
test('formulas go to Excel only when every reference stays inside the copied range',()=>{const w=make(),s=w.sheets[0];s.cells={B2:{t:'n',v:2},C2:{t:'n',v:3},D2:{t:'f',f:'=B2*$C$2+SUM(B2:C2)'},E2:{t:'f',f:'=A1+1'}};C.validate(w);assert.equal(C.blockFormula('=B2*$C$2+SUM(B2:C2)&"B2"',C.range('B2:E2')),'=A1*$B$1+SUM(A1:B1)&"B2"');assert.equal(C.blockFormula('=A1+1',C.range('B2:E2')),null);assert.equal(C.blockFormula('=LOG10(B2)',C.range('B2:E2')),'=LOG10(A1)');const h=C.clipboard(w,s,C.range('B2:E2')).html;assert.ok(h.includes('x:fmla="=A1*$B$1+SUM(A1:B1)"'));assert.equal((h.match(/x:fmla=/g)||[]).length,1);assert.ok(h.includes('x:num="11"'));});
test('text Excel would convert is detected, ordinary text is not',()=>{for(const v of ['00456','1234567890123456789','TRUE','false','1/2','2025-01-02','50%','-3','¥1,200','=A1','+81','12:30','3月4日','令和6年1月','#N/A',"'x",'1e5','(123)'])assert.equal(C.excelWouldConvert(v),true,v);for(const v of ['普通の文字','ID-A01','A-1','1行目\n2行目','abc 123','','  '])assert.equal(C.excelWouldConvert(v),false,v);});
test('Excel clipboard dialect is read back: formats, borders, colours, alignment, indent',()=>{assert.equal(C.formatFromMso('Standard'),'#,##0.00');assert.equal(C.formatFromMso('Percent'),'0.00%');assert.equal(C.formatFromMso('Fixed'),'0.00');assert.equal(C.formatFromMso('"\\0022¥\\0022\\#\\,\\#\\#0"'),'¥#,##0');assert.equal(C.formatFromMso('"yyyy\\/mm\\/dd"'),'yyyy/mm/dd');assert.equal(C.formatFromMso('"\\@"'),'@');assert.equal(C.formatFromMso('"\\#\\,\\#\\#0"'),'#,##0');assert.equal(C.formatFromMso('0.000E+00'),null);
  assert.equal(C.borderFromCss('.5pt solid windowtext'),'1px solid #000000');assert.equal(C.borderFromCss('1.0pt solid #14512E'),'2px solid #14512e');assert.equal(C.borderFromCss('.5pt hairline gray'),'0.5px solid #808080');assert.equal(C.borderFromCss('2.0pt double black'),'3px double #000000');assert.equal(C.borderFromCss('.5pt dashed #0070C0'),'1px dashed #0070c0');assert.equal(C.borderFromCss('none'),'none');
  const st=C.styleFromDecls(C.parseDecls('color:white;font-weight:700;font-family:"Yu Gothic", monospace;mso-font-charset:128;vertical-align:bottom;border-bottom:1.0pt solid #14512E;background:#217346;mso-pattern:black none;text-align:left;padding-left:30px;mso-char-indent-count:2;white-space:normal'));
  assert.deepEqual(st,{font:'Yu Gothic',bold:true,color:'#ffffff',fill:'#217346',align:'left',valign:'bottom',wrap:true,indent:2,borderBottom:'2px solid #14512e'});
  assert.deepEqual(C.parseRules('<!--td{color:black;font-size:11.0pt}\n.xl65{mso-number-format:"\\@";}-->'),[['td',{color:'black','font-size':'11.0pt'}],['.xl65',{'mso-number-format':'"\\@"'}]]);
  assert.equal(C.valueFromDisplay('¥1,200','¥#,##0'),1200);assert.equal(C.valueFromDisplay('-¥1,500','¥#,##0'),-1500);assert.equal(C.valueFromDisplay('12.35%','0.00%'),.1235);assert.equal(C.valueFromDisplay('2025/12/09','yyyy/mm/dd'),46000);assert.equal(C.valueFromDisplay('abc','0'),null);});
test('every style survives the trip out to Excel CSS and back',()=>{const styles=[{bold:true,italic:true,underline:true,color:'#c00000',fill:'#fce4d6',align:'center',valign:'top',wrap:true,numFmt:'#,##0.00',borderTop:'1px solid #000000',borderBottom:'3px double #000000',borderLeft:'0.5px solid #808080',borderRight:'2px dashed #0070c0'},{font:'Meiryo',size:9,strike:true,align:'right',indent:3,numFmt:'¥#,##0',borderBottom:'1px dotted #7030a0'},{numFmt:'yyyy/mm/dd',valign:'bottom'},{numFmt:'@'},{numFmt:'0%'},{numFmt:'0'},{numFmt:'0.00'},{numFmt:'0.00%'},{numFmt:'#,##0',borderTop:'3px solid #000000',borderBottom:'2px solid #000000',borderLeft:'1px dashed #000000'}];for(const s of styles){const full={...C.DEFAULT,...s},back={...C.DEFAULT,...C.styleFromDecls(C.parseDecls(C.xlCss(full,'x')))};assert.deepEqual(back,full,JSON.stringify(s));}});
test('defaults follow Japanese Excel and only Excel-representable borders validate',()=>{assert.equal(C.DEFAULT.valign,'middle');assert.equal(C.DEFAULT.fill,'none');const w=make(),s=w.sheets[0];s.cells.A1={t:'s',v:'x',s:{fill:'none',borderTop:'0.5px solid #808080',borderBottom:'3px double #000000'}};C.validate(w);for(const b of ['0.5px dotted #000000','3px dashed #000000','1px double #000000','1px groove #000000'])assert.throws(()=>{s.cells.A1.s={borderTop:b};C.validate(w);},/borderTop/);});
test('partial merged-cell selection expands and does not split the merge',()=>{const s=make().sheets[0];s.merges=[{r1:1,c1:1,r2:3,c2:3}];assert.deepEqual(C.expanded(s,C.rect([2,2])),{r1:1,c1:1,r2:3,c2:3});});
test('paste preserves empty cells and rejects clipping existing merges',()=>{const w=make(),s=w.sheets[0];s.cells.A1={t:'n',v:5};s.cells.A2={t:'f',f:'=A1*2'};const d=C.rangeData(w,s,C.range('A1:B2'));C.paste(s,d,5,3);assert.equal(s.cells.D7.f,'=D6*2');assert.equal(s.cells.E7.v,'');s.merges=[C.range('F10:G11')];assert.throws(()=>C.paste(s,d,9,4),/結合/);});
test('validate refuses overlapping merges, hidden values and invalid styles',()=>{const w=make(),s=w.sheets[0];s.merges=[C.range('A1:B2'),C.range('B2:C3')];assert.throws(()=>C.validate(w),/重複/);s.merges=[C.range('A1:B2')];s.cells.B2={t:'s',v:'do not discard'};assert.throws(()=>C.validate(w),/隠れた値/);s.merges=[];s.cells.B2.s={fill:'url(https://x)'};assert.throws(()=>C.validate(w),/fill/);});
test('quoted TSV preserves embedded tabs, quotes, newlines and trailing empty cells',()=>{assert.deepEqual(C.parseTSV('"a\tb"\t"c""d"\t\r\n"x\ny"\t0\t'),[['a\tb','c"d',''],['x\ny','0','']]);assert.throws(()=>C.parseTSV('"unterminated'),/引用符/);});
test('formats use values independently of strings',()=>{assert.equal(C.display(.125,'0.00%'),'12.50%');assert.equal(C.display(1234,'¥#,##0'),'¥1,234');assert.equal(C.display('00123','0'),'00123');assert.equal(C.display(45658,'yyyy/mm/dd'),'2025/01/01');});
test('view positions reject broken and out-of-bounds saved state',()=>{for(const view of [{zoom:100,active:7},{zoom:100,active:[-1,0]},{zoom:100,selection:{r1:0,c1:0,r2:500,c2:1}},{zoom:100,scrollTop:-1}]){const w=make();w.view=view;assert.throws(()=>C.validate(w),/view/);}});
test('string equality and inequality are complementary and case insensitive',()=>{assert.equal(C.evalFormula('="a"="A"',()=>0),true);assert.equal(C.evalFormula('="a"<>"A"',()=>0),false);});
test('rounding, integer, logical and IFERROR functions behave like Excel',()=>{const r=f=>C.evalFormula(f,()=>0);assert.equal(r('=ROUNDDOWN(76800.9,0)'),76800);assert.equal(r('=ROUNDDOWN(-1.25,1)'),-1.2);assert.equal(r('=ROUNDUP(1.21,1)'),1.3);assert.equal(r('=ROUNDDOWN(1234,-2)'),1200);assert.equal(r('=ROUNDDOWN(0.1*3*10,0)'),3);assert.equal(r('=INT(-1.5)'),-2);assert.equal(r('=ABS(-3)'),3);assert.equal(r('=MOD(-3,2)'),1);assert.throws(()=>r('=MOD(1,0)'),/DIV/);assert.equal(r('=IFERROR(MOD(1,0),"z")'),'z');assert.equal(r('=AND(1,TRUE)'),true);assert.equal(r('=OR(0,FALSE)'),false);assert.equal(r('=NOT(0)'),true);assert.equal(r('=IFERROR(1/0,"x")'),'x');assert.equal(r('=IFERROR(5,"x")'),5);const w=make(),s=w.sheets[0];s.cells={A1:{t:'n',v:768000},A2:{t:'f',f:'=ROUNDDOWN(A1*0.1,0)'},A3:{t:'f',f:'=IFERROR(A1/0,0)'}};assert.equal(C.calculate(w,s,1,0),76800);assert.equal(C.calculate(w,s,2,0),0);});
test('markdown table keeps what is shown, escapes pipes and line breaks, aligns numbers right',()=>{const w=make(),s=w.sheets[0];
  s.cells={A1:{t:'s',v:'品名'},B1:{t:'s',v:'金額'},C1:{t:'s',v:'備考'},A2:{t:'s',v:'ノートPC'},B2:{t:'n',v:15000,s:{numFmt:'¥#,##0'}},C2:{t:'s',v:'a|b\nc'},A3:{t:'s',v:'マウス'},B3:{t:'f',f:'=B2*2',s:{numFmt:'¥#,##0'}}};
  const md=C.markdown(w,s,C.range('A1:D5'));assert.deepEqual([md.rows,md.cols,md.empty],[3,3,false]);
  assert.equal(md.text,['| 品名     |    金額 | 備考      |','| -------- | ------: | --------- |','| ノートPC | ¥15,000 | a\\|b<br>c |','| マウス   | ¥30,000 |           |'].join('\n'));
  assert.equal(C.markdown(w,s,C.range('A1:O30')).text,md.text,'trailing empty rows and columns are left out');
  s.cells={A1:{t:'s',v:'タイトル'},A2:{t:'n',v:1},B2:{t:'n',v:2}};s.merges=[C.range('A1:B1')];
  assert.equal(C.markdown(w,s,C.range('A1:B2')).text,['| タイトル |     |','| -------: | --: |','|        1 |   2 |'].join('\n'));
  s.merges=[];s.cells={A1:{t:'s',v:'サイト',link:'https://example.com/a b(1)'}};
  assert.ok(C.markdown(w,s,C.range('A1')).text.includes('[サイト](https://example.com/a%20b%281%29)'));
  assert.equal(C.markdown(w,s,C.range('E10:F12')).empty,true);});
test('pictures need real image data and stay inside the sheet',()=>{const w=make(),s=w.sheets[0],png='data:image/png;base64,iVBORw0KGgo=';
  s.images=[{id:'p1',asset:'a1',r:2,c:3,x:5,y:6,w:120,h:80}];
  assert.throws(()=>C.validate(w),/画像の中身/,'a picture without its data is refused');
  assert.doesNotThrow(()=>C.validate(w,{assetsDetached:true}),'the editor keeps the data apart while editing');
  w.assets={a1:png};C.validate(w);
  for(const bad of ['data:image/svg+xml;base64,PHN2Zz4=','data:text/html;base64,PGI+','data:image/png;base64,<script>','javascript:alert(1)']){w.assets={a1:bad};assert.throws(()=>C.validate(w),/画像データ/,bad);}
  w.assets={a1:png};for(const patch of [{r:30},{c:15},{w:2},{x:-1},{id:''}]){s.images=[{id:'p1',asset:'a1',r:2,c:3,x:5,y:6,w:120,h:80,...patch}];assert.throws(()=>C.validate(w),/画像/,JSON.stringify(patch));}});
test('column widths and row heights stay within what Excel can set: 24-2044px wide, 16-546px high',()=>{const w=make(),s=w.sheets[0];s.colWidths=[24,2044];s.rowHeights=[16,546];assert.doesNotThrow(()=>C.validate(w),'the narrowest and widest column, the lowest and highest row');
  for(const [field,size] of [['colWidths',23],['colWidths',2045],['rowHeights',15],['rowHeights',547]]){const v=make();v.sheets[0][field]=[size];assert.throws(()=>C.validate(v),new RegExp(field),field+' '+size);}});
const T=require('../assets/toon.js');
const handoff=()=>{const w=make(),s=w.sheets[0];s.cells={A1:{t:'s',v:'見積明細',s:{bold:true,fill:'#e2efda'}},A2:{t:'s',v:'00123',s:{numFmt:'@'}},B2:{t:'f',f:'=SUM(C2:D2)'},C2:{t:'n',v:3},D2:{t:'s',v:'a, "b"\n2行目'}};
  s.merges=[{r1:5,c1:0,r2:5,c2:1},{r1:6,c1:0,r2:6,c2:1}];s.colWidths=[96,null,120];w.view={zoom:100,active:[1,2],selection:{r1:1,c1:2,r2:1,c2:2},scrollTop:0,scrollLeft:0};
  return {schema:'excel-html-state/v1',workbookId:'test',baseRevision:'sha256:'+'0'.repeat(64),exportedAt:'2026-11-17T00:00:00.000Z',workbook:w};};
test('the state text for the AI is TOON by default, between its first line and an end line, and reads back to the same state',()=>{const state=handoff(),text=C.stateText(state),lines=text.split('\n');
  assert.equal(lines[0],'EXCEL_HTML_STATE_V1 TOON');assert.equal(lines.at(-1),'EXCEL_HTML_STATE_END');assert.equal(lines.slice(1,-1).join('\n'),T.encode(state),'the lines between are the TOON of the state');
  assert.deepEqual(C.readStateText(text),state);assert.deepEqual(C.readStateText('\ufeff'+text.replace(/\n/g,'\r\n')+'\r\n'),state,'a BOM, CRLF line ends and a final line end are accepted');});
test('the JSON state text is the same as before and is still read',()=>{const state=handoff(),json=C.stateText(state,'json');
  assert.equal(json,'EXCEL_HTML_STATE_V1\n'+JSON.stringify(state,null,2));assert.deepEqual(C.readStateText(json),state);assert.ok(C.stateText(state).length<json.length,'TOON is the shorter text');});
test('a state text cut short, changed inside, or without its first line is refused rather than read in part',()=>{const state=handoff(),text=C.stateText(state),lines=text.split('\n');
  for(let n=1;n<lines.length;n++)assert.throws(()=>C.readStateText(lines.slice(0,n).join('\n')),/EXCEL_HTML_STATE_END/,'the first '+n+' lines');
  assert.throws(()=>C.readStateText(text.slice(0,-3)),/EXCEL_HTML_STATE_END/,'cut inside the end line');
  const row=lines.findIndex(l=>/^ +5,0,5,1$/.test(l));assert.ok(row>0);assert.throws(()=>C.readStateText(lines.filter((_,i)=>i!==row).join('\n')),/TOON/,'a merge row taken out');
  assert.throws(()=>C.readStateText(lines.slice(1).join('\n')),/EXCEL_HTML_STATE_V1で始まる/);assert.throws(()=>C.readStateText(text.replace('EXCEL_HTML_STATE_V1 TOON','EXCEL_HTML_STATE_V2 TOON')),/EXCEL_HTML_STATE_V1で始まる/);
  const json=C.stateText(state,'json');assert.throws(()=>C.readStateText(json.slice(0,json.length/2)),/JSON/);});
test('a selection of whole rows or whole columns is not widened by merged cells, as in Excel; any other range takes in every merged cell it touches',()=>{const s=make().sheets[0];s.merges=[C.range('A11:C12')];
  assert.deepEqual(C.selectionArea(s,C.range('B1:B30')),C.range('B1:B30'),'whole column B');assert.deepEqual(C.selectionArea(s,C.range('A11:O11')),C.range('A11:O11'),'whole row 11');
  assert.deepEqual(C.selectionArea(s,C.range('B10:B11')),C.range('A10:C12'),'a range of cells');
  assert.equal(C.splitsMerge(s,C.range('B1:B30')),true,'column B holds only part of A11:C12');assert.equal(C.splitsMerge(s,C.range('A10:C12')),false);assert.equal(C.splitsMerge(s,C.range('D1:D30')),false);});
test('copying a range that holds only part of a merged cell takes just that range as separate cells, as Excel does',()=>{const w=make(),s=w.sheets[0];
  s.cells={A1:{t:'s',v:'題名'},B4:{t:'s',v:'品名'},A11:{t:'s',v:'結合セル'}};s.merges=[C.range('A1:B1'),C.range('A11:C12')];C.validate(w);
  const d=C.rangeData(w,s,C.range('B1:B30'));assert.deepEqual([d.rows,d.cols,d.merges],[30,1,[]],'column B alone, with no merge');assert.equal(d.cells.A4.v,'品名');assert.equal(d.cells.A11.v,'','the part of A11:C12 is an empty cell');
  const clip=C.clipboard(w,s,C.range('B1:B30'));assert.equal(clip.text.split('\r\n').length,30);assert.equal((clip.html.match(/<td[ >]/g)||[]).length,30,'every row has its cell, the hidden parts of merged cells included');assert.ok(!/colspan|rowspan/.test(clip.html),'no merge goes to Excel');assert.ok(!clip.html.includes('結合セル')&&!clip.html.includes('題名'),'values of merged cells outside the range stay out');
  const md=C.markdown(w,s,C.range('B1:B4'));assert.ok(md.text.includes('品名')&&!md.text.includes('題名'));
  const left=C.rangeData(w,s,C.range('A1:A2'));assert.deepEqual([left.cells.A1.v,left.merges],['題名',[]],'a merged cell whose top-left is in the range gives its value, unmerged');
  assert.deepEqual(C.rangeData(w,s,C.range('A10:C12')).merges,[{r1:1,c1:0,r2:2,c2:2}],'a merged cell wholly inside the range stays merged');});
test('formula syntax can be checked without evaluating it (an xlsx must not hold a formula Excel cannot read)',()=>{for(const f of ['=SUM(A1:A3)*2','=vlookup(A1,B1:C3,2,FALSE)','=IF(A1>0,"正",-A1)','=A1%&"円"','=--A1','=TRUE'])assert.doesNotThrow(()=>C.evalFormula(f,null,true),f);for(const f of ['=SUM(A1:A3','=1+','=A1 B1','="abc','=Sheet2!A1','=(1','=1)','='])assert.throws(()=>C.evalFormula(f,null,true),f);});
// The lists of what a book can hold (FIELDS, CHOICES, CELL_TYPES, FUNCTIONS) are the ones validate and the formulas use,
// so a new feature has to be added to them, and tests/xlsx-features.test.mjs then asks for its xlsx samples.
const pictureBook=()=>{const w=make(),s=w.sheets[0];w.assets={a:'data:image/png;base64,iVBORw0KGgo='};w.view={zoom:100,selection:{r1:0,c1:0,r2:0,c2:0}};
  Object.assign(s,{freeze:{rows:1,cols:0},merges:[{r1:2,c1:0,r2:2,c2:1}],cells:{A1:{t:'s',v:'x'}},images:[{id:'p',asset:'a',r:3,c:0,x:0,y:0,w:10,h:10}]});return w;};
const escape=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
test('validate refuses fields the book does not have, at every level, and says where',()=>{
  const places=[[w=>w,'extra'],[w=>w.view,'view.extra'],[w=>w.view.selection,'view.selection.extra'],[w=>w.sheets[0],'sheets[0].extra'],[w=>w.sheets[0].freeze,'sheets[0].freeze.extra'],
    [w=>w.sheets[0].merges[0],'sheets[0].merges[0].extra'],[w=>w.sheets[0].cells.A1,'sheets[0].cells.A1.extra'],[w=>w.sheets[0].images[0],'sheets[0].images[0].extra']];
  for(const [at,place] of places){const w=pictureBook();C.validate(w);at(w).extra=1;assert.throws(()=>C.validate(w),new RegExp('未対応の項目 '+escape(place)+'$'),place);}});
test('validate accepts a book that uses every listed field',()=>{const w=pictureBook(),s=w.sheets[0];
  Object.assign(w,{defaultStyle:{size:10},view:{zoom:110,active:[1,1],selection:{r1:1,c1:1,r2:2,c2:2},scrollTop:0,scrollLeft:0}});
  Object.assign(s,{gridlines:false,defaultStyle:{fill:'#fff2cc'},cells:{A1:{t:'s',v:'x',s:{bold:true},link:'https://example.com/'},A2:{t:'n',v:1},A3:{t:'b',v:true},B1:{t:'f',f:'=A2'}}});
  const used={book:w,sheet:s,cell:{...s.cells.A1,...s.cells.B1},view:w.view,freeze:s.freeze,range:s.merges[0],image:s.images[0]};
  assert.deepEqual(Object.keys(C.FIELDS).sort(),Object.keys(used).sort());
  for(const [kind,list] of Object.entries(C.FIELDS))assert.deepEqual(Object.keys(used[kind]).sort(),[...list].sort(),kind+': the book uses every field');
  assert.doesNotThrow(()=>C.validate(w));
  without(C.FIELDS.sheet,'gridlines',()=>assert.throws(()=>C.validate(w),/未対応の項目 sheets\[0\]\.gridlines$/,'a field taken out of FIELDS'));});
test('gridlines is true or false',()=>{const w=make();for(const v of [true,false]){w.sheets[0].gridlines=v;assert.doesNotThrow(()=>C.validate(w),String(v));}
  for(const v of ['no',0,null]){w.sheets[0].gridlines=v;assert.throws(()=>C.validate(w),/gridlines/,String(v));}});
// Takes value out of list while fn runs, to show that core.js reads the exported list itself.
const without=(list,value,fn)=>{const i=list.indexOf(value);assert.ok(i>=0,value);list.splice(i,1);try{fn();}finally{list.splice(i,0,value);}};
const result=f=>{try{return C.evalFormula(f,()=>0);}catch(e){return e.message;}};
test('the formulas compute the functions in FUNCTIONS, and no other',()=>{
  for(const name of C.FUNCTIONS)assert.notEqual(result('='+name+'()'),'#NAME?',name+' is computed (the question read_xlsx.cjs asks)');
  without(C.FUNCTIONS,'SUM',()=>assert.equal(result('=SUM(1,2)'),'#NAME?','a function taken out of FUNCTIONS'));
  without(C.FUNCTIONS,'IF',()=>assert.equal(result('=IF(TRUE,1,2)'),'#NAME?','IF too, though its arguments are computed lazily'));
  assert.equal(result('=VLOOKUP(1,A1:B2,2)'),'#NAME?','a function the page does not compute');
  assert.equal(result('=VLOOKUP(1/0,A1:B2,2)'),'#DIV/0!','an error in an argument still comes first, as Excel shows it for VLOOKUP (desktop Excel 16.0)');});
test('a name Excel does not know is #NAME? even when an argument is an error',{todo:'the page does not know which names Excel knows: =FOO(1/0) shows #DIV/0!, while desktop Excel 16.0 shows #NAME?'},()=>{
  assert.equal(result('=FOO(1/0)'),'#NAME?');});
test('a number aligned general with an indent is shown aligned left, as Excel aligns a general cell it indents',{todo:'the screen (css) aligns the number right and indents it from the right, while the clipboard (xlCss) and the xlsx write it aligned left; only a book JSON holds this, since the ribbon has no indent button'},()=>{
  assert.match(C.css({...C.DEFAULT,indent:2},5),/text-align:left/);});
test('validate takes only the values in CHOICES and CELL_TYPES',()=>{const w=make(),s=w.sheets[0];
  for(const align of C.CHOICES.align){s.cells.A1={t:'s',v:'x',s:{align}};assert.doesNotThrow(()=>C.validate(w),align);}
  for(const valign of C.CHOICES.valign){s.cells.A1={t:'s',v:'x',s:{valign}};assert.doesNotThrow(()=>C.validate(w),valign);}
  s.cells.A1={t:'s',v:'x',s:{align:'justify'}};assert.throws(()=>C.validate(w),/align/);s.cells.A1={t:'s',v:'x',s:{valign:'center'}};assert.throws(()=>C.validate(w),/valign/);
  s.cells.A1={t:'s',v:'x',s:{align:'center'}};without(C.CHOICES.align,'center',()=>assert.throws(()=>C.validate(w),/align/,'a value taken out of CHOICES'));
  for(const t of C.CELL_TYPES){s.cells.A1=t==='f'?{t,f:'=1'}:{t,v:t==='s'?'x':t==='n'?1:true};assert.doesNotThrow(()=>C.validate(w),t);}
  s.cells.A1={t:'b',v:true};without(C.CELL_TYPES,'b',()=>assert.throws(()=>C.validate(w),/型 A1/,'a type taken out of CELL_TYPES'));
  s.cells.A1={t:'e',v:'#N/A'};assert.throws(()=>C.validate(w),/型 A1/);});
