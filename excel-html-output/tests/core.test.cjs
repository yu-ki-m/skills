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
  for(const [name,f] of [['CONCATENATE','=CONCATENATE("a")'],['CONCAT','=CONCAT("a")'],['TEXTJOIN','=TEXTJOIN(",",TRUE,"a")']])
    without(C.FUNCTIONS,name,()=>assert.equal(result(f),'#NAME?',name+' too, though it reads its arguments itself'));
  assert.equal(result('=SUMSQ(1,2)'),'#NAME?','a function the page does not compute');
  assert.equal(result('=SUMSQ(1/0)'),'#NAME?','whatever its arguments are (Excel, which knows SUMSQ, shows #DIV/0!)');});
test('a name Excel does not know is #NAME? even when an argument is an error',()=>{
  assert.equal(result('=FOO(1/0)'),'#NAME?');});
// Desktop Excel 16.0 keeps 15 significant digits of a number written in a formula and cuts the rest off: it saves
// =INT(2.9999999999999996) as INT(2.99999999999999), whose value is 2 (tests/fixtures/excel-functions.xlsx, 数学!H117).
test('a number written in a formula is read to 15 significant digits, the rest cut off, as Excel reads it',()=>{
  for(const [f,want] of [['=INT(2.9999999999999996)',2],['=TRUNC(-2.0000000000000004)',-2],['=HOUR(0.5213541666666667)',12],['=123456789012345678',123456789012345000],
    ['=0.30000000000000004=0.3',true],['=1.0000000000000002',1],['=.5',0.5],['=1E+20',1e20],['=1.234567890123456789E-5',0.0000123456789012345]])
    assert.equal(result(f),want,f);});
// Excel aligns a cell aligned "general" left once it is indented, whatever its value; the clipboard (xlCss) and the
// xlsx write such a cell aligned left, and the screen must show it so too (a book JSON can hold it: the ribbon has no
// indent button, and an xlsx read in brings it as aligned left).
test('a cell aligned general with an indent is shown aligned left and indented from the left, whatever its value, as the clipboard writes it',()=>{
  for(const value of [5,-0.5,true,'#N/A','文字']){const st={...C.DEFAULT,indent:2},shown=C.css(st,value);
    assert.match(shown,/text-align:left;/,JSON.stringify(value));
    assert.match(shown,/padding-left:21px;padding-right:3px/,JSON.stringify(value)+': two levels of 9px, on the left');
    assert.match(C.xlCss(st,value),/text-align:left;/,JSON.stringify(value)+': the clipboard writes it the same');}
  assert.match(C.css(C.DEFAULT,5),/text-align:right;.*padding-left:3px;padding-right:3px/,'a number not indented stays aligned right');
  assert.match(C.css(C.DEFAULT,true),/text-align:center;/,'and TRUE centred');
  assert.match(C.css({...C.DEFAULT,align:'right',indent:1},5),/text-align:right;.*padding-left:3px;padding-right:12px/,'aligned right, it is indented from the right');});
test('validate takes only the values in CHOICES and CELL_TYPES',()=>{const w=make(),s=w.sheets[0];
  for(const align of C.CHOICES.align){s.cells.A1={t:'s',v:'x',s:{align}};assert.doesNotThrow(()=>C.validate(w),align);}
  for(const valign of C.CHOICES.valign){s.cells.A1={t:'s',v:'x',s:{valign}};assert.doesNotThrow(()=>C.validate(w),valign);}
  s.cells.A1={t:'s',v:'x',s:{align:'justify'}};assert.throws(()=>C.validate(w),/align/);s.cells.A1={t:'s',v:'x',s:{valign:'center'}};assert.throws(()=>C.validate(w),/valign/);
  s.cells.A1={t:'s',v:'x',s:{align:'center'}};without(C.CHOICES.align,'center',()=>assert.throws(()=>C.validate(w),/align/,'a value taken out of CHOICES'));
  for(const t of C.CELL_TYPES){s.cells.A1=t==='f'?{t,f:'=1'}:{t,v:t==='s'?'x':t==='n'?1:true};assert.doesNotThrow(()=>C.validate(w),t);}
  s.cells.A1={t:'b',v:true};without(C.CELL_TYPES,'b',()=>assert.throws(()=>C.validate(w),/型 A1/,'a type taken out of CELL_TYPES'));
  s.cells.A1={t:'e',v:'#N/A'};assert.throws(()=>C.validate(w),/型 A1/);});

// Joining text: & and CONCATENATE, CONCAT, TEXTJOIN. Every expected value is what desktop Excel 16.0 (build 20326)
// gave for the same formula on the same cells (set with .Formula through COM, read back with .Value2); a value of the
// page's own is said so. tests/xlsx-import.test.mjs checks the same against a file Excel saved (fixtures/excel-text.xlsx).
const joined=x=>C.evalFormula('=A1&""',()=>x);
test('a number joined as text is written as Excel writes it: 15 significant digits, in full up to 20 characters (the sign aside), else as 1.5E-18',()=>{
  for(const [x,want] of [[1/3,'0.333333333333333'],[2/3,'0.666666666666667'],[1/7,'0.142857142857143'],[0.1+0.2,'0.3'],[46023,'46023'],[0.125,'0.125'],[-1.5,'-1.5'],
    [0,'0'],[-0,'0'],[0.9999999999999999,'1'],[9.999999999999998,'10'],[1234567890123456,'1234567890123460'],[123456789012345.6,'123456789012346'],
    [9.999999999999989e19,'99999999999999900000'],[-9.99999999999999e19,'-99999999999999900000'],[1e20,'1E+20'],[1.5e20,'1.5E+20'],[-1e20,'-1E+20'],
    [1.23456789012345e21,'1.23456789012345E+21'],[0.000123456789012345,'0.000123456789012345'],[-0.000123456789012345,'-0.000123456789012345'],
    [0.0000123456789012345,'1.23456789012345E-05'],[0.000012345678901234,'0.000012345678901234'],[1e-5,'0.00001'],[1e-12,'0.000000000001'],
    [1.5e-12,'0.0000000000015'],[1.23456789012345e-12,'1.23456789012345E-12'],[1e-18,'0.000000000000000001'],[-1e-18,'-0.000000000000000001'],
    [1.5e-18,'1.5E-18'],[1e-19,'1E-19'],[1e-307,'1E-307'],[2.2250738585072014e-308,'2.2250738585072E-308'],[1.7976931348623157e308,'1.7976931348623E+308']])
    assert.equal(joined(x),want,String(x));});
test('a number exactly halfway at the 16th significant digit is rounded toward zero, as Excel does, and one past halfway is rounded up',()=>{
  for(const [x,want] of [[100000000000000.5,'100000000000000'],[100000000000001.5,'100000000000001'],[-100000000000001.5,'-100000000000001'],
    [1000000000000.375,'1000000000000.37'],[12345678901234.75,'12345678901234.7'],[123456789012.1875,'123456789012.187'],[1.430511474609375e-6,'1.43051147460937E-06'],
    // Their exact values lie just past and just short of halfway: 1234567890.0000050067..., 12345678901.0000495910...
    [1234567890.000005,'1234567890.00001'],[12345678901.00005,'12345678901'],[1.234567890123456,'1.23456789012346'],[7.777777777777778,'7.77777777777778']])
    assert.equal(joined(x),want,String(x));});
// The sheet the formulas were measured on: A1 a, A2 empty, A3 1, A4 TRUE, A5 ="", A6 #N/A, A7 12.5% (0.125), A8 2026/01/01 (46023),
// A9 " sp ", A10 x, B1 b, B2 c, B4 2.5, C1 -, C2 +, D1 the text TRUE, D2 abc, D3 0, D4 2. An error read from Excel is held as its text.
const SHEET={A1:{t:'s',v:'a'},A3:{t:'n',v:1},A4:{t:'b',v:true},A5:{t:'f',f:'=""'},A6:{t:'s',v:'#N/A'},A7:{t:'n',v:0.125,s:{numFmt:'0%'}},
  A8:{t:'n',v:46023,s:{numFmt:'yyyy/mm/dd'}},A9:{t:'s',v:' sp '},A10:{t:'s',v:'x'},B1:{t:'s',v:'b'},B2:{t:'s',v:'c'},B4:{t:'n',v:2.5},
  C1:{t:'s',v:'-'},C2:{t:'s',v:'+'},D1:{t:'s',v:'TRUE',s:{numFmt:'@'}},D2:{t:'s',v:'abc'},D3:{t:'n',v:0},D4:{t:'n',v:2}};
/** The value of formula f in cell where (H40 unless said) of a sheet holding cells. */
const on=(f,where='H40',cells=SHEET)=>{const w=make(),s=w.sheets[0];s.rows=60;s.cells={...cells,[where]:{t:'f',f}};const [r,c]=C.coord(where);return C.calculate(w,s,r,c);};
const check=(cases,cells)=>{for(const [f,want] of cases)assert.equal(on(f,'H40',cells),want,f);};
test('& joins TRUE and FALSE, numbers and empty cells as Excel writes them',()=>check([['="a"&"b"','ab'],['=A1&B1','ab'],['=A2&"x"','x'],['=A3&"x"','1x'],
  ['=A4&"x"','TRUEx'],['=TRUE&""','TRUE'],['=FALSE&1','FALSE1'],['=(1=1)&""','TRUE'],['=1&2','12'],['="1"&2+1','13'],['="a"&-1','a-1'],['=50%&""','0.5'],
  ['=0.1+0.2&""','0.3'],['=A5&"x"','x'],['=A7&""','0.125'],['=A8&""','46023'],['=A9&"|"',' sp |'],['=SUM(1,2)&"件"','3件']]));
test('a text that reads as a number is a number in arithmetic, as in Excel (="1"&"2"+1 is 13)',()=>check([['="1"&"2"+1','13'],['="2"*"3"',6],['=" 1,000 "+1',1001],
  ['="abc"+1','#VALUE!'],['=""+1','#VALUE!'],['="2026/1/2"+0',46024],['="12:00"*2',1]]));
test('every operator shows the error of its left operand first, as Excel does (=A6+1/0 is #N/A)',()=>check([['=A6&"x"','#N/A'],['=1/0&A6','#DIV/0!'],
  ['=A6&1/0','#N/A'],['=A6+1/0','#N/A'],['=1/0+A6','#DIV/0!'],['=A6*1/0','#N/A'],['=A6-1/0','#N/A'],['=A6^(1/0)','#N/A'],['=A6=1/0','#N/A'],['=A6<1/0','#N/A']]));
test('a range used as one value, with & or in CONCATENATE, gives its cell in the row or column of the formula (implicit intersection), as Excel does',()=>{
  for(const [where,f,want] of [['F1','=A1:A3&"x"','ax'],['F2','=A1:A3&"x"','x'],['F3','=A1:A3&"x"','1x'],['F20','=A1:A3&"x"','#VALUE!'],
    ['G1','=CONCATENATE(A1:A3)','a'],['G2','=CONCATENATE(A1:A3,"!")','!'],['G3','=CONCATENATE(A1:A3)','1'],['G21','=CONCATENATE(A1:A3)','#VALUE!'],
    ['A22','=CONCATENATE(A1:B1)','a'],['B23','=CONCATENATE(A1:B1)','b'],['B24','="<"&A1:B1','<b'],['A26','=A1:A1&"x"','ax'],['B26','=CONCATENATE(B2:B2)','c'],
    ['E2','=A1:B2&"x"','#VALUE!'],['A40','=A1:B2&"x"','#VALUE!'],['E1','=CONCATENATE(A1:B2)','#VALUE!']])
    assert.equal(on(f,where),want,where+' '+f);
  const range=()=>Object.assign([1,2,3],{range:{r1:0,c1:0,r2:2,c2:0}});
  assert.throws(()=>C.evalFormula('=CONCATENATE(A1:A3)',range),/#VALUE!/,'without the cell of the formula, a range is no one value');});
test('CONCATENATE joins its arguments in order, each one value, and shows the first error',()=>check([['=CONCATENATE("a","b",1,TRUE)','ab1TRUE'],
  ['=CONCATENATE(A1)','a'],['=CONCATENATE(A2)',''],['=CONCATENATE(1/3)','0.333333333333333'],['=CONCATENATE(A7,A8)','0.12546023'],['=CONCATENATE(A6,"x")','#N/A'],
  ['=CONCATENATE("x",1/0,A6)','#DIV/0!'],
  ['=CONCATENATE()','#VALUE!']])); // Excel does not take this formula; the page shows #VALUE!, as for any wrong number of arguments
test('CONCAT joins the cells of its ranges row by row, and its arguments in order',()=>check([['=CONCAT("a","b")','ab'],['=CONCAT(A1:B2)','abc'],
  ['=CONCAT(A1:A5)','a1TRUE'],['=CONCAT(A7,A8)','0.12546023'],['=CONCAT(1/3)','0.333333333333333'],['=CONCAT(A1:B4)','abc1TRUE2.5'],['=CONCAT(B1:B4,A1:A4)','bc2.5a1TRUE'],
  ['=CONCAT(A2)',''],['=CONCAT(A5)',''],['=CONCAT(A2:B3)','c1'],['=CONCAT(A1:A10)','#N/A'],['=CONCAT(1/0,A6)','#DIV/0!'],['=IFERROR(CONCAT(A6),"err")','err'],
  ['=IF(TRUE,CONCAT("a",1),0)','a1'],['=CONCAT("a")&TEXTJOIN("-",TRUE,"b","c")','ab-c'],['=CONCAT(SUM(A3:A4))','1'],
  ['=CONCAT()','#VALUE!']])); // the page's, as for CONCATENATE()
test('TEXTJOIN puts the delimiter between the texts, takes the delimiters of a range in turn, and leaves out empty texts when ignore_empty is TRUE',()=>check([
  ['=TEXTJOIN(",",TRUE,A1:A5)','a,1,TRUE'],['=TEXTJOIN(",",FALSE,A1:A5)','a,,1,TRUE,'],['=TEXTJOIN(C1:C2,TRUE,"a","b","c","d")','a-b+c-d'],
  ['=TEXTJOIN(C1:C3,FALSE,"a","b","c","d","e")','a-b+cd-e'],['=TEXTJOIN(B1:C2,FALSE,"1","2","3","4","5","6")','1b2-3c4+5b6'],
  // The next delimiter is taken only when a text is put in: an empty text left out does not use one up.
  ['=TEXTJOIN(C1:C2,TRUE,"a","","b","c")','a-b+c'],['=TEXTJOIN(C1:C2,TRUE,"","a","b","c")','a-b+c'],['=TEXTJOIN(C1:C2,FALSE,"a","","b","c")','a-+b-c'],
  ['=TEXTJOIN(1,TRUE,"a","b")','a1b'],['=TEXTJOIN(TRUE,TRUE,"a","b")','aTRUEb'],['=TEXTJOIN(0.5,TRUE,"a","b")','a0.5b'],['=TEXTJOIN("",TRUE,"a","b")','ab'],
  ['=TEXTJOIN(A2,TRUE,"a","b")','ab'],['=TEXTJOIN(A2:A3,FALSE,"a","b","c")','ab1c'],['=TEXTJOIN(A1,TRUE,"x","y")','xay'],['=TEXTJOIN(",",TRUE,"")',''],
  ['=TEXTJOIN(",",TRUE,A2:A2)',''],['=TEXTJOIN(",",TRUE,A1:B2)','a,b,c'],['=TEXTJOIN(",",FALSE,A1:B2)','a,b,,c'],['=TEXTJOIN(", ",TRUE,A9,A10)',' sp , x'],
  ['=TEXTJOIN(",",TRUE,A1:A2,"z",A3)','a,z,1'],['=TEXTJOIN(",",TRUE,1/3,2/3)','0.333333333333333,0.666666666666667'],['=TEXTJOIN(",",FALSE,A2,A3)',',1'],
  ['=TEXTJOIN(",",TRUE,A4,FALSE)','TRUE,FALSE'],['=TEXTJOIN(",",TRUE,A1:B4)','a,b,c,1,TRUE,2.5'],['=TEXTJOIN(",",FALSE,A1:B4)','a,b,,c,1,,TRUE,2.5'],
  ['=TEXTJOIN(",",TRUE,"a")','a'],
  ['=TEXTJOIN(",",TRUE)','#VALUE!']])); // the page's, as for CONCATENATE()
test('TEXTJOIN reads ignore_empty as Excel does: a number, TRUE or FALSE also as text, an empty cell as FALSE, and anything else as #VALUE!',()=>{check([
  ['=TEXTJOIN(",",1,"a","","b")','a,b'],['=TEXTJOIN(",",0,"a","","b")','a,,b'],['=TEXTJOIN(",",2.5,"a","","b")','a,b'],['=TEXTJOIN(",","TRUE","a","","b")','a,b'],
  ['=TEXTJOIN(",","true","a","","b")','a,b'],['=TEXTJOIN(",","FALSE","a","","b")','a,,b'],['=TEXTJOIN(",","x","a","b")','#VALUE!'],['=TEXTJOIN(",","","a","b")','#VALUE!'],
  ['=TEXTJOIN(",","1","a","","b")','#VALUE!'],['=TEXTJOIN(",",A2,"a","","b")','a,,b'],['=TEXTJOIN(",",A1,"a","","b")','#VALUE!'],['=TEXTJOIN(",",A3,"a","","b")','a,b'],
  ['=TEXTJOIN(",",A4,"a","","b")','a,b'],['=TEXTJOIN(",",D1,"a","","b")','a,b'],['=TEXTJOIN(",",D2,"a","","b")','#VALUE!'],
  ['=TEXTJOIN(",",A6,"a")','#N/A'],['=TEXTJOIN(A6,TRUE,"a")','#N/A'],['=TEXTJOIN(A5:A6,TRUE,"a","b","c")','#N/A'],['=TEXTJOIN(",",TRUE,A6:A6)','#N/A']]);
  for(const [where,want] of [['I1','a,b'],['I3','a,,b'],['I25','#VALUE!']])assert.equal(on('=TEXTJOIN(",",D1:D4,"a","","b")',where),want,'ignore_empty D1:D4 in '+where);});
// Excel reads the arguments in order, but a reference given to CONCAT, or to the delimiter or a text of TEXTJOIN, only after
// all the arguments: =CONCAT(E1,NA()) is #N/A though E1 holds #DIV/0!. The ignore_empty of TEXTJOIN, and every argument of
// CONCATENATE and &, take one value, which Excel reads in its turn, from a cell too.
test('CONCAT and TEXTJOIN read a referenced cell after their other arguments, as Excel does, so that an error given directly comes first',()=>check([
  ['=CONCAT(A6,1/0)','#DIV/0!'],['=CONCAT(1/0,A6)','#DIV/0!'],['=CONCAT(A6,E1)','#N/A'],['=CONCAT(E1,A6)','#DIV/0!'],['=CONCAT(A6:A6,1/0)','#DIV/0!'],
  ['=TEXTJOIN(A6,TRUE,1/0)','#DIV/0!'],['=TEXTJOIN(A6,TRUE,E1)','#N/A'],['=TEXTJOIN(A6,1/0,"a")','#DIV/0!'],['=TEXTJOIN(",",TRUE,A6,1/0)','#DIV/0!'],
  ['=TEXTJOIN(1/0,TRUE,A6)','#DIV/0!'],['=TEXTJOIN(",",1/0,A6)','#DIV/0!'],['=TEXTJOIN(",","x",1/0)','#VALUE!'],['=TEXTJOIN(",","x",A6)','#VALUE!'],
  ['=TEXTJOIN(A6,"x","a")','#VALUE!'],['=TEXTJOIN(E1,A6,"a")','#N/A'],['=CONCATENATE(A6,1/0)','#N/A'],['=A6&1/0','#N/A']],{...SHEET,E1:{t:'f',f:'=1/0'}}));
test('a cell whose formula gives "" is no empty cell for the ignore_empty of TEXTJOIN: Excel shows #VALUE!',()=>{
  assert.equal(on('=TEXTJOIN(",",A5,"a","","b")'),'#VALUE!');assert.equal(on('=TEXTJOIN(",",A2,"a","","b")'),'a,,b','an empty cell is FALSE');});
test('an argument of CONCATENATE, CONCAT and TEXTJOIN may be left empty, as in Excel: it is empty text, and TRUE for the ignore_empty of TEXTJOIN',()=>{
  check([['=CONCATENATE("a",,"b")','ab'],['=CONCATENATE(,)',''],['=CONCATENATE("a",)','a'],['=CONCAT("a",,"b")','ab'],['=CONCAT("a",)','a'],['=TEXTJOIN(,TRUE,"a","b")','ab'],
    ['=TEXTJOIN(",",,"a","","b")','a,b'],['=TEXTJOIN(",",TRUE,"a",,"b")','a,b'],['=TEXTJOIN(",",FALSE,"a",,"b")','a,,b'],['=TEXTJOIN(",",FALSE,"a",)','a,'],
    ['=TEXTJOIN(",",TRUE,)',''],['=TEXTJOIN(,,A1:A3)','a1']]);
  // Every function takes an argument left empty, each as Excel reads it (measured): mostly as an empty cell, 0 or "" by
  // what the function wants; PRODUCT leaves it out and MODE refuses it.
  check([['=SUM(1,,2)',3],['=IF(TRUE,)',0],['=IF(FALSE,1,)',0],['=ROUND(1.5,)',2],['=AVERAGE(1,)',0.5],['=COUNT(1,)',2],['=COUNTA(1,,"")',3],['=MAX(,-1)',0],
    ['=PRODUCT(5,)',5],['=PRODUCT(2,,3)',6],['=MODE(,,1)','#VALUE!'],['=AND(,TRUE)',false],['=INDEX(A1:B2,,2)','#VALUE!']]);});
test('& and CONCATENATE cut a text at 32767 characters, as Excel does; CONCAT and TEXTJOIN show #VALUE! for a longer text (desktop Excel 16.0 shows #CALC!)',()=>{
  const y=n=>'y'.repeat(n),cells={A3:{t:'n',v:1},A6:{t:'s',v:'#N/A'},A11:{t:'s',v:y(32767)},A12:{t:'s',v:y(32766)},A13:{t:'s',v:y(32765)}};
  check([['=A11&"b"',y(32767)],['=A12&"bc"',y(32766)+'b'],['=A12&"bc"&"d"',y(32766)+'b'],['=CONCATENATE(A12,"bc")',y(32766)+'b'],['=CONCATENATE(A11,A11)',y(32767)],
    ['=CONCAT(A11,"")',y(32767)],['=CONCAT(A13,"ab")',y(32765)+'ab'],['=CONCAT(A13,"abc")','#VALUE!'],['=CONCAT(A12&"bc")',y(32766)+'b'],['=CONCAT(A11,"b",A6)','#N/A'],
    ['=TEXTJOIN(",",TRUE,A11)',y(32767)],['=TEXTJOIN("",TRUE,A12,"b")',y(32766)+'b'],['=TEXTJOIN(",",TRUE,A12,"b")','#VALUE!'],
    // TEXTJOIN counts the delimiter in front of an empty text before it leaves the text out: A11 and "" come to 32767+1.
    ['=TEXTJOIN(",",TRUE,A11,"")','#VALUE!'],['=TEXTJOIN(",",TRUE,A12,"")',y(32766)],['=TEXTJOIN(",",TRUE,A13,"","","")',y(32765)],
    ['=TEXTJOIN(",",TRUE,A13,A2:A5)','#VALUE!'],['=TEXTJOIN(",",FALSE,A13,"","")',y(32765)+',,'],['=TEXTJOIN(",",FALSE,A13,"","","")','#VALUE!'],
    ['=TEXTJOIN(",",TRUE,A11,"",A6)','#N/A']],cells);});
test('a number format shows what toLocaleString shows with the same options, and quickly for many cells',()=>{
  // Excel has no negative zero, so -0 shows as 0 in both (the next test but one).
  const old=(value,fmt)=>{if(value===0)value=0;const percent=fmt.includes('%'),decimals=fmt.includes('.00')?2:0,yen=fmt.startsWith('¥'),neg=yen&&value<0;
    return (neg?'-':'')+(yen?'¥':'')+((neg?-value:value)*(percent?100:1)).toLocaleString('en-US',{minimumFractionDigits:decimals,maximumFractionDigits:decimals,useGrouping:fmt.includes(',')})+(percent?'%':'');};
  const values=[0,1,-1,0.5,-0.5,1.005,2.675,0.125,0.0005,-0.0049,1234.5,-1234.5,999999.995,12345678.9,-98765.4321,1e15,-1e15,1e21,1.23e-7,Number.MAX_SAFE_INTEGER,Number.MIN_VALUE];
  for(let i=0;i<300;i++)values.push((Math.sin(i)*10**(i%12))*(i%3?1:-1));
  for(const fmt of ['0','0.00','#,##0','#,##0.00','0%','0.00%','¥#,##0'])for(const v of values)assert.equal(C.display(v,fmt),old(v,fmt),fmt+' '+v);
  const start=performance.now();for(let i=0;i<20000;i++)C.display(i*1.25,i%2?'#,##0':'0.00%');
  const ms=performance.now()-start;
  assert.ok(ms<250,'表示形式の書式化2万回に '+ms.toFixed(0)+'ms かかった（セルの多い表で、入力のたびに描き直す時間が延びる）');
});
test('copying a formula moves its cell references only: not the name of a function with digits (LOG10), a sheet name or a text',()=>{
  assert.equal(C.shiftFormula('=LOG10(A1)+SUM(A1:B2)',1,1),'=LOG10(B2)+SUM(B2:C3)','LOG10( は番地 LOG10 ではない');
  assert.equal(C.shiftFormula('=log10(a1)',2,0),'=log10(A3)');
  assert.equal(C.shiftFormula('=Sheet2!A1+A1',1,0),'=Sheet2!A2+A2','シート名の Sheet2 は番地ではない');
  assert.equal(C.shiftFormula('="LOG10"&A1&"B2"',0,1),'="LOG10"&B1&"B2"');
  assert.equal(C.shiftFormula('=1E5+2e3*A1',1,0),'=1E5+2e3*A2','指数で書いた数の E5 は番地ではない');
  assert.equal(C.shiftFormula('=ABCD1+A1',1,0),'=ABCD1+A2','4文字の名前の途中を番地として読まない');
  assert.equal(C.shiftFormula('=A1+$B$2+B$3',1,1),'=B2+$B$2+C$3');
  assert.equal(C.shiftFormula('=A1',-1,0),'=#REF!');
  const w=make(),s=w.sheets[0];s.cells={A1:{t:'n',v:100},A2:{t:'n',v:1000},B1:{t:'f',f:'=LOG10(A1)'},B2:{t:'f',f:C.shiftFormula('=LOG10(A1)',1,0)}};
  assert.deepEqual([C.calculate(w,s,0,1),C.calculate(w,s,1,1)],[2,3]);
});
test('two different lines on one cell edge are a border conflict; a merged cell is one block with the lines of its top-left cell, and the default style counts',()=>{
  const w=make(),s=w.sheets[0],L='1px solid #808080';
  s.merges=[C.range('B2:C3')];
  s.cells={
    A1:{t:'s',v:'',s:{borderBottom:L}},A2:{t:'s',v:'',s:{borderTop:'2px solid #217346'}},
    E1:{t:'s',v:'',s:{borderRight:L}},F1:{t:'s',v:'',s:{borderLeft:L}},
    E3:{t:'s',v:'',s:{borderRight:'1px solid #ABCDEF'}},F3:{t:'s',v:'',s:{borderLeft:'1px solid #abcdef'}},
    E5:{t:'s',v:'',s:{borderBottom:L}},
    B2:{t:'s',v:'結合',s:{borderRight:'2px solid #000000'}},C2:{t:'s',v:'',s:{borderRight:'3px double #ff0000'}},D3:{t:'s',v:'',s:{borderLeft:'1px solid #000000'}},
  };
  assert.deepEqual(C.borderConflicts(w,s),[
    {a:'A1',b:'A2',edge:'下/上',lines:[L,'2px solid #217346'],cells:['A1','A2']},
    {a:'C3',b:'D3',edge:'右/左',lines:['2px solid #000000','1px solid #000000'],cells:['B2:C3','D3']},
  ],'同じ線・大文字と小文字だけ違う色・片方だけの線・結合で隠れたセルの線は食い違いではない');
  assert.deepEqual(C.borderConflicts(w,s,C.range('A1:A1')),[],'範囲を渡すと、その中の境界だけを見る');
  assert.equal(C.borderConflicts(w,s,C.range('A1:A2')).length,1);
  const d=make(),ds=d.sheets[0];ds.defaultStyle={borderTop:L};ds.cells={B1:{t:'s',v:'',s:{borderBottom:'2px solid #000000'}},C1:{t:'s',v:'',s:{borderBottom:L}}};
  assert.deepEqual(C.borderConflicts(d,ds).map(x=>[x.a,x.b]),[['B1','B2']],'既定の書式の線とも比べる');
});
test('the references of a formula are found as formulas.js reads them, also when the formula goes to Excel: 1E5 is a number and LOG10( a function',()=>{
  assert.equal(C.blockFormula('=1E5+B2',C.range('B2:F6')),'=1E5+A1','指数で書いた数の E5 を、範囲の番地として書き換えない');
  assert.equal(C.blockFormula('=LOG10(B2)+ABCD1',C.range('B2:C3')),'=LOG10(A1)+ABCD1');
  assert.equal(C.mapRefs('=A1+"B2"&$C$3+LOG10(d4)+1E5+Sheet2!E6',(m,c$,col,r$,row)=>'['+c$+col+r$+row+']'),'=[A1]+"B2"&[$C$3]+LOG10([d4])+1E5+Sheet2![E6]');
});
test('cautions tell formulas the page cannot read, ones it cannot compute that Excel may, ones that only receive such an error, and errors Excel shows too, each with its cause',()=>{
  const w=make(),s=w.sheets[0];
  s.cells={
    A1:{t:'f',f:'=A2*Tax'},A2:{t:'f',f:'=SUM(B1'},A3:{t:'f',f:'=Sheet2!A1'},
    B1:{t:'f',f:'=TEXTBEFORE("a-b","-")'},B2:{t:'f',f:'=E99+SUM(A1:P1)'},B3:{t:'f',f:'=B3+1'},
    C1:{t:'f',f:'=B1&"x"'},C2:{t:'f',f:'=INDEX(D1:D3,5)'},C3:{t:'f',f:'=#REF!'},C4:{t:'f',f:'=C2*2'},
    D5:{t:'f',f:'=SUM(D1:D3)'},
  };
  const c=C.cautions(w);
  assert.deepEqual(c.unreadable,['Sheet1!A1 =A2*Tax','Sheet1!A2 =SUM(B1','Sheet1!A3 =Sheet2!A1']);
  assert.deepEqual(c.uncomputable,['Sheet1!B1 #NAME?（TEXTBEFORE はHTMLで計算できない関数）','Sheet1!B2 #REF!（E99・P1 はシートの外。シートは30行・15列）','Sheet1!B3 #CIRC!（循環参照）']);
  assert.deepEqual(c.dependent,['Sheet1!C1 #NAME?','Sheet1!C4 #REF!'],'参照先のエラーを受けているだけの数式');
  assert.deepEqual(c.errors,['Sheet1!C2 #REF!','Sheet1!C3 #REF!'],'Excelでも同じエラーになる数式');
  assert.deepEqual(c.textFormulas,[]);
});
test('copying a formula does not change a quoted sheet name, and a reference after a sheet name is told with its sheet',()=>{
  assert.equal(C.shiftFormula("='Q1 data'!A1+A1",1,1),"='Q1 data'!B2+B2",'シート名の中の Q1 は番地ではない');
  assert.equal(C.shiftFormula("='It''s B2'!C3",1,0),"='It''s B2'!C4");
  assert.equal(C.shiftFormula("='Q1 data'!A:A+A1",1,0),"='Q1 data'!A:A+A2",'番地の続かないシート名の中も書き換えない');
  assert.equal(C.shiftFormula('=見積!A1+A1',1,0),'=見積!A2+A2');
  const seen=[];C.mapRefs("=Sheet2!A1+'Q1 data'!$B$2+C3+見積!D4",(m,c$,col,r$,row,sheet)=>{seen.push([m,sheet]);return m;});
  assert.deepEqual(seen,[['A1','Sheet2'],['$B$2',"'Q1 data'"],['C3',undefined],['D4','見積']]);
  const range=[];C.mapRefs("=SUM(Sheet2!A1:B2)+'Q1 data'!A1:$B$5+C1:D2",(m,c$,col,r$,row,sheet)=>{range.push([m,sheet]);return m;});
  assert.deepEqual(range,[['A1','Sheet2'],['B2','Sheet2'],['A1',"'Q1 data'"],['$B$5',"'Q1 data'"],['C1',undefined],['D2',undefined]],'範囲の後ろの番地も、前の番地のシートのもの');
  assert.equal(C.shiftFormula('=売上A1+A1',1,0),'=売上A1+A2','日本語の名前の中の番地の形は番地ではない');
  assert.equal(C.shiftFormula('=A1売上+A1',1,0),'=A1売上+A2');
});
test('a negative zero shows as 0, as Excel has no negative zero',()=>{
  for(const [fmt,shown] of [['General','0'],['0','0'],['0.00','0.00'],['#,##0','0'],['0%','0%'],['¥#,##0','¥0']])assert.equal(C.display(-0,fmt),shown,fmt);
  const w=make(),s=w.sheets[0];s.cells={A1:{t:'n',v:0},B1:{t:'f',f:'=A1*-1',s:{numFmt:'¥#,##0'}},B2:{t:'f',f:'=ROUND(-0.4,0)',s:{numFmt:'0'}}};
  assert.deepEqual([C.display(C.calculate(w,s,0,1),'¥#,##0'),C.display(C.calculate(w,s,1,1),'0')],['¥0','0']);
});
test('a reference stands alone: not after a point (1.E5, x.A1) nor before an underscore (Q1_total), as formulas.js reads them',()=>{
  assert.equal(C.shiftFormula('=1.E5+A1',1,0),'=1.E5+A2');
  assert.equal(C.shiftFormula('=Q1_total+A1',1,0),'=Q1_total+A2');
  assert.equal(C.shiftFormula('=x.A1+A1',1,0),'=x.A1+A2');
});
test('a formula that only reads an error of another kind is an error of its own; a hidden function or reference the page cannot compute is told even when IFERROR hides it; SUMIF reaching beyond the sheet is told so',()=>{
  const w=make(),s=w.sheets[0];
  s.cells={
    A1:{t:'f',f:'=1/0'},A2:{t:'n',v:1},
    B1:{t:'f',f:'=VLOOKUP(1,A1:A2,3,FALSE)'},
    C1:{t:'f',f:'=IFERROR(VLOOKUP(1,$A$1:$B$100,2,FALSE),"なし")'},C2:{t:'f',f:'=IFERROR(TEXTBEFORE("a-b","-"),"")'},
    D1:{t:'f',f:'=SUMIF(A1:A30,">0",B5)'},
  };
  const c=C.cautions(w);
  assert.deepEqual(c.errors,['Sheet1!B1 #REF!'],'読んだのは #DIV/0! で、式の #REF! は VLOOKUP の列番号から来る');
  assert.deepEqual(c.dependent,[]);
  assert.deepEqual(c.uncomputable,['Sheet1!C1 "なし"（B100 はシートの外。シートは30行・15列）','Sheet1!C2 ""（TEXTBEFORE はHTMLで計算できない関数）','Sheet1!D1 #REF!（B34 はシートの外。シートは30行・15列）']);
});
test('copying a formula keeps a reference that does not move (absolute, or not moved that way) as written, and turns a moved one into #REF! only below row 1 or column A, or beyond Excel',()=>{
  assert.equal(C.shiftFormula('=SUM($A$1:$A$1000)+B1',1,0),'=SUM($A$1:$A$1000)+B2','動かない番地は、シートの外でも書いたまま');
  assert.equal(C.shiftFormula('=A600',0,1),'=B600');
  assert.equal(C.shiftFormula('=A499',1,0),'=A500','Excelの行の範囲なら、シートの外へ動いても番地のまま（HTMLでは #REF! と計算し、ビルドが知らせる）');
  assert.equal(C.shiftFormula('=A1048576',1,0),'=#REF!');
  assert.equal(C.shiftFormula('=XFD1',0,1),'=#REF!');
  assert.equal(C.shiftFormula('=B2',-2,0),'=#REF!');
});
test('a reference is not taken from inside a name before a point or after an underscore (A1.x, _A1)',()=>{
  assert.equal(C.shiftFormula('=A1.x+A1',1,0),'=A1.x+A2');
  assert.equal(C.shiftFormula('=_A1+A1',1,0),'=_A1+A2');
});
test('validate says where a wrong sheet field is and what it takes',()=>{
  const sheetWith=patch=>{const w=make();Object.assign(w.sheets[0],patch);return w;};
  const cases=[
    [sheetWith({colWidths:[20]}),/sheets\[0\]\.colWidths（24〜2044 のピクセルの数か null を、列の数まで並べる）/],
    [sheetWith({rowHeights:[600]}),/sheets\[0\]\.rowHeights（16〜546 のピクセルの数か null を、行の数まで並べる）/],
    [sheetWith({name:'見積/2026'}),/sheets\[0\]\.name "見積\/2026"（シート名は31文字以内で/],
    [sheetWith({rows:501}),/sheets\[0\]\.rows と cols（上限500行 \/ 100列の、1以上の整数）/],
    [sheetWith({cells:{A99:{t:'s',v:'x'}}}),/sheets\[0\]\.cells\.A99 範囲外のセル（シートは30行・15列）/],
    [sheetWith({cells:{A1:{t:'n',v:'1'}}}),/sheets\[0\]\.cells\.A1\.v（t が s なら文字、n なら有限の数、b なら true \/ false）/],
    [sheetWith({merges:[C.range('A1:B2'),C.range('B2:C3')]}),/結合の重複（sheets\[0\]\.merges\[1\] B2:C3）/],
  ];
  for(const [w,message] of cases)assert.throws(()=>C.validate(w),message,String(message));
  const w=make();w.activeSheet='Sheet1';
  assert.throws(()=>C.validate(w),/activeSheet "Sheet1"（最初に開くシートの id。s のどれか）/);
});
