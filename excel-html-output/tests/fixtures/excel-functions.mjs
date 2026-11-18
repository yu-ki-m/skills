// The formulas of tests/fixtures/excel-functions.xlsx: for each kind of function, a sheet holds the cells the formulas
// refer to (columns A to G, and a few more) and the formulas themselves (column H, from row 1; a few in other cells,
// where the row or column of the formula matters). Excel computes them and saves its values with them, and
// tests/xlsx-functions.test.mjs checks that the page computes the same value for every formula. Most cases were first
// found by trying many formulas in Excel; a comment above a group says what Excel was seen to do.
//
// To make the file again (Windows, desktop Excel, PowerShell 7):  node tests/fixtures/excel-functions.mjs
// The values must not depend on the day the file is made: TODAY() and NOW() appear only in formulas whose value does not.
// A date written without its year ("1/2") is in the year the file was made; the test computes with its clock set to the
// time the file says it was made (docProps/core.xml).
import {spawnSync} from 'node:child_process';
import {mkdtempSync, rmSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';

/** Every function name with every argument: [f(a1), f(a2), ...] for each name. */
const each = (names, args) => names.flatMap(name => args.map(a => `=${name}(${a})`));

// Cells of each kind, which most sheets share: a text, an empty cell, a number, TRUE, a formula giving "", #N/A, #DIV/0!,
// a number written as text, 0, the text TRUE, FALSE, negative and fractional numbers.
const KINDS = {
  A1: 'a', A3: 1, A4: true, A5: {f: '=""'}, A6: {f: '=NA()'}, A7: {f: '=1/0'}, A8: '2', A9: 0, A10: 'TRUE', A11: false, A12: -3.5, A13: 2.9,
  A14: 'A', A15: ' sp ',
};
const KIND_REFS = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11'];

const logical = {
  name: '論理', cells: {...KINDS},
  formulas: [
    '=IF(A3>0,"正","負")', '=IF(A1="a","はい")', '=IF(A11,1)', ...['A9', 'A10', 'A8', 'A1', 'A6', 'A2', 'A5', '"TRUE"', '"true"', '"abc"', '2', '0.5', '" "', '""'].map(c => `=IF(${c},1,2)`),
    '=IF(TRUE,)', '=IF(FALSE,1,)', '=IF(FALSE,1)', '=IF(TRUE,,1)', '=IF(TRUE,1/0,1)', '=IF(FALSE,1/0,1)', '=IF(TRUE,A2)', '=IF(TRUE,A5)', '=IF(A6,1/0)', '=IF(1<2,"a","b")&"!"',
    '=IFS(A3>1,"a",A3=1,"b")', '=IFS(FALSE,1)', '=IFS(A9,1,TRUE,2)', '=IFS("TRUE",1)', '=IFS("x",1)', '=IFS(A6,1)', '=IFS(TRUE,1/0)', '=IFS(FALSE,1/0,TRUE,2)',
    '=IFS(A2,1,TRUE,2)', '=IFS(1,"x")', '=IFS(A1,1)', '=IFS(FALSE,1,FALSE,2)', '=IFS(TRUE,A2)', '=IFS(A8,1)',
    '=IFERROR(A6,"x")', '=IFERROR(A7,"x")', '=IFERROR(A3,"x")', '=IFERROR(1/0,)', '=IFERROR(A2,"x")', '=IFERROR(A6,A7)', '=IFERROR(A5,"x")', '=IFERROR("a"+1,"数値でない")',
    '=IFNA(A6,"x")', '=IFNA(A7,"x")', '=IFNA(A3,"x")', '=IFNA(A2,"x")', '=IFNA(NA(),1/0)', '=IFNA(A1,1)', '=IFNA(NA(),)',
    '=AND(TRUE,TRUE)', '=AND(TRUE,FALSE)', '=AND(A3:A4)', '=AND(A1:A4)', '=AND(A2)', '=AND("TRUE")', '=AND("a")', '=AND(1,"a")', '=AND(A10)', '=AND(A6,TRUE)', '=AND(,TRUE)',
    '=AND(A1:A3,FALSE)', '=AND(A9:A11)', '=AND(1/0,A6)', '=AND(A8)', '=AND(2)', '=AND(A5)',
    '=OR(FALSE,0)', '=OR(A9,A11)', '=OR(A1)', '=OR(1,"a")', '=OR(,TRUE)', '=OR(A1:A3)', '=OR("FALSE","TRUE")', '=OR(A6,TRUE)', '=OR(A2,A9)',
    '=XOR(TRUE,FALSE)', '=XOR(TRUE,TRUE)', '=XOR(TRUE,TRUE,TRUE)', '=XOR(A3:A4,A11)', '=XOR(A1)', '=XOR(1,0)', '=XOR("TRUE",FALSE)', '=XOR(A6)', '=XOR(A9:A11)',
    ...['A9', 'A3', '"TRUE"', '"a"', 'A2', 'A1', 'A6', '2.5', 'A10', 'A8', '""'].map(c => `=NOT(${c})`),
    '=SWITCH(A3,1,"一",2,"二")', '=SWITCH(2,1,"一",2,"二")', '=SWITCH(3,1,"一",2,"二")', '=SWITCH(3,1,"一","他")', '=SWITCH(A1,"A","大文字")', '=SWITCH(A2,0,"零","空")',
    '=SWITCH(A2,"","空文字")', '=SWITCH(A6,1,2)', '=SWITCH(1,A6,2,1,3)', '=SWITCH(1,1,1/0,2)', '=SWITCH("1",1,"数値","1","文字")', '=SWITCH(TRUE,1,"1",TRUE,"真")',
    '=SWITCH(A5,"","空文字",0,"零")', '=SWITCH(2,1,"a",2,A2)',
    '=TRUE()', '=FALSE()', '=NA()', '=TRUE()+1', '=ISNA(NA())',
    '=CHOOSE(2,"a","b","c")', '=CHOOSE(1.9,"a","b")', '=CHOOSE(0,"a")', '=CHOOSE(4,"a","b")', '=CHOOSE("2","a","b")', '=CHOOSE(TRUE,"a","b")', '=CHOOSE(A2,"a")',
    '=CHOOSE(2,1/0,"b")', '=CHOOSE(A6,1)', '=CHOOSE(2,"a",)', '=SUM(CHOOSE(2,A3:A4,A12:A13))', '=CHOOSE(-1,"a")', '=CHOOSE(A1,"a")',
    ...each(['ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL', 'ISERROR', 'ISERR', 'ISNA'], [...KIND_REFS, '"x"', '1', '1/0', 'NA()', 'TRUE', '""', 'A1:A3', 'A20:A22']),
    ...each(['ISEVEN', 'ISODD'], ['2', '3', '-3', '2.9', '-2.5', '0', '"4"', 'A8', 'TRUE', 'A2', '"a"', 'A6', 'A9', 'A12', 'A13', '-0.5']),
    ...each(['N'], ['A3', 'A1', 'A4', 'A6', 'A2', '"5"', 'TRUE', 'A8', 'A5', 'DATE(2026,1,2)', '1/0', 'A11']),
    ...each(['T'], ['A1', 'A3', 'A4', 'A6', 'A2', '"x"', '1', 'A5', 'TRUE', '1/0', 'A8']),
  ],
};

const math = {
  name: '数学',
  cells: {
    B1: 3, B2: 5, B3: '7', B4: false, B5: 2.5, B7: -4, B8: {f: '=1/0'}, B9: 0.5, C1: 1, C2: 2, C3: 3, D1: 4, D2: 5, D3: 6, E1: 'a', E2: 2, E3: true,
    F1: {v: 46023, fmt: 'yyyy/mm/dd'}, F2: {v: 46054, fmt: 'yyyy/mm/dd'}, F3: {v: 46024, fmt: 'yyyy/mm/dd'}, G1: 'abc', G2: 'ｱｲ',
  },
  formulas: [
    '=SUM(B1:B5)', '=SUM(B1:B9)', '=SUM("2",1)', '=SUM(TRUE,1)', '=SUM("a",1)', '=SUM(B3)', '=SUM(B4)', '=SUM(1,,2)', '=SUM(C1:C3,D1:D3)', '=SUM(C1:D3)', '=SUM(B6)',
    '=SUM(B1:B5,"7")', '=SUM(0.1,0.2)', '=SUM(E1:E3)', '=SUM(E1:E3,TRUE)',
    '=PRODUCT(C1:C3)', '=PRODUCT(B1:B5)', '=PRODUCT("2",3)', '=PRODUCT(TRUE,5)', '=PRODUCT(B6)', '=PRODUCT(B3)', '=PRODUCT(,5)', '=PRODUCT(C1:D3)', '=PRODUCT(B7,B9)', '=PRODUCT("a")',
    ...[['1.005', 2], ['2.5', 0], ['-2.5', 0], ['1234.5678', -2], ['0.285', 2], ['1.15', 1], ['-1.15', 1], ['123456789.125', 2], ['0.5', 0], ['1.45', 1], ['2.675', 2], ['10.555', 2],
      ['0.1+0.2', 15], ['"1.25"', 1], ['TRUE', 0], ['B8', 1], ['2.5', '"a"'], ['2.5', 1.9], ['2.55', -0.5], ['5555', -5], ['1E+20', 2], ['-0.0000000001', 5], ['1234.5', '"0"'], ['B6', 0],
      ['4.35', 1], ['1.0000000000000002', 0], ['0.1+0.7', 1], ['8.345', 2], ['-8.345', 2]]
      .flatMap(([x, d]) => ['ROUND', 'ROUNDUP', 'ROUNDDOWN'].map(f => `=${f}(${x},${d})`)),
    '=ROUNDUP(0.1+0.2,1)', '=ROUNDDOWN(0.1*3*10,0)', '=ROUND(B5,)',
    ...each(['INT', 'TRUNC'], ['-1.5', '2.9999999999999996', '-0.1', '"2.5"', 'TRUE', 'B6', '1E+20', '-2.0000000000000004', '"a"', 'B8']),
    '=TRUNC(1.239,2)', '=TRUNC(1234,-2)', '=TRUNC(-1.239,1)', '=TRUNC(1.239,"1")', '=TRUNC(0.1+0.2,16)',
    ...each(['ABS'], ['-3', '"-2"', 'TRUE', 'B6', '"a"', 'B7', 'B8']),
    ...[['7', '4'], ['-7', '4'], ['7', '-4'], ['-7', '-4'], ['7.5', '2'], ['1', '0'], ['0', '5'], ['2.6', '0.2'], ['1E+15', '7'], ['"7"', '2'], ['5', '2.5'], ['-1', '3'], ['0.3', '0.1'], ['TRUE', '2']]
      .flatMap(([a, b]) => [`=MOD(${a},${b})`, `=QUOTIENT(${a},${b})`]),
    ...each(['SIGN'], ['-2', '0', '3.5', '"-1"', 'TRUE', 'B6', '"a"']),
    ...each(['SQRT'], ['16', '2', '-1', '"9"', '0', 'B6', 'TRUE']),
    '=POWER(2,10)', '=POWER(2,0.5)', '=POWER(-8,1/3)', '=POWER(0,0)', '=POWER(0,-1)', '=POWER(-2,2)', '=POWER("2",3)', '=POWER(10,400)', '=POWER(-2,0.5)', '=POWER(2,-1)',
    '=2^10', '=(-8)^(1/3)', '=0^0', '=0^-1', '=(-2)^2', '=-2^2', '=2^3^2', '=-2^0.5', '=2^-1', '=10^400', '=-B5^2', '=--"5"', '=-"5"', '=+"5"&"x"', '=50%%', '=10%*2', '=-50%',
    '=1<2<3', '=3>2>1', '=1+2&3', '=1&2+3', '=2*3^2', '=-(2^2)', '=-(-2)^2',
    ...each(['EXP'], ['1', '0', '-1', '710', '"2"']), ...each(['LN'], ['EXP(2)', '0', '-1', '1', '"10"']), '=LOG(100)', '=LOG(8,2)', '=LOG(1,1)', '=LOG(10,0)', '=LOG(-10)',
    '=LOG(1000,10)', '=LOG(2,8)', ...each(['LOG10'], ['1000', '0', '0.01', '"100"', '-1']), '=PI()', '=PI()*2', '=ROUND(PI(),10)',
    ...[['2.5', '1'], ['-2.5', '-1'], ['-2.5', '1'], ['2.5', '-1'], ['1.5', '0.1'], ['0', '5'], ['5', '0'], ['4.42', '0.05'], ['1234', '100'], ['1.58', '0.1'], ['-4.42', '-0.05'],
      ['7', '"2"'], ['0.3', '0.1'], ['2.1', '0.7']].flatMap(([x, s]) => [`=CEILING(${x},${s})`, `=FLOOR(${x},${s})`]),
    '=CEILING.MATH(-2.5)', '=CEILING.MATH(-2.5,1,1)', '=CEILING.MATH(2.5,-1)', '=CEILING.MATH(24.3,5)', '=CEILING.MATH(2.5)', '=CEILING.MATH(-24.3,5,-1)', '=CEILING.MATH(0)',
    '=CEILING.MATH(6.7,0)', '=FLOOR.MATH(-2.5)', '=FLOOR.MATH(-2.5,1,1)', '=FLOOR.MATH(24.3,5)', '=FLOOR.MATH(2.5)', '=FLOOR.MATH(-24.3,-5)', '=FLOOR.MATH(6.7,0)', '=FLOOR.MATH(-8.1,2,TRUE)',
    ...[['10', '3'], ['-10', '-3'], ['1.3', '0.2'], ['5', '-2'], ['7.5', '5'], ['2.5', '0'], ['-7.5', '-5'], ['6.05', '0.1'], ['1.5', '1'], ['0', '3'], ['3.5', '1']].map(([a, b]) => `=MROUND(${a},${b})`),
    '=SUMPRODUCT(C1:C3,D1:D3)', '=SUMPRODUCT(C1:C3)', '=SUMPRODUCT(C1:C3,D1:D2)', '=SUMPRODUCT((C1:C3>1)*D1:D3)', '=SUMPRODUCT(--(C1:C3>1))', '=SUMPRODUCT((E1:E3="a")*1)',
    '=SUMPRODUCT(C1:C3*D1:D3)', '=SUMPRODUCT(C1:C3+1)', '=SUMPRODUCT(E1:E3)', '=SUMPRODUCT(E1:E3,C1:C3)', '=SUMPRODUCT((C1:C3=2)+(D1:D3=6))', '=SUMPRODUCT(C1:C3,B7:B9)',
    '=SUMPRODUCT(C1:C3*2)', '=SUMPRODUCT(LEN(G1:G3))', '=SUMPRODUCT((MONTH(F1:F3)=1)*1)', '=SUMPRODUCT(C1:D3)', '=SUMPRODUCT(C1:D3,C1:D3)', '=SUMPRODUCT((C1:C3>1)*(D1:D3<6))',
    '=SUMPRODUCT(B1:B5)', '=SUMPRODUCT(B1:B5*1)', '=SUMPRODUCT(C1:C3/D1:D3)', '=SUMPRODUCT(1/C1:C3)', '=SUMPRODUCT(C1:C3^2)', '=SUMPRODUCT(-C1:C3)', '=SUMPRODUCT(C1:C3&"")',
    '=SUMPRODUCT((C1:C3>=2)*(C1:C3<=3)*D1:D3)', '=SUMPRODUCT(ISNUMBER(E1:E3)*1)', '=SUMPRODUCT(--ISTEXT(E1:E3))', '=SUMPRODUCT(ABS(B7:B7))', '=SUMPRODUCT(ROUND(B9*C1:C3,0))',
    // MOD is the exact remainder of the numbers as held; a quotient beyond 2^50/1000 is #NUM!.
    ...[['6', '0.1'], ['1', '0.1'], ['0.7', '0.1'], ['3', '0.3'], ['-6', '0.1'], ['6', '-0.1'], ['2^40', '0.9999'], ['1.1E+12', '1.1'], ['2^40+0.5', '1'],
      ['1125899906842.62', '1'], ['1.1259E+12', '1'], ['2^41', '1.99999'], ['2^41', '1.9'], ['-1.1259E+12', '1'], ['1', '2^-40'], ['1', '2^-41']].map(([a, b]) => `=MOD(${a},${b})`),
    // A negative number has a root only where 1/y is an odd whole number.
    ...['(-8)^(1/3)', '(-8)^(2/3)', '(-32)^0.2', '(-8)^(-1/3)', '(-64)^(1/6)', '(-1)^(1/7)', '(-8)^(1/3+2^-50)', '(-128)^(1/7)', '(-8)^(-2/3)', 'POWER(-27,1/3)']
      .map(e => `=${e}`),
    '=MROUND(0.15,0.1)', '=MROUND(0.25,0.1)', '=MROUND(0.35,0.1)', '=MROUND(-0.15,-0.1)', '=MROUND(TRUE,1)', '=MROUND(5,TRUE)', '=MROUND(B4,1)', '=MROUND("5",2)',
    '=QUOTIENT(7,TRUE)', '=QUOTIENT("7","2")', '=QUOTIENT(B4,1)', '=QUOTIENT(B3,2)', '=PRODUCT(5,)', '=PRODUCT(,,5)', '=PRODUCT(2,,3)', '=PRODUCT(,)',
    // Numbers are equal when they are the same to 15 significant digits.
    '=1+22*2^-52=1', '=1+23*2^-52=1', '=1-10*2^-53=1', '=1-2^-53=1', '=1000+30*2^-43=1000', '=1<1+23*2^-52', '=1<1+22*2^-52', '=1.5+22*2^-52=1.5', '=1.99+23*2^-52=1.99',
  ],
};

const stats = {
  name: '統計',
  cells: {A1: 5, A2: 3, A3: 8, A4: 3, A5: 9, A6: 1, A7: 3, A8: 8, A10: 'x', A11: true, A12: '4', A13: {f: '=""'}, B1: 10, B2: 20, B3: 30, B4: 40, C1: 1, C2: 1, C3: 1,
    D1: 2.5, D2: -1, D3: 2.5, D4: -1, E1: {f: '=1/0'}, E2: 4},
  formulas: [
    '=AVERAGE(A1:A13)', '=AVERAGE(A1:A8)', '=AVERAGE(TRUE,1)', '=AVERAGE("2",4)', '=AVERAGE(A10:A13)', '=AVERAGE(B1:B4,100)', '=AVERAGE(E1:E2)',
    ...each(['MEDIAN'], ['A1:A8', 'A1:A7', '1,2,"3"', 'A10', 'A1:A13', 'B1:B4', 'D1:D4', 'TRUE,3', 'E1:E2', '"a",1']),
    ...each(['MODE', 'MODE.SNGL'], ['A1:A8', 'B1:B4', '1,2,2,1', '"3",3', 'D1:D4', 'C1:C3', 'A1:A13', '2,1,1,2', 'E1:E2', 'A10:A13']),
    ...[1, 2, 3, 8, 9, 0, 1.5, 2.9, '"2"', 'TRUE', -1].flatMap(k => [`=LARGE(A1:A8,${k})`, `=SMALL(A1:A8,${k})`]),
    '=LARGE(A1:A13,1)', '=SMALL(A1:A13,1)', '=LARGE(1,1)', '=LARGE(E1:E2,1)', '=SMALL(A10:A13,1)', '=LARGE(D1:D4,2)', '=SMALL(D1:D4,3)', '=LARGE((A1:A8),9-1)',
    ...['8,A1:A8', '3,A1:A8', '3,A1:A8,1', '9,A1:A8,0', '7,A1:A8', '"8",A1:A8', 'A1,A1:A8', '8,A1:A13', '1,A1:A8,1', '2.5,D1:D4', '-1,D1:D4,1', '8,A1:A8,2', '8,A1:A8,"1"', 'A10,A1:A8',
      'E1,A1:A8', '4,E1:E2'].flatMap(a => [`=RANK(${a})`, `=RANK.EQ(${a})`]),
    ...each(['COUNTBLANK'], ['A1:A13', 'A1:A3', 'A9', 'A13', 'A9:A13', 'E1:E2']),
    ...each(['STDEV', 'STDEV.S', 'STDEV.P', 'STDEVP', 'VAR', 'VAR.S', 'VAR.P', 'VARP'], ['A1:A8', '1', '1,"2",TRUE', 'A1:A13', 'C1:C3', 'B1:B4,50', 'D1:D4', 'E1:E2', '"a",1']),
    ...each(['MAX', 'MIN'], ['A1:A13', '"10",A1:A8', 'TRUE,A1:A8', 'A10:A13', 'D1:D4', 'E1:E2', '"a"', ',-1', 'A11', 'TRUE', '"3",1']),
    ...each(['COUNT', 'COUNTA'], ['A1:A13', 'A9:A13', '1,,"",TRUE,"a"', 'E1:E2', '1/0', 'A10', '"1"']),
    // MODE counts only numbers, also of what is given directly, and refuses an empty argument.
    ...each(['MODE'], ['TRUE,TRUE,1', '"1","1",2,2', '1,1,"a"', '"a",1,1', '1,2,TRUE,TRUE', ',,1', '1,2,3,3,2', '3,1,1,3', 'A11,A11,1']),
    // k between two whole numbers: LARGE takes the next, SMALL the one before; k must lie between 1 and the count.
    ...['0.5', '8.5', '7.1', '1.01', '7.99', '8', '1'].flatMap(k => [`=SMALL(A1:A8,${k})`, `=LARGE(A1:A8,${k})`]),
    '=RANK(8,A1:A8,"0")', '=RANK(8,A1:A8,A10)', '=RANK(8,A1:A8,A12)', '=RANK(8,A1:A8,A11)', '=RANK(TRUE,A1:A8)', '=RANK(8,A1:A8,0.5)', '=RANK(8,A1:A8,-1)',
  ],
};

const conditions = {
  name: '条件',
  cells: {
    A1: '東京', A2: '大阪', A3: '東京', A4: '名古屋', A5: '東京都', A6: 'abc', A7: 'ABC', A8: 'a*c', A10: '5', A11: 5, A12: true, A13: {f: '=NA()'}, A14: {f: '=""'},
    B1: 100, B2: 200, B3: 300, B4: 400, B5: 500, B6: 600, B7: 700, B8: 800, B9: 900, B10: '1000', B11: 1100, B12: false, B13: 1300, B14: 1400,
    C1: {v: 46023, fmt: 'yyyy/mm/dd'}, C2: {v: 46024, fmt: 'yyyy/mm/dd'}, C3: {v: 46030, fmt: 'yyyy/mm/dd'}, C4: {v: 46054, fmt: 'yyyy/mm/dd'}, C5: {v: 46055, fmt: 'yyyy/mm/dd'},
    C6: {v: 46100, fmt: 'yyyy/mm/dd'}, C7: '2026/1/2', C8: 46024.5,
    E1: '東京', E2: '>200', F1: 100, F2: 200, F3: 300,
    J1: 0, J2: '0', J4: {f: '=""'}, J5: 'x', J6: 1000, K1: '1000', K2: 5, K3: 'a',
  },
  formulas: [
    ...['"東京"', '"東京*"', '"*京"', '"?京"', '"東?"', '"~*"', '"a~*c"', '"a*c"', '"<>東京"', '""', '"="', '"<>"', '">5"', '">=5"', '"<5"', '"5"', '5', '"=5"', '"<>5"', '">a"', '"<b"',
      '"TRUE"', 'TRUE', '"#N/A"', '"*"', '"?*"', '"東京 "', '"ABC"', '"abc"', 'A1', 'B1', '"="&A1', '"<>"&""', '"*東京*"', '0', '"0"', '">"&5', '"=TRUE"', '"<"&"東京"', 'A9', 'A14',
      'E1', '"<>*"', '"=*"', '"東*京"', '"> 5"', '">=5.0"', '"=5.0"', '"*5*"', '"<=東京"', '"<>a*c"', '">A"', 'NA()', '"=東京都"', '"a?c"', '"~?"', '"<>#N/A"', '1/0']
      .map(c => `=COUNTIF(A1:A14,${c})`),
    // An empty cell as the criterion is 0; "<>0" leaves out only the number 0; spaces around a number are allowed.
    ...['J3', '0', '"0"', '"=0"', '"<>0"', '"<>x"', '"<>"', '""', '"> 0"', '" 0"', '"0 "', '"=0 "', '"=x "', '">=0"', '"<1"', '"=1E3"', '1000', '"<>1000"', '"> -1"', '"< 1"']
      .map(c => `=COUNTIF(J1:J6,${c})`),
    ...['"<>1000"', '"=1000"', '1000', '"<>"&1000', '">999"', '"<>5"', '"<>a"', '">"&4'].map(c => `=COUNTIF(K1:K3,${c})`),
    ...['">500"', '"<=300"', '">=1000"', '1000', '"1000"', '">"&B3', '"<>"&100', 'FALSE', '"<>FALSE"', '"1*"', '"<>"', '""', '">=1E3"', '"=1000"', 'E2', '">1,000"', '">¥500"', '">50%"']
      .map(c => `=COUNTIF(B1:B14,${c})`),
    ...['">="&DATE(2026,1,2)', '"<2026/1/10"', '"2026/1/2"', 'DATE(2026,1,2)', '">=46024"', '"<=2026年1月3日"', '"1/2"', '">"&C1'].map(c => `=COUNTIF(C1:C8,${c})`),
    '=COUNTIF(B1:B14,">"&AVERAGE(B1:B9))', '=COUNTIF(A1:A14,"*")+COUNTIF(A1:A14,"<>*")', '=COUNTIF(A1:C3,"東京")', '=COUNTIF(A2,"大阪")', '=COUNTIF(A1:A14,1/0)',
    '=COUNTIFS(A1:A14,"東京",B1:B14,">100")', '=COUNTIFS(A1:A14,"東京*",B1:B14,"<600",C1:C14,">0")', '=COUNTIFS(A1:A14,"東京",B1:B3,">0")', '=COUNTIFS(A1:A5,"東京")',
    '=COUNTIFS(B1:B14,">200",B1:B14,"<=600")', '=COUNTIFS(A1:A14,"<>東京",A1:A14,"<>")', '=COUNTIFS(A1:B2,"東京",B1:C2,">0")',
    '=SUMIF(A1:A14,"東京",B1:B14)', '=SUMIF(B1:B14,">500")', '=SUMIF(A1:A14,"東京",B1:B1)', '=SUMIF(A1:A14,"東京*",B2:B3)', '=SUMIF(A1:A14,"*",B1:B14)', '=SUMIF(A1:A14,"<>東京",B1:B14)',
    '=SUMIF(B1:B14,">0",A1:A14)', '=SUMIF(A1:A14,"東京",C1:C14)', '=SUMIF(A1:A3,"東京",B1:B3)', '=SUMIF(A1:A14,A1,B1:B14)', '=SUMIF(A1:A14,"none",B1:B14)', '=SUMIF(A1:A14,5,B1:B14)',
    '=SUMIF(A1:A14,TRUE,B1:B14)', '=SUMIF(A1:A14,"",B1:B14)', '=SUMIF(A1:A14,"=",B1:B14)', '=SUMIF(A1:A14,"#N/A",B1:B14)', '=SUMIF(C1:C8,">="&DATE(2026,1,5),B1:B8)', '=SUMIF(A1:B3,"東京")',
    '=SUMIFS(B1:B14,A1:A14,"東京",B1:B14,">100")', '=SUMIFS(B1:B14,A1:A14,"東京",C1:C3,">0")', '=SUMIFS(B1:B14,A1:A14,"東京*")', '=SUMIFS(B1:B14,B1:B14,">=500",B1:B14,"<1000")',
    '=SUMIFS(C1:C8,B1:B8,">200")', '=SUMIFS(B1:B3,A1:A3,"none")',
    '=AVERAGEIF(A1:A14,"東京",B1:B14)', '=AVERAGEIF(A1:A14,"none",B1:B14)', '=AVERAGEIF(B1:B14,">500")', '=AVERAGEIF(A1:A14,"東京*",B1:B1)', '=AVERAGEIF(A1:A14,"*",A1:A14)',
    '=AVERAGEIFS(B1:B14,A1:A14,"東京*",B1:B14,">100")', '=AVERAGEIFS(B1:B14,A1:A14,"none")', '=AVERAGEIFS(B1:B14,B1:B14,">0")',
    '=MAXIFS(B1:B14,A1:A14,"東京*")', '=MINIFS(B1:B14,A1:A14,"東京*")', '=MAXIFS(B1:B14,A1:A14,"none")', '=MINIFS(B1:B14,A1:A14,"none")', '=MAXIFS(B1:B14,B1:B14,"<1000")',
    '=MINIFS(B1:B14,A1:A14,"<>東京",B1:B14,">300")', '=MAXIFS(B1:B14,A1:A14,"東京",C1:C3,">0")', '=MAXIFS(A1:A14,B1:B14,">0")',
  ],
};

const lookup = {
  name: '検索',
  cells: {
    A1: 1, A2: 3, A3: 5, A4: 7, A5: 9, A6: 11, A7: 13, A8: 15, B1: 'りんご', B2: 'みかん', B3: 'ぶどう', B4: 'もも', B5: 'Apple', B6: 'banana', B7: 'Cherry',
    C1: 100, C2: 80, C3: 300, C4: 250, C5: 120, C6: 90, C7: 400, C8: 50, D1: true, D2: 'x', D3: 3.5, D4: {f: '=NA()'}, D5: {f: '=""'}, D6: 7, D7: '終わり', D8: 0,
    E1: 30, E2: 20, E3: 10, F1: 5, F2: 1, F3: 9, F4: 3, F5: 7, F6: 2, G1: 'b', G2: 'A', G3: 'c', G4: 'a*', G5: 'x?', G6: '~',
    A10: 1, B10: 3, C10: 5, D10: 7, E10: 9, F10: 11, G10: 13, A11: '一', B11: '三', C11: '五', D11: '七', E11: '九', F11: '十一', G11: '十三', A12: 10, B12: 30, C12: 50, D12: 70, E12: 90,
    F12: 110, G12: 130, A14: 'apple', A15: 'banana', A16: 'cherry', A17: 'date', A18: 'fig', B14: 1, B15: 2, B16: 3, B17: 4, B18: 5, A20: '2', A21: 2, A22: true,
    I1: 1, I2: 2, I3: 2, I4: 2, I5: 3, I7: 1, I8: 2, I9: 2, I10: 2, I11: 2, I12: 2, I13: 3, I15: 10, I16: 30, I17: 20, I18: 50, I19: 40, I20: 70, I21: 60,
  },
  placed: {J1: '=ROW()', J5: '=ROW()', K3: '=COLUMN()', J3: '=ROW(B2:B4)', J10: '=ROW(B2:B4)', K10: '=COLUMN(B2:D2)', L2: '=COLUMN(B2:D2)', J2: '=INDEX(A1:C8,0,2)',
    M4: '=INDEX(A1:C8,2,0)', B9: '=INDEX(A1:C8,0,2)'},
  formulas: [
    '=VLOOKUP(5,A1:C8,2,FALSE)', '=VLOOKUP(4,A1:C8,2,FALSE)', '=VLOOKUP(4,A1:C8,2,TRUE)', '=VLOOKUP(4,A1:C8,2)', '=VLOOKUP(0,A1:C8,2,TRUE)', '=VLOOKUP(100,A1:C8,3,TRUE)',
    '=VLOOKUP(15,A1:C8,2,0)', '=VLOOKUP("apple",B1:C8,2,FALSE)', '=VLOOKUP("App*",B1:C8,2,FALSE)', '=VLOOKUP("?anana",B1:C8,2,FALSE)', '=VLOOKUP(5,A1:C8,4,FALSE)',
    '=VLOOKUP(5,A1:C8,0,FALSE)', '=VLOOKUP(5,A1:C8,"2",FALSE)', '=VLOOKUP(5,A1:C8,2.9,FALSE)', '=VLOOKUP("5",A1:C8,2,FALSE)', '=VLOOKUP(5,A1:C8,2,"FALSE")', '=VLOOKUP(5,A1:C8,2,)',
    '=VLOOKUP(7,A1:D8,4,FALSE)', '=VLOOKUP(9,A1:D8,4,FALSE)', '=VLOOKUP(13,A1:D8,4,FALSE)', '=VLOOKUP(2,A20:B22,1,FALSE)', '=VLOOKUP("2",A20:B22,1,FALSE)', '=VLOOKUP(TRUE,A20:B22,1,FALSE)',
    '=VLOOKUP(A9,A1:C8,2,FALSE)', '=VLOOKUP(NA(),A1:C8,2,FALSE)', '=VLOOKUP(5,A1:C8,-1,FALSE)', '=VLOOKUP("*",B1:C8,2,FALSE)', '=VLOOKUP("b*",B1:C8,2,FALSE)', '=VLOOKUP("みかん",B1:C8,2,FALSE)',
    '=VLOOKUP("a~*",G1:G6,1,FALSE)', '=VLOOKUP("x~?",G1:G6,1,FALSE)', '=VLOOKUP("~~",G1:G6,1,FALSE)', '=VLOOKUP("a",G1:G6,1,FALSE)', '=VLOOKUP("a*",G1:G6,1,FALSE)',
    ...[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].flatMap(x => [`=VLOOKUP(${x},F1:G6,2,TRUE)`, `=MATCH(${x},F1:F6,1)`]),
    '=VLOOKUP("c",A14:B18,2,TRUE)', '=VLOOKUP("coconut",A14:B18,2)', '=VLOOKUP("zzz",A14:B18,2)', '=VLOOKUP("a",A14:B18,2)', '=VLOOKUP(5,A14:B18,2)', '=VLOOKUP("CHERRY",A14:B18,2,TRUE)',
    '=HLOOKUP(5,A10:G12,2,FALSE)', '=HLOOKUP(6,A10:G12,3,TRUE)', '=HLOOKUP(5,A10:G12,4,FALSE)', '=HLOOKUP(0,A10:G12,2)', '=HLOOKUP(20,A10:G12,3)', '=HLOOKUP(5,A10:G12,2)',
    '=MATCH(5,A1:A8,0)', '=MATCH(6,A1:A8)', '=MATCH(6,A1:A8,1)', '=MATCH(0,A1:A8)', '=MATCH(20,E1:E3,-1)', '=MATCH(25,E1:E3,-1)', '=MATCH(35,E1:E3,-1)', '=MATCH(5,E1:E3,-1)',
    '=MATCH("APPLE",B1:B8,0)', '=MATCH("*an*",B1:B8,0)', '=MATCH("a~*",G1:G6,0)', '=MATCH("x~?",G1:G6,0)', '=MATCH("~~",G1:G6,0)', '=MATCH("cherry",A14:A18)', '=MATCH("coconut",A14:A18)',
    '=MATCH(5,A1:B2,0)', '=MATCH(TRUE,D1:D8,0)', '=MATCH(5,A1:A8,2)', '=MATCH(5,A1:A8,-0.5)', '=MATCH(NA(),A1:A8,0)', '=MATCH(5,A10:G10,0)', '=MATCH(12,A10:G10)', '=MATCH("",D1:D8,0)',
    '=MATCH(0,D1:D8,0)', '=MATCH(7,D1:D8,0)', '=MATCH("x",D1:D8,0)', '=MATCH(2,A20:A22,0)', '=MATCH("2",A20:A22,0)', '=MATCH(16,A1:A8)', '=MATCH(1,A1:A8,-1)', '=MATCH(5,A1:A8,"0")',
    '=INDEX(A1:C8,3,2)', '=INDEX(A1:C8,3)', '=INDEX(A1:A8,3)', '=INDEX(A10:G10,3)', '=INDEX(A1:C8,9,1)', '=SUM(INDEX(A1:C8,0,3))', '=SUM(INDEX(A1:C8,2,0))', '=INDEX(A1:C8,8,2)',
    '=INDEX(A1:C8,"2",1)', '=INDEX(A1:C8,1.9,1)', '=INDEX(A1:C8,-1,1)', '=COUNTA(INDEX(A1:D8,0,4))', '=INDEX(A1:C8,2,4)', '=INDEX(D1:D8,4)', '=INDEX(A1:A8,0)', '=INDEX(A1:C8,0,0)',
    '=SUM(INDEX(A1:C8,0,0))', '=INDEX(A1:C8,,2)', '=INDEX(A1:C8,2,)', '=INDEX(A1:A8,2,1)', '=INDEX(A1:A8,1,2)', '=INDEX(A10:G10,1,3)', '=INDEX(A10:G10,2)', '=INDEX(C1:C8,MATCH("もも",B1:B8,0))',
    '=XLOOKUP(5,A1:A8,B1:B8)', '=XLOOKUP(4,A1:A8,B1:B8)', '=XLOOKUP(4,A1:A8,B1:B8,"なし")', '=XLOOKUP(4,A1:A8,B1:B8,,-1)', '=XLOOKUP(4,A1:A8,B1:B8,,1)', '=XLOOKUP(16,A1:A8,B1:B8,,1)',
    '=XLOOKUP("app*",B1:B8,C1:C8,,2)', '=XLOOKUP("app*",B1:B8,C1:C8)', '=XLOOKUP(3,F1:F6,G1:G6,,0,-1)', '=XLOOKUP(7,A1:A8,C1:C8,,0,2)', '=XLOOKUP(20,E1:E3,E1:E3,,0,-2)',
    '=XLOOKUP(15,A1:A8,B1:B8)', '=XLOOKUP(5,A1:A8,B1:B7)', '=XLOOKUP(5,A10:G10,A11:G11)', '=XLOOKUP("APPLE",B1:B8,A1:A8)', '=XLOOKUP(5,A1:A8,B1:B8,"なし",3)', '=XLOOKUP(2,F1:F6,G1:G6)',
    '=XLOOKUP(0,A1:A8,B1:B8,,-1)', '=XLOOKUP(100,A1:A8,B1:B8,,-1)', '=XLOOKUP(8,F1:F6,G1:G6,,-1)', '=XLOOKUP(8,F1:F6,G1:G6,,1)', '=XLOOKUP(9,A1:A8,D1:D8)', '=XLOOKUP(7,A1:A8,D1:D8)',
    '=XLOOKUP(NA(),A1:A8,B1:B8)', '=XLOOKUP(4,A1:A8,B1:B8,1/0)', '=XLOOKUP("?",G1:G6,F1:F6,,2)', '=XLOOKUP("a~*",G1:G6,F1:F6,,2)', '=XLOOKUP(6,A1:A8,B1:B8,,-1,2)', '=XLOOKUP(6,A1:A8,B1:B8,,1,2)',
    '=XLOOKUP(5,A1:A8,B1:B8,,0,3)', '=XLOOKUP(2,A20:A22,A20:A22)', '=XLOOKUP("2",A20:A22,A20:A22)',
    '=XMATCH(5,A1:A8)', '=XMATCH(4,A1:A8,-1)', '=XMATCH(4,A1:A8,1)', '=XMATCH("*an*",B1:B8,2)', '=XMATCH(3,F1:F6,0,-1)', '=XMATCH(9,A1:A8,0,2)', '=XMATCH(4,A1:A8)', '=XMATCH(25,E1:E3,1,-2)',
    '=XMATCH(7,A10:G10)', '=XMATCH("CHERRY",A14:A18)',
    '=LOOKUP(6,A1:A8,B1:B8)', '=LOOKUP(6,A1:B8)', '=LOOKUP(0,A1:A8,B1:B8)', '=LOOKUP(100,A1:A8,C1:C8)', '=LOOKUP("c",A14:A18,B14:B18)', '=LOOKUP(6,A10:G11)', '=LOOKUP(4,F1:F6,G1:G6)',
    '=LOOKUP(6,A1:A8,A10:G10)', '=LOOKUP("zz",A14:A18)',
    '=ROW(C7)', '=COLUMN(C7)', '=ROWS(A1:C8)', '=COLUMNS(A1:C8)', '=ROWS(A1)', '=ROW()+ROW()', '=COLUMN()', '=COLUMNS(A10:G12)', '=ROW(A1:A3)', '=SUM(ROW(A1:A3))',
    // MATCH -1 searches from the first value on; MATCH 1 and LOOKUP search by halves and go on over equal values.
    ...[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(x => `=MATCH(${x},F1:F6,-1)`), ...[0, 1, 2, 15, 16].map(x => `=MATCH(${x},A1:A8,-1)`),
    ...[0, 10, 15, 30, 80].map(x => `=MATCH(${x},I15:I21,-1)`), ...[1, 2, 3].flatMap(x => [`=MATCH(${x},I1:I5)`, `=MATCH(${x},I7:I13,1)`, `=LOOKUP(${x},I1:I5,A1:A5)`]),
    '=VLOOKUP(2,I1:I5,1)', '=MATCH(6,A1:A8,0.5)', '=MATCH(6,A1:A8,-0.5)', '=MATCH(6,A1:A8,-1.5)', '=MATCH(25,E1:E3,-0.5)', '=MATCH(20,E1:E3,-0.1)',
    // XMATCH and XLOOKUP search by halves with search mode 2 (ascending) and -2 (descending).
    ...[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].flatMap(x => [0, -1, 1].flatMap(m => [`=XMATCH(${x},F1:F6,${m},2)`, `=XMATCH(${x},F1:F6,${m},-2)`])),
    ...[0, 1, 5, 15, 16].flatMap(x => [`=XMATCH(${x},A1:A8,-1,-2)`, `=XMATCH(${x},A1:A8,1,-2)`]), '=XLOOKUP(8,F1:F6,G1:G6,"なし",-1,-2)', '=XLOOKUP(8,F1:F6,G1:G6,"なし",1,2)',
  ],
};

const texts = {
  name: '文字列',
  cells: {A1: 'Excel 関数', A3: 12345, A4: true, A5: 'abc def', A6: '  a  b  ', A7: 'ＡＢＣ', A8: '𩸽の刺身', A9: 'a-b-c-d', A10: {f: '=NA()'}, A11: {f: '=1/3'}, A12: 'Hello World',
    A13: '東京都千代田区', A14: 'a*b?c~d', A15: {f: '=""'}, A16: 'ｶﾞｷﾞ', B1: 2, B2: 'ab', B3: '１２３'},
  formulas: [
    ...each(['LEN'], ['A1', 'A2', 'A3', 'A4', 'A8', 'A11', 'A15', '""', '" a "', '123.45', '1/3', 'TRUE', '"😀"', 'A10', '-0.5', 'A16', '1E+21', '1E-7']),
    ...['A1', 'A1,0', 'A1,5', 'A1,100', 'A1,-1', 'A1,"2"', 'A1,2.9', 'A3,2', 'A4,2', 'A8,1', 'A8,2', 'A2,1', 'A10,1', 'A1,A2', 'A11,4', 'A1,TRUE', '1234.5,3', 'A16,1']
      .flatMap(a => [`=LEFT(${a})`, `=RIGHT(${a})`]),
    ...['A1,1,5', 'A1,0,1', 'A1,20,1', 'A1,2,0', 'A1,2,-1', 'A1,1.9,2.9', 'A3,2,3', 'A8,1,1', 'A8,2,2', 'A1,7,100', 'A1,"2","3"', 'A4,2,2', 'A10,1,1', 'A1,1,0', 'A2,1,1', 'A5,4,1']
      .map(a => `=MID(${a})`),
    ...['"a",A5', '"A",A5', '"d",A5,5', '"d",A5,7', '"z",A5', '"",A5', '"",A5,3', '"",A5,8', '"",A5,9', '"",""', '"都",A13', '"c",A5,0', '"b",A12', '"o",A12,6', '"o",A12,9', '"?",A14',
      '"*",A14', '1,A3', '"345",A3', '"E",A1', '"関",A1', '"𩸽",A8', '"の",A8']
      .flatMap(a => [`=FIND(${a})`, `=SEARCH(${a})`]),
    '=SEARCH("b?d",A9)', '=SEARCH("b*d",A9)', '=SEARCH("~?",A14)', '=SEARCH("~*",A14)', '=SEARCH("~~",A14)', '=SEARCH("e*o",A12)', '=SEARCH("*",A12)', '=SEARCH("w",A12,8)',
    ...['A9,"-","+"', 'A9,"-","+",2', 'A9,"-","+",0', 'A9,"-","+",5', 'A9,"x","+"', 'A9,"","+"', '"aaa","aa","b"', '"aaa","aa","b",2', 'A3,"3","x"', 'A9,"-",""', 'A9,"-","+","2"',
      'A12,"o","0"', 'A12,"O","0"', 'A10,"a","b"', 'A2,"a","b"', 'A9,"-",A2', '"東京都東京","東京","大阪",2']
      .map(a => `=SUBSTITUTE(${a})`),
    ...['A12,1,5,"Bye"', 'A12,7,5,"Excel"', 'A12,20,1,"!"', 'A12,0,1,"x"', 'A12,1,0,"x"', 'A12,1,-1,"x"', 'A3,2,3,"x"', 'A12,6,1,""', 'A13,1,2,"大阪"', 'A12,1.9,1,"x"', 'A12,12,0,"!"']
      .map(a => `=REPLACE(${a})`),
    ...each(['TRIM'], ['A6', '"  a   b  "', '"　a　　b　"', '"a" & CHAR(9) & "b"', 'A1', 'A3', 'A2', 'A10', '"a" & UNICHAR(160) & " b"', '" "', '"a  "&"  b"']),
    ...each(['CLEAN'], ['"a"&CHAR(10)&"b"', '"a"&CHAR(9)&"b"&CHAR(13)', '"a"&UNICHAR(127)&"b"', 'A1', 'A3', '"a"&UNICHAR(129)&"b"']),
    ...each(['UPPER', 'LOWER', 'PROPER'], ['"hello World"', 'A7', '"ａｂｃ"', '"éÀ"', 'A13', 'A3', 'A4', 'A2', '"o\'neil"', '"abc-def ghi"', '"123abc"', '"日本abc"', '"ﾊﾛｰabc"',
      '"ÀB"', '"σς"', 'A10', '"x1y"', '"é1e"']),
    ...['"ab",3', '"a",0', '"a",-1', '"a",2.9', '"abc",11000', '"",5', 'A3,2', 'A4,2', '"あ","3"', 'A10,2', '"ab",A2'].map(a => `=REPT(${a})`),
    ...['"a","A"', '"a","a"', '1,"1"', 'TRUE,"TRUE"', 'A2,""', '"ａ","a"', 'A10,"a"', '"a ","a"', '1/3,"0.333333333333333"'].map(a => `=EXACT(${a})`),
    ...each(['CHAR'], ['65', '10', '0', '256', '33440', '130', '177', '"66"', '65.9', '255', '128', '160', '223', '224', '33088', '-1', '40959', '65535', '65536']),
    ...each(['CODE'], ['"A"', '"あ"', '"ｱ"', '""', '"𩸽"', '"€"', '"ABC"', '1', 'TRUE', 'A2', '"①"', '"髙"', '"〜"', '"～"', '"¥"', '"\\"', '"ä"', '"Ⅰ"']),
    ...each(['UNICHAR'], ['9731', '0', '65', '171581', '1114112', '"12354"', '55296']),
    ...each(['UNICODE'], ['"あ"', '"𩸽"', '""', '"A"', '1', 'TRUE']),
    // DBCS is the function the Japanese Excel shows as JIS (a formula set through COM, as here, takes the English name).
    ...each(['ASC', 'DBCS'], ['"ＡＢＣ１２３"', '"アイウ"', '"ガ"', '"パ"', '"ー"', '"、。「」"', '"あ"', '"　"', '"ヴ"', '"ヵ"', '"ABC123"', '"ｶﾞ"', '"ﾊﾟ"', '"ｳﾞ"', '" "', '"~"',
      '"\\"', '"￥"', '"¥"', '"’"', '"\'"', 'UNICHAR(65282)', '""""', '123', 'TRUE', 'A2', 'A16', '"ﾞ"', '"ヰ"', '"ァ"', '"ｧ"', '"〜"', '"～"', '"－"', '"-"', '"−"', '"ー-"', '"A1"', 'A10', '"ヷ"', '"ゝ"',
      '"・"', '"￣"', '"‐"', '"｀"', '"`"', '"ﾟ"', '"ゔ"']),
  ],
};

// More text functions on the same cells (a sheet holds at most 500 rows that the page reads from xlsx).
const texts2 = {
  name: '文字列2', cells: texts.cells,
  formulas: [
    ...each(['VALUE'], ['"123"', '"abc"', '""', 'TRUE', '1', 'A3', 'B3', 'A2', '" 1,234 "', '"10%"', '"2026/1/2"', '"12:30"', '"¥1,000"', '"(5)"', '"1e3"', '"１２３"', '"令和8年1月2日"', 'A10',
      '"2026/1/2 12:30"', '"1 1/2"']),
    ...[['1234.567', '"0"'], ['1234.567', '"0.00"'], ['1234567.891', '"#,##0"'], ['1234567.891', '"#,##0.00"'], ['0.1234', '"0%"'], ['0.1234', '"0.0%"'], ['1234.5', '"#"'], ['1234.5', '"#.##"'],
      ['12', '"0000"'], ['1.5', '"00.0"'], ['1234', '"#,##0円"'], ['1234', '"\\#,##0"'], ['1234', '"¥#,##0"'], ['1234', '"$#,##0"'], ['1234567', '"#,##0,"'], ['1234567890', '"#,##0,,"'],
      ['12345.678', '"0.00E+00"'], ['0.000123', '"0.0E+0"'], ['-1234', '"#,##0;(#,##0)"'], ['-1234', '"#,##0;-#,##0"'], ['-1234', '"#,##0;[赤]-#,##0"'], ['0', '"0;-0;ゼロ"'],
      ['5', '"0;-0;ゼロ"'], ['-5', '"0;-0;ゼロ"'], ['"abc"', '"0;-0;0;@です"'], ['1500', '"[>=1000]#,##0,""千"";0"'], ['500', '"[>=1000]#,##0,""千"";0"'], ['"abc"', '"@"'],
      ['12', '"""金額:""0"'], ['1234.5', '"G/標準"'], ['1234.5', '"General"'], ['0.5', '"0.0?"'], ['2.5', '"0"'], ['-2.5', '"0"'], ['0.005', '"0.00"'], ['1E+16', '"0"'], ['TRUE', '"0"'],
      ['"abc"', '"0"'], ['-0.5', '"0"'], ['0', '"#"'], ['0', '"#,##0"'], ['-0.4', '"0"'], ['1234.5', '"#,##0.0_);(#,##0.0)"'], ['1/3', '"0.000000000000000000"'], ['123', '"000-0000"'],
      ['1234567', '"0000-000"'], ['46024', '"yyyy/mm/dd"'], ['46024', '"yyyy/m/d"'], ['46024', '"yy/mm/dd"'], ['46024', '"yyyy年m月d日"'], ['46024', '"m月d日"'], ['46024', '"mm/dd"'],
      ['46024', '"dd"'], ['46024', '"d"'], ['46024', '"ddd"'], ['46024', '"dddd"'], ['46024', '"aaa"'], ['46024', '"aaaa"'], ['46024', '"mmm"'], ['46024', '"mmmm"'], ['46024', '"mmmmm"'],
      ['46024', '"ge.m.d"'], ['46024', '"gge年m月d日"'], ['46024', '"ggge年m月d日"'], ['46024', '"e"'], ['46024', '"ee"'], ['46024.5208333333', '"yyyy-mm-dd hh:mm:ss"'], ['0.5208333333', '"h:mm"'],
      ['0.0423611111', '"hh:mm"'], ['0.5213541667', '"h:mm:ss"'], ['0.5625', '"h:mm AM/PM"'], ['0.0625', '"h:mm:ss AM/PM"'], ['1.5', '"[h]:mm"'], ['0.0625', '"[mm]"'], ['0.0001', '"[ss]"'],
      ['0.0423611111', '"mm:ss"'], ['0.00001', '"s.00"'], ['0.5208333333', '"h時mm分"'], ['46024', '"yyyy/mm/dd(aaa)"'], ['46024', '"m/d/yyyy"'], ['"2026/1/2"', '"yyyy年m月d日"'],
      ['46024.75', '"yyyy/mm/dd h:mm"'], ['0.5625', '"上午/下午 h:mm"'], ['0.5625', '"AM/PM h:mm"'], ['0.5625', '"午前/午後h:mm"'], ['43585', '"ggge年m月d日"'], ['43586', '"ggge年m月d日"'],
      ['32515', '"gge"'], ['1', '"yyyy/mm/dd"'], ['60', '"yyyy/mm/dd"'], ['61', '"yyyy/mm/dd"'], ['0', '"yyyy/mm/dd"'], ['46024', '"[$-ja-JP]ggge年m月d日"'], ['46024', '"[$-411]ggge"'],
      ['0.999999999', '"h:mm:ss"'], ['0.5', '"hh"'], ['46024', '"yyyymmdd"'], ['46024', '"d-mmm-yy"'], ['46024', '"m"'], ['0.5', '"m"'], ['0.5', '"h:m"'], ['46024', '"ggggggge"'],
      ['-1', '"yyyy/mm/dd"'], ['A2', '"0"'], ['A10', '"0"'], ['1234.5', '""'], ['1234.5', '"0.00;"'], ['0.1', '"0.0%%"'], ['7', '"0 ""個"""'], ['7', '"0個"'], ['12.3', '"???.??"'],
      ['3.75', '"# ?/?"'], ['0.333', '"?/?"'],
      // At most 255 characters; b is the Buddhist year; B, n and N are no format characters; every section must be right;
      // ! shows the next character.
      ['1E+254', '"0"'], ['1E+255', '"0"'], ['REPT("a",255)', '"@"'], ['REPT("a",256)', '"@"'], ['46024', '"bbbb"'], ['46024', '"b"'], ['1', '"b"'], ['46024', '"n"'],
      ['1', '"0;B"'], ['1', '"0;N"'], ['1', '"0b"'], ['1', '"0;0y"'], ['-1', '"0;0x"'], ['1', '"0;0;0;@y"'], ['1', '"0;yyyy"'], ['1', '"0!x"'], ['1', '"0!!"'], ['1', '"0!"'],
      ['1', '"0x"'], ['1', '"0 kg"']]
      .map(([v, f]) => `=TEXT(${v},${f})`),
    // LEN, MID, FIND and SEARCH count UTF-16 units, LEFT and RIGHT characters; ? of SEARCH is one unit.
    ...['LEFT("𩸽a",1)', 'LEFT("a𩸽b",2)', 'LEFT("𩸽𩸽",3)', 'RIGHT("a𩸽",1)', 'RIGHT("ab𩸽",2)', 'RIGHT("𩸽𩸽",1)', 'LEN(LEFT("𩸽𩸽",2))', 'MID("a𩸽b",3,2)', 'LEN("😀")',
      'REPLACE("a𩸽b",3,1,"x")', 'FIND("b","a𩸽b")', 'SEARCH("?b","a𩸽b")', 'SEARCH("a?b","a𩸽b")', 'PROPER("𩸽abc")', 'UPPER("𩸽a")', 'SUBSTITUTE("a𩸽b","𩸽","x")',
      // SEARCH may not start after the last character (FIND may); a ~ at the end of a pattern stands for nothing.
      'SEARCH("","abc",3)', 'SEARCH("","abc",4)', 'FIND("","abc",4)', 'SEARCH("","a",2)', 'SEARCH("~","a~b")', 'SEARCH("*~","a~b")', 'SEARCH("b*","abcabc",3)',
      // SUBSTITUTE counts its occurrences overlapping.
      'SUBSTITUTE("aaaa","aa","b",2)', 'SUBSTITUTE("aaaa","aa","b",3)', 'SUBSTITUTE("ababab","aba","x",2)', 'SUBSTITUTE("aaaaa","aa","b")', 'SUBSTITUTE("aAa","a","x",2)']
      .map(e => `=${e}`),
  ],
};

const dates = {
  name: '日付',
  cells: {A1: {v: 46024, fmt: 'yyyy/mm/dd'}, A3: {v: 46053, fmt: 'yyyy/mm/dd'}, A4: {v: 46027, fmt: 'yyyy/mm/dd'}, A5: {v: 46035, fmt: 'yyyy/mm/dd'}, A6: {v: 46036, fmt: 'yyyy/mm/dd'},
    A7: 'abc', A8: {v: 0.5, fmt: 'h:mm'}, A9: '2026/1/10', A10: {f: '=NA()'}, A11: true},
  formulas: [
    ...['2026,1,2', '2026,13,1', '2026,0,1', '2026,1,0', '2026,2,29', '1900,2,29', '1900,3,0', '26,1,1', '1899,12,31', '-1,1,1', '10000,1,1', '9999,12,31', '"2026","1","2"',
      '2026.9,1.9,2.9', '2026,-1,1', '1900,1,1', '1900,1,0', '1900,2,30', '2026,1,-1', '1900,0,1', '0,1,1', '2024,2,29', '2023,2,29', '9999,12,32', '2026,1,"x"', 'A10,1,1', '1904,1,1',
      '2026,24,31', '1,1,1', '1899,1,1', '1900,1,-1']
      .map(a => `=DATE(${a})`),
    ...each(['YEAR', 'MONTH', 'DAY'], ['46024', '60', '59', '61', '1', '0', '-1', '2958465', '2958466', '"2026/1/2"', '46024.99', 'TRUE', 'A2', 'A7', 'A10', '"abc"', 'A8']),
    ...['46024', '46024,2', '46024,3', '46024,11', '46024,12', '46024,13', '46024,14', '46024,15', '46024,16', '46024,17', '46024,4', '1', '0', '60', '61', '"2026/1/2"', '46024.9', '-1',
      '46024,1.5', '46024,"2"'].map(a => `=WEEKDAY(${a})`),
    ...['46024', '46024,2', '46024,21', '46024,11', '46024,12', '46024,13', '46024,14', '46024,15', '46024,16', '46024,17', '46024,3', 'DATE(2026,12,31)', 'DATE(2026,12,31),21',
      'DATE(2027,1,1),21', 'DATE(2027,1,1)', 'DATE(2027,1,4),2', '1', '0'].map(a => `=WEEKNUM(${a})`),
    ...each(['ISOWEEKNUM'], ['46024', 'DATE(2027,1,1)', 'DATE(2026,12,28)', 'DATE(2021,1,3)', 'DATE(2020,12,31)', '"2026/1/2"', '1']),
    ...['DATE(2026,1,31),1', 'DATE(2026,3,31),-1', 'DATE(2024,1,31),1', '46024,1.9', '"2026/1/31",1', '46024,-1000000', '46024,0', 'DATE(2026,1,31),-1.9', 'A7,1', '46024.7,1', '61,-1']
      .flatMap(a => [`=EDATE(${a})`, `=EOMONTH(${a})`]),
    '=EOMONTH(DATE(2024,2,10),0)', '=EOMONTH(DATE(2026,12,15),1)',
    ...[['DATE(2020,1,15)', 'DATE(2026,1,14)'], ['DATE(2020,1,15)', 'DATE(2026,1,15)'], ['DATE(2026,1,31)', 'DATE(2026,3,1)'], ['DATE(2026,1,31)', 'DATE(2026,2,28)'],
      ['DATE(2024,2,29)', 'DATE(2025,2,28)'], ['DATE(2024,2,29)', 'DATE(2025,3,1)'], ['DATE(2026,3,1)', 'DATE(2026,1,1)'], ['DATE(2026,1,1)', 'DATE(2026,1,1)'], ['DATE(2025,12,31)', 'DATE(2026,1,1)']]
      .flatMap(([a, b]) => ['"Y"', '"M"', '"D"', '"MD"', '"YM"', '"YD"'].map(u => `=DATEDIF(${a},${b},${u})`)),
    '=DATEDIF(DATE(2020,1,1),DATE(2026,1,1),"y")', '=DATEDIF(DATE(2020,1,1),DATE(2026,1,1),"X")', '=DATEDIF("2026/1/1","2026/2/1","D")', '=DATEDIF(46024.9,46025.1,"D")',
    '=DAYS(DATE(2026,1,10),DATE(2026,1,1))', '=DAYS("2026/1/10","2026/1/1")', '=DAYS(46024.9,46024.1)', '=DAYS(1,2)', '=DAYS("a",1)', '=DAYS(-1,1)',
    '=NETWORKDAYS(DATE(2026,1,1),DATE(2026,1,31))', '=NETWORKDAYS(DATE(2026,1,1),DATE(2026,1,31),A4:A6)', '=NETWORKDAYS(DATE(2026,1,31),DATE(2026,1,1))', '=NETWORKDAYS(46024,46024)',
    '=NETWORKDAYS(DATE(2026,1,3),DATE(2026,1,4))', '=NETWORKDAYS(46024.9,46030.1)', '=NETWORKDAYS(DATE(2026,1,1),DATE(2026,1,31),A4)', '=NETWORKDAYS(DATE(2026,1,1),DATE(2026,1,31),A7)',
    ...['1', '2', '7', '11', '17', '"0000011"', '"1111111"', '"0000000"', '"1000001"', '8', '"000001"', '"0000012"']
      .map(w => `=NETWORKDAYS.INTL(DATE(2026,1,1),DATE(2026,1,31),${w})`),
    '=NETWORKDAYS.INTL(DATE(2026,1,1),DATE(2026,1,31),1,A4:A6)', '=NETWORKDAYS.INTL(DATE(2026,1,1),DATE(2026,1,31))',
    '=WORKDAY(DATE(2026,1,2),5)', '=WORKDAY(DATE(2026,1,2),-5)', '=WORKDAY(DATE(2026,1,2),5,A4:A6)', '=WORKDAY(DATE(2026,1,3),0)', '=WORKDAY(DATE(2026,1,2),0.9)', '=WORKDAY(DATE(2026,1,2),1.9)',
    '=WORKDAY(DATE(2026,1,3),1)', '=WORKDAY(46024.5,1)', '=WORKDAY(DATE(2026,1,2),-1.5)', '=WORKDAY(DATE(2026,1,2),30,A4:A6)', '=WORKDAY(A7,1)',
    ...['1', '2', '11', '"0000011"', '"0000000"', '"1111111"'].map(w => `=WORKDAY.INTL(DATE(2026,1,2),5,${w})`), '=WORKDAY.INTL(DATE(2026,1,2),-3,"0000011",A4:A6)',
    ...['12,30,0', '25,0,0', '0,-1,0', '12,90,0', '0,0,86400', '-1,0,0', '"12","30","0"', '0,0,0', '23,59,59', '0,0,-1', '32767,0,0', '1.9,1.9,1.9', '24,0,0'].map(a => `=TIME(${a})`),
    ...each(['HOUR', 'MINUTE', 'SECOND'], ['0.5', '0.5213541666666667', '46024.75', '0.99999999', '"12:30:45"', 'TRUE', '-0.5', 'A2', '0.999994', '0.000005787', '0.0000057', '"abc"', '1.5',
      '0.4999999999', 'A8', '2958466']),
    ...each(['DATEVALUE'], ['"2026/1/2"', '"2026/1/2 12:30"', '"12:30"', '"abc"', '46024', '"令和8年1月2日"', 'A9', '"2026-01-02"', '"1900/2/29"', 'A1']),
    ...each(['TIMEVALUE'], ['"12:30"', '"2026/1/2 12:30"', '"25:00"', '"1:30 PM"', '"abc"', '0.5', '"12:30:45.5"', '"2026/1/2"', '"12時30分"']),
    '=TODAY()=INT(NOW())', '=NOW()-TODAY()<1', '=YEAR(TODAY())>=2026', '=TODAY()-TODAY()', '=ISNUMBER(NOW())', '=TODAY()>DATE(2026,1,1)',
    // DATEDIF "YD" around leap days (the rule is in formulas.js).
    ...[['2021,1,15', '2026,1,14'], ['2020,3,1', '2021,2,28'], ['2019,3,1', '2020,2,29'], ['2020,2,28', '2021,3,1'], ['2024,1,1', '2025,12,31'], ['2021,2,28', '2024,2,29'],
      ['2024,2,29', '2028,2,28'], ['2023,3,1', '2024,2,29'], ['2019,1,1', '2020,12,31'], ['2024,3,31', '2025,2,28'], ['2023,12,31', '2024,3,1'], ['2024,12,31', '2025,3,1'],
      ['2024,1,1', '2025,3,1'], ['2024,1,1', '2026,12,31'], ['2024,2,28', '2025,12,31'], ['2024,2,1', '2025,3,1'], ['2020,1,31', '2021,3,1'], ['2024,1,15', '2025,3,1'],
      ['2023,1,1', '2024,3,1'], ['2024,2,10', '2025,3,10'], ['2024,1,10', '2025,3,20'], ['2024,2,29', '2025,12,31'], ['2024,2,29', '2026,3,1'], ['2024,2,29', '2028,3,1'],
      ['2020,2,28', '2024,3,1'], ['2019,2,28', '2020,3,1'], ['2019,2,1', '2020,3,1'], ['2024,12,1', '2025,12,31']]
      .map(([a, b]) => `=DATEDIF(DATE(${a}),DATE(${b}),"YD")`),
    // EDATE and EOMONTH keep to the calendar, where 1900/02 has 28 days.
    ...['DATE(1900,1,29),1', 'DATE(1900,1,31),1', '1,1', '60,1', '60,0', '60,12', '59,1', 'DATE(1900,3,31),-1', 'DATE(1900,2,28),0', '60,-1'].flatMap(a => [`=EDATE(${a})`, `=EOMONTH(${a})`]),
    // WORKDAY takes the days down to a whole number; the weekend numbers 11 to 17 are one day, Sunday to Saturday.
    ...['46024,1.5', '46024,-0.5', '46024,-2.5', '46024,-0.9', '46025,-1', '46026,1', '46025,0'].map(a => `=WORKDAY(${a})`), '=WORKDAY.INTL(46024,-1.5,1)',
    ...[11, 12, 13, 14, 15, 16, 17].flatMap(k => [`=NETWORKDAYS.INTL(DATE(2026,1,1),DATE(2026,1,31),${k})`, `=WORKDAY.INTL(DATE(2026,1,2),5,${k})`]),
    '=NETWORKDAYS.INTL(DATE(2026,1,2),DATE(2026,1,9),"1111111")', '=WORKDAY.INTL(DATE(2026,1,2),0,"1111111")', '=WORKDAY.INTL(DATE(2026,1,2),1,"1111111")',
    // The functions that came from the Analysis ToolPak take no TRUE or FALSE, not even from a cell.
    ...['EDATE(TRUE,1)', 'EOMONTH(46024,TRUE)', 'WEEKNUM(TRUE)', 'WEEKNUM(46024,TRUE)', 'WORKDAY(TRUE,1)', 'WORKDAY(46024,TRUE)', 'WORKDAY(46024,1,TRUE)', 'NETWORKDAYS(TRUE,46024)',
      'NETWORKDAYS(46024,46030,TRUE)', 'WORKDAY.INTL(TRUE,1)', 'NETWORKDAYS.INTL(46024,TRUE)', 'EDATE(A11,1)', 'EDATE("46024",1)', 'EDATE(46024,"1")', 'ISOWEEKNUM(TRUE)', 'DAYS(TRUE,1)',
      'DATEDIF(TRUE,46024,"D")', 'WEEKDAY(TRUE)', 'NETWORKDAYS(46024,46030,A7)'].map(e => `=${e}`),
  ],
};

const subtotals = {
  name: '小計',
  cells: {A1: 1, A2: 2, A3: 3, A4: {f: '=SUBTOTAL(9,A1:A3)'}, A5: 10, A6: {f: '=SUM(A1:A3)'}, A7: {f: '=SUBTOTAL(109,A5)'}, B1: 'x', B2: true, B3: '5', C1: {f: '=1/0'}, C2: 4},
  formulas: [
    ...[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111].map(n => `=SUBTOTAL(${n},A1:A7)`),
    '=SUBTOTAL(9,A1:A3,A5)', '=SUBTOTAL(12,A1:A3)', '=SUBTOTAL(0,A1:A3)', '=SUBTOTAL("9",A1:A3)', '=SUBTOTAL(9,B1:B3)', '=SUBTOTAL(1,B1:B3)', '=SUBTOTAL(3,B1:B3)', '=SUBTOTAL(9.9,A1:A3)',
    '=SUBTOTAL(9,C1:C2)', '=SUBTOTAL(2,C1:C2)', '=SUBTOTAL(3,C1:C2)', '=SUBTOTAL(9,A4)', '=SUBTOTAL(9,A1:A7)+SUBTOTAL(9,A1:A3)', '=SUBTOTAL(4,A1:A7,C2)', '=SUBTOTAL(9,A1:C3)',
  ],
};

export const SHEETS = [logical, math, stats, conditions, lookup, texts, texts2, dates, subtotals];

/** The book for make-excel-book.ps1: each sheet's cells, its formulas in column H from row 1, and the placed formulas. */
export function bookSpec() {
  return {sheets: SHEETS.map(s => {
    const cells = {...s.cells};
    s.formulas.forEach((f, i) => { cells['H' + (i + 1)] = {f}; });
    for (const [a, f] of Object.entries(s.placed || {})) cells[a] = {f};
    return {name: s.name, cells};
  })};
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const here = dirname(fileURLToPath(import.meta.url)), dir = mkdtempSync(join(tmpdir(), 'excel-functions-')), spec = join(dir, 'spec.json');
  writeFileSync(spec, JSON.stringify(bookSpec()));
  const args = ['-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', join(here, 'make-excel-book.ps1'), '-Spec', spec, '-Out', join(here, 'excel-functions.xlsx')];
  if (process.argv[2]) args.push('-Results', process.argv[2]);
  const run = spawnSync('pwsh', args, {encoding: 'utf8'});
  rmSync(dir, {recursive: true, force: true});
  process.stdout.write(run.stdout); process.stderr.write(run.stderr);
  process.exitCode = run.status ?? 1;
}
