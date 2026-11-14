const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const assert = require('node:assert/strict');
const dir = __dirname;
const html = fs.readFileSync(path.join(dir, 'Index.html'), 'utf8');
const fixture = [
 ['Criteria', 'คำนำหน้า', 'ชื่อ - สกุล', 'ชื่อเล่น', 'รหัสนิสิต', 'เบอร์โทร', 'Dancer', 'Staff', 'First Date'],
 ['Qualified', '', 'Demo One', 'Min', '6800000001', '081-000-0001', '2', '1', 'Staff'],
 ['Unqualified', '', 'Demo Two', 'Ploy', '6800000002', '0810000002', '0', '0', ''],
 ['Choreo', '', 'Demo Three', 'Jay', '6800000003', '+66 81 000 0002', '', '1', 'Staff'],
 ['หมายเหตุ', '', '', '', '', '', '', '', '']
];
let rows = structuredClone(fixture), reads = 0, fail = false;
const sheet = {getSheetId: () => 0, getName: () => 'Attendance', getLastColumn: () => rows[0].length,
 getLastRow: () => rows.length, getRange: () => ({getDisplayValues: () => { reads++; return structuredClone(rows); }})};
let choreoRows = [[...fixture[0].slice(0,6), 'จำนวนโชว์', 'First Date', 'Show'], ['Unqualified', '', 'Demo Three', 'Jay', '6800000003', '', '1', 'TRUE', 'FALSE']];
const choreoSheet = {getSheetId: () => 1, getName: () => 'Choreo', getLastColumn: () => choreoRows[0].length, getLastRow: () => choreoRows.length, getRange: () => ({getDisplayValues: () => structuredClone(choreoRows)})};
const context = vm.createContext({SpreadsheetApp: {openById: () => { if(fail) throw Error('private details'); return {getSheets: () => [sheet, choreoSheet]}; }}});
vm.runInContext(fs.readFileSync(path.join(dir,'Code.gs'),'utf8'),context);
const call = (method,...args) => JSON.parse(JSON.stringify(context[method](...args)));
let checks = 0;
function check(label, action) { action(); checks++; console.log('PASS',label); }
check('Client JS parses; production HTML has no demo identifiers', () => {
 new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1]);
 assert(!html.includes('6800000001')); assert(!html.includes('0810000001'));
});
check('Exact student ID lookup', () => assert.equal(call('searchMembers','6800000001').result.member.name,'Demo One'));
check('Phone separators, country code, Thai numerals', () => {
 for (const query of ['0810000001','081-000-0001','+66 81 000 0001','๐๘๑๐๐๐๐๐๐๑'])
  assert.equal(call('searchMembers',query).result.member.name,'Demo One');
});
check('Formatted and unformatted phone numbers match in both directions', () => {
 rows[1][5] = '0123456789';
 assert.equal(call('searchMembers','012-345-6789').result.member.name,'Demo One');
 rows[1][5] = '012-345-6789';
 assert.equal(call('searchMembers','0123456789').result.member.name,'Demo One');
 rows = structuredClone(fixture);
});
check('Names, partial numbers, malformed inputs rejected before reading', () => {
 const before = reads;
 for(const query of ['Min','68000','','*','letters0810000001','0810000001/','++66810000001',null,{},123,'1'.repeat(31)])
  assert.equal(call('searchMembers',query).code,'INVALID_QUERY');
 assert.equal(reads,before);
});
check('Duplicate phone sends only member choices', () => {
 const result = call('searchMembers','0810000002').result;
 assert.equal(result.kind,'matches'); assert.equal(result.matches.length,2);
 assert.deepEqual(Object.keys(result.matches[0]).sort(),['name','nickname','row']);
});
check('Selected member returns Choreo without exposing identifiers', () => {
 const choice = call('searchMembers','0810000002').result.matches[1];
 const response = call('getMember','+66810000002',choice);
 assert.equal(response.result.member.status,'Unqualified'); assert.equal(response.result.member.isChoreo,true); assert.equal(response.result.member.choreo,1); assert.deepEqual(response.result.member.activities,[{name:'First Date',role:'Choreo'}]);
 assert(!JSON.stringify(response).includes('6800000003')); assert(!JSON.stringify(response).includes('66810000002'));
 assert.deepEqual(Object.keys(response.result.member).sort(),['activities','choreo','hasUnrecognizedRole','isChoreo','name','nickname','status']);
});
check('Every search and selection reads again', () => {
 const before = reads; call('searchMembers','6800000001'); call('searchMembers','6800000001'); assert.equal(reads,before+2);
 rows[1][0]='Unqualified'; assert.equal(call('searchMembers','6800000001').result.member.status,'Unqualified'); rows=structuredClone(fixture);
});
check('Moved row, edited ID and mismatched query cannot select stale record', () => {
 const choice = call('searchMembers','0810000002').result.matches[0];
 [rows[2],rows[3]]=[rows[3],rows[2]]; assert.equal(call('getMember','0810000002',choice).code,'STALE_SELECTION'); rows=structuredClone(fixture);
 assert.equal(call('getMember','6800000001',choice).code,'STALE_SELECTION');
});
check('Valid unknown number gives empty; empty source identifiers never match', () => {
 assert.equal(call('searchMembers','6899999999').result.kind,'empty');
 rows[1][4]=''; rows[1][5]=''; assert.equal(call('searchMembers','6800000001').result.kind,'empty'); rows=structuredClone(fixture);
});
check('Student ID and phone searches use separate columns', () => {
 rows[1][4] = '0810000009'; rows[1][5] = '6800000009';
 assert.equal(call('searchMembers','0810000009').result.kind,'empty');
 assert.equal(call('searchMembers','6800000009').result.kind,'empty');
 rows=structuredClone(fixture);
});
check('E/F header drift fails closed', () => {
 for(const column of [4,5]) {rows[0][column]='Changed'; assert.equal(call('searchMembers','6800000001').code,'SHEET_STRUCTURE'); rows=structuredClone(fixture);}
});
check('Inserted Workshop column does not shift counts or activities', () => {
 rows.forEach((row, index) => row.splice(6, 0, index === 0 ? 'Workshop' : '7'));
 const member = call('searchMembers','6800000001').result.member;
 assert.equal(member.dancer,2); assert.equal(member.staff,1);
 assert.deepEqual(member.activities.map(a => a.name), ['First Date']);
 rows = structuredClone(fixture);
});
check('Ambiguous required headers fail closed', () => {
 rows[0][1]='Dancer'; assert.equal(call('searchMembers','6800000001').code,'SHEET_STRUCTURE');
 rows=structuredClone(fixture);
});
check('Service error hides raw exception', () => { fail=true; assert.deepEqual(call('searchMembers','6800000001'),{ok:false,code:'SERVICE_UNAVAILABLE'}); fail=false; });
check('Unknown criteria remains Unknown; counts are not inferred', () => {
 rows[1][0]='constructor'; const member=call('searchMembers','6800000001').result.member;
 assert.equal(member.status,'constructor'); assert.equal(member.dancer,2); assert.equal(member.activities.length,1);
});
check('Workshop activity and count are returned', () => { rows.forEach((r,i)=>r.splice(6,0,i===0?'Workshop':'3')); rows[1][9]='Workshop'; const m=call('searchMembers','6800000001').result.member; assert.equal(m.workshop,3); assert.equal(m.activities[0].role,'Workshop'); rows=structuredClone(fixture); });
check('Choreo missing and duplicate matches fail safely', () => { const saved=structuredClone(choreoRows); choreoRows.pop(); assert.equal(call('searchMembers','6800000003').code,'CHOREO_LOOKUP'); choreoRows=structuredClone(saved); choreoRows.push([...saved[1]]); assert.equal(call('searchMembers','6800000003').code,'CHOREO_LOOKUP'); choreoRows=saved; });
check('Board and Core Team use sheet values', () => { for (const status of ['Board','Core Team']) { rows[1][0]=status; assert.equal(call('searchMembers','6800000001').result.member.status,status); } rows=structuredClone(fixture); });
console.log(`${checks} identifier-search checks passed. SpreadsheetApp is mocked; no live deployment test.`);
