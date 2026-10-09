// In-memory worksheet model. All network operations belong to google.js.
(function(global){'use strict';
class Range{
 constructor(sheet,row,col,height=1,width=1){Object.assign(this,{sheet,row,col,height,width});}
 getRow(){return this.row;}
 getValues(){return Array.from({length:this.height},(_,r)=>Array.from({length:this.width},(_,c)=>this.sheet.values[this.row+r-1]?.[this.col+c-1]??''));}
 getValue(){return this.getValues()[0][0];}
 setValues(values){values.forEach((row,r)=>{const i=this.row+r-1;this.sheet.values[i]??=[];row.forEach((v,c)=>this.sheet.values[i][this.col+c-1]=v);});return this;}
 setValue(value){return this.setValues([[value]]);}
 clearContent(){return this.setValues(Array.from({length:this.height},()=>Array(this.width).fill('')));}
 createTextFinder(text){const range=this;let regex=false,entire=false;return {useRegularExpression(v){regex=v;return this;},matchEntireCell(v){entire=v;return this;},findAll(){const cells=[],pattern=regex?new RegExp(text):null;range.getValues().forEach((row,r)=>row.forEach((v,c)=>{const s=String(v);if(pattern?pattern.test(s):entire?s===text:s.includes(text))cells.push(new Range(range.sheet,range.row+r,range.col+c));}));return cells;}};}
}
class Sheet{
 constructor(name,values=[],sheetId=0,capacity=1000){Object.assign(this,{name,values,sheetId,capacity});}
 getLastRow(){let n=this.values.length;while(n&&!this.values[n-1].some(x=>x!==''&&x!=null))n--;return n;}
 getRange(...args){return new Range(this,...args);}
 getDataRange(){return this.getRange(1,1,Math.max(1,this.getLastRow()),Math.max(1,...this.values.slice(0,10).map(r=>r.length)));}
 appendRow(row){return this.getRange(this.getLastRow()+1,1,1,row.length).setValues([row]);}
}
class Book{
 constructor(tables=[]){this.sheets={};tables.forEach(t=>this.sheets[t.name]=new Sheet(t.name,structuredClone(t.values||[]),t.sheetId,t.capacity));}
 getSheetByName(name){return this.sheets[name]||null;}
 toJSON(){return Object.values(this.sheets).map(s=>({name:s.name,values:s.values,sheetId:s.sheetId,capacity:s.capacity}));}
 clone(){return new Book(this.toJSON());}
}
const cell=v=>({userEnteredValue:typeof v==='number'?{numberValue:v}:typeof v==='boolean'?{boolValue:v}:{stringValue:String(v??'')}});
// Preserve stable rows on deletion, update changed cells, append new rows server-side.
function changes(before,after,widths){const requests=[];Object.values(after.sheets).forEach(s=>{const old=before.sheets[s.name],end=old.getLastRow(),width=widths[s.name];if(!width)return;
 for(let r=0;r<end;r++){const a=s.values[r]||[],b=old.values[r]||[];for(let c=0;c<width;c++){if((a[c]??'')!==(b[c]??''))requests.push({updateCells:{range:{sheetId:s.sheetId,startRowIndex:r,endRowIndex:r+1,startColumnIndex:c,endColumnIndex:c+1},rows:[{values:[cell(a[c])]}],fields:'userEnteredValue'}});}}
 const rows=s.values.slice(end,s.getLastRow());if(rows.length)requests.push({appendCells:{sheetId:s.sheetId,rows:rows.map(r=>({values:Array.from({length:width},(_,c)=>cell(r[c]))})),fields:'userEnteredValue'}});
 });return requests;}
global.ChoirModel={Book,changes,cell};
})(globalThis);
