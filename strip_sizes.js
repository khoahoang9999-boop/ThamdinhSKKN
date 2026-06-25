const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

const parts = content.split('ref={printRef}');
if (parts.length === 2) {
  let printArea = parts[1];
  printArea = printArea.replace(/text-\[13px\] /g, '');
  printArea = printArea.replace(/text-\[14px\]/g, '');
  printArea = printArea.replace(/text-sm /g, '');
  printArea = printArea.replace(/ text-sm/g, '');
  printArea = printArea.replace(/text-xs /g, '');
  
  fs.writeFileSync('src/App.tsx', parts[0] + 'ref={printRef}' + printArea);
  console.log('done');
}
