const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

// replace text-[13px] starting from line 2500 onwards
const lines = content.split('\n');
for (let i = 2500; i < lines.length; i++) {
  lines[i] = lines[i].replace(/text-\[13px\] /g, 'text-justify ');
}

fs.writeFileSync('src/App.tsx', lines.join('\n'));
