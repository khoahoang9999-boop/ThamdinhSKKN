const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

// replace className="text-justify mb-4..." with className="mb-4..." style={{ textAlign: 'justify' }}
content = content.replace(/className="text-justify mb-4/g, 'style={{ textAlign: \'justify\' }} className="mb-4');

// In renderList:
// return <p key={idx} className="mb-1 text-justify">- {text}</p>;
content = content.replace(/className="mb-1 text-justify"/g, 'className="mb-1" style={{ textAlign: \'justify\' }}');

// In uuDiem and hanChe map:
// <div className="space-y-1 text-justify">
content = content.replace(/className="space-y-1 text-justify"/g, 'className="space-y-1" style={{ textAlign: \'justify\' }}');

// <div className="text-justify italic text-gray-700">
content = content.replace(/className="text-justify italic text-gray-700"/g, 'className="italic text-gray-700" style={{ textAlign: \'justify\' }}');


fs.writeFileSync('src/App.tsx', content);
