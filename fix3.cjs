const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

content = content.replace(/<h4 className="font-bold mb-2 mt-4 text-\[14px\]">III\. TỔNG HỢP ĐIỂM<\/h4>/g, '<h4 className="font-bold mb-2 mt-4">III. TỔNG HỢP ĐIỂM</h4>');

content = content.replace(/<h4 className="font-bold mb-2 text-\[14px\]">IV\. XẾP LOẠI ĐỀ NGHỊ<\/h4>/g, '<h4 className="font-bold mb-2">IV. XẾP LOẠI ĐỀ NGHỊ</h4>');

content = content.replace(/<div className="mb-8 text-\[14px\]">/g, '<div className="mb-8">');

fs.writeFileSync('src/App.tsx', content);
