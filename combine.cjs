const fs = require('fs');
const files = ['chunk1.txt','chunk2.txt','chunk3.txt','chunk4.txt','chunk5.txt','chunk6.txt','chunk7.txt','chunk8.txt','chunk9.txt','chunk10.txt','chunk11.txt','chunk12.txt','chunk13.txt','chunk14.txt'];
let full = '';
for(const f of files) {
  full += fs.readFileSync('src/' + f, 'utf8') + '\n';
}
fs.writeFileSync('src/App.tsx', full);
