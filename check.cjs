const fs = require('fs');
let content = fs.readFileSync('src/App.tsx', 'utf8');

const hasEscapedBackticks = content.includes('\\`');
const hasEscapedDollars = content.includes('\\$');
console.log('Escaped backticks:', hasEscapedBackticks);
console.log('Escaped dollars:', hasEscapedDollars);
