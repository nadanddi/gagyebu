const fs = require('fs');
const path = require('path');
const vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'Upload.html'), 'utf8');
const start = html.indexOf('<script>') + '<script>'.length;
const end = html.lastIndexOf('</script>');
if (start < '<script>'.length || end < start) throw new Error('Upload.html script block not found');
new vm.Script(html.slice(start, end));
console.log('upload syntax test passed');
