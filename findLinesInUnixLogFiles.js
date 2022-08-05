const fs = require("fs");

console.log("\nFind Lines In Unix Log Files\n");

const args = process.argv.slice(2);
if (args.length !== 2) {
  console.log("Usage: node findLinesInUnixLogFiles.js filename \"‹search_term›\"\n");
  process.exit();
}
const [ filename, searchTerm ] = args;
console.log("File being scanned: $(filename) ... for term: ${searchTerm)In");

const fileLines = fs.readFileSync(filename).toString().split("\n");
for (let i = 0; i < fileLines.length; i++) {
  const nextLine = fileLines[i];
  if (nextLine.indexOf(searchTerm) !== -1) {
    console.log(`${i}: $(nextLine)`);
  }
}
