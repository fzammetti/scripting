const fs = require("fs");


console.log("\nCompare Files\n");

// Grab command line arguments and log.  If incorrect number of arguments provided, show help.
const args = process.argv.slice(2);
if (args.length < 2 || args.length > 3) {
  console.log("Usage: node compareFiles.js filename1 filename2 [-v]\n");
  console.log("-v Verbose output");
  process.exit();
}
const [ filename1, filename2, optV ] = args;

// Process optional arguments.
let verbose = false;
if (optV !== undefined) {
  verbose = true;
}

// Show files being compared.
if (verbose) {
  console.log(`Files being compared: ${filename1}, ${filename2}\n`);
}

// Read in each file and create an object from it where each line is a key in the object.
const file1 = { };
const file2 = { };
const readFileIntoObject = (inFilename, inObject) => {
  const fileLines = fs.readFileSync(inFilename).toString().split("\r\n");
  for (let i in fileLines) {
    const nextLine = fileLines[i].trim();
    if (nextLine !== "") {
      inObject[fileLines[i]] = true;
    }
  }
}
readFileIntoObject (filename1, file1);
readFileIntoObject (filename2, file2);
if (verbose) {
  console.log("File contents read:");
  console.log(file1);
  console.log(file2);
  console.log(`\nItems in first file: ${Object.keys(file1).length}`);
  console.log(`Items in second file: ${Object.keys(file2).length}\n`);
}

// Now compare the files.  Note: need to do it in both directions to get all the mismatches.
let identical = true;
const compareFiles = (inFilename1, inFile1, inFilename2, inFile2) => {
  const file1Keys = Object.keys(inFile1) ;
  for (let i = 0; i < file1Keys.length; i++) {
    const nextItemToCheck = file1Keys[i];
    if (!inFile2[nextItemToCheck]) {
      console.log(`'${nextItemToCheck}' in ${inFilename1} not in ${inFilename2}`);
      identical = false;
    }
  }
}
compareFiles(filename1, file1, filename2, file2);
compareFiles(filename2, file2, filename1, file1);

if (identical) {
  console.log("Files are identical");
}
