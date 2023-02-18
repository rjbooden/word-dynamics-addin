/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
// copy script
const { execSync } = require("child_process");
let fs =  require('fs');
let path = require('path');

const args = process.argv;
const buildCmd = (args.length && args[args.length - 1] == 'dev') ? "npm run build:dev" : "npm run build";

console.log('starting build');

execSync(buildCmd, {stdio: 'inherit'});

console.log('updating the addin in the api');

const source = "./dist";
const destination = "../word-dynamics-api/wwwroot";

const files = fs.readdirSync(destination);
for (const file of files) {
    let removePath = path.join(destination, file);
    if (fs.lstatSync(removePath).isDirectory()) {
        console.log("Removing directory: " + removePath);
        fs.rmSync(removePath, { recursive: true, force: true });
    }
    else {
        console.log("Removing file: " + removePath);
        fs.unlinkSync(removePath);
    }
}

console.log("Copy contents from '" + source + "' to '" + destination + "'");
fs.cpSync(source, destination, {recursive: true});