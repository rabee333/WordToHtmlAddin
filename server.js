'use strict'

const express = require('express');
const fs = require('fs');
const http = require('http');
const bodyParser = require('body-parser');
const atob = require('atob');
const mammoth = require("mammoth");

const app = express();
const fileName = 'tranFile.docx'
const fileNameHtml = 'tranFile.html';
var bytesDataArr = [];

app.get('/GetDoc_App.html', async (req, res) => {
	try {
		await fs.readFile('static/GetDoc_App.html', 'utf-8', function(err, result){
            if(err){
                return res.status(500).json({Error: err});
            }
            return res.header('content-type', 'text/html').end(result);
        });
	}
	catch(err) {
		console.error(`error rendering GetDoc_App.xml file: ${err.message}`);
		return res.status(500).json({ error: err.message });
	}
});

app.get('/GetDoc_App.js', async (req, res) => {
	try {
		await fs.readFile('static/GetDoc_App.js', 'utf-8', function(err, result){
            if(err){
                return res.status(500).json({Error: err});
            }
            return res.header('content-type', 'text/html').end(result);
        });
	}
	catch(err) {
		console.error(`error rendering GetDoc_App.xml file: ${err.message}`);
		return res.status(500).json({ error: err.message });
	}
});

app.use(bodyParser.json());

function b64DecodeUnicode(str) {
    // Going backwards: from bytestream, to percent-encoding, to original string.
    return decodeURIComponent(atob(str).split('').map(function(c) {
        return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));
}


app.post('/submit', function (req,res) {
    if(!req.body){
        return res.end(`body is empty`);
    }

    var sliceNumber = req.headers["slice-number"];
    if(req.headers["is-first"]==="true"){
        bytesDataArr = [];
    }
    bytesDataArr[sliceNumber] = b64DecodeUnicode(req.body.data);

    if(req.headers["is-last"]==="true"){
        var allBytesArray = (bytesDataArr.join(",")).split(",");
        fs.writeFile(fileName, new Buffer(allBytesArray), (err) => {
            if(err) {
                console.error(err.message);
            }
            convertToHtml(fileName);
        });
        return res.end(`finish`);
    }

    return res.end(`ok`);
});

function convertToHtml(docxFile){
    mammoth.convertToHtml({path: fileName})
    .then(function(result){
        var html = result.value; // The generated HTML
        var messages = result.messages; // Any messages, such as warnings during conversion

        fs.writeFile(fileNameHtml, html, (err) => {
            if(err) console.error(err);
        })
    })
    .done();
}

app.get('/', async (req,res) => {
    return res.end(`server is on`);
});

http.createServer(app).listen(8080, err => {
    if (err) return console.error(err);
    console.info(`server is listening on port 8080`);
});

process.on('uncaughtException', err => {
	console.error(`uncaught exception: ${err.message}`);
	setTimeout(() =>  process.exit(1), 1000);
});
