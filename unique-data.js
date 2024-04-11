const fs = require('fs');
const path = require('path');

// UNIQUE TITLES
let americas_titles = require('./lists/americas-titles.json');
let emeaa_titles = require('./lists/emeaa-titles.json');
let concatinated_titles = americas_titles.concat(emeaa_titles);
let unique_titles = concatinated_titles.filter((item,index) => concatinated_titles.indexOf(item) === index);

// Alternatively, you can write the extracted data to a new JSON file
const uniqueTitlesPathJson = `./unique-data/titles.json`;

// Ensure that the directories leading up to the file path exist
const uniqueTitlesPath = path.dirname(uniqueTitlesPathJson);

if (!fs.existsSync(uniqueTitlesPath)) {
    // Create the directory structure recursively
    fs.mkdirSync(uniqueTitlesPath, { recursive: true });
}

// Now you can write the data to the file
fs.writeFileSync(uniqueTitlesPathJson, JSON.stringify(unique_titles.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })), null, 2));

// UNIQUE TIMEFRAMES
let americas_timeframe = require('./lists/americas-timeframe.json');
let emeaa_timeframe = require('./lists/emeaa-timeframe.json');
let concatinated_timeframe = americas_timeframe.concat(emeaa_timeframe);
let unique_timeframe = concatinated_timeframe.filter((item,index) => concatinated_timeframe.indexOf(item) === index);

// Alternatively, you can write the extracted data to a new JSON file
const uniqueTimeframePathJson = `./unique-data/timeframe.json`;

// Ensure that the directories leading up to the file path exist
const uniqueTimeframe = path.dirname(uniqueTimeframePathJson);

if (!fs.existsSync(uniqueTimeframe)) {
    // Create the directory structure recursively
    fs.mkdirSync(uniqueTimeframe, { recursive: true });
}

// Now you can write the data to the file
fs.writeFileSync(uniqueTimeframePathJson, JSON.stringify(unique_timeframe.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })), null, 2));

// UNIQUE TIMEFRAMES
let americas_notes = require('./lists/americas-note.json');
let emeaa_notes = require('./lists/emeaa-note.json');
let concatinated_notes = americas_notes.concat(emeaa_notes);
let unique_notes = concatinated_notes.filter((item,index) => concatinated_notes.indexOf(item) === index);

// Alternatively, you can write the extracted data to a new JSON file
const uniqueNotesPathJson = `./unique-data/notes.json`;

// Ensure that the directories leading up to the file path exist
const uniqueNotes = path.dirname(uniqueNotesPathJson);

if (!fs.existsSync(uniqueNotes)) {
    // Create the directory structure recursively
    fs.mkdirSync(uniqueNotes, { recursive: true });
}

// Now you can write the data to the file
fs.writeFileSync(uniqueNotesPathJson, JSON.stringify(unique_notes.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })), null, 2));