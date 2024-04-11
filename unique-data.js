const fs = require('fs');
const path = require('path');

// UNIQUE TITLES
let americas_titles = require('./lists/americas-titles.json');
let emeaa_titles = require('./lists/emeaa-titles.json');
let concatinated_titles = americas_titles.concat(emeaa_titles);
let unique_titles = removeDuplicates(concatinated_titles);
const uniqueTitlesPathJson = `./unique-data/titles.json`;
const uniqueTitlesPath = path.dirname(uniqueTitlesPathJson);
if (!fs.existsSync(uniqueTitlesPath)) {
    fs.mkdirSync(uniqueTitlesPath, { recursive: true });
}
fs.writeFileSync(uniqueTitlesPathJson, JSON.stringify(unique_titles.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })), null, 2));

// UNIQUE TIMEFRAMES
let americas_timeframe = require('./lists/americas-timeframe.json');
let emeaa_timeframe = require('./lists/emeaa-timeframe.json');
let concatinated_timeframe = americas_timeframe.concat(emeaa_timeframe);
let unique_timeframe = removeDuplicates(concatinated_timeframe);
const uniqueTimeframePathJson = `./unique-data/timeframe.json`;
const uniqueTimeframe = path.dirname(uniqueTimeframePathJson);
if (!fs.existsSync(uniqueTimeframe)) {
    fs.mkdirSync(uniqueTimeframe, { recursive: true });
}
fs.writeFileSync(uniqueTimeframePathJson, JSON.stringify(unique_timeframe.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })), null, 2));

// UNIQUE TIMEFRAMES
let americas_notes = require('./lists/americas-note.json');
let emeaa_notes = require('./lists/emeaa-note.json');
let concatinated_notes = americas_notes.concat(emeaa_notes);
let unique_notes = removeDuplicates(concatinated_notes);
const uniqueNotesPathJson = `./unique-data/notes.json`;
const uniqueNotes = path.dirname(uniqueNotesPathJson);
if (!fs.existsSync(uniqueNotes)) {
    fs.mkdirSync(uniqueNotes, { recursive: true });
}
fs.writeFileSync(uniqueNotesPathJson, JSON.stringify(unique_notes.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })), null, 2));

function removeDuplicates(arr) {
    let unique = [];
    arr.forEach(element => {
        const lowerCaseElement = element.toLowerCase(); // Convert to lowercase
        if (!unique.some(item => item.toLowerCase() === lowerCaseElement)) {
            unique.push(element);
        }
    });
    return unique;
}