const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const he = require('he');



// Replace 'your-file.xlsx' with the name of your Excel file
const excelFilePath = 'OJT Template_AMER.xlsx';

// Load the Excel file
const workbook = XLSX.readFile(excelFilePath);
//  role_headers

const brand_list = [
    {   
        region: 'americas',
        brand_id: 'regent',
        brand_name: 'Regent',
        sheet_index: 1,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'intercontinental',
        brand_name: 'Intercontinental',
        sheet_index: 2,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'vignette',
        brand_name: 'Vignette',
        sheet_index: 3,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'hotel-indigo',
        brand_name: 'Hotel Indigo',
        sheet_index: 4,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'voco-hotels',
        brand_name: 'Voco Hotels',
        sheet_index: 5,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'crowne-plaza',
        brand_name: 'Crowne Plaza',
        sheet_index: 6,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'even-hotels',
        brand_name: 'Even Hotels',
        sheet_index: 7,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'holiday-inn',
        brand_name: 'Holiday Inn',
        sheet_index: 8,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'holiday-inn-express',
        brand_name: 'Holiday Inn Express',
        sheet_index: 9,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'avid-hotels',
        brand_name: 'Avid Hotels',
        sheet_index: 10,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'garner',
        brand_name: 'Garner',
        sheet_index: 11,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'atwell-suites',
        brand_name: 'Atwell Suites',
        sheet_index: 12,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'staybridge-suites',
        brand_name: 'Staybridge Suites',
        sheet_index: 13,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'candlewood-suites',
        brand_name: 'Candlewood Suites',
        sheet_index: 14,
        row_index: 1,
    },
    {   
        region: 'americas',
        brand_id: 'special-project',
        brand_name: 'IHG  Hotel',
        sheet_index: 15,
        row_index: 1,
    }
];

let americas = {
    americas: {},
}
brand_list.map((i) => {
    americas.americas[i.brand_id] = {
        name: i.brand_name,
        pictures: {
            logo: `./img/icons/${i.brand_id}.png`
        },
        departments: {

        }
    };
    for (const role in i.role_headers){
        americas.americas[i.brand_id]['departments'][role] = { name: capitalizeEachWord(role.replaceAll("-", " "))}
    }
});

// Alternatively, you can write the extracted data to a new JSON file
const jsonFilePath = `brand-departments-america.json`;
// Ensure that the directories leading up to the file path exist
const directoryPath = path.dirname(jsonFilePath);

if (!fs.existsSync(directoryPath)) {
    // Create the directory structure recursively
    fs.mkdirSync(directoryPath, { recursive: true });
}

// Now you can write the data to the file
fs.writeFileSync(jsonFilePath, JSON.stringify(americas, null, 2));

console.log(`Data has been written to ${jsonFilePath}`);


var trainingTitleList = [];
var timeframeList = [];
var notesList = [];

function capitalizeEachWord(sentence) {
    // Split the sentence into an array of words
    var words = sentence.split(' ');

    // Capitalize the first letter of each word
    var capitalizedWords = words.map(function(word) {
        return word.charAt(0).toUpperCase() + word.slice(1);
    });

    // Join the words back into a sentence
    var capitalizedSentence = capitalizedWords.join(' ');

    return capitalizedSentence;
}


const ojt_worksheet = workbook.Sheets['OJT Links'];
// Extract the lookup and return columns
let lookupColumn = [];
let returnColumn = [];
for (let rowIndex = 1; ; rowIndex++) {
    const cellAddressE = XLSX.utils.encode_cell({ r: rowIndex, c: 4 }); // Column E
    const cellAddressH = XLSX.utils.encode_cell({ r: rowIndex, c: 7 }); // Column H

    const cellE = ojt_worksheet[cellAddressE];
    const cellH = ojt_worksheet[cellAddressH];

    if (!cellE && !cellH) {
        break;
    }

    lookupColumn.push(cellE ? cellE.v : '');
    returnColumn.push(cellH ? cellH.v : '');
}

// Function to perform the lookup
function xlookup(lookupValue, lookupColumn, returnColumn, notFoundValue = "Not found") {
    for (let rowIndex = 1; rowIndex <= lookupColumn.length; rowIndex++) {
        if (lookupColumn[rowIndex - 1] === lookupValue) {
            return returnColumn[rowIndex - 1];
        }
    }
    return notFoundValue;
}

generateJson(brand_list);

function generateJson(data){
    
    var count = 0;
    data.map((item, key) => {
        // Assuming you want to read data from the second sheet (index 1)
        const sheetIndex = item.sheet_index;
        const sheetName = workbook.SheetNames[sheetIndex];
        const worksheet = workbook.Sheets[sheetName];
    
    
        let brand = {};
    
        brand[item.brand_id] = {
            'name': item.brand_name,
            'hero-image': './images/',
        }
    
        // Headers to identify the columns in the Excel sheet
        let headers = {
        };
    
        headers = {
            ...headers,
            'course-id': '',
            'timeframe': '',
            'notes': '',
        }
        // Initialize an array to store the extracted data
        const extractedData = [];
        // Iterate through each row in the worksheet
        for (let rowIndex = item.row_index; ; rowIndex++) {
            // Construct the cell address for each column in the current row
            const rowValues = Object.keys(headers).reduce((acc, key, colIndex) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                const cell = worksheet[cellAddress];
                acc[key] = cell?.v || '';
                return acc;
            }, {});
        
            // Check if all cells in the row are empty
            const isRowEmpty = Object.values(rowValues).every(value => value === '');
        
            // If the row is empty, stop the iteration
            if (isRowEmpty) {
                break;
            }

            // Extract the hyperlink URL from the 'Course ID' column if it exists
            let hyperlinkURLAccess = he.decode(worksheet[`C${rowIndex+1}`]?.l?.Target || '');

            // Perform the lookup for the value in column E and get the corresponding value from column H
            const lookupValue = worksheet[`E${rowIndex + 1}`]?.v; // Adjust this if the lookup value is from another column
            let hyperlinkURLCreateOjt = xlookup(lookupValue, lookupColumn, returnColumn);

            console.log(hyperlinkURLAccess, 'ACCESS');
            console.log(hyperlinkURLCreateOjt, 'CREATE OJT');
            let courseID = rowValues['course-id'];

            
        }

        /*  ---------------- ADJUSTMENT #3 ----------------
            • Add the following:
            IHG Way of Clean Resource Library for ALL roles:
            https://ihg.bravais.com/s/hwo8uaeFvG8WYA8aQ0pu
            button should say “Access”
            Notes should say:
            Explore additional resources to integrate IHG Way of Clean into your daily routine.
        */

        extractedData.push({
            'title': 'IHG Way of Clean Resource Library',
            'timeframe': '',
            'notes': 'Explore additional resources to integrate IHG Way of Clean into your daily routine.',
            'sorting': 100,
            'link': 'https://ihg.bravais.com/s/hwo8uaeFvG8WYA8aQ0pu',
            'course-id': 'Access',
        })

        let brand_parsed = {
            ...brand,
            trainings: [
                ...extractedData
            ],
        }

        // Alternatively, you can write the extracted data to a new JSON file
        const jsonFilePath = `./${item.region}/${item.region}.${item.brand_id}.json`;
        
        // Ensure that the directories leading up to the file path exist
        const directoryPath = path.dirname(jsonFilePath);

        // if (!fs.existsSync(directoryPath)) {
        //     // Create the directory structure recursively
        //     fs.mkdirSync(directoryPath, { recursive: true });
        // }

        // // Now you can write the data to the file
        // fs.writeFileSync(jsonFilePath, JSON.stringify(brand_parsed, null, 2));

        console.log(`Data has been written to ${jsonFilePath}`);
        count++;
    });
    

    // Save unique titles on a json file
    const titlesFilePath=`./lists/americas-titles.json`;
    const titlesDirectoryPath = path.dirname(titlesFilePath);

    if (!fs.existsSync(titlesDirectoryPath)) {
        // Create the directory structure recursively
        fs.mkdirSync(titlesDirectoryPath, { recursive: true });
    }

    fs.writeFileSync(titlesFilePath, JSON.stringify(trainingTitleList, null, 2));

    // Save unique timeframes on a json file
    const timeframeFilePath=`./lists/americas-timeframe.json`;
    const timeframeDirectoryPath = path.dirname(timeframeFilePath);

    if (!fs.existsSync(timeframeDirectoryPath)) {
        // Create the directory structure recursively
        fs.mkdirSync(timeframeDirectoryPath, { recursive: true });
    }

    fs.writeFileSync(timeframeFilePath, JSON.stringify(timeframeList, null, 2));

    // Save unique notes on a json file
    const noteFilePath=`./lists/americas-note.json`;
    const noteDirectoryPath = path.dirname(noteFilePath);

    if (!fs.existsSync(noteDirectoryPath)) {
        // Create the directory structure recursively
        fs.mkdirSync(noteDirectoryPath, { recursive: true });
    }

    fs.writeFileSync(noteFilePath, JSON.stringify(notesList, null, 2));

    console.log(count, 'files');
}
