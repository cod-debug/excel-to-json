const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const he = require('he');

// Replace 'your-file.xlsx' with the name of your Excel file
const excelFilePath = 'OJT Template_EMEAA.xlsx';

// Load the Excel file
const workbook = XLSX.readFile(excelFilePath);

//  role_headers
const region = 'emeaa';
const brand_list = [
    {   
        brand_id: 'regent',
        brand_name: 'Regent',
        row_index: 1,
    },
    {   
        brand_id: 'intercontinental',
        brand_name: 'Intercontinental',
        row_index: 1,
    },
    {   
        brand_id: 'kimpton',
        brand_name: 'Kimpton',
        row_index: 1,
    },
    {   
        brand_id: 'vignette',
        brand_name: 'Vignette',
        row_index: 1,
    },
    {   
        brand_id: 'hotel-indigo',
        brand_name: 'Hotel Indigo',
        row_index: 1,
    },
    {   
        brand_id: 'voco-hotels',
        brand_name: 'Voco Hotels',
        row_index: 1,
    },
    {   
        brand_id: 'crowne-plaza',
        brand_name: 'Crowne Plaza',
        row_index: 1,
    },
    {   
        brand_id: 'holiday-inn',
        brand_name: 'Holiday Inn',
        row_index: 1,
    },
    {   
        brand_id: 'holiday-inn-express',
        brand_name: 'Holiday Inn Express',
        row_index: 1,
    },
    {   
        brand_id: 'staybridge-suites',
        brand_name: 'Staybridge Suites',
        row_index: 1,
    },
];

let emeaa = {
    emeaa: {},
}

brand_list.map((i) => {
    emeaa.emeaa[i.brand_id] = {
        name: i.brand_name,
        pictures: {
            logo: `./img/icons/${i.brand_id}.png`
        },
    };
});

// Alternatively, you can write the extracted data to a new JSON file
const jsonFilePath = `brands-${region}.json`;
// Ensure that the directories leading up to the file path exist
const directoryPath = path.dirname(jsonFilePath);

if (!fs.existsSync(directoryPath)) {
    // Create the directory structure recursively
    fs.mkdirSync(directoryPath, { recursive: true });
}

// Now you can write the data to the file
fs.writeFileSync(jsonFilePath, JSON.stringify(emeaa, null, 2));

console.log(`Data has been written to ${jsonFilePath}`);


var trainingTitleList = [];
var timeframeList = [];
var notesList = [];

const ojt_links_worksheet = workbook.Sheets['OJT Links'];
// Extract the lookup and return columns
let lookupColumn = [];
let returnColumn = [];
for (let rowIndex = 1; ; rowIndex++) {
    const cellAddressE = XLSX.utils.encode_cell({ r: rowIndex, c: 4 }); // Column E
    const cellAddressH = XLSX.utils.encode_cell({ r: rowIndex, c: 7 }); // Column H

    const cellE = ojt_links_worksheet[cellAddressE];
    const cellH = ojt_links_worksheet[cellAddressH];

    if (!cellE && !cellH) {
        break;
    }

    lookupColumn.push(cellE ? cellE.v : '');
    returnColumn.push(cellH ? cellH.v : '');
}

// Function to perform the lookup
function xlookup(lookupValue, lookupColumn, returnColumn, notFoundValue = "") {
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
        const sheetIndex = key+1;
        const sheetName = workbook.SheetNames[sheetIndex];
        const worksheet = workbook.Sheets[sheetName];
    
    
        let brand = {};
    
        brand[item.brand_id] = {
            'name': item.brand_name,
            'hero-image': './images/',
        }
    
        // Headers to identify the columns in the Excel sheet
        let headers = {
            'course_title': '',
            'instructions': '',
            'access_link': '',
            'access_text': ''
        };

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

            const hyperLinkURLAccess =  he.decode(worksheet[`C${rowIndex+1}`]?.l?.Target || '');; 
            // Perform the lookup for the value in column E and get the corresponding value from column H
            const lookupValue = worksheet[`E${rowIndex + 1}`]?.v; // Adjust this if the lookup value is from another column
            let hyperlinkURLCreateOjt = xlookup(lookupValue, lookupColumn, returnColumn);

            // restructure data to polish and trim values
            const restructuredData = {
                'course_title': rowValues['course_title'].replace(/\s+/g, ' ').trim(),
                'instructions': rowValues['instructions'].replace(/\s+/g, ' ').trim(),
                'access_text': rowValues['access_text'] ? 'Access ' + rowValues['access_text'].replace(/\s+/g, ' ').trim() : '',
                'access_link': hyperLinkURLAccess,
                'create_ojt_link': he.decode(hyperlinkURLCreateOjt),
            };

            // push to extracted data all trainings found
            if(restructuredData.course_title !== ""){
                extractedData.push(restructuredData);
            }
        }

        let brand_parsed = {
            ...brand,
            trainings: [
                ...extractedData
            ],
        }

        // Alternatively, you can write the extracted data to a new JSON file
        const jsonFilePath = `./${region}/${region}.${item.brand_id}.json`;
        
        // Ensure that the directories leading up to the file path exist
        const directoryPath = path.dirname(jsonFilePath);

        if (!fs.existsSync(directoryPath)) {
            // Create the directory structure recursively
            fs.mkdirSync(directoryPath, { recursive: true });
        }

        // Now you can write the data to the file
        fs.writeFileSync(jsonFilePath, JSON.stringify(brand_parsed, null, 2));

        console.log(`Data has been written to ${jsonFilePath}`);
        count++;
    });
    

    // Save unique titles on a json file
    const titlesFilePath=`./lists/emeaa-titles.json`;
    const titlesDirectoryPath = path.dirname(titlesFilePath);

    if (!fs.existsSync(titlesDirectoryPath)) {
        // Create the directory structure recursively
        fs.mkdirSync(titlesDirectoryPath, { recursive: true });
    }

    fs.writeFileSync(titlesFilePath, JSON.stringify(trainingTitleList, null, 2));

    // Save unique timeframes on a json file
    const timeframeFilePath=`./lists/emeaa-timeframe.json`;
    const timeframeDirectoryPath = path.dirname(timeframeFilePath);

    if (!fs.existsSync(timeframeDirectoryPath)) {
        // Create the directory structure recursively
        fs.mkdirSync(timeframeDirectoryPath, { recursive: true });
    }

    fs.writeFileSync(timeframeFilePath, JSON.stringify(timeframeList, null, 2));

    // Save unique notes on a json file
    const noteFilePath=`./lists/emeaa-note.json`;
    const noteDirectoryPath = path.dirname(noteFilePath);

    if (!fs.existsSync(noteDirectoryPath)) {
        // Create the directory structure recursively
        fs.mkdirSync(noteDirectoryPath, { recursive: true });
    }

    fs.writeFileSync(noteFilePath, JSON.stringify(notesList, null, 2));

    console.log(count, 'files');
}
