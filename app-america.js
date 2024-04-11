const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const he = require('he');



// Replace 'your-file.xlsx' with the name of your Excel file
const excelFilePath = 'americas-v3.xlsx';

// Load the Excel file
const workbook = XLSX.readFile(excelFilePath);
//  role_headers

const brand_list = [
    {   
        region: 'americas',
        brand_id: 'regent',
        brand_name: 'Regent',
        sheet_index: 1,
        row_index: 2,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'intercontinental',
        brand_name: 'Intercontinental',
        sheet_index: 2,
        row_index: 1,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'vignette',
        brand_name: 'Vignette',
        sheet_index: 3,
        row_index: 1,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'hotel-indigo',
        brand_name: 'Hotel Indigo',
        sheet_index: 4,
        row_index: 1,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'voco-hotels',
        brand_name: 'Voco Hotels',
        sheet_index: 5,
        row_index: 2,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'crowne-plaza',
        brand_name: 'Crowne Plaza',
        sheet_index: 6,
        row_index: 1,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'even-hotels',
        brand_name: 'Even Hotels',
        sheet_index: 7,
        row_index: 2,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'revenue-manager': '',
            'food-and-beverage-manager-1': '',
            'hotel-experience-champion': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage-manager-2': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
    },
    {   
        region: 'americas',
        brand_id: 'holiday-inn',
        brand_name: 'Holiday Inn',
        sheet_index: 8,
        row_index: 1,
        role_headers: {
            'general-manager': null,
            'front-office-manager': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'food-and-beverage-director': null,
            'director-of-sales': null,
            'revenue-manager': null,
            'hotel-experience-champion': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'all-other-management-colleagues': null,
            'all-other-non-management-colleagues': null,
        },
    },
    {   
        region: 'americas',
        brand_id: 'holiday-inn-express',
        brand_name: 'Holiday Inn Express',
        sheet_index: 9,
        row_index: 1,
        role_headers: {
            'general-manager': null,
            'director-of-sales': null,
            'revenue-manager': null,
            'front-office-manager': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'hotel-experience-champion': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'all-other-management-colleagues': null,
            'breakfast-host': null,
        },
    },
    {   
        region: 'americas',
        brand_id: 'avid-hotels',
        brand_name: 'Avid Hotels',
        sheet_index: 10,
        row_index: 2,
        role_headers: {
            'general-manager': null,
            'front-office-manager': null,
            'director-of-sales': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'revenue-manager': null,
            'hotel-experience-champion': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'all-other-management-colleagues': null,
            'all-other-non-management-colleagues': null,
        },
    },
    {   
        region: 'americas',
        brand_id: 'garner',
        brand_name: 'Garner',
        sheet_index: 11,
        row_index: 2,
        role_headers: {
            'general-manager': null,
            'front-office-manager': null,
            'director-of-sales': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'revenue-manager': null,
            'hotel-experience-champion': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'all-other-management-colleagues': null,
            'all-other-non-management-colleagues': null,
        },
    },
    {   
        region: 'americas',
        brand_id: 'atwell-suites',
        brand_name: 'Atwell Suites',
        sheet_index: 12,
        row_index: 1,
        role_headers: {
            'general-manager': null,
            'front-office-manager': null,
            'revenue-manager': null,
            'director-of-sales': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'hotel-experience-champion': null,
            'fand-b-director': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'food-and-beverage': null,
            'all-other-management-colleagues': null,
            'all-other-non-management-colleagues': null,
        }
    },
    {   
        region: 'americas',
        brand_id: 'staybridge-suites',
        brand_name: 'Staybridge Suites',
        sheet_index: 13,
        row_index: 1,
        role_headers: {
            'general-manager': null,
            'front-office-manager': null,
            'director-of-sales': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'revenue-manager': null,
            'hotel-experience-champion': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'breakfast-host': null,
            'all-other-non-management-colleagues': null,
        },
    },
    {   
        region: 'americas',
        brand_id: 'candlewood-suites',
        brand_name: 'Candlewood Suites',
        sheet_index: 14,
        row_index: 1,
        role_headers: {
            'general-manager': null,
            'front-office-manager': null,
            'director-of-sales': null,
            'executive-housekeeper': null,
            'chief-engineer': null,
            'revenue-manager': null,
            'hotel-experience-champion': null,
            'front-desk': null,
            'housekeeping': null,
            'engineering': null,
            'all-other-management-colleagues': null,
            'all-other-non-management-colleagues': null,
        },
    },
    {   
        region: 'americas',
        brand_id: 'special-project',
        brand_name: 'IHG  Hotel',
        sheet_index: 15,
        row_index: 2,
        role_headers: {
            'general-manager': '',
            'front-office-manager': '',
            'revenue-manager': '',
            'director-of-sales': '',
            'executive-housekeeper': '',
            'chief-engineer': '',
            'hotel-experience-champion': '',
            'food-and-beverage-director': '',
            'front-desk': '',
            'housekeeping': '',
            'engineering': '',
            'food-and-beverage': '',
            'all-other-management-colleagues': '',
            'all-other-non-management-colleagues': '',
        },
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
console.log(americas);

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
    
        headers[item.brand_id] = "General Manager Operates";
    
        headers = {
            ...headers,
            'course-id': '',
            'timeframe': '',
            'notes': '',
            ...item.role_headers,
        }
        for (const role_id in item.role_headers) {
            // Initialize an array to store the extracted data
            const extractedData = [];
            if (item.role_headers.hasOwnProperty(role_id)) {
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
                    var hyperlinkURL = he.decode(worksheet[`B${rowIndex+1}`]?.l?.Target || '');
                    var courseID = rowValues['course-id'];

                    // ADDITIONAL CONDITIONS FOR CLIENT ADJUSTMENTS ON EXCEL FILE
                    /*  ---------------- ADJUSTMENT #1 ----------------
                        IHG Way of Clean
                        For all roles, brands, and regions:
                        • Replace link and course ID for IHG Way of Clean 5-S Cleaning Program to:
                        IHG7376524
                        https://mylearning.sumtotal.host/core/pillarRedirect?relyingParty=LM&amp;url=app%2Fmanagement%2FL
                        MS_ActDetails.aspx%3FActivityId%3D280596%26UserMode%3D0
                    */
                    if(rowValues[item.brand_id].toLowerCase() === "IHG Way of Clean 5-S Cleaning Program".toLowerCase()){
                        hyperlinkURL = `https://mylearning.sumtotal.host/core/pillarRedirect?relyingParty=LM&amp;url=app%2Fmanagement%2FLMS_ActDetails.aspx%3FActivityId%3D280596%26UserMode%3D0`;
                        courseID = `IHG7376524`;
                    }

                    /*  ---------------- ADJUSTMENT #2 ----------------
                        • Replace link and course ID for IHG Way of Clean for Non-Housekeeping Colleagues to:
                        IHG1227263
                        https://mylearning.sumtotal.host/core/pillarRedirect?relyingParty=LM&amp;url=app%2Fmanagement%2FL
                        MS_ActDetails.aspx%3FActivityId%3D310962%26UserMode%3D0
                    */
                    if(rowValues[item.brand_id].toLowerCase() === "IHG Way of Clean for Non-Housekeeping Colleagues".toLowerCase()){
                        hyperlinkURL = he.decode(`https://mylearning.sumtotal.host/core/pillarRedirect?relyingParty=LM&amp;url=app%2Fmanagement%2FLMS_ActDetails.aspx%3FActivityId%3D310962%26UserMode%3D0`);
                        courseID = `IHG1227263`;
                    }
                    // Construct the object with the extracted data
                    const rowData = {
                        ...rowValues,
                        'Course ID Link': hyperlinkURL,
                        'Course ID': courseID,
                    };

                    // Restructure the data with 'holiday-inn-express' as the key
                    const restructuredData = {
                        'title': rowData[item.brand_id].replace(/\s+/g, ' ').trim(),
                        'timeframe': rowData['timeframe'].replace(/\s+/g, ' ').trim(),
                        'notes': rowData['notes'].replace(/\s+/g, ' ').trim(),
                        'sorting': rowData[role_id],
                        'link': rowData['Course ID Link'],
                        'course-id': rowData['Course ID'],
                    };
                
                    let to_be_removed = [
                        'IHG Culture of Clean Implementation',
                        'IHG Way of Clean Bedding and Duvet',
                        'Duvet Training Materials and Care Guides',
                        'IHG Way of Clean Daily Room Refresh'
                    ];

                    if(!isNaN(rowData[role_id]) && rowData[role_id] != ''){
                        if(!to_be_removed.includes(restructuredData.title)){
                            extractedData.push(restructuredData);
                            
                            if(!['','remove', 'Remove', 'x', 'X'].includes(rowData[item.brand_id].replace(/\s+/g, ' ').trim())){
                                if(!trainingTitleList.includes(rowData[item.brand_id].replace(/\s+/g, ' ').trim())){
                                    trainingTitleList.push(rowData[item.brand_id].replace(/\s+/g, ' ').trim());
                                }
                            }

                            if(!['','remove', 'Remove', 'x', 'X'].includes(rowData['timeframe'].replace(/\s+/g, ' ').trim())){
                                if(!timeframeList.includes(rowData['timeframe'].replace(/\s+/g, ' ').trim())){
                                    timeframeList.push(rowData['timeframe'].replace(/\s+/g, ' ').trim());
                                }
                            }

                            if(!['','remove', 'Remove', 'x', 'X'].includes(rowData['notes'].replace(/\s+/g, ' ').trim())){
                                if(!notesList.includes(rowData['notes'].replace(/\s+/g, ' ').trim())){
                                    notesList.push(rowData['notes'].replace(/\s+/g, ' ').trim());
                                }
                            }
                        }
                    }
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
                const jsonFilePath = `./${item.region}/${item.region}.${item.brand_id}.${role_id}.json`;
                
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
            }
        }
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
