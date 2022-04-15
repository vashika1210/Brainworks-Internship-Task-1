
// Client ID and API key from the Developer Console
var CLIENT_ID = "855455246635-ri45s8cmq0lj8310ve9g622sl0dh2oku.apps.googleusercontent.com";
var API_KEY = "AIzaSyDYCqtvJVndtTGQFgrsm5gYEYMPvtqTaIk";

// Array of API discovery doc URLs for APIs used by the app
var DISCOVERY_DOCS = ["https://sheets.googleapis.com/$discovery/rest?version=v4"];

// Authorization scopes required by the API
var SCOPES = "https://www.googleapis.com/auth/spreadsheets";

// Target spreadsheet ID
var SHEET_ID = "1vReWYcdne_unhJyp44CKSBRp7dOe0j9aoOLsQpScknU";

// Form element controls
var fName = document.getElementById('fName');
var lName = document.getElementById('lName');
var fetchName = document.getElementById('fetchName');
var saveForm = document.getElementById("saveForm");
var fetchForm = document.getElementById("fetchForm");


// Add event listeners to prevent page reload
function handleForm(event) { event.preventDefault();}  
saveForm.addEventListener('submit', handleForm);
fetchForm.addEventListener('submit', handleForm);

/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {
    gapi.load('client:auth2', initClient);
}

/**
 *  Initializes the API client library
 */
function initClient() {
    gapi.client.init({
        apiKey: API_KEY,
        clientId: CLIENT_ID,
        discoveryDocs: DISCOVERY_DOCS,
        scope: SCOPES
    }).catch((err) => {
        handleError({phase: "initialization", type: err.error, details: err.details});
    });
}

/**
 *  Handle form save requests
 */
function handleSave(event) {
    // Display loading icon while processing
    toggleLoadingIcon(true, "save");

    // If not already signed in, call OAuth Sign-in module
    // Sign-in only required to save data  
    if (!gapi.auth2.getAuthInstance().isSignedIn.get()) {
        gapi.auth2.getAuthInstance().signIn().catch(err => {
            handleError({phase: "authentication", type: err.error, details: err.details})
        })
    }

    var values = [
        [
            fName.value, lName.value
        ]
    ];
    var body = {
        values: values
    };

    // Get existing names to check for duplicates            
    gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: 'Sheet1!A2:B',
    }).then(response => {
        if (findName(response.result.values, fName.value, lName.value)) {
            return handleError({phase: "name save", type:"Duplicate name", details: "Given name already exists! Please provide a unique value."})
        }
        // Save valid name to the spreadsheet
        gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SHEET_ID,
            range: 'Sheet1!A1:B1',
            valueInputOption: "USER_ENTERED",
            resource: body
        }).then((response) => {
            var result = response.result;
            // Hide loading icon
            toggleLoadingIcon(false, "save");
            // Display success message
            document.getElementById("error").innerText = `"${fName.value} ${lName.value}" saved successfully!`;
            document.getElementById("error").style.color = "#1BD073";
        });

    }).catch(err => {
        handleError({phase: "data load", type: err.error, details: err.details});       
    })
}

/**
 * Handle form fetch requests
 */
function handleFetch(event) {
    // Show loading icon
    toggleLoadingIcon(true, "fetch");

    gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: 'Sheet1!A2:B',
    }).then(response => {
        // If requested name does not exist
        var name = findName(response.result.values, fetchName.value)
        if (!name) {
            toggleLoadingIcon(false, "fetch");
            return alert("Given name does not exist! Please provide a valid name.")
        }
        // If requested name is found
        document.getElementById("result").innerText = "Last name: " + name[1]
        // Hide loading icon
        toggleLoadingIcon(false, "fetch");
    }).catch(err => {
        handleError({phase: "data load", type: err.error, details: err.details});  
    })
}

/* Helper Functions */

/**
 * Find a matching name in the given nested array
 * @param {Array} arr Array to be searched
 * @param {String} fName Target first name 
 * @param {String} lName Target last name [optional]
 * @returns {[fName, lName]} if found
 * or:
 * @returns {undefined} if not found 
 */

function findName(arr, fName, lName) {
    return arr.find(val => {
        if (lName) {            
            return (val[0] === fName && val[1] === lName)
        } else return (val[0] === fName)
    })
}

/**
 * Handle loading icon toggle
 *
 * @param {Boolean} shouldShow Whether to show or hide the loading icon
 * @param {String} target "save" or "fetch" to toggle respective loading icons
 */
function toggleLoadingIcon(shouldShow, target) {
    var loadingIcon = target === "save" ? document.getElementById("saveLoadingIcon") : document.getElementById("fetchLoadingIcon");
    var btnLabel = target === "save" ? document.getElementById("btnSaveLabel") : document.getElementById("btnFetchLabel");
    if (shouldShow) {        
        loadingIcon.style.display = "inline-block";
        btnLabel.innerText = "";
    } else {
        loadingIcon.style.display = "none";
        btnLabel.innerText = target;
    }
    
}

/**
 * Handle errors or display success status
 *
 * @param {Object} error Error object 
 * @param {String} error.phase Phase during which the error occured
 * @param {String} error.type Type of given error 
 * @param {String} error.details Details of given error
 */
function handleError(error = {phase, type, details}) {
    toggleLoadingIcon(false, "save");
    toggleLoadingIcon(false, "fetch");
    document.getElementById("error").style.color = "#FB6080";
    document.getElementById("error").innerHTML = `Error occured during ${error.phase}!<br>Type: ${error.type}<br>Details: ${error.details}`
}