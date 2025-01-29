// Initialize the keyword list
let keywordList = {};

// Fetch and parse the XML data from the hosted URL
function fetchKeywordData() {
    fetch('https://sziszi38.github.io/Keyword-information/')
        .then(response => response.text()) // Fetch the XML as text
        .then(data => parseXML(data)) // Parse the XML when it's fetched
        .catch(error => console.error('Error fetching XML:', error));
}

// Parse the XML string to extract keywords and associated pop-up words
function parseXML(xmlString) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "application/xml");

    // Clear the existing keyword list
    keywordList = {};

    // Get all <keyword> elements from the XML
    const keywords = xmlDoc.getElementsByTagName("keyword");

    // Iterate over the keywords and store them in the keywordList object
    for (let i = 0; i < keywords.length; i++) {
        const keyword = keywords[i].getAttribute('name'); // Get the keyword name
        const info = keywords[i].getElementsByTagName('info')[0].textContent; // Get the info (pop-up word)
        keywordList[keyword] = info; // Store in the keywordList object
    }

    // Optionally: display the keyword list in the UI
    displayKeywordList();
}

// Display the current list of keywords in the UI (side panel)
function displayKeywordList() {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = ''; // Clear current content

    Object.keys(keywordList).forEach(keyword => {
        const info = keywordList[keyword];
        const keywordItem = document.createElement('div');
        keywordItem.textContent = `${keyword}: ${info}`;
        resultDiv.appendChild(keywordItem);
    });
}

// Handle the form submission to add a new keyword
document.getElementById('keywordForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const keyword = document.getElementById('keyword').value;
    const popupWord = document.getElementById('popupWord').value;

    // Add the new keyword and its pop-up word to the list
    keywordList[keyword] = popupWord;

    // Optionally: Update the displayed keyword list
    displayKeywordList();

    // Clear the input fields
    document.getElementById('keyword').value = '';
    document.getElementById('popupWord').value = '';
});

// Function to check for a keyword in the email content
function checkForKeyword(emailText) {
    // Loop through the stored keyword list and check for matches
    for (const keyword in keywordList) {
        // Create a regex pattern to match the keyword with any additional characters (e.g., B3232 and B32324)
        const regex = new RegExp(`\\b${keyword}\\w*\\b`, 'g'); 

        if (regex.test(emailText)) {
            // Display the popup with the associated info
            displayPopup(keyword, keywordList[keyword]);
        }
    }
}

// Function to display a pop-up with the keyword info
function displayPopup(keyword, info) {
    const popup = document.createElement('div');
    popup.className = 'popup';
    popup.innerHTML = `${keyword}: ${info}`;
    document.body.appendChild(popup);

    // Remove the popup after 5 seconds
    setTimeout(() => {
        popup.remove();
    }, 5000);
}

// Monitor the email body for keyword matches (trigger this when composing an email)
function monitorEmailBody() {
    Office.context.mailbox.item.body.getAsync('text', function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            checkForKeyword(result.value); // Check the email content for keywords
        }
    });
}

// Call the function to fetch and load the keywords when the add-in is initialized
fetchKeywordData();

// You may want to set up event listeners for when the email body is being typed or changed
monitorEmailBody();
