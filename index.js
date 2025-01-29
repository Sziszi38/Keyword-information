let keywordList = {};

function fetchKeywordData() {
    fetch('https://sziszi38.github.io/Keyword-information/keywords.xml')
        .then(response => response.text())
        .then(data => parseXML(data))
        .catch(error => console.error('Error fetching XML:', error));
}

function parseXML(xmlString) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, "application/xml");

    const keywords = xmlDoc.getElementsByTagName("keyword");

    for (let i = 0; i < keywords.length; i++) {
        const keyword = keywords[i].getAttribute('name');
        const info = keywords[i].getElementsByTagName('info')[0].textContent;
        keywordList[keyword] = info;
    }

    displayKeywordList();
}

function displayKeywordList() {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = '';

    Object.keys(keywordList).forEach(keyword => {
        const info = keywordList[keyword];
        const keywordItem = document.createElement('div');
        keywordItem.textContent = `${keyword}: ${info}`;
        resultDiv.appendChild(keywordItem);
    });
}

document.getElementById('keywordForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const keyword = document.getElementById('keyword').value;
    const popupWord = document.getElementById('popupWord').value;

    keywordList[keyword] = popupWord;

    displayKeywordList();

    document.getElementById('keyword').value = '';
    document.getElementById('popupWord').value = '';
});

function checkForKeyword(emailText) {
    for (const keyword in keywordList) {
        const regex = new RegExp(`\\b${keyword}\\w*\\b`, 'g');

        if (regex.test(emailText)) {
            displayPopup(keyword, keywordList[keyword]);
        }
    }
}

function displayPopup(keyword, info) {
    const popup = document.createElement('div');
    popup.className = 'popup';
    popup.innerHTML = `${keyword}: ${info}`;
    document.body.appendChild(popup);

    setTimeout(() => {
        popup.remove();
    }, 5000);
}

function monitorEmailBody() {
    Office.context.mailbox.item.body.getAsync('text', function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            checkForKeyword(result.value);
        }
    });
}

fetchKeywordData();
monitorEmailBody();
