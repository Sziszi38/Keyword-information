Office.initialize = function (reason) {
  $(document).ready(function () {
    var item = Office.context.mailbox.item;
    var keywords = {};

    // Function to check for keywords in the email body
    function checkForKeywords(bodyText) {
      for (var keyword in keywords) {
        if (bodyText.includes(keyword)) {
          document.getElementById("result").innerText = "Pop-up Word: " + keywords[keyword];
          break;
        }
      }
    }

    // Get the email body and check for keywords
    if (item.body.getAsync) {
      item.body.getAsync("text", function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var bodyText = result.value;
          checkForKeywords(bodyText);
        }
      });
    }

    // Add new keyword and pop-up word
    document.getElementById("keywordForm").addEventListener("submit", function (event) {
      event.preventDefault();
      var keyword = document.getElementById("keyword").value;
      var popupWord = document.getElementById("popupWord").value;
      keywords[keyword] = popupWord;
      document.getElementById("keywordForm").reset();
      alert("Keyword and pop-up word added!");
    });
  });
};
