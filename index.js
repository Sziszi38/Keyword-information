Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize the add-in
    initializeAddIn();
  }
});

function initializeAddIn() {
  // Add an event listener for the ItemChanged event
  Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
}

function onItemChanged(eventArgs) {
  // Get the body of the email
  Office.context.mailbox.item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailBody = result.value;
      // Check if the keyword exists in the email body
      if (emailBody.includes("Valeo")) {
        // Perform the desired action
        showKeywordDetectedMessage();
      }
    } else {
      console.error("Failed to get email body:", result.error);
    }
  });
}

function showKeywordDetectedMessage() {
  // Display a message to the user
  Office.context.mailbox.item.notificationMessages.addAsync("keywordDetected", {
    type: "informationalMessage",
    message: "Keyword 'Valeo' detected in the email.",
    icon: "iconid",
    persistent: true
  });
}
