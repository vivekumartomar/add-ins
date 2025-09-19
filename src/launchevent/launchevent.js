Office.onReady(() => {
    console.log("Add-in is ready for launch events");
  });
  
  // This runs when user clicks Send
  function onMessageSendHandler(event) {
    console.log("User clicked Send button!");
  
    // Example: get subject
    
    const item = Office.context.mailbox.item;
    console.log("Subject:", item.subject);
  
    // Always call event.completed() to allow sending
    event.completed();
  }
  
  // Export for classic Outlook (needed)
  if (typeof module !== "undefined") {
    module.exports = {
      onMessageSendHandler
    };
  }
  