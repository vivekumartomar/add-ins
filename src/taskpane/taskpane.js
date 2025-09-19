Office.onReady(() => {
  if (Office.context.mailbox.item) {
    // This is a compose item
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
});

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // Get subject
  const subject = item.subject;
  
  // Alert the subject
  alert("Subject: " + subject);

  // Indicate success so the email can send
  event.completed({ allowEvent: true });
}
