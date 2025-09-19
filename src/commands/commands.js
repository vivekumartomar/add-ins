Office.onReady(() => {
  console.log("Add-in ready");
  console.log('test send 1');
});

// function onMessageSendHandler(event) {
//   const item = Office.context.mailbox.item;

//   // Read the subject
//   const subject = item.subject || "No subject";

//   // Alert it
//   alert("Subject: " + subject);

//   // Allow sending
//   event.completed({ allowEvent: true });
// }

Office.onReady(() => {
  console.log("Commands.js ready");
});

function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    item.subject.getAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const subject = asyncResult.value || "";

        console.log("Subject:", subject);

        if (subject.length < 3) {

          event.completed({
            allowEvent: false,
            errorMessage: "Subject must be at least 3 characters long.",
            errorMessageMarkdown: " **Subject too short**\n\nThe subject must be at least 3 characters.",
            cancelLabel: "Fix Subject",
            commandId: "msgComposeOpenPaneButton"
          });
          return;
        }

        event.completed({ allowEvent: true });

      } else {
        console.error("Failed to get subject:", asyncResult.error);
        event.completed({ allowEvent: true }); // fallback
      }
    });

  } catch (err) {
    console.error("Error in onMessageSendHandler:", err);
    event.completed({ allowEvent: true }); // fallback
  }
}





