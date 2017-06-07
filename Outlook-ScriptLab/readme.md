Disclaimer: This project has nothing to do with ScriptLab - https://github.com/OfficeDev/script-lab

It's sample read\compose mail add-in that allows you to test offcejs APIs.
Here is a list of couple of such commands -


/* Get Email Subject */
if (Office.context.mailbox.item.subject.getAsync) {
  Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
    if (asyncResult.status == "failed") {
      showNotification("Action failed with error: " + asyncResult.error.message);
    }
    else {
      showNotification("Composed subject is: " + asyncResult.value);
    }
  });
}
else {
  showNotification("Subject is: " + Office.context.mailbox.item.subject);
}

/* Get Email Recepient */
if (Office.context.mailbox.item.to.getAsync) {
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    if (asyncResult.status == "failed") {
      showNotification("Action failed with error: " + asyncResult.error.message);
    }
    else {
      showNotification("First Recipient is: " + asyncResult.value[0].displayName);
    }
  });
}
else {
  showNotification("Recipient is: " + Office.context.mailbox.item.to[0].displayName);
}



/* Set Email Body */
Office.context.mailbox.item.body.prependAsync("Go Hawks!", function (asyncResult) {
  if (asyncResult.status == "failed") {
    showMessage("Action failed with error: " + asyncResult.error.message);
  }
});



/* Set Email Subject */
Office.context.mailbox.item.subject.setAsync("New subject!!!!!!!!!!!!!!!!!!!!", function (asyncResult) {
  if (asyncResult.status == "failed") {
    showMessage("Action failed with error: " + asyncResult.error.message);
  }
});



/* Set Email Recepient */
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );


