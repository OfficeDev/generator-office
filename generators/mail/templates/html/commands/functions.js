function setSubject(event){
  Office.context.mailbox.item.subject.setAsync('Hello world!', function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      Office.context.mailbox.item.notificationMessages.addAsync('setSubjectError', {
        type: 'errorMessage',
        message: 'Failed to set subject: ' + result.error
      });
      
      event.completed();
    }
    else {
      showMessage('Subject set', 'icon-16', event);
    }
  });

}

function getSubject(event){
  Office.context.mailbox.item.subject.getAsync(function(result){
    if (result.status === Office.AsyncResultStatus.Failed) {
      Office.context.mailbox.item.notificationMessages.addAsync('getSubjectError', {
        type: 'errorMessage',
        message: 'Failed to get subject: ' + result.error
      });
      
      event.completed();
    }
    else {
      showMessage('The current subject is: ' + result.value, 'icon-16', event);
    }
  });
}

function addToRecipients(event){
  var item = Office.context.mailbox.item;
  var addressToAdd = {
    displayName: Office.context.mailbox.userProfile.displayName,
    emailAddress: Office.context.mailbox.userProfile.emailAddress
  };

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    item.to.addAsync([addressToAdd], { asyncContext: event }, addRecipCallback);
  } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    item.requiredAttendees.addAsync([addressToAdd], { asyncContext: event }, addRecipCallback);
  }
}

function addRecipCallback(result) {
  var event = result.asyncContext;
  if (result.status === Office.AsyncResultStatus.Failed) {
    Office.context.mailbox.item.notificationMessages.addAsync('addRecipError', {
      type: 'errorMessage',
      message: 'Failed to add recipient: ' + result.error
    });
    
    event.completed();
  }
  else {
    showMessage('Recipient added', 'icon-16', event);
  }
}