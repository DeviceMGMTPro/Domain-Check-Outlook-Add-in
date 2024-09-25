Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, checkRecipientDomains);
    }
  });
  
  function checkRecipientDomains(eventArgs) {
    const item = Office.context.mailbox.item;
    
    item.to.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const recipients = result.value;
        const domains = new Set();
  
        recipients.forEach((recipient) => {
          const email = recipient.emailAddress;
          const domain = email.split('@')[1];
          domains.add(domain);
        });
  
        if (domains.size > 1) {
          const domainList = Array.from(domains).join(", ");
          const message = `Warning: You are sending this email to multiple domains: ${domainList}. Do you want to proceed?`;
  
          if (!confirm(message)) {
            eventArgs.completed({ allowEvent: false }); // Block sending
          } else {
            eventArgs.completed({ allowEvent: true }); // Allow sending
          }
        } else {
          eventArgs.completed({ allowEvent: true }); // Allow sending if one domain
        }
      } else {
        console.error(result.error.message);
        eventArgs.completed({ allowEvent: true }); // Fail-safe
      }
    });
  }
  