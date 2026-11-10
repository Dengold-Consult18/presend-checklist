// Runs every time the user presses Send. Shows the checklist dialog; the email is sent only if
// every item is ticked and the user presses "Send email". Nothing is stored anywhere.

function getAsync(obj, method, ...args) {
  return new Promise((resolve, reject) => {
    obj[method](...args, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error.message)));
  });
}

async function collectDetails() {
  const item = Office.context.mailbox.item;
  const [to, cc, bcc, subject, attachments] = await Promise.all([
    getAsync(item.to, "getAsync"), getAsync(item.cc, "getAsync"), getAsync(item.bcc, "getAsync"),
    getAsync(item.subject, "getAsync"), getAsync(item, "getAttachmentsAsync"),
  ]);
  const fmt = (l) => l.map((r) => (r.displayName && r.displayName !== r.emailAddress ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress));
  return {
    to: fmt(to), cc: fmt(cc), bcc: fmt(bcc), subject,
    attachments: attachments.filter((a) => !a.isInline).map((a) => a.name),
  };
}

// Resolves true only if the user completed the checklist and confirmed.
function askUser(details) {
  return new Promise((resolve) => {
    const url = new URL("dialog.html", window.location.href).href;
    Office.context.ui.displayDialogAsync(url, { height: 75, width: 40, displayInIframe: true }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return resolve(false);
      const dialog = res.value;
      let done = false;
      const finish = (ok) => {
        if (done) return;
        done = true;
        try { dialog.close(); } catch (e) { /* already closed */ }
        resolve(ok);
      };
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let msg = {};
        try { msg = JSON.parse(arg.message); } catch (e) { /* ignore */ }
        if (msg.type === "ready") dialog.messageChild(JSON.stringify(details));
        else if (msg.type === "confirmed") finish(true);
        else if (msg.type === "cancel") finish(false);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish(false)); // closed with the X
    });
  });
}

async function onMessageSendHandler(event) {
  try {
    const ok = await askUser(await collectDetails());
    if (ok) event.completed({ allowEvent: true });
    else event.completed({ allowEvent: false, errorMessage: "The email was not sent because the pre-send checklist was not completed.", cancelLabel: "OK" });
  } catch (e) {
    event.completed({ allowEvent: false, errorMessage: "The pre-send checklist could not be shown, so the email was not sent. Please try again." });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
