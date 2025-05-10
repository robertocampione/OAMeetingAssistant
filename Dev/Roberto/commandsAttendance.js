Office.onReady(function (info) {
  // Office.js is ready
});

function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function inPersonAttendance(event) {
  sendAttendance("inPerson").then(() => {
    event.completed();
  }).catch(error => {
    console.error("Error in inPersonAttendance:", error);
    event.completed();
  });
}

function onlineAttendance(event) {
  sendAttendance("online").then(() => {
    event.completed();
  }).catch(error => {
    console.error("Error in onlineAttendance:", error);
    event.completed();
  });
}

async function sendAttendance(mode) {
  const item = Office.context.mailbox.item;
  const profile = Office.context.mailbox.userProfile;
  
  const payload = {
    "Contextual-add-in": {
      "meetingId": item.itemId,
      "response": "Accepted",
      "attendanceMode": mode,
      "respondent": {
        "name": profile.displayName,
        "email": profile.emailAddress
      },
      "timestamp": new Date().toISOString(),
      "organizerMail": item.organizer?.emailAddress || item.from?.emailAddress || null
    }
  };

  console.log("Payload to send:", payload);

  try {
    const response = await fetch(window.appConfig.endpointFlow3Url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (response.ok) {
      statusUpdate("Icon.16x16", `Attendance sent: ${mode}`);
      
      // Dopo aver mandato la conferma a Flow3, accetta il meeting!
      if (item && item.calendar) {
        item.calendar.acceptAsync(function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Meeting accepted automatically.");
          } else {
            console.error("Failed to accept the meeting:", asyncResult.error.message);
          }
        });
      }

    } else {
      statusUpdate("Icon.16x16", `Error sending attendance`);
      console.error("Flow3 returned error:", await response.text());
    }
  } catch (error) {
	  statusUpdate("Icon.16x16", `Network error`);
	  console.error("SendAttendance failed:", error);
	  console.log("Flow3 endpoint was:", window?.appConfig?.endpointFlow3Url);
  }
}

// Obbligatorio per ExecuteFunction
if (typeof Office !== "undefined") {
  Office.actions.associate("inPersonAttendance", inPersonAttendance);
  Office.actions.associate("onlineAttendance", onlineAttendance);
}
