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

function virtualAttendance(event) {
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

  // STEP 1: Extract eventId from the body using marker
  const eventId = await new Promise((resolve) => {
    item.body.getAsync("text", function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const body = result.value;
        const match = body.match(/Smart Meeting Event ID[:\s]*([A-Za-z0-9+=\/_\\-]+)/i);
        const extractedId = match ? match[1] : null;
        console.log("Extracted Event ID from marker:", extractedId);
        resolve(extractedId);
      } else {
        console.error("Failed to read body for marker:", result.error.message);
        resolve(null);
      }
    });
  });

  // STOP if no marker is found
  if (!eventId) {
    statusUpdate("Icon.16x16", "❌ This is not a Proximus Smart Meeting. Interaction with this element is not possible.");
    console.warn("No valid Smart Meeting Event ID found. Aborting sendAttendance.");
    return;
  }

  const payload = {
    body: {
      "Contextual-add-in": {
        "meetingId": eventId,
        "response": "Accepted",
        "attendanceMode": mode,
        "respondent": {
          "name": profile.displayName,
          "email": profile.emailAddress
        },
        "timestamp": new Date().toISOString(),
        "organizerMail": item.organizer?.emailAddress || item.from?.emailAddress || null
      }
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
      statusUpdate("Icon.16x16", `✅ Your attendance has been recorded as ${mode}. You can change it at any time before the meeting.`);
      
      // Accept the meeting after successful confirmation
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

/*TESTING ROBERTO*/
async function getUserProfile(event) {
	try {
		console.log("Host:", Office.context.diagnostics.host);
		console.log("Platform:", Office.context.diagnostics.platform);
		console.log("OfficeRuntime.auth:", typeof OfficeRuntime.auth);
		
		const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
		console.log("Access token acquired:", accessToken);

		const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
		  method: "GET",
		  headers: {
			Authorization: `Bearer ${accessToken}`
		  }
		});

		if (graphResponse.ok) {
		  const profile = await graphResponse.json();
		  console.log("Graph profile response:", profile);

		  Office.context.mailbox.item.notificationMessages.addAsync("profileMsg", {
			type: "informationalMessage",
			message: `👤 ${profile.displayName}`,
			icon: "Icon.16x16",
			persistent: false
		  });
		} else {
		  console.error("Graph error:", await graphResponse.text());
		  Office.context.mailbox.item.notificationMessages.addAsync("profileMsg", {
			type: "errorMessage",
			message: "❌ Failed to fetch user info from Graph"
		  });
		}
	  } catch (error) {
		console.error("SSO error:", error);

		  const message = error.message || JSON.stringify(error);

		  Office.context.mailbox.item.notificationMessages.addAsync("profileMsg", {
			type: "errorMessage",
			message: `❌ SSO failed: ${message.substring(0, 150)}`
		  });
	  } finally {
		event.completed();
	  }
	}


// mandatory for ExecuteFunction
if (typeof Office !== "undefined") {
  Office.actions.associate("inPersonAttendance", inPersonAttendance);
  Office.actions.associate("virtualAttendance", virtualAttendance);
}
