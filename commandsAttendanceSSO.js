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

	function getAccessTokenViaPopup() {
	  return new Promise((resolve, reject) => {
		const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e8f07e89-4853-4f09-abd9-bbd398d2875c&response_type=token&redirect_uri=https%3A%2F%2Frobertocampione.github.io%2FOAMeetingAssistant%2Ffallbackauthdialog.html&scope=openid%20profile%20User.Read&response_mode=fragment&prompt=select_account`;

		const authWindow = window.open(authUrl, "_blank", "width=600,height=600");

		if (!authWindow) {
		  reject(new Error("Popup was blocked by the browser"));
		  return;
		}

		const timeout = setTimeout(() => {
		  window.removeEventListener("message", receiveMessage);
		  reject(new Error("Login timed out"));
		}, 60_000); // 60s timeout

		const receiveMessage = (event) => {
		  if (event.data.message === "token") {
			clearTimeout(timeout);
			window.removeEventListener("message", receiveMessage);
			resolve(event.data.accessToken);
		  }
		};

		window.addEventListener("message", receiveMessage);
	  });
	}

	async function getUserProfile(event) {
	  try {
		statusUpdate("Icon.16x16", "🔐 Waiting for login popup...");

		const token = await getAccessTokenViaPopup();
		console.log("Token received:", token);

		//  Request /me for displayName and officeLocation
		const meResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
		  headers: {
			Authorization: `Bearer ${token}`
		  }
		});
		const meData = await meResponse.json();
		console.log("/me response:", meData);

		const displayName = meData.displayName || "Not specified";
		const officeLocation = meData.officeLocation || "Not specified";

		// Request /me?$select=department
		const deptResponse = await fetch("https://graph.microsoft.com/v1.0/me?$select=department", {
		  headers: {
			Authorization: `Bearer ${token}`
		  }
		});
		const deptData = await deptResponse.json();
		console.log("/me?$select=department:", deptData);

		const department = deptData.department || "Not specified";

		// Request /me/mailboxSettings for language
		const langResponse = await fetch("https://graph.microsoft.com/v1.0/me/mailboxSettings", {
		  headers: {
			Authorization: `Bearer ${token}`
		  }
		});
		const langData = await langResponse.json();
		console.log("/me/mailboxSettings:", langData);

		//const language = langData.language?.locale || "Not specified";
		const officeLanguage = Office.context.displayLanguage || "Not specified";

		// Compose message
		const text = `👤 ${displayName}\n🌐 Language: ${officeLanguage}\n🏢 Location: ${officeLocation}\n📂 Department: ${department}`;
		console.log("Final message:", text);

		statusUpdate("Icon.16x16", text);

	  } catch (error) {
		console.error("Fallback auth error:", error);
		statusUpdate("Icon.16x16", `❌ Login failed: ${error.message || error}`);
	  } finally {
		if (event) event.completed();
	  }
	}

if (typeof Office !== "undefined") {
  Office.actions.associate("inPersonAttendance", inPersonAttendance);
  Office.actions.associate("virtualAttendance", virtualAttendance);
  Office.actions.associate("getUserProfile", getUserProfile);
}
