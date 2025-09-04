// Centralized LoggerApp object
const LoggerApp = {
  info: (...args) => console.info("[INFO]", ...args),
  warn: (...args) => console.warn("[WARN]", ...args),
  error: (...args) => console.error("[ERROR]", ...args),
  debug: (...args) => console.debug("[DEBUG]", ...args)
};

let creating_meeting_loading_inner_html, 
    additional_seats_inner_html,
    status_cancelation,
    status_confirming,
    confirmation_message,
    cancelation_message;

// Auth module to encapsulate token retrieval logic
const Auth = {
  acquireGraphToken: function () {
    const now = Date.now();
    const tokenDataRaw = sessionStorage.getItem("sso_token_data");

    if (tokenDataRaw) {
      try {
        const tokenData = JSON.parse(tokenDataRaw);
        if (tokenData.expiry && tokenData.token && now < tokenData.expiry - 5 * 60 * 1000) {
          LoggerApp.info("✅ Using cached access token");
          return Promise.resolve(tokenData.token);
        }
      } catch (e) {
        LoggerApp.warn("⚠️ Failed to parse cached token, falling back to popup");
      }
    }

    const clientId = window.appConfig.clientId;
    const redirectUri = encodeURIComponent(window.appConfig.redirectUri);
    const scope = encodeURIComponent(window.appConfig.scope.join(" "));
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=${scope}&response_mode=fragment&prompt=select_account`;

    return new Promise((resolve, reject) => {
      const authWindow = window.open(authUrl, "_blank", "width=600,height=600");

      if (!authWindow) {
        reject(new Error("Popup was blocked by the browser"));
        return;
      }

      const timeout = setTimeout(() => {
        window.removeEventListener("message", receiveMessage);
        reject(new Error("Login timed out"));
      }, 60000);

      const receiveMessage = (event) => {
        if (event.data.message === "token") {
          clearTimeout(timeout);
          window.removeEventListener("message", receiveMessage);
          const accessToken = event.data.accessToken;
          const expiry = now + 60 * 60 * 1000;
          sessionStorage.setItem("sso_token_data", JSON.stringify({ token: accessToken, expiry }));
          LoggerApp.info("Access token acquired and stored");
          resolve(accessToken);
        }
      };

      window.addEventListener("message", receiveMessage);
    });
  },

  acquireTokenWithDialog: function () {
    const clientId = window.appConfig.clientId;
    const redirectUri = window.appConfig.redirectUri;
    const scope = window.appConfig.scope.join(" ");
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scope)}&response_mode=fragment`;

    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(authUrl, { width: 30, height: 50 }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error("Dialog API failed: " + asyncResult.error.message));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          const hash = arg.message;
          const params = new URLSearchParams(hash.substring(1));
          const token = params.get('access_token');
          if (token) {
            resolve(token);
          } else {
            reject(new Error("No token found"));
          }
          dialog.close();
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          reject(new Error("Dialog closed prematurely"));
        });
      });
    });
  }
};

// Graph module to isolate Graph API calls
const Graph = {
  getBusinessUnit: async function (token) {
    const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=department", {
      headers: { Authorization: `Bearer ${token}` }
    });
    const data = await response.json();
    const department = data.department || "Not specified";
    const businessUnitInput = document.getElementById("businessUnit");
    if (businessUnitInput) {
      businessUnitInput.value = department;
      businessUnitInput.readOnly = !!data.department;
    }
  },

  getAmenitiesFromSharePoint: async function (token) {
      const siteId = window.appConfig.siteID;
      const listId = window.appConfig.listAmenitiesID;
      const items = await Graph.getListItemsFromSharePoint(token, siteId, listId);
      //return items.map(item => item.fields.Amenity);
      return items.map(item => ({
                name: item.fields.Amenity,
                id: item.fields.AmenityID
              }));
  },
  getListItemsFromSharePoint: async function (token, siteId, listId, select = "fields") {
      const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=${select}`;
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json"
        }
      });

      if (!response.ok) {
        throw new Error(`Error fetching list items: ${response.statusText}`);
      }

      const data = await response.json();
      return data.value;
   },
   getCalendarId: async function (token, userPrincipalName) {
      const url = `https://graph.microsoft.com/v1.0/users/${userPrincipalName}/calendars`;
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json"
        }
      });

      if (!response.ok) {
        throw new Error(`Error fetching calendars: ${response.statusText}`);
      }

      const data = await response.json();
      const mainCalendar = data.value.find(cal => cal.name === "Calendar");

      if (!mainCalendar) {
        throw new Error("Default 'Calendar' not found.");
      }
      LoggerApp.debug("Main calendar ID:", mainCalendar.id);
      window.cachedCalendarId = mainCalendar.id;
      return mainCalendar.id;
    }
};

// Helper function to retrieve all attachments with their content
const getAttachmentsWithContent = (item) => {
  return new Promise((resolve) => {
    item.getAttachmentsAsync((attachmentResult) => {
      if (attachmentResult.status !== Office.AsyncResultStatus.Succeeded) {
        LoggerApp.warn("Failed to get attachments:", attachmentResult.error.message);
        resolve([]);
        return;
      }

      const attachments = attachmentResult.value || [];
      if (attachments.length === 0) {
        resolve([]);
        return;
      }

      Promise.all(
        attachments.map(
          (att) =>
            new Promise((res) => {
              item.getAttachmentContentAsync(att.id, { asyncContext: att }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  res({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    name: att.name,
                    contentType: result.value.contentType,
                    contentBytes: result.value.content
                  });
                } else {
                  LoggerApp.warn("Failed to get attachment content:", result.error.message);
                  res(null);
                }
              });
            })
        )
      ).then((res) => resolve(res.filter(Boolean)));
    });
  });
};

Office.onReady(info => {
  console.log("CACHE TRACKER: UAT TEST: 123");
 if (info.host === Office.HostType.Outlook) {
    const addinContainer = document.getElementById("add-in-container");
    const envStatus = document.getElementById("env-status");
    const devToggle = document.getElementById("devModeToggle");
    const uatEnvToggle = document.getElementById("uatEnvToggle");
    const uatEnvToggleText = document.getElementById("uatEnvToggleText");
    const jsonOutput = document.getElementById("jsonOutput");
    let endpointFlow1Url, endpointFlow2Url;
    let tag = "[ Smart Meeting Draft ]";

  // Add the Smart Meeting Draft to the subject, with UAT prefix by default
  updateSubjectWithTag("UAT");
  
  // Listen to the ENV toggle and change the Enviroment footer status bar / Update the subject
  uatEnvToggle.addEventListener("change", function () {
    if (uatEnvToggle.checked) {
      uatEnvToggleText.innerText = "UAT";
      addinContainer.classList.remove('green-gradient');
      addinContainer.classList.add('red-gradient');
      envStatus.innerText = "UAT Enviroment"
      updateSubjectWithTag("UAT");
    } else {
      uatEnvToggleText.innerText = "DEV";
      addinContainer.classList.remove('red-gradient');
      addinContainer.classList.add('green-gradient');
      envStatus.innerText = "DEV Enviroment"
      updateSubjectWithTag("DEV");
    }
  });

  // Get Translation from the JSON 
  function setAddinLanguage(lang) {
    //lang is outlookLanguage, here we get the first two letters from the outlooklanguage string "en-US" will be "en"
    //lang = "fr";
    let selectedLang;
    if (lang.includes("-nav")) {
      // HTML nav format, e.g., "fr-nav"
      selectedLang = lang.substring(0, 2);
      sessionStorage.setItem("selectedLanguage", selectedLang + "-Session"); // override session
    } else {
      // Outlook and session format, e.g., "fr-BE or fr-BE-Session"
      selectedLang = lang.substring(0, 2);
    }

    async function loadResources(selectedLang) {
      try {
        const res = await fetch(`${selectedLang}.json`);
        if (!res.ok) throw new Error(`Missing ${selectedLang}.json`);
        return await res.json();
      } catch (error) {
        console.warn(`Falling back to en.json due to: ${error.message}`);
        const fallbackRes = await fetch(`en.json`);
        if (!fallbackRes.ok) throw new Error("Fallback en.json also missing!");
        return await fallbackRes.json();
      }
    }
    // Find the HTML elements and apply the translation
    function applyTranslations(resources) {
      // ensure key messages are available even if no matching DOM element exists
      confirmation_message = resources["confirmation.message"];
      cancelation_message = resources["cancelation.message"];
      status_confirming = resources["status.confirming"];
      status_cancelation = resources["status.cancelation"];

      document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');

        if (key === "creating.meeting.loading"){
          creating_meeting_loading_inner_html = resources[key];
        } else {
          el.innerText = resources[key] || key; // fallback to key if missing
        }

        if (key === "additional.seats") { additional_seats_inner_html = resources[key]; }
        if (key === "confirmation.message") { confirmation_message = resources[key]; }
        if (key === "cancelation.message") { cancelation_message = resources[key]; }
        if (key === "status.confirming") { status_confirming = resources[key]; }
        if (key === "status.cancelation") { status_cancelation = resources[key]; }
      });
    }

    loadResources(selectedLang).then(function (resources) {
     applyTranslations(resources);
    });
  }
    
    // Read the Language from the Session Storage
    const savedLang = sessionStorage.getItem('selectedLanguage'); 
    // Determine the current language earlier so it can update everything in time
    const item = Office.context.mailbox.item;
    let outlooklanguage;
    
    if (savedLang){
      outlooklanguage = savedLang;
    } else {
      outlooklanguage = Office.context.displayLanguage;
    }
    LoggerApp.info("Outlook display language: ", outlooklanguage);

    // If the language stored within the session, set the language and the navigator class accorindgly
    if(savedLang){
        try {
          setAddinLanguage(outlooklanguage);
          // Add the class to the clicked link
          document.getElementById('link-' + savedLang.substring(0, 2)).classList.add('active-language');
        } catch (e) {
          LoggerApp.warn("⚠️ No Language found in the Session Storage");
        }
    }
  
  document.getElementById('language-bar').addEventListener('click', function(event) {
    const link = event.target.closest('a');
    if (link && this.contains(link)) {
      event.preventDefault();
      const selectedLang = link.getAttribute('data-lang');

      // Save to sessionStorage
      // sessionStorage.setItem('selectedLanguage', selectedLang);
      setAddinLanguage(selectedLang);
      // Remove the class from all links
      document.querySelectorAll('#language-bar a').forEach(a => a.classList.remove('active-language'));
      // Add the class to the clicked link
      link.classList.add('active-language');
    }
  });
  // End of Localization Language Functionality 

    (async () => {
      let token;
      try {
        token = await Auth.acquireGraphToken();
      } catch (error) {
        LoggerApp.warn("Standard token acquisition failed, falling back to Dialog API:", error);
        token = await Auth.acquireTokenWithDialog();
      }

      if (token) {
        try {
          await Graph.getBusinessUnit(token);
          const amenitiesList = await Graph.getAmenitiesFromSharePoint(token);
          LoggerApp.debug("Amenities: ", amenitiesList);
          renderAmenitiesCheckboxes(amenitiesList);
          setAddinLanguage(outlooklanguage);
          const calendarId = await Graph.getCalendarId(token, Office.context.mailbox.userProfile.emailAddress); 
        } catch (err) {
          LoggerApp.error("Failed to fetch business unit or amenities:", err);
        }
      }
    })();
	
  /* 
    Append [Smart Meeting Draft] to the subject. This is currently for testing purposes only. 
    The idea is to add a tag (in the subject or body) to mark a Smart Meeting event immediately after 
    interacting with the add-in. This marker could potentially be used by Smart Alert to detect cases 
    where the user tries to press buttons like "Send", which could result in duplicate invitations.
  */  
    
    function updateSubjectWithTag(tag) {
      const item = Office.context.mailbox.item;
    
      item.subject.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const currentSubject = result.value.replace(/(UAT|DEV) \[ Smart Meeting Draft \]/, "").trim() || "";
          
          if (!currentSubject.includes(tag.trim() + "[ Smart Meeting Draft ]")) {
            const newSubject = tag + " [ Smart Meeting Draft ] " + currentSubject;
    
            item.subject.setAsync(newSubject, function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                LoggerApp.info("Subject updated successfully.");
              } else {
                LoggerApp.error("Failed to update subject:", asyncResult.error.message);
              }
            });
          } else {
            LoggerApp.info("Subject already contains tag. No update needed.");
          }
        } else {
          LoggerApp.error("Failed to get current subject:", result.error.message);
        }
      });
    }
    
    // Close Event Planner Window after the Confirmation
    const closeBtn = document.getElementById("closeEventPlannerWindow");
    const closeOverlay = document.getElementById("closeConfirmOverlay");
    const closeYes = document.getElementById("confirmCloseYes");
    const closeNo = document.getElementById("confirmCloseNo");
    
    const settingsGearButton = document.getElementById("settings-gear-button");
    const settingsOverlay = document.getElementById("settingsOverlay");
    const closeSettingsButton = document.getElementById("closeSettingsButton");

    if (closeBtn && closeOverlay && closeYes && closeNo) {
      closeBtn.addEventListener("click", function() {
        closeOverlay.style.display = "flex";
      });

      closeYes.addEventListener("click", function() {
        closeOverlay.style.display = "none";
        closeEventFunction();
      });

      closeNo.addEventListener("click", function() {
        closeOverlay.style.display = "none";
      });
    }

    function closeEventFunction() {
      Office.context.mailbox.item.close();
    }
    
    // Settings Overlay
    if (settingsGearButton && settingsOverlay && closeSettingsButton) {
      settingsGearButton.addEventListener("click", function() {
        settingsOverlay.style.display = "flex";
      });

      closeSettingsButton.addEventListener("click", function() {
        settingsOverlay.style.display = "none";
      });
    }

    // Utility function: ensures number input is between min and max
    function validateInputNum(input, min, max) {
      let v = parseInt(input.value, 10) || 0;
      v = Math.max(min, Math.min(max, v));
      input.value = v;
      return v;
    }
	
    // Remove any previous event listener and attach new one to the main button
    const btn = document.getElementById("generateJson");
    btn.replaceWith(btn.cloneNode(true));
    const newBtn = document.getElementById("generateJson");

    newBtn.addEventListener("click", function (e) {
      e.preventDefault();

      collectAndShowData(function(payload) {
        if (devToggle.checked) {
          jsonOutput.textContent = "⏳"+ creating_meeting_loading_inner_html;

          if(uatEnvToggle.checked){
            endpointFlow1Url = window.appConfig.endpointFlow1Url;
            endpointFlow2Url = window.appConfig.endpointFlow2Url;
          }else {
            endpointFlow1Url = window.appConfig.DEVendpointFlow1Url;
            endpointFlow2Url = window.appConfig.DEVendpointFlow2Url;
          }

          fetch(endpointFlow1Url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
          })
          .then(response => {
            if (!response.ok)
              throw new Error("Flow1 error: HTTP status " + response.status);

            return response.text().then(txt => {
              LoggerApp.info("Flow1 RAW response:", txt);

              let result;
              try {
                result = JSON.parse(txt);
              } catch {
                throw new Error("Flow1 did not return valid JSON: " + txt);
              }

              if (!result["Placeholder Meeting Room Name"]) {
                throw new Error("Missing 'Placeholder Meeting Room Name' in Flow1 response: " + txt);
              }

              showConfirmationUI(result["Placeholder Meeting Room Name"], result);
            });
          })
          .catch(err => {
            jsonOutput.textContent = "Error sending to Flow: " + err;
            // Keep form data so the user can correct/resubmit
          });

        } else {
          jsonOutput.textContent = JSON.stringify(payload, null, 2);
        }
      });
    });

  // Additional new functions:
	function renderAmenitiesCheckboxes(amenitiesList) {
	  const container = document.getElementById('amenitiesContainer');
	  container.innerHTML = '';
	  container.style.display = 'flex';
	  container.style.flexWrap = 'wrap';
	  container.style.gap = '14px';

	  const checkboxGroup = document.createElement('div');
	  checkboxGroup.style.display = 'flex';
	  checkboxGroup.style.flexWrap = 'wrap';
	  checkboxGroup.style.gap = '14px';

	  amenitiesList.forEach((amenity) => {
		const label = document.createElement('label');
		label.className = 'ms-Checkbox';

		const checkbox = document.createElement('input');
		checkbox.type = 'checkbox';
		checkbox.name = 'amenities';
		checkbox.value = amenity.id;
		checkbox.checked = false;

		const span = document.createElement('span');
		span.className = 'ms-Checkbox-label';
		span.textContent = amenity.name;

		label.appendChild(checkbox);
		label.appendChild(span);
		checkboxGroup.appendChild(label);
	  });

	  container.appendChild(checkboxGroup);

	  const additionalSeatsDiv = document.createElement('div');
	  additionalSeatsDiv.style = 'margin-top:10px; display:flex; align-items:center; gap:8px; width:100%;';

	  const seatsLabel = document.createElement('label');
	  seatsLabel.className = 'ms-Label';
	  seatsLabel.textContent = additional_seats_inner_html || "Additional seats";
	  seatsLabel.setAttribute('for', 'additionalSeats');
    seatsLabel.setAttribute('data-i18n', 'additional.seats'); // Added for localization
	  seatsLabel.style.marginBottom = '0';
    
	  const seatsInput = document.createElement('input');
	  seatsInput.type = 'number';
	  seatsInput.id = 'additionalSeats';
	  seatsInput.min = '0';
	  seatsInput.max = '20';
	  seatsInput.value = '0';
	  seatsInput.style.width = '40px';

	  additionalSeatsDiv.appendChild(seatsLabel);
	  additionalSeatsDiv.appendChild(seatsInput);

	  container.appendChild(additionalSeatsDiv);
	}
	
  // Displays the confirmation (pre-action) UI and wires the handlers
function showConfirmationUI(roomName, flowResponse = {}) {
  const form     = document.getElementById('meetingForm');
  const section  = document.getElementById('confirmationSection');

  // 1) Show the confirmation section, hide the form
  form.style.display = 'flex';   // keep flex for gap layout when you come back
  form.style.display = 'none';
  section.style.display = 'block';

  // 2) Reset visual state to PRE-ACTION
  section.classList.remove('show-result', 'is-cancel', 'is-error');
  section.classList.add('is-success'); // neutral/success tint for pre-action

  // Title + messages
  const titleEl = document.getElementById('confirmationTitle');
  const msgEl   = document.getElementById('confirmationMessage');
  const status  = document.getElementById('confirmationStatus');
  titleEl.textContent = 'Temporary placeholder reserved!';
  msgEl.style.display = '';       // show instructions paragraph
  status.textContent  = '';       // clear any previous result

  // Room block (make sure it is visible for pre-action)
  const roomBlock = section.querySelector('.room-block');
  if (roomBlock) roomBlock.style.display = '';

  // 3) Fill room info (inline + block link)
  const roomInline = document.getElementById('placeholderRoomName');
  const roomLinkEl = document.getElementById('placeholderRoomName_link');
  if (roomInline) roomInline.textContent = roomName || '';
  if (roomLinkEl) roomLinkEl.textContent = roomName || '';

  // Optional room URL + building (if provided by Flow1)
  const link = document.getElementById('placeholderRoomLink');
  const buildingEl = document.getElementById('placeholderBuilding');
  const roomUrl = flowResponse.RoomUrl || flowResponse.RoomURL || '';
  if (link) {
    link.href = roomUrl || '#';
    link.style.pointerEvents = roomUrl ? 'auto' : 'none';
  }
  if (buildingEl) {
    const b = flowResponse.BuildingName || flowResponse.Building || '';
    buildingEl.textContent = b;
    buildingEl.style.display = b ? '' : 'none';
  }

  // 4) Hidden fields for Flow2
  document.getElementById('confirmEventID').value     = flowResponse.EventID     || '';
  document.getElementById('confirmDataverseID').value = flowResponse.DataverseID || '';
  document.getElementById('confirmIcalUid').value     = flowResponse.iCalUid     || '';
  document.getElementById('organizerResponse').value  = '';

  // 5) Reset CTA row + follow-up row
  const ctas = document.getElementById('confirmCancelBtns');
  ctas.style.display = 'flex';
  const backRow = document.getElementById('backBtnDiv');
  backRow.style.display = 'none';

  // 6) Enable and wire buttons
  const confirmBtn = document.getElementById('confirmBtn');
  const cancelBtn  = document.getElementById('cancelBtn');
  confirmBtn.disabled = false;
  cancelBtn.disabled  = false;

  confirmBtn.onclick = function () {
    document.getElementById('organizerResponse').value = 'Confirmed';
    handleConfirmation();
  };
  cancelBtn.onclick = function () {
    document.getElementById('organizerResponse').value = 'Cancel';
    handleConfirmation();
  };

  // 7) UX: ensure the section is in view (optional)
  try { section.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch (_) {}
}

    // Handles confirmation or cancellation of meeting room
    function handleConfirmation() {
      
      const statusDiv = document.getElementById("confirmationStatus");

      document.getElementById('confirmBtn').disabled = true;
      document.getElementById('cancelBtn').disabled = true;
      document.getElementById('confirmCancelBtns').style.display = "none";
      // statusDiv.textContent = (document.getElementById("organizerResponse").value === "Confirmed" ? "Confirming..." : "Cancelling...");
      statusDiv.textContent =
        document.getElementById("organizerResponse").value === "Confirmed"
          ? status_confirming || "Confirming..."
          : status_cancelation || "Cancelling...";

          const item = Office.context.mailbox.item;
          let payload;

          getAttachmentsWithContent(item)
            .then(attachments => {
              payload = {
                EventID: document.getElementById("confirmEventID").value,
                DataverseID: document.getElementById("confirmDataverseID").value,
                iCalUid: document.getElementById("confirmIcalUid").value || "",
                OrganizerResponse: document.getElementById("organizerResponse").value,
                Attachments: attachments
              };

              LoggerApp.info("EventID:" + document.getElementById("confirmEventID").value);

              return fetch(endpointFlow2Url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
              });
            })
            .then(response => response.text())
            .then(result => {
				  const section = document.getElementById('confirmationSection');
				  const isConfirm = payload.OrganizerResponse === "Confirmed";
				
				  document.getElementById('confirmationTitle').textContent =
				    isConfirm ? 'Reservation confirmed' : 'Reservation cancelled';
				
				  document.getElementById('confirmationStatus').textContent =
				    isConfirm
				      ? (confirmation_message || 'Reservation confirmed! Meeting room is now visible to all attendees.')
				      : (cancelation_message || 'Reservation cancelled. The placeholder has been removed.');
				
				  // hide pre-action blocks for the result view
				  document.getElementById('confirmationMessage').style.display = 'none';
				  const roomBlock = section.querySelector('.room-block');
				  if (roomBlock) roomBlock.style.display = 'none';
				  document.getElementById('confirmCancelBtns').style.display = 'none';
				
				  section.classList.toggle('is-success', isConfirm);
				  section.classList.toggle('is-cancel', !isConfirm);
				  section.classList.add('show-result');         // if you have CSS tied to this
				
				  document.getElementById('backBtnDiv').style.display = 'flex';
				  document.getElementById('backBtn').onclick = resetToForm;
				})
            .catch(err => {
              statusDiv.textContent = "Error: " + err;
              document.getElementById('backBtnDiv').style.display = 'block';
              document.getElementById('backBtn').onclick = resetToForm;
            });
        }

    // Resets UI back to the initial form
    function resetToForm() {
	  document.getElementById('confirmCancelBtns').style.display = 'flex';
	  document.getElementById('cancelBtn').style.display = 'flex';
	  document.getElementById('confirmBtn').style.display = 'flex';
		
      document.getElementById('meetingForm').style.display = 'flex';
      document.getElementById('confirmationSection').style.display = 'none';
      jsonOutput.textContent = "";
    }

    // Collects all required meeting data
     function collectAndShowData(callback) {
      Office.context.mailbox.item.subject.getAsync(function(subjectResult) {
        const subject = subjectResult.value;
        let organizerEmail, calendarId;
        if(uatEnvToggle.checked){
          organizerEmail = Office.context.mailbox.userProfile.emailAddress;
          calendarId = window.cachedCalendarId || "";
        }else {
          organizerEmail = "roberto.campione.ext@proximusuat.com";
          calendarId = "AAMkAGExZGQxMDAxLTg1ZjktNGMwNy1iMzg4LTY0OTAxYWFjYmUzMQBGAAAAAAAEAAQ7FEmVSotJoNetmaQ4BwBUjqXA3xWGTapiAE_7zGePAAAAAAEGAAAifHiMQ6ciS5lzQ5O2Zt71AAAqgIKzAAA=";
        }
        LoggerApp.debug("cachedCalendarId:", calendarId);

        Office.context.mailbox.item.requiredAttendees.getAsync(function(requiredResult) {
          const requiredAttendeesArr = (requiredResult.value || []).map(a => a.emailAddress);

          Office.context.mailbox.item.optionalAttendees.getAsync(function(optionalResult) {
            const optionalAttendeesArr = (optionalResult.value || []).map(a => a.emailAddress);

            Office.context.mailbox.item.start.getAsync(function(startResult) {
              const startDateTime = startResult.value;
              Office.context.mailbox.item.end.getAsync(function(endResult) {
                const endDateTime = endResult.value;

                Office.context.mailbox.item.recurrence.getAsync(function(recurrenceResult) {
                  const recurrence = recurrenceResult.status === Office.AsyncResultStatus.Succeeded
                    ? recurrenceResult.value : null;

                  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
                    const body = (bodyResult.value || "");
                    const businessUnit = document.getElementById("businessUnit")?.value || "";
                    const onBehalfOf = document.getElementById("onBehalfOf")?.checked || false;
                    const additionalSeats = validateInputNum(document.getElementById("additionalSeats"), 0, 20);
                    const extraTimeBefore = validateInputNum(document.getElementById("extraTimeBefore"), 0, 15);
                    const extraTimeAfter = validateInputNum(document.getElementById("extraTimeAfter"), 0, 15);
                    const building = document.getElementById("building").value;
                    const amenities = Array.from(document.querySelectorAll("input[name='amenities']:checked")).map(el => el.value).join(",");
                    const meetingType = document.querySelector("input[name='meetingType']:checked")?.value || "";

                    let sensitivity = "unknown";
                    let isPrivate = false;
                    
                    try {
                      sensitivity = Office.context.mailbox.item.sensitivity || "unknown";
                      isPrivate = sensitivity === "private";
                    } catch (ex) {}

                    if (document.getElementById("isPrivate")?.checked) {
                      isPrivate = true;
                    }

                    // Handle recurrence patterns for compatibility
                    let Recurrence = "none";
                    let RecurrenceEndDate = "";
                    let RecurrenceDaysOfWeek = "";

                    if (recurrence?.recurrencePattern) {
                      Recurrence = recurrence.recurrencePattern.type || "";
                      RecurrenceDaysOfWeek = (recurrence.recurrencePattern.daysOfWeek || []).join(",");
                      RecurrenceEndDate = recurrence.recurrenceRange?.endDate || "";

                    } else if (recurrence?.recurrenceProperties) {
                      Recurrence = recurrence.recurrenceType || "";
                      RecurrenceDaysOfWeek = Array.isArray(recurrence.recurrenceProperties.days)
                        ? recurrence.recurrenceProperties.days.join(",") : "";
                      if (recurrence.seriesTime && recurrence.seriesTime.endYear) {
                        RecurrenceEndDate =
                          recurrence.seriesTime.endYear + "-" +
                          String(recurrence.seriesTime.endMonth).padStart(2, '0') + "-" +
                          String(recurrence.seriesTime.endDay).padStart(2, '0');
                      }
                    }
                    
                    getAttachmentsWithContent(Office.context.mailbox.item)
                      .then((attachments) => {
                        const payload = {
                          Title: subject,
                          Body: body,
                          Attachments: attachments,
                          Startdatetime: startDateTime,
                          Enddatetime: endDateTime,
                          // OrganizerEmailAddress: "roberto.campione.ext@proximusuat.com", //organizerEmail // to be replaced with dynamic mapping organizerEmail
                          OrganizerEmailAddress: organizerEmail,
                          //OrganizerCalendarID: "AAMkAGExZGQxMDAxLTg1ZjktNGMwNy1iMzg4LTY0OTAxYWFjYmUzMQBGAAAAAAAEAAQ7FEmVSotJoNetmaQ4BwBUjqXA3xWGTapiAE_7zGePAAAAAAEGAAAifHiMQ6ciS5lzQ5O2Zt71AAAqgIKzAAA=", //to be replaced with dynamic mapping calendarId
                          OrganizerCalendarID: calendarId,
                          BuildingName: building,
                          BusinessUnit: businessUnit,
                          MeetingType: meetingType,
                          Attendees: {
                            RequiredAttendees: requiredAttendeesArr,
                            OptionalAttendees: optionalAttendeesArr
                          },
                          isPrivate: isPrivate,
                          OnBehalfOf: onBehalfOf,
                          Recurrence,
                          RecurrenceEndDate,
                          RecurrenceDaysOfWeek,
                          ExtraFeatures: {
                            AdditionalSeats: additionalSeats,
                            Amenities: amenities,
                            ExtraTimeBefore: extraTimeBefore,
                            ExtraTimeAfter: extraTimeAfter
                          }
                        };
                        if (callback) callback(payload);
                      })
                      .catch((err) => {
                        LoggerApp.warn("Failed to get attachments:", err.message || err);
                        const payload = {
                          Title: subject,
                          Body: body,
                          Attachments: [],
                          Startdatetime: startDateTime,
                          Enddatetime: endDateTime,
                          OrganizerEmailAddress: organizerEmail,
                          OrganizerCalendarID: calendarId,
                          BuildingName: building,
                          BusinessUnit: businessUnit,
                          MeetingType: meetingType,
                          Attendees: {
                            RequiredAttendees: requiredAttendeesArr,
                            OptionalAttendees: optionalAttendeesArr
                          },
                          isPrivate: isPrivate,
                          OnBehalfOf: onBehalfOf,
                          Recurrence,
                          RecurrenceEndDate,
                          RecurrenceDaysOfWeek,
                          ExtraFeatures: {
                            AdditionalSeats: additionalSeats,
                            Amenities: amenities,
                            ExtraTimeBefore: extraTimeBefore,
                            ExtraTimeAfter: extraTimeAfter
                          }
                        };
                        if (callback) callback(payload);
                      });
                  });
                });
              });
            });
          });
        });
      });
    }
  }
});
