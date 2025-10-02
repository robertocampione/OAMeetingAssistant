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
    cancelation_message,
    no_meeting_found,
    show_technical_details,
    hide_technical_details,
    business_unit_amenities_error,
    createing_allocating_room_error,
    warning_24hour_meeting,
    sso_error;

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

        showOverlayMessage(
          "error",
          sso_error
        );
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
	return department;
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
      // console.log("Default Calendar: ", mainCalendar);
      const data = await response.json();
      // const mainCalendar = data.value.find(cal => cal.name === "Calendar" || "Agenda");

      const mainCalendar = data.value.find(cal => cal.isDefaultCalendar == true );
          
      if (!mainCalendar) {
        throw new Error("Default 'Calendar' not found.");
      }
      LoggerApp.debug("Main calendar ID:", mainCalendar.id);
      window.cachedCalendarId = mainCalendar.id;
      return mainCalendar.id;
    }
};

// Get the Business-Unit from PXS Employee API
// const getEmployeeBusinessUnit = {
//   get: async function () {
//     try {
//       const token = await Auth.acquireGraphToken();
//       // TODO: It will be replaced actual Employee API endpoint that returns the Business-Unit
//       const url = "https://TO_BE_ADDED_employee-api-url.PLACEHOLDER.com";
//       const response = await fetch(url, {
//         headers: { Authorization: `Bearer ${token}` }
//       });
//       if (!response.ok) {
//         throw new Error(`HTTP ${response.status}`);
//       }
//       const data = await response.json();
//       return data.businessUnit || "";
//     } catch (err) {
//       LoggerApp.warn("Failed to fetch business unit:", err);
//       return document.getElementById("businessUnit")?.value || "";
//     }
//   }
// };

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

function showLoading(kind /* 'finding' | 'confirming' | 'cancelling' */) {
  const form = document.getElementById('meetingForm');
  const loading = document.getElementById('loadingSection');
  const confirm = document.getElementById('confirmationSection');

  // Hide everything except loading
  form.style.display = 'none';
  confirm.style.display = 'none';
  loading.style.display = 'block';

  // Set copy/icons depending on stage
  const title = document.getElementById('loadingTitle');
  const msg1  = document.getElementById('loadingMsg1');
  const msg2  = document.getElementById('loadingMsg2');
  const ill   = loading.querySelector('.loading-ill');

  if (kind === 'finding') {
    title.textContent = 'Looking for a meeting room…';
    // Reuse localized string if available:
    msg1.textContent  = (typeof creating_meeting_loading_inner_html === 'string' && creating_meeting_loading_inner_html.trim())
      ? creating_meeting_loading_inner_html
      : 'Creating the meeting for you…';
    msg2.textContent  = 'Please don’t close this window or the add-in.';
    ill.textContent   = '⏳';
  } else if (kind === 'confirming') {
    title.textContent = 'Confirming your reservation…';
    msg1.textContent  = (typeof status_confirming === 'string' && status_confirming.trim()) ? status_confirming : 'Confirming…';
    msg2.textContent  = 'Please keep this window open.';
    ill.textContent   = '✅';
  } else { // 'cancelling'
    title.textContent = 'Cancelling your reservation…';
    msg1.textContent  = (typeof status_cancelation === 'string' && status_cancelation.trim()) ? status_cancelation : 'Cancelling…';
    msg2.textContent  = 'Please keep this window open.';
    ill.textContent   = '🗑️';
  }
}

function hideLoading() {
  const loading = document.getElementById('loadingSection');
  loading.style.display = 'none';
}

function getSelectedBuildingName() {
  const candidates = ['buildingSelect','building','meetingBuilding','BuildingName','buildingName'];
  for (const id of candidates) {
    const el = document.getElementById(id);
    if (!el) continue;

    if (el.tagName === 'SELECT') {
      const opt = el.selectedOptions && el.selectedOptions[0];
      if (opt) return (opt.textContent || opt.value || '').trim();
    }
    if ('value' in el) return (el.value || '').trim();
  }
  return '';
}



Office.onReady(info => {
  
 if (info.host === Office.HostType.Outlook) {
    const addinContainer = document.getElementById("add-in-container");
    const envStatus = document.getElementById("env-status");
    const devToggle = document.getElementById("devModeToggle");
    const uatEnvToggle = document.getElementById("uatEnvToggle");
    const uatEnvToggleText = document.getElementById("uatEnvToggleText");
    const jsonOutput = document.getElementById("jsonOutput");
    let endpointFlow1Url, endpointFlow2Url;
    let tag = "[ OA Draft ]";
    let cachedBusinessUnit = undefined;
    const appVersionContainer = document.getElementById("appVersion");
    appVersionContainer.innerHTML = "v"+window.appConfig.appVersion + " | Updated: " + window.appConfig.appDateTimeUpdate;


	function applyEnvironment(isUAT) {
		  const container = document.getElementById('add-in-container');
		  const statusEl  = document.getElementById('env-status');
		  const labelEl   = document.getElementById('uatEnvToggleText');

		  // Text
		  statusEl.textContent = isUAT ? 'UAT Environment' : 'DEV Environment';
		  if (labelEl) labelEl.textContent = isUAT ? 'UAT' : 'DEV';

		  // Color theme
		  container.classList.toggle('green-gradient',  isUAT);
		  container.classList.toggle('red-gradient',   !isUAT);
      
      const envSubject = isUAT ? 'UAT' : 'DEV'
      console.log("envSubject :::::::::". envSubject);
      updateSubjectWithTag(envSubject);
		}
  // Add the Smart Meeting Draft to the subject, with UAT prefix by default
  
	const uatToggle = document.getElementById('uatEnvToggle');

	// If you persist env somewhere, restore it here; otherwise default true.
	let isUAT = true;
	// Example if you store it:
	// const saved = localStorage.getItem('oa.env');
	// let isUAT = saved ? saved === 'uat' : true;

	uatToggle.checked = isUAT;
	applyEnvironment(isUAT);

	// Keep it in sync on change
	uatToggle.addEventListener('change', () => {
	  const nowUAT = uatToggle.checked;
	  applyEnvironment(nowUAT);
	  // localStorage.setItem('oa.env', nowUAT ? 'uat' : 'dev'); // optional
	});

  // Get Translation from the JSON 
  function setAddinLanguage(lang) {
    //lang is outlookLanguage, here we get the first two letters from the outlooklanguage string "en-US" will be "en"
    //lang = "fr";
    console.log("LANGUAGE SELECTED:   ",lang)
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
      no_meeting_found = resources["no.room.found.message"];
      show_technical_details = resources["show.technical.details"];
      hide_technical_details = resources["hide.technical.details"];

      business_unit_amenities_error = resources["Business.unit.amenities.error"];
      createing_allocating_room_error = resources["createing.allocating.room.error"];
      warning_24hour_meeting = resources["24hour.meeting.warning"];
      sso_error = resources["sso.error"];

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
        if (key === "no.room.found.message") { no_meeting_found = resources[key]; }
        if (key === "show.technical.details") { show_technical_details = resources[key]; }
        if (key === "hide.technical.details") { hide_technical_details = resources[key]; }

        if (key === "Business.unit.amenities.error") { business_unit_amenities_error = resources[key]; }
        if (key === "createing.allocating.room.error") { createing_allocating_room_error = resources[key]; }
        if (key === "24hour.meeting.warning") { warning_24hour_meeting = resources[key]; }
        if (key === "sso.error") { sso_error = resources[key]; }
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
      } else {
      document.getElementById('link-' + outlooklanguage.substring(0, 2)).classList.add('active-language');
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
          const businessUnit= await Graph.getBusinessUnit(token);
          cachedBusinessUnit = businessUnit;
          const amenitiesList = await Graph.getAmenitiesFromSharePoint(token);
          LoggerApp.debug("Amenities: ", amenitiesList);
          renderAmenitiesCheckboxes(amenitiesList);
          setAddinLanguage(outlooklanguage);
          const calendarId = await Graph.getCalendarId(token, Office.context.mailbox.userProfile.emailAddress); 
        } catch (err) {
          LoggerApp.error("Failed to fetch business unit or amenities:", err);

          const errMsg = business_unit_amenities_error;
          showOverlayMessage(
            "error",
            errMsg,
            err
          );
          return;
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
          const currentSubject = result.value.replace(/(UAT|DEV) \[ OA Draft \]/, "").trim() || "";
          
          if (!currentSubject.includes(tag.trim() + "[ OA Draft ]")) {
            const newSubject = tag + " [ OA Draft ] " + currentSubject;
    
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

            // Convert short weekday names to full names
        function RecurrenceDayOfWeek(day) {
          const mapping = {
            sun: "Sunday",
            sunday: "Sunday",
            mon: "Monday",
            monday: "Monday",
            tue: "Tuesday",
            tuesday: "Tuesday",
            wed: "Wednesday",
            wednesday: "Wednesday",
            thu: "Thursday",
            thursday: "Thursday",
            fri: "Friday",
            friday: "Friday",
            sat: "Saturday",
            saturday: "Saturday"
          };
          const lower = (day || "").toLowerCase();
          return mapping[lower] || day;
        }

function getStaticBusinessUnit(organizerEmail) {
    const emailToBU = {
        "gerald.zanelli.ext@proximus.com": "S&S",
        "tanguy.orban.ext@proximus.com": "S&S",
        "inge.wauters@proximus.com": "S&S",
        "frederic.cornu@proximus.com": "S&S",
        "marleen.demolder@proximus.com": "S&S",
        "leslie.baudoin@proximus.com": "S&S",
        "david.lytton@proximus.com": "S&S",
        "dirk.vandervelden@proximus.com": "DTI",
        "marc.vanduerm@proximus.com": "DTI",
        "kevin.van.schoor@proximus.com": "DTI",
        "margaret.denis@proximus.com": "S&S",
        "wael.al.sabbouh.ext@proximus.com": "DTI",
        "yunting.zhang.ext@proximus.com": "DTI",
        "roberto.campione.ext@proximus.com": "DTI",
        "lars.maubach.ext@proximus.com": "DTI",
        "charlotte.sagaert@proximus.com": "S&S"
    };

    // Normalize email to lowercase and trim spaces
    const normalizedEmail = organizerEmail.trim().toLowerCase();
    return emailToBU[normalizedEmail] || "N/A";
}



        // Map Outlook recurrence types to Microsoft Graph pattern types
        function mapRecurrencePatternType(type) {
          if (!type) return "";
          const mapping = {
            daily: "daily",
            weekly: "weekly",
            monthly: "absoluteMonthly",
            monthlynth: "relativeMonthly",
            yearly: "absoluteYearly",
            yearlynth: "relativeYearly"
          };
          return mapping[type.toLowerCase()] || type;
        }

        // Checks if meeting duration exceeds 24 hours
        function durationExceeds24Hours(start, end) {
          const startTime = new Date(start);
          const endTime = new Date(end);
          return endTime - startTime > 24 * 60 * 60 * 1000;
        }

        function addMonths(date, months) {
          const d = new Date(date);
          d.setMonth(d.getMonth() + months);
          return d;
        }

        function formatDateYYYYMMDD(date) {
          return date.toISOString().split("T")[0];
        }

        function calculateTotalOccurrences(start, end, pattern) {
          const startDate = new Date(start);
          const endDate = new Date(end);
          if (!pattern || !pattern.type) return 1;
          const type = pattern.type;
          const interval = pattern.interval || 1;
          let count = 0;

          if (type === "daily") {
            const msPerDay = 24 * 60 * 60 * 1000;
            count = Math.floor((endDate - startDate) / (msPerDay * interval)) + 1;
          } else if (type === "weekly") {
            const days = pattern.daysOfWeek || [startDate.toLocaleString("en-US", { weekday: "long" })];
            let current = new Date(startDate);
            while (current <= endDate) {
              const dayName = current.toLocaleString("en-US", { weekday: "long" });
              if (days.includes(dayName)) {
                const weeksBetween = Math.floor((current - startDate) / (7 * 24 * 60 * 60 * 1000));
                if (weeksBetween % interval === 0) count++;
              }
              current.setDate(current.getDate() + 1);
            }
          } else if (type === "absoluteMonthly" || type === "relativeMonthly") {
            let current = new Date(startDate);
            while (current <= endDate) {
              count++;
              current.setMonth(current.getMonth() + interval);
            }
          } else {
            count = 1;
          }

          return count;
        }

        // Displays overlay messages with configurable type, main text, and optional technical details
        function showOverlayMessage(messageType, messageText, technicalDetails = "") {
          const overlay = document.getElementById("durationWarningOverlay");
          if (!overlay) {
            LoggerApp.warn("Overlay message element not found");
            return;
          }

          const iconSpan = overlay.querySelector("#durationWarningIcon");
          const messageSpan = overlay.querySelector("#durationWarningMessage");
          const closeButton = overlay.querySelector("#durationWarningClose");
          const modalCard = overlay.querySelector(".modal-card");
          const techSection = overlay.querySelector("#durationTechSection");
          const techToggle = overlay.querySelector("#durationTechToggle");
          const techToggleText = overlay.querySelector("#durationTechToggleText");
          const techMessage = overlay.querySelector("#durationTechnicalMessage");

          const typeMap = {
            info: { icon: "🤷‍♀️", cardClass: "modal-card--info" },
            warning: { icon: "⚠️", cardClass: "modal-card--warning" },
            error: { icon: "❌", cardClass: "modal-card--error" }
          };

          const selectedType = typeMap[messageType] || typeMap.warning;

          if (iconSpan) {
            iconSpan.textContent = selectedType.icon;
          }

          if (messageSpan) {
            messageSpan.textContent = messageText;
          }

          const trimmedTechnicalDetails = (technicalDetails ?? "").toString().trim();
          const hasTechnicalDetails = trimmedTechnicalDetails.length > 0;

          if (techSection) {
            techSection.hidden = !hasTechnicalDetails;
          }

          if (techMessage) {
            techMessage.textContent = trimmedTechnicalDetails;
            techMessage.hidden = true;
          }

          if (techToggle) {
            const showLabel = show_technical_details;
            const hideLabel = hide_technical_details;

            techToggle.setAttribute("aria-expanded", "false");
            techToggle.classList.remove("is-expanded");

            if (techToggleText) {
              techToggleText.textContent = showLabel;
            }

            techToggle.hidden = !hasTechnicalDetails;

            if (!techToggle.dataset.listenerAttached) {
              techToggle.addEventListener("click", () => {
                const isExpanded = techToggle.getAttribute("aria-expanded") === "true";
                const newExpandedState = !isExpanded;

                techToggle.setAttribute("aria-expanded", String(newExpandedState));
                techToggle.classList.toggle("is-expanded", newExpandedState);

                if (techMessage) {
                  techMessage.hidden = !newExpandedState;
                }

                if (techToggleText) {
                  techToggleText.textContent = newExpandedState ? hideLabel : showLabel;
                }
              });
              techToggle.dataset.listenerAttached = "true";
            }
          }

          if (modalCard) {
            modalCard.classList.remove("modal-card--info", "modal-card--warning", "modal-card--error");
            modalCard.classList.add(selectedType.cardClass);
          }

          if (closeButton && !closeButton.dataset.listenerAttached) {
            closeButton.addEventListener("click", () => {
              overlay.classList.remove("is-visible");
              overlay.setAttribute("aria-hidden", "true");
            });
            closeButton.dataset.listenerAttached = "true";
          }

          overlay.classList.add("is-visible");
          overlay.setAttribute("aria-hidden", "false");
        }

        // Remove any previous event listener and attach new one to the main button
        const btn = document.getElementById("generateJson");
        btn.replaceWith(btn.cloneNode(true));
        const newBtn = document.getElementById("generateJson");

        newBtn.addEventListener("click", function (e) {
          e.preventDefault();

          collectAndShowData(function(payload) {
        if (devToggle.checked) {
          // OLD:
          // jsonOutput.textContent = "⏳"+ creating_meeting_loading_inner_html;

          // NEW:
          showLoading('finding');

          if(uatEnvToggle.checked){
            endpointFlow1Url = window.appConfig.endpointFlow1Url;
            endpointFlow2Url = window.appConfig.endpointFlow2Url;
          } else {
            endpointFlow1Url = window.appConfig.DEVendpointFlow1Url;
            endpointFlow2Url = window.appConfig.DEVendpointFlow2Url;
          }

          fetch(endpointFlow1Url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
          })
          .then(response => {
            if (!response.ok) throw new Error("Flow1 error: HTTP status " + response.status);
            return response.text();
          })
          .then(txt => {
            LoggerApp.info("Flow1 RAW response:", txt);
            let result;
            try { result = JSON.parse(txt); }
            catch { throw new Error("Flow1 did not return valid JSON: " + txt); }

            if (!result["Placeholder Meeting Room Name"]) {
              
              

              throw new Error("Missing 'Placeholder Meeting Room Name' in Flow1 response: " + txt);
            }

            hideLoading();
            const buildingDisplay =
              payload.BuildingName || payload.Building || payload.buildingName || payload.building || getSelectedBuildingName();

            showConfirmationUI(result["Placeholder Meeting Room Name"], result, buildingDisplay);
          })
          .catch(err => {
            hideLoading();
            // Return to form so user can retry
            document.getElementById('meetingForm').style.display = 'flex';
            // jsonOutput.textContent = "Error sending to Flow: " + err;
            
            const errorStr = String(err); // Convert to string
            let errorMsg;
            if (errorStr.indexOf("No room booked") !== -1) {
              errorMsg = no_meeting_found;
            } else {
              errorMsg = "Error While Creating A Meeting or Allocating Meeting Room!";
            }

            showOverlayMessage(
              "error",
              errorMsg,
              errorStr
              );
          });

        } else {
          // showOverlayMessage(
          //   "error",
          //   createing_allocating_room_error,
          //   JSON.stringify(payload, null, 2)
          //   );
           jsonOutput.textContent = JSON.stringify(payload, null, 2);
        }
      });

    });

  // Additional new functions:
 
      function renderAmenitiesCheckboxes(amenitiesList) {
        const container = document.getElementById('amenitiesContainer');
        container.innerHTML = '';
        container.classList.remove('row'); // just in case
        container.style.display = 'block';

        // Wrapper
        const wrapper = document.createElement('div');
        wrapper.className = 'amenities-wrapper';

        // Toggle (header)
        const toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = 'amenities-toggle';
        toggle.setAttribute('aria-expanded', 'false');

        const left = document.createElement('div');
        left.style.display = 'flex';
        left.style.flexDirection = 'column';
        left.style.gap = '2px';

        const title = document.createElement('div');
        title.className = 'amenities-title';
        title.setAttribute('data-i18n', 'room.amenities');
        title.textContent = 'Room Amenities'; 

        const summary = document.createElement('div');
        summary.className = 'amenities-summary';
        summary.textContent = '—';

        left.appendChild(title);
        left.appendChild(summary);

        const right = document.createElement('div');
        right.className = 'amenities-right';

        const badge = document.createElement('span');
        badge.className = 'amenities-badge';
        badge.textContent = '0';

        const chev = document.createElement('i');
        chev.className = 'amenities-chevron';

        right.appendChild(badge);
        right.appendChild(chev);

        toggle.appendChild(left);
        toggle.appendChild(right);

        // Panel
        const panel = document.createElement('div');
        panel.className = 'amenities-panel';
        panel.style.display = 'none';

        const list = document.createElement('div');
        list.className = 'amenities-list';

        amenitiesList.forEach((amenity) => {
          const item = document.createElement('label');
          item.className = 'amenities-item ms-Checkbox';

          const cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.name = 'amenities';           // keep contract
          cb.value = String(amenity.id);   // (optional) normalize to string
          cb.dataset.label = amenity.name; // <-- add this line
          cb.checked = false;

          const txt = document.createElement('span');
          txt.className = 'ms-Checkbox-label';
          txt.textContent = amenity.name;

          item.appendChild(cb);
          item.appendChild(txt);
          list.appendChild(item);
        });

        // Additional Seats row (stays in panel, spans full width)
        const extra = document.createElement('div');
        extra.className = 'amenities-extra';

        const seatsLabel = document.createElement('label');
        seatsLabel.className = 'ms-Label';
        seatsLabel.setAttribute('for', 'additionalSeats');
        seatsLabel.setAttribute('data-i18n', 'additional.seats');
        seatsLabel.textContent = (typeof additional_seats_inner_html === 'string' && additional_seats_inner_html.trim())
        ? additional_seats_inner_html : 'Additional Seats';

        const seatsInput = document.createElement('input');
        seatsInput.type = 'number';
        seatsInput.id = 'additionalSeats';
        seatsInput.min = '0';
        seatsInput.max = '20';
        seatsInput.value = '0';

        extra.appendChild(seatsLabel);
        extra.appendChild(seatsInput);

        panel.appendChild(list);
        panel.appendChild(extra);

        wrapper.appendChild(toggle);
        wrapper.appendChild(panel);
        container.appendChild(wrapper);

        // ---- behavior
        const updateSummaryAndBadge = () => {
          const selected = Array
          .from(container.querySelectorAll('input[name="amenities"]:checked'))
          .map(cb =>
            cb.dataset.label ||
            cb.closest('label')?.querySelector('.ms-Checkbox-label')?.textContent?.trim() ||
            ''
          )
          .filter(Boolean);

          badge.textContent = String(selected.length);
          if (selected.length === 0) {
          summary.textContent = 'Select…';
          wrapper.classList.remove('has-value');
          } else if (selected.length <= 2) {
          summary.textContent = selected.join(', ');
          wrapper.classList.add('has-value');
          } else {
          summary.textContent = `${selected.slice(0, 2).join(', ')}, …`;
          wrapper.classList.add('has-value');
          }
          summary.title = selected.join(', ');
        };
        const open = () => {
        panel.style.display = 'block';
        wrapper.classList.add('amenities-open');
        toggle.setAttribute('aria-expanded', 'true');
        // optional: focus first checkbox
        const first = panel.querySelector('input[name="amenities"]');
        if (first) first.focus({ preventScroll: true });
        };
        const close = () => {
        panel.style.display = 'none';
        wrapper.classList.remove('amenities-open');
        toggle.setAttribute('aria-expanded', 'false');
        };

        toggle.addEventListener('click', () => {
        const isOpen = panel.style.display !== 'none';
        (isOpen ? close : open)();
        });

        // Close on outside click / Esc
        document.addEventListener('click', (e) => {
        if (!wrapper.contains(e.target)) close();
        });
        document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') close();
        });

        // Update summary live
        list.addEventListener('change', (e) => {
        if (e.target && e.target.name === 'amenities') {
          updateSummaryAndBadge();
        }
        });

        // Initial text
        updateSummaryAndBadge();
      }
        
      // Displays confirmation (pre-action) UI
        function showConfirmationUI(roomName, flowResponse, buildingDisplayName) {
          const form = document.getElementById('meetingForm');
          const section = document.getElementById('confirmationSection');

          // Show section, hide form
          form.style.display = 'none';
          section.style.display = 'block';

          // Reset layout
          section.classList.remove('show-result', 'is-cancel', 'is-error');
          section.classList.add('is-success');

          document.getElementById('confirmationTitle').textContent = 'Temporary placeholder reserved!';
          document.getElementById('confirmationStatus').textContent = '';

          const msg = document.getElementById('confirmationMessage');
          const roomBlock = section.querySelector('.room-block');
          const ctas = document.getElementById('confirmCancelBtns');
          msg.style.display = '';
          if (roomBlock) roomBlock.style.display = '';
          ctas.style.display = 'flex';
          document.getElementById('backBtnDiv').style.display = 'none';

          // Fill room name (main sentence + block)
          document.getElementById('placeholderRoomName').textContent = roomName || '';
          const roomNameBlock = document.getElementById('placeholderRoomName_block');
          if (roomNameBlock) roomNameBlock.textContent = roomName || '';

          // IDs from flow
          document.getElementById("confirmEventID").value     = flowResponse.EventID     || "";
          document.getElementById("confirmDataverseID").value = flowResponse.DataverseID || "";
          document.getElementById("confirmIcalUid").value     = flowResponse.iCalUid     || "";
          document.getElementById("organizerResponse").value  = "";

          // Building
          (function setBuilding() {
            const buildingEl = document.getElementById('placeholderBuilding');
            if (!buildingEl) return;

            const block = buildingEl.closest('.room-block');
            const b = (buildingDisplayName || '').trim();

            if (b) {
              buildingEl.textContent = b;
              if (block) block.style.display = 'flex';
            } else {
              if (block) block.style.display = 'none';
            }
          })();

          // Buttons
          const confirmBtn = document.getElementById('confirmBtn');
          const cancelBtn  = document.getElementById('cancelBtn');
          confirmBtn.disabled = false;
          cancelBtn.disabled  = false;

          confirmBtn.onclick = function () {
            document.getElementById("organizerResponse").value = "Confirmed";
            handleConfirmation();
          };
          cancelBtn.onclick = function () {
            document.getElementById("organizerResponse").value = "Cancel";
            handleConfirmation();
          };
        }

	
    // Handles confirmation or cancellation of meeting room
      function handleConfirmation() {
        const isConfirming = document.getElementById("organizerResponse").value === "Confirmed";

        // Show loading screen for this stage
        showLoading(isConfirming ? 'confirming' : 'cancelling');

        const item = Office.context.mailbox.item;
        let payload;

        getAttachmentsWithContent(item)
          .then(attachments => {
            payload = {
              EventID:      document.getElementById("confirmEventID").value,
              DataverseID:  document.getElementById("confirmDataverseID").value,
              iCalUid:      document.getElementById("confirmIcalUid").value || "",
              OrganizerResponse: document.getElementById("organizerResponse").value,
              Attachments:  attachments
            };
            LoggerApp.info("EventID:" + payload.EventID);

            return fetch(endpointFlow2Url, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(payload)
            });
          })
          .then(r => r.text())
          .then(() => {
            hideLoading();

            const section  = document.getElementById('confirmationSection');
            const statusEl = document.getElementById('confirmationStatus');
            const ctas     = document.getElementById('confirmCancelBtns');

            // Bring back confirmation card and show final state
            section.style.display = 'block';
            const isConfirm = payload.OrganizerResponse === 'Confirmed';

            section.classList.toggle('is-success', isConfirm);
            section.classList.toggle('is-cancel', !isConfirm);
            document.getElementById('confirmationTitle').textContent =
              isConfirm ? 'Reservation confirmed' : 'Reservation cancelled';

            const okMsg     = 'Reservation confirmed! Meeting room is now visible to all attendees.';
            const cancelMsg = 'Reservation cancelled. The placeholder has been removed.';
            statusEl.textContent = isConfirm ? okMsg : cancelMsg;

            document.getElementById('confirmationMessage').style.display = 'none';
            document.getElementById('cancelBtn').style.display = 'none';
            document.getElementById('confirmBtn').style.display = 'none';
            const roomBlock = section.querySelector('.room-block');
            if (roomBlock) roomBlock.style.display = 'none';
            ctas.style.display = 'none';
            section.classList.add('show-result');

            // Show the follow-up actions
            document.getElementById('backBtnDiv').style.display = 'flex';
            document.getElementById('backBtn').onclick = resetToForm;
          })
          .catch(err => {
            hideLoading();

            // Fall back to confirmation section to display the error
            const section  = document.getElementById('confirmationSection');
            const statusEl = document.getElementById('confirmationStatus');

            section.style.display = 'block';
            section.classList.add('is-error');
            document.getElementById('confirmationTitle').textContent = 'Something went wrong';
            statusEl.textContent = "Error: " + err;

            document.getElementById('backBtnDiv').style.display = 'flex';
            document.getElementById('backBtn').onclick = resetToForm;
          });
      }

    // Resets UI back to the initial form
    function resetToForm() {
      document.getElementById('backBtnDiv').style.display = 'flex';
      document.getElementById('confirmCancelBtns').style.display = 'flex';
	    document.getElementById('cancelBtn').style.display = 'flex';
	    document.getElementById('confirmBtn').style.display = 'flex';

      document.getElementById('meetingForm').style.display = 'flex';
      document.getElementById('confirmationSection').style.display = 'none';
      document.getElementById('jsonOutput').textContent = "";
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
        LoggerApp.info("Business Unit:", getStaticBusinessUnit(organizerEmail));
        Office.context.mailbox.item.requiredAttendees.getAsync(function(requiredResult) {
          const requiredAttendeesArr = (requiredResult.value || []).map(a => a.emailAddress);

          Office.context.mailbox.item.optionalAttendees.getAsync(function(optionalResult) {
            const optionalAttendeesArr = (optionalResult.value || []).map(a => a.emailAddress);

            Office.context.mailbox.item.start.getAsync(function(startResult) {
              const startDateTime = startResult.value;
              Office.context.mailbox.item.end.getAsync(function(endResult) {
                const endDateTime = endResult.value;

                //Calculate if duration of the meeting exceeds 24 hours
                if (durationExceeds24Hours(startDateTime, endDateTime)) {
                  showOverlayMessage(
                    "warning",
                    warning_24hour_meeting
                  );
                  return;
                }

                Office.context.mailbox.item.recurrence.getAsync(function(recurrenceResult) {
                  const recurrence = recurrenceResult.status === Office.AsyncResultStatus.Succeeded
                    ? recurrenceResult.value : null;

                  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(bodyResult) {
                    const body = (bodyResult.value || "");
                    //const businessUnit = cachedBusinessUnit;
                    const businessUnit = getStaticBusinessUnit(organizerEmail);
                    const onBehalfOf = document.getElementById("onBehalfOf")?.checked || false;
                    const additionalSeats = validateInputNum(document.getElementById("additionalSeats"), 0, 20);
                    const extraTimeBefore = validateInputNum(document.getElementById("extraTimeBefore"), 0, 15);
                    const extraTimeAfter = validateInputNum(document.getElementById("extraTimeAfter"), 0, 15);
                    const building = document.getElementById("building").value;
                    const amenities = Array.from(document.querySelectorAll("input[name='amenities']:checked")).map(el => el.value).join(",");
                    const meetingTypeEl = document.getElementById('meetingType');
				            const meetingType = meetingTypeEl ? meetingTypeEl.value : '';

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
                    let RecurrencePattern = null;

                    if (recurrence?.recurrencePattern) {
                      Recurrence = recurrence.recurrencePattern.type || "";

                      RecurrenceDaysOfWeek = (recurrence.recurrencePattern.daysOfWeek || [])
                      .map(RecurrenceDayOfWeek)
                      .join(",");

                      RecurrenceEndDate = recurrence.recurrenceRange?.endDate || "";

                      RecurrencePattern = {
                        type: mapRecurrencePatternType(recurrence.recurrencePattern.type || ""),
                        interval: recurrence.recurrencePattern.interval || 1
                      };

                      if (Array.isArray(recurrence.recurrencePattern.daysOfWeek) && recurrence.recurrencePattern.daysOfWeek.length) {
                        RecurrencePattern.daysOfWeek = recurrence.recurrencePattern.daysOfWeek.map(RecurrenceDayOfWeek);
                      }
                      if (recurrence.recurrencePattern.dayOfMonth != null) {
                        RecurrencePattern.dayOfMonth = recurrence.recurrencePattern.dayOfMonth;
                      }
                      if (recurrence.recurrencePattern.firstDayOfWeek) {
                        RecurrencePattern.firstDayOfWeek = RecurrenceDayOfWeek(recurrence.recurrencePattern.firstDayOfWeek);
                      }
                      if (recurrence.recurrencePattern.index) {
                        RecurrencePattern.index = recurrence.recurrencePattern.index.toLowerCase();
                      }

                    } else if (recurrence?.recurrenceProperties) {
                      Recurrence = recurrence.recurrenceType || "";
                      RecurrenceDaysOfWeek = Array.isArray(recurrence.recurrenceProperties.days)
                      ? recurrence.recurrenceProperties.days.map(RecurrenceDayOfWeek).join(",") : "";

                      if (recurrence.seriesTime && recurrence.seriesTime.endYear) {
                        RecurrenceEndDate =
                          recurrence.seriesTime.endYear + "-" +
                          String(recurrence.seriesTime.endMonth).padStart(2, '0') + "-" +
                          String(recurrence.seriesTime.endDay).padStart(2, '0');
                      }

                      RecurrencePattern = {
                        type: mapRecurrencePatternType(recurrence.recurrenceType || ""),
                        interval: recurrence.recurrenceProperties.interval || 1
                      };
                      if (Array.isArray(recurrence.recurrenceProperties.days) && recurrence.recurrenceProperties.days.length) {
                        RecurrencePattern.daysOfWeek = recurrence.recurrenceProperties.days.map(RecurrenceDayOfWeek);
                      }
                      if (recurrence.recurrenceProperties.dayOfMonth != null) {
                        RecurrencePattern.dayOfMonth = recurrence.recurrenceProperties.dayOfMonth;
                      }
                      if (recurrence.recurrenceProperties.firstDayOfWeek) {
                        RecurrencePattern.firstDayOfWeek = RecurrenceDayOfWeek(recurrence.recurrenceProperties.firstDayOfWeek);
                      }
                      if (recurrence.recurrenceProperties.weekIndex) {
                        RecurrencePattern.index = recurrence.recurrenceProperties.weekIndex.toLowerCase();
                      }
                    }

                    let TotalRecurringMeetings = 1;
                    if (Recurrence !== "none") {
                      const startDateObj = new Date(startDateTime);
                      const threeMonthsLater = addMonths(startDateObj, 3);
                      const existingEnd = RecurrenceEndDate ? new Date(RecurrenceEndDate) : null;
                      const finalEnd = (!existingEnd || existingEnd > threeMonthsLater) ? threeMonthsLater : existingEnd;
                      if (!existingEnd || existingEnd > threeMonthsLater) {
                        RecurrenceEndDate = formatDateYYYYMMDD(finalEnd);
                      }
                      TotalRecurringMeetings = calculateTotalOccurrences(startDateObj, finalEnd, RecurrencePattern);
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
                          BusinessUnit: getStaticBusinessUnit(organizerEmail),
                          MeetingType: meetingType,
                          Attendees: {
                            RequiredAttendees: requiredAttendeesArr,
                            OptionalAttendees: optionalAttendeesArr
                          },
                          isPrivate: isPrivate,
                          OnBehalfOf: onBehalfOf,
                          RecurrencePattern,
                          Recurrence,
                          RecurrenceEndDate,
                          RecurrenceDaysOfWeek,
                          TotalRecurringMeetings,
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
                          BusinessUnit: getStaticBusinessUnit(organizerEmail),
                          MeetingType: meetingType,
                          Attendees: {
                            RequiredAttendees: requiredAttendeesArr,
                            OptionalAttendees: optionalAttendeesArr
                          },
                          isPrivate: isPrivate,
                          OnBehalfOf: onBehalfOf,
                          RecurrencePattern,
                          Recurrence,
                          RecurrenceEndDate,
                          RecurrenceDaysOfWeek,
                          TotalRecurringMeetings,
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
