Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    const devToggle = document.getElementById("devModeToggle");
    const jsonOutput = document.getElementById("jsonOutput");
    const endpointFlow1Url = window.appConfig.endpointFlow1Url;
    const endpointFlow2Url = window.appConfig.endpointFlow2Url;

    // Utility: Validates numeric input
    function validateInputNum(input, min, max) {
      let v = parseInt(input.value, 10) || 0;
      v = Math.max(min, Math.min(max, v));
      input.value = v;
      return v;
    }

    // Main button: Add your meeting room
	//Remove precedent listner
	const btn = document.getElementById("generateJson");
	btn.replaceWith(btn.cloneNode(true));
	const newBtn = document.getElementById("generateJson");

	newBtn.addEventListener("click", function (e) {
      e.preventDefault();
      collectAndShowData(function(payload) {
        if (devToggle.checked) {
          jsonOutput.textContent = "Sending...";
			fetch(endpointFlow1Url, {
			  method: "POST",
			  headers: { "Content-Type": "application/json" },
			  body: JSON.stringify(payload)
			})
			.then(response => {
			  // 1. Se status HTTP non è OK, segnala errore
			  if (!response.ok) throw new Error("Flow1 error: HTTP status " + response.status);

			  // 2. Leggi la risposta come testo per fare debug/log
			  return response.text().then(txt => {
				console.log("Flow1 RAW response:", txt); // <-- Puoi vedere in console devtools

				// 3. Prova a fare il parsing come JSON
				let result;
				try { result = JSON.parse(txt); }
				catch { throw new Error("Flow1 did not return valid JSON: " + txt); }

				// 4. Verifica che il dato atteso sia presente
				if (!result["Placeholder Meeting Room Name"]) {
				  throw new Error("Missing 'Placeholder Meeting Room Name' in Flow1 response: " + txt);
				}

				// 5. Successo: mostra la conferma
				showConfirmationUI(result["Placeholder Meeting Room Name"], result);
			  });
			})
			.catch(err => {
			  jsonOutput.textContent = "Error sending to Flow: " + err;
			  // **Non resettare subito la form**: lascia così, l'utente può correggere/riprovarci.
			});
        } else {
          jsonOutput.textContent = JSON.stringify(payload, null, 2);
        }
      });
    });

    // Show confirmation/cancel UI
		function showConfirmationUI(roomName, flowResponse) {
		  document.getElementById('meetingForm').style.display = 'none';
		  document.getElementById('confirmationSection').style.display = 'block';
		  document.getElementById('confirmationStatus').textContent = "";
		  document.getElementById('backBtnDiv').style.display = 'none';
		  document.getElementById('placeholderRoomName').textContent = roomName;

		  // Popola hidden fields dal flowResponse
		  document.getElementById("confirmEventID").value = flowResponse.EventID || "";
		  document.getElementById("confirmDataverseID").value = flowResponse.DataverseID || "";
		  document.getElementById("confirmIcalUid").value = flowResponse.iCalUid || "";
		  document.getElementById("organizerResponse").value = "";

		  // Riabilita bottoni (in caso di retry/back)
		  document.getElementById('confirmBtn').disabled = false;
		  document.getElementById('cancelBtn').disabled = false;
		  document.getElementById('confirmCancelBtns').style.display = "block";

		  document.getElementById('confirmBtn').onclick = function() {
			document.getElementById("organizerResponse").value = "Confirmed";
			handleConfirmation();
		  };
		  document.getElementById('cancelBtn').onclick = function() {
			document.getElementById("organizerResponse").value = "Cancel";
			handleConfirmation();
		  };
		}

		// Confirm-Cancel handler
		function handleConfirmation() {
		  const statusDiv = document.getElementById("confirmationStatus");
		  // Disabilita (o nascondi) i bottoni
		  document.getElementById('confirmBtn').disabled = true;
		  document.getElementById('cancelBtn').disabled = true;
		  document.getElementById('confirmCancelBtns').style.display = "none";
		  statusDiv.textContent = (document.getElementById("organizerResponse").value === "Confirmed" ? "Confirming..." : "Cancelling...");

		  const payload = {
			EventID: document.getElementById("confirmEventID").value,
			DataverseID: document.getElementById("confirmDataverseID").value,
			iCalUid: document.getElementById("confirmIcalUid").value || "",
			OrganizerResponse: document.getElementById("organizerResponse").value
		  };

		  fetch(endpointFlow2Url, {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify(payload)
		  })
		  .then(response => response.text())
		  .then(result => {
			statusDiv.textContent =
			  payload.OrganizerResponse === "Confirmed"
				? "Reservation confirmed! Meeting room is now visible to all attendees."
				: "Reservation cancelled. The placeholder has been removed.";
			// Mostra Back solo ora
			document.getElementById('backBtnDiv').style.display = 'block';
			document.getElementById('backBtn').onclick = resetToForm;
		  })
		  .catch(err => {
			statusDiv.textContent = "Error: " + err;
			document.getElementById('backBtnDiv').style.display = 'block';
			document.getElementById('backBtn').onclick = resetToForm;
		  });
		}

    // Allows retry/reset to the form
    function resetToForm() {
	  document.getElementById('backBtnDiv').style.display = 'flex';
      document.getElementById('meetingForm').style.display = 'block';
      document.getElementById('confirmationSection').style.display = 'none';
      jsonOutput.textContent = "";
    }

    // Mapping all fields into payload
    function collectAndShowData(callback) {
      Office.context.mailbox.item.subject.getAsync(function(subjectResult) {
        const subject = subjectResult.value;
        const calendarId = Office.context.mailbox.item.calendarId || "";
        const organizerEmail = Office.context.mailbox.userProfile.emailAddress;
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
					const body = (bodyResult.value || "") + "<!--OA_MEETING_ROOM_ACTIVATED-->";
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
                    // Recurrence mapping compatibile
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
                    const payload = {
                      Title: subject,
                      Body: body,
                      Startdatetime: startDateTime,
                      Enddatetime: endDateTime,
                      OrganizerEmailAddress: "roberto.campione.ext@proximusuat.com",
                      OrganizerCalendarID: "AAMkAGExZGQxMDAxLTg1ZjktNGMwNy1iMzg4LTY0OTAxYWFjYmUzMQBGAAAAAAAEAAQ7FEmVSotJoNetmaQ4BwBUjqXA3xWGTapiAE_7zGePAAAAAAEGAAAifHiMQ6ciS5lzQ5O2Zt71AAAqgIKzAAA=",
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
    }
  }
});
