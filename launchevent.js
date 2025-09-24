// Classic fallback script for Windows clients
(function () {
  // Register when Office.js is available
  function registerHandler() {
    try {
      if (window.Office && Office.actions && typeof Office.actions.associate === "function") {
        Office.actions.associate("onAppointmentSendHandler", function (event) {
          try {
            event.completed({ allowEvent: true });
          } catch (e) {
            try { event.completed({ allowEvent: true }); } catch (_) {}
          }
        });
      } else {
        // Some classic hosts expect a global function with the manifest FunctionName
        window.onAppointmentSendHandler = function (event) {
          try { event.completed({ allowEvent: true }); } catch (e) {}
        };
      }
    } catch (err) {
      console.warn("OA launchevent stub error", err);
      // If we cannot register, do nothing — manifest still resolves the resource.
    }
  }
 
  if (window.Office && Office.onReady) {
    Office.onReady(registerHandler);
  } else {
    // wait a short time for office.js to load
    var t = setInterval(function () {
      if (window.Office && Office.onReady) {
        clearInterval(t);
        Office.onReady(registerHandler);
      }
    }, 200);
    // fail-safe timeout
    setTimeout(function () { clearInterval(t); }, 5000);
  }
})();
