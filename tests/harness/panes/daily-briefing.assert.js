setTimeout(() => {
      const result = {
        errors: window.__harness.errors,
        status: document.getElementById("briefingStatus").textContent || "",
        rightNow: document.getElementById("briefingRightNow").textContent || "",
        nextFocus: document.getElementById("briefingNextFocus").textContent || "",
        selectedLanguage: document.querySelector(".language-pill.is-selected") ? document.querySelector(".language-pill.is-selected").textContent || "" : "",
        lastBriefingRequestLanguage: window.__harness.lastBriefingRequestLanguage || ""
      };
      const pass = result.errors.length === 0 &&
        result.status.indexOf("Daily briefing ready.") >= 0 &&
        result.rightNow.indexOf("2 meetings") >= 0 &&
        result.nextFocus.indexOf("pricing reply") >= 0 &&
        result.selectedLanguage === "English" &&
        result.lastBriefingRequestLanguage === "";
      document.getElementById("result").textContent = JSON.stringify(result, null, 2);
      document.title = pass ? "PASS" : "FAIL";
    }, 2000);
