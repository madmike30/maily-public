(function () {
      const url = new URL(window.location.href);
      const scenarioName = String(url.searchParams.get("scenario") || "english").trim().toLowerCase();
      const errors = [];
      window.__harness = { errors };
      window.__harness.scenarioName = scenarioName;
      window.__harness.summaryRequestLanguages = [];
      window.addEventListener("error", (event) => errors.push(`error:${event.message}`));
      window.addEventListener("unhandledrejection", (event) => {
        const reason = event.reason && event.reason.message ? event.reason.message : String(event.reason);
        errors.push(`rejection:${reason}`);
      });
      window.__harness.lastSummaryRequestLanguage = "";

      function runScenarioFlow() {
        if (scenarioName !== "hebrew") {
          return;
        }

        const trySwitchLanguage = () => {
          const hebrewButton = document.getElementById("summaryLanguageHebrewBtn");
          const summarizeButton = document.getElementById("summarizeBtn");
          if (!hebrewButton || !summarizeButton || summarizeButton.disabled) {
            setTimeout(trySwitchLanguage, 200);
            return;
          }

          hebrewButton.click();
        };

        setTimeout(trySwitchLanguage, 900);
      }

      function asyncSuccess(value, callback) {
        setTimeout(() => callback({ status: "succeeded", value }), 0);
      }

      const bodyHtml = "<div>We can move the target to next week.</div><div id='divRplyFwdMsg'><div>From: Team</div><div>Subject: Plan</div></div>";

      window.Office = {
        AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
        CoercionType: { Html: "html", Text: "text" },
        onReady(callback) { setTimeout(callback, 0); },
        context: {
          diagnostics: { hostName: "Outlook", hostVersion: "mock" },
          mailbox: {
            userProfile: { displayName: "Alex Example", emailAddress: "alex@example.com" },
            item: {
              subject: { getAsync(callback) { asyncSuccess("RE: Plan", callback); } },
              to: { getAsync(callback) { asyncSuccess([{ emailAddress: "team@example.com" }], callback); } },
              cc: { getAsync(callback) { asyncSuccess([], callback); } },
              from: { getAsync(callback) { asyncSuccess({ displayName: "Alex Example", emailAddress: "alex@example.com" }, callback); } },
              getComposeTypeAsync(callback) { asyncSuccess({ composeType: "reply" }, callback); },
              body: {
                getTypeAsync(callback) { asyncSuccess("html", callback); },
                getAsync(type, callback) {
                  const text = "We can move the target to next week.\n\nFrom: Team\nSubject: Plan";
                  asyncSuccess(type === "text" ? text : bodyHtml, callback);
                }
              }
            }
          }
        }
      };

      window.fetch = async function (url, options) {
        const textUrl = String(url);
        if (textUrl.includes("/api/health")) {
          return new Response(JSON.stringify({ status: "ok" }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/readiness")) {
          return new Response(JSON.stringify({ status: "ready", ready: true, codexReady: true, userMessage: "Maily is ready.", recommendedAction: "", checks: { codex: { ok: true, details: {} } } }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/settings")) {
          return new Response(JSON.stringify({ profileId: "default" }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/generate-reply")) {
          window.__harness.replyRequest = JSON.parse(options.body);
          return new Response(JSON.stringify({ draft: { text: "Synthetic reply preview." } }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/summary")) {
          const payload = options && options.body ? JSON.parse(options.body) : {};
          window.__harness.lastSummaryRequestLanguage = payload.outputLanguageOverride || "";
          window.__harness.summaryRequestLanguages.push(window.__harness.lastSummaryRequestLanguage);
          const isHebrew = window.__harness.lastSummaryRequestLanguage === "Hebrew";
          return new Response(JSON.stringify({
            sections: {
              executiveSummary: isHebrew ? "התוכנית זזה לשבוע הבא." : "The plan is moving to next week.",
              mainPoints: isHebrew
                ? ["הצוות הסכים לדחות את התזמון.", "תאריך ההשקה נדחה."]
                : ["The team agreed to shift timing.", "The launch date is being pushed."],
              actionItems: isHebrew
                ? ["לאשר את לוח הזמנים המעודכן.", "לעדכן את הצוות בתאריך החדש."]
                : ["Confirm the updated schedule.", "Tell the team about the new date."],
              recommendedNextSteps: isHebrew ? "להשיב עם התאריך המעודכן." : "Reply with the revised date."
            },
            meta: { requestId: "summary-mock", model: "gpt-5.4", surface: "compose" }
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/client-log")) {
          return new Response("", { status: 202 });
        }
        return new Response("", { status: 200 });
      };

      window.Response = window.Response || class Response {
        constructor(body, init) { this._body = body; this.status = init.status; this.ok = this.status >= 200 && this.status < 300; this.headers = init.headers || {}; }
        async json() { return JSON.parse(this._body || "{}"); }
        async text() { return this._body; }
      };

      runScenarioFlow();
    })();
