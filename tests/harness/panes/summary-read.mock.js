(function () {
      const errors = [];
      window.__harness = { errors };
      window.addEventListener("error", (event) => errors.push(`error:${event.message}`));
      window.addEventListener("unhandledrejection", (event) => {
        const reason = event.reason && event.reason.message ? event.reason.message : String(event.reason);
        errors.push(`rejection:${reason}`);
      });
      window.__harness.lastSummaryRequestLanguage = "";

      function asyncSuccess(value, callback) {
        setTimeout(() => callback({ status: "succeeded", value }), 0);
      }

      window.Office = {
        AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
        CoercionType: { Html: "html", Text: "text" },
        onReady(callback) { setTimeout(callback, 0); },
        context: {
          diagnostics: { hostName: "Outlook", hostVersion: "mock" },
          mailbox: {
            userProfile: { displayName: "Alex Example", emailAddress: "alex@example.com" },
            item: {
              subject: "Welcome to the synthetic training workshop",
              to: [
                { displayName: "Alex Example", emailAddress: "alex@example.com" }
              ],
              cc: [],
              from: {
                displayName: "Example Training Team",
                emailAddress: "updates@example.com"
              },
              body: {
                getAsync(typeOrCallback, maybeCallback) {
                  const callback = typeof typeOrCallback === "function" ? typeOrCallback : maybeCallback;
                  asyncSuccess(
                    "Welcome to the synthetic training workshop. Please review the training resources and start testing when it becomes available.",
                    callback);
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
        if (textUrl.includes("/api/outlook/summary")) {
          const payload = options && options.body ? JSON.parse(options.body) : {};
          window.__harness.lastSummaryRequestLanguage = payload.outputLanguageOverride || "";
          const isHebrew = window.__harness.lastSummaryRequestLanguage === "Hebrew";
          return new Response(JSON.stringify({
            sections: {
              executiveSummary: isHebrew
                ? "צוות Example Training Team הכריז על סדנת ההדרכה לדוגמה."
                : "Example Training Team announced the synthetic training workshop.",
              mainPoints: isHebrew
                ? ["ההודעה מיועדת לעובדי Example Company.", "יש להתחיל הדרכה ובדיקות כשהשירות יהיה זמין."]
                : ["The note is for Example Company employees.", "Training and testing should begin once available."],
              actionItems: isHebrew
                ? ["לעבור על חומרי ההדרכה.", "להתחיל בדיקות במכשירים הזמינים."]
                : ["Review the training resources.", "Start testing on available devices."]
            },
            meta: { requestId: "summary-read-mock", model: "gpt-5.4", surface: "read" }
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
    })();
