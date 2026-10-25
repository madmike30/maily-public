(function () {
      const errors = [];
      window.__harness = { errors };
      window.addEventListener("error", (event) => errors.push(`error:${event.message}`));
      window.addEventListener("unhandledrejection", (event) => {
        const reason = event.reason && event.reason.message ? event.reason.message : String(event.reason);
        errors.push(`rejection:${reason}`);
      });
      window.__harness.lastBriefingRequestLanguage = "";

      window.Office = {
        AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
        onReady(callback) { setTimeout(callback, 0); },
        context: {
          diagnostics: { hostName: "Outlook", hostVersion: "mock" },
          mailbox: {
            userProfile: { displayName: "Alex Example", emailAddress: "alex@example.com" }
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
          return new Response(JSON.stringify({
            profileId: "default",
            values: { language: "English" }
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/daily-briefing")) {
          const payload = options && options.body ? JSON.parse(options.body) : {};
          window.__harness.lastBriefingRequestLanguage = payload.outputLanguageOverride || "";
          return new Response(JSON.stringify({
            sections: {
              rightNow: "You have 2 meetings, 3 open action items, and 2 unread emails worth checking.",
              meetingsToday: "- 09:00 Budget sync\n- 13:00 Customer follow-up",
              openActionItems: "- Reply to the pricing thread\n- Send the revised schedule",
              importantUnread: "- Customer escalation\n- Security review update",
              nextFocus: "Start with the pricing reply before the 09:00 sync."
            },
            meta: {
              requestId: "briefing-mock",
              model: "gpt-5.4",
              meetingCount: 2,
              openActionItemCount: 3,
              importantUnreadCount: 2
            }
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
