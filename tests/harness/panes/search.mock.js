(function () {
      const errors = [];
      window.__harness = { errors };
      window.addEventListener("error", (event) => errors.push(`error:${event.message}`));
      window.addEventListener("unhandledrejection", (event) => {
        const reason = event.reason && event.reason.message ? event.reason.message : String(event.reason);
        errors.push(`rejection:${reason}`);
      });

      window.Office = {
        onReady(callback) { setTimeout(callback, 0); },
        context: {
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
          return new Response(JSON.stringify({ profileId: "default" }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/search/start")) {
          return new Response(JSON.stringify({
            query: "What did Alex commit to?",
            filters: { scope: "all-mail", dateRange: "90d", sender: "", focus: "all" },
            availableRefinements: {
              scope: [{ value: "all-mail", label: "All mail", count: 0 }, { value: "inbox", label: "Inbox", count: 0 }],
              dateRange: [{ value: "90d", label: "Last 90 days", count: 0 }, { value: "all", label: "Any time", count: 0 }],
              focus: [{ value: "all", label: "All matches", count: 0 }, { value: "commitments", label: "Commitments", count: 0 }],
              sender: [{ value: "sam@example.com", label: "Sam <sam@example.com>", count: 1 }]
            },
            answerSummary: "Initial search suggests Alex committed to update the schedule next week.",
            commitments: "Preview evidence points to a schedule update commitment.",
            results: [
              { id: "1", subject: "Project schedule", folderName: "Inbox", whyItMatches: "Mentions the schedule change.", fromLine: "Sam <sam@example.com>", receivedDateTime: "2026-03-30T09:00:00Z", bodyPreview: "Let us update the schedule next week." }
            ],
            meta: {
              requestId: "search-start-mock",
              sessionId: "session-123",
              model: "gpt-5.4",
              queryStage: "start",
              scopeLabel: "All mail",
              evidenceMode: "preview-ranked",
              candidateCount: 14,
              resultCount: 1,
              hydratedCount: 0,
              lastUpdatedUtc: "2026-03-30T12:00:00Z"
            }
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/search/refine")) {
          return new Response(JSON.stringify({
            query: "What did Alex commit to?",
            filters: { scope: "all-mail", dateRange: "90d", sender: "sam@example.com", focus: "commitments" },
            availableRefinements: {
              scope: [{ value: "all-mail", label: "All mail", count: 0 }],
              dateRange: [{ value: "90d", label: "Last 90 days", count: 0 }],
              focus: [{ value: "all", label: "All matches", count: 0 }, { value: "commitments", label: "Commitments", count: 0 }],
              sender: [{ value: "sam@example.com", label: "Sam <sam@example.com>", count: 1 }]
            },
            answerSummary: "Alex committed to update the schedule next week.",
            commitments: "Update the schedule and confirm with the team.",
            results: [
              { id: "1", subject: "Project schedule", folderName: "Inbox", whyItMatches: "Hydrated full message confirms the commitment.", fromLine: "Sam <sam@example.com>", receivedDateTime: "2026-03-30T09:00:00Z", bodyPreview: "Let us update the schedule next week." }
            ],
            meta: {
              requestId: "search-refine-mock",
              warning: "Results cover a limited subset of this mailbox. Some messages may be missing; narrow the date range or scope.",
              usedFallback: false,
              sessionId: "session-123",
              model: "gpt-5.4",
              queryStage: "refine",
              scopeLabel: "All mail",
              evidenceMode: "hydrated-top-results",
              candidateCount: 3,
              resultCount: 1,
              hydratedCount: 3,
              lastUpdatedUtc: "2026-03-30T12:01:00Z"
            }
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/client-log")) {
          return new Response("", { status: 202 });
        }
        return new Response(JSON.stringify({ status: "ok" }), { status: 200, headers: { "Content-Type": "application/json" } });
      };

      window.Response = window.Response || class Response {
        constructor(body, init) { this._body = body; this.status = init.status; this.ok = this.status >= 200 && this.status < 300; this.headers = init.headers || {}; }
        async json() { return JSON.parse(this._body || "{}"); }
        async text() { return this._body; }
      };
    })();
