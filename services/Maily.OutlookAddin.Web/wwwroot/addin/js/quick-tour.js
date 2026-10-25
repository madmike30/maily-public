(function () {
  'use strict';
  const byId = id => document.getElementById(id);
  const panel = byId('tourPanel');
  if (!panel) return;
  const steps = [
    ['Start with the email', 'Summarize an email, then choose relevant action items to help draft a reply. Always review the result before sending.'],
    ['Your draft comes first', 'Write your rough reply in Outlook, then ask Maily to improve it. Maily rewrites that reply while keeping your signature and conversation history.'],
    ['Language follows your message', 'Summaries follow the main language of the email. Replies follow your existing draft, or the incoming email when the draft is empty. You can also choose a language.', 'card-language'],
    ['Make it sound right', 'Choose your preferred tone and length in Writing defaults. You can change these whenever you need.', 'card-tone'],
    ['Keep model choice simple', 'Automatic chooses the recommended fast model available to your connection. Choose another model only if you prefer it.', 'card-model'],
    ['Optional: learn your writing style', 'Style learning reads sent emails only when you choose Learn my style. It builds a writing profile to guide future drafts. Review your organization’s rules before starting; you can leave personalization off.', 'personalStyleSection']
  ];
  let index = 0;
  function clearHighlight() {
    for (const step of steps) if (step[2]) byId(step[2])?.classList.remove('tour-highlight');
  }
  function show() {
    clearHighlight();
    panel.hidden = false;
    const [title, text, target] = steps[index];
    byId('tourTitle').textContent = title;
    byId('tourText').textContent = text;
    byId('tourProgress').textContent = `${index + 1} of ${steps.length}`;
    byId('tourBack').disabled = index === 0;
    byId('tourNext').textContent = index === steps.length - 1 ? 'Finish' : 'Next';
    if (target) {
      const node = byId(target);
      if (node) {
        node.classList.add('tour-highlight');
        if (target === 'personalStyleSection') node.open = true;
      }
    }
    byId('tourTitle').focus();
  }
  function close() { panel.hidden = true; clearHighlight(); byId('tourStart').focus(); }
  panel.hidden = true;
  byId('tourStart').addEventListener('click', () => { index = 0; show(); });
  byId('tourBack').addEventListener('click', () => { if (index > 0) { index--; show(); } });
  byId('tourNext').addEventListener('click', () => {
    if (index < steps.length - 1) { index++; show(); return; }
    try { window.localStorage.setItem('maily.quickTour.v1.completed', 'true'); } catch (_) { /* Optional preference. */ }
    close();
  });
  byId('tourDismiss').addEventListener('click', close);
}());
