const STORAGE_KEY = 'crewSchedulerV5Data';

const state = {
  currentMonth: startOfMonth(new Date()),
  selectedEventId: null,
  selectedPersonId: null,
  data: null,
};

const els = {};

document.addEventListener('DOMContentLoaded', init);

function init() {
  cacheElements();

  const confirmParams = new URLSearchParams(window.location.search);
  if (confirmParams.get('confirm') === '1') {
    initConfirmView(confirmParams);
    return;
  }

  loadData();
  bindUI();
  renderAll();
}

function cacheElements() {
  const ids = [
    'calendarTitle','calendarGrid','personFormTitle','eventsList','alertsList','peopleList','personForm','personId','personName','personType','personRole','personEmail','personPhone','personContactPreference','personNotes','personFormResetBtn',
    'selectedEventTitle','selectedEventMeta','requirementsList','replacementSuggestions','assignmentsList','sendAllEventBtn','eventForm','eventId','eventTitleInput','eventLocationInput','eventStartInput','eventEndInput','eventDescriptionInput','eventFormResetBtn',
    'notificationLogList','messageTypePreview','messageSubjectPreview','messageBodyPreview','outlookConfigForm','clientIdInput','tenantIdInput','redirectUriInput','outlookConfigSummary','fakeConnectMicrosoftBtn','webhookConfigForm','emailWebhookInput','smsWebhookInput','webhookConfigSummary',
    'icsInput','newEventBtn','newPersonBtn','prevMonthBtn','nextMonthBtn','todayBtn','eventSearchInput','peopleSearchInput','simulateOutlookSyncBtn','exportDataBtn','importJsonInput','resetDemoBtn','clearLogBtn','addRequirementBtn','duplicateEventBtn'
  ];
  ids.forEach(id => els[id] = document.getElementById(id));
  els.navButtons = Array.from(document.querySelectorAll('.nav-btn'));
  els.sections = Array.from(document.querySelectorAll('.section'));
}

function bindUI() {
  els.navButtons.forEach(btn => btn.addEventListener('click', () => switchSection(btn.dataset.section)));
  els.prevMonthBtn.addEventListener('click', () => { state.currentMonth = addMonths(state.currentMonth, -1); renderCalendar(); });
  els.nextMonthBtn.addEventListener('click', () => { state.currentMonth = addMonths(state.currentMonth, 1); renderCalendar(); });
  els.todayBtn.addEventListener('click', () => { state.currentMonth = startOfMonth(new Date()); renderCalendar(); });
  els.newEventBtn.addEventListener('click', () => { switchSection('event-detail'); resetEventForm(); });
  els.newPersonBtn.addEventListener('click', () => { switchSection('people'); resetPersonForm(); });
  els.personForm.addEventListener('submit', savePersonFromForm);
  els.personFormResetBtn.addEventListener('click', resetPersonForm);
  els.eventForm.addEventListener('submit', saveEventFromForm);
  els.eventFormResetBtn.addEventListener('click', resetEventForm);
  els.sendAllEventBtn.addEventListener('click', sendAllForSelectedEvent);
  els.clearLogBtn.addEventListener('click', () => { state.data.notificationLog = []; persist(); renderNotificationLog(); });
  els.addRequirementBtn.addEventListener('click', addRequirementToSelectedEvent);
  els.duplicateEventBtn.addEventListener('click', duplicateSelectedEvent);
  els.icsInput.addEventListener('change', importIcsFile);
  els.simulateOutlookSyncBtn.addEventListener('click', simulateOutlookSync);
  els.exportDataBtn.addEventListener('click', exportJson);
  els.importJsonInput.addEventListener('change', importJson);
  els.resetDemoBtn.addEventListener('click', resetDemo);
  els.eventSearchInput.addEventListener('input', renderEventsList);
  els.peopleSearchInput.addEventListener('input', renderPeopleList);
  els.outlookConfigForm.addEventListener('submit', saveOutlookConfig);
  els.webhookConfigForm.addEventListener('submit', saveWebhookConfig);
  els.fakeConnectMicrosoftBtn.addEventListener('click', fakeConnectMicrosoft);
  els.messageTypePreview.addEventListener('change', updateMessagePreviewForSelectedEvent);
}

function switchSection(sectionId) {
  els.navButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.section === sectionId));
  els.sections.forEach(section => section.classList.toggle('active', section.id === sectionId));
}

function loadData() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    state.data = createDemoData();
    persist();
    return;
  }
  try {
    state.data = JSON.parse(raw);
  } catch (error) {
    state.data = createDemoData();
    persist();
  }
  if (!state.data.notificationLog) state.data.notificationLog = [];
  if (!state.data.config) state.data.config = { outlook: {}, webhooks: {} };
  if (!state.selectedEventId && state.data.events[0]) state.selectedEventId = state.data.events[0].id;
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.data));
}

function renderAll() {
  renderCalendar();
  renderEventsList();
  renderAlerts();
  renderPeopleList();
  renderSelectedEvent();
  renderNotificationLog();
  renderConfigSummaries();
}

function renderCalendar() {
  const month = state.currentMonth;
  els.calendarTitle.textContent = month.toLocaleDateString('fr-CA', { month: 'long', year: 'numeric' });
  const grid = els.calendarGrid;
  grid.innerHTML = '';
  ['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'].forEach(d => {
    const el = document.createElement('div');
    el.className = 'day-name';
    el.textContent = d;
    grid.appendChild(el);
  });

  const days = buildCalendarDays(month);
  const todayKey = dateKey(new Date());
  days.forEach(day => {
    const dayCell = document.createElement('div');
    dayCell.className = 'calendar-day';
    if (!day.inCurrentMonth) dayCell.classList.add('muted-day');
    if (day.key === todayKey) dayCell.classList.add('today');

    const number = document.createElement('div');
    number.className = 'calendar-day-number';
    number.textContent = day.date.getDate();
    dayCell.appendChild(number);

    const events = getEventsForDay(day.key);
    events.slice(0, 4).forEach(event => {
      const chip = document.createElement('div');
      chip.className = 'calendar-event-chip';
      chip.textContent = `${formatTime(event.start)} • ${event.title}`;
      chip.addEventListener('click', () => selectEvent(event.id, true));
      dayCell.appendChild(chip);
    });

    if (events.length > 4) {
      const more = document.createElement('div');
      more.className = 'small-note';
      more.textContent = `+${events.length - 4} autre(s)`;
      dayCell.appendChild(more);
    }

    dayCell.addEventListener('dblclick', () => {
      resetEventForm();
      els.eventStartInput.value = toDateTimeLocal(day.date);
      const end = new Date(day.date.getTime() + 2 * 60 * 60 * 1000);
      els.eventEndInput.value = toDateTimeLocal(end);
      switchSection('event-detail');
    });

    grid.appendChild(dayCell);
  });
}

function renderEventsList() {
  const query = (els.eventSearchInput.value || '').trim().toLowerCase();
  const wrap = els.eventsList;
  wrap.innerHTML = '';
  const events = [...state.data.events].sort((a, b) => a.start.localeCompare(b.start)).filter(event => {
    if (!query) return true;
    return `${event.title} ${event.location || ''} ${event.description || ''}`.toLowerCase().includes(query);
  });

  if (!events.length) {
    wrap.innerHTML = '<div class="list-item muted">Aucun événement trouvé.</div>';
    return;
  }

  events.forEach(event => {
    const coverage = getCoverageSummary(event);
    const div = document.createElement('div');
    div.className = 'list-item event-card clickable' + (state.selectedEventId === event.id ? ' selected' : '');
    div.innerHTML = `
      <div class="row">
        <strong>${escapeHtml(event.title)}</strong>
        <span class="pill ${coverage.variant}">${coverage.label}</span>
      </div>
      <div class="kv"><span>${formatDateTime(event.start)} → ${formatDateTime(event.end)}</span><span>${escapeHtml(event.location || 'Sans lieu')}</span></div>
      <div class="small-note">${event.requirements.length} rôle(s) requis • ${event.assignments.length} assignation(s)</div>
    `;
    div.addEventListener('click', () => selectEvent(event.id, false));
    wrap.appendChild(div);
  });
}

function renderAlerts() {
  const wrap = els.alertsList;
  wrap.innerHTML = '';
  const alerts = [];
  state.data.events.forEach(event => {
    const uncovered = getUncoveredRequirements(event);
    uncovered.forEach(req => alerts.push({ type: 'missing', event, req }));
    event.assignments.filter(a => a.status === 'declined').forEach(a => alerts.push({ type: 'declined', event, assignment: a }));
  });

  if (!alerts.length) {
    wrap.innerHTML = '<div class="list-item alert-card"><strong>Aucune alerte.</strong><div class="muted">Tous les événements semblent couverts pour l’instant.</div></div>';
    return;
  }

  alerts.forEach(alert => {
    const div = document.createElement('div');
    div.className = 'list-item alert-card clickable';
    if (alert.type === 'missing') {
      div.innerHTML = `<div class="row"><strong>${escapeHtml(alert.event.title)}</strong><span class="pill red">Rôle manquant</span></div><div class="muted">${escapeHtml(alert.req.role)} • ${alert.req.count} requis</div>`;
    } else {
      const person = getPersonById(alert.assignment.personId);
      div.innerHTML = `<div class="row"><strong>${escapeHtml(alert.event.title)}</strong><span class="pill amber">Refus reçu</span></div><div class="muted">${escapeHtml(person?.name || 'Personne inconnue')} • ${escapeHtml(alert.assignment.role)}</div>`;
    }
    div.addEventListener('click', () => selectEvent(alert.event.id, true));
    wrap.appendChild(div);
  });
}

function renderPeopleList() {
  const query = (els.peopleSearchInput.value || '').trim().toLowerCase();
  const wrap = els.peopleList;
  wrap.innerHTML = '';
  const people = [...state.data.people].sort((a,b) => a.name.localeCompare(b.name, 'fr')).filter(person => {
    if (!query) return true;
    return `${person.name} ${person.role || ''} ${person.email || ''} ${person.phone || ''}`.toLowerCase().includes(query);
  });
  if (!people.length) {
    wrap.innerHTML = '<div class="list-item muted">Aucune personne trouvée.</div>';
    return;
  }
  people.forEach(person => {
    const div = document.createElement('div');
    div.className = 'list-item person-card clickable' + (state.selectedPersonId === person.id ? ' selected' : '');
    div.innerHTML = `
      <div class="row">
        <strong>${escapeHtml(person.name)}</strong>
        <span class="pill ${person.type === 'pigiste' ? 'blue' : 'green'}">${person.type === 'pigiste' ? 'Pigiste' : 'Employé'}</span>
      </div>
      <div class="muted">${escapeHtml(person.role || 'Rôle non précisé')}</div>
      <div class="small-note">Préférence: ${contactPreferenceLabel(person.contactPreference)} • ${escapeHtml(person.email || 'pas d’email')} • ${escapeHtml(person.phone || 'pas de téléphone')}</div>
      <div class="assignment-actions">
        <button class="ghost-btn sm" data-action="edit">Modifier</button>
        <button class="ghost-btn sm danger" data-action="delete">Supprimer</button>
      </div>
    `;
    div.querySelector('[data-action="edit"]').addEventListener('click', e => { e.stopPropagation(); loadPersonIntoForm(person.id); });
    div.querySelector('[data-action="delete"]').addEventListener('click', e => { e.stopPropagation(); deletePerson(person.id); });
    div.addEventListener('click', () => { state.selectedPersonId = person.id; renderPeopleList(); });
    wrap.appendChild(div);
  });
}

function renderSelectedEvent() {
  const event = getSelectedEvent();
  if (!event) {
    els.selectedEventTitle.textContent = 'Aucun événement sélectionné';
    els.selectedEventMeta.innerHTML = '<div class="muted">Choisis un événement dans le calendrier ou crée-en un nouveau.</div>';
    els.requirementsList.innerHTML = '<div class="list-item muted">Aucun rôle à afficher.</div>';
    els.assignmentsList.innerHTML = '<div class="list-item muted">Aucune assignation.</div>';
    els.replacementSuggestions.innerHTML = '<div class="list-item muted">Aucune suggestion.</div>';
    updateMessagePreviewForSelectedEvent();
    return;
  }

  const coverage = getCoverageSummary(event);
  els.selectedEventTitle.textContent = event.title;
  els.selectedEventMeta.innerHTML = `
    <div class="row"><span>${formatDateTime(event.start)} → ${formatDateTime(event.end)}</span><span class="pill ${coverage.variant}">${coverage.label}</span></div>
    <div>${escapeHtml(event.location || 'Sans lieu')}</div>
    <div>${escapeHtml(event.description || 'Aucune description')}</div>
  `;

  renderRequirements(event);
  renderAssignments(event);
  renderReplacementSuggestions(event);
  loadEventIntoForm(event.id, false);
  updateMessagePreviewForSelectedEvent();
}

function renderRequirements(event) {
  const wrap = els.requirementsList;
  wrap.innerHTML = '';
  if (!event.requirements.length) {
    wrap.innerHTML = '<div class="list-item muted">Aucun rôle requis pour cet événement.</div>';
    return;
  }
  event.requirements.forEach(req => {
    const assignedCount = event.assignments.filter(a => a.role === req.role && a.status !== 'declined').length;
    const div = document.createElement('div');
    div.className = 'list-item requirement-card';
    div.innerHTML = `
      <div class="row">
        <strong>${escapeHtml(req.role)}</strong>
        <span class="pill ${assignedCount >= req.count ? 'green' : 'amber'}">${assignedCount}/${req.count}</span>
      </div>
      <div class="small-note">${escapeHtml(req.notes || 'Aucune note')}</div>
      <div class="requirement-actions">
        <button class="ghost-btn sm" data-action="assign">Assigner</button>
        <button class="ghost-btn sm" data-action="edit">Modifier</button>
        <button class="ghost-btn sm danger" data-action="delete">Supprimer</button>
      </div>
    `;
    div.querySelector('[data-action="assign"]').addEventListener('click', () => assignPersonPrompt(event.id, req.role));
    div.querySelector('[data-action="edit"]').addEventListener('click', () => editRequirement(event.id, req.id));
    div.querySelector('[data-action="delete"]').addEventListener('click', () => deleteRequirement(event.id, req.id));
    wrap.appendChild(div);
  });
}

function renderAssignments(event) {
  const wrap = els.assignmentsList;
  wrap.innerHTML = '';
  if (!event.assignments.length) {
    wrap.innerHTML = '<div class="list-item muted">Aucune personne assignée pour l’instant.</div>';
    return;
  }
  event.assignments.forEach(assignment => {
    const person = getPersonById(assignment.personId);
    const div = document.createElement('div');
    div.className = 'list-item assignment-card';
    div.innerHTML = `
      <div class="row">
        <strong>${escapeHtml(person?.name || 'Personne supprimée')}</strong>
        <span class="status-badge ${assignment.status}">${statusLabel(assignment.status)}</span>
      </div>
      <div class="muted">${escapeHtml(assignment.role)} • ${contactPreferenceLabel(person?.contactPreference || 'none')}</div>
      <div class="small-note">${escapeHtml(person?.email || 'pas d’email')} • ${escapeHtml(person?.phone || 'pas de téléphone')}</div>
      <div class="small-note">Lien : <code>${escapeHtml(buildConfirmUrl(event.id, assignment.id))}</code></div>
      <div class="assignment-actions">
        <button class="primary-btn sm" data-action="send">Envoyer</button>
        <button class="ghost-btn sm" data-action="confirm">Confirmer</button>
        <button class="ghost-btn sm danger" data-action="decline">Refuser</button>
        <button class="ghost-btn sm danger" data-action="remove">Retirer</button>
      </div>
    `;
    div.querySelector('[data-action="send"]').addEventListener('click', () => sendAssignment(event.id, assignment.id));
    div.querySelector('[data-action="confirm"]').addEventListener('click', () => setAssignmentStatus(event.id, assignment.id, 'confirmed'));
    div.querySelector('[data-action="decline"]').addEventListener('click', () => setAssignmentStatus(event.id, assignment.id, 'declined'));
    div.querySelector('[data-action="remove"]').addEventListener('click', () => removeAssignment(event.id, assignment.id));
    wrap.appendChild(div);
  });
}

function renderReplacementSuggestions(event) {
  const wrap = els.replacementSuggestions;
  wrap.innerHTML = '';
  const uncovered = getUncoveredRequirements(event);
  const suggestions = [];
  uncovered.forEach(req => {
    const candidates = getCandidatesForRole(event, req.role).slice(0, 3);
    candidates.forEach(person => suggestions.push({ role: req.role, person }));
  });
  event.assignments.filter(a => a.status === 'declined').forEach(a => {
    const candidates = getCandidatesForRole(event, a.role).slice(0, 3);
    candidates.forEach(person => suggestions.push({ role: a.role, person }));
  });

  if (!suggestions.length) {
    wrap.innerHTML = '<div class="list-item muted">Aucune suggestion pour le moment.</div>';
    return;
  }

  suggestions.forEach(item => {
    const div = document.createElement('div');
    div.className = 'list-item';
    div.innerHTML = `
      <div class="row">
        <strong>${escapeHtml(item.person.name)}</strong>
        <span class="pill blue">${escapeHtml(item.role)}</span>
      </div>
      <div class="small-note">${escapeHtml(item.person.role || 'Rôle non précisé')} • ${contactPreferenceLabel(item.person.contactPreference)}</div>
      <div class="assignment-actions">
        <button class="primary-btn sm">Assigner</button>
      </div>
    `;
    div.querySelector('button').addEventListener('click', () => {
      assignPersonToEvent(event.id, item.role, item.person.id);
      selectEvent(event.id, true);
    });
    wrap.appendChild(div);
  });
}

function renderNotificationLog() {
  const wrap = els.notificationLogList;
  wrap.innerHTML = '';
  const logs = [...state.data.notificationLog].reverse();
  if (!logs.length) {
    wrap.innerHTML = '<div class="list-item muted">Aucun envoi dans le journal.</div>';
    return;
  }
  logs.forEach(log => {
    const div = document.createElement('div');
    div.className = 'list-item log-card';
    div.innerHTML = `
      <div class="row">
        <strong>${escapeHtml(log.personName)}</strong>
        <span class="pill ${log.channel === 'sms' ? 'amber' : 'blue'}">${log.channel.toUpperCase()}</span>
      </div>
      <div class="muted">${escapeHtml(log.eventTitle)} • ${escapeHtml(log.role)}</div>
      <div class="small-note">${escapeHtml(log.destination)} • ${formatDateTime(log.sentAt)}</div>
    `;
    wrap.appendChild(div);
  });
}

function renderConfigSummaries() {
  const outlook = state.data.config.outlook || {};
  els.clientIdInput.value = outlook.clientId || '';
  els.tenantIdInput.value = outlook.tenantId || '';
  els.redirectUriInput.value = outlook.redirectUri || '';
  els.outlookConfigSummary.innerHTML = `
    <div class="kv"><strong>Client ID</strong><span>${escapeHtml(outlook.clientId || 'Non configuré')}</span></div>
    <div class="kv"><strong>Tenant ID</strong><span>${escapeHtml(outlook.tenantId || 'Non configuré')}</span></div>
    <div class="kv"><strong>Redirect URI</strong><span>${escapeHtml(outlook.redirectUri || 'Non configuré')}</span></div>
  `;

  const webhooks = state.data.config.webhooks || {};
  els.emailWebhookInput.value = webhooks.email || '';
  els.smsWebhookInput.value = webhooks.sms || '';
  els.webhookConfigSummary.innerHTML = `
    <div class="kv"><strong>Email webhook</strong><span>${escapeHtml(webhooks.email || 'Non configuré')}</span></div>
    <div class="kv"><strong>SMS webhook</strong><span>${escapeHtml(webhooks.sms || 'Non configuré')}</span></div>
  `;
}

function updateMessagePreviewForSelectedEvent() {
  const event = getSelectedEvent();
  const type = els.messageTypePreview.value;
  if (!event) {
    els.messageSubjectPreview.value = '';
    els.messageBodyPreview.value = '';
    return;
  }
  const assignment = event.assignments[0];
  const person = assignment ? getPersonById(assignment.personId) : null;
  const preview = buildMessageContent(event, assignment, person, type);
  els.messageSubjectPreview.value = preview.subject;
  els.messageBodyPreview.value = preview.body;
}

function selectEvent(eventId, goToEventSection) {
  state.selectedEventId = eventId;
  if (goToEventSection) switchSection('event-detail');
  renderEventsList();
  renderSelectedEvent();
  renderAlerts();
}

function getSelectedEvent() {
  return state.data.events.find(event => event.id === state.selectedEventId) || null;
}

function savePersonFromForm(e) {
  e.preventDefault();
  const id = els.personId.value || uid('person');
  const person = {
    id,
    name: els.personName.value.trim(),
    type: els.personType.value,
    role: els.personRole.value.trim(),
    email: els.personEmail.value.trim(),
    phone: els.personPhone.value.trim(),
    contactPreference: els.personContactPreference.value,
    notes: els.personNotes.value.trim(),
  };
  if (!person.name) return;
  upsert(state.data.people, person);
  state.selectedPersonId = person.id;
  persist();
  resetPersonForm(false);
  renderPeopleList();
  renderSelectedEvent();
}

function resetPersonForm(clearSelected = true) {
  els.personForm.reset();
  els.personId.value = '';
  if (clearSelected) state.selectedPersonId = null;
  els.personFormTitle.textContent = 'Ajouter / modifier une personne';
  renderPeopleList();
}

function loadPersonIntoForm(personId) {
  const person = getPersonById(personId);
  if (!person) return;
  state.selectedPersonId = person.id;
  switchSection('people');
  els.personFormTitle.textContent = `Modifier ${person.name}`;
  els.personId.value = person.id;
  els.personName.value = person.name || '';
  els.personType.value = person.type || 'pigiste';
  els.personRole.value = person.role || '';
  els.personEmail.value = person.email || '';
  els.personPhone.value = person.phone || '';
  els.personContactPreference.value = person.contactPreference || 'email';
  els.personNotes.value = person.notes || '';
  renderPeopleList();
}

function deletePerson(personId) {
  const person = getPersonById(personId);
  if (!person) return;
  if (!confirm(`Supprimer ${person.name} ?`)) return;
  state.data.people = state.data.people.filter(p => p.id !== personId);
  state.data.events.forEach(event => {
    event.assignments = event.assignments.filter(a => a.personId !== personId);
  });
  if (state.selectedPersonId === personId) state.selectedPersonId = null;
  persist();
  renderAll();
}

function saveEventFromForm(e) {
  e.preventDefault();
  const start = els.eventStartInput.value;
  const end = els.eventEndInput.value;
  if (!start || !end || start > end) {
    alert('Vérifie les dates: le début doit être avant ou égal à la fin.');
    return;
  }
  const existing = els.eventId.value ? state.data.events.find(e => e.id === els.eventId.value) : null;
  const event = {
    id: els.eventId.value || uid('event'),
    title: els.eventTitleInput.value.trim(),
    location: els.eventLocationInput.value.trim(),
    start,
    end,
    description: els.eventDescriptionInput.value.trim(),
    requirements: existing?.requirements || [],
    assignments: existing?.assignments || [],
    importedFrom: existing?.importedFrom || 'manual',
  };
  if (!event.title) return;
  upsert(state.data.events, event);
  state.selectedEventId = event.id;
  persist();
  renderAll();
  switchSection('event-detail');
}

function resetEventForm(clearSelected = false) {
  els.eventForm.reset();
  els.eventId.value = '';
  if (clearSelected) state.selectedEventId = null;
  const start = new Date();
  const end = new Date(start.getTime() + 2 * 60 * 60 * 1000);
  els.eventStartInput.value = toDateTimeLocal(start);
  els.eventEndInput.value = toDateTimeLocal(end);
}

function loadEventIntoForm(eventId, switchTab = true) {
  const event = state.data.events.find(e => e.id === eventId);
  if (!event) return;
  if (switchTab) switchSection('event-detail');
  els.eventId.value = event.id;
  els.eventTitleInput.value = event.title || '';
  els.eventLocationInput.value = event.location || '';
  els.eventStartInput.value = event.start;
  els.eventEndInput.value = event.end;
  els.eventDescriptionInput.value = event.description || '';
}

function addRequirementToSelectedEvent() {
  const event = getSelectedEvent();
  if (!event) return alert('Choisis un événement d’abord.');
  const role = prompt('Nom du rôle requis ?', 'Technicien son');
  if (!role) return;
  const count = Number(prompt('Combien de personnes pour ce rôle ?', '1') || '1');
  const notes = prompt('Note pour ce rôle ?', '') || '';
  event.requirements.push({ id: uid('req'), role: role.trim(), count: Math.max(1, count || 1), notes: notes.trim() });
  persist();
  renderSelectedEvent();
  renderEventsList();
  renderAlerts();
}

function editRequirement(eventId, reqId) {
  const event = state.data.events.find(e => e.id === eventId);
  const req = event?.requirements.find(r => r.id === reqId);
  if (!req) return;
  const role = prompt('Modifier le rôle', req.role);
  if (!role) return;
  const count = Number(prompt('Modifier le nombre requis', String(req.count)) || req.count);
  const notes = prompt('Modifier la note', req.notes || '') || '';
  req.role = role.trim();
  req.count = Math.max(1, count || 1);
  req.notes = notes.trim();
  persist();
  renderSelectedEvent();
  renderEventsList();
  renderAlerts();
}

function deleteRequirement(eventId, reqId) {
  const event = state.data.events.find(e => e.id === eventId);
  if (!event) return;
  event.requirements = event.requirements.filter(r => r.id !== reqId);
  persist();
  renderSelectedEvent();
  renderEventsList();
  renderAlerts();
}

function assignPersonPrompt(eventId, role) {
  const candidates = getCandidatesForRole(state.data.events.find(e => e.id === eventId), role);
  if (!candidates.length) {
    alert(`Aucun candidat trouvé pour le rôle: ${role}`);
    return;
  }
  const lines = candidates.map((person, index) => `${index + 1}. ${person.name} (${person.role || 'sans rôle'})`).join('\n');
  const choice = Number(prompt(`Choisis une personne pour “${role}”:\n\n${lines}`, '1') || '0');
  const person = candidates[choice - 1];
  if (!person) return;
  assignPersonToEvent(eventId, role, person.id);
  selectEvent(eventId, true);
}

function assignPersonToEvent(eventId, role, personId) {
  const event = state.data.events.find(e => e.id === eventId);
  if (!event) return;
  const existing = event.assignments.find(a => a.personId === personId && a.role === role);
  if (existing) return;
  event.assignments.push({ id: uid('asg'), personId, role, status: 'pending', invitedAt: null, respondedAt: null });
  persist();
  renderAll();
}

function removeAssignment(eventId, assignmentId) {
  const event = state.data.events.find(e => e.id === eventId);
  if (!event) return;
  event.assignments = event.assignments.filter(a => a.id !== assignmentId);
  persist();
  renderAll();
}

function setAssignmentStatus(eventId, assignmentId, status) {
  const event = state.data.events.find(e => e.id === eventId);
  const assignment = event?.assignments.find(a => a.id === assignmentId);
  if (!assignment) return;
  assignment.status = status;
  assignment.respondedAt = new Date().toISOString().slice(0,16);
  persist();
  renderAll();
}

function sendAssignment(eventId, assignmentId) {
  const event = state.data.events.find(e => e.id === eventId);
  const assignment = event?.assignments.find(a => a.id === assignmentId);
  if (!event || !assignment) return;
  const person = getPersonById(assignment.personId);
  if (!person) return;
  const channels = resolveChannels(person.contactPreference);
  if (!channels.length) {
    alert(`${person.name} a la préférence “aucune”.`);
    return;
  }
  channels.forEach(channel => {
    const preview = buildMessageContent(event, assignment, person, channel);
    const destination = channel === 'sms' ? (person.phone || 'numéro manquant') : (person.email || 'email manquant');
    state.data.notificationLog.push({
      id: uid('log'),
      sentAt: new Date().toISOString().slice(0,16),
      channel,
      eventId: event.id,
      eventTitle: event.title,
      assignmentId: assignment.id,
      role: assignment.role,
      personId: person.id,
      personName: person.name,
      destination,
      subject: preview.subject,
      body: preview.body,
    });
  });
  assignment.status = 'sent';
  assignment.invitedAt = new Date().toISOString().slice(0,16);
  persist();
  renderAll();
  switchSection('outbox');
}

function sendAllForSelectedEvent() {
  const event = getSelectedEvent();
  if (!event) return;
  event.assignments.forEach(a => sendAssignment(event.id, a.id));
}

function buildMessageContent(event, assignment, person, channel) {
  const confirmUrl = buildConfirmUrl(event.id, assignment?.id || '');
  const subject = `Disponibilité demandée • ${event.title}`;
  const body = channel === 'sms'
    ? `Bonjour ${person?.name || ''}, dispo pour ${event.title} (${assignment?.role || 'équipe'}) le ${formatDateTime(event.start)} à ${event.location || 'lieu à confirmer'} ? Confirme ici: ${confirmUrl}`
    : `Bonjour ${person?.name || ''},\n\nNous aimerions confirmer ta disponibilité pour l’événement “${event.title}”.\n\nRôle: ${assignment?.role || 'À confirmer'}\nDate: ${formatDateTime(event.start)} → ${formatDateTime(event.end)}\nLieu: ${event.location || 'À confirmer'}\n\nLien de confirmation:\n${confirmUrl}\n\nMerci.`;
  return { subject, body };
}

function buildConfirmUrl(eventId, assignmentId) {
  const base = `${window.location.origin}${window.location.pathname}`;
  return `${base}?confirm=1&eventId=${encodeURIComponent(eventId)}&assignmentId=${encodeURIComponent(assignmentId)}`;
}

function saveOutlookConfig(e) {
  e.preventDefault();
  state.data.config.outlook = {
    clientId: els.clientIdInput.value.trim(),
    tenantId: els.tenantIdInput.value.trim(),
    redirectUri: els.redirectUriInput.value.trim(),
  };
  persist();
  renderConfigSummaries();
}

function saveWebhookConfig(e) {
  e.preventDefault();
  state.data.config.webhooks = {
    email: els.emailWebhookInput.value.trim(),
    sms: els.smsWebhookInput.value.trim(),
  };
  persist();
  renderConfigSummaries();
}

function fakeConnectMicrosoft() {
  alert('Bouton prêt. Une fois hébergé en HTTPS et connecté à Entra ID, ce bouton pourra lancer la vraie connexion Microsoft.');
}

function simulateOutlookSync() {
  const now = new Date();
  const event = {
    id: uid('event'),
    title: 'Sync Outlook démo',
    location: 'Centre-ville Montréal',
    start: toDateTimeLocal(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 4, 9, 0)),
    end: toDateTimeLocal(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 4, 17, 0)),
    description: 'Événement ajouté par la simulation de sync Outlook.',
    importedFrom: 'outlook-simulated',
    requirements: [
      { id: uid('req'), role: 'Technicien son', count: 2, notes: '' },
      { id: uid('req'), role: 'Technicien éclairage', count: 1, notes: '' },
    ],
    assignments: [],
  };
  state.data.events.push(event);
  state.selectedEventId = event.id;
  persist();
  renderAll();
  switchSection('event-detail');
}

function importIcsFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const text = String(reader.result || '');
    const imported = parseICS(text);
    if (!imported.length) {
      alert('Aucun événement reconnu dans ce fichier .ics');
      return;
    }
    imported.forEach(event => state.data.events.push(event));
    state.selectedEventId = imported[0].id;
    persist();
    renderAll();
    switchSection('dashboard');
    e.target.value = '';
  };
  reader.readAsText(file);
}

function exportJson() {
  const blob = new Blob([JSON.stringify(state.data, null, 2)], { type: 'application/json' });
  downloadBlob(blob, `crew-scheduler-export-${new Date().toISOString().slice(0,10)}.json`);
}

function importJson(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      state.data = JSON.parse(String(reader.result || '{}'));
      if (!state.data.config) state.data.config = { outlook: {}, webhooks: {} };
      if (!state.data.notificationLog) state.data.notificationLog = [];
      state.selectedEventId = state.data.events?.[0]?.id || null;
      persist();
      renderAll();
      alert('Import JSON réussi.');
    } catch (error) {
      alert('Le fichier JSON est invalide.');
    }
    e.target.value = '';
  };
  reader.readAsText(file);
}

function resetDemo() {
  if (!confirm('Réinitialiser toutes les données locales de la démo ?')) return;
  state.data = createDemoData();
  state.currentMonth = startOfMonth(new Date());
  persist();
  renderAll();
}

function duplicateSelectedEvent() {
  const event = getSelectedEvent();
  if (!event) return;
  const clone = JSON.parse(JSON.stringify(event));
  clone.id = uid('event');
  clone.title = `${clone.title} - copie`;
  clone.assignments = [];
  clone.requirements = clone.requirements.map(req => ({ ...req, id: uid('req') }));
  state.data.events.push(clone);
  state.selectedEventId = clone.id;
  persist();
  renderAll();
}

function initConfirmView(params) {
  const tpl = document.getElementById('confirmViewTemplate');
  document.body.classList.add('confirm-mode');
  document.body.innerHTML = '';

  if (!tpl) {
    document.body.innerHTML = '<div style="padding:24px">Template de confirmation introuvable.</div>';
    return;
  }
  document.body.appendChild(tpl.content.cloneNode(true));

  loadData();
  const event = state.data.events.find(e => e.id === params.get('eventId'));
  const assignment = event?.assignments.find(a => a.id === params.get('assignmentId'));
  const person = assignment ? getPersonById(assignment.personId) : null;
  const resultEl = document.getElementById('confirmResultMessage');

  if (!event || !assignment) {
    document.getElementById('confirmEventTitle').textContent = 'Invitation introuvable';
    document.getElementById('confirmPersonLine').textContent = 'Le lien semble invalide ou expiré.';
    document.getElementById('confirmRoleLine').textContent = '';
    document.getElementById('confirmDateLine').textContent = '';
    document.getElementById('confirmAcceptBtn').remove();
    document.getElementById('confirmDeclineBtn').remove();
    return;
  }

  document.getElementById('confirmEventTitle').textContent = event.title;
  document.getElementById('confirmPersonLine').textContent = `Personne: ${person?.name || 'Inconnue'}`;
  document.getElementById('confirmRoleLine').textContent = `Rôle: ${assignment.role}`;
  document.getElementById('confirmDateLine').textContent = `Date: ${formatDateTime(event.start)} → ${formatDateTime(event.end)} • ${event.location || 'Sans lieu'}`;

  document.getElementById('confirmAcceptBtn').addEventListener('click', () => {
    assignment.status = 'confirmed';
    assignment.respondedAt = new Date().toISOString().slice(0,16);
    persist();
    resultEl.textContent = 'Merci, ta présence est confirmée.';
  });
  document.getElementById('confirmDeclineBtn').addEventListener('click', () => {
    assignment.status = 'declined';
    assignment.respondedAt = new Date().toISOString().slice(0,16);
    persist();
    resultEl.textContent = 'Merci, le refus a été enregistré.';
  });
}

function parseICS(text) {
  const normalized = text.replace(/\r\n[ \t]/g, '').replace(/\r/g, '');
  const parts = normalized.split('BEGIN:VEVENT').slice(1);
  return parts.map(part => {
    const endPart = part.split('END:VEVENT')[0];
    const summary = pickIcsLine(endPart, 'SUMMARY') || 'Événement importé';
    const location = pickIcsLine(endPart, 'LOCATION') || '';
    const description = pickIcsLine(endPart, 'DESCRIPTION') || '';
    const dtStartRaw = pickIcsLineFlexible(endPart, 'DTSTART');
    const dtEndRaw = pickIcsLineFlexible(endPart, 'DTEND');
    if (!dtStartRaw) return null;
    const start = parseIcsDate(dtStartRaw);
    const end = parseIcsDate(dtEndRaw) || new Date(start.getTime() + 2 * 60 * 60 * 1000);
    return {
      id: uid('event'),
      title: summary,
      location,
      description,
      start: toDateTimeLocal(start),
      end: toDateTimeLocal(end),
      importedFrom: 'ics',
      requirements: [],
      assignments: [],
    };
  }).filter(Boolean);
}

function pickIcsLine(block, key) {
  const match = block.match(new RegExp(`^${key}:(.*)$`, 'm'));
  return match ? decodeIcsText(match[1].trim()) : '';
}

function pickIcsLineFlexible(block, key) {
  const match = block.match(new RegExp(`^${key}(;[^:]*)?:(.*)$`, 'm'));
  return match ? match[2].trim() : '';
}

function parseIcsDate(value) {
  if (!value) return null;
  if (/^\d{8}T\d{6}Z$/.test(value)) {
    const y = Number(value.slice(0,4));
    const m = Number(value.slice(4,6)) - 1;
    const d = Number(value.slice(6,8));
    const hh = Number(value.slice(9,11));
    const mm = Number(value.slice(11,13));
    const ss = Number(value.slice(13,15));
    return new Date(Date.UTC(y,m,d,hh,mm,ss));
  }
  if (/^\d{8}T\d{6}$/.test(value)) {
    const y = Number(value.slice(0,4));
    const m = Number(value.slice(4,6)) - 1;
    const d = Number(value.slice(6,8));
    const hh = Number(value.slice(9,11));
    const mm = Number(value.slice(11,13));
    const ss = Number(value.slice(13,15));
    return new Date(y,m,d,hh,mm,ss);
  }
  if (/^\d{8}$/.test(value)) {
    const y = Number(value.slice(0,4));
    const m = Number(value.slice(4,6)) - 1;
    const d = Number(value.slice(6,8));
    return new Date(y,m,d,8,0,0);
  }
  const maybe = new Date(value);
  return Number.isNaN(maybe.getTime()) ? null : maybe;
}

function decodeIcsText(str) {
  return str.replace(/\\n/g, '\n').replace(/\\,/g, ',').replace(/\\;/g, ';');
}

function getEventsForDay(dayKeyValue) {
  return state.data.events.filter(event => dateKey(new Date(event.start)) === dayKeyValue).sort((a,b) => a.start.localeCompare(b.start));
}

function getCoverageSummary(event) {
  const uncovered = getUncoveredRequirements(event).length;
  const declined = event.assignments.filter(a => a.status === 'declined').length;
  if (uncovered === 0 && declined === 0 && event.requirements.length > 0) return { label: 'Complet', variant: 'green' };
  if (declined > 0) return { label: 'Remplacement requis', variant: 'red' };
  if (event.requirements.length === 0) return { label: 'Sans besoins', variant: 'blue' };
  return { label: 'Incomplet', variant: 'amber' };
}

function getUncoveredRequirements(event) {
  return event.requirements.filter(req => {
    const activeAssigned = event.assignments.filter(a => a.role === req.role && a.status !== 'declined').length;
    return activeAssigned < req.count;
  });
}

function getCandidatesForRole(event, role) {
  const takenIds = new Set(event.assignments.filter(a => a.status !== 'declined').map(a => a.personId));
  return state.data.people.filter(person => {
    const matchesRole = !person.role || person.role.toLowerCase().includes(role.toLowerCase()) || role.toLowerCase().includes(person.role.toLowerCase());
    return matchesRole && !takenIds.has(person.id);
  }).sort((a,b) => scorePersonForRole(b, role) - scorePersonForRole(a, role));
}

function scorePersonForRole(person, role) {
  let score = 0;
  if ((person.role || '').toLowerCase() === role.toLowerCase()) score += 5;
  if ((person.role || '').toLowerCase().includes(role.toLowerCase())) score += 3;
  if (person.type === 'employe') score += 1;
  if (person.contactPreference === 'both') score += 1;
  return score;
}

function getPersonById(personId) {
  return state.data.people.find(person => person.id === personId) || null;
}

function resolveChannels(pref) {
  switch (pref) {
    case 'email': return ['email'];
    case 'sms': return ['sms'];
    case 'both': return ['email', 'sms'];
    default: return [];
  }
}

function contactPreferenceLabel(pref) {
  return pref === 'both' ? 'Email + SMS' : pref === 'sms' ? 'SMS' : pref === 'none' ? 'Aucune' : 'Email';
}

function statusLabel(status) {
  return status === 'sent' ? 'Envoyé' : status === 'confirmed' ? 'Confirmé' : status === 'declined' ? 'Refusé' : 'En attente';
}

function formatDateTime(value) {
  const date = new Date(value);
  return date.toLocaleString('fr-CA', { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
}

function formatTime(value) {
  const date = new Date(value);
  return date.toLocaleTimeString('fr-CA', { hour: '2-digit', minute: '2-digit' });
}

function toDateTimeLocal(date) {
  const pad = n => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date, months) {
  return new Date(date.getFullYear(), date.getMonth() + months, 1);
}

function buildCalendarDays(monthDate) {
  const first = startOfMonth(monthDate);
  const firstWeekday = (first.getDay() + 6) % 7;
  const start = new Date(first);
  start.setDate(first.getDate() - firstWeekday);
  const days = [];
  for (let i = 0; i < 42; i++) {
    const date = new Date(start);
    date.setDate(start.getDate() + i);
    days.push({ date, key: dateKey(date), inCurrentMonth: date.getMonth() === monthDate.getMonth() });
  }
  return days;
}

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
}

function uid(prefix) {
  return `${prefix}_${Math.random().toString(36).slice(2,8)}${Date.now().toString(36).slice(-4)}`;
}

function upsert(array, item) {
  const index = array.findIndex(x => x.id === item.id);
  if (index === -1) array.push(item);
  else array[index] = item;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, char => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[char]));
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function createDemoData() {
  const people = [
    { id: 'p1', name: 'Marc Dubois', type: 'pigiste', role: 'Technicien son', email: 'marc@example.com', phone: '514-555-0101', contactPreference: 'both', notes: '' },
    { id: 'p2', name: 'Julie Tremblay', type: 'pigiste', role: 'Technicien éclairage', email: 'julie@example.com', phone: '514-555-0102', contactPreference: 'sms', notes: '' },
    { id: 'p3', name: 'Alex Caron', type: 'employe', role: 'Technicien son', email: 'alex@example.com', phone: '514-555-0103', contactPreference: 'email', notes: '' },
    { id: 'p4', name: 'Samuel Gagnon', type: 'pigiste', role: 'Vidéo', email: 'samuel@example.com', phone: '514-555-0104', contactPreference: 'both', notes: '' },
    { id: 'p5', name: 'Kevin Roy', type: 'pigiste', role: 'Technicien son', email: 'kevin@example.com', phone: '514-555-0105', contactPreference: 'sms', notes: '' },
  ];

  const now = new Date();
  const event1Start = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2, 8, 0);
  const event1End = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2, 17, 0);
  const event2Start = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 5, 12, 0);
  const event2End = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 5, 23, 0);

  const events = [
    {
      id: 'e1',
      title: 'Festival du centre-ville',
      location: 'Place des Arts',
      start: toDateTimeLocal(event1Start),
      end: toDateTimeLocal(event1End),
      description: 'Montage et opération audio/éclairage.',
      importedFrom: 'demo',
      requirements: [
        { id: 'r1', role: 'Technicien son', count: 2, notes: 'Montage + show' },
        { id: 'r2', role: 'Technicien éclairage', count: 1, notes: '' },
      ],
      assignments: [
        { id: 'a1', personId: 'p1', role: 'Technicien son', status: 'sent', invitedAt: toDateTimeLocal(new Date()), respondedAt: null },
        { id: 'a2', personId: 'p3', role: 'Technicien son', status: 'confirmed', invitedAt: toDateTimeLocal(new Date()), respondedAt: toDateTimeLocal(new Date()) },
        { id: 'a3', personId: 'p2', role: 'Technicien éclairage', status: 'pending', invitedAt: null, respondedAt: null },
      ],
    },
    {
      id: 'e2',
      title: 'Corporate Gala',
      location: 'Hôtel Bonaventure',
      start: toDateTimeLocal(event2Start),
      end: toDateTimeLocal(event2End),
      description: 'Soirée corporative avec vidéo et éclairage.',
      importedFrom: 'demo',
      requirements: [
        { id: 'r3', role: 'Vidéo', count: 1, notes: 'Projection + cues' },
        { id: 'r4', role: 'Technicien son', count: 1, notes: '' },
      ],
      assignments: [
        { id: 'a4', personId: 'p4', role: 'Vidéo', status: 'declined', invitedAt: toDateTimeLocal(new Date()), respondedAt: toDateTimeLocal(new Date()) },
      ],
    },
  ];

  return {
    people,
    events,
    notificationLog: [],
    config: {
      outlook: { clientId: '', tenantId: 'common', redirectUri: '' },
      webhooks: { email: '', sms: '' },
    },
  };
}
