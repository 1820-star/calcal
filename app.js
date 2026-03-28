const STORAGE_KEY = 'calcal-appointments-v1';
const NGB_STORAGE_KEY = 'calcal-ngb-v1';
const SLOT_TIMES = buildSlotTimes();

const today = new Date();
const state = {
  appointments: loadAppointments(),
  ngbs: loadNgbs(),
  currentMonthStart: startOfMonth(today),
  selectedDate: '',
  editingKey: null,
  editingNgbId: null,
  moveAppointmentId: null,
  archiveQuery: '',
  ngbQuery: '',
  deferredInstallPrompt: null,
};

const hoverCard = createHoverCard();
const hoverState = {
  x: 0,
  y: 0,
  targetX: 0,
  targetY: 0,
  rafId: null,
  showTimer: null,
  pointerX: 0,
  pointerY: 0,
};

const elements = {
  navTabs: [...document.querySelectorAll('.nav-tab')],
  views: {
    calendar: document.getElementById('calendar-view'),
    travel: document.getElementById('travel-view'),
    ngb: document.getElementById('ngb-view'),
    archive: document.getElementById('archive-view'),
  },
  monthTitle: document.getElementById('month-title'),
  prevMonth: document.getElementById('prev-month'),
  nextMonth: document.getElementById('next-month'),
  todayLabel: document.getElementById('today-label'),
  viewTitle: document.getElementById('view-title'),
  selectedDateInput: document.getElementById('selected-date'),
  selectedDateLabel: document.getElementById('selected-date-label'),
  selectedDayCount: document.getElementById('selected-day-count'),
  newAppointment: document.getElementById('new-appointment'),
  printDaylist: document.getElementById('print-daylist'),
  moveHint: document.getElementById('move-hint'),
  nextThree: document.getElementById('next-three'),
  calendarGrid: document.getElementById('calendar-grid'),
  archiveContainer: document.getElementById('archive-container'),
  archiveSearch: document.getElementById('archive-search'),
  archiveSearchInfo: document.getElementById('archive-search-info'),
  exportArchiveCsv: document.getElementById('export-archive-csv'),
  ngbContainer: document.getElementById('ngb-container'),
  ngbSearch: document.getElementById('ngb-search'),
  ngbSearchInfo: document.getElementById('ngb-search-info'),
  newNgb: document.getElementById('new-ngb'),
  travelPrice: document.getElementById('travel-price'),
  travelStart: document.getElementById('travel-start'),
  travelEnd: document.getElementById('travel-end'),
  travelDuration: document.getElementById('travel-duration'),
  travelCalc: document.getElementById('travel-calc'),
  travelResults: document.getElementById('travel-results'),
  installApp: document.getElementById('install-app'),
  exportMonthOverview: document.getElementById('export-month-overview'),
  exportBothWeeks: document.getElementById('export-both-weeks'),
  exportData: document.getElementById('export-data'),
  importData: document.getElementById('import-data'),
  importDataFile: document.getElementById('import-data-file'),
  dialog: document.getElementById('appointment-dialog'),
  dialogTitle: document.getElementById('dialog-slot-title'),
  appointmentForm: document.getElementById('appointment-form'),
  closeDialog: document.getElementById('close-dialog'),
  cancelDialog: document.getElementById('cancel-dialog'),
  deleteAppointment: document.getElementById('delete-appointment'),
  moveAppointment: document.getElementById('move-appointment'),
  ngbDialog: document.getElementById('ngb-dialog'),
  ngbForm: document.getElementById('ngb-form'),
  ngbDialogTitle: document.getElementById('dialog-ngb-title'),
  closeNgbDialog: document.getElementById('close-ngb-dialog'),
  cancelNgbDialog: document.getElementById('cancel-ngb-dialog'),
  deleteNgb: document.getElementById('delete-ngb'),
};

bootstrap();

function bootstrap() {
  archivePastAppointments();
  state.selectedDate = getInitialSelectedDate();
  registerServiceWorker();
  bindEvents();
  renderAll();
}

function bindEvents() {
  elements.navTabs.forEach((button) => {
    button.addEventListener('click', () => setView(button.dataset.view));
  });

  elements.prevMonth.addEventListener('click', () => {
    state.currentMonthStart = addMonths(state.currentMonthStart, -1);
    ensureValidSelectedDate();
    elements.calendarGrid.classList.add('animate-month-change');
    setTimeout(() => elements.calendarGrid.classList.remove('animate-month-change'), 200);
    renderAll();
  });

  elements.nextMonth.addEventListener('click', () => {
    state.currentMonthStart = addMonths(state.currentMonthStart, 1);
    ensureValidSelectedDate();
    elements.calendarGrid.classList.add('animate-month-change');
    setTimeout(() => elements.calendarGrid.classList.remove('animate-month-change'), 200);
    renderAll();
  });

  elements.selectedDateInput.addEventListener('change', handleSelectedDateChange);
  elements.newAppointment.addEventListener('click', openNewAppointmentDialog);
  elements.printDaylist.addEventListener('click', printSelectedDayTwice);
  elements.archiveSearch.addEventListener('input', handleArchiveSearchInput);
  elements.exportArchiveCsv.addEventListener('click', exportArchiveCsv);
  elements.installApp.addEventListener('click', installApp);
  elements.exportMonthOverview.addEventListener('click', exportMonthlyOverview);
  elements.exportBothWeeks.addEventListener('click', exportBothWeeksOverview);
  elements.exportData.addEventListener('click', exportDataJson);
  elements.importData.addEventListener('click', () => elements.importDataFile.click());
  elements.importDataFile.addEventListener('change', importDataJson);

  elements.appointmentForm.addEventListener('submit', handleSubmit);
  elements.closeDialog.addEventListener('click', closeDialog);
  elements.cancelDialog.addEventListener('click', closeDialog);
  elements.deleteAppointment.addEventListener('click', handleDelete);
  elements.moveAppointment.addEventListener('click', handleMoveStart);
  
  if (elements.newNgb && elements.ngbSearch && elements.ngbForm && elements.closeNgbDialog && elements.cancelNgbDialog && elements.deleteNgb) {
    elements.newNgb.addEventListener('click', openNgbDialog);
    elements.ngbSearch.addEventListener('input', handleNgbSearchInput);
    elements.ngbForm.addEventListener('submit', handleNgbSubmit);
    elements.closeNgbDialog.addEventListener('click', closeNgbDialog);
    elements.cancelNgbDialog.addEventListener('click', closeNgbDialog);
    elements.deleteNgb.addEventListener('click', handleNgbDelete);
  }

  if (elements.travelCalc && elements.travelPrice && elements.travelStart && elements.travelEnd && elements.travelDuration) {
    elements.travelCalc.addEventListener('click', calculateTravelCosts);
    elements.travelPrice.addEventListener('change', calculateTravelCosts);
    elements.travelStart.addEventListener('change', calculateTravelCosts);
    elements.travelEnd.addEventListener('change', calculateTravelCosts);
    elements.travelDuration.addEventListener('change', calculateTravelCosts);
  }

  window.addEventListener('resize', fitCalendarToViewport);

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    state.deferredInstallPrompt = event;
    elements.installApp.classList.remove('hidden');
  });

  window.addEventListener('appinstalled', () => {
    state.deferredInstallPrompt = null;
    elements.installApp.classList.add('hidden');
  });
}

function setView(view) {
  Object.entries(elements.views).forEach(([key, node]) => {
    if (!node) {
      return;
    }
    node.classList.toggle('active', key === view);
  });

  elements.navTabs.forEach((button) => {
    button.classList.toggle('active', button.dataset.view === view);
  });

  const titles = {
    calendar: 'Kalender',
    travel: 'Fahrtkosten',
    ngb: 'NGB',
    archive: 'Archiv',
  };
  if (titles[view]) {
    elements.viewTitle.textContent = titles[view];
  }
}

function renderAll() {
  archivePastAppointments();
  ensureValidSelectedDate();
  updateHeaderLabels();
  renderCalendar();
  renderNextThree();
  renderArchive();
  renderNgb();
  updateSelectionSummary();
  updateMoveHint();
  elements.selectedDateInput.value = state.selectedDate;
}

function updateHeaderLabels() {
  elements.monthTitle.textContent = new Intl.DateTimeFormat('de-DE', {
    month: 'long',
    year: 'numeric',
  }).format(state.currentMonthStart);

  elements.todayLabel.textContent = `Heute: ${formatDateLong(formatDateISO(new Date()))}`;
}

function renderCalendar() {
  const monthDays = buildWorkDaysForMonth(state.currentMonthStart);
  const visibleSlotTimes = getVisibleSlotTimes(monthDays);
  const todayIso = formatDateISO(new Date());
  const table = document.createElement('table');
  table.className = 'calendar-table';

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.appendChild(createCell('th', 'Zeit'));

  monthDays.forEach((date) => {
    const th = document.createElement('th');
    th.className = `day-header ${getDayStatusClass(date)}`;
    if (date === todayIso) {
      th.classList.add('today');
    }
    th.innerHTML = `
      <div class="day-weekday">${formatWeekdayCompact(date)}</div>
      <div class="day-number">${formatDayNumber(date)}</div>
      <small>${countAppointmentsForDay(date)}</small>
    `;
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  visibleSlotTimes.forEach((time) => {
    const row = document.createElement('tr');
    row.appendChild(createCell('td', time));

    monthDays.forEach((date) => {
      const td = document.createElement('td');
      td.className = 'slot-cell';
      if (date === todayIso) {
        td.classList.add('today-column');
      }
      td.dataset.date = date;
      td.dataset.time = time;
      td.addEventListener('click', () => handleSlotClick(date, time));

      const appointment = getAppointment(date, time);
      if (appointment) {
        td.classList.add(getSlotBookingClass(date, time));
        td.dataset.tooltip = [
          `${formatDateLong(date)} · ${time}`,
          `Name: ${appointment.name}`,
          `Telefon: ${appointment.phone}`,
          `KG: ${appointment.kg || '-'}`,
          `EZ: ${appointment.ez || '-'}`,
          `Anliegen: ${appointment.concern || '-'}`,
        ].join('\n');
      } else {
        td.classList.add('slot-empty-cell');
        td.dataset.tooltip = [
          `${formatDateLong(date)} · ${time}`,
          'Status: frei',
          'Klick: Termin anlegen',
        ].join('\n');
      }

      td.addEventListener('mouseenter', handleSlotHoverStart);
      td.addEventListener('mousemove', handleSlotHoverMove);
      td.addEventListener('mouseleave', handleSlotHoverEnd);
      td.innerHTML = '<span class="slot-plus">+</span>';

      row.appendChild(td);
    });

    tbody.appendChild(row);
  });

  table.appendChild(tbody);
  elements.calendarGrid.replaceChildren(table);
  fitCalendarToViewport();
}

function handleSlotClick(date, time) {
  state.selectedDate = date;

  if (state.moveAppointmentId) {
    moveAppointmentToSlot(date, time);
    return;
  }

  openDialog(date, time);
}

function handleSelectedDateChange(event) {
  const picked = event.target.value;
  if (!picked) {
    return;
  }

  const pickedDate = new Date(`${picked}T00:00:00`);
  if (!isWeekDay(pickedDate)) {
    alert('Bitte einen Werktag (Mo-Fr) auswählen.');
    event.target.value = state.selectedDate;
    return;
  }

  state.selectedDate = picked;
  state.currentMonthStart = startOfMonth(pickedDate);
  renderAll();
}

function openNewAppointmentDialog() {
  const defaultDate = state.selectedDate || formatDateISO(new Date());
  const defaultTime = findFirstFreeSlotForDate(defaultDate) || SLOT_TIMES[0];
  openDialog(defaultDate, defaultTime, true);
}

function moveAppointmentToSlot(date, time) {
  const movingAppointment = state.appointments.find(
    (appointment) => appointment.status === 'open' && appointment.id === state.moveAppointmentId
  );

  if (!movingAppointment) {
    state.moveAppointmentId = null;
    renderAll();
    return;
  }

  const targetAppointment = getAppointment(date, time);
  if (targetAppointment) {
    alert('Der Zielslot ist bereits belegt. Bitte einen freien Slot wählen.');
    return;
  }

  movingAppointment.date = date;
  movingAppointment.time = time;
  movingAppointment.updatedAt = new Date().toISOString();
  state.moveAppointmentId = null;
  persistAppointments();
  renderAll();
}

function renderArchive() {
  const archiveItems = state.appointments
    .slice()
    .sort((left, right) => `${right.date} ${right.time}`.localeCompare(`${left.date} ${left.time}`));

  const query = state.archiveQuery.trim().toLowerCase();
  const filteredItems = query
    ? archiveItems.filter((appointment) => {
        const haystack = [
          appointment.name,
          appointment.phone,
          appointment.kg,
          appointment.ez,
          appointment.concern,
          appointment.status,
          appointment.time,
          appointment.date,
          formatDateGerman(appointment.date),
          appointment.createdAt ? formatDateTimeGerman(appointment.createdAt) : '',
        ]
          .join(' ')
          .toLowerCase();
        return haystack.includes(query);
      })
    : archiveItems;

  const headers = ['Terminzeit', 'Name', 'Telefon', 'KG', 'EZ', 'Anliegen', 'Status', 'Eingetragen am'];
  const rows = filteredItems.map((appointment) => [
    `${formatDateGerman(appointment.date)} ${appointment.time}`,
    appointment.name,
    appointment.phone,
    appointment.kg || '-',
    appointment.ez || '-',
    appointment.concern || '-',
    getStatusLabel(appointment.status),
    appointment.createdAt ? formatDateTimeGerman(appointment.createdAt) : '-',
  ]);

  elements.archiveSearchInfo.textContent = `${filteredItems.length} Treffer`;
  elements.archiveContainer.replaceChildren(buildTable(headers, rows, 'Keine Eintragungen zur Suche gefunden.'));
}

function handleArchiveSearchInput(event) {
  state.archiveQuery = event.target.value || '';
  renderArchive();
}

function exportArchiveCsv() {
  const rows = state.appointments
    .slice()
    .sort((left, right) => `${left.date} ${left.time}`.localeCompare(`${right.date} ${right.time}`));

  if (!rows.length) {
    alert('Keine Daten fuer den CSV-Export vorhanden.');
    return;
  }

  const headers = [
    'Datum',
    'Uhrzeit',
    'Name',
    'Telefon',
    'KG',
    'EZ',
    'Anliegen',
    'Status',
    'EingetragenAm',
  ];

  const csvRows = [
    headers.join(';'),
    ...rows.map((item) =>
      [
        formatDateGerman(item.date),
        item.time,
        item.name,
        item.phone,
        item.kg || '',
        item.ez || '',
        item.concern || '',
        getStatusLabel(item.status),
        item.createdAt ? formatDateTimeGerman(item.createdAt) : '',
      ]
        .map(escapeCsvValue)
        .join(';')
    ),
  ];

  const csvContent = `\ufeff${csvRows.join('\n')}`;
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const stamp = formatDateISO(new Date());
  link.href = url;
  link.download = `scal-kontakte-${stamp}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function escapeCsvValue(value) {
  const text = String(value ?? '');
  return `"${text.replaceAll('"', '""')}"`;
}

function renderNextThree() {
  const now = new Date();
  const upcoming = state.appointments
    .filter((appointment) => {
      if (appointment.status !== 'open') {
        return false;
      }

      const dateTime = new Date(`${appointment.date}T${appointment.time}:00`);
      return dateTime >= now;
    })
    .sort((left, right) => `${left.date} ${left.time}`.localeCompare(`${right.date} ${right.time}`))
    .slice(0, 3);

  const slots = [0, 1, 2].map((index) => {
    const appointment = upcoming[index];
    if (!appointment) {
      return '<div class="next-three-item empty">Frei</div>';
    }

    return `<div class="next-three-item">${escapeHtml(formatDateGerman(appointment.date))} ${escapeHtml(
      appointment.time
    )} · ${escapeHtml(appointment.name)} · KG ${escapeHtml(appointment.kg || '-')} · EZ ${escapeHtml(
      appointment.ez || '-'
    )} · ${escapeHtml(appointment.concern || '-')}</div>`;
  });

  elements.nextThree.innerHTML = slots.join('');
}

function updateSelectionSummary() {
  elements.selectedDateLabel.textContent = formatDateLong(state.selectedDate);
  elements.selectedDayCount.textContent = `${countAppointmentsForDay(state.selectedDate)} Termine`;
}

function updateMoveHint() {
  if (state.moveAppointmentId) {
    elements.moveHint.textContent = 'Verschiebemodus aktiv: freien Zielslot anklicken.';
    return;
  }
  elements.moveHint.textContent = 'Mo-Fr, 08:00 bis 12:00. Klick auf Termin: bearbeiten oder verschieben.';
}

function openDialog(date, time, forceNew = false) {
  const appointment = getAppointment(date, time);
  const isEditingExisting = Boolean(appointment) && !forceNew;
  state.editingKey = isEditingExisting ? `${date}__${time}` : null;
  elements.dialogTitle.textContent = `${formatDateLong(date)} · ${time}`;
  elements.deleteAppointment.classList.toggle('hidden', !isEditingExisting);
  elements.moveAppointment.classList.toggle('hidden', !isEditingExisting);

  fillTimeSelect();
  elements.appointmentForm.date.value = date;
  elements.appointmentForm.time.value = time;

  elements.appointmentForm.name.value = isEditingExisting ? appointment.name : '';
  elements.appointmentForm.kg.value = isEditingExisting ? appointment.kg || '' : '';
  elements.appointmentForm.ez.value = isEditingExisting ? appointment.ez || '' : '';
  elements.appointmentForm.phone.value = isEditingExisting ? appointment.phone : '';
  elements.appointmentForm.concern.value = isEditingExisting ? appointment.concern || '' : '';

  elements.dialog.showModal();
}

function closeDialog() {
  elements.dialog.close();
  state.editingKey = null;
  elements.appointmentForm.reset();
}

function handleMoveStart() {
  if (!state.editingKey) {
    return;
  }

  const [date, time] = state.editingKey.split('__');
  const appointment = getAppointment(date, time);
  if (!appointment) {
    return;
  }

  state.moveAppointmentId = appointment.id;
  closeDialog();
  setView('calendar');
  updateMoveHint();
}

function handleSubmit(event) {
  event.preventDefault();

  const targetDate = elements.appointmentForm.date.value;
  const targetTime = elements.appointmentForm.time.value;
  const editingParts = state.editingKey ? state.editingKey.split('__') : null;
  const sourceDate = editingParts ? editingParts[0] : null;
  const sourceTime = editingParts ? editingParts[1] : null;
  const existingAppointment = sourceDate && sourceTime ? getAppointment(sourceDate, sourceTime) : null;

  const conflictingAppointment = getAppointment(targetDate, targetTime);
  if (
    conflictingAppointment &&
    (!existingAppointment || conflictingAppointment.id !== existingAppointment.id)
  ) {
    alert('Dieser Termin-Slot ist bereits belegt. Bitte andere Uhrzeit wählen.');
    return;
  }

  const payload = {
    id: existingAppointment?.id || crypto.randomUUID(),
    date: targetDate,
    time: targetTime,
    name: elements.appointmentForm.name.value.trim(),
    kg: elements.appointmentForm.kg.value.trim(),
    ez: elements.appointmentForm.ez.value.trim(),
    phone: elements.appointmentForm.phone.value.trim(),
    concern: elements.appointmentForm.concern.value.trim(),
    status: 'open',
    archivedAt: null,
    createdAt: existingAppointment?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  if (!payload.name || !payload.phone) {
    alert('Name und Telefonnummer sind Pflichtfelder.');
    return;
  }

  if (existingAppointment) {
    state.appointments = state.appointments.filter((appointment) => appointment.id !== existingAppointment.id);
  }
  state.appointments.push(payload);
  state.selectedDate = targetDate;
  state.currentMonthStart = startOfMonth(new Date(`${targetDate}T00:00:00`));
  persistAppointments();
  renderAll();
  closeDialog();
}

function handleDelete() {
  if (!state.editingKey) {
    return;
  }

  const [date, time] = state.editingKey.split('__');
  state.appointments = state.appointments.filter(
    (appointment) => !(appointment.date === date && appointment.time === time && appointment.status === 'open')
  );
  persistAppointments();
  renderAll();
  closeDialog();
}

function archivePastAppointments() {
  const now = new Date();
  const twoHoursAgo = new Date(now.getTime() - 2 * 60 * 60 * 1000);
  let changed = false;

  state.appointments = state.appointments.map((appointment) => {
    if (appointment.status === 'open') {
      const appointmentDate = new Date(`${appointment.date}T${appointment.time}:00`);
      if (appointmentDate < twoHoursAgo) {
        changed = true;
        return {
          ...appointment,
          status: 'done',
          archivedAt: new Date().toISOString(),
        };
      }
    }
    return appointment;
  });

  if (changed) {
    persistAppointments();
  }
}

function getAppointmentsForDay(date) {
  return state.appointments.filter((appointment) => appointment.status === 'open' && appointment.date === date);
}

function countAppointmentsForDay(date) {
  return getAppointmentsForDay(date).length;
}

function getDayStatusClass(date) {
  const count = countAppointmentsForDay(date);
  if (count <= 1) {
    return 'green';
  }
  if (count <= 4) {
    return 'yellow';
  }
  return 'red';
}

function getAppointment(date, time) {
  return state.appointments.find(
    (appointment) => appointment.status === 'open' && appointment.date === date && appointment.time === time
  );
}

function getSlotBookingClass(date, time) {
  const slotStart = new Date(`${date}T${time}:00`);
  const slotEnd = new Date(slotStart.getTime() + 10 * 60 * 1000);
  const now = new Date();

  if (now < slotStart) {
    return 'slot-booked-future';
  }
  if (now >= slotStart && now <= slotEnd) {
    return 'slot-booked-now';
  }
  return 'slot-booked-past';
}

function getStatusLabel(status) {
  if (status === 'open') {
    return 'Aktiv';
  }
  if (status === 'done' || status === 'archived') {
    return 'Erledigt';
  }
  return String(status || '-');
}

function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) {
    return;
  }

  window.addEventListener('load', () => {
    void (async () => {
      try {
        const host = window.location.hostname;
        const isLocalHost = host === 'localhost' || host === '127.0.0.1';

        if (isLocalHost) {
          const registrations = await navigator.serviceWorker.getRegistrations();
          await Promise.all(registrations.map((registration) => registration.unregister()));

          if ('caches' in window) {
            const keys = await caches.keys();
            await Promise.all(keys.map((key) => caches.delete(key)));
          }
        }

        await navigator.serviceWorker.register('./sw.js?v=scal-milestone-20260328');
      } catch {
        // Leise fehlschlagen, App soll ohne SW weiterhin funktionieren.
      }
    })();
  });
}

async function installApp() {
  if (!state.deferredInstallPrompt) {
    return;
  }

  state.deferredInstallPrompt.prompt();
  await state.deferredInstallPrompt.userChoice;
  state.deferredInstallPrompt = null;
  elements.installApp.classList.add('hidden');
}

function exportDataJson() {
  const payload = {
    exportedAt: new Date().toISOString(),
    app: 'CalCal',
    version: 1,
    appointments: state.appointments,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `calcal-backup-${formatDateISO(new Date())}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function exportMonthlyOverview() {
  const html = buildMonthlyOverviewHtml();

  if (window.calcalDesktop?.isDesktopApp && typeof window.calcalDesktop.saveMonthlyOverview === 'function') {
    try {
      const result = await window.calcalDesktop.saveMonthlyOverview(html);
      if (result?.ok) {
        alert(`Monatsuebersicht gespeichert: ${result.filePath}`);
        return;
      }
    } catch {
      alert('Speichern im Zielordner fehlgeschlagen. Fallback-Download wird gestartet.');
    }
  }

  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'Monatsuebersicht-Aktuell.html';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildMonthlyOverviewHtml() {
  const monthDays = buildWorkDaysForMonth(state.currentMonthStart);
  const visibleSlotTimes = getVisibleSlotTimes(monthDays);
  const monthLabel = new Intl.DateTimeFormat('de-DE', { month: 'long', year: 'numeric' }).format(state.currentMonthStart);

  const headerCells = monthDays
    .map((date) => `<th>${escapeHtml(formatWeekdayCompact(date))} ${escapeHtml(formatDayNumber(date))}</th>`)
    .join('');

  const bodyRows = visibleSlotTimes
    .map((time) => {
      const dayCells = monthDays
        .map((date) => {
          const appointment = getAppointment(date, time);
          if (!appointment) {
            return '<td></td>';
          }

          const lines = [
            `${appointment.time}`,
            appointment.name,
            appointment.phone || '',
            appointment.concern || '',
          ].filter(Boolean);

          return `<td>${lines.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</td>`;
        })
        .join('');

      return `<tr><th>${escapeHtml(time)}</th>${dayCells}</tr>`;
    })
    .join('');

  return `<!doctype html>
<html lang="de">
  <head>
    <meta charset="UTF-8" />
    <title>Monatsuebersicht ${escapeHtml(monthLabel)}</title>
    <style>
      @page { size: A4 landscape; margin: 8mm; }
      * { box-sizing: border-box; }
      body { margin: 0; font-family: Arial, sans-serif; color: #232323; }
      h1 { margin: 0 0 6mm; font-size: 18px; }
      table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10px; }
      th, td { border: 1px solid #cfcfcf; padding: 2px; vertical-align: top; }
      thead th { background: #f2f2f2; }
      tbody th { background: #f7f7f7; width: 44px; }
      td div { white-space: normal; line-height: 1.2; }
    </style>
  </head>
  <body>
    <h1>scal Monatsuebersicht ${escapeHtml(monthLabel)}</h1>
    <table>
      <thead>
        <tr><th>Zeit</th>${headerCells}</tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
    </table>
  </body>
</html>`;
}

async function exportBothWeeksOverview() {
  const today = new Date();
  
  // Calculate last week start
  const dayOfWeek = today.getDay();
  const daysToMonday = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
  const currentWeekStart = new Date(today);
  currentWeekStart.setDate(currentWeekStart.getDate() + daysToMonday);
  
  const lastWeekStart = new Date(currentWeekStart);
  lastWeekStart.setDate(lastWeekStart.getDate() - 7);
  
  const html = buildBothWeeksHtml(formatDateISO(lastWeekStart), formatDateISO(currentWeekStart));

  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const stamp = formatDateISO(new Date());
  link.href = url;
  link.download = `scal-wochen-${stamp}.html`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildBothWeeksHtml(lastWeekStartIso, currentWeekStartIso) {
  const lastWeekStart = new Date(`${lastWeekStartIso}T00:00:00`);
  const currentWeekStart = new Date(`${currentWeekStartIso}T00:00:00`);
  
  const getWeekDays = (weekStart) => {
    const days = [];
    for (let i = 0; i < 5; i++) {
      const date = new Date(weekStart);
      date.setDate(date.getDate() + i);
      const iso = formatDateISO(date);
      if (isWeekDay(date)) {
        days.push(iso);
      }
    }
    return days;
  };

  const lastWeekDays = getWeekDays(lastWeekStart);
  const currentWeekDays = getWeekDays(currentWeekStart);
  const visibleSlotTimes = getVisibleSlotTimes([...lastWeekDays, ...currentWeekDays]);

  const lastWeekLabel = `${formatDateGerman(lastWeekDays[0])} - ${formatDateGerman(lastWeekDays[lastWeekDays.length - 1])}`;
  const currentWeekLabel = `${formatDateGerman(currentWeekDays[0])} - ${formatDateGerman(currentWeekDays[currentWeekDays.length - 1])}`;

  const buildWeekTable = (weekDays, weekTitle) => {
    const headerCells = weekDays
      .map((date) => `<th>${escapeHtml(formatWeekdayCompact(date))} ${escapeHtml(formatDayNumber(date))}</th>`)
      .join('');

    const bodyRows = visibleSlotTimes
      .map((time) => {
        const dayCells = weekDays
          .map((date) => {
            const appointment = getAppointment(date, time);
            if (!appointment) {
              return '<td></td>';
            }

            const lines = [
              `${appointment.time}`,
              appointment.name,
              appointment.phone || '',
              appointment.concern || '',
            ].filter(Boolean);

            return `<td>${lines.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</td>`;
          })
          .join('');

        return `<tr><th>${escapeHtml(time)}</th>${dayCells}</tr>`;
      })
      .join('');

    return `
        <div style="margin-bottom: 24px;">
          <p style="margin: 12px 0; font-size: 11px; color: #666;">${escapeHtml(weekTitle)}</p>
          <table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10px;">
        <thead>
          <tr><th>Zeit</th>${headerCells}</tr>
        </thead>
        <tbody>
          ${bodyRows}
        </tbody>
      </table>
        </div>
    `;
  };

  return `<!doctype html>
<html lang="de">
  <head>
    <meta charset="UTF-8" />
      <title>Terminuebersicht</title>
    <style>
      @page { size: A4 landscape; margin: 8mm; }
      * { box-sizing: border-box; }
      body { margin: 0; font-family: Arial, sans-serif; color: #232323; }
        h1 { margin: 0 0 12mm; font-size: 18px; }
      table { border-collapse: collapse; table-layout: fixed; font-size: 10px; }
      th, td { border: 1px solid #cfcfcf; padding: 2px; vertical-align: top; }
      thead th { background: #f2f2f2; }
      tbody th { background: #f7f7f7; width: 44px; }
      td div { white-space: normal; line-height: 1.2; }
    </style>
  </head>
  <body>
      <h1>scal Terminübersicht</h1>
      ${buildWeekTable(lastWeekDays, lastWeekLabel)}
      ${buildWeekTable(currentWeekDays, currentWeekLabel)}
  </body>
</html>`;
}

async function exportWeeklyOverview(weekType) {
  const today = new Date();
  let weekStart;

  if (weekType === 'current') {
    const dayOfWeek = today.getDay();
    const daysToMonday = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    weekStart = new Date(today);
    weekStart.setDate(weekStart.getDate() + daysToMonday);
  } else {
    const dayOfWeek = today.getDay();
    const daysToMonday = (dayOfWeek === 0 ? -6 : 1) - dayOfWeek;
    weekStart = new Date(today);
    weekStart.setDate(weekStart.getDate() + daysToMonday - 7);
  }

  const html = buildWeeklyOverviewHtml(formatDateISO(weekStart));

  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const weekLabel = weekType === 'current' ? 'Aktuelle' : 'Letzte';
  const stamp = formatDateISO(weekStart);
  link.href = url;
  link.download = `scal-woche-${weekLabel.toLowerCase()}-${stamp}.html`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildWeeklyOverviewHtml(weekStartIso) {
  const weekStart = new Date(`${weekStartIso}T00:00:00`);
  const weekDays = [];
  for (let i = 0; i < 5; i++) {
    const date = new Date(weekStart);
    date.setDate(date.getDate() + i);
    const iso = formatDateISO(date);
    if (isWeekDay(date)) {
      weekDays.push(iso);
    }
  }

  const visibleSlotTimes = getVisibleSlotTimes(weekDays);
  const weekLabel = `${formatDateGerman(weekDays[0])} - ${formatDateGerman(weekDays[weekDays.length - 1])}`;

  const headerCells = weekDays
    .map((date) => `<th>${escapeHtml(formatWeekdayCompact(date))} ${escapeHtml(formatDayNumber(date))}</th>`)
    .join('');

  const bodyRows = visibleSlotTimes
    .map((time) => {
      const dayCells = weekDays
        .map((date) => {
          const appointment = getAppointment(date, time);
          if (!appointment) {
            return '<td></td>';
          }

          const lines = [
            `${appointment.time}`,
            appointment.name,
            appointment.phone || '',
            appointment.concern || '',
          ].filter(Boolean);

          return `<td>${lines.map((line) => `<div>${escapeHtml(line)}</div>`).join('')}</td>`;
        })
        .join('');

      return `<tr><th>${escapeHtml(time)}</th>${dayCells}</tr>`;
    })
    .join('');

  return `<!doctype html>
<html lang="de">
  <head>
    <meta charset="UTF-8" />
    <title>Wochenuebersicht ${escapeHtml(weekLabel)}</title>
    <style>
      @page { size: A4 landscape; margin: 8mm; }
      * { box-sizing: border-box; }
      body { margin: 0; font-family: Arial, sans-serif; color: #232323; }
      h1 { margin: 0 0 6mm; font-size: 18px; }
      table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 10px; }
      th, td { border: 1px solid #cfcfcf; padding: 2px; vertical-align: top; }
      thead th { background: #f2f2f2; }
      tbody th { background: #f7f7f7; width: 44px; }
      td div { white-space: normal; line-height: 1.2; }
    </style>
  </head>
  <body>
    <h1>scal Wochenuebersicht ${escapeHtml(weekLabel)}</h1>
    <table>
      <thead>
        <tr><th>Zeit</th>${headerCells}</tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
    </table>
  </body>
</html>`;
}

async function importDataJson(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const appointments = Array.isArray(parsed) ? parsed : parsed.appointments;

    if (!Array.isArray(appointments)) {
      throw new Error('Ungueltiges Format');
    }

    state.appointments = appointments.filter((item) => item && item.date && item.time && item.name);
    persistAppointments();
    renderAll();
    alert('Daten wurden erfolgreich importiert.');
  } catch {
    alert('Import fehlgeschlagen. Bitte eine gueltige CalCal-JSON Datei waehlen.');
  } finally {
    event.target.value = '';
  }
}

function getVisibleSlotTimes(monthDays) {
  const filledTimes = new Set(
    state.appointments
      .filter((appointment) => appointment.status === 'open' && monthDays.includes(appointment.date))
      .map((appointment) => appointment.time)
  );

  return SLOT_TIMES.filter((time) => isPrimarySlotTime(time) || filledTimes.has(time));
}

function isPrimarySlotTime(time) {
  return time.endsWith(':00') || time.endsWith(':30');
}

function fitCalendarToViewport() {
  const table = elements.calendarGrid.querySelector('.calendar-table');
  if (!table) {
    return;
  }

  table.style.transform = 'none';
  table.style.width = '100%';
  table.style.height = 'auto';
}

function fillTimeSelect() {
  const select = elements.appointmentForm.time;
  select.innerHTML = '';
  SLOT_TIMES.forEach((slotTime) => {
    const option = document.createElement('option');
    option.value = slotTime;
    option.textContent = slotTime;
    select.appendChild(option);
  });
}

function findFirstFreeSlotForDate(date) {
  return SLOT_TIMES.find((time) => !getAppointment(date, time)) || null;
}

function printSelectedDayTwice() {
  const printDate = elements.selectedDateInput.value || state.selectedDate;
  const appointments = getAppointmentsForDay(printDate)
    .sort((left, right) => left.time.localeCompare(right.time))
    .map((appointment) => `${appointment.time} - ${escapeHtml(appointment.name)}`);

  const rowsHtml = appointments.length
    ? appointments.map((line) => `<tr><td>${line}</td></tr>`).join('')
    : '<tr><td>Heute keine Termine</td></tr>';

  const title = `${escapeHtml(formatDateLong(printDate))}`;

  const existingRoot = document.getElementById('print-root');
  if (existingRoot) {
    existingRoot.remove();
  }

  const printRoot = document.createElement('div');
  printRoot.id = 'print-root';
  printRoot.className = 'print-root';
  printRoot.innerHTML = `
    <div class="print-page">
      <section class="print-copy"><h1>${title}</h1><table>${rowsHtml}</table></section>
      <section class="print-copy"><h1>${title}</h1><table>${rowsHtml}</table></section>
    </div>
  `;

  document.body.appendChild(printRoot);
  document.body.classList.add('print-daylist-mode');

  const cleanup = () => {
    document.body.classList.remove('print-daylist-mode');
    printRoot.remove();
    window.removeEventListener('afterprint', cleanup);
  };

  window.addEventListener('afterprint', cleanup);
  window.setTimeout(() => {
    window.print();
  }, 0);
}

function buildTable(headers, rows, emptyText) {
  const template = document.getElementById('table-template');
  const fragment = template.content.firstElementChild.cloneNode(true);
  const thead = fragment.querySelector('thead');
  const tbody = fragment.querySelector('tbody');

  const headerRow = document.createElement('tr');
  headers.forEach((header) => headerRow.appendChild(createCell('th', header)));
  thead.appendChild(headerRow);

  if (!rows.length) {
    const emptyRow = document.createElement('tr');
    const emptyCell = document.createElement('td');
    emptyCell.colSpan = headers.length;
    emptyCell.textContent = emptyText;
    emptyRow.appendChild(emptyCell);
    tbody.appendChild(emptyRow);
    return fragment;
  }

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    row.forEach((value) => tr.appendChild(createCell('td', value)));
    tbody.appendChild(tr);
  });

  return fragment;
}

function createCell(tagName, text) {
  const cell = document.createElement(tagName);
  cell.textContent = text;
  return cell;
}

function createHoverCard() {
  const card = document.createElement('div');
  card.className = 'slot-hover-card hidden';
  document.body.appendChild(card);
  return card;
}

function handleSlotHoverStart(event) {
  const tooltip = event.currentTarget?.dataset?.tooltip;
  if (!tooltip) {
    return;
  }

  const [title, ...rest] = tooltip.split('\n');
  const html = `<div class="slot-hover-title">${escapeHtml(title)}</div>${rest
    .map((line) => `<div>${escapeHtml(line)}</div>`)
    .join('')}`;
  hoverState.pointerX = event.clientX;
  hoverState.pointerY = event.clientY;

  if (hoverState.showTimer) {
    window.clearTimeout(hoverState.showTimer);
  }

  hoverState.showTimer = window.setTimeout(() => {
    hoverState.showTimer = null;
    hoverCard.innerHTML = html;
    hoverCard.classList.remove('hidden');
    moveHoverCard({ clientX: hoverState.pointerX, clientY: hoverState.pointerY });

    const startY = Math.min(window.innerHeight - hoverCard.offsetHeight - 12, hoverState.pointerY + 14);
    hoverState.x = hoverState.targetX;
    hoverState.y = startY;
    hoverCard.style.left = `${hoverState.x}px`;
    hoverCard.style.top = `${hoverState.y}px`;

    startHoverAnimation();
  }, 1000);
}

function handleSlotHoverMove(event) {
  hoverState.pointerX = event.clientX;
  hoverState.pointerY = event.clientY;

  if (hoverCard.classList.contains('hidden')) {
    return;
  }
  moveHoverCard(event);
}

function handleSlotHoverEnd() {
  if (hoverState.showTimer) {
    window.clearTimeout(hoverState.showTimer);
    hoverState.showTimer = null;
  }
  hoverCard.classList.add('hidden');
  stopHoverAnimation();
}

function moveHoverCard(event) {
  const offset = 14;
  const maxLeft = window.innerWidth - hoverCard.offsetWidth - 12;
  const minLeft = 12;
  const minTop = 12;

  const centeredLeft = event.clientX - hoverCard.offsetWidth / 2;
  const aboveTop = event.clientY - hoverCard.offsetHeight - offset;

  hoverState.targetX = Math.min(maxLeft, Math.max(minLeft, centeredLeft));
  hoverState.targetY = Math.max(minTop, aboveTop);
}

function startHoverAnimation() {
  if (hoverState.rafId) {
    return;
  }

  const tick = () => {
    hoverState.x += (hoverState.targetX - hoverState.x) * 0.12;
    hoverState.y += (hoverState.targetY - hoverState.y) * 0.12;
    hoverCard.style.left = `${hoverState.x}px`;
    hoverCard.style.top = `${hoverState.y}px`;

    if (hoverCard.classList.contains('hidden')) {
      hoverState.rafId = null;
      return;
    }

    hoverState.rafId = window.requestAnimationFrame(tick);
  };

  hoverState.rafId = window.requestAnimationFrame(tick);
}

function stopHoverAnimation() {
  if (!hoverState.rafId) {
    return;
  }

  window.cancelAnimationFrame(hoverState.rafId);
  hoverState.rafId = null;
}

function loadAppointments() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function persistAppointments() {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state.appointments));
}

function ensureValidSelectedDate() {
  const monthDays = buildWorkDaysForMonth(state.currentMonthStart);
  if (!monthDays.length) {
    state.selectedDate = formatDateISO(state.currentMonthStart);
    return;
  }

  const selectedInMonth = monthDays.includes(state.selectedDate);
  if (selectedInMonth) {
    return;
  }

  const todayIso = formatDateISO(new Date());
  if (monthDays.includes(todayIso)) {
    state.selectedDate = todayIso;
    return;
  }

      state.selectedDate = monthDays[0];
}

function getInitialSelectedDate() {
  const now = new Date();
  if (isWeekDay(now)) {
    return formatDateISO(now);
  }

  const fallback = new Date(now);
  while (!isWeekDay(fallback)) {
    fallback.setDate(fallback.getDate() + 1);
  }
  return formatDateISO(fallback);
}

function buildWorkDaysForMonth(monthStartDate) {
  const days = [];
  const date = new Date(monthStartDate);

  while (date.getMonth() === monthStartDate.getMonth()) {
    if (isWeekDay(date)) {
      days.push(formatDateISO(date));
    }
    date.setDate(date.getDate() + 1);
  }

  return days;
}

function buildSlotTimes() {
  const times = [];
  let hour = 8;
  let minute = 0;

  while (hour < 11 || (hour === 11 && minute <= 30)) {
    times.push(`${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`);
    minute += 10;
    if (minute === 60) {
      minute = 0;
      hour += 1;
    }
  }

  return times;
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date, amount) {
  return new Date(date.getFullYear(), date.getMonth() + amount, 1);
}

function isWeekDay(date) {
  const day = date.getDay();
  return day >= 1 && day <= 5;
}

function formatDayHeader(isoDate) {
  const date = new Date(`${isoDate}T00:00:00`);
  const weekday = new Intl.DateTimeFormat('de-DE', { weekday: 'short' }).format(date).slice(0, 2);
  const day = new Intl.DateTimeFormat('de-DE', { day: '2-digit' }).format(date);
  return `${weekday} ${day}`;
}

function formatWeekdayCompact(isoDate) {
  const date = new Date(`${isoDate}T00:00:00`);
  return new Intl.DateTimeFormat('de-DE', { weekday: 'short' })
    .format(date)
    .replace('.', '')
    .slice(0, 2)
    .toUpperCase();
}

function formatDayNumber(isoDate) {
  return String(new Date(`${isoDate}T00:00:00`).getDate());
}

function formatDateISO(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDateGerman(isoDate) {
  return new Intl.DateTimeFormat('de-DE', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(new Date(`${isoDate}T00:00:00`));
}

function formatDateLong(isoDate) {
  return new Intl.DateTimeFormat('de-DE', {
    weekday: 'long',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(new Date(`${isoDate}T00:00:00`));
}

function formatDateTimeGerman(isoDateTime) {
  return new Intl.DateTimeFormat('de-DE', {
    dateStyle: 'short',
    timeStyle: 'short',
  }).format(new Date(isoDateTime));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

// ====== NGB-Funktionen ======
function loadNgbs() {
  try {
    const stored = localStorage.getItem(NGB_STORAGE_KEY);
    return stored ? JSON.parse(stored) : [];
  } catch {
    return [];
  }
}

function persistNgbs() {
  localStorage.setItem(NGB_STORAGE_KEY, JSON.stringify(state.ngbs));
}

function renderNgb() {
  if (!elements.ngbContainer || !elements.ngbSearchInfo) {
    return;
  }

  const query = state.ngbQuery.trim().toLowerCase();
  const filteredItems = query
    ? state.ngbs.filter((ngb) => {
        const haystack = [
          ngb.ngbNumber,
          ngb.recipientName,
          ngb.iban,
          ngb.amount || '',
          ngb.status,
          normalizeTzList(ngb.tzs || ngb.tz).join(' '),
        ]
          .join(' ')
          .toLowerCase();
        return haystack.includes(query);
      })
    : state.ngbs;

  const headers = ['NGB-Nummer', 'Name des Empfängers', 'IBAN', 'TZ', 'Betrag', 'Status'];
  const rows = filteredItems.map((ngb) => [
    ngb.ngbNumber,
    ngb.recipientName,
    ngb.iban,
    formatTzList(normalizeTzList(ngb.tzs || ngb.tz)),
    ngb.amount ? `${ngb.amount}€` : '-',
    ngb.status === 'done' ? '✓ Erledigt' : 'Offen',
  ]);

  elements.ngbSearchInfo.textContent = `${filteredItems.length} Einträge`;
  elements.ngbContainer.replaceChildren(buildTable(headers, rows, 'Keine Einträge vorhanden.'));
}

function openNgbDialog(existingNgbId = null) {
  if (!elements.ngbDialog || !elements.ngbForm) {
    return;
  }

  const existingNgb = state.ngbs.find((ngb) => ngb.id === existingNgbId);
  state.editingNgbId = existingNgbId || null;
  elements.ngbDialogTitle.textContent = existingNgb ? 'NGB-Eintrag bearbeiten' : 'Neuer NGB-Eintrag';
  elements.deleteNgb.classList.toggle('hidden', !existingNgb);

  elements.ngbForm['ngb-number'].value = existingNgb?.ngbNumber || '';
  elements.ngbForm['recipient-name'].value = existingNgb?.recipientName || '';
  elements.ngbForm.iban.value = existingNgb?.iban || '';
  elements.ngbForm.tz.value = formatTzList(normalizeTzList(existingNgb?.tzs || existingNgb?.tz), '');
  elements.ngbForm.amount.value = existingNgb?.amount || '';

  elements.ngbDialog.showModal();
}

function closeNgbDialog() {
  if (!elements.ngbDialog || !elements.ngbForm) {
    return;
  }

  elements.ngbDialog.close();
  state.editingNgbId = null;
  elements.ngbForm.reset();
}

function normalizeTzList(rawTz) {
  if (Array.isArray(rawTz)) {
    return rawTz.map((item) => String(item).trim()).filter(Boolean);
  }

  if (!rawTz) {
    return [];
  }

  return String(rawTz)
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function formatTzList(tzList, fallback = '-') {
  return tzList.length ? tzList.join(', ') : fallback;
}

function handleNgbSubmit(event) {
  event.preventDefault();

  const ngbNumber = elements.ngbForm['ngb-number'].value.trim();
  const recipientName = elements.ngbForm['recipient-name'].value.trim();
  const iban = elements.ngbForm.iban.value.trim();
  const tzs = normalizeTzList(elements.ngbForm.tz.value);
  const amount = elements.ngbForm.amount.value ? parseFloat(elements.ngbForm.amount.value) : null;

  if (!ngbNumber || !recipientName || !iban) {
    alert('NGB-Nummer, Name und IBAN sind erforderlich.');
    return;
  }

  if (state.editingNgbId) {
    const existingNgb = state.ngbs.find((ngb) => ngb.id === state.editingNgbId);
    if (existingNgb) {
      existingNgb.ngbNumber = ngbNumber;
      existingNgb.recipientName = recipientName;
      existingNgb.iban = iban;
      existingNgb.tzs = tzs;
      existingNgb.amount = amount;
    }
  } else {
    state.ngbs.push({
      id: crypto.randomUUID(),
      ngbNumber,
      recipientName,
      iban,
      tzs,
      amount,
      status: 'open',
      createdAt: new Date().toISOString(),
    });
  }

  persistNgbs();
  renderNgb();
  closeNgbDialog();
}

function handleNgbDelete() {
  if (!state.editingNgbId) return;
  const index = state.ngbs.findIndex((ngb) => ngb.id === state.editingNgbId);
  if (index >= 0) {
    state.ngbs.splice(index, 1);
    persistNgbs();
    renderNgb();
    closeNgbDialog();
  }
}

function handleNgbSearchInput(event) {
  state.ngbQuery = event.target.value || '';
  renderNgb();
}

// ====== Travel Cost Calculator Functions ======
function timeToMinutes(timeString) {
  if (!timeString) return null;
  const [h, m] = timeString.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function minutesToTime(minutes) {
  if (minutes === null) return '--:--';
  const adjustedMinutes = ((minutes % 1440) + 1440) % 1440;
  const h = Math.floor(adjustedMinutes / 60);
  const m = adjustedMinutes % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function formatCurrency(amount) {
  return new Intl.NumberFormat('de-DE', {
    style: 'currency',
    currency: 'EUR',
  }).format(amount);
}

function calculateTravelCosts() {
  if (!elements.travelPrice || !elements.travelStart || !elements.travelEnd || !elements.travelDuration || !elements.travelResults) {
    return;
  }

  const priceInput = (elements.travelPrice.value || '').replace(',', '.').trim();
  const price = Number(priceInput);
  const startMin = timeToMinutes(elements.travelStart.value);
  const endMin = timeToMinutes(elements.travelEnd.value);
  const durationMin = timeToMinutes(elements.travelDuration.value);

  if (!Number.isFinite(price) || price < 0 || startMin === null || endMin === null || durationMin === null) {
    elements.travelResults.classList.add('hidden');
    return;
  }

  const setText = (id, value) => {
    const node = document.getElementById(id);
    if (node) {
      node.textContent = value;
    }
  };

  const setActive = (id, active) => {
    const node = document.getElementById(id);
    if (node) {
      node.classList.toggle('active', active);
    }
  };

  // Calculate timeline
  const departHome = startMin - 30 - durationMin;
  const arrivalAppointment = startMin - 30;
  const departAppointment = endMin + 30;
  const arrivalHome = endMin + 30 + durationMin;

  // Display timeline
  setText('travel-t0', minutesToTime(departHome));
  setText('travel-t1', minutesToTime(arrivalAppointment));
  setText('travel-t2', minutesToTime(startMin));
  setText('travel-t3', minutesToTime(endMin));
  setText('travel-t4', minutesToTime(departAppointment));
  setText('travel-t5', minutesToTime(arrivalHome));

  // Calculate costs with explanations
  const trainCost = price * 2;
  
  const earlyCost = departHome < 7 * 60 ? 5.80 : 0;
  const lunchCost = departHome < 11 * 60 && arrivalHome > 14 * 60 ? 12.30 : 0;
  const eveningCost = arrivalHome > 19 * 60 ? 12.30 : 0;
  const totalCost = trainCost + earlyCost + lunchCost + eveningCost;

  // Build explanations
  const trainExpl = `${formatCurrency(price)} × 2 (Hin- & Rückfahrt)`;
  
  const earlyExpl = departHome < 7 * 60 
    ? `Abfahrt ${minutesToTime(departHome)} vor 7:00 Uhr`
    : `Abfahrt ${minutesToTime(departHome)} nach 7:00 Uhr – nicht berechtigt`;
  
  const lunchExpl = departHome < 11 * 60 && arrivalHome > 14 * 60
    ? `Abfahrt vor 11:00 Uhr (${minutesToTime(departHome)}) & Ankunft nach 14:00 Uhr (${minutesToTime(arrivalHome)})`
    : `Abfahrt ${minutesToTime(departHome)} oder Ankunft ${minutesToTime(arrivalHome)} – Bedingung nicht erfüllt`;
  
  const eveningExpl = arrivalHome > 19 * 60
    ? `Ankunft ${minutesToTime(arrivalHome)} nach 19:00 Uhr`
    : `Ankunft ${minutesToTime(arrivalHome)} vor 19:00 Uhr – nicht berechtigt`;

  // Display costs with explanations
  setText('travel-cost-train', formatCurrency(trainCost));
  setText('travel-exp-train', trainExpl);
  
  setText('travel-cost-early', earlyCost > 0 ? formatCurrency(earlyCost) : '–');
  setText('travel-exp-early', earlyExpl);
  
  setText('travel-cost-lunch', lunchCost > 0 ? formatCurrency(lunchCost) : '–');
  setText('travel-exp-lunch', lunchExpl);
  
  setText('travel-cost-evening', eveningCost > 0 ? formatCurrency(eveningCost) : '–');
  setText('travel-exp-evening', eveningExpl);
  
  setText('travel-cost-total', formatCurrency(totalCost));

  // Toggle visibility of meal allowances
  setActive('travel-row-early', earlyCost > 0);
  setActive('travel-row-lunch', lunchCost > 0);
  setActive('travel-row-evening', eveningCost > 0);

  // Show results
  elements.travelResults.classList.remove('hidden');
}
