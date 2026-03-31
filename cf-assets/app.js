const STORAGE_KEY = 'calcal-appointments-v1';
const NGB_STORAGE_KEY = 'calcal-ngb-v1';
const TRAVEL_DESTINATION_KEY = 'calcal-travel-destination-v1';
const NGB_PAGE_RATE = 1.6;
const SLOT_TIMES = buildSlotTimes();

const today = new Date();
const state = {
  appointments: loadAppointments(),
  ngbs: loadNgbs(),
  importPreview: [],
  currentMonthStart: startOfMonth(today),
  selectedDate: '',
  editingKey: null,
  editingAppointmentId: null,
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
  hideTimer: null,
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
  printDaylistTools: document.getElementById('print-daylist-tools'),
  ngbOpenCount: document.getElementById('ngb-open-count'),
  ngbLatest: document.getElementById('ngb-latest'),
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
  travelOrigin: document.getElementById('travel-origin'),
  travelDestination: document.getElementById('travel-destination'),
  travelPrice: document.getElementById('travel-price'),
  travelStart: document.getElementById('travel-start'),
  travelEnd: document.getElementById('travel-end'),
  travelDuration: document.getElementById('travel-duration'),
  travelCalc: document.getElementById('travel-calc'),
  travelMaps: document.getElementById('travel-maps'),
  travelResults: document.getElementById('travel-results'),
  installApp: document.getElementById('install-app'),
  exportMonthOverview: document.getElementById('export-month-overview'),
  exportBothWeeks: document.getElementById('export-both-weeks'),
  chooseOverviewDir: document.getElementById('choose-overview-dir'),
  exportData: document.getElementById('export-data'),
  importData: document.getElementById('import-data'),
  importDataFile: document.getElementById('import-data-file'),
  importAssistantDialog: document.getElementById('import-assistant-dialog'),
  importAssistantSummary: document.getElementById('import-assistant-summary'),
  importAssistantList: document.getElementById('import-assistant-list'),
  closeImportAssistant: document.getElementById('close-import-assistant'),
  cancelImportAssistant: document.getElementById('cancel-import-assistant'),
  confirmImportAssistant: document.getElementById('confirm-import-assistant'),
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
  ngbCreatedAt: document.getElementById('ngb-created-at'),
  ngbTzList: document.getElementById('ngb-tz-list'),
  addNgbTz: document.getElementById('add-ngb-tz'),
  saveNgb: document.getElementById('save-ngb'),
  closeNgbDialog: document.getElementById('close-ngb-dialog'),
  cancelNgbDialog: document.getElementById('cancel-ngb-dialog'),
  deleteNgb: document.getElementById('delete-ngb'),
};

bootstrap();

function bootstrap() {
  archivePastAppointments();
  state.selectedDate = getInitialSelectedDate();
  hydrateTravelFields();
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
  if (elements.printDaylist) {
    elements.printDaylist.addEventListener('click', printSelectedDayTwice);
  }
  if (elements.printDaylistTools) {
    elements.printDaylistTools.addEventListener('click', printSelectedDayTwice);
  }
  elements.archiveSearch.addEventListener('input', handleArchiveSearchInput);
  elements.exportArchiveCsv.addEventListener('click', exportArchiveCsv);
  elements.installApp.addEventListener('click', installApp);
  elements.exportMonthOverview.addEventListener('click', exportMonthlyOverview);
  elements.exportBothWeeks.addEventListener('click', exportBothWeeksOverview);
  if (elements.chooseOverviewDir) {
    elements.chooseOverviewDir.addEventListener('click', chooseOverviewDirectory);
  }
  elements.exportData.addEventListener('click', exportDataJson);
  elements.importData.addEventListener('click', () => elements.importDataFile.click());
  elements.importDataFile.addEventListener('change', importDataJson);
  if (elements.closeImportAssistant) {
    elements.closeImportAssistant.addEventListener('click', closeImportAssistantDialog);
  }
  if (elements.cancelImportAssistant) {
    elements.cancelImportAssistant.addEventListener('click', closeImportAssistantDialog);
  }
  if (elements.confirmImportAssistant) {
    elements.confirmImportAssistant.addEventListener('click', confirmImportAssistant);
  }

  elements.appointmentForm.addEventListener('submit', handleSubmit);
  elements.closeDialog.addEventListener('click', closeDialog);
  elements.cancelDialog.addEventListener('click', closeDialog);
  elements.deleteAppointment.addEventListener('click', handleDelete);
  elements.moveAppointment.addEventListener('click', handleMoveStart);
  
  if (elements.newNgb && elements.ngbSearch && elements.ngbForm && elements.closeNgbDialog && elements.cancelNgbDialog && elements.deleteNgb) {
    elements.newNgb.addEventListener('click', () => openNgbDialog(null));
    elements.ngbSearch.addEventListener('input', handleNgbSearchInput);
    elements.ngbForm.addEventListener('submit', handleNgbSubmit);
    const pageCountInput = elements.ngbForm.querySelector('[name="page-count"]');
    if (pageCountInput) {
      pageCountInput.addEventListener('input', updateNgbAmountPreview);
    }
    if (elements.addNgbTz) {
      elements.addNgbTz.addEventListener('click', () => addNgbTzInput(''));
    }
    if (elements.saveNgb) {
      elements.saveNgb.addEventListener('click', () => {
        if (typeof elements.ngbForm.requestSubmit === 'function') {
          elements.ngbForm.requestSubmit();
          return;
        }
        elements.ngbForm.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
      });
    }
    elements.closeNgbDialog.addEventListener('click', closeNgbDialog);
    elements.cancelNgbDialog.addEventListener('click', closeNgbDialog);
    elements.deleteNgb.addEventListener('click', handleNgbDelete);
  }

  if (elements.travelCalc && elements.travelPrice && elements.travelStart && elements.travelEnd && elements.travelDuration) {
    elements.travelCalc.addEventListener('click', calculateTravelCosts);
    elements.travelPrice.addEventListener('input', calculateTravelCosts);
    elements.travelStart.addEventListener('input', calculateTravelCosts);
    elements.travelEnd.addEventListener('input', calculateTravelCosts);
    elements.travelDuration.addEventListener('input', calculateTravelCosts);
  }

  if (elements.travelDestination) {
    elements.travelDestination.addEventListener('input', persistTravelDestination);
  }

  if (elements.travelMaps) {
    elements.travelMaps.addEventListener('click', openGoogleMapsRoutePlanner);
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
    travel: 'ZGB',
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
  renderNgbSummary();
  renderArchive();
  renderNgb();
  updateSelectionSummary();
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
      <small>${countCalendarAppointmentsForDay(date)}</small>
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

      const appointment = getCalendarAppointment(date, time);
      if (appointment) {
        const bookingClass = getSlotBookingClass(date, time, appointment);
        td.classList.add(bookingClass);
        td.dataset.tooltip = [
          `${formatDateLong(date)}`,
          `${time}`,
          `Status: ${getStatusLabel(appointment.status)}`,
          `Name: ${appointment.name}`,
          `Telefon: ${appointment.phone}`,
          `KG: ${appointment.kg || '-'}`,
          `EZ: ${appointment.ez || '-'}`,
          `Anliegen: ${appointment.concern || '-'}`,
        ].join('\n');
      } else {
        td.classList.add('slot-empty-cell');
        td.dataset.tooltip = [
          `${formatDateLong(date)}`,
          'Status: frei',
          `${time}`,
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
  const movingAppointment = state.appointments.find((appointment) => appointment.id === state.moveAppointmentId);

  if (!movingAppointment) {
    state.moveAppointmentId = null;
    renderAll();
    return;
  }

  const targetAppointment = getCalendarAppointment(date, time);
  if (targetAppointment && targetAppointment.id !== movingAppointment.id) {
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

function renderNgbSummary() {
  const openNgbCount = state.ngbs.filter((ngb) => ngb.status !== 'done').length;
  if (elements.ngbOpenCount) {
    elements.ngbOpenCount.textContent = String(openNgbCount);
  }

  if (!elements.ngbLatest) {
    return;
  }

  const latestNgbs = [...state.ngbs]
    .sort((left, right) => {
      const leftDate = left.createdAt || '';
      const rightDate = right.createdAt || '';
      return rightDate.localeCompare(leftDate);
    })
    .slice(0, 3);

  const rows = [0, 1, 2].map((index) => {
    const ngb = latestNgbs[index];
    if (!ngb) {
      return '<div class="next-three-item empty">-</div>';
    }

    return `<div class="next-three-item">${escapeHtml(ngb.recipientName || '-')} · NGB ${escapeHtml(
      ngb.ngbNumber || '-'
    )}</div>`;
  });

  elements.ngbLatest.innerHTML = rows.join('');
}

function updateSelectionSummary() {
  elements.selectedDateLabel.textContent = formatDateLong(state.selectedDate);
  elements.selectedDayCount.textContent = `${countAppointmentsForDay(state.selectedDate)} Termine`;
}

function openDialog(date, time, forceNew = false) {
  const appointment = getCalendarAppointment(date, time);
  const isEditingExisting = Boolean(appointment) && !forceNew;
  state.editingKey = isEditingExisting ? `${date}__${time}` : null;
  state.editingAppointmentId = isEditingExisting ? appointment.id : null;
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
  state.editingAppointmentId = null;
  elements.appointmentForm.reset();
}

function handleMoveStart() {
  if (!state.editingAppointmentId) {
    return;
  }

  const appointment = state.appointments.find((item) => item.id === state.editingAppointmentId);
  if (!appointment) {
    return;
  }

  state.moveAppointmentId = appointment.id;
  closeDialog();
  setView('calendar');
}

function handleSubmit(event) {
  event.preventDefault();

  const targetDate = elements.appointmentForm.date.value;
  const targetTime = elements.appointmentForm.time.value;
  const existingAppointment = state.editingAppointmentId
    ? state.appointments.find((appointment) => appointment.id === state.editingAppointmentId)
    : null;

  const conflictingAppointment = getCalendarAppointment(targetDate, targetTime);
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
    status: existingAppointment?.status || 'open',
    archivedAt: existingAppointment?.archivedAt || null,
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
  if (!state.editingAppointmentId) {
    return;
  }

  state.appointments = state.appointments.filter((appointment) => appointment.id !== state.editingAppointmentId);
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

function getCalendarAppointmentsForDay(date) {
  return state.appointments.filter((appointment) => appointment.date === date);
}

function countAppointmentsForDay(date) {
  return getAppointmentsForDay(date).length;
}

function countCalendarAppointmentsForDay(date) {
  return getCalendarAppointmentsForDay(date).length;
}

function getDayStatusClass(date) {
  const count = countCalendarAppointmentsForDay(date);
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

function getCalendarAppointment(date, time) {
  const matches = state.appointments.filter(
    (appointment) => appointment.date === date && appointment.time === time
  );

  if (!matches.length) {
    return null;
  }

  matches.sort((left, right) => {
    if (left.status === 'open' && right.status !== 'open') {
      return -1;
    }
    if (left.status !== 'open' && right.status === 'open') {
      return 1;
    }
    const leftUpdated = left.updatedAt || left.createdAt || '';
    const rightUpdated = right.updatedAt || right.createdAt || '';
    return rightUpdated.localeCompare(leftUpdated);
  });

  return matches[0];
}

function getSlotBookingClass(date, time, appointment = null) {
  if (appointment && appointment.status !== 'open') {
    return 'slot-booked-past';
  }

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

async function chooseOverviewDirectory() {
  if (!window.calcalDesktop?.isDesktopApp || typeof window.calcalDesktop.chooseOverviewDirectory !== 'function') {
    alert('Ordnerwahl ist nur in der Desktop-App verfügbar.');
    return;
  }

  try {
    const result = await window.calcalDesktop.chooseOverviewDirectory();
    if (result?.ok && result.directory) {
      alert(`Speicherordner gesetzt: ${result.directory}`);
      return;
    }
    if (!result?.canceled) {
      alert('Speicherordner konnte nicht gesetzt werden.');
    }
  } catch {
    alert('Ordnerwahl fehlgeschlagen.');
  }
}

async function saveOverviewDesktop(kind, htmlContent, fallbackFileName) {
  if (!window.calcalDesktop?.isDesktopApp || typeof window.calcalDesktop.saveOverviewHtml !== 'function') {
    return false;
  }

  try {
    const result = await window.calcalDesktop.saveOverviewHtml({ kind, htmlContent });
    if (result?.ok) {
      alert(`Übersicht gespeichert: ${result.filePath}`);
      return true;
    }
    alert('Speichern im Zielordner fehlgeschlagen. Fallback-Download wird gestartet.');
  } catch {
    alert('Speichern im Zielordner fehlgeschlagen. Fallback-Download wird gestartet.');
  }

  const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fallbackFileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  return true;
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

  if (window.calcalDesktop?.isDesktopApp) {
    await saveOverviewDesktop('monthly', html, 'Monatsuebersicht-Aktuell.html');
    return;
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

  if (window.calcalDesktop?.isDesktopApp) {
    await saveOverviewDesktop('weekly', html, 'Wochenuebersicht-Aktuell.html');
    return;
  }

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
    const appointments = await parseImportFile(file);
    if (!appointments.length) {
      throw new Error('Keine importierbaren Termine erkannt.');
    }
    openImportAssistantDialog(appointments, file.name);
  } catch {
    alert('Import fehlgeschlagen. Bitte JSON oder Excel-Datei prüfen.');
  } finally {
    event.target.value = '';
  }
}

async function parseImportFile(file) {
  const lowerName = (file.name || '').toLowerCase();
  if (lowerName.endsWith('.json')) {
    return parseJsonImport(file);
  }

  if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls')) {
    return parseExcelImport(file);
  }

  throw new Error('Dateiformat nicht unterstützt');
}

async function parseJsonImport(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const appointments = Array.isArray(parsed) ? parsed : parsed.appointments;
  if (!Array.isArray(appointments)) {
    throw new Error('Ungueltiges JSON-Format');
  }

  return appointments
    .filter((item) => item && item.date && item.time && item.name)
    .map((item) => ({
      id: item.id || createId(),
      date: item.date,
      time: item.time,
      name: item.name,
      kg: item.kg || '',
      ez: item.ez || '',
      phone: item.phone || '-',
      concern: item.concern || '',
      status: item.status || 'done',
      archivedAt: item.archivedAt || new Date().toISOString(),
      createdAt: item.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    }));
}

async function parseExcelImport(file) {
  if (!window.XLSX) {
    throw new Error('Excel-Bibliothek nicht verfügbar');
  }

  const buffer = await file.arrayBuffer();
  const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error('Leeres Excel');
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
  return parseExcelRowsToAppointments(rows);
}

function parseExcelRowsToAppointments(rows) {
  const colDateMap = {};

  for (let r = 0; r < Math.min(rows.length, 12); r += 1) {
    const row = rows[r] || [];
    row.forEach((cell, colIndex) => {
      if (colDateMap[colIndex]) {
        return;
      }
      const parsedDate = parsePossibleDate(cell);
      if (parsedDate) {
        colDateMap[colIndex] = parsedDate;
      }
    });
  }

  const collected = [];
  rows.forEach((row) => {
    const rowDate = parsePossibleDate(row[0]) || parsePossibleDate(row[1]);
    const rowTime = findTimeInRow(row);

    row.forEach((cell, colIndex) => {
      const text = String(cell || '').replace(/\s+/g, ' ').trim();
      if (!text || !/\b66\d+\b/.test(text)) {
        return;
      }

      const date = colDateMap[colIndex] || rowDate;
      if (!date) {
        return;
      }

      const parsed = parseLegacyCellText(text);
      if (!parsed.name) {
        return;
      }

      collected.push({
        id: createId(),
        date,
        time: rowTime || '08:00',
        name: parsed.name,
        kg: parsed.kg,
        ez: parsed.ez,
        phone: '-',
        concern: parsed.concern,
        status: 'done',
        archivedAt: new Date().toISOString(),
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      });
    });
  });

  return dedupeImportedAppointments(collected);
}

function dedupeImportedAppointments(items) {
  const map = new Map();
  items.forEach((item) => {
    const key = `${item.date}__${item.time}__${item.name}`;
    if (!map.has(key)) {
      map.set(key, item);
    }
  });
  return [...map.values()];
}

function parsePossibleDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatDateISO(value);
  }

  const text = String(value || '').trim();
  if (!text) {
    return null;
  }

  const germanMatch = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (germanMatch) {
    const day = Number(germanMatch[1]);
    const month = Number(germanMatch[2]);
    const year = Number(germanMatch[3].length === 2 ? `20${germanMatch[3]}` : germanMatch[3]);
    const date = new Date(year, month - 1, day);
    if (!Number.isNaN(date.getTime())) {
      return formatDateISO(date);
    }
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return text;
  }

  return null;
}

function findTimeInRow(row) {
  for (let i = 0; i < Math.min(row.length, 6); i += 1) {
    const text = String(row[i] || '').trim();
    const match = text.match(/^([01]?\d|2[0-3]):([0-5]\d)$/);
    if (match) {
      return `${String(match[1]).padStart(2, '0')}:${match[2]}`;
    }
  }
  return null;
}

function parseLegacyCellText(text) {
  const clean = text.replace(/\s+/g, ' ').trim();
  const kgMatches = [...clean.matchAll(/\b66\d+\b/g)].map((m) => m[0]);
  const uniqueKg = [...new Set(kgMatches)];

  const ezMatches = [...clean.matchAll(/\bEZ[:\s-]*([A-Za-z0-9/.,-]+)/gi)].map((m) => m[1]);
  const normalizedEz = ezMatches
    .flatMap((part) => String(part).split(/[;,/]+/))
    .map((part) => part.trim())
    .filter(Boolean);

  const firstKgIndex = clean.search(/\b66\d+\b/);
  const name = firstKgIndex > 0 ? clean.slice(0, firstKgIndex).trim() : clean;

  let concern = clean;
  uniqueKg.forEach((kg) => {
    concern = concern.replace(kg, '').trim();
  });
  concern = concern.replace(/\bEZ[:\s-]*[A-Za-z0-9/.,-]+/gi, '').replace(/\s+/g, ' ').trim();
  if (name && concern.startsWith(name)) {
    concern = concern.slice(name.length).trim();
  }

  return {
    name,
    kg: uniqueKg.join(', '),
    ez: normalizedEz.join(', '),
    concern,
  };
}

function openImportAssistantDialog(appointments, sourceName) {
  if (!elements.importAssistantDialog || !elements.importAssistantSummary || !elements.importAssistantList) {
    state.appointments = [...state.appointments, ...appointments];
    persistAppointments();
    renderAll();
    alert('Daten wurden erfolgreich importiert.');
    return;
  }

  state.importPreview = appointments;
  elements.importAssistantSummary.textContent = `${appointments.length} Termine aus ${sourceName} erkannt`;
  elements.importAssistantList.innerHTML = appointments
    .slice(0, 120)
    .map((item) => {
      const detail = [item.name, item.kg ? `KG ${item.kg}` : '', item.ez ? `EZ ${item.ez}` : '', item.concern || '']
        .filter(Boolean)
        .join(' · ');
      return `<div class="import-assistant-row"><span>${escapeHtml(formatDateGerman(item.date))}</span><span>${escapeHtml(
        item.time
      )}</span><span>${escapeHtml(detail)}</span></div>`;
    })
    .join('');

  elements.importAssistantDialog.showModal();
}

function closeImportAssistantDialog() {
  if (!elements.importAssistantDialog) {
    return;
  }
  elements.importAssistantDialog.close();
  state.importPreview = [];
}

function confirmImportAssistant() {
  if (!state.importPreview.length) {
    closeImportAssistantDialog();
    return;
  }

  state.appointments = [...state.appointments, ...state.importPreview];
  persistAppointments();
  renderAll();
  alert(`${state.importPreview.length} Termine importiert.`);
  closeImportAssistantDialog();
}

function getVisibleSlotTimes(monthDays) {
  const filledTimes = new Set(
    state.appointments
      .filter((appointment) => monthDays.includes(appointment.date))
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
  if (event.currentTarget?.classList?.contains('slot-booked-past')) {
    event.currentTarget.classList.add('slot-booked-past-animate');
  }

  if (event.currentTarget?.classList?.contains('slot-booked-future')) {
    event.currentTarget.classList.add('slot-booked-future-animate');
  }

  const tooltip = event.currentTarget?.dataset?.tooltip;
  if (!tooltip) {
    return;
  }

  const [title, ...rest] = tooltip.split('\n');
  const html = `<div class="slot-hover-title">${escapeHtml(title)}</div>${rest
    .map((line) => {
      if (/^\d{2}:\d{2}$/.test(line.trim())) {
        return `<div><strong>${escapeHtml(line)}</strong></div>`;
      }
      return `<div>${escapeHtml(line)}</div>`;
    })
    .join('')}`;
  hoverState.pointerX = event.clientX;
  hoverState.pointerY = event.clientY;

  if (hoverState.showTimer) {
    window.clearTimeout(hoverState.showTimer);
  }

  if (hoverState.hideTimer) {
    window.clearTimeout(hoverState.hideTimer);
    hoverState.hideTimer = null;
  }

  hoverState.showTimer = window.setTimeout(() => {
    hoverState.showTimer = null;
    hoverCard.innerHTML = html;
    hoverCard.classList.remove('fade-out');
    hoverCard.classList.remove('hidden');
    moveHoverCard({ clientX: hoverState.pointerX, clientY: hoverState.pointerY });

    const startY = Math.min(window.innerHeight - hoverCard.offsetHeight - 12, hoverState.pointerY + 14);
    hoverState.x = hoverState.targetX;
    hoverState.y = startY;
    hoverCard.style.left = `${hoverState.x}px`;
    hoverCard.style.top = `${hoverState.y}px`;

    startHoverAnimation();
  }, 350);
}

function handleSlotHoverMove(event) {
  hoverState.pointerX = event.clientX;
  hoverState.pointerY = event.clientY;

  if (hoverCard.classList.contains('hidden')) {
    return;
  }
  moveHoverCard(event);
}

function handleSlotHoverEnd(event) {
  if (event.currentTarget?.classList?.contains('slot-booked-past-animate')) {
    event.currentTarget.classList.remove('slot-booked-past-animate');
  }

  if (event.currentTarget?.classList?.contains('slot-booked-future-animate')) {
    event.currentTarget.classList.remove('slot-booked-future-animate');
  }

  if (hoverState.showTimer) {
    window.clearTimeout(hoverState.showTimer);
    hoverState.showTimer = null;
  }

  if (hoverCard.classList.contains('hidden')) {
    return;
  }

  hoverCard.classList.add('fade-out');
  if (hoverState.hideTimer) {
    window.clearTimeout(hoverState.hideTimer);
  }
  hoverState.hideTimer = window.setTimeout(() => {
    hoverCard.classList.add('hidden');
    hoverCard.classList.remove('fade-out');
    stopHoverAnimation();
    hoverState.hideTimer = null;
  }, 240);
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

  const sortedNgbs = [...state.ngbs].sort((left, right) => {
    const leftDate = left.createdAt || '';
    const rightDate = right.createdAt || '';
    return leftDate.localeCompare(rightDate);
  });

  const query = state.ngbQuery.trim().toLowerCase();
  const filteredItems = query
    ? sortedNgbs.filter((ngb) => {
        const haystack = [
          ngb.ngbNumber,
          ngb.recipientName,
          ngb.iban,
          ngb.pageCount ?? '',
          getNgbAmount(ngb) ?? '',
          ngb.status,
          ngb.createdAt ? formatDateTimeGerman(ngb.createdAt) : '',
          normalizeTzList(ngb.tzs || ngb.tz).join(' '),
        ]
          .join(' ')
          .toLowerCase();
        return haystack.includes(query);
      })
    : sortedNgbs;

  const headers = ['NGB-Nummer', 'Name des Empfängers', 'IBAN', 'TZ', 'Seiten', 'Betrag', 'Status', 'Angelegt am'];
  const rows = filteredItems.map((ngb) => [
    ngb.ngbNumber,
    ngb.recipientName,
    ngb.iban,
    formatTzList(normalizeTzList(ngb.tzs || ngb.tz)),
    ngb.pageCount ?? '-',
    getNgbAmount(ngb) !== null ? `${formatAmount(getNgbAmount(ngb))}€` : '-',
    ngb.status === 'done' ? '✓ Erledigt' : 'Offen',
    ngb.createdAt ? formatDateTimeGerman(ngb.createdAt) : '-',
  ]);

  elements.ngbSearchInfo.textContent = `${filteredItems.length} Einträge`;
  const table = buildTable(headers, rows, 'Keine Einträge vorhanden.');
  const bodyRows = table.querySelectorAll('tbody tr');
  bodyRows.forEach((row, index) => {
    const item = filteredItems[index];
    if (!item) {
      return;
    }
    row.style.cursor = 'pointer';
    row.addEventListener('click', () => openNgbDialog(item.id));
  });

  elements.ngbContainer.replaceChildren(table);
  renderNgbSummary();
}

function getNgbFormFields() {
  if (!elements.ngbForm) {
    return null;
  }

  return {
    ngbNumber: elements.ngbForm.querySelector('[name="ngb-number"]'),
    recipientName: elements.ngbForm.querySelector('[name="recipient-name"]'),
    iban: elements.ngbForm.querySelector('[name="iban"]'),
    pageCount: elements.ngbForm.querySelector('[name="page-count"]'),
    amount: elements.ngbForm.querySelector('[name="amount"]'),
    statusDone: elements.ngbForm.querySelector('[name="status-done"]'),
  };
}

function formatAmount(amount) {
  const value = Number(amount);
  if (!Number.isFinite(value)) {
    return '-';
  }
  return value.toFixed(2).replace('.', ',');
}

function getNgbAmount(ngb) {
  if (typeof ngb.pageCount === 'number' && Number.isFinite(ngb.pageCount)) {
    return ngb.pageCount * NGB_PAGE_RATE;
  }
  if (typeof ngb.amount === 'number' && Number.isFinite(ngb.amount)) {
    return ngb.amount;
  }
  return null;
}

function createNgbTzRow(value = '') {
  const row = document.createElement('div');
  row.className = 'ngb-tz-row';

  const input = document.createElement('input');
  input.type = 'text';
  input.name = 'tz-item';
  input.placeholder = 'TZ';
  input.value = value;

  const removeButton = document.createElement('button');
  removeButton.type = 'button';
  removeButton.className = 'ngb-remove-tz';
  removeButton.textContent = '−';
  removeButton.title = 'TZ entfernen';
  removeButton.addEventListener('click', () => {
    row.remove();
    ensureAtLeastOneTzRow();
  });

  row.append(input, removeButton);
  return row;
}

function ensureAtLeastOneTzRow() {
  if (!elements.ngbTzList) {
    return;
  }
  if (!elements.ngbTzList.querySelector('[name="tz-item"]')) {
    elements.ngbTzList.appendChild(createNgbTzRow(''));
  }
}

function setNgbTzRows(values) {
  if (!elements.ngbTzList) {
    return;
  }

  elements.ngbTzList.innerHTML = '';
  const normalized = normalizeTzList(values);
  const list = normalized.length ? normalized : [''];
  list.forEach((item) => {
    elements.ngbTzList.appendChild(createNgbTzRow(item));
  });
}

function addNgbTzInput(value = '') {
  if (!elements.ngbTzList) {
    return;
  }
  elements.ngbTzList.appendChild(createNgbTzRow(value));
}

function collectNgbTzs() {
  if (!elements.ngbTzList) {
    return [];
  }

  return [...elements.ngbTzList.querySelectorAll('[name="tz-item"]')]
    .map((input) => input.value.trim())
    .filter(Boolean);
}

function updateNgbAmountPreview() {
  const fields = getNgbFormFields();
  if (!fields || !fields.pageCount || !fields.amount) {
    return;
  }

  const pageCountRaw = fields.pageCount.value.trim();
  if (!pageCountRaw) {
    fields.amount.value = '';
    return;
  }

  const pageCount = Number(pageCountRaw.replace(',', '.'));
  if (!Number.isFinite(pageCount) || pageCount < 0) {
    fields.amount.value = '';
    return;
  }

  fields.amount.value = `${formatAmount(pageCount * NGB_PAGE_RATE)} €`;
}

function createId() {
  if (window.crypto && typeof window.crypto.randomUUID === 'function') {
    return window.crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function openNgbDialog(existingNgbId = null) {
  if (!elements.ngbDialog || !elements.ngbForm) {
    return;
  }

  const fields = getNgbFormFields();
  if (!fields || !fields.ngbNumber || !fields.recipientName || !fields.iban || !fields.pageCount || !fields.amount || !fields.statusDone) {
    return;
  }

  const existingNgb = state.ngbs.find((ngb) => ngb.id === existingNgbId);
  state.editingNgbId = existingNgbId || null;
  elements.ngbDialogTitle.textContent = existingNgb ? 'NGB-Eintrag bearbeiten' : 'Neuer NGB-Eintrag';
  if (elements.ngbCreatedAt) {
    elements.ngbCreatedAt.textContent = existingNgb?.createdAt
      ? `Angelegt am: ${formatDateTimeGerman(existingNgb.createdAt)}`
      : 'Angelegt am: wird beim Speichern gesetzt';
  }
  elements.deleteNgb.classList.toggle('hidden', !existingNgb);

  fields.ngbNumber.value = existingNgb?.ngbNumber || '';
  fields.recipientName.value = existingNgb?.recipientName || '';
  fields.iban.value = existingNgb?.iban || '';
  setNgbTzRows(existingNgb?.tzs || existingNgb?.tz || []);
  fields.pageCount.value = existingNgb?.pageCount ?? '';
  updateNgbAmountPreview();
  fields.statusDone.checked = existingNgb?.status === 'done';

  elements.ngbDialog.showModal();
}

function closeNgbDialog() {
  if (!elements.ngbDialog || !elements.ngbForm) {
    return;
  }

  elements.ngbDialog.close();
  state.editingNgbId = null;
  if (elements.ngbCreatedAt) {
    elements.ngbCreatedAt.textContent = 'Angelegt am: wird beim Speichern gesetzt';
  }
  setNgbTzRows([]);
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

  const fields = getNgbFormFields();
  if (!fields || !fields.ngbNumber || !fields.recipientName || !fields.iban || !fields.pageCount || !fields.amount || !fields.statusDone) {
    alert('NGB-Formular ist nicht vollständig geladen. Bitte Seite neu laden.');
    return;
  }

  const ngbNumber = fields.ngbNumber.value.trim();
  const recipientName = fields.recipientName.value.trim();
  const iban = fields.iban.value.trim();
  const tzs = collectNgbTzs();
  const pageCountRaw = fields.pageCount.value.trim();
  const pageCount = pageCountRaw ? Number(pageCountRaw.replace(',', '.')) : null;
  const amount = pageCount === null ? null : pageCount * NGB_PAGE_RATE;
  const status = fields.statusDone.checked ? 'done' : 'open';

  if (!ngbNumber || !recipientName || !iban) {
    alert('NGB-Nummer, Name und IBAN sind erforderlich.');
    return;
  }

  if (pageCount !== null && (!Number.isFinite(pageCount) || pageCount < 0)) {
    alert('Seitenzahl ist ungültig. Bitte eine Zahl ab 0 eingeben.');
    return;
  }

  if (state.editingNgbId) {
    const existingNgb = state.ngbs.find((ngb) => ngb.id === state.editingNgbId);
    if (existingNgb) {
      existingNgb.ngbNumber = ngbNumber;
      existingNgb.recipientName = recipientName;
      existingNgb.iban = iban;
      existingNgb.tzs = tzs;
      existingNgb.pageCount = pageCount;
      existingNgb.amount = amount;
      existingNgb.status = status;
    }
  } else {
    state.ngbs.push({
      id: createId(),
      ngbNumber,
      recipientName,
      iban,
      tzs,
      pageCount,
      amount,
      status,
      createdAt: new Date().toISOString(),
    });
  }

  state.ngbQuery = '';
  if (elements.ngbSearch) {
    elements.ngbSearch.value = '';
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

function parseDurationMinutes(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return 0;
  }

  if (raw.includes(':')) {
    return timeToMinutes(raw);
  }

  const minutes = Number(raw.replace(',', '.'));
  if (!Number.isFinite(minutes) || minutes < 0) {
    return null;
  }
  return Math.round(minutes);
}

function hydrateTravelFields() {
  if (!elements.travelDestination) {
    return;
  }

  try {
    const savedDestination = window.localStorage.getItem(TRAVEL_DESTINATION_KEY);
    if (savedDestination) {
      elements.travelDestination.value = savedDestination;
    }
  } catch {
    // localStorage kann in restriktiven Umgebungen blockiert sein.
  }
}

function persistTravelDestination() {
  if (!elements.travelDestination) {
    return;
  }

  try {
    window.localStorage.setItem(TRAVEL_DESTINATION_KEY, elements.travelDestination.value || '');
  } catch {
    // Speichern optional; App muss auch ohne localStorage funktionieren.
  }
}

function openGoogleMapsRoutePlanner() {
  const origin = (elements.travelOrigin?.value || '').trim();
  const destination = (elements.travelDestination?.value || '').trim();

  const params = new URLSearchParams({ api: '1' });
  if (origin) {
    params.set('origin', origin);
  }
  if (destination) {
    params.set('destination', destination);
  }

  const url = `https://www.google.com/maps/dir/?${params.toString()}`;
  window.open(url, '_blank', 'noopener,noreferrer');
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
  const durationMin = parseDurationMinutes(elements.travelDuration.value);

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
