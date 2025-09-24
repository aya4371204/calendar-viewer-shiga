// --- GAPI と GIS の初期化 (グローバルスコープ) ---
let tokenClient;
let gapiInited = false;
let gisInited = false;

// ★★★ GCPで作成したウェブアプリケーション用の情報を設定 ★★★
// APIキーは既存のものを使用
const API_KEY = 'AIzaSyCpRjx_lkdpcp-eePb-_psrh5MUB-T06aA'; 
// 滋賀工場専用に作成した新しいクライアントID
const CLIENT_ID = '976002357617-gfbg1g3obnb323cr845a6nmebucpe67h.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly';

function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
    try {
        await gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"],
        });
        gapiInited = true;
        document.dispatchEvent(new Event('gapiReady'));
    } catch (error) {
        console.error("Error initializing GAPI client: ", error);
        document.getElementById('error').textContent = 'Google APIの初期化に失敗しました。APIキーの設定を確認してください。';
        document.getElementById('error').style.display = 'block';
    }
}

function gisLoaded() {
    try {
        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: CLIENT_ID,
            scope: SCOPES,
            callback: '',
        });
        gisInited = true;
        document.dispatchEvent(new Event('gisReady'));
    } catch (error) {
        console.error("Error initializing GIS client: ", error);
        document.getElementById('error').textContent = 'Google認証の初期化に失敗しました。クライアントIDの設定を確認してください。';
        document.getElementById('error').style.display = 'block';
    }
}

// --- メインのアプリケーションロジック ---
document.addEventListener('DOMContentLoaded', () => {

    // --- DOM Elements ---
    const signInButton = document.getElementById('signInButton');
    const signOutButton = document.getElementById('signOutButton');
    const controlsWrapper = document.getElementById('controls-wrapper');
    const dataDisplayArea = document.getElementById('dataDisplayArea');
    const loadingDiv = document.getElementById('loading');
    const errorDiv = document.getElementById('error');
    const dailyDatePicker = document.getElementById('dailyDatePicker');
    const prevDayBtn = document.getElementById('prevDayBtn');
    const nextDayBtn = document.getElementById('nextDayBtn');
    const bookingModalBackdrop = document.getElementById('bookingModalBackdrop');
    const bookingModal = document.getElementById('bookingModal');
    const bookingResourceName = document.getElementById('bookingResourceName');
    const bookingStartTimeSelect = document.getElementById('bookingStartTime');
    const bookingEndTimeSelect = document.getElementById('bookingEndTime');
    const eventTitleInput = document.getElementById('eventTitle');
    const targetCalendarSelect = document.getElementById('targetCalendarSelect');
    const saveBookingBtn = document.getElementById('saveBookingBtn');
    const cancelBookingBtn = document.getElementById('cancelBookingBtn');
    
    let bookingData = {};

    // --- Resource Data ---
    const resourceCalendarItems = [
        { name: "1-4号館/第一応接室", id: "c_1884ojmrqcv0iitpiq88jj7i96msk@resource.calendar.google.com", type: "room" },
        { name: "1-4号館/第一会議室", id: "c_188a7oshg5c5ijatnn3hs9cc4meie@resource.calendar.google.com", type: "room" },
        { name: "1-4号館/第二会議室", id: "c_1881acanrq6f0jitl0grsoqhfg1b6@resource.calendar.google.com", type: "room" },
        { name: "食堂/第三会議室", id: "c_1881v7mm3lr06i72lar5oliqu7seg@resource.calendar.google.com", type: "room" },
        { name: "食堂/第四会議室", id: "c_1887eobagp11ohunnvljr96o10vhm@resource.calendar.google.com", type: "room" },
        { name: "食堂/第五会議室", id: "c_1887f81iqru2qhpul4ish83rmmjsc@resource.calendar.google.com", type: "room" },
        { name: "品証棟/応接室", id: "c_18860s77j0vlohejl4ioo2bpi56mk@resource.calendar.google.com", type: "room" }
    ];

    // --- State Variables ---
    const jstFormatter = new Intl.DateTimeFormat('ja-JP', { timeZone: 'Asia/Tokyo', year: 'numeric', month: 'numeric', day: 'numeric' });
    const parts = jstFormatter.formatToParts(new Date());
    const year = parseInt(parts.find(p => p.type === 'year').value, 10);
    const month = parseInt(parts.find(p => p.type === 'month').value, 10);
    const day = parseInt(parts.find(p => p.type === 'day').value, 10);
    let selectedDate = new Date(year, month - 1, day);

    // --- Helper Functions ---
    function showLoading(isLoading) { if (loadingDiv) loadingDiv.style.display = isLoading ? 'block' : 'none'; }
    function showError(message) { if (errorDiv) { errorDiv.textContent = message; errorDiv.style.display = message ? 'block' : 'none'; } }
    function formatTime(date) {
        const h = String(date.getHours()).padStart(2, '0');
        const m = String(date.getMinutes()).padStart(2, '0');
        return `${h}:${m}`;
    }
    function formatEventTime(eventStart, eventEnd) {
        const options = { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' };
        const startTime = new Date(eventStart.dateTime || eventStart.date).toLocaleTimeString('ja-JP', options);
        const endTime = new Date(eventEnd.dateTime || eventEnd.date).toLocaleTimeString('ja-JP', options);
        return `${startTime}～${endTime}`;
    }

    // --- Auth Logic ---
    function handleAuthResponse(resp) {
        if (resp.error !== undefined) {
            if (resp.error === 'popup_closed' || resp.error === 'user_cancel' || resp.error === 'immediate_failed') {
                signInButton.style.display = 'block';
            } else {
                showError('認証エラー: ' + resp.error);
                console.error('Auth error:', resp);
                signInButton.style.display = 'block';
            }
            return;
        }
        signOutButton.style.display = 'block';
        signInButton.style.display = 'none';
        controlsWrapper.style.display = 'flex';
        fetchData();
        fetchAndPopulateCalendarList();
    }

    function trySilentSignIn() {
        if (gapiInited && gisInited) {
            tokenClient.callback = (resp) => {
                // Prevent silent sign-in errors from showing to the user
                if (resp.error !== undefined) {
                    if (resp.error !== 'immediate_failed') {
                        console.error('Silent sign-in failed:', resp);
                    }
                    signInButton.style.display = 'block'; // Show sign in button if silent auth fails
                } else {
                     handleAuthResponse(resp);
                }
            };
            tokenClient.requestAccessToken({ prompt: 'none' });
        }
    }

    signInButton.onclick = () => {
        if (gisInited && tokenClient) {
            tokenClient.callback = handleAuthResponse;
            tokenClient.requestAccessToken({ prompt: 'consent' });
        } else {
            showError('Google認証が準備できていません。ページを再読み込みしてください。');
        }
    };
    
    signOutButton.onclick = () => {
        const token = gapi.client.getToken();
        if (token !== null) {
            google.accounts.oauth2.revoke(token.access_token, () => {
                gapi.client.setToken('');
                dataDisplayArea.innerHTML = '';
                controlsWrapper.style.display = 'none';
                signInButton.style.display = 'block';
                signOutButton.style.display = 'none';
                showError('サインアウトしました。');
            });
        }
    };

    // --- Data Fetching & Calendar List ---
    async function fetchAndPopulateCalendarList() {
        try {
            const response = await gapi.client.calendar.calendarList.list();
            const calendars = response.result.items;
            targetCalendarSelect.innerHTML = ''; 
            let primaryCalendarId = null; 
            calendars.forEach((calendar) => {
                if (calendar.accessRole === 'owner' || calendar.accessRole === 'writer') {
                    const option = document.createElement('option');
                    option.value = calendar.id;
                    option.textContent = calendar.summary;
                    targetCalendarSelect.appendChild(option);
                    if (calendar.primary) {
                        primaryCalendarId = calendar.id;
                    }
                }
            });
            const desiredCalendarId = 'rteikabo0e4p6gkd6gdfdbvmf4@group.calendar.google.com';
            const desiredOptionExists = Array.from(targetCalendarSelect.options).some(opt => opt.value === desiredCalendarId);
            if (desiredOptionExists) {
                targetCalendarSelect.value = desiredCalendarId;
            } else if (primaryCalendarId) {
                targetCalendarSelect.value = primaryCalendarId;
            }
        } catch (err) {
            console.error("Error fetching calendar list", err);
            showError("カレンダーリストの取得に失敗しました。");
        }
    }

    async function fetchData() {
        if (!gapi.client.getToken()) { showError("Googleアカウントでサインインしてください。"); return; }
        showLoading(true); showError(''); dataDisplayArea.innerHTML = '';
        const timeMin = new Date(selectedDate);
        timeMin.setHours(0, 0, 0, 0);
        const timeMax = new Date(selectedDate);
        timeMax.setHours(23, 59, 59, 999);
        
        const idsToFetchDetails = resourceCalendarItems.map(item => item.id);
        if (idsToFetchDetails.length === 0) { showError("取得対象のリソースがありません。"); showLoading(false); return; }
        
        const allCalendarData = {};
        try {
            const batch = gapi.client.newBatch();
            idsToFetchDetails.forEach(calendarId => {
                batch.add(gapi.client.calendar.events.list({
                    'calendarId': calendarId,
                    'timeMin': timeMin.toISOString(),
                    'timeMax': timeMax.toISOString(),
                    'showDeleted': false,
                    'singleEvents': true,
                    'maxResults': 50,
                    'orderBy': 'startTime'
                }), {id: calendarId});
            });

            const batchResponse = await batch;
            const results = batchResponse.result;

            for (const calendarId in results) {
                const response = results[calendarId];
                if (response.status === 200 && response.result) {
                    allCalendarData[calendarId] = { items: response.result.items ? response.result.items.map(event => ({ id: event.id, summary: event.summary || '(タイトルなし)', start: event.start, end: event.end, creator: event.creator ? (event.creator.displayName || event.creator.email) : '(作成者不明)', organizer: event.organizer ? (event.organizer.displayName || event.organizer.email) : '(主催者不明)', attendees: event.attendees ? event.attendees.map(att => att.displayName || att.email) : [] })) : [] };
                } else {
                     console.error(`Error fetching events for ${calendarId}:`, response.result.error);
                     allCalendarData[calendarId] = { items: [], error: response.result.error };
                }
            }
            showLoading(false);
            renderDailyMatrixView(allCalendarData);
        } catch (err) {
            showLoading(false);
            console.error("A critical error occurred: ", err);
            showError(`重大なエラーが発生しました: ${err.message || JSON.stringify(err)}`);
        }
    }

    // --- Booking Modal Logic ---
    function populateTimeSelects(selectElement, selectedTime) {
        selectElement.innerHTML = '';
        const startHour = 8; const endHour = 19; const timeSlotInterval = 15;
        for (let h = startHour; h <= endHour; h++) {
            for (let m = 0; m < 60; m += timeSlotInterval) {
                if (h === endHour && m > 0) continue;
                const option = document.createElement('option');
                const time = new Date(selectedDate);
                time.setHours(h, m, 0, 0);
                option.value = time.toISOString();
                option.textContent = formatTime(time);
                if (formatTime(time) === formatTime(selectedTime)) {
                    option.selected = true;
                }
                selectElement.appendChild(option);
            }
        }
    }
    
    function openBookingModal(room, startTime) {
        const endTime = new Date(startTime.getTime() + 30 * 60000); // Default to 30 min booking
        bookingData = { room: room };
        populateTimeSelects(bookingStartTimeSelect, startTime);
        populateTimeSelects(bookingEndTimeSelect, endTime);
        bookingResourceName.textContent = room.name;
        eventTitleInput.value = '';
        bookingModal.style.display = 'block';
        bookingModalBackdrop.style.display = 'block';
    }

    function closeBookingModal() {
        bookingModal.style.display = 'none';
        bookingModalBackdrop.style.display = 'none';
    }

    async function createCalendarEvent() {
        const summary = eventTitleInput.value;
        if (!summary) { alert('会議名を入力してください。'); return; }
        const targetCalendarId = targetCalendarSelect.value;
        if (!targetCalendarId) { alert('作成先のカレンダーを選択してください。'); return; }
        
        const eventResource = {
            'summary': summary,
            'start': { 'dateTime': bookingStartTimeSelect.value, 'timeZone': 'Asia/Tokyo' },
            'end': { 'dateTime': bookingEndTimeSelect.value, 'timeZone': 'Asia/Tokyo' },
            'attendees': [{ 'email': bookingData.room.id }]
        };
        
        try {
            const response = await gapi.client.calendar.events.insert({
                'calendarId': targetCalendarId,
                'resource': eventResource,
                'sendUpdates': 'all'
            });
            
            closeBookingModal();
            setTimeout(() => { fetchData(); }, 2000); 

            if (confirm('予約が作成されました。\n続けて詳細情報（ゲストや添付資料など）を編集しますか？')) {
                const eventUrl = response.result.htmlLink;
                if (eventUrl) {
                    window.open(eventUrl, '_blank');
                } else {
                    alert('詳細情報の編集画面を開けませんでした。');
                }
            }
        } catch (err) {
            console.error('Error creating event:', err);
            alert(`予約の作成に失敗しました: ${err.result.error.message}`);
        }
    }
    
    cancelBookingBtn.onclick = closeBookingModal;
    bookingModalBackdrop.onclick = closeBookingModal;
    saveBookingBtn.onclick = createCalendarEvent;

    // --- Rendering Logic ---
    function renderDailyMatrixView(calendarsEventData) {
        dataDisplayArea.innerHTML = '';
        const dateDisplay = document.createElement('h3');
        const y = selectedDate.getFullYear();
        const m = selectedDate.getMonth() + 1;
        const d = selectedDate.getDate();
        const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][selectedDate.getDay()];
        dateDisplay.textContent = `${y}年${m}月${d}日 (${dayOfWeek}曜日)`;
        dateDisplay.style.textAlign = 'center';
        dateDisplay.style.color = '#333';
        dateDisplay.style.fontSize = '1.0em';
        dateDisplay.style.marginBottom = '5px';
        dataDisplayArea.appendChild(dateDisplay);

        const table = document.createElement('table'); table.id = 'dailyMatrixTable';
        const thead = table.createTHead();
        const headerRow1 = thead.insertRow(); headerRow1.className = 'header-row-1';
        const headerRow2 = thead.insertRow(); headerRow2.className = 'header-row-2';
        
        const thRoomHeader = document.createElement('th');
        thRoomHeader.rowSpan = 2;
        thRoomHeader.textContent = 'リソース';
        headerRow1.appendChild(thRoomHeader);
        
        const startHour = 8; const endHour = 19; const timeSlotInterval = 15;
        const slotsPerHour = 60 / timeSlotInterval;
        for (let h = startHour; h < endHour; h++) {
            const thHour = document.createElement('th');
            thHour.colSpan = slotsPerHour;
            thHour.textContent = `${String(h).padStart(2, '0')}:00`;
            headerRow1.appendChild(thHour);
            for (let m = 0; m < 60; m += timeSlotInterval) {
                const thMin = document.createElement('th');
                thMin.textContent = `:${String(m).padStart(2, '0')}`;
                headerRow2.appendChild(thMin);
            }
        }
        
        const tbody = table.createTBody();
        resourceCalendarItems.forEach((room, index) => {
            const roomRow = tbody.insertRow();
            if (index > 0 && room.type !== resourceCalendarItems[index - 1].type) { roomRow.classList.add('group-separator'); }
            if (room.type === 'car') { roomRow.classList.add('car-row'); }
            const tdRoomName = roomRow.insertCell();
            tdRoomName.textContent = room.name;
            tdRoomName.title = room.name;
            const roomData = calendarsEventData[room.id];
            
            const totalSlots = (endHour - startHour) * slotsPerHour;
            const slots = new Array(totalSlots).fill(null);

            if (roomData && roomData.items) {
                roomData.items.forEach(event => {
                    const eventStart = new Date(event.start.dateTime || event.start.date + 'T00:00:00');
                    const eventEnd = new Date(event.end.dateTime || event.end.date + 'T23:59:59');
                    const dayStart = new Date(selectedDate); dayStart.setHours(startHour, 0, 0, 0);
                    const dayEnd = new Date(selectedDate); dayEnd.setHours(endHour, 0, 0, 0);

                    const effectiveStart = eventStart > dayStart ? eventStart : dayStart;
                    const effectiveEnd = eventEnd < dayEnd ? eventEnd : dayEnd;

                    const startMinutes = (effectiveStart.getHours() - startHour) * 60 + effectiveStart.getMinutes();
                    const endMinutes = (effectiveEnd.getHours() - startHour) * 60 + effectiveEnd.getMinutes();
                    
                    const startIndex = Math.floor(startMinutes / timeSlotInterval);
                    const endIndex = Math.ceil(endMinutes / timeSlotInterval);
                    
                    for (let i = startIndex; i < endIndex; i++) {
                        if (i >= 0 && i < totalSlots && slots[i] === null) {
                            slots[i] = { event: event, isStart: (i === startIndex) };
                        }
                    }
                });
            }
            
            for (let i = 0; i < totalSlots; ) {
                const slotData = slots[i];
                if (slotData) {
                    let colspanCount = 1;
                    for (let j = i + 1; j < totalSlots; j++) {
                        if (slots[j] && slots[j].event.id === slotData.event.id) {
                            colspanCount++;
                        } else {
                            break;
                        }
                    }
                    const tdHourStatus = roomRow.insertCell();
                    tdHourStatus.colSpan = colspanCount;
                    
                    const eventDiv = document.createElement('div');
                    eventDiv.classList.add('event-bar');
                    eventDiv.textContent = `${formatEventTime(slotData.event.start, slotData.event.end)} ${slotData.event.summary}`;
                    tdHourStatus.appendChild(eventDiv);
                    
                    tdHourStatus.title = `会議時間: ${formatEventTime(slotData.event.start, slotData.event.end)}\n会議名: ${slotData.event.summary}\n作成者: ${slotData.event.creator || slotData.event.organizer || '(不明)'}`;
                    tdHourStatus.classList.add('matrix-cell-busy');
                    
                    i += colspanCount;
                } else {
                    const slotStartTime = new Date(selectedDate);
                    const h = startHour + Math.floor(i / slotsPerHour);
                    const m = (i % slotsPerHour) * timeSlotInterval;
                    slotStartTime.setHours(h, m, 0, 0);
                    const tdHourStatus = roomRow.insertCell();
                    tdHourStatus.classList.add('matrix-cell-available');
                    tdHourStatus.onclick = () => openBookingModal(room, slotStartTime);
                    i++;
                }
            }
        });
        dataDisplayArea.appendChild(table);
    }
    
    // --- UI Control Logic ---
    function navigateDay(offset) {
        selectedDate.setDate(selectedDate.getDate() + offset);
        const y = selectedDate.getFullYear();
        const m = String(selectedDate.getMonth() + 1).padStart(2, '0');
        const d = String(selectedDate.getDate()).padStart(2, '0');
        dailyDatePicker.value = `${y}-${m}-${d}`;
        if (gapi.client.getToken()) fetchData();
    }

    // --- App Initialization ---
    (function initializeApp() {
        if (resourceCalendarItems.length === 0) { showError("確認するリソースカレンダーが設定されていません。"); return; }
        
        const y = selectedDate.getFullYear();
        const m = String(selectedDate.getMonth() + 1).padStart(2, '0');
        const d = String(selectedDate.getDate()).padStart(2, '0');
        dailyDatePicker.value = `${y}-${m}-${d}`;
    
        dailyDatePicker.addEventListener('change', () => {
            const parts = dailyDatePicker.value.split('-').map(Number);
            selectedDate = new Date(parts[0], parts[1] - 1, parts[2]);
            if (gapi.client.getToken()) fetchData();
        });
        prevDayBtn.addEventListener('click', () => navigateDay(-1));
        nextDayBtn.addEventListener('click', () => navigateDay(1));
        document.addEventListener('gapiReady', trySilentSignIn);
        document.addEventListener('gisReady', trySilentSignIn);
    })();
});
