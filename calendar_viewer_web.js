// --- GAPI と GIS の初期化 (グローバルスコープ) ---
let tokenClient;
let gapiInited = false;
let gisInited = false;

// ★★★ GCPで作成したウェブアプリケーション用の情報を設定 ★★★
const API_KEY = 'AIzaSyCpRjx_lkdpcp-eePb-_psrh5MUB-T06aA';
const CLIENT_ID = '976002357617-j8i08ahpphr1frt1ko7tltvfbp847alk.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.readonly';

function gapiLoaded() {
    gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
    await gapi.client.init({
        apiKey: API_KEY,
        discoveryDocs: ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"],
    });
    gapiInited = true;
    document.dispatchEvent(new Event('gapiReady'));
}

function gisLoaded() {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CLIENT_ID,
        scope: SCOPES,
        callback: '',
    });
    gisInited = true;
    document.dispatchEvent(new Event('gisReady'));
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
    // ★★★ 滋賀工場の会議室リスト ★★★
    const resourceCalendarItems = [
        { name: "滋賀工場-1-4号館第一応接室 (6)", id: "c_1884ojmrqcv0iitpiq88jj7i96msk@resource.calendar.google.com", type: "room" },
        { name: "滋賀工場-1-4号館第一会議室 (20)", id: "c_188a7oshg5c5ijatnn3hs9cc4meie@resource.calendar.google.com", type: "room" },
        { name: "滋賀工場-1-4号館第二会議室 (4)", id: "c_1881acanrq6f0jitl0grsoqhfg1b6@resource.calendar.google.com", type: "room" },
        { name: "滋賀工場-1-階-食堂第三会議室 (4)", id: "c_1881v7mm3lr06i72lar5oliqu7seg@resource.calendar.google.com", type: "room" },
        { name: "滋賀工場-1-階-食堂第四会議室 (4)", id: "c_1887eobagp11ohunnvljr96o10vhm@resource.calendar.google.com", type: "room" },
        { name: "滋賀工場-1-階-食堂第五会議室 (4)", id: "c_1887f81iqru2qhpul4ish83rmmjsc@resource.calendar.google.com", type: "room" },
        { name: "滋賀工場-1-品証棟応接室 (4)", id: "c_18860s77j0vlohejl4ioo2bpi56mk@resource.calendar.google.com", type: "room" }
    ];


// --- State Variables ---
    // ★★★ 修正箇所: 日本時間での今日の日付を取得 ★★★
    // 'Asia/Tokyo'タイムゾーンでの日付部分を取得するためのフォーマッターを作成
    const jstFormatter = new Intl.DateTimeFormat('ja-JP', {
        timeZone: 'Asia/Tokyo',
        year: 'numeric',
        month: 'numeric',
        day: 'numeric'
    });

    // 現在時刻をフォーマットして、日本の年・月・日の部分を取得
    const parts = jstFormatter.formatToParts(new Date());
    const year = parseInt(parts.find(p => p.type === 'year').value, 10);
    const month = parseInt(parts.find(p => p.type === 'month').value, 10);
    const day = parseInt(parts.find(p => p.type === 'day').value, 10);

    // 取得した年・月・日を使用してDateオブジェクトを作成
    // Dateコンストラクタの月は0から始まるため、-1 する
    let selectedDate = new Date(year, month - 1, day);

    // --- Helper Functions ---
    function showLoading(isLoading) { if (loadingDiv) loadingDiv.style.display = isLoading ? 'block' : 'none'; }
    function showError(message) { if (errorDiv) { errorDiv.textContent = message; errorDiv.style.display = message ? 'block' : 'none'; } }
    function getMonday(d) { d = new Date(d); const day = d.getDay(), diff = d.getDate() - day + (day === 0 ? -6 : 1); return new Date(d.setDate(diff)); }
    function formatDate(date) { const y = date.getFullYear(), m = ('0' + (date.getMonth() + 1)).slice(-2), d = ('0' + date.getDate()).slice(-2); return `${y}/${m}/${d}`; }
    function formatTime(date) {
        const h = String(date.getHours()).padStart(2, '0');
        const m = String(date.getMinutes()).padStart(2, '0');
        return `${h}:${m}`;
    }
    function formatEventTime(eventStart, eventEnd) {
        const options = { hour: '2-digit', minute: '2-digit' };
        const startTime = new Date(eventStart.dateTime || eventStart.date).toLocaleTimeString('ja-JP', options);
        const endTime = new Date(eventEnd.dateTime || eventEnd.date).toLocaleTimeString('ja-JP', options);
        return `${startTime}～${endTime}`;
    }

    // --- Auth Logic ---
    function handleAuthResponse(resp) {
        if (resp.error !== undefined) {
            if (resp.error === 'popup_closed' || resp.error === 'user_cancel' || resp.error === 'immediate_failed') { signInButton.style.display = 'block'; } else { showError('認証中にエラーが発生しました: ' + JSON.stringify(resp)); signInButton.style.display = 'block'; }
            return;
        }
        signOutButton.style.display = 'block'; signInButton.style.display = 'none'; controlsWrapper.style.display = 'flex';
        gapi.client.setToken({ access_token: resp.access_token });
        fetchData();
        fetchAndPopulateCalendarList();
    }
    function trySilentSignIn() { if (gapiInited && gisInited) { tokenClient.callback = handleAuthResponse; tokenClient.requestAccessToken({ prompt: 'none' }); } }
    signInButton.onclick = () => { tokenClient.callback = handleAuthResponse; tokenClient.requestAccessToken({ prompt: 'consent' }); };
    signOutButton.onclick = () => {
        const token = gapi.client.getToken();
        if (token !== null) {
            google.accounts.oauth2.revoke(token.access_token, () => {
                gapi.client.setToken(''); dataDisplayArea.innerHTML = ''; controlsWrapper.style.display = 'none';
                signInButton.style.display = 'block'; signOutButton.style.display = 'none'; showError('サインアウトしました。');
            });
        }
    };

    // --- Data Fetching & Calendar List ---
    async function fetchAndPopulateCalendarList() {
    try {
        const response = await gapi.client.calendar.calendarList.list();
        const calendars = response.result.items;
        targetCalendarSelect.innerHTML = ''; // 既存の選択肢をクリア

        let primaryCalendarId = null; // マイカレンダーのIDを一時的に保存する変数

        // 手順1：まず、すべてのカレンダーを選択肢に追加する
        calendars.forEach((calendar) => {
            if (calendar.accessRole === 'owner' || calendar.accessRole === 'writer') {
                const option = document.createElement('option');
                option.value = calendar.id;
                option.textContent = calendar.summary;
                targetCalendarSelect.appendChild(option);

                // マイカレンダーのIDを覚えておく
                if (calendar.primary) {
                    primaryCalendarId = calendar.id;
                }
            }
        });

        // 手順2：リスト作成後、初期値にしたいカレンダーのIDを定義
        const desiredCalendarId = 'rteikabo0e4p6gkd6gdfdbvmf4@group.calendar.google.com';

        // 手順3：目的のIDが選択肢の中に存在するか確認
        const desiredOptionExists = Array.from(targetCalendarSelect.options).some(opt => opt.value === desiredCalendarId);

        if (desiredOptionExists) {
            // 存在すれば、それを初期値として設定
            targetCalendarSelect.value = desiredCalendarId;
        } else if (primaryCalendarId) {
            // 手順4：存在しない場合のみ、予備としてマイカレンダーを設定
            targetCalendarSelect.value = primaryCalendarId;
        }

    } catch (err) {
        console.error("Error fetching calendar list", err);
    }
}

    async function fetchData() {
        if (!gapi.client.getToken()) { showError("Googleアカウントでサインインしてください。"); return; }
        showLoading(true); showError(''); dataDisplayArea.innerHTML = '';
        const timeMinISO = new Date(new Date(selectedDate).setHours(0, 0, 0, 0)).toISOString();
        const timeMaxISO = new Date(new Date(selectedDate).setHours(23, 59, 59, 999)).toISOString();
        const idsToFetchDetails = resourceCalendarItems.map(item => item.id);
        if (idsToFetchDetails.length === 0) { showError("取得対象のリソースがありません。"); showLoading(false); return; }
        const allCalendarData = {};
        try {
            const requests = idsToFetchDetails.map(calendarId => gapi.client.calendar.events.list({ 'calendarId': calendarId, 'timeMin': timeMinISO, 'timeMax': timeMaxISO, 'showDeleted': false, 'singleEvents': true, 'maxResults': 50, 'orderBy': 'startTime' }));
            const responses = await Promise.allSettled(requests);
            responses.forEach((response, index) => {
                const calendarId = idsToFetchDetails[index];
                if (response.status === 'fulfilled' && response.value.result) {
                     allCalendarData[calendarId] = { items: response.value.result.items ? response.value.result.items.map(event => ({ id: event.id, summary: event.summary || '(タイトルなし)', start: event.start, end: event.end, creator: event.creator ? (event.creator.displayName || event.creator.email) : '(作成者不明)', organizer: event.organizer ? (event.organizer.displayName || event.organizer.email) : '(主催者不明)', attendees: event.attendees ? event.attendees.map(att => att.displayName || att.email) : [] })) : [] };
                } else {
                    console.error(`Error fetching events for ${calendarId}:`, response.reason || response.value);
                    allCalendarData[calendarId] = { items: [], errorDetails: response.reason || response.value };
                }
            });
            showLoading(false); renderDailyMatrixView(allCalendarData);
        } catch (err) {
            showLoading(false); console.error("A critical error occurred: ", err); showError(`重大なエラーが発生しました: ${err.message || JSON.stringify(err)}`);
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
        const endTime = new Date(startTime.getTime() + 30 * 60000);
        bookingData = { room: room };
        populateTimeSelects(bookingStartTimeSelect, startTime);
        populateTimeSelects(bookingEndTimeSelect, endTime);
        bookingResourceName.textContent = room.name;
        eventTitleInput.value = '';
        bookingModal.style.display = 'block';
        bookingModalBackdrop.style.display = 'block';
    }
    function closeBookingModal() { bookingModal.style.display = 'none'; bookingModalBackdrop.style.display = 'none'; }
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
        // 予定作成APIのレスポンスを受け取るように変更
        const response = await gapi.client.calendar.events.insert({ 'calendarId': targetCalendarId, 'resource': eventResource, 'sendUpdates': 'all' });

        // 先にモーダルを閉じる
        closeBookingModal();

        // 予約作成後に少し待ってからデータを再取得
        setTimeout(() => {
            fetchData();
        }, 3000); // 3000ミリ秒 = 3秒待つ

        // confirmダイアログで、詳細情報を編集するかユーザーに確認
        const openDetails = confirm('予約が作成されました。\n続けて詳細情報（ゲストや添付資料など）を編集しますか？');

        // 「はい」(OK)が押された場合
        if (openDetails) {
            // レスポンスからイベントのURL(htmlLink)を取得
            const eventUrl = response.result.htmlLink;
            if (eventUrl) {
                // 新しいタブでGoogleカレンダーのイベント編集画面を開く
                window.open(eventUrl, '_blank');
            } else {
                alert('詳細情報の編集画面を開けませんでした。');
            }
        }
    } catch (err) { console.error('Error creating event:', err); alert(`予約の作成に失敗しました: ${err.result.error.message}`); }
}
    cancelBookingBtn.onclick = closeBookingModal;
    bookingModalBackdrop.onclick = closeBookingModal;
    saveBookingBtn.onclick = createCalendarEvent;

    // --- Rendering Logic ---
    function renderDailyMatrixView(calendarsEventData) {
        dataDisplayArea.innerHTML = '';

        // 表示する日付の要素を作成
        const dateDisplay = document.createElement('h3');

        // selectedDate変数から年月日と曜日を取得
        const y = selectedDate.getFullYear();
        const m = selectedDate.getMonth() + 1;
        const d = selectedDate.getDate();
        const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][selectedDate.getDay()];

        // テキストとスタイルを設定
        dateDisplay.textContent = `${y}年${m}月${d}日 (${dayOfWeek}曜日)`;
        dateDisplay.style.textAlign = 'center';
        dateDisplay.style.color = '#333';
        dateDisplay.style.fontSize = '1.0em';
        dateDisplay.style.marginBottom = '5px';

        // マトリクス表示エリア (dataDisplayArea) の先頭に日付を追加
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
            thHour.classList.add('hour-header');
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
            if (index > 0 && room.type === 'car' && resourceCalendarItems[index - 1].type === 'room') { roomRow.classList.add('group-separator'); }
            if (room.type === 'car') { roomRow.classList.add('car-row'); }
            const tdRoomName = roomRow.insertCell();
            tdRoomName.textContent = room.name;
            tdRoomName.title = room.name;
            const roomData = calendarsEventData[room.id];

            const totalSlots = (endHour - startHour) * slotsPerHour;
            const slots = new Array(totalSlots).fill(null);
            if (roomData && roomData.items) {
                roomData.items.forEach(event => {
                    const eventStart = new Date(event.start.dateTime || event.start.date);
                    const eventEnd = new Date(event.end.dateTime || event.end.date);
                    const dayStart = new Date(selectedDate); dayStart.setHours(startHour, 0, 0, 0);
                    const dayEnd = new Date(selectedDate); dayEnd.setHours(endHour, 0, 0, 0);

                    const effectiveStart = eventStart > dayStart ? eventStart : dayStart;
                    const effectiveEnd = eventEnd < dayEnd ? eventEnd : dayEnd;

                    const startMinutes = (effectiveStart.getHours() - startHour) * 60 + effectiveStart.getMinutes();
                    const endMinutes = (effectiveEnd.getHours() - startHour) * 60 + effectiveEnd.getMinutes();

                    const startIndex = Math.floor(startMinutes / timeSlotInterval);
                    const endIndex = Math.ceil(endMinutes / timeSlotInterval);

                    for (let i = startIndex; i < endIndex; i++) {
                        if (i >= 0 && i < totalSlots) {
                            if (slots[i] === null) {
                                slots[i] = { event: event, isStart: (i === startIndex) };
                            }
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

                    let titleDetails = `会議時間: ${formatEventTime(slotData.event.start, slotData.event.end)}\n会議名: ${slotData.event.summary}\n作成者: ${slotData.event.creator || slotData.event.organizer || '(不明)'}\nゲスト: ${slotData.event.attendees && slotData.event.attendees.length > 0 ? slotData.event.attendees.join(', ') : "なし"}`;
                    tdHourStatus.title = titleDetails;
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
    // selectedDate変数の日付を更新
    selectedDate.setDate(selectedDate.getDate() + offset);

    // Dateオブジェクトから 'YYYY-MM-DD' 形式の文字列を生成
    const y = selectedDate.getFullYear();
    const m = String(selectedDate.getMonth() + 1).padStart(2, '0');
    const d = String(selectedDate.getDate()).padStart(2, '0');

    // タイムゾーン問題のある valueAsDate の代わりに、value プロパティに文字列を直接設定
    dailyDatePicker.value = `${y}-${m}-${d}`;

    // 更新された日付でデータを取得
    if (gapi.client.getToken()) fetchData();
}

    // --- App Initialization ---
    (function initializeApp() {
        if (resourceCalendarItems.length === 0) { showError("確認するリソースカレンダーが設定されていません。"); return; }
        // selectedDateオブジェクトから 'YYYY-MM-DD' 形式の文字列を生成
    const y = selectedDate.getFullYear();
    const m = String(selectedDate.getMonth() + 1).padStart(2, '0');
    const d = String(selectedDate.getDate()).padStart(2, '0');
    dailyDatePicker.value = `${y}-${m}-${d}`;

    // changeイベントリスナーも、valueプロパティから日付を読み取るように修正
    dailyDatePicker.addEventListener('change', () => {
    // 'YYYY-MM-DD'文字列を分割して、ローカルタイムゾーンのDateオブジェクトを再生成
    const parts = dailyDatePicker.value.split('-').map(Number);
    selectedDate = new Date(parts[0], parts[1] - 1, parts[2]);
    if (gapi.client.getToken()) fetchData();
});
        prevDayBtn.addEventListener('click', () => navigateDay(-1));
        nextDayBtn.addEventListener('click', () => navigateDay(1));
        document.addEventListener('gapiReady', trySilentSignIn);
        document.addEventListener('gisReady', trySilentSignIn);
        trySilentSignIn();
    })();
});