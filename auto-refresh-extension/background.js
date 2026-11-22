const ALARM_NAME = 'auto-refresh';
const INTERVAL_MINUTES = 20;

chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.get({ enabled: true }, (result) => {
    if (result.enabled) startAlarm();
  });
});

chrome.runtime.onStartup.addListener(() => {
  chrome.storage.local.get({ enabled: true }, (result) => {
    if (result.enabled) startAlarm();
  });
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === ALARM_NAME) {
    chrome.storage.local.get({ enabled: true }, (result) => {
      if (!result.enabled) return;
      chrome.tabs.query({}, (tabs) => {
        for (const tab of tabs) {
          if (tab.id && tab.url && /^https?:\/\//.test(tab.url)) {
            chrome.tabs.reload(tab.id);
          }
        }
      });
    });
  }
});

function startAlarm() {
  chrome.alarms.create(ALARM_NAME, { periodInMinutes: INTERVAL_MINUTES });
}

function stopAlarm() {
  chrome.alarms.clear(ALARM_NAME);
}

chrome.runtime.onMessage.addListener((message) => {
  if (message.action === 'toggle') {
    chrome.storage.local.get({ enabled: true }, (result) => {
      const newState = !result.enabled;
      chrome.storage.local.set({ enabled: newState }, () => {
        if (newState) startAlarm();
        else stopAlarm();
        chrome.runtime.sendMessage({ enabled: newState });
      });
    });
  }
});
