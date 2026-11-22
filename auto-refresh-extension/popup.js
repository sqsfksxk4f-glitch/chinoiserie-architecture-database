const statusEl = document.getElementById('status');
const toggleBtn = document.getElementById('toggle');

chrome.storage.local.get({ enabled: true }, (result) => {
  updateUI(result.enabled);
});

toggleBtn.addEventListener('click', () => {
  chrome.storage.local.get({ enabled: true }, (result) => {
    const newState = !result.enabled;
    chrome.storage.local.set({ enabled: newState }, () => {
      if (newState) {
        chrome.alarms.create('auto-refresh', { periodInMinutes: 20 });
      } else {
        chrome.alarms.clear('auto-refresh');
      }
      updateUI(newState);
    });
  });
});

function updateUI(enabled) {
  statusEl.textContent = enabled ? '状态：运行中' : '状态：已暂停';
  toggleBtn.textContent = enabled ? '关闭' : '开启';
}
