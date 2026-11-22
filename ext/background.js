chrome.action.onClicked.addListener(async (tab) => {
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        // Try to get the main content, fall back to body text
        const article = document.querySelector('article');
        const main = document.querySelector('main');
        const content = article || main || document.body;
        const metaDesc = document.querySelector('meta[name="description"]');
        // Try multiple meta tags for author
        let author = '';
        const authorSelectors = [
          'meta[name="citation_author"]',
          'meta[name="dc.creator"]',
          'meta[name="author"]',
          'meta[property="article:author"]',
        ];
        for (const sel of authorSelectors) {
          const el = document.querySelector(sel);
          if (el && el.content) { author = el.content; break; }
        }
        // Try multiple meta tags for publication date
        let pubDate = '';
        const dateSelectors = [
          'meta[name="citation_publication_date"]',
          'meta[name="dc.date"]',
          'meta[property="article:published_time"]',
          'meta[name="date"]',
        ];
        for (const sel of dateSelectors) {
          const el = document.querySelector(sel);
          if (el && el.content) { pubDate = el.content; break; }
        }
        return {
          text: content.innerText,
          url: window.location.href,
          title: document.title,
          description: metaDesc ? metaDesc.content : '',
          author: author,
          pub_date: pubDate
        };
      }
    });

    const data = results[0].result;

    const response = await fetch('http://localhost:9876/extract', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });

    const result = await response.json();

    chrome.action.setBadgeText({ text: 'OK' });
    chrome.action.setBadgeBackgroundColor({ color: '#4CAF50' });
    setTimeout(() => chrome.action.setBadgeText({ text: '' }), 1500);
  } catch (e) {
    chrome.action.setBadgeText({ text: 'ERR' });
    chrome.action.setBadgeBackgroundColor({ color: '#F44336' });
    setTimeout(() => chrome.action.setBadgeText({ text: '' }), 2000);
  }
});
