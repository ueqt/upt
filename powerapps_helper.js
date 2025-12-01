
/**
 * @description 查找当前可见的对话框
 * @returns {HTMLElement | null} 可见的对话框元素，如果找不到则返回 null
 */
function getVisibleDialog() {
  const dialogs = document.querySelectorAll('div[role="dialog"]');
  for (const dialog of dialogs) {
    if (!dialog.closest('[aria-hidden="true"]')) {
      return dialog;
    }
  }
  return null;
}

/**
 * @description 在 `ppuxOfficeHeaderCenterRegion` 元素中显示版本号和抓取按钮
 */
function displayVersion() {
  const version = chrome.runtime.getManifest().version;
  const headerCenterRegion = document.getElementById('ppuxOfficeHeaderCenterRegion');

  if (headerCenterRegion) {
    headerCenterRegion.style.display = 'flex';
    headerCenterRegion.style.justifyContent = 'center';
    headerCenterRegion.style.alignItems = 'center';

    let versionBanner = document.getElementById('extension-version-banner');
    if (!versionBanner) {
      versionBanner = document.createElement('div');
      versionBanner.id = 'extension-version-banner';
      versionBanner.textContent = `Helper v${version}`;
      versionBanner.style.cssText = `
        color: #605e5c;
        background-color: #f3f2f1;
        padding: 4px 12px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
        line-height: 1;
        text-align: center;
        margin-right: 10px;
      `;
      headerCenterRegion.appendChild(versionBanner);
    }

    let scrapeButton = document.getElementById('scrape-button');
    if (!scrapeButton) {
      scrapeButton = document.createElement('button');
      scrapeButton.id = 'scrape-button';
      scrapeButton.style.cssText = `
        background-color: #0078d4;
        color: white;
        border: none;
        padding: 4px 12px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
        cursor: pointer;
      `;
      headerCenterRegion.appendChild(scrapeButton);
      scrapeButton.addEventListener('click', fetchAllItems);
    }
    scrapeButton.textContent = 'Fetch All Items (API)';
  }
}

/**
 * @description 显示一个全局加载遮罩
 */
function showLoadingOverlay() {
  const overlay = document.createElement('div');
  overlay.id = 'loading-overlay';
  overlay.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0, 0, 0, 0.5);
    z-index: 9999;
    display: flex;
    justify-content: center;
    align-items: center;
  `;

  const spinner = document.createElement('div');
  spinner.style.cssText = `
    border: 8px solid #f3f3f3;
    border-top: 8px solid #3498db;
    border-radius: 50%;
    width: 60px;
    height: 60px;
    animation: spin 2s linear infinite;
  `;

  overlay.appendChild(spinner);
  document.body.appendChild(overlay);

  const style = document.createElement('style');
  style.textContent = `
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  `;
  document.head.appendChild(style);
}

/**
 * @description 隐藏加载遮罩
 */
function hideLoadingOverlay() {
  const overlay = document.getElementById('loading-overlay');
  if (overlay) {
    overlay.remove();
  }
}

/**
 * @description 在对话框中添加“Filter items already in solution”复选框
 */
function addFilterCheckbox() {
  const dialog = getVisibleDialog();
  if (dialog && !dialog.querySelector('#filter-checkbox')) {
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'filter-checkbox';
    checkbox.style.marginRight = '8px';

    const label = document.createElement('label');
    label.htmlFor = 'filter-checkbox';
    label.textContent = 'Filter items already in solution';
    label.style.display = 'flex';
    label.style.alignItems = 'center';
    label.style.cursor = 'pointer';
    label.style.fontWeight = '600';
    label.style.marginBottom = '10px';

    label.insertBefore(checkbox, label.firstChild);

    const titleElement = dialog.querySelector('h1');
    if (titleElement && titleElement.parentElement) {
      titleElement.parentElement.insertBefore(label, titleElement.nextSibling);
    }

    checkbox.addEventListener('change', (event) => {
      filterSolutionItems(event.target.checked);
    });
  }
}

/**
 * @description 根据复选框状态过滤解决方案中的项目
 * @param {boolean} isChecked - 复选框是否被选中
 */
function filterSolutionItems(isChecked) {
  if (!window.allApiItems) {
    alert('Please click "Fetch All Items (API)" first to get the data for filtering.');
    const checkbox = document.getElementById('filter-checkbox');
    if (checkbox) checkbox.checked = false;
    return;
  }

  const gridRows = document.querySelectorAll('div[role="grid"] div[role="row"]');
  const apiItemsMap = new Map(window.allApiItems.map(item => [item.msdyn_displayname, item]));

  gridRows.forEach(row => {
    const nameElement = row.querySelector('div[data-automation-id="solution-component-name"]');
    if (nameElement) {
      const displayName = nameElement.textContent.trim();
      const apiItem = apiItemsMap.get(displayName);

      if (isChecked && apiItem && apiItem.msdyn_ismanaged) {
        row.style.display = 'none';
      } else {
        row.style.display = '';
      }
    }
  });
}

/**
 * @description Injects a script into the page to intercept fetch requests and capture credentials.
 */
function injectScript() {
  const script = document.createElement('script');
  script.src = chrome.runtime.getURL('injected_script.js');
  (document.head || document.documentElement).appendChild(script);
  script.onload = () => script.remove();
}

/**
 * @description Listens for messages from the injected script containing captured credentials.
 */
window.addEventListener('message', (event) => {
  if (event.source === window && event.data.type === 'CREDENTIALS_CAPTURED') {
    const { bearerToken, instanceUrl } = event.data.payload;
    console.log('[Content Script] Received credentials:', { instanceUrl, bearerToken });
    if (bearerToken && instanceUrl) {
      localStorage.setItem('bearerToken', bearerToken);
      localStorage.setItem('instanceUrl', instanceUrl);
      console.log('Credentials captured and stored.');
    }
  }
});

/**
 * @description Fetches all components for a given solution from the PowerApps API.
 * It retrieves credentials from localStorage and handles pagination.
 */
async function fetchAllItems() {
  showLoadingOverlay();
  const scrapeButton = document.getElementById('scrape-button');
  scrapeButton.textContent = 'Fetching...';
  scrapeButton.disabled = true;

  try {
    // 1. Get credentials from localStorage
    const bearerToken = localStorage.getItem('bearerToken');
    const instanceUrl = localStorage.getItem('instanceUrl');

    if (!bearerToken || !instanceUrl) {
      alert('Credentials not found. Please navigate around the PowerApps interface to capture them automatically.');
      return; // The finally block will reset the button
    }

    // 2. From URL get Solution ID
    const params = new URLSearchParams(window.location.search);
    let solutionId = params.get('id');
    if (!solutionId) {
      const urlParts = window.location.pathname.split('/');
      const solutionIndex = urlParts.indexOf('solutions');
      if (solutionIndex !== -1 && urlParts.length > solutionIndex + 1) {
        solutionId = urlParts[solutionIndex + 1];
      }
    }
    if (!solutionId) {
      alert('Solution ID not found in URL. Please navigate to a solution page.');
      throw new Error('Solution ID not found');
    }
    solutionId = solutionId.replace(/[{}]/g, "");

    // 3. Make API request
    let allItems = [];
    let nextLink = `${instanceUrl}/api/data/v9.0/msdyn_solutioncomponentsummaries?$filter=(msdyn_solutionid eq ${solutionId})&$orderby=msdyn_displayname asc`;

    while (nextLink) {
      const response = await fetch(nextLink, {
        headers: {
          'Authorization': bearerToken,
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          'Accept': 'application/json',
          'Content-Type': 'application/json; charset=utf-8'
        }
      });

      if (!response.ok) {
        if (response.status === 401) {
          alert('The provided Bearer Token is invalid or has expired. Please refresh the page to capture a new one.');
          localStorage.removeItem('bearerToken'); // Clear the expired token
        } else {
          const errorText = await response.text();
          alert(`HTTP error! status: ${response.status}, message: ${errorText}`);
        }
        throw new Error(`API request failed with status ${response.status}`);
      }
      const data = await response.json();
      if (data.value) {
        allItems = allItems.concat(data.value);
      }
      nextLink = data['@odata.nextLink'];
    }

    window.allApiItems = allItems; // Store all items globally as per instruction
    
    // Store msdyn_objectid in localStorage for the injected script
    const objectIds = allItems.map(item => item.msdyn_objectid);
    localStorage.setItem('solutionComponentObjectIds', JSON.stringify(objectIds));

    console.table(allItems.map(item => ({ 
      msdyn_objectid: item.msdyn_objectid, 
      msdyn_displayname: item.msdyn_displayname 
    })));
    alert(`Successfully fetched ${allItems.length} items. See console for data.`);

  } catch (error) {
    console.error('Error fetching items:', error);
    alert('An error occurred while fetching items. See the console for details.');
  } finally {
    hideLoadingOverlay();
    scrapeButton.textContent = 'Fetch All Items (API)';
    scrapeButton.disabled = false;
  }
}

const observer = new MutationObserver(() => {
  observer.disconnect();
  displayVersion();
  addFilterCheckbox();
  observer.observe(document.body, {
    childList: true,
    subtree: true,
  });
});

observer.observe(document.body, {
  childList: true,
  subtree: true,
});


displayVersion();
injectScript();