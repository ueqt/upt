
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
 * @description 在“Add”按钮前添加一个复选框，用于过滤已在 solution 中的内容
 */
function addFilterCheckbox() {
  const visibleDialog = getVisibleDialog();
  if (!visibleDialog) {
    const checkbox = document.getElementById('filter-checkbox');
    if (checkbox) checkbox.remove();
    const label = document.querySelector('label[for="filter-checkbox"]');
    if (label) label.remove();
    return;
  }

  const nextButton = visibleDialog.querySelector('button[aria-label="Next"]');
  const addButton = visibleDialog.querySelector('button[aria-label="Add"]');
  const tabList = visibleDialog.querySelector('[role="tablist"]');

  const checkbox = document.getElementById('filter-checkbox');
  const label = document.querySelector('label[for="filter-checkbox"]');

  const shouldShowCheckbox = addButton && tabList && !nextButton;

  if (shouldShowCheckbox) {
    if (!checkbox) {
      const newCheckbox = document.createElement('input');
      newCheckbox.type = 'checkbox';
      newCheckbox.id = 'filter-checkbox';
      newCheckbox.style.marginRight = '10px';

      const newLabel = document.createElement('label');
      newLabel.htmlFor = 'filter-checkbox';
      newLabel.innerText = 'Filter items already in solution';
      newLabel.style.marginRight = '10px';

      addButton.parentNode.insertBefore(newCheckbox, addButton);
      addButton.parentNode.insertBefore(newLabel, addButton);

      newCheckbox.addEventListener('change', (e) => {
        filterSolutionItems(e.target.checked, visibleDialog);
      });
    }
  } else {
    if (checkbox) checkbox.remove();
    if (label) label.remove();
  }
}

/**
 * @description 根据复选框的状态过滤 solution 中的项目
 * @param {boolean} isChecked 复选框是否选中
 * @param {HTMLElement} context 在哪个对话框中进行过滤
 */
function filterSolutionItems(isChecked, context) {
  if (!context) return;

  const grid = context.querySelector('div[role="grid"]');
  if (!grid) return;

  const rows = Array.from(grid.querySelectorAll('div[role="row"]'));
  rows.forEach(row => {
    const managedCell = row.querySelector('div[data-automation-key="managed"]');
    if (managedCell) {
      const isManaged = managedCell.innerText.toLowerCase() === 'yes';
      if (isChecked && isManaged) {
        row.style.display = 'none';
      } else {
        row.style.display = '';
      }
    }
  });
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
    console.table(window.allApiItems);
    alert(`Successfully fetched ${allItems.length} items. See console for data.`);

  } catch (error) {
    console.error('Error in fetchAllItems:', error);
  } finally {
    scrapeButton.textContent = 'Fetch All Items (API)';
    scrapeButton.disabled = false;
  }
}

const observer = new MutationObserver(() => {
  observer.disconnect();
  addFilterCheckbox();
  displayVersion();
  observer.observe(document.body, {
    childList: true,
    subtree: true,
  });
});

observer.observe(document.body, {
  childList: true,
  subtree: true,
});


addFilterCheckbox();
displayVersion();
injectScript();