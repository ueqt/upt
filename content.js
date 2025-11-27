// Global state
let allRowsMap = new Map();
let listProcessed = false;
let lastMaxIndex = -1;

function addExpandButtons() {
  const grid = document.querySelector('div[role="grid"]');
  if (!grid) {
    return;
  }

  const rows = Array.from(grid.querySelectorAll('div[role="row"]'));

  // Ensure all rows are visible by default unless explicitly hidden by this script
  rows.forEach(row => {
    if (row.style.display === 'none' && !row.dataset.customHidden) {
      row.style.display = '';
    }
  });

  const headerRow = rows.find(row => row.querySelector('div[role="columnheader"]'));
  if (headerRow && !headerRow.querySelector('.custom-expand-column-header')) {
    const newHeaderCell = document.createElement('div');
    newHeaderCell.setAttribute('role', 'columnheader');
    newHeaderCell.className = 'custom-expand-column-header';
    newHeaderCell.style.width = '40px';
    newHeaderCell.style.minWidth = '40px';
    headerRow.prepend(newHeaderCell);
  }

  rows.forEach((row, index) => {
    if (row.querySelector('div[role="columnheader"]')) return;

    let customCell = row.querySelector('.custom-expand-cell');
    if (!customCell) {
        customCell = document.createElement('div');
        customCell.setAttribute('role', 'gridcell');
        customCell.className = 'custom-expand-cell';
        customCell.style.width = '40px';
        customCell.style.minWidth = '40px';
        customCell.style.display = 'flex';
        customCell.style.alignItems = 'center';
        customCell.style.justifyContent = 'center';
        row.prepend(customCell);
    }

    let shouldHaveButton = false;
    const nativeExpandButton = row.querySelector('button[aria-expanded]');
    const nextRow = rows[index + 1];

    if (nativeExpandButton && nextRow) {
        const currentRowContentWrapper = row.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');
        const nextRowContentWrapper = nextRow.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');

        if (currentRowContentWrapper && nextRowContentWrapper && nextRowContentWrapper.querySelector('img')) {
          if(currentRowContentWrapper.querySelector('img')) {
            const currentRowPadding = parseFloat(window.getComputedStyle(currentRowContentWrapper.firstChild).paddingLeft);
            const nextRowPadding = parseFloat(window.getComputedStyle(nextRowContentWrapper.firstChild).paddingLeft);
            shouldHaveButton = (nextRowPadding > currentRowPadding);
          } else {
            shouldHaveButton = true;
          }
        }
    }

    const existingButton = customCell.querySelector('.custom-expand-button');

    if (shouldHaveButton) {
      if (!existingButton) {
        const newButton = document.createElement('button');
        newButton.textContent = '-'; // Default to expanded
        newButton.className = 'custom-expand-button';
        newButton.style.cssText = 'border: 1px solid #ccc; background-color: #f0f0f0; cursor: pointer; min-width: 20px; height: 20px; line-height: 18px; padding: 0;';

        newButton.addEventListener('click', (e) => {
          e.stopPropagation();
          
          const isExpanded = newButton.textContent === '-';
          newButton.textContent = isExpanded ? '+' : '-';

          const currentRow = newButton.closest('div[role="row"]');
          const allRows = Array.from(grid.querySelectorAll('div[role="row"]'));
          const currentRowIndex = allRows.findIndex(r => r === currentRow);

          const currentRowContentWrapper = currentRow.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');
          if (!currentRowContentWrapper) return;

          const hasImage = currentRowContentWrapper.querySelector('img');

          if (!hasImage) {
            // Logic for rows without images
            for (let i = currentRowIndex + 1; i < allRows.length; i++) {
              const nextRow = allRows[i];
              const nextRowContentWrapper = nextRow.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');
              if (nextRowContentWrapper && !nextRowContentWrapper.querySelector('img')) {
                break;
              }
              if (isExpanded) { // Collapse
                nextRow.parentElement.style.display = 'none';
                nextRow.dataset.customHidden = 'true';
              } else { // Expand
                nextRow.parentElement.style.display = '';
                delete nextRow.dataset.customHidden;
              }
            }
          } else {
            // Logic for rows with images (based on padding)
            if (!currentRowContentWrapper.firstChild) return;
            const basePadding = parseFloat(window.getComputedStyle(currentRowContentWrapper.firstChild).paddingLeft);

            if (isExpanded) { // Collapse
              for (let i = currentRowIndex + 1; i < allRows.length; i++) {
                const nextRow = allRows[i];
                const nextRowContentWrapper = nextRow.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');
                if (!nextRowContentWrapper || !nextRowContentWrapper.firstChild) continue;
                const nextPadding = parseFloat(window.getComputedStyle(nextRowContentWrapper.firstChild).paddingLeft);

                if (nextPadding > basePadding) {
                  nextRow.parentElement.style.display = 'none';
                  nextRow.dataset.customHidden = 'true';
                } else {
                  break;
                }
              }
            } else { // Expand
              for (let i = currentRowIndex + 1; i < allRows.length; i++) {
                const nextRow = allRows[i];
                const nextRowContentWrapper = nextRow.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');
                if (!nextRowContentWrapper || !nextRowContentWrapper.firstChild) continue;
                const nextPadding = parseFloat(window.getComputedStyle(nextRowContentWrapper.firstChild).paddingLeft);

                if (nextPadding > basePadding) {
                  nextRow.parentElement.style.display = '';
                  delete nextRow.dataset.customHidden;
                  
                  const childButton = nextRow.querySelector('.custom-expand-button');
                  if (childButton && childButton.textContent === '+') {
                    const childsPadding = nextPadding;
                    let j = i + 1;
                    while (j < allRows.length) {
                      const subRow = allRows[j];
                      const subRowContentWrapper = subRow.querySelector('div[role="gridcell"]:not(.custom-expand-cell) > div:first-child');
                      if (!subRowContentWrapper || !subRowContentWrapper.firstChild) {
                          j++;
                          continue;
                      }
                      const subPadding = parseFloat(window.getComputedStyle(subRowContentWrapper.firstChild).paddingLeft);
                      if (subPadding > childsPadding) {
                        j++;
                      } else {
                        break;
                      }
                    }
                    i = j - 1;
                  }
                } else {
                  break;
                }
              }
            }
          }
          window.dispatchEvent(new Event('resize'));
        });
        customCell.appendChild(newButton);
      }
      // If button exists, do nothing, preserving its state.
    } else {
      if (existingButton) {
        existingButton.remove();
      }
    }
  });
}

function displayVersion() {
    const version = chrome.runtime.getManifest().version;
    const headerCenterRegion = document.getElementById('ppuxOfficeHeaderCenterRegion');

    if (headerCenterRegion) {
        // Ensure the container is a flex container to allow centering
        headerCenterRegion.style.display = 'flex';
        headerCenterRegion.style.justifyContent = 'center';
        headerCenterRegion.style.alignItems = 'center';

        let versionBanner = document.getElementById('extension-version-banner');
        if (!versionBanner) {
            versionBanner = document.createElement('div');
            versionBanner.id = 'extension-version-banner';
            headerCenterRegion.appendChild(versionBanner);
        }
        versionBanner.textContent = `Helper v${version}`;
        versionBanner.style.cssText = `
            color: #605e5c; /* Fluent UI neutral secondary text color */
            background-color: #f3f2f1; /* Fluent UI neutralLighter background */
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 600;
            line-height: 1;
            text-align: center;
        `;
    }
}

function adjustTableLayout() {
    const fixedWidth = '500px';
    const columnName = 'permissionName';

    const headerCell = document.querySelector(`div[role="columnheader"][data-item-key="${columnName}"]`);
    if (headerCell) {
        headerCell.style.width = fixedWidth;
        headerCell.style.minWidth = fixedWidth;
    }

    const dataCells = document.querySelectorAll(`div[role="gridcell"][data-automation-key="${columnName}"]`);
    dataCells.forEach(cell => {
        cell.style.width = fixedWidth;
        cell.style.minWidth = fixedWidth;
    });

    const grid = document.querySelector('div[role="grid"]');
    if (grid) {
        let parent = grid.parentElement;
        while (parent && parent !== document.body) {
            if (parent.style.maxWidth) {
                parent.style.maxWidth = 'none';
            }
            const computedStyle = window.getComputedStyle(parent);
            if (computedStyle.overflow === 'hidden' || computedStyle.overflowY === 'hidden') {
                parent.style.overflow = 'visible';
            }
            parent = parent.parentElement;
        }
    }
}

function adjustParentScroll() {
    const viewport = document.querySelector('.ms-Viewport');
    if (!viewport) {
        return;
    }

    const viewportParent = viewport.parentElement;
    if (viewportParent) {
        // Set a fixed height on the parent container to establish a boundary
        viewportParent.style.height = 'calc(100vh - 250px)'; // Adjust this value as needed
        viewportParent.style.overflow = 'hidden'; // Prevent this container from scrolling
    }

    // Make the viewport itself scrollable within the new boundary
    viewport.style.height = '100%';
    viewport.style.overflowY = 'auto';
}

function forceLoadAllDetailsListItems() {
    if (listProcessed) return;
    const grid = document.querySelector('div[role="grid"]');
    if (!grid || !grid.closest('.ms-DetailsList') || !grid.querySelector('.ms-List-page')) return;

    const scrollableContainer = grid.closest('.ms-Viewport');
    if (!scrollableContainer) return;

    console.log('DetailsList page found. Starting process to load all rows.');
    listProcessed = true;

    scrollableContainer.addEventListener('scroll', () => {
        window.dispatchEvent(new Event('resize'));
    });

    let lastScrollHeight = 0;

    const collectRows = () => {
        const currentRows = Array.from(grid.querySelectorAll('div[role="row"]'));
        let newKeys = [];
        currentRows.forEach(row => {
            const key = row.getAttribute('data-item-index');
            if (key && !allRowsMap.has(key)) {
                allRowsMap.set(key, row);
                newKeys.push(parseInt(key, 10));
            }
        });
        return newKeys.sort((a, b) => a - b);
    };

    const scrollAndLoad = () => {
        const newKeys = collectRows();

        if (newKeys.length > 0) {
            if (lastMaxIndex !== -1 && newKeys[0] > lastMaxIndex + 1) {
                console.warn(`Data gap detected. Expected index ${lastMaxIndex + 1}, but found ${newKeys[0]}. Scrolling up to retry.`);
                // Scroll up a bit to try and load the missing rows
                scrollableContainer.scrollTop -= 500; // Adjust this value as needed
                setTimeout(scrollAndLoad, 500); // Wait a bit before retrying
                return;
            }
            lastMaxIndex = newKeys.length > 0 ? Math.max(lastMaxIndex, ...newKeys) : lastMaxIndex;
        }

        const firstRow = grid.querySelector('div[role="row"][data-item-index]');
        if (!firstRow) {
            console.log("No data rows found, retrying in 500ms...");
            setTimeout(scrollAndLoad, 500);
            return;
        }
        const rowHeight = firstRow.offsetHeight;
        const scrollIncrement = rowHeight * 20;

        const lastScrollTop = scrollableContainer.scrollTop;
        lastScrollHeight = scrollableContainer.scrollHeight;

        scrollableContainer.scrollTop += scrollIncrement;
        scrollableContainer.dispatchEvent(new Event('scroll'));
        scrollableContainer.dispatchEvent(new WheelEvent('wheel', { bubbles: true, deltaY: 1 }));
        window.dispatchEvent(new Event('resize'));

        setTimeout(() => {
            const newScrollTop = scrollableContainer.scrollTop;
            const newScrollHeight = scrollableContainer.scrollHeight;

            if (newScrollHeight > lastScrollHeight || newScrollTop > lastScrollTop) {
                scrollAndLoad();
            } else {
                collectRows();
                finalizeLoading();
            }
        }, 1000);
    };

    const finalizeLoading = () => {
        const allRows = Array.from(allRowsMap.values());
        console.log('All loaded and merged rows:', allRows);

        const totalHeight = scrollableContainer.scrollHeight;
        scrollableContainer.style.height = `${totalHeight}px`;
        console.log(`Virtualization disabled. Container height set to ${totalHeight}px.`);

        let parent = scrollableContainer.parentElement;
        while (parent && parent !== document.body) {
            const computedStyle = window.getComputedStyle(parent);
            if (computedStyle.overflow === 'hidden' || computedStyle.overflowY === 'hidden') {
                parent.style.overflow = 'visible';
            }
            if (parent.style.maxHeight && parent.style.maxHeight !== 'none') {
                parent.style.maxHeight = 'none';
            }
            parent = parent.parentElement;
        }
    };

    scrollAndLoad();
}

function runAll() {
  addExpandButtons();
  displayVersion();
  adjustTableLayout();
  adjustParentScroll();
  forceLoadAllDetailsListItems();
}

function resetStateAndReload() {
    console.log('Data has been refreshed. Resetting state and reloading all rows.');
    allRowsMap.clear();
    listProcessed = false;
    lastMaxIndex = -1;

    const scrollableContainer = document.querySelector('.ms-Viewport');
    if (scrollableContainer) {
        scrollableContainer.scrollTop = 0;
    }

    // Delay the restart slightly to allow the UI to settle
    setTimeout(runAll, 500);
}

// Monkey-patch fetch to detect data refreshes
const originalFetch = window.fetch;
window.fetch = function(...args) {
    const [url] = args;

    const promise = originalFetch.apply(this, args);

    if (typeof url === 'string' && url.includes('api/data/v9.2/powerpagecomponents')) {
        promise.then(res => {
            if (res.ok) {
                resetStateAndReload();
            }
        });
    }

    return promise;
};

const observer = new MutationObserver(() => {
  // Disconnect the observer to prevent an infinite loop
  observer.disconnect();

  // Run the functions that modify the DOM
  runAll();

  // Reconnect the observer to watch for future changes
  observer.observe(document.body, {
    childList: true,
    subtree: true,
  });
});

observer.observe(document.body, {
  childList: true,
  subtree: true,
});

runAll();