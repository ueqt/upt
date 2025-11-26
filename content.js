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

  if (headerCenterRegion && !document.querySelector('.version-display')) {
    const versionDisplay = document.createElement('div');
    versionDisplay.textContent = `V${version}`;
    versionDisplay.className = 'version-display';

    versionDisplay.style.cssText = `
      font-family: 'Segoe UI', 'Roboto', 'Helvetica Neue', 'Arial', sans-serif;
      font-size: 14px;
      font-weight: 600;
      color: #FFFFFF;
      text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.2);
      width: 100%;
      text-align: center;
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      pointer-events: none;
    `;

    headerCenterRegion.style.position = 'relative';
    headerCenterRegion.appendChild(versionDisplay);
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

const observer = new MutationObserver(() => {
  // Disconnect the observer to prevent an infinite loop
  observer.disconnect();

  // Run the functions that modify the DOM
  addExpandButtons();
  displayVersion();
  adjustTableLayout();

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

addExpandButtons();
displayVersion();
adjustTableLayout();