(function() {
    'use strict';

    let instanceUrl, bearerToken;

    // Intercept XMLHttpRequest to capture credentials
    const originalOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function(method, url) {
        if (url.includes('/api/data/')) {
            instanceUrl = new URL(url, window.location.origin).origin;
        }
        originalOpen.apply(this, arguments);
    };

    const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
        if (header.toLowerCase() === 'authorization') {
            bearerToken = value;
        }
        originalSetRequestHeader.apply(this, arguments);

        if (instanceUrl && bearerToken) {
            console.log('[Injected Script] Captured Credentials:', { instanceUrl, bearerToken });
            window.postMessage({ type: 'CREDENTIALS_CAPTURED', payload: { bearerToken, instanceUrl } }, '*');
            // Reset to avoid sending multiple messages
            instanceUrl = undefined;
            bearerToken = undefined;
        }
    };
})();