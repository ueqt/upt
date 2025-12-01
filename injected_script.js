
(function() {
    'use strict';

    console.log('[Injected Script] Loaded and intercepting XHR.');

    const originalXhrOpen = XMLHttpRequest.prototype.open;
    const originalXhrSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;

    XMLHttpRequest.prototype.open = function(method, url) {
        // When open is called, tag the request instance if the URL is relevant
        if (typeof url === 'string' && url.includes('/api/data')) {
            this._interceptUrl = url; // Store URL on the instance
            console.log('[Injected Script] XHR.open detected target URL:', url);
        } else {
            delete this._interceptUrl; // Clean up for other requests
        }
        return originalXhrOpen.apply(this, arguments);
    };

    XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
        // When a header is set, check if it's the one we want on a tagged request
        if (this._interceptUrl && header.toLowerCase() === 'authorization') {
            const instanceUrl = new URL(this._interceptUrl).origin;
            const bearerToken = value;

            // We have everything we need. Log it and send it.
            console.log('[Injected Script] Captured Credentials from XHR:', {
                instanceUrl: instanceUrl,
                bearerToken: bearerToken
            });

            window.postMessage({
                type: 'CREDENTIALS_CAPTURED',
                payload: {
                    bearerToken: bearerToken,
                    instanceUrl: instanceUrl,
                },
            }, '*');
            
            // Clean up the property to avoid sending multiple times for the same request
            delete this._interceptUrl; 
        }
        return originalXhrSetRequestHeader.apply(this, arguments);
    };

})();