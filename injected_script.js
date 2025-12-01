(function() {
    'use strict';

    // --- XMLHttpRequest Interception for credentials and $batch filtering ---
    let capturedInstanceUrl, capturedBearerToken;
    const originalOpen = XMLHttpRequest.prototype.open;

    XMLHttpRequest.prototype.open = function(method, url) {
        if (typeof url === 'string' && url.includes('/api/data/')) {
            capturedInstanceUrl = new URL(url, window.location.origin).origin;
        }

        // Add ready state change listener to intercept the response
        this.addEventListener('readystatechange', function() {
            // Check if the request is done and was successful
            if (this.readyState === 4 && this.status === 200) {
                const contentType = this.getResponseHeader('content-type');
                const isBatchRequest = typeof url === 'string' && (url.includes('%24batch') || url.includes('$batch'));

                const storedObjectIds = localStorage.getItem('solutionComponentObjectIds');

                // Check if it's a $batch response and we have items to filter against
                if (contentType && contentType.includes('multipart/mixed') && isBatchRequest && storedObjectIds) {
                    console.log('[Injected Script] Intercepted XHR $batch response for filtering.');

                    const solutionComponentIds = new Set(JSON.parse(storedObjectIds));
                    const responseText = this.responseText;

                    const boundaryMatch = contentType.match(/boundary=(.+)/);
                    if (!boundaryMatch) {
                        return; // Not a valid multipart response
                    }
                    const boundary = boundaryMatch[1];

                    const parts = responseText.split(`--${boundary}`);
                    let modified = false;
                    const newParts = [];

                    for (const part of parts) {
                        if (part.trim() === '' || part.trim() === '--') {
                            newParts.push(part);
                            continue;
                        }

                        const jsonStartIndex = part.indexOf('{');
                        if (jsonStartIndex === -1 || !part.includes('Content-Type: application/json')) {
                            newParts.push(part);
                            continue;
                        }

                        const httpHeadersPart = part.substring(0, jsonStartIndex);
                        const jsonPart = part.substring(jsonStartIndex);

                        try {
                            const data = JSON.parse(jsonPart);
                            if (data && data.value && Array.isArray(data.value)) {
                                const originalCount = data.value.length;
                                data.value = data.value.filter(item => !solutionComponentIds.has(item.powerpagecomponentid));
                                const newCount = data.value.length;

                                if (originalCount !== newCount) {
                                    modified = true;
                                    console.log(`[Injected Script] Filtered ${originalCount - newCount} items from $batch response part.`);
                                }

                                const newJsonPart = JSON.stringify(data);
                                newParts.push(httpHeadersPart + newJsonPart);
                            } else {
                                newParts.push(part);
                            }
                        } catch (e) {
                            console.error('[Injected Script] Error parsing JSON in $batch part:', e);
                            newParts.push(part);
                        }
                    }

                    if (modified) {
                        const newResponseBody = newParts.join(`--${boundary}`);
                        // Redefine responseText and response to reflect the filtered data
                        Object.defineProperty(this, 'responseText', { value: newResponseBody, writable: true });
                        Object.defineProperty(this, 'response', { value: newResponseBody, writable: true });
                        console.log('[Injected Script] Modified XHR $batch response.');
                    }
                }
            }
        }, false);

        originalOpen.apply(this, arguments);
    };

    const originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
        if (header.toLowerCase() === 'authorization') {
            capturedBearerToken = value;
        }
        originalSetRequestHeader.apply(this, arguments);

        if (capturedInstanceUrl && capturedBearerToken) {
            console.log('[Injected Script] Captured Credentials from XHR:', { instanceUrl: capturedInstanceUrl, bearerToken: capturedBearerToken });
            window.postMessage({ type: 'CREDENTIALS_CAPTURED', payload: { bearerToken: capturedBearerToken, instanceUrl: capturedInstanceUrl } }, '*');
            capturedInstanceUrl = undefined;
            capturedBearerToken = undefined;
        }
    };
})();