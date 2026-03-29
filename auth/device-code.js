const https = require('https');
const querystring = require('querystring');

function postForm(url, payload) {
  const body = querystring.stringify(payload);

  return new Promise((resolve, reject) => {
    const request = https.request(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (response) => {
      let data = '';

      response.on('data', (chunk) => {
        data += chunk;
      });

      response.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          resolve({
            statusCode: response.statusCode,
            body: parsed
          });
        } catch (error) {
          reject(new Error(`Failed to parse Microsoft auth response: ${error.message}`));
        }
      });
    });

    request.on('error', reject);
    request.write(body);
    request.end();
  });
}

async function requestDeviceCode(authConfig) {
  const response = await postForm(authConfig.deviceCodeEndpoint, {
    client_id: authConfig.clientId,
    scope: authConfig.scopes.join(' ')
  });

  if (response.statusCode < 200 || response.statusCode >= 300) {
    throw new Error(response.body.error_description || 'Failed to start device code flow.');
  }

  return response.body;
}

async function pollDeviceCodeToken(authConfig) {
  const response = await postForm(authConfig.tokenEndpoint, {
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    client_id: authConfig.clientId,
    device_code: authConfig.deviceCode
  });

  if (response.statusCode >= 200 && response.statusCode < 300) {
    return {
      status: 'authorized',
      tokens: response.body
    };
  }

  const errorCode = response.body.error;
  if (errorCode === 'authorization_pending') {
    return {
      status: 'pending',
      error: errorCode,
      message: response.body.error_description
    };
  }

  if (errorCode === 'slow_down') {
    return {
      status: 'slow_down',
      error: errorCode,
      message: response.body.error_description
    };
  }

  if (errorCode === 'expired_token' || errorCode === 'authorization_declined' || errorCode === 'bad_verification_code') {
    return {
      status: 'failed',
      error: errorCode,
      message: response.body.error_description || 'Device code authentication failed.'
    };
  }

  throw new Error(response.body.error_description || 'Device code token polling failed.');
}

module.exports = {
  requestDeviceCode,
  pollDeviceCodeToken
};
