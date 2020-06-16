const APP_ID = '1ce20985-a426-4e30-9543-e7e855a4bba6';
const APP_SECERET = 'v-L4a14k-B3m-YUD2Ya544awY4.0.fZjBz';
const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/c44b6c35-9deb-4ed4-ba89-9f6efbc4606b/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const axios = require('axios');
const qs = require('qs');
const sso = require("office-addin-sso");
let token = '';
let data;

const postData = {
    client_id: APP_ID,
    scope: MS_GRAPH_SCOPE,
    client_secret: APP_SECERET,
    grant_type: 'client_credentials'
};

axios.defaults.headers.post['Content-Type'] =
    'application/x-www-form-urlencoded';

axios
    .post(TOKEN_ENDPOINT, qs.stringify(postData))
    .then(response => {
        token = response.data.access_token;
        console.log(token);
    })
    .then(() => {
        data = getUserData(token);
    })
    .then(() => {
        console.log(data);
    })
    .catch(error => {
        console.log(error);
    });

async function getUserData(token) {
    return new Promise(async function(resolve, reject) {
        try {
            const response = await sso.makeGraphApiCall(token);
            resolve(response)
        } catch (err) {
            reject(err);
        }
    });
}