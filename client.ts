require('isomorphic-fetch')
const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { AuthorizationCodeCredential, ClientSecretCredential, } = require("@azure/identity");

const clientId = '4ecc15f0-3ba6-4333-b937-7235e3688945'; //client id of the registered application
const scopes: string[] = [
    'https://graph.microsoft.com/.default'
]
const tenantId = '4a7fbee1-6041-40b1-a613-86577b4f827f';
const clientSecret = '4l77Q~72d7QkWiCFsHfTZZDfb1D_TkFfZfv-a';

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
const authProvider = new TokenCredentialAuthenticationProvider(credential,  {
    scopes: [scopes]
});

export class MsClient {
    _getToken = async () => {
        return await authProvider.getAccessToken();
    };

    getClient(): any {
        let token = this._getToken()
        return Client.init({
            defaultVersion: "v1.0",
            debugLogging: true,
            authProvider: function (authDone: any) {
                authDone(null, token)
            },
        });
    }
}


