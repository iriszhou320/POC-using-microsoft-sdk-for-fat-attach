require('isomorphic-fetch')
const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const { AuthorizationCodeCredential, ClientSecretCredential, } = require("@azure/identity");

const clientId = 'go to Azure AD and find it'; //client id of the registered application
const scopes: string[] = [
    'https://graph.microsoft.com/.default'
]
const tenantId = 'go to Azure AD and find it';
const clientSecret = 'go to Azure AD and find it';

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


