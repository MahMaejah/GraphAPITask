import { Client } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { User } from 'microsoft-graph';

let graphClient: Client | undefined = undefined;

function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider) {
  if (!graphClient) {
    graphClient = Client.initWithMiddleware({
      authProvider: authProvider
    });
  }

  return graphClient;
}

export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
  ensureClient(authProvider);

  // Return the /me API endpoint result as a User object
  const user: User = await graphClient!.api('/me')
    .get();

  return user;
}

export async function getMail(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<any> {
  ensureClient(authProvider);

  // Return the /me/messages API endpoint result
  const mail: any = await graphClient!.api('/me/messages')
    .header('Prefer', 'outlook.body-content-type="text"')
    .select('subject, body, bodyPreview,uniqueBody')
    .get();

  return mail;
}