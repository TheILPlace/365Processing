import { Client, ClientOptions,AuthenticationProvider } from "@microsoft/microsoft-graph-client";



export const getGraphClientByToken = async (accessToken: string): Promise<Client> => {

    class AccessTokenAuthProvider implements AuthenticationProvider {
        private token: string;
        constructor(token: string) {
          this.token = token;
        }
        async getAccessToken(): Promise<string> {
          return Promise.resolve(this.token);
        }
      }
      
      const createGraphClient = async (token: string): Promise<Client> => {
        const options: ClientOptions = {
          authProvider: new AccessTokenAuthProvider(token),
          defaultVersion: 'v1.0',
          debugLogging: false,
        };
        
        const client = Client.initWithMiddleware(options);        
        return client;
      };
    
      const client = await createGraphClient(accessToken);

      return client;
    
}

