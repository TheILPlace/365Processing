
import { AzureFunction, Context } from "@azure/functions"

import { DefaultAzureCredential } from '@azure/identity';
import { TelemetryClient } from 'applicationinsights';

const telemetryClient = new TelemetryClient(process.env["APPLICATIONINSIGHTS_CONNECTION_STRING"])



const activityFunction: AzureFunction = async function (context: Context): Promise<string> {   
    try{
    const scope = 'https://graph.microsoft.com/.default';

    let accessToken = "";

    //const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const credential = new DefaultAzureCredential();


    accessToken = (await credential.getToken(scope)).token;

    return accessToken
    } catch(error){
        telemetryClient.trackException({exception:error});
        telemetryClient.trackEvent({ name: "GeneralFetchGraphAccessToken", properties: { status: "error",error} });
        return "";
    }
};

export default activityFunction;
