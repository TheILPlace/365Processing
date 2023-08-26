import * as df from "durable-functions"
import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { EventsFileOrchestratorInput } from "../Interfaces/events-file-orchestrator-input";
import { TelemetryClient } from "applicationinsights";

const telemetryClient = new TelemetryClient(process.env["APPLICATIONINSIGHTS_CONNECTION_STRING"])

const httpStart: AzureFunction = async function (context: Context, req: HttpRequest): Promise<any> {
    const client = df.getClient(context);
    context.log('manual working on file ', req.body.blobName);


    telemetryClient.trackEvent({ name: "EventsProcessingHttpStart", properties: { status: "started", containerName: req.body.containerName, blobName: req.body.blobName} });

    const instanceId = await client.startNew(req.params.functionName, undefined, {containerName: req.body.containerName, blobName: req.body.blobName } as EventsFileOrchestratorInput);

    context.log(`Started orchestration with ID = '${instanceId}'.`);

    return client.createCheckStatusResponse(context.bindingData.req, instanceId);
};

export default httpStart;
