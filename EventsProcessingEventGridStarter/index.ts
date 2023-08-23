import * as df from "durable-functions"
import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { EventGridEvent } from "@azure/eventgrid";
import { EventsFileOrchestratorInput } from "../Interfaces/events-file-orchestrator-input";

const httpStart: AzureFunction = async function (context: Context, eventGridEvent: EventGridEvent<any> ): Promise<any> {


    const eventSubject = eventGridEvent.subject;



    const parts = eventSubject.split("/");
    if (parts.length !== 7 || parts[0] !== "" || parts[1] !== "blobServices" || parts[2] !== "default" || parts[3] !== "containers") {
        context.res = {
            status: 400,
            body: "Invalid event subject format.",
        };
        context.log("EventsProcessingEventGridStarter: Error in eventSubject", eventSubject);
        return false;
    }

    const containerName = parts[4];
    const blobName = parts[6];





    context.log("event grid data: ", eventGridEvent)

    const client = df.getClient(context);
    const instanceId = await client.startNew("EventsFileOrchestrator", undefined, {containerName,blobName} as EventsFileOrchestratorInput);

    context.log(`Started orchestration with ID = '${instanceId}'.`);

    context.log("status: ", client.createCheckStatusResponse(context.bindingData.req, instanceId));
};

export default httpStart;
