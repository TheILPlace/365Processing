
import { AzureFunction, Context } from "@azure/functions"
import { TelemetryClient } from 'applicationinsights';
import { BlobServiceClient } from "@azure/storage-blob";
import { AppointmentsFile } from "../Interfaces/appointments-file";
import { EventsFileOrchestratorInput } from "../Interfaces/events-file-orchestrator-input";

const telemetryClient = new TelemetryClient(process.env["APPLICATIONINSIGHTS_CONNECTION_STRING"])

const activityFunction: AzureFunction = async function (context: Context, input: EventsFileOrchestratorInput): Promise<AppointmentsFile> {

    //console.log(context.bindings.inputFile);


    return context.bindings.inputFile as AppointmentsFile;

    // const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AzureWebJobsStorage);
    // const containerName = "incoming"; // Replace with your container name
    // const blobName = input.fileName; // Replace with your JSON file name




    // const containerClient = blobServiceClient.getContainerClient(containerName);
    // const blobClient = containerClient.getBlobClient(blobName);


    // const response = await blobClient.download();
    // const content = await streamToBuffer(response.readableStreamBody);
    // const appointmentsFile: AppointmentsFile = JSON.parse(content.toString());
    // return appointmentsFile;
};


// async function streamToBuffer(readableStream) {
//     return new Promise((resolve, reject) => {
//         const chunks = [];
//         readableStream.on("data", (data) => {
//             chunks.push(data instanceof Buffer ? data : Buffer.from(data));
//         });
//         readableStream.on("end", () => {
//             resolve(Buffer.concat(chunks));
//         });
//         readableStream.on("error", reject);
//     });
// }

export default activityFunction;
