

import * as df from "durable-functions"
import { AppointmentsFile } from "../Interfaces/appointments-file";
import { AppointmentUser } from "../Interfaces/appointment-user";
import { EventsFileOrchestratorInput } from "../Interfaces/events-file-orchestrator-input";
import { EventsProcessUserAppointmentsInput } from "../Interfaces/events-process-user-appointments-input";
import { TelemetryClient } from "applicationinsights";

const telemetryClient = new TelemetryClient(process.env["APPLICATIONINSIGHTS_CONNECTION_STRING"])

const orchestrator = df.orchestrator(function* (context) {


    const eventsFileOrchestratorInput = context.df.getInput<EventsFileOrchestratorInput>();

    const runId = context.df.instanceId;

    if (!context.df.isReplaying) telemetryClient.trackEvent({ name: "EventsFileOrchestrator", properties: { status: "started", runId: runId, input: eventsFileOrchestratorInput} });

    context.log('EventsFileOrchestrator input ', eventsFileOrchestratorInput);


    //generate a new access token
    if (!context.df.isReplaying) telemetryClient.trackEvent({ name: "EventsFileOrchestrator", properties: { status: "call", activity: "GeneralFetchGraphAccessToken", runId: runId} });
    const accessToken = yield context.df.callActivity("GeneralFetchGraphAccessToken");
   // context.log('EventsFileOrchestrator: got token: ', accessToken);

    //load the file from blob and return array of user appointments
    if (!context.df.isReplaying) telemetryClient.trackEvent({ name: "EventsFileOrchestrator", properties: { status: "call", activity: "EventsFileLoader", runId: runId} });
    const appointmentsFile: AppointmentsFile = yield context.df.callActivity("EventsFileLoader", eventsFileOrchestratorInput);

    let appointmentUser = {} as AppointmentUser;
    const tasks = [];
    for (let index = 0; index < appointmentsFile.Changes.length; index++) {
        appointmentUser = appointmentsFile.Changes[index];
        // add runid
        tasks.push(context.df.callActivity("EventsProcessUserAppointments", { accessToken, appointmentUser, runId: runId } as EventsProcessUserAppointmentsInput));
    }


    if (!context.df.isReplaying) telemetryClient.trackEvent({ name: "EventsFileOrchestrator", properties: { status: "call", activity: "EventsProcessUserAppointments", runId: runId, appointmentUser} });

    // wait for all the calendar fetch activities to complete.
    const result: any = yield context.df.Task.all(tasks);

    if (!context.df.isReplaying) telemetryClient.trackEvent({ name: "EventsFileOrchestrator", properties: { status: "ended", activity: "EventsProcessUserAppointments", runId: runId, appointmentUser} });


    return true;
});

export default orchestrator;
