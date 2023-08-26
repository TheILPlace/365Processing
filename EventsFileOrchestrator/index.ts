

import * as df from "durable-functions"
import { AppointmentsFile } from "../Interfaces/appointments-file";
import { AppointmentUser } from "../Interfaces/appointment-user";
import { EventsFileOrchestratorInput } from "../Interfaces/events-file-orchestrator-input";
import { EventsProcessUserAppointmentsInput } from "../Interfaces/events-process-user-appointments-input";

const orchestrator = df.orchestrator(function* (context) {


    const eventsFileOrchestratorInput = context.df.getInput<EventsFileOrchestratorInput>();

    context.log('EventsFileOrchestrator input ', eventsFileOrchestratorInput);


    //generate a new access token
    const accessToken = yield context.df.callActivity("GeneralFetchGraphAccessToken");
   // context.log('EventsFileOrchestrator: got token: ', accessToken);

    //load the file from blob and return array of user appointments
    const appointmentsFile: AppointmentsFile = yield context.df.callActivity("EventsFileLoader", eventsFileOrchestratorInput);

    let appointmentUser = {} as AppointmentUser;
    const tasks = [];
    for (let index = 0; index < appointmentsFile.Changes.length; index++) {
        appointmentUser = appointmentsFile.Changes[index];
        // add runid
        tasks.push(context.df.callActivity("EventsProcessUserAppointments", { accessToken, appointmentUser, runId: "" } as EventsProcessUserAppointmentsInput));
    }


    // wait for all the calendar fetch activities to complete.
    const result: any = yield context.df.Task.all(tasks);



    return true;
});

export default orchestrator;
