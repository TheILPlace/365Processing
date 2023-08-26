
import { AzureFunction, Context } from "@azure/functions"
import { EventsProcessUserAppointmentsInput } from "../Interfaces/events-process-user-appointments-input";
import { BatchRequestContent, BatchRequestStep, Client, PageCollection } from "@microsoft/microsoft-graph-client";
import { getGraphClientByToken } from "../utils/graph-utils";
import { TelemetryClient } from 'applicationinsights';
import { Appointment } from "../Interfaces/appointment";
import { GraphAppointment } from "../Interfaces/graph-appointment";


let client: Client;
const telemetryClient = new TelemetryClient(process.env["APPLICATIONINSIGHTS_CONNECTION_STRING"])


const activityFunction: AzureFunction = async function (context: Context, input: EventsProcessUserAppointmentsInput): Promise<boolean> {
    
    telemetryClient.trackEvent({ name: "EventsProcessUserAppointments", properties: { status: "starting", user: input.appointmentUser.MailAddress, runId: input.runId } })

    
    context.log('processing user: ', input.appointmentUser.MailAddress);


    // get userID from graph
    client = await getGraphClientByToken(input.accessToken);

    try {

        //check if the user exists
        try {
            const user = await client
                .api(`/users/${input.appointmentUser.MailAddress}`)
                .select('id')
                .get();

        } catch (error) {
            context.log('ERROR processing user: ', input.appointmentUser.MailAddress, error);
            return false;

        }


        const userId = input.appointmentUser.MailAddress; //user.id;


        //delete all the user's events for the current week
        await deleteAllEvents(input.appointmentUser.MailAddress, context, input.runId);


        //create new user events

        await createAllEvents(input.appointmentUser.MailAddress, input.appointmentUser.Appointments, context, input.runId);

        return true;


    } catch (error) {
        context.log('failed processing user: ', input.appointmentUser.MailAddress);
        context.log('error: ', error);
    }





};

// const createAllEvents2 = async (udn: string, appointments: Array<Appointment>, context: any) => {

//     context.log('about to create all events for user ', udn);
//     context.log(`retrieved ${appointments.length} events for user ${udn} `);


//     const createPromises = appointments.map(async (app) => {

//         const graphAppointment = createEventFromAppointmentObject(udn, app.ItemId, app);



//         return client
//             .api(`/users/${udn}/calendar/events`)
//             .post(graphAppointment)
//             .then(response => response)
//             .catch(error => {
//                 context.log('error creating event ', udn, graphAppointment, error);
//                 return null;
//             });
//     });

//     await Promise.all(createPromises);

//     context.log('created all events for user ', udn);

// }

const createAllEvents = async (udn: string, appointments: Array<Appointment>, context: any, runId: string) => {


    telemetryClient.trackEvent({ name: "EventsProcessUserAppointments", properties: { status: "create events start", runId,user: udn, itemsCount: appointments.length} });

    context.log('about to create all events for user ', udn);
    context.log(`retrieved ${appointments.length} events for user ${udn} `);

    let eventChunks: Array<Array<any>> = [];
    const batchSize = 20;

    for (let i = 0; i < appointments.length; i += batchSize) {
        const chunk = appointments.slice(i, i + batchSize);
        eventChunks.push(chunk);
    }



    for (let index = 0; index < eventChunks.length; index++) {
        const chunk = eventChunks[index];


        const batchRequestSteps = chunk.map((app, index2) => {
            const graphAppointment = createEventFromAppointmentObject(udn, app.ItemId, app);

            const request = {
                id: index2.toString(),
                method: 'POST',
                url: `/users/${udn}/calendar/events`,
                body: graphAppointment,
                headers: {
                    "Content-Type": "application/json"
                  }

            };
            return request;
        });

        //deletePromises.push(client.api('$batch').post({ requests: batchRequestSteps }))
        const res = await client.api('$batch').post({ requests: batchRequestSteps })
        console.log('create-finished batch ', index);

    }



    context.log('created all events for user ', udn);
    telemetryClient.trackEvent({ name: "EventsProcessUserAppointments", properties: { status: "create events end", user: udn, itemsCount: appointments.length} });

}



const createEventFromAppointmentObject = (udn: string, id: string, app: Appointment): GraphAppointment => {
    
    const data = app.Appointment;

    if (data.IsAllDayEvent) {
        data.Start = data.Start.split("T")[0] + "T00:00:00";

        //calculate end date
        const dateObject = new Date(data.End.split("T")[0]);

        // Adding one day
        dateObject.setDate(dateObject.getDate() + 1);

        data.End = dateObject.toISOString().slice(0, 10) + "T00:00:00";

    }


    const graphAppointment: GraphAppointment = {
        body: { content: 'meeting', contentType: 'text' },
        categories: [],
        importance: data.Importance.toLocaleLowerCase(),
        isOrganizer: data.IsOrganizer,
        organizer: { emailAddress: { address: data.IsOrganizer ? udn : "someoneelse@outlook.com", name: 'Name' } },
        // originalStart: string,
        // originalStartTimeZone: string,
        sensitivity: data.Sensitivity.toLocaleLowerCase(),
        start: { dateTime: data.Start.split('+')[0], timeZone: 'Israel Standard Time' },
        end: { dateTime: data.End.split('+')[0], timeZone: 'Israel Standard Time' },
        subject: "Meeting +" + data.NumberOfParticipants + (data.AppointmentType == 'Occurrence' ? ' R' : ''),
        type: data.AppointmentType == 'Occurrence' ? 'occurrence' : 'singleInstance',
        isAllDay: data.IsAllDayEvent,
        transactionId: id
    }

    return graphAppointment;
}



const deleteAllEvents = async (udn: string, context: any, runId: string) => {

    context.log('about to delete all events for user ', udn);

    telemetryClient.trackEvent({ name: "EventsProcessUserAppointments", properties: { status: "delete events start", runId, user: udn} })

    let allResults = [];
    let existingEvents;

    existingEvents = await client
        .api(`/users/${udn}/events`)
        .select('id')
        .get();



    allResults = allResults.concat(existingEvents.value);

    let nextPage = await getNextPage(existingEvents, udn, context);
    while (nextPage) {
        existingEvents = nextPage;
        allResults = allResults.concat(existingEvents.value);
        nextPage = await getNextPage(existingEvents, udn, context);
    }

    context.log(`retrieved ${allResults.length} events for user ${udn} `);

    //now delete all items

    telemetryClient.trackEvent({ name: "EventsProcessUserAppointments", properties: { status: "delete events stat", runId, user: udn, itemsCount: allResults.length} })

    //split into chunks of 20 (for batch processing)


    let eventChunks: Array<Array<any>> = [];
    const batchSize = 20;

    for (let i = 0; i < allResults.length; i += batchSize) {
        const chunk = allResults.slice(i, i + batchSize);
        eventChunks.push(chunk);
    }

    //let deletePromises = [];
    //eventChunks.forEach(chunk => {
    for (let index = 0; index < eventChunks.length; index++) {
        const chunk = eventChunks[index];


        const batchRequestSteps = chunk.map((event, index2) => {
            const request = {
                id: index2.toString(),
                method: 'DELETE',
                url: `/users/${udn}/events/${event.id}`

            };
            return request;
        });

        //deletePromises.push(client.api('$batch').post({ requests: batchRequestSteps }))
        const res = await client.api('$batch').post({ requests: batchRequestSteps })
        console.log('finixhed batch ', index);

    }




    // const deletePromises = allResults.map(async (event) => {
    //     return  client.api(`/users/${udn}/events/${event.id}`)
    //         .delete();
    //     //console.log(`Deleted event with ID: ${event.id}`);
    // });

    //const res = await Promise.all(deletePromises);
    // for (let index = 0; index < deletePromises.length; index++) {
    //     const res = await deletePromises[index];
    //     console.log('finixhed batch ', index);

    // }

    context.log('deleted all events for user ', udn);

    telemetryClient.trackEvent({ name: "EventsProcessUserAppointments", properties: { status: "delete events end", runId, user: udn, itemsCount: allResults.length} })

}

const getNextPage = async (collection: PageCollection, udn: string, context: any) => {
    // context.log("about to get more data for : " + emailBox);
    if (collection['@odata.nextLink']) {
        const nextPageUrl = collection['@odata.nextLink'];

        context.log("we have more  data for : " + udn);
        const nextPageCollection = await client.api(nextPageUrl).get();
        return nextPageCollection;
    }
    return null;
};



export default activityFunction;
