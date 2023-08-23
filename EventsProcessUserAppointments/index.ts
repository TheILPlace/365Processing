
import { AzureFunction, Context } from "@azure/functions"
import { EventsProcessUserAppointmentsInput } from "../Interfaces/events-process-user-appointments-input";
import { Client, PageCollection } from "@microsoft/microsoft-graph-client";
import { getGraphClientByToken } from "../utils/graph-utils";
import { TelemetryClient } from 'applicationinsights';
import { Appointment } from "../Interfaces/appointment";
import { GraphAppointment } from "../Interfaces/graph-appointment";


let client: Client;
const telemetryClient = new TelemetryClient(process.env["APPLICATIONINSIGHTS_CONNECTION_STRING"])


const activityFunction: AzureFunction = async function (context: Context, input: EventsProcessUserAppointmentsInput): Promise<boolean> {
    context.log('processing user: ', input.appointmentUser.MailAddress);


    // get userID from graph
    client = await getGraphClientByToken(input.accessToken);

    try {
        // const user = await client
        //     .api(`/users/${input.appointmentUser.MailAddress}`)
        //     .select('id')
        //     .get();

        const userId = input.appointmentUser.MailAddress; //user.id;


        //delete all the user's events for the current week

        await deleteAllEvents(input.appointmentUser.MailAddress, context);


        //create new user events

        await createAllEvents(input.appointmentUser.MailAddress, input.appointmentUser.Appointments, context);

        return true;


    } catch (error) {
        context.log('failed processing user: ', input.appointmentUser.MailAddress);
        context.log('error: ', error);
    }





};

const createAllEvents = async (udn: string, appointments: Array<Appointment>, context: any) => {

    context.log('about to create all events for user ', udn);
    context.log(`retrieved ${appointments.length} events for user ${udn} `);


    const createPromises = appointments.map(async (app) => {

        const graphAppointment = createEventFromAppointmentObject(udn, app.ItemId, app);



        return await client
            .api(`/users/${udn}/calendar/events`)
            .post(graphAppointment);
    });

    await Promise.all(createPromises);

    context.log('created all events for user ', udn);

}


const createEventFromAppointmentObject = (udn: string, id: string, app: Appointment): GraphAppointment => {
    const data = app.Appointment;
    const graphAppointment: GraphAppointment = {
        body: {content: 'meeting', contentType: 'text'},
        categories: [],
        importance: 'normal',
        isOrganizer: data.IsOrganizer,
        organizer: { emailAddress: {address:data.IsOrganizer ? udn : "someoneelse@outlook.com", name: 'Name' }},
        // originalStart: string,
        // originalStartTimeZone: string,
        sensitivity: 'normal',
        start: { dateTime: data.Start, timeZone: 'UTC' },
        end: { dateTime: data.End, timeZone: 'UTC' },
        subject: "Meeting +" + data.NumberOfParticipants + (data.AppointmentType == 'Occurrence' ? ' R' : ''),
        type: data.AppointmentType == 'Occurrence' ? 'occurrence' : ' singleInstance',
        isAllDay: data.IsAllDayEvent,
        transactionId: id
    }

    return graphAppointment;
}



const deleteAllEvents = async (udn: string, context: any) => {

    context.log('about to delete all events for user ', udn);
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

    const deletePromises = allResults.map(async (event) => {
        return await client.api(`/users/${udn}/events/${event.id}`)
            .delete();
        //console.log(`Deleted event with ID: ${event.id}`);
    });

    await Promise.all(deletePromises);
    context.log('deleted all events for user ', udn);

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
