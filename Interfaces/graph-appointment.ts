export interface GraphAppointment {
    body?: ItemBody,
    categories?: Array<string>,
    importance?: string,
    isOrganizer: boolean,
    organizer?: Organizer,
    // originalStart: string,
    // originalStartTimeZone: string,
    sensitivity: string,
    start: DateTimeTimeZone,
    end: DateTimeTimeZone,
    subject?: string,
    type?: string,
    isAllDay: boolean,
    transactionId: string

}

export interface DateTimeTimeZone {
    dateTime: string,
    timeZone: string
}

export interface Organizer {
    emailAddress: EmailAddress;
}

export interface EmailAddress {
    address: string;
    name: string;
}

export interface ItemBody {

    content: string;
    contentType: string;

}