export interface AppointmentData {
    LegacyFreeBusyStatus: string, //Busy,
    IsRecurring: boolean,
    MyResponseType: string, //Unknown,
    IsAllDayEvent: boolean,
    Sensitivity: string, // Normal,
    Importance: string, //Normal,
    AppointmentType: string, //Single,
    Categories: [],
    Start: string, //2023-08-13T09:00:00+03:00,
    End: string, //2023-08-13T09:30:00+03:00,
    NumberOfParticipants: number,
    IsOrganizer: boolean
}


export interface Appointment {

    ItemId: string,
          ChangeType: string, //enum
          Appointment: AppointmentData 
}