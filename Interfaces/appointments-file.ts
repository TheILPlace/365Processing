import { AppointmentUser } from "./appointment-user";

export interface AppointmentsFile {
    JobType: string,
    Changes: Array<AppointmentUser>,
    StartDateTime: string, //format : 2023-08-16T22:34:33
    EndDateTime: string
}