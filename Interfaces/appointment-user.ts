import { Appointment } from "./appointment";

export interface AppointmentUser {
    MailAddress: string,
    Appointments: Array<Appointment>
}