import { AppointmentUser } from "./appointment-user";

export interface EventsProcessUserAppointmentsInput {
    accessToken: string;
    appointmentUser: AppointmentUser;
    runId: string;
}