import { IEmployee } from "../EmployeeOtmWebPart";

export interface IEmployeeOtmProps {
  employees: Array<IEmployee>;
  assign_eotm: (prevEotm: IEmployee, newEotm: IEmployee, reason: string)=>Promise<IEmployee | null>;
  rootLink: string;
}
