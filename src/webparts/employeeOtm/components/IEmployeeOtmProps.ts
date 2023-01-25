import { IEmployee } from "../EmployeeOtmWebPart";

export interface IEmployeeOtmProps {
  Eotm: IEmployee;
  assignEotm: (prevEotm: IEmployee, newEotm: IEmployee, reason: string)=>Promise<IEmployee | null>;
  rootLink: string;
}
