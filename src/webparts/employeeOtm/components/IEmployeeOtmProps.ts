import { IMicrosoftTeams } from "@microsoft/sp-webpart-base";
import { IEmployee } from "../EmployeeOtmWebPart";

export interface IEmployeeOtmProps {
  Eotm: IEmployee;
  rootLink: string;
  TeamsContext: IMicrosoftTeams;
}
