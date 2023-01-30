declare interface IEmployeeOtmWebPartStrings {
  PropertyPaneListName: string;
  PropertyPaneDescription: string;
  PropertyPaneEmployee: string;
  PropertyPaneMonth: string;
  PropertyPaneReason: string;
  PropertyPanePanelButton: string;
  PropertyPaneChoiceButton: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'EmployeeOtmWebPartStrings' {
  const strings: IEmployeeOtmWebPartStrings;
  export = strings;
}
