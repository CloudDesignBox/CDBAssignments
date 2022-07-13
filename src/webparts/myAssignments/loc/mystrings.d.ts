declare interface IMyAssignmentsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'MyAssignmentsWebPartStrings' {
  const strings: IMyAssignmentsWebPartStrings;
  export = strings;
}
