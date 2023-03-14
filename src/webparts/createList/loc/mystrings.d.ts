declare interface ICreateListWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  CreateBtn: string;
  DeleteBtn: string;
}

declare module 'CreateListWebPartStrings' {
  const strings: ICreateListWebPartStrings;
  export = strings;
}
