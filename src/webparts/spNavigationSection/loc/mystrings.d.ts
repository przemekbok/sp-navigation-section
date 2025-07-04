declare interface ISpNavigationSectionWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  NavigationListLabel: string;
  CreateNewListText: string;
  ViewSelectedListText: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'SpNavigationSectionWebPartStrings' {
  const strings: ISpNavigationSectionWebPartStrings;
  export = strings;
}
