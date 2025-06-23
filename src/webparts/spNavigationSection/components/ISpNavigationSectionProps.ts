export interface ISpNavigationSectionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  selectedListId: string;
  navigationItems: INavigationItem[];
  siteUrl: string;
  isLoading: boolean;
  errorMessage: string;
}

export interface INavigationItem {
  displayText: string;
  link: string;
}

export interface IListInfo {
  id: string;
  title: string;
}
