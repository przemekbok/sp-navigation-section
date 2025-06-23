export interface ISpNavigationSectionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  selectedListId: string;
  navigationSections: INavigationSection[];
  siteUrl: string;
  isLoading: boolean;
  errorMessage: string;
}

export interface INavigationSection {
  section: string;
  items: INavigationItem[];
}

export interface INavigationItem {
  displayText: string;
  link: string;
  section: string;
}

export interface IListInfo {
  id: string;
  title: string;
}
