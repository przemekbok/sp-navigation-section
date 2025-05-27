import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLink,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SpNavigationSectionWebPartStrings';
import SpNavigationSection from './components/SpNavigationSection';
import { ISpNavigationSectionProps, INavigationItem, IListInfo } from './components/ISpNavigationSectionProps';

export interface ISpNavigationSectionWebPartProps {
  description: string;
  selectedListId: string;
}

export default class SpNavigationSectionWebPart extends BaseClientSideWebPart<ISpNavigationSectionWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _lists: IListInfo[] = [];
  private _navigationItems: INavigationItem[] = [];

  public render(): void {
    const element: React.ReactElement<ISpNavigationSectionProps> = React.createElement(
      SpNavigationSection,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        selectedListId: this.properties.selectedListId,
        navigationItems: this._navigationItems,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      return this._loadLists();
    }).then(() => {
      if (this.properties.selectedListId) {
        return this._loadNavigationItems();
      }
      return Promise.resolve();
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._lists.length === 0) {
      this._loadLists().then(() => {
        this.context.propertyPane.refresh();
      });
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    
    if (propertyPath === 'selectedListId' && newValue !== oldValue) {
      this.properties.selectedListId = newValue;
      this._loadNavigationItems().then(() => {
        this.render();
        this.context.propertyPane.refresh();
      });
    } else if (propertyPath === 'description' && newValue !== oldValue) {
      this.properties.description = newValue;
      this.render();
    }
  }

  private _loadLists(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title&$orderby=Title`;
    
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error(`Failed to load lists: ${response.statusText}`);
        }
      })
      .then((data: any) => {
        this._lists = data.value.map((list: any) => ({
          id: list.Id,
          title: list.Title
        }));
        console.log('Lists loaded:', this._lists);
      })
      .catch((error) => {
        console.error('Error loading lists:', error);
        this._lists = [];
      });
  }

  private _loadNavigationItems(): Promise<void> {
    if (!this.properties.selectedListId) {
      this._navigationItems = [];
      return Promise.resolve();
    }

    console.log('Loading navigation items for list:', this.properties.selectedListId);
    
    // Try multiple field name variations for better compatibility
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists('${this.properties.selectedListId}')/items?$select=Id,Title,Display_x0020_Text,DisplayText,Link&$orderby=ID`;
    
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error(`Failed to load navigation items: ${response.statusText}`);
        }
      })
      .then((data: any) => {
        console.log('Raw list data:', data);
        this._navigationItems = data.value.map((item: any) => {
          // Handle different field name possibilities
          const displayText = item.Display_x0020_Text || item.DisplayText || item.Title || 'Untitled';
          let link = '#';
          
          // Handle different link field formats
          if (item.Link) {
            if (typeof item.Link === 'string') {
              link = item.Link;
            } else if (item.Link.Url) {
              link = item.Link.Url;
            } else if (item.Link.Description) {
              link = item.Link.Description;
            }
          }
          
          return {
            displayText: displayText,
            link: link
          };
        });
        
        console.log('Navigation items loaded:', this._navigationItems);
      })
      .catch((error) => {
        console.error('Error loading navigation items:', error);
        this._navigationItems = [];
      });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const listOptions: IPropertyPaneDropdownOption[] = [
      { key: '', text: 'Select a list...' },
      ...this._lists.map(list => ({
        key: list.id,
        text: list.title
      }))
    ];

    const selectedList = this._lists.find(list => list.id === this.properties.selectedListId);
    const newListUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/new.aspx`;
    const listViewUrl = selectedList ? `${this.context.pageContext.web.absoluteUrl}/Lists/${selectedList.title.replace(/\s+/g, '')}` : '#';

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Navigation Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Header Text',
                  placeholder: 'Enter header text...'
                }),
                PropertyPaneDropdown('selectedListId', {
                  label: 'Select Navigation List',
                  options: listOptions,
                  selectedKey: this.properties.selectedListId
                }),
                PropertyPaneLink('', {
                  href: newListUrl,
                  text: 'Create New List',
                  target: '_blank'
                }),
                PropertyPaneLink('', {
                  href: listViewUrl,
                  text: 'View Selected List',
                  target: '_blank',
                  disabled: !this.properties.selectedListId
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
