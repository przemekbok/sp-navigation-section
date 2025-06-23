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
import { ISpNavigationSectionProps, INavigationItem, INavigationSection, IListInfo } from './components/ISpNavigationSectionProps';

export interface ISpNavigationSectionWebPartProps {
  description: string;
  selectedListId: string;
}

export default class SpNavigationSectionWebPart extends BaseClientSideWebPart<ISpNavigationSectionWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _lists: IListInfo[] = [];
  private _navigationSections: INavigationSection[] = [];
  private _isLoading: boolean = false;
  private _errorMessage: string = '';

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
        navigationSections: this._navigationSections,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        isLoading: this._isLoading,
        errorMessage: this._errorMessage
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
      this._errorMessage = '';
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
          throw new Error(`Failed to load lists: ${response.status} ${response.statusText}`);
        }
      })
      .then((data: any) => {
        this._lists = data.value.map((list: any) => ({
          id: list.Id,
          title: list.Title
        }));
        console.log('Lists loaded successfully:', this._lists.length);
      })
      .catch((error) => {
        console.error('Error loading lists:', error);
        this._lists = [];
        this._errorMessage = `Failed to load lists: ${error.message}`;
      });
  }

  private async _loadNavigationItems(): Promise<void> {
    if (!this.properties.selectedListId) {
      this._navigationSections = [];
      this._errorMessage = '';
      return Promise.resolve();
    }

    this._isLoading = true;
    this._errorMessage = '';
    this.render(); // Show loading state

    console.log('Loading navigation items for list ID:', this.properties.selectedListId);
    
    try {
      // First, get the list fields to understand the structure
      const fieldsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.properties.selectedListId}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false`;
      
      const fieldsResponse = await this.context.spHttpClient.get(fieldsUrl, SPHttpClient.configurations.v1);
      if (!fieldsResponse.ok) {
        throw new Error(`Failed to load list fields: ${fieldsResponse.status} ${fieldsResponse.statusText}`);
      }
      
      const fieldsData = await fieldsResponse.json();
      const fields = fieldsData.value;
      console.log('Available fields:', fields);
      
      // Find display text field (look for common variations)
      const displayTextField = fields.find((field: any) => 
        field.InternalName === 'Display_x0020_Text' || 
        field.InternalName === 'DisplayText' ||
        field.InternalName === 'NavigationText' ||
        field.Title === 'Display Text'
      );
      
      // Find link field
      const linkField = fields.find((field: any) => 
        field.InternalName === 'Link' ||
        field.InternalName === 'URL' ||
        field.InternalName === 'NavigationLink' ||
        field.TypeAsString === 'URL' ||
        field.Title === 'Link'
      );

      // Find section field
      const sectionField = fields.find((field: any) => 
        field.InternalName === 'Section' ||
        field.InternalName === 'NavigationSection' ||
        field.Title === 'Section'
      );
      
      console.log('Display field found:', displayTextField);
      console.log('Link field found:', linkField);
      console.log('Section field found:', sectionField);
      
      // Build select fields - always include Title, add others if they exist
      let selectFields = 'Id,Title';
      if (displayTextField) {
        selectFields += `,${displayTextField.InternalName}`;
      }
      if (linkField) {
        selectFields += `,${linkField.InternalName}`;
      }
      if (sectionField) {
        selectFields += `,${sectionField.InternalName}`;
      }
      
      // Now get the list items
      const itemsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.properties.selectedListId}')/items?$select=${selectFields}&$orderby=ID&$top=500`;
      console.log('Items URL:', itemsUrl);
      
      const itemsResponse = await this.context.spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1);
      if (!itemsResponse.ok) {
        throw new Error(`Failed to load navigation items: ${itemsResponse.status} ${itemsResponse.statusText}`);
      }
      
      const itemsData = await itemsResponse.json();
      console.log('Raw list items:', itemsData);
      
      if (!itemsData.value || itemsData.value.length === 0) {
        this._navigationSections = [];
        this._errorMessage = 'No items found in the selected list. Please add items to the list or create the required columns (Display Text, Link, Section).';
      } else {
        // Map items to navigation items
        const navigationItems: INavigationItem[] = itemsData.value.map((item: any) => {
          // Get display text
          let displayText = item.Title || 'Untitled';
          if (displayTextField && item[displayTextField.InternalName]) {
            displayText = item[displayTextField.InternalName];
          }
          
          // Get link
          let link = '#';
          if (linkField && item[linkField.InternalName]) {
            const linkValue = item[linkField.InternalName];
            if (typeof linkValue === 'string') {
              link = linkValue;
            } else if (linkValue && linkValue.Url) {
              link = linkValue.Url;
            } else if (linkValue && linkValue.Description) {
              link = linkValue.Description;
            }
          }

          // Get section
          let section = 'General'; // Default section
          if (sectionField && item[sectionField.InternalName]) {
            section = item[sectionField.InternalName];
          }
          
          return {
            displayText: displayText,
            link: link,
            section: section
          };
        });
        
        // Group items by section
        const sectionsMap = new Map<string, INavigationItem[]>();
        navigationItems.forEach(item => {
          const section = item.section;
          if (!sectionsMap.has(section)) {
            sectionsMap.set(section, []);
          }
          sectionsMap.get(section)!.push(item);
        });

        // Convert map to array of sections
        this._navigationSections = Array.from(sectionsMap.entries())
          .map(([section, items]) => ({
            section: section,
            items: items
          }))
          .sort((a, b) => a.section.localeCompare(b.section)); // Sort sections alphabetically
        
        console.log('Navigation sections processed:', this._navigationSections);
        
        if (this._navigationSections.length === 0) {
          this._errorMessage = 'No valid navigation items found. Please check that your list has items with the required fields.';
        }
      }
      
    } catch (error) {
      console.error('Error loading navigation items:', error);
      this._navigationSections = [];
      this._errorMessage = `Error loading navigation items: ${error.message}. Please check that the list exists and has the required permissions.`;
    } finally {
      this._isLoading = false;
    }
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
    const listViewUrl = selectedList ? `${this.context.pageContext.web.absoluteUrl}/Lists/${encodeURI(selectedList.title)}` : '#';

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
