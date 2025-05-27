import * as React from 'react';
import styles from './SpNavigationSection.module.scss';
import type { ISpNavigationSectionProps } from './ISpNavigationSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpNavigationSection extends React.Component<ISpNavigationSectionProps> {
  
  private renderNavigationItems(): React.ReactElement[] {
    const { navigationItems } = this.props;
    const elements: React.ReactElement[] = [];
    
    if (!navigationItems || navigationItems.length === 0) {
      return [];
    }

    // Group items into chunks of 6
    for (let i = 0; i < navigationItems.length; i += 6) {
      const chunk = navigationItems.slice(i, i + 6);
      const navigationLine = this.createNavigationLine(chunk, i);
      elements.push(navigationLine);
    }

    return elements;
  }

  private createNavigationLine(items: any[], startIndex: number): React.ReactElement {
    const elements: React.ReactElement[] = [];
    
    items.forEach((item, index) => {
      const key = `nav-item-${startIndex + index}`;
      
      // Add hyperlink
      elements.push(
        <a 
          key={key} 
          href={item.link} 
          className={styles.navigationLink}
          target={item.link.startsWith('http') ? '_blank' : '_self'}
          rel={item.link.startsWith('http') ? 'noreferrer' : undefined}
        >
          {escape(item.displayText)}
        </a>
      );
      
      // Add slash separator if not the last item
      if (index < items.length - 1) {
        elements.push(
          <span key={`separator-${startIndex + index}`} className={styles.separator}> / </span>
        );
      }
    });

    return (
      <div key={`nav-line-${startIndex}`} className={styles.navigationLine}>
        {elements}
      </div>
    );
  }

  private renderContent(): React.ReactElement {
    const { isLoading, errorMessage, navigationItems, selectedListId } = this.props;

    if (isLoading) {
      return (
        <div className={styles.loadingContainer}>
          <div className={styles.loadingSpinner}></div>
          <div className={styles.loadingText}>Loading navigation items...</div>
        </div>
      );
    }

    if (errorMessage) {
      return (
        <div className={styles.errorContainer}>
          <div className={styles.errorIcon}>⚠️</div>
          <div className={styles.errorText}>{errorMessage}</div>
          {selectedListId && (
            <div className={styles.errorHelp}>
              <strong>Troubleshooting tips:</strong>
              <ul>
                <li>Ensure the list has items</li>
                <li>Create columns named "Display Text" and "Link"</li>
                <li>Check that you have permission to read the list</li>
              </ul>
            </div>
          )}
        </div>
      );
    }

    if (!selectedListId) {
      return (
        <div className={styles.configurationMessage}>
          <div className={styles.configIcon}>⚙️</div>
          <div className={styles.configText}>
            Please select a navigation list in the webpart properties to display navigation items.
          </div>
        </div>
      );
    }

    if (!navigationItems || navigationItems.length === 0) {
      return (
        <div className={styles.noItems}>
          <div className={styles.noItemsIcon}>📋</div>
          <div className={styles.noItemsText}>
            No navigation items found in the selected list.
          </div>
          <div className={styles.noItemsHelp}>
            Add items to your list with "Display Text" and "Link" columns.
          </div>
        </div>
      );
    }

    const navigationElements = this.renderNavigationItems();
    return (
      <div className={styles.navigationContent}>
        {navigationElements}
      </div>
    );
  }

  public render(): React.ReactElement<ISpNavigationSectionProps> {
    const {
      description,
      hasTeamsContext
    } = this.props;

    return (
      <section className={`${styles.spNavigationSection} ${hasTeamsContext ? styles.teams : ''}`}>
        {description && (
          <h2 className={styles.header}>{escape(description)}</h2>
        )}
        {this.renderContent()}
      </section>
    );
  }
}
