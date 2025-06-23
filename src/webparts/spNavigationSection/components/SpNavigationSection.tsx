import * as React from 'react';
import styles from './SpNavigationSection.module.scss';
import type { ISpNavigationSectionProps, INavigationSection } from './ISpNavigationSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpNavigationSection extends React.Component<ISpNavigationSectionProps> {
  
  private renderNavigationItems(items: any[]): React.ReactElement[] {
    const elements: React.ReactElement[] = [];
    
    if (!items || items.length === 0) {
      return [];
    }

    // Group items into chunks of 6
    for (let i = 0; i < items.length; i += 6) {
      const chunk = items.slice(i, i + 6);
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

  private renderSections(): React.ReactElement[] {
    const { navigationSections } = this.props;
    
    if (!navigationSections || navigationSections.length === 0) {
      return [];
    }

    return navigationSections.map((section: INavigationSection, sectionIndex: number) => {
      const navigationElements = this.renderNavigationItems(section.items);
      
      return (
        <div key={`section-${sectionIndex}`} className={styles.navigationSectionContainer}>
          <h3 className={styles.sectionHeader}>{escape(section.section)}</h3>
          <div className={styles.navigationContent}>
            {navigationElements}
          </div>
        </div>
      );
    });
  }

  private renderContent(): React.ReactElement {
    const { isLoading, errorMessage, navigationSections, selectedListId } = this.props;

    if (isLoading) {
      return (
        <div className={styles.loadingContainer}>
          <div className={styles.loadingSpinner}/>
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
                <li>Create columns named &quot;Display Text&quot;, &quot;Link&quot;, and &quot;Section&quot;</li>
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

    if (!navigationSections || navigationSections.length === 0) {
      return (
        <div className={styles.noItems}>
          <div className={styles.noItemsIcon}>📋</div>
          <div className={styles.noItemsText}>
            No navigation items found in the selected list.
          </div>
          <div className={styles.noItemsHelp}>
            Add items to your list with &quot;Display Text&quot;, &quot;Link&quot;, and &quot;Section&quot; columns.
          </div>
        </div>
      );
    }

    const sectionElements = this.renderSections();
    return (
      <div className={styles.sectionsContainer}>
        {sectionElements}
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
