import * as React from 'react';
import styles from './SpNavigationSection.module.scss';
import type { ISpNavigationSectionProps } from './ISpNavigationSectionProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpNavigationSection extends React.Component<ISpNavigationSectionProps> {
  
  private renderNavigationItems(): React.ReactElement[] {
    const { navigationItems } = this.props;
    const elements: React.ReactElement[] = [];
    
    if (!navigationItems || navigationItems.length === 0) {
      return [<div key="no-items" className={styles.noItems}>No navigation items found. Please configure the webpart properties.</div>];
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

  public render(): React.ReactElement<ISpNavigationSectionProps> {
    const {
      description,
      hasTeamsContext
    } = this.props;

    const navigationElements = this.renderNavigationItems();

    return (
      <section className={`${styles.spNavigationSection} ${hasTeamsContext ? styles.teams : ''}`}>
        {description && (
          <h2 className={styles.header}>{escape(description)}</h2>
        )}
        <div className={styles.navigationContent}>
          {navigationElements}
        </div>
      </section>
    );
  }
}
