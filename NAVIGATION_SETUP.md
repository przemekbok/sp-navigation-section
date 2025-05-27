# SharePoint Navigation Section WebPart

## Summary

A SharePoint Framework (SPFx) webpart that displays navigation links from a SharePoint list. The webpart shows navigation items as clickable links separated by slashes, with 6 items per line. Each link comes from a SharePoint list with configurable "Display Text" and "Link" columns.

![SharePoint Framework Version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Features

- **Dynamic List Selection**: Choose any SharePoint list from the current site
- **Configurable Header**: Set custom header text for the navigation section
- **Slash-Separated Links**: Navigation links displayed with slash separators (no styling)
- **Responsive Layout**: 6 navigation items per line, automatically wrapping
- **Property Pane Integration**: Easy configuration with dropdown list selection
- **Quick Actions**: Direct links to create new lists or view selected list
- **Custom Font Support**: Built-in directions for implementing custom fonts

## List Structure

Create a SharePoint list with the following columns:

| Column Name | Type | Description |
|-------------|------|-------------|
| Title | Single line of text | Default SharePoint title column |
| Display Text | Single line of text | Text to display for the navigation link |
| Link | Hyperlink or Single line of text | URL for the navigation link |

### Sample List Data:

| Title | Display Text | Link |
|-------|-------------|------|
| Home | Home | https://contoso.sharepoint.com |
| About | About Us | https://contoso.sharepoint.com/about |
| Services | Our Services | https://contoso.sharepoint.com/services |
| Contact | Contact | https://contoso.sharepoint.com/contact |
| Blog | Company Blog | https://contoso.sharepoint.com/blog |
| Careers | Join Us | https://contoso.sharepoint.com/careers |

## Installation & Setup

### Prerequisites

- SharePoint Online environment
- SharePoint Framework development environment
- Node.js (version 18.17.1 or higher)

### Build and Deploy

1. **Clone the repository**
   ```bash
   git clone https://github.com/przemekbok/sp-navigation-section.git
   cd sp-navigation-section
   git checkout feature/navigation-webpart
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Build the solution**
   ```bash
   gulp build
   gulp bundle --ship
   gulp package-solution --ship
   ```

4. **Deploy to SharePoint**
   - Upload the `.sppkg` file from `./sharepoint/solution/` to your App Catalog
   - Deploy the solution to your site collection

### Configuration

1. **Create Navigation List**
   - Go to your SharePoint site
   - Create a new list (or use the "Create New List" link in webpart properties)
   - Add the required columns: "Display Text" and "Link"
   - Add your navigation items

2. **Add WebPart to Page**
   - Edit a SharePoint page
   - Add the "SP Navigation Section" webpart
   - Configure the webpart properties:
     - **Header Text**: Enter the title for your navigation section
     - **Select Navigation List**: Choose your navigation list from the dropdown
     - Use quick links to create new list or view selected list

3. **Customize Appearance**
   - The webpart automatically displays 6 navigation items per line
   - Links are separated by forward slashes
   - No special styling is applied to maintain clean appearance

## Custom Font Usage

To implement custom fonts in the webpart, you have several options:

### Option 1: Google Fonts (Recommended)
Edit `src/webparts/spNavigationSection/components/SpNavigationSection.module.scss`:

```scss
// Add at the top of the file
@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');

// Apply to main container
.spNavigationSection {
  font-family: 'Roboto', Arial, sans-serif;
}

// Apply to header specifically
.header {
  font-family: 'Roboto', Arial, sans-serif;
  font-weight: 700;
}

// Apply to navigation links
.navigationLink {
  font-family: 'Roboto', Arial, sans-serif;
  font-weight: 400;
}
```

### Option 2: Local Font Files
1. Create an `assets/fonts/` folder in your webpart directory
2. Add your font files (.woff2, .woff, .ttf)
3. Add font-face declarations in the SCSS:

```scss
@font-face {
  font-family: 'MyCustomFont';
  src: url('../assets/fonts/MyCustomFont.woff2') format('woff2'),
       url('../assets/fonts/MyCustomFont.woff') format('woff');
  font-weight: normal;
  font-style: normal;
}

.spNavigationSection {
  font-family: 'MyCustomFont', Arial, sans-serif;
}
```

### Option 3: SharePoint Theme Fonts
Use SharePoint's theme-aware font variables:

```scss
.spNavigationSection {
  font-family: var(--font-family-primary, $ms-font-family-fallbacks);
}
```

## Development

### Local Development
```bash
gulp serve
```

### Testing
```bash
gulp test
```

### Building for Production
```bash
gulp clean
gulp build
gulp bundle --ship
gulp package-solution --ship
```

## WebPart Properties

| Property | Type | Description |
|----------|------|-------------|
| description | string | Header text displayed above navigation items |
| selectedListId | string | GUID of the selected SharePoint list |

## Technical Implementation

- **Framework**: SharePoint Framework (SPFx) 1.20.0
- **React Version**: 17.0.1
- **UI Framework**: Fluent UI React 8.106.4
- **Build Tools**: Gulp, TypeScript 4.7.4
- **Styling**: SCSS modules with theme awareness

## Browser Support

- Microsoft Edge
- Google Chrome
- Mozilla Firefox
- Safari

## Version History

| Version | Date | Comments |
|---------|------|----------|
| 1.0.0 | 2025-05-27 | Initial navigation webpart implementation |

## Authors

- **Developer**: [Your Name] ([Your Email])
- **Company**: [Your Company]

## Support

For issues and questions:
1. Check the GitHub Issues page
2. Create a new issue with detailed description
3. Include browser version and SharePoint environment details

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
