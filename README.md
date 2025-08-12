# Skill Matrix SharePoint WebPart

A SharePoint Framework (SPFx) webpart that displays employee skill matrix data from a SharePoint list called "SkillLogs".

## Features

- 📊 Displays skill matrix data in a clean, professional table format
- 🔄 Real-time data fetching from SharePoint SkillLogs list
- 📱 Responsive design that works on desktop and mobile
- 🎨 Modern UI with blue theme matching SharePoint design
- ⚡ Loading states and error handling
- 🔗 Support for document attachments

## SharePoint List Requirements

This webpart expects a SharePoint list named "SkillLogs" with the following columns:

| Column Name | Type | Description |
|-------------|------|-------------|
| EmployeeEmail | Single line of text | Employee's email address |
| Skills | Single line of text | Skill name (e.g., "Java", "React") |
| ExpectedProficiency | Single line of text | Expected skill level (e.g., "P4") |
| ActualProficiency | Single line of text | Current skill level (e.g., "P3") |
| ActionPlan | Multiple lines of text | Action plan to improve skill |
| Timeline | Single line of text | Timeline for skill improvement |
| SkillGap | Single line of text | Gap between expected and actual |
| Comments | Multiple lines of text | Employee comments |
| Attachments | Attachment | Supporting documents |

## Installation

### Prerequisites

- Node.js (version 16.13.0 - 18.x)
- SharePoint Online tenant
- SharePoint Framework development environment

### Setup

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Trust the development certificate:**
   ```bash
   gulp trust-dev-cert
   ```

3. **Update serve.json:**
   Edit `config/serve.json` and replace the placeholder URL with your SharePoint site:
   ```json
   {
     "initialPage": "https://yourtenant.sharepoint.com/sites/yoursite/_layouts/workbench.aspx"
   }
   ```

### Development

1. **Start the local server:**
   ```bash
   gulp serve
   ```

2. **Add the webpart to your SharePoint workbench page**

### Deployment

1. **Build the solution:**
   ```bash
   gulp build
   gulp bundle --ship
   gulp package-solution --ship
   ```

2. **Deploy to SharePoint:**
   - Upload the `.sppkg` file from `solution/` folder to your App Catalog
   - Install the app on your SharePoint site
   - Add the webpart to a SharePoint page

## Configuration

The webpart includes property pane settings where you can:
- Set a custom description
- Specify the SharePoint list name (defaults to "SkillLogs")

## Usage

1. Ensure your SharePoint site has a list named "SkillLogs" with the required columns
2. Add some sample data to the list
3. Add the Skill Matrix webpart to a SharePoint page
4. The webpart will automatically fetch and display the data

## Data Structure

The webpart displays data in the following format:

| Skills | Expected Proficiency | Actual Proficiency | Action Plan | Timeline | Skill Gap | Employee Comments | Attach Document |
|--------|---------------------|-------------------|-------------|----------|-----------|-------------------|-----------------|
| Java | P4 | P3 | Training course | Q2 2024 | 1 level | Need more practice | 📎 View |

## Troubleshooting

### Common Issues

1. **List not found error:**
   - Ensure the "SkillLogs" list exists in your SharePoint site
   - Check that all required columns are present with correct names

2. **Permission errors:**
   - Ensure the webpart has read permissions to the SharePoint list
   - Check that the user has access to the site and list

3. **No data displayed:**
   - Verify that the list contains data
   - Check browser console for any JavaScript errors
   - Use the refresh button to reload data

### API Calls

The webpart uses SharePoint REST API to fetch data:
```
GET /_api/web/lists/getbytitle('SkillLogs')/items?$select=EmployeeEmail,Skills,ExpectedProficiency,ActualProficiency,ActionPlan,Timeline,SkillGap,Comments,Attachments
```

## File Structure

```
src/
├── models/
│   └── ISkillLog.ts              # TypeScript interfaces
├── services/
│   └── SharePointService.ts      # SharePoint API service
└── webparts/
    └── skillMatrix/
        ├── components/
        │   ├── SkillMatrix.tsx           # Main React component
        │   └── SkillMatrix.module.scss   # Component styles
        ├── loc/
        │   ├── en-us.js                  # Localization strings
        │   └── mystrings.d.ts            # String type definitions
        ├── SkillMatrixWebPart.ts         # Main webpart class
        └── SkillMatrixWebPart.manifest.json # Webpart manifest
```

## Technologies Used

- SharePoint Framework (SPFx) 1.18.2
- React 17.0.1
- TypeScript 4.5.5
- Office UI Fabric React
- SCSS for styling

## License

This project is provided as-is for educational and development purposes.

## Contributing

Feel free to submit issues and enhancement requests!

## Support

For support with this webpart:
1. Check the troubleshooting section above
2. Review SharePoint Framework documentation
3. Check SharePoint developer community forums