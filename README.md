# Employee Dashboard

## Privacy and Security

⚠️ All processing is done client-side in the browser. No data is sent to any server or third-party service, making this application suitable for sensitive HR data.

## Features of the application

1. **File Upload Interface**:
   - Drag and drop Excel files
   - Supports .xlsx, .xls, .xlsb, and .xlsm formats

2. **File Structure Visualization**:
   - See all sheets, columns, and data types
   - Preview sample data rows

3. **Employee Data Analysis**:
   - Highlighting for non-participating employees
   - Part-time percentage correction
   - Statistics by manager
   - Visual charts for data distribution

4. **Export Options**:
   - Download processed data as Excel file
   - Formatted statistics sheet included

## Customization

To customize the application for your specific needs:

1. Modify the column mappings in the `processFile` function if your Excel file has different headers
2. Adjust the logic in `isPresentClassName` to match your participation criteria
3. Change the styling by modifying the Tailwind CSS classes

## Troubleshooting

- If your Excel file isn't loading correctly, check that it has the required columns (Leidinggevende, Parttime (%))
- For large files, you may need to increase the browser memory limit by adding `NODE_OPTIONS=--max_old_space_size=4096` to your build commands


## Deployment

1. Deploy to GitHub Pages:
   ```bash
   npm run deploy
   ```

3. Configure GitHub Pages in the repository settings:
   - Go to your repository on GitHub
   - Click Settings > Pages
   - Select the `gh-pages` branch as the source
   - Save your changes
