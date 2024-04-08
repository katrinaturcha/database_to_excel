This script automates the extraction, transformation, and loading of sales data into an Excel spreadsheet for analysis. 
It uses Python libraries such as pandas, numpy, and sqlalchemy to manipulate data and openpyxl to manage Excel files.

Key Features:
- Environment Setup: Uses dotenv to load environment variables securely, establishing database connections without hardcoding sensitive information.
- Database Connection: Establishes a connection to a MySQL database using credentials fetched from environment variables. It supports both direct connections and socket-based connections for server environments.
- Data Fetching and Aggregation: Implements a function to fetch sales data from the database based on the last update date. It then aggregates this data daily and monthly for each marketplace and product, providing summaries of quantities sold and total sales.
- Excel Reporting: Generates an Excel report where each sheet corresponds to a marketplace for a specific year. It includes daily and monthly sales data, formatted for clarity and ease of analysis.
- Dynamic Date Handling: Calculates date ranges dynamically and groups Excel columns by days and months, hiding daily columns while leaving monthly summary columns visible.
- Styling and Formatting: Applies Excel styles such as center alignment, bold fonts for headers, and custom column widths to improve readability.
- File Management: Checks if the target Excel file exists and decides whether to create a new file or update an existing one, preventing data duplication.
- Performance Metrics: Tracks and prints the script’s execution time, providing insight into the efficiency of data processing tasks.

This script is ideal for businesses looking to automate their sales reporting processes, allowing for easy tracking of sales trends across different marketplaces and periods. 
It's structured to handle large datasets efficiently and can be customized to fit specific business needs.
