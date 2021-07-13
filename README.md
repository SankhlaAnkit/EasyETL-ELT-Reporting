# Easy ETL-ELT-Reporting Application
Python based GUI solution for File to Database load, transformation, derived columns and generating custom report output.

Hi All,

Welcome to one stop solution for all your ETL/ELT requirements without the hassle of using big shot ETL tools and applications that cost you a lot of money,lot of training, plus the need to install heavy softwares on your system. Add to it the permission requirements for installing these softwares if you belong to one of those reputed IT Firms.
Here all you require is Python installed on your system, import some libraries and voila, you are good to go.

Problem that I faced with most of the ETL/ELT applications:-
1.Most ETL application require defining the source files, format, connections, define database tables before hand. All these steps are very cumbersome and require lot of manual efforts.
2.Almost none of the ETL/ELT applications provide you the capability to create Stored procedures that you can edit and re-use at will.
3.Next, if you have requriement to generate ad-hoc reports that require writing huge queries, customizing the columns, saving the queries for reusability in different environments ( Dev/QA/PROD)

All this lead me to create a simple Drag Drop operation style application.Here leave everything up to the application to automate the desired outcome

What all you can do using these applications?? here's a list:
1. Load any file data in any format ( Excel, Csv, Text etc.) to a Database table.
   -- Specify file location, database connection parameters and few file related parameters.
   -- Select one or mutiple file to load at once.
   -- No need to create DB tables in advance, the application will take care of this step.
   -- Easy load fixed width files with option to update field settings before load and ability to view preview data before load to tables.
 2. Transform existing data or newly loaded table data.
   -- Use the Drag and drop feature to select the data field/columns required for data transformation.
   -- type functions and linkages of your choice on the selected fields.
   -- Map source fields to target Table fields.
   -- setup joins, where clauses etc. by selecting as many tables as required.
   -- Auto Creation of Stored Procecures and ability to view the SQL script and preview the data.
   -- Create Custom/Derived Columns and save them for future use.
 3. Custom Report Extraction.
   -- Ability to generate Custom data reports in different formats ( Excel,Csv, Fixed width).
   -- Data sequencing, field manipulation and custom fields, data preview.
   
 Feel free to use these applications and please share your Feedback as to what new features you would like to add in this utility.
