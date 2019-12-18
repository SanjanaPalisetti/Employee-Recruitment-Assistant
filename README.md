# Employee-Recruitment-Assistant
An selenium-based web-scraping application that updates employee recruitment and job availability related details on the [Sony Talent Tracker website](http://app.talentrackr.com/sony/Vendor/Login.aspx)

# Motivation
The Sony Talent Tracker website provides details regarding __Available jobs__, __Employees__, and __Employee Recruitment Status__ in the form of tables. The staffing executive of the company is responsible for tracking any new updates that show up in the Job openings or Employee recruitment status to manage the recruitment process. Each row in the _Employees_ table links to a webpage containing the skillset and status of the employee's recruitment. There are over 30 jobs and 150 employees to track on the website. In addition to this, the tables are available on different webpages and are linked through Job IDs, making it difficult to navigate and keep track.


This python program utilizes selenium to web scrape the information provided in the tables by automatically navigating to the linked webpages, and outputs an [excel sheet](Recruitment Info.xls) providing the details of Employees and Jobs along with the updated columns labeled so. 

