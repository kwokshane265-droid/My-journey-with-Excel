# My journey with Excel

This project marks my first experience using Excel for data analysis. After completing a guided course, I independently recreated two projects without video assistance: an interactive salary dashboard that compares median salaries by job role, country, and work schedule, and an in-demand skills analysis using Power Query, DAX, and Power Pivot to clean, organize, and model job-posting data. Together, these projects helped me develop practical Excel skills while identifying valuable data-industry skills based on demand and median salary.

## [Project 1: Salary Dashboard](Salary DashBoard Project) 

<img width="1647" height="478" alt="Salary Dashbard" src="https://github.com/user-attachments/assets/00991168-072e-4c34-b6da-c2b42758514a" />

### Key Objectives
- Build an interactive Excel dashboard for exploring data-job salaries.
- Compare median salaries across different job roles and countries.
- Filter salary data by job title, location, and work schedule.
- Present job-market information in a clear and accessible format.

### Key Lessons
- Excel formulas: Used dynamic IF statements and combined ISNUMBER with SEARCH to categorize data.
- Created structured tables with conditional formatting to highlight relevant values.
- Built map visualizations to compare salary data across locations.

#### How to find the specific salaries?
<img width="1298" height="60" alt="Screenshot 2026-09-02 222536" src="https://github.com/user-attachments/assets/953c6dd7-60b0-4c06-b7c0-e2e6f693a919" />

```
=MEDIAN(
  IF(
   (jobs[job_title_short]=job_title_slot)*
   (jobs[job_country]=job_country_slot)*
   (ISNUMBER(SEARCH(type_slot,jobs[job_schedule_type]))),
   jobs[salary_year_all]
  )
 ) 
```
The above formula was used to deliver the Output "Median Salary" and a similar formula was used to deliver "Job Count" (median was replaced with count)

#### How to create the highlighted bar charts?
<img width="1571" height="407" alt="Screenshot 2026-09-02 205625" src="https://github.com/user-attachments/assets/f4586ab6-f4cf-41c5-bf5a-338fba1375fd" />
Above is the method in which I created Bar charts that would have 1 data point highlighted. (From left to right).

Essentially, there are 2 series being displayed, each showing a different color. One is responsible for all the "Non-focus" items. Another is responsible for ONLY the "Focus Item".

---

## [Project 2](Data Skills Analysis Project)

All data used was processed with PowerQuery and PowerPivot. 
In **Power Query**, I duplicated the main dataset to isolate and clean the skills data. I removed all unnecessary columns—keeping only `Job ID` and the grouped skills—and then split those skill strings out into individual rows. This gave me two clean, lightweight tables: one for main job details and a secondary table linking each `Job ID` to its specific skills.

In **Power Pivot**, I built a relational data model by creating a one-to-many relationship between the two tables using `Job ID`. Connecting them this way avoided a bloated dataset while allowing my Pivot Tables and slicers to stay synchronized. This setup made it seamless to cross-filter across roles, locations, and individual technical skills.

This allowed me to produce the following graphs
### Linear relationship between skills and Median Salary
<img width="1175" height="438" alt="Screenshot 2026-09-02 233508" src="https://github.com/user-attachments/assets/a37602cc-ad92-43d4-aa9f-de8ffbe94a92" />


<img width="1322" height="400" alt="Screenshot 2026-09-02 233548" src="https://github.com/user-attachments/assets/eaacd2d3-b451-4cdc-a518-1d511674142a" />

<img width="1450" height="442" alt="Screenshot 2026-09-02 233535" src="https://github.com/user-attachments/assets/396b1ddb-97e4-46a4-8a5e-5650c2a80a8b" />


