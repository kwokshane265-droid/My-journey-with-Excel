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
