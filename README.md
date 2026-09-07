# My journey with Excel

This project marks my first experience using Excel for data analysis. After completing a guided course, I independently recreated two projects without video assistance: an interactive salary dashboard that compares median salaries by job role, country, and work schedule, and an in-demand skills analysis using Power Query, DAX, and Power Pivot to clean, organize, and model job-posting data. Together, these projects helped me develop practical Excel skills while identifying valuable data-industry skills based on demand and median salary.

---
## [Salary DashBoard Project](https://github.com/kwokshane265-droid/My-journey-with-Excel/tree/main/Salary%20DashBoard%20Project)

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
---
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
The above formula was used to deliver the Output "Median Salary". For each category (job title, job country and job schedule), this formula would comb through the dataset and look for match with the options (the words in the orange boxes). It would only return values if all 3 categories matched the 3 different options and once those values were returned, a median would be calculated.

The options (orange boxes) were given through data validation and each had a list of the unique items within each catgory. 

#### How to find specific job counts?
```
=COUNT(
  IF(
   (jobs[job_title_short]=job_title_slot)*
   (jobs[job_country]=job_country_slot)*
   (ISNUMBER(SEARCH(type_slot,jobs[job_schedule_type]))),
   jobs[salary_year_all]
  )
 ) 
```
The above formula is similar to previous and performs "Count" instead of "Median", which is what calculated the job count.


#### How to create the highlighted bar charts?
<img width="1571" height="407" alt="Screenshot 2026-09-02 205625" src="https://github.com/user-attachments/assets/f4586ab6-f4cf-41c5-bf5a-338fba1375fd" />
Above is the method in which I created Bar charts that would have 1 data point highlighted. (From left to right).

Essentially, there are 2 series being displayed, each showing a different color. One is responsible for all the "Non-focus" items. Another is responsible for ONLY the "Focus Item".

#### How to find the best platform?
```
=COUNT(
 IF(
  (jobs[job_title_short]=job_title_slot)*
  (jobs[job_country]=job_country_slot)*
  (ISNUMBER(SEARCH(type_slot,jobs[job_schedule_type])))*
  (jobs[job Platform]=M2),
  jobs[salary_year_all]
 )
)
```
First, the above formula was used to match categories to options like the previous formulas. Then an extra condition was added to the IF: The job platform had to be equal to the one stated in column M, which is the unique list of job platforms(currently this is focused on M2, the formula is replicated along M). 

This meant that a count would displayed for each relevant job platform

```
=XLOOKUP(
 MAX(
  'Data Validation'!$N$2:$N$594),
 'Data Validation'!$N$2:$N$594,
 'Data Validation'!$M$2#)
```
Then on the dashboard, the above formula was used to find the maximum count and return the corresponding job platform.

---

## [Data Skills Analysis Project](https://github.com/kwokshane265-droid/My-journey-with-Excel/tree/main/Data%20Skills%20Analysis%20Project)

All data used was processed with PowerQuery and PowerPivot. 
In **Power Query**, I duplicated the main dataset to isolate and clean the skills data. I removed all unnecessary columns—keeping only `Job ID` and the grouped skills—and then split those skill strings out into individual rows. This gave me two clean, lightweight tables: one for main job details and a secondary table linking each `Job ID` to its specific skills.

<img width="457" height="250" alt="Screenshot 2026-09-03 194251" src="https://github.com/user-attachments/assets/e62240cd-f525-4ebc-90ad-430d7900a358" />
<img width="401" height="228" alt="Screenshot 2026-09-03 194501" src="https://github.com/user-attachments/assets/b136e456-e005-49f4-bfc1-5877e69cf21e" />

In **Power Pivot**, I built a relational data model by creating a one-to-many relationship between the two tables using `Job ID`. Connecting them this way avoided a bloated dataset while allowing my Pivot Tables and slicers to stay synchronized. This setup made it seamless to cross-filter across roles, locations, and individual technical skills.

<img width="718" height="191" alt="Screenshot 2026-09-04 093721" src="https://github.com/user-attachments/assets/4662b6f3-d53d-4d5f-ab14-531c4cdfd280" />

This allowed me to produce the following graphs

---
### Linear relationship between kills and median salary
<img width="687" height="405" alt="Screenshot 2026-09-04 093501" src="https://github.com/user-attachments/assets/9af71115-613c-4121-a5e8-8d842091224d" />

This chart explores whether acquiring a broader range of skills leads to higher compensation by plotting median salaries against the average number of skills requested per job posting across major data roles.

* **More Skills Mean Higher Pay:** Compensation generally scales alongside technical breadth across entry- and mid-level roles, starting from Business Analysts ($85k; 3.3 skills) and rising as roles require broader software management.
* **Data Engineers Need the Broadest Toolsets:** Data Engineering demands the highest tool volume overall, peaking with Senior Data Engineers (8.1 skills; $147.5k median) due to complex pipeline, database, and cloud infrastructure requirements.
* **Domain Depth Beats Tool Quantity:** Senior Data Scientists command the highest market compensation ($155k median) while requiring a moderate skill count (5.3 skills), proving that deep technical expertise in specialized domains yields higher financial returns than accumulating a wide list of secondary tools.
---
### Median salary and chance of requirement of top 10 skills
<img width="683" height="345" alt="Screenshot 2026-09-04 093245" src="https://github.com/user-attachments/assets/8340d157-fcc0-4022-9b95-8579c090d4de" />


This chart evaluates the top 10 skills for each job title by comparing their median salaries against the likelihood of each skill appearing in a job requirement, balancing overall market demand with compensation potential. This takes median salary and skill likelihood into consideration.

This case in particular shows the combined information of a **Data Analyst**.

* **SQL and Excel Form the High-Demand Foundation:** SQL (52% likelihood; $92.5k median) and Excel (40% likelihood; $84.5k median) are the most frequently requested skills, proving that database querying and spreadsheet mechanics remain mandatory baseline competencies for Data Analysts.
* **Python Commands the Top Salary Premium:** Python yields the highest median salary ($98.5k) among the top 10 skills while maintaining strong market demand (29% likelihood), showing that programming and automation capabilities unlock higher pay tiers.
* **Niche Tools Offer Targeted Leverage:** Enterprise technologies like Oracle ($95k median; 7% likelihood) and Tableau ($95k median; 28% likelihood) offer above-average compensation, demonstrating that specialized data management and visualization tools carry strong market value even with lower overall posting frequency.

---
### Likelihood of top 10 most common skills required of data-related jobs
<img width="593" height="393" alt="Screenshot 2026-09-04 093056" src="https://github.com/user-attachments/assets/0a4fa6da-b415-4515-94ec-729015756570" />

This chart highlights the top 10 most frequently requested skills for each job role, ranking them by their likelihood of appearing in job descriptions to identify essential technical requirements for the role. It takes the job title and country into account.

This case in particular shows the top 10 skills necessary to be a **Data Engineer** in the **United States**

* **SQL and Python Form the Core Engine:** SQL (49% likelihood) and Python (46% likelihood) lead by a massive margin, proving that strong database querying and scripting capabilities are mandatory foundational tools for building data pipelines.
* **AWS Leads Cloud Infrastructure Requirements:** AWS is the most demanded cloud platform (30% likelihood), outpacing Azure (22% likelihood), reflecting strong market preference for cloud-native storage and compute environments.
* **Big Data and Processing Tools Define modern Pipelines:** Technologies like Spark (22%), Snowflake (17%), Hadoop (12%), and Kafka (12%) highlight that handling large-scale distributed computing and real-time streaming data are standard operational expectations for Data Engineers.
