# My Journey with Excel

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=flat&logo=microsoft-excel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power_Query-004B8D?style=flat&logo=powerbi&logoColor=white)
![Power Pivot](https://img.shields.io/badge/Power_Pivot-005A9C?style=flat)
![DAX](https://img.shields.io/badge/DAX-2E7D32?style=flat)

This project marks my first experience using Excel for data analysis. After completing a guided course, I independently recreated two projects **without video assistance**:

1. An **interactive salary dashboard** comparing median salaries by job role, country, and work schedule.
2. An **in-demand skills analysis** using Power Query, DAX, and Power Pivot to clean, organize, and model job-posting data.

Together, these projects helped me build practical Excel skills while identifying valuable data-industry skills based on demand and median salary.

## Table of Contents
- [Salary Dashboard Project](#salary-dashboard-project)
  - [Key Objectives](#key-objectives)
  - [Key Lessons](#key-lessons)
  - [Formula Breakdown](#formula-breakdown)
- [Data Skills Analysis Project](#data-skills-analysis-project)
  - [Key Objectives](#key-objectives-1)
  - [Key Lessons](#key-lessons-1)
  - [Insights](#insights)
- [Takeaways](#takeaways)

---

## Salary Dashboard Project
📁 [View project files](https://github.com/kwokshane265-droid/My-journey-with-Excel/tree/main/Salary%20DashBoard%20Project)

<img width="1647" height="478" alt="Interactive salary dashboard showing median salary and job counts filtered by job title, country, and work schedule" src="https://github.com/user-attachments/assets/00991168-072e-4c34-b6da-c2b42758514a" />

### Key Objectives
- Build an interactive Excel dashboard for exploring data-job salaries.
- Compare median salaries across different job roles and countries.
- Filter salary data by job title, location, and work schedule.
- Present job-market information in a clear and accessible format.

### Key Lessons
- **Dynamic filtering:** Combined `IF` statements with `ISNUMBER` and `SEARCH` to categorize and filter data by criteria.
- **Structured tables:** Built structured tables with conditional formatting to highlight relevant values.
- **Map visualizations:** Compared salary data across locations using Excel's built-in map charts.

### Formula Breakdown

#### 1. Finding a specific median salary
<img width="1298" height="60" alt="Data validation dropdowns for job title, country, and schedule type used as formula inputs" src="https://github.com/user-attachments/assets/953c6dd7-60b0-4c06-b7c0-e2e6f693a919" />

```excel
=MEDIAN(
  IF(
   (jobs[job_title_short]=job_title_slot)*
   (jobs[job_country]=job_country_slot)*
   (ISNUMBER(SEARCH(type_slot,jobs[job_schedule_type]))),
   jobs[salary_year_all]
  )
)
```
For each category — job title, job country, and job schedule — this formula scans the dataset for matches against the selected options (the orange boxes on the dashboard). It only returns a value when all three categories match, and a median is calculated from the results.

The dropdown options in the orange boxes are powered by data validation, each listing the unique items within its category.

#### 2. Finding job counts
```excel
=COUNT(
  IF(
   (jobs[job_title_short]=job_title_slot)*
   (jobs[job_country]=job_country_slot)*
   (ISNUMBER(SEARCH(type_slot,jobs[job_schedule_type]))),
   jobs[salary_year_all]
  )
)
```
Same logic as the median formula above, but using `COUNT` instead of `MEDIAN` to calculate the number of matching job postings.

#### 3. Highlighted bar charts
<img width="1571" height="407" alt="Bar chart with one highlighted data point standing out in a different color from the rest of the series" src="https://github.com/user-attachments/assets/f4586ab6-f4cf-41c5-bf5a-338fba1375fd" />

To highlight a single bar (left to right), the chart uses **two data series**:
- One series covers all the "non-focus" items, shown in one color.
- A second series covers only the "focus" item, shown in a contrasting color.

#### 4. Finding the best platform
```excel
=COUNT(
 IF(
  (jobs[job_title_short]=job_title_slot)*
  (jobs[job_country]=job_country_slot)*
  (ISNUMBER(SEARCH(type_slot,jobs[job_schedule_type])))*
  (jobs[job_platform]=M2),
  jobs[salary_year_all]
 )
)
```
This extends the earlier matching logic with one more condition: the job platform must equal the value in column M (a unique list of job platforms). This formula is copied across column M, producing a count of matching postings for each platform.

```excel
=XLOOKUP(
 MAX('Data Validation'!$N$2:$N$594),
 'Data Validation'!$N$2:$N$594,
 'Data Validation'!$M$2#
)
```
On the dashboard itself, this formula finds the highest count from that list and returns the corresponding job platform — i.e., the platform with the most matching postings.

---

## Data Skills Analysis Project
📁 [View project files](https://github.com/kwokshane265-droid/My-journey-with-Excel/tree/main/Data%20Skills%20Analysis%20Project)

### Key Objectives
- Clean and model raw job-posting data using Power Query and Power Pivot.
- Identify the most in-demand skills for each data role.
- Compare skill count, skill demand, and median salary to find where compensation and demand intersect.

### Key Lessons
- **Power Query cleanup:** Duplicated the main dataset to isolate and clean the skills data — keeping only `Job ID` and the grouped skills, then splitting those skill strings into individual rows. This produced two clean, lightweight tables: one for job details, and one linking each `Job ID` to its specific skills.

  <img width="457" height="250" alt="Power Query steps showing the job skills column split into individual rows" src="https://github.com/user-attachments/assets/e62240cd-f525-4ebc-90ad-430d7900a358" />
  <img width="401" height="228" alt="Cleaned skills table linked to Job ID after Power Query transformation" src="https://github.com/user-attachments/assets/b136e456-e005-49f4-bfc1-5877e69cf21e" />

- **Power Pivot data modeling:** Built a one-to-many relationship between the two tables using `Job ID`. This avoided a bloated dataset while keeping pivot tables and slicers synchronized, making it seamless to cross-filter across roles, locations, and individual technical skills.

  <img width="718" height="191" alt="Power Pivot data model showing a one-to-many relationship between the jobs table and skills table via Job ID" src="https://github.com/user-attachments/assets/4662b6f3-d53d-4d5f-ab14-531c4cdfd280" />

### Insights

#### Skill count vs. median salary
<img width="687" height="405" alt="Scatter chart plotting median salary against average number of skills requested, by data role" src="https://github.com/user-attachments/assets/9af71115-613c-4121-a5e8-8d842091224d" />

This chart explores whether acquiring a broader range of skills leads to higher compensation, plotting median salary against the average number of skills requested per job posting across major data roles.

- **More skills generally mean higher pay.** Compensation scales alongside technical breadth across entry- and mid-level roles, starting from Business Analysts ($85k; 3.3 skills) and rising as roles require broader software management.
- **Data Engineers need the broadest toolset.** Data Engineering demands the highest tool volume overall, peaking with Senior Data Engineers (8.1 skills; $147.5k median) due to complex pipeline, database, and cloud infrastructure requirements.
- **Domain depth beats tool quantity.** Senior Data Scientists command the highest market compensation ($155k median) with a moderate skill count (5.3 skills) — deep expertise in a specialized domain pays more than accumulating a wide list of secondary tools.

#### Median salary vs. demand for top 10 skills (Data Analyst)
<img width="683" height="345" alt="Bar and line chart comparing median salary against likelihood of appearing in a job posting for the top 10 Data Analyst skills" src="https://github.com/user-attachments/assets/8340d157-fcc0-4022-9b95-8579c090d4de" />

This chart evaluates the top 10 skills for each job title by comparing median salary against the likelihood of that skill appearing in a job posting — balancing market demand with compensation potential. The example below shows a **Data Analyst**.

- **SQL and Excel form the high-demand foundation.** SQL (52% likelihood; $92.5k median) and Excel (40% likelihood; $84.5k median) are the most frequently requested skills — database querying and spreadsheet mechanics remain mandatory baseline competencies.
- **Python commands the top salary premium.** Python yields the highest median salary ($98.5k) among the top 10 skills while maintaining strong demand (29% likelihood) — programming and automation capabilities unlock higher pay tiers.
- **Niche tools offer targeted leverage.** Enterprise technologies like Oracle ($95k median; 7% likelihood) and Tableau ($95k median; 28% likelihood) command above-average compensation despite lower posting frequency — specialized tools carry strong market value.

#### Most common skills required (Data Engineer, United States)
<img width="593" height="393" alt="Bar chart ranking the top 10 most frequently requested skills for Data Engineer roles in the United States" src="https://github.com/user-attachments/assets/0a4fa6da-b415-4515-94ec-729015756570" />

This chart ranks the top 10 most frequently requested skills for a role by likelihood of appearing in job descriptions, identifying essential technical requirements. The example below shows **Data Engineer** roles in the **United States**.

- **SQL and Python form the core engine.** SQL (49% likelihood) and Python (46% likelihood) lead by a wide margin — strong database querying and scripting capability are foundational for building data pipelines.
- **AWS leads cloud infrastructure requirements.** AWS is the most in-demand cloud platform (30% likelihood), ahead of Azure (22% likelihood), reflecting a market preference for cloud-native storage and compute.
- **Big data and processing tools define modern pipelines.** Spark (22%), Snowflake (17%), Hadoop (12%), and Kafka (12%) show that distributed computing and real-time streaming are standard expectations for Data Engineers.

---

## Takeaways

Working through both projects independently — from raw data to a finished dashboard — reinforced how much of "advanced Excel" comes down to a small set of tools used well: dynamic array formulas, structured data modeling with Power Query and Power Pivot, and clear visual design. It also surfaced a practical, data-backed answer to "what should I learn next": SQL, Python, and cloud platforms show up repeatedly as the highest-leverage skills across data roles.
