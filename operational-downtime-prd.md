# Product Requirements Document (PRD)

## Product
Operational Downtime Reporting Application

## Version
Draft v0.1 - 2026-08-04

## 1. Purpose
Create a simple browser-based application for Maersk Contract Logistics that enables site operations teams worldwide to manually submit operational downtime impacts caused by warehouse management system outages, performance degradation, and related technology incidents. The application will standardize global reporting, provide leadership dashboards, and prepare the foundation for later integration with system outage data.

## 2. Background and Problem Statement
Operational downtime is currently captured manually by sites, typically in spreadsheet format. This creates several issues:
- inconsistent data capture across countries and sites
- limited visibility for regional and global leadership
- slow manual consolidation for monthly reviews
- weak trend analysis across customer, application, site, and root cause
- no scalable path to connect operational impact with source system outage data later

The current reporting example shows site-level entries with repeated incident descriptions across multiple impacted sites. This indicates a near-term need for standardized site-by-site submission rather than a complex incident hierarchy.

## 3. Product Vision
Provide one simple, globally accessible internal web application where site teams can record lost operational hours in a consistent format, and leadership can immediately view standardized dashboards and exports for decision-making and service improvement.

## 4. Goals
### 4.1 Business Goals
- Standardize operational downtime reporting globally
- Reduce manual consolidation effort for Operations and TCO teams
- Improve monthly leadership visibility into downtime impact trends
- Identify recurring technology issues and high-impact sites faster
- Establish a clean data foundation for future cost modeling and system integration

### 4.2 Product Goals
- Allow authenticated Maersk users to submit downtime records from any browser
- Store all submissions in a centralized shared data source
- Provide immediate dashboard visibility after submission
- Support Excel and Power BI export/consumption
- Keep the application simple enough for rapid adoption and minimal training

## 5. Non-Goals for Version 1
- automated ingestion from system outage reports
- automatic cost calculation using a formal financial model
- complex workflow orchestration or multi-step approval
- offline entry and sync
- advanced role-based administration beyond basic submit/view/admin behavior
- incident-to-multiple-sites shared record model

## 6. Confirmed Scope Decisions
- **Data store:** SharePoint/OneDrive-backed shared data source
- **Access:** Microsoft 365 / SSO
- **Who can submit/view:** all authenticated Maersk users
- **Who can edit submitted records:** central admins only
- **Reporting visibility:** submissions appear immediately
- **Data entry standardization:** centrally managed drop-down lists
- **Financial impact in v1:** deferred; focus on lost hours first
- **Incident structure in v1:** each site submits its own separate record

## 7. Users and Roles
### 7.1 Site Submitters
Site operations or local support users who create downtime records for their site.

### 7.2 Regional / Global Leadership
Consumers of dashboards, KPI summaries, and exports.

### 7.3 Central Admins
TCO / central operations users who:
- maintain master data lists
- correct or update submitted records
- monitor data quality

## 8. Core User Stories
### 8.1 Submission
- As a site user, I want to submit a downtime impact record in a few minutes so that my site’s loss is reflected in global reporting.
- As a site user, I want standardized drop-down selections so that my submission aligns with global reporting categories.

### 8.2 Reporting
- As a regional leader, I want to view lost hours by month, region, country, site, customer, and application so that I can identify trends and issues.
- As a global leader, I want headline KPIs so that I can use the data in monthly reviews.

### 8.3 Administration
- As a central admin, I want to edit incorrect submissions so that reported data stays accurate.
- As a central admin, I want to manage drop-down master data so that reporting remains standardized globally.

## 9. MVP Scope
### 9.1 Included in MVP
1. Browser-based form for manual downtime submission
2. Microsoft 365 authenticated access
3. Centralized storage in SharePoint List or equivalent Microsoft 365-backed structure
4. Standardized drop-down master data for:
   - region
   - country
   - site ID
   - customer
   - affected application/system
   - incident category
   - root cause
5. Dashboard and KPI view
6. Table/report view with filtering
7. Excel export
8. Power BI-ready data structure
9. Central admin editing capability

### 9.2 Deferred to Later Phases
1. cost model and cost calculation
2. integration to source outage/system reports
3. automated linkage between system downtime and operational downtime
4. advanced audit workflow / formal approval
5. automated alerts and escalations

## 10. Functional Requirements
### 10.1 Authentication and Access
- The application shall require Microsoft 365 authentication.
- The application shall be accessible from a standard web browser without local installation.
- All authenticated Maersk users shall be able to create records and view reporting.
- Only central admins shall be allowed to edit existing records or maintain master data.

### 10.2 Record Submission
The application shall allow users to create a downtime record with the following fields:
- record ID (system generated)
- incident reference number
- incident date
- outage start date/time
- outage end date/time
- outage duration
- region
- country
- site ID
- customer
- affected application/system
- incident category
- root cause
- estimated operational lost hours
- impact description / business narrative
- submitter name
- submitter email
- submission timestamp
- record status (default: submitted)

### 10.3 Data Validation
- Site, country, region, customer, and application shall be selected from centrally managed drop-down lists.
- The application shall validate required fields before submission.
- Lost hours shall accept numeric values including decimals.
- End date/time shall not be earlier than start date/time.
- The application shall allow outage duration and operational lost hours to differ.
- The application shall support the business case where one hour of system outage can create multiple hours of operational impact.

### 10.4 Editing and Data Quality
- Submitted records shall appear immediately in reporting.
- Central admins shall be able to update incorrect or incomplete records.
- The application should preserve created date, created by, modified date, and modified by metadata from the underlying platform where available.

### 10.5 Reporting and Analytics
The application shall provide:
- total lost hours
- total downtime incidents
- average lost hours per incident
- monthly trend of lost hours
- reporting by region
- reporting by country
- reporting by site
- reporting by customer
- reporting by application/system
- reporting by incident category
- reporting by root cause

### 10.6 Export
- Users shall be able to export the underlying data to Excel.
- The data structure shall be suitable for direct Power BI consumption.

## 11. Initial Dashboard Requirements
### 11.1 KPI Tiles
- Total Lost Hours
- Total Incidents
- Average Lost Hours per Incident
- Reporting Period Selected

### 11.2 Charts for Initial Release
1. **Lost Hours by Month**  
   Primary leadership KPI showing monthly trend

2. **Lost Hours by Region / Country**  
   Comparative view to highlight concentration of impact

3. **Cost Impact by Month**  
   Placeholder chart for future phase; in v1 this can be hidden, disabled, or labeled "coming later" until a cost model is approved

### 11.3 Recommended Additional MVP Views
- Top 10 impacted sites by lost hours
- Top 10 affected applications/systems
- Root cause breakdown
- Detailed filterable records table

## 12. Reporting Filters
The reporting view should support filters for:
- date range
- month
- region
- country
- site ID
- customer
- application/system
- incident category
- root cause

## 13. Data Model
### 13.1 Primary Entity: Downtime Record
Each record represents one site-level submission for one incident occurrence.

### 13.2 Suggested Fields
| Field | Type | Notes |
|---|---|---|
| RecordID | Text / GUID | System generated unique identifier |
| IncidentReference | Text | External or internal incident/ticket number |
| IncidentDate | Date | Date incident occurred |
| OutageStart | DateTime | Start of system issue |
| OutageEnd | DateTime | End of system issue |
| OutageDurationHours | Decimal | Calculated or entered |
| OperationalLostHours | Decimal | Manual estimate from site |
| Region | Lookup | Master data |
| Country | Lookup | Master data |
| SiteID | Lookup | Master data |
| Customer | Lookup | Master data |
| Application | Lookup | Master data |
| IncidentCategory | Lookup | Master data |
| RootCause | Lookup | Master data |
| ImpactDescription | Long text | Business context |
| Status | Lookup/Text | Submitted, corrected, etc. |
| SubmittedBy | Text | Auto from user identity |
| SubmittedAt | DateTime | Auto captured |
| LastModifiedBy | Text | Auto captured |
| LastModifiedAt | DateTime | Auto captured |

## 14. Recommended v1 Solution Design
### 14.1 Architecture
Use a lightweight HTML/CSS/JavaScript front end hosted internally, with Microsoft 365 authentication and a SharePoint List as the system of record.

### 14.2 Why this fits the requirement
- runs in any modern browser
- low setup overhead
- works well with Microsoft 365 identity
- easy export to Excel
- straightforward connection to Power BI
- simple enough for fast MVP delivery

### 14.3 Recommended Components
- **Frontend:** simple HTML/CSS/JavaScript
- **Identity:** Microsoft 365 / Entra ID
- **Data store:** SharePoint List
- **Reporting:** in-app charts plus Power BI consumption
- **Admin data management:** SharePoint list maintenance or lightweight admin screen

## 15. Non-Functional Requirements
### 15.1 Usability
- The submission form should take less than 3 minutes for a trained user to complete.
- The UI should be simple, mobile-laptop friendly, and optimized for common enterprise browsers.

### 15.2 Performance
- Form load should feel near-instant on standard corporate networks.
- Dashboard filters should refresh within a few seconds for expected MVP volumes.

### 15.3 Reliability
- The application should prevent duplicate accidental submissions as far as practical.
- Submitted records should be persisted immediately once confirmed.

### 15.4 Security
- Access shall be restricted to authenticated internal users.
- Edit rights shall be limited to central admins.
- User actions should be attributable through platform metadata.

### 15.5 Scalability
- The design should support global adoption across many sites and monthly reporting growth.
- The data model should remain compatible with future automated system outage integration.

## 16. Future-State Requirements
### Phase 2
- introduce agreed cost model
- capture or calculate estimated financial impact
- add cost KPIs and cost trend charts

### Phase 3
- connect to system outage reports
- associate operational records with technology incidents
- compare system downtime duration vs operational impact duration
- support cases where 1 hour of outage results in 3+ hours of operational loss

### Phase 4
- automate exception analysis
- identify recurring incidents and chronic root causes
- introduce action tracking for service improvement

## 17. Success Metrics
### Adoption
- % of target sites submitting monthly data
- number of active submitting sites

### Data Quality
- % of records with complete mandatory fields
- % of records corrected by central admins

### Reporting Value
- time saved in monthly reporting preparation
- number of recurring issues identified through dashboards
- leadership satisfaction with reporting usefulness

## 18. Risks and Mitigations
| Risk | Impact | Mitigation |
|---|---|---|
| inconsistent submission behavior | weak reporting trust | enforce centralized drop-downs and validation |
| duplicate reporting of same event | inflated impact | require incident reference and clear site-level guidance |
| weak adoption | incomplete global view | keep form simple and provide basic user guidance |
| unclear cost model | misleading financial reporting | defer cost from v1 |
| future integration complexity | rework later | keep outage duration and operational lost hours as separate fields now |

## 19. Open Questions for Next Draft
- Should submitters be able to save drafts before final submission?
- Should incident category and root cause be mandatory in v1 or optional when unknown?
- Should the dashboard default to current month, rolling 12 months, or user-selected dates?
- Should duplicate detection warn when the same site, date, and incident reference are entered twice?
- Should region derive automatically from site ID rather than being separately selected?

## 20. Recommended Next Steps
1. Confirm the open questions above
2. Define the master data lists for site, customer, application, incident category, and root cause
3. Decide whether SharePoint List alone is sufficient or whether a thin API layer is needed
4. Create a wireframe for:
   - submission form
   - KPI dashboard
   - detailed report table
5. Build a v1 prototype in plain HTML/CSS/JavaScript

