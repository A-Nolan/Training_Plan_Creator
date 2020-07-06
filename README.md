# Training_Plan_Creator

Every week I spend roughly two hours creating a Training Plan for my restaurant.

The object of the project is to help automate this process and cut it down into only a couple of minutes.

## Tasks

- [x] Scrape schedule data from MySchedule and save it into an Excel file `schedule.xlsx`
- [x] Check which employees need training this week from `Training Plan.xlsx`
    - [ ] Prioritise `Hygiene & Food Safety` and `Health & Safety` SOCs
    - [x] Prioritise employees who have completed the least amount of SOCs
- [x] Check which SOC the employee needs to do from `Suggested_SOCs.xlsx` file downloaded from PeopleStuff
    - [ ] Prioritise `Hygiene & Food Safety` and `Health & Safety` SOCs
    - [x] Generally take the first SOC from the list
        - [ ] If the first soc is `Hygiene & Food Safety` or `Health & Safety`, choose the next
        - [ ] Set rules for CELs to not do any Kitchen related SOCs
        - [ ] Set rules for Managers to only do certain SOCs

## Features to Add

- [ ] Document code
- [ ] Comment on the more abstract pieces of code
- [ ] Refractor many of the functions and parts of code

## Issues

**None**