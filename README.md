# jira_time_google
Connective tissue between Google Calendar, Jira tasks and logging time.


# TO DO LIST:
- "Pagination" When pulling Jira tasks, do the initial pull, then if there is a nextPageToken, run a secondary pull and repeat that secondary pull until nextPageToken returns empty.
- Recurring events - only need to allocate onces not each time (lookup by eventID)
- Send calendar details to time card after assigning jiras to events.
- Page to close out jiras (don't need to do it in Jira)
- On the History tab, when submitting time, after time records are written to the History tab, identify the range that needs the client, Jira project code and allocation values, and look those up based on each jira id.
- Complete work on the "Create Jira" modal.
