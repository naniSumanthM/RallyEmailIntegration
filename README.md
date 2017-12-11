# RallyEmailIntegration
https://rally1.rallydev.com/slm/login.op

Rally is a scrum tool developed by CA agile. It helps create user stories, tasks and tracks various agile metrics.
Small firms maintain email accounts to support user queries for key enterprise systems.
I have written a custom .NET integration that automates the process of creating user stories, and its assignment to the respective developer.
A notification service has been developed to communicate the creation of new user stories and the assigned developer on a Slack channel.

The solution syncs a Rally server with an email server as a source for new user queries, and a Slack server as a notification service.
