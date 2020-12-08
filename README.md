# Outlook virtually categorized
A simple tool to make Outlook / Office365 mail categories available as virtual folders on all devices and email clients.

*If you used labels in Gmail, switched to Outlook and used categories only to realize these do not sync across devices - this tool is for you!*

**Visit: https://berndverst.github.io/outlook-categorized/**

### Motivation
Email is not a piece of paper that needs to be filed into a single bin (folder). Often an email covers multiple topics or categories. **Wouldn't it be nice if we had multiple ways to get to a particular message?**

Gmail got this right by introducing the concept of **labels** which act like tags that can be independently applied to a message and used for message retrieval. Outlook also supports this concept through what is known as **Categories**.

Just like with Gmail Filters, Outlook allows you to create mail rules that apply categories (including custom categories) to messages based on conditions you define. Multiple categories can be applied to a single message in this manner.

### The Problem
Unfortunately, Outlook categories are not supported in the Outlook mobile apps. Most email clients cannot display Outlook categories.

### The Solution

This simple application uses the Microsoft Graph API to create virtual mail folders on the server that correspond to all of your categories. For simplicity, simply obtain an authorization token from the [Microsoft Graph API Explorer](https://developer.microsoft.com/graph/graph-explorer?WT.mc_id=academic-0000-beverst) and paste it into this simple app. No data is sent to another entity other than Microsoft, and only for the sake of calling the Graph API.

The Outlook Desktop app does not persists settings on the server, and therefore other email clients cannot benefit its configuration. Additionally, the configuration performed by this tool is not directly possible through Outlook web or any email client.

### How it works ( in depth )

This is a static web app that does not store data or send data to any third party. The app only communicates with the Microsoft Graph API for the sake of configuration your Outlook account.

To avoid the need to authorize a third party app with your Outlook account we utilize the short-lived access token generated for the [Microsoft Graph API Explorer](https://developer.microsoft.com/graph/graph-explorer?WT.mc_id=academic-0000-beverst). This access token expires in under an hour and is neither stored nor sent anywhere but to the Microsoft Graph API.

The following Graph API calls are bing made:
1. List all mail categories.
1. List mail folders to obtain the ID of the root folder (the parent of the Inbox folder).
1. Create mail search folders with the name of each of the categories, configured to track these categories.

The "mail search folder" will appear as regular folders in all Outlook clients, but will automatically track messages with the corresponding category.
