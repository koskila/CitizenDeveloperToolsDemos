# CitizenDeveloperToolsDemos

This repo contains some of my citizen dev demos - basically mostly Azure functions that are meant to be hooked into SharePoint, Flow or both.

## Koskila.CitizenDeveloperTools

This project contains the Azure Functions I'm attaching to Flows and webhooks on SharePoint lists for my demos. They work the following way:
- BetterCopyFunction: This Azure Function can copy a classic publishing page from site 1 to site 2. See it in action here: https://www.youtube.com/watch?v=DaX6V_fFqy8
- SharePointWebhook: Add this Azure Function's URL (via POST) to a SharePoint list's webhook, and it'll run and push the notification of a list item change to another Flow, that can then decide what it does with the info. Requires app setting "NotificationFlowAddress"

### Prerequisites

The following app settings are required for these functions:
- ClientId
- ClientSecret
- SiteCollectionRequests_TenantAdminSite

Additionally, the following environments/tools:

#### For local dev
- Azure CLI
- Postman (or similar) to simulate webhook payloads

#### For production deployment
- SharePoint site
- Azure subscription
