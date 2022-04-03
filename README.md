# ![Enhanced Power Apps Embed Webpart icon](resources/icon-small.png "Enhanced Power Apps Embed Webpart icon") Enhanced Power Apps Embed Webpart

## Table of Contents
  - [Summary](#summary)
  - [Used SharePoint Framework Version](#used-sharepoint-framework-version)
  - [Applies to](#applies-to)
  - [Prerequisites](#prerequisites)
  - [Solution](#solution)
  - [Version history](#version-history)
  - [Supported languages](#supported-languages)
  - [How to implement](#how-to-implement)

## Summary

This SPFx (SharePoint Framework) webpart improves on the existing react-enhanced-powerapps solution developed by Hugo Bernier ([GitHub](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-enhanced-powerapps)) so I highly recommend you check out his work and see what the orginal version already could do. When using his webpart myself, I thought to myself "passing a SharePoint context parameter to your embedded Power App is great, but why should it be limited to only one?". That is why this enhanced webpart allows you to embed a Power App in SharePoint and use various page context parameters within your app by calling the ```Param("{contextParamName}")``` function, as well as configuring your own dynamic parameter in the property panel. See [how to implement](#how-to-implement) for a detailed description of how it works.

**[<img src="https://external-content.duckduckgo.com/iu/?u=https%3A%2F%2Fwww.iconsdb.com%2Ficons%2Fpreview%2Froyal-blue%2Fdata-transfer-download-xxl.png&f=1&nofb=1" alt="Download .sppkg file" style="width:15px;margin-right:10px;"/> __Download the .sppkg file here!__](sharepoint/solution/enhanced-power-apps-embed.sppkg)**

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.14.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

> - Node.js v10/12/14
> - (optional) An M365 account. Get your own free Microsoft 365 tenant from [M365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)

## Solution

Solution|Author(s)
--------|---------
Enhanced Power Apps Embed Webpart | cup o'365 ([Contact](mailto:info@cupo365.gg), [Website](https://cupo365.gg))

## Version history

Version|Date|Comments
-------|----|--------
1.0|April 3, 2022|Initial release


## Supported languages

- English
- Dutch
---

## How to implement

By default, the webpart provides various SharePoint context variables by passing them to the Power App via the HTML iframe element URL as parameters.
These parameters are (by default) as follows:

Name|Type|Description|Example
-------|----|--------|--------
tenantId|guid|Unique identifier of the tenant|"f494b1c5-bb1c-46fc-9c9b-b3d6bbf41e43"
languageId|number|LCID value of the configured language on the page the app is embedded on ([see also](https://www.pkbullock.com/resources/reference-sharepoint-online-languages-ids/))|1033
languageName|text|The language tag name of the configured language on the page the app is embedded on ([see also](https://www.pkbullock.com/resources/reference-sharepoint-online-languages-ids/))|"en-US"
siteTitle|text|The name of site on which the app is embedded on|"Dev team"
portalUrl|text|The SharePoint base URL of the page on which the app is embedded on|"https://cupo365.sharepoint.com/"
absoluteUrl|text|The site URL of the page which the app is embedded on|"https://cupo365.sharepoint.com/sites/DevTeam"
serverRelativeUrl|text|The relative path from the SharePoint base URL of the page on which the app is embedded on|"/sites/DevTeam"
serverRequestPath|text|The relative path from the server relative URL of the page on whic the app is embedded on|"/SitePages/Enhanced Power Apps Embed test.aspx"
groupId|guid|Unique identifier of the Office-group intertwined with the SharePoint site on which the app is embedded on|"d0496f03-563d-4f50-8b29-f84560ddfdd1"
groupType|text|Whether the group intertwined with the SharePoint site on which the app is embedded on is public or private|"Public"
isTeamsConnectedSite|booolean|Whether the site of the page the app is embedded on is a Teams site|false
isTeamsChannelSite|boolean|Whether the site of the page the app is embedded on is a Teams channel|false
isArchived|boolean|Whether the site on which the app is embedded on is archived|false
isHubSite|boolean|Whether the site on which the app is embedded on is a hub site|false
userId|guid|Unique identifier of the user|"d581c3b1-011f-4cef-bf15-cf4074eca852"
displayName|text|The Office display name of the user|"Lennart de Waart"
email|text|The email address of the user|"len@cupo365.gg"
isAnonymousGuestUser|boolean|Whether the user is an anonymous guest|false
isExternalGuestUser|boolean|Whether the user is an external guest|false
isSiteAdmin|boolean|Whether the user is the admin of the site on which the app is embedded on|false
isSiteOwner|boolean|Whether the user is the owner of the site on which the app is embedded on|false

You can fetch these context parameters in your Power App by calling the ````Params("{contextParameterName}")```` function. You can fetch and use them individually, or use this formula in your app's OnStart to create a global variable of a compiled record of all of these context parameters which you can call later on: 
```` 
Set(
    varEmbeddedContext,
    {
        aadInfo: {tenantId: Param("tenantId")},
        cultureInfo: {
            languageId: Param("languageId"),
            languageName: Param("languageName")
        },
        site: {
            title: Param("siteTitle"),
            portalUrl: Param("portalUrl"),
            absoluteUrl: Param("absoluteUrl"),
            serverRelativeUrl: Param("serverRelativeUrl"),
            serverRequestPath: Param("serverRequestPath"),
            group: {
                id: Param("groupId"),
                type: Param("groupType")
            },
            isTeamsConnectedSite: Param("isTeamsConnectedSite"),
            isTeamsChannelSite: Param("isTeamsChannelSite"),
            isArchived: Param("isArchived"),
            isHubSite: Param("isHubSite")
        },
        user: {
            id: Param("userId"),
            displayName: Param("displayName"),
            email: Param("email"),
            isAnonymousGuestUser: Param("isAnonymousGuestUser"),
            isExternalGuestUser: Param("isExternalGuestUser"),
            isSiteAdmin: Param("isSiteAdmin"),
            isSiteOwner: Param("isSiteOwner")
        }
    }
);
````
Furthermore, you can configure your own (dynamic) context parameter in the property pane of the webpart, just like the version of Hugo Bernier.


**[<img src="https://external-content.duckduckgo.com/iu/?u=https%3A%2F%2Fwww.iconsdb.com%2Ficons%2Fpreview%2Froyal-blue%2Fdata-transfer-download-xxl.png&f=1&nofb=1" alt="Download .sppkg file" style="width:15px;margin-right:10px;"/> __Download the .sppkg file here!__](sharepoint/solution/enhanced-power-apps-embed.sppkg)**
