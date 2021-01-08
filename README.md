# Staff Directory with optional department dropdown and advanced query

## Summary

This code is forked from the great work of [this repository](https://github.com/pnp/sp-dev-fx-webparts/tree/master/samples/react-people-directory). Thank you to developers João Mendes, Peter Paul Kirschner, and Sudharsan K.

This web part searches the organizational directory and has letter filters. You can optionally set up department names and add a dropdown list to filter by department.

The web part is built with SPFx, ReactJS, and Microsoft Fabric UI.

## Applies to

* [SharePoint Online](https://docs.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
* [Microsoft Teams](https://products.office.com/en-US/microsoft-teams/group-chat-software)
* [Office 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)

## Minimal Path to Awesome

- Clone this repository
- in the command line run:
  - `npm install`
  - `gulp build`
  - `gulp bundle --ship`
  - `gulp package-solution --ship`
  - `Add to AppCatalog and deploy`