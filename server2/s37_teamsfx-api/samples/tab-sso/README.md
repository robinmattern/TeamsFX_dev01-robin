# Teams Tab Single Sign-on (SSO) Sample

## Summary

This sample shows how to create a tab for Teams that uses the built-in [Single Sign-On (SSO)](https://docs.microsoft.com/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso) capabilities to get a secure token containing important user information like user display name, Azure AD Object Id, UPN and Tenant ID. The sample also shows how to create a web service to enable exchange for an [On-Behalf-Of](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) token that can be used to call another service, such as the Microsoft Graph, to access additional capabilities.

Please note that using SSO does *not* eliminate the user seeing consent popups - these are still required in order to ensure the application has approval to access the user's information. To assist in this, read [here](https://docs.microsoft.com/en-gb/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#tenant-admin-consent) on how to have an administrator pre-consent on behalf of all the users in the organisation.

Below is a screenshot of the SSO tab in action (this is for the .NET version - the Node.js version shows the same information but in a slightly different way).

![picture of the app in action](assets/images/TabScreenshot.png)

**Please note** that this sample actually contains two separate, alternate projects: a .NET version and a Node.js version:

  * The .NET version uses a more 'vanilla JS' approach, in order to show how SSO can be implemented in such a scenario. The .NET back-end is protected by [Microsoft.Identity.Web](https://github.com/AzureAD/microsoft-identity-web), which validates the SSO token and calls Graph using the Graph SDK.
  * The Node.js version uses a React front end, in order to demonstrate SSO using a front end framework. It also has been updated with an alternative pattern that does Graph queries server-side rather than passing the access token back to the web browser. It also adds logic to verify the SSO token on the server side.

## Frameworks

**Dotnet Version**

![drop](https://img.shields.io/badge/aspnetcore-3.1-green.svg)

**Node.js Version**

![drop](https://img.shields.io/badge/Node.js-12.6.1-green.svg)
![drop](https://img.shields.io/badge/React-16.13.1-green.svg)

## Prerequisites

**Node.js Version**
* [Node.js](https://nodejs.org) version 12.6.1 or higher

```bash
    # determine node version
    node --version
```

**Dotnet Version**
* [Dotnet Core](https://dotnet.microsoft.com/) version 3.1 or higher

**Both Versions**
* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)
* Optional: Static hosting such as [github pages](https://pages.github.com/)

## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|November 25, 2020|Hilton Giesenow|Initial release
1.1|December 05, 2020|Hilton Giesenow|Node.js version added
2.0|February 09, 2021|Doğan Erişen|Added MSAL.js support to Node version only
2.1|March 05, 2021|Doğan Erişen|Added server-side token validation and moved Graph calls to server in Node version only

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

### 1. Clone this repository

```bash
    git clone https://github.com/pnp/teams-dev-samples.git
```

### 2. Register Azure AD application

For detailed steps on how to do this, read [Registering your app through the Azure Active Directory portal in-depth](https://docs.microsoft.com/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#registering-your-app-through-the-azure-active-directory-portal-in-depth). We summarize the steps below:

1. Navigate to the [Azure portal](https://portal.azure.com) and select the **Azure AD** service.
1. Select the **App Registrations** blade on the left, then select **New registration**.
1. In the **Register an application page** that appears, enter your application's registration information:
   - In the **Name** section, enter a meaningful application name that will be displayed to users of the app, for example `teams-sso-sample`.
   - Under **Supported account types**, select **Accounts in any organizational directory**.
   - In the **Redirect URI (optional)** section, select **Single-page application** in the combo-box and enter a redirect URI (e.g. `https://contoso.ngrok.io/auth-end`).
   - Select **Register** to create the application.
1. In the app's registration screen, find and note the **Application (client) ID**. You use this value in your app's configuration file(s) later in your code.
1. In the app's registration screen, select the **Certificates & secrets** blade in the left to open the page where we can generate secrets and upload certificates.
1. In the **Client secrets** section, select **New client secret**:
   - Type a key description (for instance `app secret`),
   - Select one of the available key durations (**6 months**, **12 months** or **Custom**) as per your security posture.
   - The generated key value will be displayed when you select the **Add** button. Copy and save the generated value for use in later steps.
1. In the app's registration screen, select the **API permissions** blade in the left to open the page where we add access to the APIs that your application needs.
   - Select the **Add a permission** button and then,
   - Ensure that the **Microsoft APIs** tab is selected.
   - In the *Commonly used Microsoft APIs* section, select **Microsoft Graph**
   - In the **Delegated permissions** section, select **openid**, **profile**, **email**, **offline_access** and **user.read** in the list. Use the search box if necessary.
   - Select the **Add permissions** button at the bottom.
1. In the app's registration screen, select the **Expose an API** blade to the left to open the page where you can declare the parameters to expose this app as an API for which client applications can obtain [access tokens](https://docs.microsoft.com/azure/active-directory/develop/access-tokens) for.
The first thing that we need to do is to declare the unique [resource](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow) URI that the clients will be using to obtain access tokens for this Api. To declare an resource URI, follow the following steps:
   - Select `Set` next to the **Application ID URI** to generate a URI that is unique for this app.
   - For this sample, edit the proposed Application ID URI (e.g. `api://contoso.ngrok.io/{appId}`) and selecting **Save**.
1. All APIs have to publish a minimum of one [scope](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-auth-code-flow#request-an-authorization-code) for the client's to obtain an access token successfully. To publish a scope, follow the following steps:
   - Select **Add a scope** button open the **Add a scope** screen and Enter the values as indicated below:
        - For **Scope name**, use `access_as_user`.
        - Select **Admins and users** options for **Who can consent?**.
        - For **Admin consent display name** type `Access teams-sso-sample`.
        - For **Admin consent description** type `Allows the app to access teams-sso-sample as the signed-in user.`
        - For **User consent display name** type `Access teams-sso-sample`.
        - For **User consent description** type `Allow the application to access teams-sso-sample on your behalf.`
        - Keep **State** as **Enabled**.
        - Select the **Add scope** button on the bottom to save this scope.
1. Still in the **Expose an API** blade, find the **Authorized client applications** section, follow the steps below:
        - Select **Add a client application**
        - In the panel that opens to the right, enter `1fec8e78-bce4-4aaf-ab1b-5451cc387264`
        - Check the **Authorized scopes** checkbox for the scope you've just exposed (e.g. `access_as_user`)
        - Select **Add application** to save.
        - In the **Authorized client applications** section, select **Add a client application** again.
        - In the panel that opens to the right, enter `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`
        - Check the **Authorized scopes** checkbox for the scope you've just exposed (e.g. `access_as_user`)
        - Select **Add application** to save.

### 3. Update app configuration & run the web application

**Dotnet Version**

* See the [`dotnet`](src/dotnet) folder inside this sample's `src` folder

**Node.js Version**

* See [`node`](src/nodejs) folder inside this sample's `src` folder

### 4. Update & package the Teams app manifest

Inside the [`src`](src) folder for this sample is a [`manifest.json`](src/manifest.json) file. The following needs to be changed in this file:

1. The `"id"` value must be populated with a new Guid value. You can do this in various ways depending on your platform of choice, but a simple PowerShell command is:

```powershell
New-Guid
```

2. The `{appId}` values (near the bottom of the manifest) must be replaced with the Azure Application ID you generated in step 2 above, when generating the new Azure AD application.

3. The `{ngrokSubdomain}` value must be replaced with whatever ngrok subdomain you are using. If you are using another tunneling tool, you might need to replace the entire `{ngrokSubdomain}.ngrok.io` value, and when you create a Production version of your application you will similarly need to supply a complete production URL.

4. From within the [`src`](src) folder, in the command line, run the `gulp` command (you will need the [gulp.js](https://gulpjs.com/) task runner installed to do this). This will generate a .zip manifest file that can be easily uploaded to Microsoft Teams.

### 5. Upload the manifest to Teams

There are a few possible options to do this, depending on your development tools and platform. The easiest is simple to use Teams' [App Studio](https://docs.microsoft.com/microsoftteams/platform/concepts/build-and-test/app-studio-overview) tool, in particular the [`manifest editor`](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/app-studio-overview#manifest-editor) tab which allows you to import a manifest (i.e. the one you created in step 4 above) and immediately install it.

## Features

This sample demonstrates how to create a tab for Teams that uses Single-Sign-On (SSO) as well as how to exchange the SSO token for an On-Behalf-Of (OBO) token.

## Further Reading

* [What Are Tabs (Overview of Tabs in Teams)](https://docs.microsoft.com/microsoftteams/platform/tabs/what-are-tabs)
* [Tab Authentication](https://docs.microsoft.com/microsoftteams/platform/tabs/how-to/authentication/auth-flow-tab)
* [Single Sign-On (SSO) For Tabs](https://docs.microsoft.com/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso)
* [Microsoft Authentication Library (MSAL) 2.0](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)

<img src="https://pnptelemetry.azurewebsites.net/teams-dev-samples/samples/tab-sso />