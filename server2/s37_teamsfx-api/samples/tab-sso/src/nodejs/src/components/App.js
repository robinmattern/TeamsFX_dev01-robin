// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { BrowserRouter as Router, Route } from "react-router-dom";

import { PublicClientApplication, BrowserCacheLocation } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";

import Privacy from "./Privacy";
import TermsOfUse from "./TermsOfUse";
import TabConfig from "./TabConfig";
import ConsentPopup from "./ConsentPopup";
import ClosePopup from "./ClosePopup";
import Tab from "./Tab";

import './App.css';

/**
 * Configuration object to be passed to MSAL instance on creation. 
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md 
 */
const msalConfig = {
    auth: {
        clientId: process.env.REACT_APP_AZURE_APP_REGISTRATION_ID,
        authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_INFO}`,
        redirectUri: `${window.location.origin}/auth-end`,
        navigateToLoginRequestUrl: false,
    },
    cache: {
        cacheLocation: BrowserCacheLocation.LocalStorage // Ensure cache is shared between windows/tabs
    }
}

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
function App() {

    // Initialize the Microsoft Teams SDK
    microsoftTeams.initialize();

    // You will first need to initialize an instance of PublicClientApplication then pass this to MsalProvider as a prop
    const msalInstance = new PublicClientApplication(msalConfig);

    // Display the app home page hosted in Teams
    return (
        <MsalProvider instance={msalInstance}>
            <Router>
                <Route exact path="/privacy" component={Privacy} />
                <Route exact path="/termsofuse" component={TermsOfUse} />
                <Route exact path="/configure" component={TabConfig} />
                <Route exact path="/ssoDemo" component={Tab} />
                <Route exact path="/auth-start" component={ConsentPopup} />
                <Route exact path="/auth-end" component={ClosePopup} />
            </Router>
        </MsalProvider>
    );
}

export default App;
