import * as msal from "@azure/msal-browser";
import { CacheOptions } from "@azure/msal-browser/dist/src/config/Configuration";

import {
  iMSAL,
  DataObject,
  Options,
  Auth,
  Request,
  ResetRequest,
} from "./types";

export class MSAL implements iMSAL {
  private msalLibrary: any;
  private tokenExpirationTimers: { [key: string]: undefined | number } = {};
  public data: DataObject = {
    isAuthenticated: false,
    accessToken: "",
    idToken: "",
    user: { name: "", userName: "" },
    custom: {},
    account: {
      homeAccountId: "",
      environment: "",
      tenantId: "",
      username: "",
      localAccountId: "",
      name: "",
      idTokenClaims: {},
    },
  };
  // Config object to be passed to Msal on creation.
  // For a full list of msal.js configuration parameters,
  // visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
  private auth: Auth = {
    clientId: "",
    authority: "",
    redirectUri: "",
    knownAuthorities: [],
    onAuthentication: (error, response) => {},
    onToken: (error, response) => {},
    beforeSignOut: () => {},
  };
  private cache: CacheOptions = {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  };
  // Add here scopes for id token to be used at MS Identity Platform endpoints.
  private loginRequest: Request = {
    scopes: ["openid", "profile", "User.Read"],
  };

  // Add here scopes for access token to be used at MS Graph API endpoints.
  private tokenRequest: ResetRequest = {
    scopes: ["User.Read"],
  };

  private resetRequest: Request = {
    scopes: ["User.Read"],
  };

  constructor(options: Options) {
    if (!options.auth.clientId) {
      throw new Error("auth.clientId is required");
    }
    this.auth = Object.assign(this.auth, options.auth);
    this.cache = Object.assign(this.cache, options.cache);
    this.loginRequest = Object.assign(this.loginRequest, options.loginRequest);
    this.tokenRequest = Object.assign(this.tokenRequest, options.tokenRequest);
    this.resetRequest = Object.assign(this.resetRequest, options.resetRequest);

    const config: msal.Configuration = {
      auth: this.auth,
      cache: this.cache,
    };
    this.msalLibrary = new msal.PublicClientApplication(config);
    this.signIn();
  }
  signIn() {
    return this.msalLibrary
      .loginPopup(this.loginRequest)
      .then((loginResponse) => {
        if (loginResponse !== null) {
          this.data.user.userName = loginResponse.account.username;
          this.data.accessToken = loginResponse.accessToken;
          this.data.idToken = loginResponse.idToken;
          this.data.account = loginResponse.account;
        } else {
          // need to call getAccount here?
          const currentAccounts = this.msalLibrary.getAllAccounts();
          console.log("all accounts: ");
          console.log(currentAccounts);
          if (currentAccounts === null) {
            return;
          } else if (currentAccounts.length > 1) {
            // Add choose account code here
          } else if (currentAccounts.length === 1) {
            this.data.user.userName = currentAccounts[0].username;
            this.data.user.userName = currentAccounts[0].name;
            console.log("this.data: ");
            console.log(this.data);
          }
        }
      })
      .catch(function (error) {
        console.log(error);
      });
  }
  signOut() {
    const logoutRequest = {
      account: this.msalLibrary.getAccountByUsername(this.data.user.userName),
    };
    this.msalLibrary.logout(logoutRequest);
    this.data.accessToken = "";
    this.data.idToken = "";
    this.data.user.userName = "";
  }
  async acquireToken(request = this.loginRequest, retries = 0) {
    this.loginRequest.account = this.data.account;
    console.log("in acquireToken! retries: " + retries);
    try {
      const response = await this.msalLibrary.acquireTokenSilent(request);
      this.handleTokenResponse(null, response);
    } catch (error) {
      console.log("silent token acquisition fails.");
      if (error instanceof msal.InteractionRequiredAuthError) {
        console.log("acquiring token using popup");
        return this.msalLibrary.acquireTokenPopup(request).catch((error) => {
          console.error(error);
        });
      } else if (retries > 0) {
        console.log("in acquireToken with retries: " + retries);
        return await new Promise((resolve) => {
          console.log("setting timeout 5 seconds");
          setTimeout(async () => {
            const res = await this.acquireToken(request, retries - 1);
            resolve(res);
          }, 5 * 1000);
        });
      }
      return false;
    }
  }
  async getTokenRedirect() {
    try {
      let tokenResponse = await this.msalLibrary.handleRedirectPromise();

      let accountObj;
      if (tokenResponse) {
        accountObj = tokenResponse;
      } else {
        accountObj = this.isAuthenticated();
      }

      if (accountObj && tokenResponse) {
        console.log(
          "[AuthService.init] Got valid accountObj and tokenResponse"
        );
        return accountObj;
      }
      if (accountObj) {
        console.log("[AuthService.init] User has logged in, but no tokens.");
        try {
          tokenResponse = await this.msalLibrary.acquireTokenSilent({
            account: this.msalLibrary.getAllAccounts()[0],
            scopes: this.loginRequest.scopes,
          });

          if (tokenResponse.state) {
            this.loginRequest.state = tokenResponse.state;
          }
          return tokenResponse;
        } catch (err) {
          await this.msalLibrary.acquireTokenRedirect({
            scopes: this.loginRequest.scopes,
            state: this.loginRequest.state,
          });
        }
      } else {
        console.log(
          "[AuthService.init] No accountObject or tokenResponse present. User must now login."
        );
        await this.msalLibrary.loginRedirect({
          scopes: this.loginRequest.scopes,
          state: this.loginRequest.state,
        });
      }
    } catch (error) {
      console.error(
        "[AuthService.init] Failed to handleRedirectPromise()",
        error
      );

      if (error.errorMessage.indexOf("AADB2C90118") > -1) {
        try {
          return this.msalLibrary.loginRedirect(this.resetRequest);
        } catch (err) {
          return console.log(err);
        }
      }
      throw new Error("Auth Failed!");
    }
    throw new Error();
  }
  isAuthenticated() {
    if (this.msalLibrary.getAllAccounts() === null) {
      return false;
    } else {
      return true;
    }
  }
  private handleTokenResponse(error, response) {
    if (error) {
      return;
    }
    if (this.data.accessToken !== response.accessToken) {
      this.setToken(
        "accessToken",
        response.accessToken,
        response.expiresOn,
        response.scopes
      );
      console.log("got new accessToken: " + response.accessToken);
    }
    if (this.data.idToken !== response.idToken.rawIdToken) {
      this.setToken(
        "idToken",
        response.idToken.rawIdToken,
        new Date(response.idToken.expiration * 1000),
        [this.auth.clientId]
      );
      console.log("got new idToken: " + response.idToken.rawIdToken);
    }
  }
  private setToken(
    tokenType: string,
    token: string,
    expiresOn: Date,
    scopes: string[]
  ) {
    const expirationOffset = 10000000;
    const expiration =
      expiresOn.getTime() - new Date().getTime() - expirationOffset;
    console.log("set token: " + expiration);
    if (expiration >= 0) {
      console.log("setting token: " + tokenType + " with val: " + token);
      this.data[tokenType] = token;
    }
    if (this.tokenExpirationTimers[tokenType])
      clearTimeout(this.tokenExpirationTimers[tokenType]);
    this.tokenExpirationTimers[tokenType] = window.setTimeout(async () => {
      console.log("auto refreshing token: " + this.auth.autoRefreshToken);
      if (this.auth.autoRefreshToken) {
        await this.acquireToken({ scopes }, 3);
      } else {
        this.data[tokenType] = "";
        console.log("setting token to none:" + this.data.accessToken);
      }
    }, expiration);
  }
}
